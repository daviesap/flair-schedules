// mealsPivot.js
import { tmpdir } from "os";
import { join } from "path";
import ExcelJS from "exceljs";
import { parseISO, format } from "date-fns";
import { getStorage } from "firebase-admin/storage";

export async function mealsPivotHandler(req, res) {
  try {
    const { eventName = "Event", slots = [], data = [] } = req.body;

    if (!Array.isArray(slots) || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid or missing 'slots' or 'data'" });
    }

    const sortedSlots = [...slots].sort((a, b) => a.slot - b.slot);

    // Map slot keys to abbreviations (e.g., slot1 -> B)
    const slotMap = new Map();
    for (const { slot, abb } of sortedSlots) {
      slotMap.set(`slot${slot}`, abb);
    }

    const allDates = [...new Set(data.map(d => d.Date))].sort();

    // Build pivot structure
    const pivot = {};
    for (const row of data) {
      const name = row.name || "Unknown";
      const role = row.role || "";
      const key = `${name}__${role}`;
      const date = row.Date;

      if (!pivot[key]) pivot[key] = {};
      if (!pivot[key][date]) pivot[key][date] = {};

      for (const [slotKey, abb] of slotMap.entries()) {
        const qty = row[slotKey]?.qty;
        if (qty) {
          pivot[key][date][abb] = qty;
        }
      }
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Meals");

    // Row 1: Title
    const titleRow = sheet.addRow([eventName]);
    titleRow.font = { bold: true };

    // Row 2: Generated timestamp
    const generatedAt = format(new Date(), "EEEE d MMM yyyy, h:mm a");
    sheet.addRow([`Generated ${generatedAt}`]);

    // Row 3: Spacer
    sheet.addRow([]);

    // Row 4: Name + Role + merged date cells
    const headerRow = ["Name", "Role"];
    const mergeRanges = [];

    let currentCol = 3;
    for (const date of allDates) {
      const formattedDate = format(parseISO(date), "EEE d MMM");
      const slotCount = sortedSlots.length;
      const startCol = currentCol;
      const endCol = currentCol + slotCount - 1;

      headerRow.push(formattedDate);
      mergeRanges.push({ start: startCol, end: endCol });

      for (let i = 1; i < slotCount; i++) {
        headerRow.push(""); // fill merged cells
      }

      currentCol += slotCount;
    }

    // Manually populate row 4 so it stays at the expected position
    const headerRowNumber = 4;
    headerRow.forEach((value, index) => {
      const cell = sheet.getCell(headerRowNumber, index + 1);
      cell.value = value;
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Merge the date columns
    mergeRanges.forEach(({ start, end }) => {
      sheet.mergeCells(4, start, 4, end);
      const cell = sheet.getCell(4, start);
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { bold: true };
    });

    // Row 5: Meal abbreviations
    const slotRow = ["", ""];
    for (let i = 0; i < allDates.length; i++) {
      for (const { abb } of sortedSlots) {
        slotRow.push(abb);
      }
    }

    const slotRowRef = sheet.addRow(slotRow);
    slotRowRef.font = { bold: true };
    slotRowRef.alignment = { vertical: "middle", horizontal: "center" };

    // Row 6+: Data rows
    for (const [key, dateData] of Object.entries(pivot)) {
      const [name, role] = key.split("__");
      const row = [name, role];

      for (const date of allDates) {
        const meals = dateData[date] || {};
        for (const { abb } of sortedSlots) {
          row.push(meals[abb] || "");
        }
      }

      sheet.addRow(row);
    }

    // Row for totals
    const lastRowNumber = sheet.lastRow.number + 1;
    const totalRow = sheet.getRow(lastRowNumber);
    totalRow.getCell(1).value = "TOTAL";
    totalRow.getCell(1).font = { bold: true };
    totalRow.getCell(1).alignment = { vertical: "middle" };
    totalRow.getCell(2).value = "";
    totalRow.getCell(2).alignment = { vertical: "middle" };

    const startRow = 6; // First row of data
    const endRow = sheet.lastRow.number - 1; // Last data row before total
    const colCount = sheet.columnCount;

    for (let col = 3; col <= colCount; col++) {
      const colLetter = sheet.getColumn(col).letter;
      totalRow.getCell(col).value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
      totalRow.getCell(col).font = { bold: true };
      totalRow.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
    }

    totalRow.commit();

    // Set column widths:
    // - The first two columns ("Name" and "Role") are made wider (20 units) to accommodate longer text.
    // - All other columns (meal abbreviations) are narrower (4 units) since they only hold small numeric values.
    sheet.columns.forEach((col, idx) => {
      col.width = idx < 2 ? 20 : 3;
    });


    // Define dataStartCol and dataEndCol before use
    const dataStartCol = 1; // Column A
    const dataEndCol = sheet.columnCount;

    // Thin border below row 5
    const slotBorderRow = sheet.getRow(5);
    for (let colNum = dataStartCol; colNum <= dataEndCol; colNum++) {
      const cell = slotBorderRow.getCell(colNum);
      cell.border = {
        ...cell.border,
        bottom: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }

    // Apply thick black border around the entire data block
    const dataStartRow = 4; // Row with headers
    const dataEndRow = sheet.lastRow.number; // Includes totals

    for (let rowNum = dataStartRow; rowNum <= dataEndRow; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = dataStartCol; colNum <= dataEndCol; colNum++) {
        const cell = row.getCell(colNum);
        const border = {};
        if (rowNum === dataStartRow) border.top = { style: 'medium', color: { argb: 'FF000000' } };
        if (rowNum === dataEndRow) border.bottom = { style: 'medium', color: { argb: 'FF000000' } };
        if (colNum === dataStartCol) border.left = { style: 'medium', color: { argb: 'FF000000' } };
        if (colNum === dataEndCol) border.right = { style: 'medium', color: { argb: 'FF000000' } };
        cell.border = border;
      }
      // Thin border above total row
      if (rowNum === dataEndRow) {
        const totalBorderRow = sheet.getRow(dataEndRow);
        for (let colNum = dataStartCol; colNum <= dataEndCol; colNum++) {
          const cell = totalBorderRow.getCell(colNum);
          cell.border = {
            ...cell.border,
            top: { style: 'thin', color: { argb: 'FF000000' } }
          };
        }
      }
    }

    // Save to temp file
    const fileName = `${eventName}_Catering_${Date.now()}.xlsx`;
    const filePath = join(tmpdir(), fileName);
    await workbook.xlsx.writeFile(filePath);

    // Upload to Firebase Storage and make public
    const bucket = getStorage().bucket();
    const destPath = `meals/${fileName}`;
    await bucket.upload(filePath, {
      destination: destPath,
      metadata: {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    });
    await bucket.file(destPath).makePublic();

    const publicUrl = `https://storage.googleapis.com/${bucket.name}/${destPath}`;
    return res.json({ status: "success", fileUrl: publicUrl });

  } catch (err) {
    console.error("❌ Error in mealsPivotHandler:", err);
    return res.status(500).json({ error: "Server error", details: err.message });
  }
}