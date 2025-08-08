// mealsPivot.js
import { tmpdir } from "os";
import { join } from "path";
import ExcelJS from "exceljs";
import { parseISO, format } from "date-fns";
import { getStorage } from "firebase-admin/storage";

const greyFill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFEFEFEF' }
};

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
        const qty = row[slotKey];
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

    // New Row 2: Catering Grid
    sheet.addRow(["Catering Grid"]);

    // Row 3: Generated timestamp (bold)
    const generatedAt = format(new Date(), "EEEE d MMM yyyy, h:mm a");
    const generatedRow = sheet.addRow([`Generated ${generatedAt}`]);
    generatedRow.font = { bold: true };

    // Row 4: Spacer
    sheet.addRow([]);

    // Row 5: Name + Role + merged date cells
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

    // Manually populate row 5 so it stays at the expected position
    const headerRowNumber = 5;
    headerRow.forEach((value, index) => {
      const cell = sheet.getCell(headerRowNumber, index + 1);
      cell.value = value;
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Merge the date columns
    mergeRanges.forEach(({ start, end }) => {
      sheet.mergeCells(5, start, 5, end);
      const cell = sheet.getCell(5, start);
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { bold: true };
    });

    // Row 6: Meal abbreviations
    const slotRow = ["", ""];
    for (let i = 0; i < allDates.length; i++) {
      for (const { abb } of sortedSlots) {
        slotRow.push(abb);
      }
    }

    const slotRowRef = sheet.addRow(slotRow);
    slotRowRef.font = { bold: true };
    slotRowRef.alignment = { vertical: "middle", horizontal: "center" };

    // Row 7+: Data rows
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

    const startRow = 7; // First row of data
    const endRow = sheet.lastRow.number - 1; // Last data row before total
    const colCount = sheet.columnCount;

    for (let col = 3; col <= colCount; col++) {
      const colLetter = sheet.getColumn(col).letter;
      totalRow.getCell(col).value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
      totalRow.getCell(col).font = { bold: true };
      totalRow.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
    }

    totalRow.commit();

    // Center-align all data and total cells except for columns A and B
    for (let rowNum = 7; rowNum <= sheet.lastRow.number; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let col = 3; col <= sheet.columnCount; col++) {
        const cell = row.getCell(col);
        cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
      }
    }

    // Borders: horizontal below row 6 and below last data row, and vertical left borders for slot groupings
    // Get last data row (before totals)
    const lastDataRow = sheet.lastRow.number - 1;
    // Horizontal border below row 6 (meal abbreviations)
    const borderStyle = { style: 'thin', color: { argb: 'FF000000' } };
    const keyBorderStyle = { style: 'medium', color: { argb: 'FF000000' } };
    // Below row 6
    const mealAbbRowNum = 6;
    const mealAbbRow = sheet.getRow(mealAbbRowNum);
    for (let col = 1; col <= sheet.columnCount; col++) {
      mealAbbRow.getCell(col).border = {
        ...(mealAbbRow.getCell(col).border || {}),
        bottom: borderStyle
      };
    }
    // Below last data row (just above totals)
    const lastDataRowObj = sheet.getRow(lastDataRow);
    for (let col = 1; col <= sheet.columnCount; col++) {
      lastDataRowObj.getCell(col).border = {
        ...(lastDataRowObj.getCell(col).border || {}),
        bottom: borderStyle
      };
    }

    // Vertical left borders for each slot grouping
    // Slot groupings start at column 3 (C), then every (number of slots) columns after
    const slotCount = sortedSlots.length;
    const totalDates = allDates.length;
    const slotStartCols = [];
    let colIdx = 3;
    for (let i = 0; i < totalDates; i++) {
      slotStartCols.push(colIdx);
      colIdx += slotCount;
    }
    // For each group start col, apply left border from row 5 to totals row
    const firstRowWithSlots = 5;
    const lastRowWithTotals = sheet.lastRow.number;
    for (const slotCol of slotStartCols) {
      for (let rowNum = firstRowWithSlots; rowNum <= lastRowWithTotals; rowNum++) {
        const cell = sheet.getRow(rowNum).getCell(slotCol);
        cell.border = {
          ...(cell.border || {}),
          left: borderStyle
        };
      }
    }

    // Outer medium border around full table
    const outerBorderStyle = { style: 'medium', color: { argb: 'FF000000' } };
    const topRow = 5;
    const bottomRow = sheet.lastRow.number;
    const leftCol = 1;
    const rightCol = sheet.columnCount;

    for (let col = leftCol; col <= rightCol; col++) {
      sheet.getRow(topRow).getCell(col).border = {
        ...(sheet.getRow(topRow).getCell(col).border || {}),
        top: outerBorderStyle
      };
      sheet.getRow(bottomRow).getCell(col).border = {
        ...(sheet.getRow(bottomRow).getCell(col).border || {}),
        bottom: outerBorderStyle
      };
    }

    for (let row = topRow; row <= bottomRow; row++) {
      sheet.getRow(row).getCell(leftCol).border = {
        ...(sheet.getRow(row).getCell(leftCol).border || {}),
        left: outerBorderStyle
      };
      sheet.getRow(row).getCell(rightCol).border = {
        ...(sheet.getRow(row).getCell(rightCol).border || {}),
        right: outerBorderStyle
      };
    }

    // Set column widths:
    // - The first two columns ("Name" and "Role") are made wider (20 units) to accommodate longer text.
    // - All other columns (meal abbreviations) are narrower (4 units) since they only hold small numeric values.
    sheet.columns.forEach((col, idx) => {
      col.width = idx < 2 ? 20 : 3;
    });


    // Add key table below totals
    const keyStartRow = sheet.lastRow.number + 2;
    sheet.getRow(keyStartRow).getCell(1).value = "Key";
    sheet.getRow(keyStartRow).getCell(1).font = { bold: true };
    // Add light grey background to all 12 key title cells
    const keyTitleRowObj = sheet.getRow(keyStartRow);
    for (let col = 1; col <= 12; col++) {
      keyTitleRowObj.getCell(col).fill = greyFill;
    }

    // Header for key
    const keyHeaderRow = sheet.getRow(keyStartRow + 1);
    keyHeaderRow.getCell(1).value = "Meal";
    keyHeaderRow.getCell(2).value = "Abbreviation";
    keyHeaderRow.getCell(3).value = "Location";

    keyHeaderRow.getCell(1).font = { italic: true };
    keyHeaderRow.getCell(2).font = { italic: true };
    keyHeaderRow.getCell(3).font = { italic: true };
    // Add light grey background to all 12 header cells
    for (let col = 1; col <= 12; col++) {
      keyHeaderRow.getCell(col).fill = greyFill;
    }

    // Data for key
    sortedSlots.forEach((slot, index) => {
      const row = sheet.getRow(keyStartRow + 2 + index);

      const mealCell = row.getCell(1);
      const abbCell = row.getCell(2);
      const locationCell = row.getCell(3);

      mealCell.value = slot.name;
      abbCell.value = slot.abb;
      locationCell.value = slot.location;

      // Apply light grey background
      [mealCell, abbCell, locationCell].forEach(cell => {
        cell.fill = greyFill;
      });

      sheet.mergeCells(keyStartRow + 2 + index, 3, keyStartRow + 2 + index, 12);
    });

    // Apply medium border around the key table
    const keyTopRow = keyStartRow;
    const keyBottomRow = keyStartRow + 1 + sortedSlots.length;
    const keyLeftCol = 1;
    const keyRightCol = 12;

    for (let col = keyLeftCol; col <= keyRightCol; col++) {
      sheet.getRow(keyTopRow).getCell(col).border = {
        ...(sheet.getRow(keyTopRow).getCell(col).border || {}),
        top: keyBorderStyle
      };
      sheet.getRow(keyBottomRow).getCell(col).border = {
        ...(sheet.getRow(keyBottomRow).getCell(col).border || {}),
        bottom: keyBorderStyle
      };
    }

    for (let row = keyTopRow; row <= keyBottomRow; row++) {
      sheet.getRow(row).getCell(keyLeftCol).border = {
        ...(sheet.getRow(row).getCell(keyLeftCol).border || {}),
        left: keyBorderStyle
      };
      sheet.getRow(row).getCell(keyRightCol).border = {
        ...(sheet.getRow(row).getCell(keyRightCol).border || {}),
        right: keyBorderStyle
      };
    }


    // Save to temp file
    const fileName = `${eventName}_Catering_${Date.now()}.xlsx`;
    const filePath = join(tmpdir(), fileName);

    const isRunningLocally = process.env.FUNCTIONS_EMULATOR === 'true';

    if (isRunningLocally) {
      const fs = await import('fs');
      const localDir = '/Users/apndavies/Coding/Flair Schedules/output';
      if (!fs.existsSync(localDir)) {
        fs.mkdirSync(localDir);
      }
      const localPath = `${localDir}/${fileName}`;
      await workbook.xlsx.writeFile(localPath);
      console.log(`✅ File saved locally at ${localPath}`);
      return res.json({ status: "success", localPath });
    } else {
      // Save to temp file
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
    }

  } catch (err) {
    console.error("❌ Error in mealsPivotHandler:", err);
    return res.status(500).json({ error: "Server error", details: err.message });
  }
}