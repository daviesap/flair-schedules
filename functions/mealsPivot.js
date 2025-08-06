// functions/mealsPivot.js
import ExcelJS from "exceljs";
import { getStorage } from "firebase-admin/storage";
import { v4 as uuidv4 } from "uuid";

// Format long date (e.g. "Tuesday 12th May 1970")
function formatLongDate(date) {
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];
  const day = dayNames[date.getUTCDay()];
  const dayNum = date.getUTCDate();
  const suffix = (n) => (n > 3 && n < 21 ? "th" : ["st", "nd", "rd"][n % 10 - 1] || "th");
  const month = monthNames[date.getUTCMonth()];
  const year = date.getUTCFullYear();
  return `${day} ${dayNum}${suffix(dayNum)} ${month} ${year}`;
}

// Format short date (e.g. "Tue 12 May")
function formatShortDate(dateStr) {
  const date = new Date(dateStr);
  return date.toLocaleDateString("en-GB", {
    weekday: "short",
    day: "numeric",
    month: "short",
  });
}

export async function mealsPivotHandler(req, res) {
  try {
    const { eventName, data } = req.body;
    if (!eventName || !Array.isArray(data)) {
      return res.status(400).json({ error: "Missing eventName or data" });
    }

    const slotKeys = ["slot1", "slot2", "slot3", "slot4"];

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Meals Pivot");

    // Title and generation date
    worksheet.addRow([eventName]).font = { bold: true, size: 14 };
    worksheet.addRow([`Generated ${formatLongDate(new Date())}`]);
    worksheet.addRow([]);

    // Header rows
    const headerRow1 = worksheet.addRow(["Name", "Role"]);
    const headerRow2 = worksheet.addRow(["", ""]);
    headerRow1.font = { bold: true };
    headerRow2.font = { bold: true };

    // Create map of dates to slot data
    const dateMap = {};
    for (const entry of data) {
      const date = entry.Date.split("T")[0];
      if (!dateMap[date]) dateMap[date] = {};

      slotKeys.forEach((key) => {
        const slotQty = entry[key]?.qty || 0;
        if (slotQty > 0) {
          if (!dateMap[date][key]) dateMap[date][key] = 0;
          dateMap[date][key] += slotQty;
        }
      });
    }

    const dateKeys = Object.keys(dateMap).sort();
    let colIndex = 3;

    // Create header columns with merged date and slot abb
    for (const date of dateKeys) {
      const slots = slotKeys.filter((key) => dateMap[date][key] !== undefined);
      const startCol = colIndex;
      const endCol = colIndex + slots.length - 1;

      if (startCol !== endCol) {
        worksheet.mergeCells(headerRow1.number, startCol, headerRow1.number, endCol);
      }

      worksheet.getCell(headerRow1.number, startCol).value = formatShortDate(date);

      slots.forEach((slotKey, i) => {
        const meta = req.body[slotKey];
        headerRow2.getCell(colIndex + i).value = meta?.abb || slotKey;
      });

      colIndex += slots.length;
    }

    // Build participant → date → slot qty map
    const personMap = {};
    for (const entry of data) {
      const key = entry.name;
      if (!personMap[key]) personMap[key] = { role: entry.role || "", entries: {} };
      const date = entry.Date.split("T")[0];
      if (!personMap[key].entries[date]) personMap[key].entries[date] = {};
      slotKeys.forEach((slotKey) => {
        const qty = entry[slotKey]?.qty;
        if (qty) personMap[key].entries[date][slotKey] = qty;
      });
    }

    const startRowNumber = worksheet.lastRow.number + 1;

    // Add each person's data row
    for (const [name, { role, entries }] of Object.entries(personMap)) {
      const rowValues = [name, role];
      for (const date of dateKeys) {
        for (const slotKey of slotKeys) {
          const value = entries[date]?.[slotKey] || "";
          rowValues.push(value);
        }
      }
      worksheet.addRow(rowValues);
    }

    const lastRowNumber = worksheet.lastRow.number;

    // Totals row with formulas
    const totalsRow = worksheet.addRow(["TOTAL", ""]);
    let formulaCol = 3;
    dateKeys.forEach(() => {
      slotKeys.forEach(() => {
        const colLetter = worksheet.getColumn(formulaCol).letter;
        totalsRow.getCell(formulaCol).value = {
          formula: `SUM(${colLetter}${startRowNumber}:${colLetter}${lastRowNumber})`,
        };
        formulaCol++;
      });
    });
    totalsRow.font = { bold: true };

    // Formatting
    worksheet.eachRow((row) => {
      row.eachCell((cell, colNumber) => {
        if (colNumber > 2) {
          cell.alignment = { vertical: "middle", horizontal: "center" };
        }
      });
    });

    worksheet.columns.forEach((col, i) => {
      col.width = i < 2 ? 15 : 6;
    });

    // Upload to Firebase
    const storage = getStorage();
    const bucket = storage.bucket();
    const generatedDate = formatLongDate(new Date()).replace(/\s+/g, "_");
    const fileName = `pivots/${eventName.replace(/\s+/g, "_")}_pivot_${generatedDate}.xlsx`;
    const file = bucket.file(fileName);
    const buffer = await workbook.xlsx.writeBuffer();
    const token = uuidv4();

    await file.save(buffer, {
      metadata: {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        metadata: { firebaseStorageDownloadTokens: token },
      },
    });

    const url = `https://firebasestorage.googleapis.com/v0/b/${bucket.name}/o/${encodeURIComponent(fileName)}?alt=media&token=${token}`;
    return res.json({ url });
  } catch (err) {
    console.error("❌ mealsPivotHandler error:", err);
    return res.status(500).json({ error: "Failed to generate pivot" });
  }
}