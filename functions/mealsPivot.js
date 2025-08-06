// functions/mealsPivot.js
import ExcelJS from "exceljs";
import { getStorage } from "firebase-admin/storage";
import { v4 as uuidv4 } from "uuid";

// Helper: format generation date nicely
function formatDateWithTime(date) {
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const day = dayNames[date.getUTCDay()];
  const dateNum = date.getUTCDate();

  // Ordinal suffix
  const suffix = (n) => {
    if (n > 3 && n < 21) return "th";
    switch (n % 10) {
      case 1: return "st";
      case 2: return "nd";
      case 3: return "rd";
      default: return "th";
    }
  };

  const month = monthNames[date.getUTCMonth()];
  const year = date.getUTCFullYear();
  const hours = date.getUTCHours().toString().padStart(2, "0");
  const minutes = date.getUTCMinutes().toString().padStart(2, "0");

  return `${day} ${dateNum}${suffix(dateNum)} ${month} ${year} ${hours}_${minutes}`;
}

export async function mealsPivotHandler(req, res) {
  try {
    const { eventName, data } = req.body;

    if (!eventName || !Array.isArray(data)) {
      return res.status(400).json({ error: "Missing eventName or data" });
    }

    // Create workbook & sheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Meals Pivot");

    // Title row
    worksheet.mergeCells("A1:D1");
    worksheet.getCell("A1").value = `Event: ${eventName}`;
    worksheet.getCell("A1").font = { bold: true, size: 14 };

    // Add header rows
    const headerRow1 = worksheet.addRow([]);
    const headerRow2 = worksheet.addRow([]);

    headerRow1.getCell(1).value = "Name";
    headerRow1.getCell(2).value = "Role";
    headerRow2.getCell(1).value = "";
    headerRow2.getCell(2).value = "";

    // Build date → slots map
    const dateMap = {};
    for (const entry of data) {
      const date = entry.Date.split("T")[0]; // YYYY-MM-DD
      if (!dateMap[date]) dateMap[date] = [];

      ["slot1", "slot2", "slot3", "slot4"].forEach(key => {
        const slot = entry[key];
        if (slot && slot.type) {
          dateMap[date].push({
            type: slot.type,
            abb: slot.abb,
            sort: slot.sort,
          });
        }
      });
    }

    // Deduplicate & sort slots for each date
    for (const date in dateMap) {
      dateMap[date] = Array.from(
        new Map(dateMap[date].map(s => [s.type, s])).values()
      ).sort((a, b) => a.sort - b.sort);
    }

    // Write headers
    let colIndex = 3;
    const dateKeys = Object.keys(dateMap).sort();
    for (const date of dateKeys) {
      const slots = dateMap[date];
      const startCol = colIndex;
      const endCol = colIndex + slots.length - 1;
      worksheet.mergeCells(headerRow1.number, startCol, headerRow1.number, endCol);
      worksheet.getCell(headerRow1.number, startCol).value = date;

      slots.forEach((slot, i) => {
        headerRow2.getCell(colIndex + i).value = slot.abb;
      });

      colIndex += slots.length;
    }

    // Write data rows & track totals
    const totals = Array(colIndex - 3).fill(0);

    for (const entry of data) {
      const rowValues = [entry.name, entry.role || ""];
      let slotCounter = 0;
      for (const date of dateKeys) {
        const slots = dateMap[date];
        slots.forEach(slot => {
          const match = Object.values(entry).find(v => v && v.type === slot.type);
          const qty = match?.qty || 0;
          rowValues.push(qty || "");
          totals[slotCounter] += qty;
          slotCounter++;
        });
      }
      worksheet.addRow(rowValues);
    }

    // Totals row
    const totalsRowValues = ["TOTAL", ""];
    totals.forEach(total => totalsRowValues.push(total));
    const totalsRow = worksheet.addRow(totalsRowValues);
    totalsRow.font = { bold: true };

    // Auto-size columns
    worksheet.columns.forEach(column => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, cell => {
        const cellValue = cell.value ? cell.value.toString() : "";
        maxLength = Math.max(maxLength, cellValue.length);
      });
      column.width = maxLength + 2;
    });

    // Upload to Firebase Storage
    const storage = getStorage();
    const bucket = storage.bucket();

    const niceDate = formatDateWithTime(new Date());
    const fileName = `pivots/${eventName.replace(/\s+/g, "_")}_pivot_${niceDate.replace(/\s+/g, "_")}.xlsx`;
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