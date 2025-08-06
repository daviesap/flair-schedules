// functions/mealsPivot.js
import ExcelJS from "exceljs";
import { getStorage } from "firebase-admin/storage";
import { v4 as uuidv4 } from "uuid";

// Format for file name and generation note
function formatLongDate(date) {
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const day = dayNames[date.getUTCDay()];
  const dayNum = date.getUTCDate();
  const suffix = (n) => (n > 3 && n < 21 ? "th" : ["st", "nd", "rd"][n % 10 - 1] || "th");
  const month = monthNames[date.getUTCMonth()];
  const year = date.getUTCFullYear();
  return `${day} ${dayNum}${suffix(dayNum)} ${month} ${year}`;
}

function formatShortDate(dateStr) {
  const date = new Date(dateStr);
  return date.toLocaleDateString("en-GB", {
    weekday: "short",
    day: "numeric",
    month: "short",
  }); // e.g. "Wed 6 Aug"
}

export async function mealsPivotHandler(req, res) {
  try {
    const { eventName, data } = req.body;
    if (!eventName || !Array.isArray(data)) {
      return res.status(400).json({ error: "Missing eventName or data" });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Meals Pivot");

    // Title row
    worksheet.mergeCells("A1:D1");
    worksheet.getCell("A1").value = eventName;
    worksheet.getCell("A1").font = { bold: true, size: 14 };

    // Generation date row
    const generatedDate = formatLongDate(new Date());
    worksheet.mergeCells("A2:D2");
    worksheet.getCell("A2").value = `Generated ${generatedDate}`;

    // Blank row
    worksheet.addRow([]);

    // Header rows
    const headerRow1 = worksheet.addRow(["Name", "Role"]);
    const headerRow2 = worksheet.addRow(["", ""]);
    headerRow1.font = { bold: true };
    headerRow2.font = { bold: true };

    // Build date→slots map
    const dateMap = {};
    for (const entry of data) {
      const date = entry.Date.split("T")[0];
      if (!dateMap[date]) dateMap[date] = [];

      ["slot1", "slot2", "slot3", "slot4"].forEach((key) => {
        const slot = entry[key];
        if (slot?.type) {
          dateMap[date].push({ type: slot.type, abb: slot.abb, sort: slot.sort });
        }
      });
    }

    // Deduplicate & sort
    for (const date in dateMap) {
      dateMap[date] = Array.from(new Map(dateMap[date].map(s => [s.type, s])).values())
        .sort((a, b) => a.sort - b.sort);
    }

    // Column headers
    let colIndex = 3;
    const dateKeys = Object.keys(dateMap).sort();
    for (const date of dateKeys) {
      const slots = dateMap[date];
      const startCol = colIndex;
      const endCol = colIndex + slots.length - 1;

      worksheet.mergeCells(headerRow1.number, startCol, headerRow1.number, endCol);
      worksheet.getCell(headerRow1.number, startCol).value = formatShortDate(date);

      slots.forEach((slot, i) => {
        headerRow2.getCell(colIndex + i).value = slot.abb;
      });

      colIndex += slots.length;
    }

    // Write data rows and totals
    const totals = Array(colIndex - 3).fill(0);
    for (const entry of data) {
      const rowValues = [entry.name, entry.role || ""];
      let slotCounter = 0;
      for (const date of dateKeys) {
        const slots = dateMap[date];
        slots.forEach((slot) => {
          const match = Object.values(entry).find((v) => v?.type === slot.type);
          const qty = match?.qty || 0;
          rowValues.push(qty || "");
          totals[slotCounter] += qty;
          slotCounter++;
        });
      }
      worksheet.addRow(rowValues);
    }

    // Totals row
    const totalsRowValues = ["TOTAL", "", ...totals];
    const totalsRow = worksheet.addRow(totalsRowValues);
    totalsRow.font = { bold: true };

    // Column widths
    worksheet.columns.forEach((col, i) => {
      col.width = i < 2 ? 15 : 6; // name/role wider, rest narrow
    });

    // Upload
    const storage = getStorage();
    const bucket = storage.bucket();

    const fileName = `pivots/${eventName.replace(/\s+/g, "_")}_pivot_${generatedDate.replace(/\s+/g, "_")}.xlsx`;
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