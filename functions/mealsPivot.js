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

    // Add title and generation date
    sheet.addRow([eventName]);
    sheet.addRow([`Generated ${format(new Date(), "EEEE d MMM yyyy")}`]);
    sheet.addRow([]); // Spacer

    // Header Rows
    const dummyDateHeader = ["Name", "Role"]; // Not used for data insertion
    const slotHeader = ["", ""];

    let currentCol = 3;
    for (const date of allDates) {
      const formattedDate = format(parseISO(date), "EEE d MMM");
      const slotCount = sortedSlots.length;
      const start = currentCol;
      const end = currentCol + slotCount - 1;

      // ✅ Merge header range and set visible label
      sheet.mergeCells(4, start, 4, end);
      sheet.getCell(4, start).value = formattedDate;

      // Row 5: meal abbreviations
      for (const { abb } of sortedSlots) {
        slotHeader.push(abb);
      }

      currentCol += slotCount;
    }

    // Row 4 already set manually, Row 5 is slot header
    sheet.addRow(dummyDateHeader); // Row 4 placeholder for alignment
    const slotHeaderRow = sheet.addRow(slotHeader); // Row 5

    slotHeaderRow.font = { bold: true };
    slotHeaderRow.alignment = { vertical: "middle", horizontal: "center" };

    // Data rows start from row 6
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

    // Set column widths
    sheet.columns.forEach((col, idx) => {
      col.width = idx < 2 ? 20 : 6;
    });

    // Save to temp file
    const fileName = `${eventName}_Catering.xlsx`;
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