// functions/mealsPivot.js
import ExcelJS from "exceljs";
import { getStorage } from "firebase-admin/storage";
import { v4 as uuidv4 } from "uuid";

// Format for generation date
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
        const { eventName, data, slot1, slot2, slot3, slot4 } = req.body;
        if (!eventName || !Array.isArray(data)) {
            return res.status(400).json({ error: "Missing eventName or data" });
        }

        // Collect slots in order
        const slots = [
            { key: "slot1", ...slot1 },
            { key: "slot2", ...slot2 },
            { key: "slot3", ...slot3 },
            { key: "slot4", ...slot4 }
        ].filter(slot => slot && slot.abb);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Meals Pivot");

        // Title row
        worksheet.getCell("A1").value = eventName;
        worksheet.getCell("A1").font = { bold: true, size: 14 };

        // Generated date row
        const generatedDate = formatLongDate(new Date());
        worksheet.getCell("A2").value = `Generated ${generatedDate}`;

        // Blank row
        worksheet.addRow([]);

        // Header rows
        const headerRow1 = worksheet.addRow(["Name", "Role"]);
        const headerRow2 = worksheet.addRow(["", ""]);
        headerRow1.font = { bold: true };
        headerRow2.font = { bold: true };

        // Collect unique dates
        const dateKeys = [...new Set(data.map(entry => entry.Date.split("T")[0]))].sort();

        // Column headers
        let colIndex = 3;
        for (const date of dateKeys) {
            const startCol = colIndex;
            const endCol = colIndex + slots.length - 1;

            worksheet.mergeCells(headerRow1.number, startCol, headerRow1.number, endCol);
            worksheet.getCell(headerRow1.number, startCol).value = formatShortDate(date);

            slots.forEach((slot, i) => {
                headerRow2.getCell(colIndex + i).value = slot.abb;
            });

            colIndex += slots.length;
        }

        // Group data by person (name+role)
        const grouped = {};
        for (const entry of data) {
            const key = `${entry.name}__${entry.role || ""}`;
            if (!grouped[key]) {
                grouped[key] = { name: entry.name, role: entry.role || "", meals: {} };
            }
            const dateKey = entry.Date.split("T")[0];
            if (!grouped[key].meals[dateKey]) grouped[key].meals[dateKey] = {};
            slots.forEach(slot => {
                if (entry[slot.key]?.qty) {
                    grouped[key].meals[dateKey][slot.key] =
                        (grouped[key].meals[dateKey][slot.key] || 0) + entry[slot.key].qty;
                }
            });
        }

        // Write grouped rows
        for (const personKey in grouped) {
            const person = grouped[personKey];
            const rowValues = [person.name, person.role];
            for (const date of dateKeys) {
                slots.forEach(slot => {
                    const qty = person.meals[date]?.[slot.key] || "";
                    rowValues.push(qty);
                });
            }
            worksheet.addRow(rowValues);
        }

        // Totals row using formulas
        const lastRowNumber = worksheet.lastRow.number;
        const totalsRow = worksheet.addRow(["TOTAL", ""]);
        totalsRow.font = { bold: true };

        let formulaCol = 3;
        for (const date of dateKeys) {
            slots.forEach(() => {
                const colLetter = worksheet.getColumn(formulaCol).letter;
                totalsRow.getCell(formulaCol).value = {
                    formula: `SUM(${colLetter}${headerRow2.number + 1}:${colLetter}${lastRowNumber})`
                };
                formulaCol++;
            });
        }

        // Center align all cells except A/B
        worksheet.eachRow(row => {
            row.eachCell((cell, colNumber) => {
                if (colNumber > 2) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                }
            });
        });

        // Column widths
        worksheet.columns.forEach((col, i) => {
            col.width = i < 2 ? 15 : 6;
        });

        // Upload to Firebase Storage
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