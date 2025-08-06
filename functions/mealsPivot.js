// functions/mealsPivot.js
import ExcelJS from "exceljs";
import { getStorage } from "firebase-admin/storage";
import { v4 as uuidv4 } from "uuid";

// Format for file name and generation note
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
        const { eventName, data } = req.body;
        if (!eventName || !Array.isArray(data)) {
            return res.status(400).json({ error: "Missing eventName or data" });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Meals Pivot");

        // Title row
        worksheet.getCell("A1").value = eventName;
        worksheet.getCell("A1").font = { bold: true, size: 14 };

        // Generation date row
        const generatedDate = formatLongDate(new Date());
        worksheet.getCell("A2").value = `Generated ${generatedDate}`;

        // Blank row
        worksheet.addRow([]);

        // Header rows
        const headerRow1 = worksheet.addRow(["Name", "Role"]);
        const headerRow2 = worksheet.addRow(["", ""]);
        headerRow1.font = { bold: true };
        headerRow2.font = { bold: true };

        // Build date → slots map
        const dateMap = {};
        for (const entry of data) {
            const date = entry.Date.split("T")[0];
            if (!dateMap[date]) dateMap[date] = [];

            ["slot1", "slot2", "slot3", "slot4"].forEach((key) => {
                const slot = entry[key];
                const meta = req.body[key];
                if (slot && meta) {
                    dateMap[date].push({ key, abb: meta.abb, order: parseInt(key.replace("slot", "")) });
                }
            });
        }

        // Deduplicate & sort by slot order
        for (const date in dateMap) {
            dateMap[date] = Array.from(new Map(dateMap[date].map(s => [s.key, s])).values())
                .sort((a, b) => a.order - b.order);
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

            slots.forEach(({ abb }, i) => {
                headerRow2.getCell(colIndex + i).value = abb;
            });

            colIndex += slots.length;
        }

        // Write data rows (grouped by name+role)
        const dataStartRow = worksheet.lastRow.number + 1;

        // Group data by name+role
        const groupedData = {};
        for (const entry of data) {
            const personKey = `${entry.name}||${entry.role || ""}`;
            if (!groupedData[personKey]) {
                groupedData[personKey] = { name: entry.name, role: entry.role || "", dates: {} };
            }
            const date = entry.Date.split("T")[0];
            if (!groupedData[personKey].dates[date]) {
                groupedData[personKey].dates[date] = {};
            }
            ["slot1", "slot2", "slot3", "slot4"].forEach((key) => {
                if (entry[key]?.qty) {
                    groupedData[personKey].dates[date][key] = (groupedData[personKey].dates[date][key] || 0) + entry[key].qty;
                }
            });
        }

        // Write grouped rows
        for (const personKey in groupedData) {
            const person = groupedData[personKey];
            const rowValues = [person.name, person.role];
            for (const date of dateKeys) {
                const slots = dateMap[date];
                slots.forEach(({ key }) => {
                    const qty = person.dates[date]?.[key] || 0;
                    rowValues.push(qty || "");
                });
            }
            worksheet.addRow(rowValues);
        }

        // Totals row with formulas
        const totalsRowValues = ["TOTAL", ""];
        for (let i = 0; i < colIndex - 3; i++) {
            const colLetter = worksheet.getColumn(i + 3).letter;
            totalsRowValues.push({ formula: `SUM(${colLetter}${dataStartRow}:${colLetter}${worksheet.lastRow.number})` });
        }
        const totalsRow = worksheet.addRow(totalsRowValues);
        totalsRow.font = { bold: true };

        // Center align all cells except A and B
        worksheet.eachRow((row) => {
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

        // Add borders
        const firstDataRow = headerRow1.number;
        const lastRow = totalsRow.number;
        const firstMealCol = 3;
        const lastCol = colIndex - 1;

        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (
                    rowNumber >= firstDataRow &&
                    rowNumber <= lastRow &&
                    colNumber >= 1 &&
                    colNumber <= lastCol
                ) {
                    cell.border = {
                        top: rowNumber === firstDataRow + 1 ? { style: "thin" } : cell.border?.top,
                        bottom: rowNumber === lastRow ? { style: "thin" } : cell.border?.bottom,
                        left: colNumber === 1 || colNumber === firstMealCol ? { style: "thin" } : cell.border?.left,
                        right: colNumber === lastCol ? { style: "thin" } : cell.border?.right,
                    };
                }
            });
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