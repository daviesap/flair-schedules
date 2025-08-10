// mealsPivot.js
import { tmpdir } from "os";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import ExcelJS from "exceljs";
import { parseISO, format } from "date-fns";
import { getStorage } from "firebase-admin/storage";

// Font size for the Excel description row (adjustable)
const DESC_FONT_SIZE = 10;

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// HTML escape helper
const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const greyFill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFEFEFEF' }
};

export async function mealsPivotHandler(req, res) {
  try {
    const { eventName = "Event", dates = [], slots = [], names = [], tags = [], data = [] } = req.body;

    if (!Array.isArray(slots) || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid or missing 'slots' or 'data'" });
    }
    if (!Array.isArray(names) || !Array.isArray(tags)) {
      return res.status(400).json({ error: "Invalid or missing 'names' or 'tags'" });
    }

    // Require explicit dates[]
    if (!Array.isArray(dates) || dates.length === 0) {
      return res.status(400).json({ error: "Missing 'dates' array in payload" });
    }

    const sortedSlots = [...slots].sort((a, b) => a.slot - b.slot);

    // Map slot keys to abbreviations (e.g., slot1 -> B)
    const slotMap = new Map();
    for (const { slot, abb } of sortedSlots) {
      slotMap.set(`slot${slot}`, abb);
    }

    // Merge people from names and tags into a single directory
    const people = [...names, ...tags];
    const peopleMap = new Map();
    for (const p of people) {
      if (!p || !p.id) continue;
      peopleMap.set(p.id, {
        name: p.name || "Unknown",
        company: p.company || "",
        role: p.role || ""
      });
    }
    // Ensure any IDs only present in data are also represented (fallback name to the ID)
    for (const r of data) {
      const pid = r?.name; // "name" field carries the person/tag id in the new payload
      if (pid && !peopleMap.has(pid)) {
        peopleMap.set(pid, { name: pid, company: "", role: "" });
      }
    }
    const allPersonIds = Array.from(peopleMap.keys());

    // Build ordered date list from the provided dates[] (ascending) and a description map
    const normalized = dates
      .filter(d => d && d.date)
      .map(d => ({ date: new Date(d.date).toISOString(), description: d.description || "" }));
    normalized.sort((a, b) => new Date(a.date) - new Date(b.date));
    const allDates = normalized.map(d => d.date);
    // Map descriptions by YYYY-MM-DD to avoid timezone/offset mismatches
    const descByDateKey = new Map(
      normalized.map(d => [format(parseISO(d.date), "yyyy-MM-dd"), d.description || ""])
    );
    // Also keep an exact ISO -> description map to avoid any timezone key drift
    const descByIso = new Map(normalized.map(d => [d.date, d.description || ""]));

    // Build pivot: personId -> date -> { abb: qty }
    const pivot = {};
    // Track if a person was ever accommodated=true in the dataset
    const accommodatedByPerson = new Map();

    for (const row of data) {
      const personId = row.name; // in the new JSON this is the person/tag id
      //const person = peopleMap.get(personId) || { name: "Unknown", company: "", role: "" };
      const date = row.Date;

      if (!pivot[personId]) pivot[personId] = {};
      if (!pivot[personId][date]) pivot[personId][date] = {};

      // mark accommodated if any row says true
      if (row.accommodated === true) accommodatedByPerson.set(personId, true);

      for (const [slotKey, abb] of slotMap.entries()) {
        const qty = row[slotKey];
        if (typeof qty === 'number' && qty > 0) {
          pivot[personId][date][abb] = qty;
        }
      }
    }

    // Define section membership
    const isAccommodated = (pid) => accommodatedByPerson.get(pid) === true;

    // Sort by company, then name, then role (case-insensitive)
    const cmp = (a, b) => {
      const A = peopleMap.get(a) || { company: "", name: a, role: "" };
      const B = peopleMap.get(b) || { company: "", name: b, role: "" };
      const c1 = (A.company || "").localeCompare(B.company || "", undefined, { sensitivity: 'base' });
      if (c1 !== 0) return c1;
      const c2 = (A.name || "").localeCompare(B.name || "", undefined, { sensitivity: 'base' });
      if (c2 !== 0) return c2;
      return (A.role || "").localeCompare(B.role || "", undefined, { sensitivity: 'base' });
    };

    const accommodatedIds = allPersonIds.filter(isAccommodated).sort(cmp);
    const otherIds = allPersonIds.filter(pid => !isAccommodated(pid)).sort(cmp);

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

    // Row 5: Name + Company + Role + merged date cells
    const headerRow = ["Name", "Company", "Role"];
    const mergeRanges = [];

    let currentCol = 4;
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

    // Row 6: Description row (centered, wrapping)
    // Create a full-width row so merges have backing cells
    const totalColsForDesc = 3 + allDates.length * sortedSlots.length;
    const descRowArr = new Array(totalColsForDesc).fill("");
    const descRowObj = sheet.addRow(descRowArr);
    const descRowNum = descRowObj.number;

    // Helper to obtain description with multiple fallbacks
    const getDescForDate = (iso) => {
      if (!iso) return "";
      const direct = descByIso.get(iso);
      if (typeof direct === "string" && direct.trim().length > 0) return direct;
      const key = format(parseISO(iso), "yyyy-MM-dd");
      const byKey = descByDateKey.get(key);
      if (typeof byKey === "string" && byKey.trim().length > 0) return byKey;
      return "";
    };

    let curCol = 4;
    for (const date of allDates) {
      const desc = getDescForDate(date);
      const startCol = curCol;
      const endCol = curCol + sortedSlots.length - 1;

      // Merge the span for this date block
      sheet.mergeCells(descRowNum, startCol, descRowNum, endCol);

      // Write the value to the top-left cell of the merge
      const tlCell = sheet.getCell(descRowNum, startCol);
      tlCell.value = desc || ""; // leave blank if none
      tlCell.font = { size: DESC_FONT_SIZE };
      tlCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

      // Also apply alignment/font to all cells in the merged region to avoid Excel rendering quirks
      for (let c = startCol + 1; c <= endCol; c++) {
        const cell = sheet.getCell(descRowNum, c);
        // Do not set a value here; merged region takes value from the TL cell
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.font = { size: DESC_FONT_SIZE };
      }

      curCol = endCol + 1;
    }

    // Give the description row a bit more height for readability
    descRowObj.height = 42;

    // Row 7: Meal abbreviations
    const slotRow = ["", "", ""];
    for (let i = 0; i < allDates.length; i++) {
      for (const { abb } of sortedSlots) {
        slotRow.push(abb);
      }
    }
    const slotRowRef = sheet.addRow(slotRow);
    slotRowRef.font = { bold: true };
    slotRowRef.alignment = { vertical: "middle", horizontal: "center" };

    // Row 8+: Section header + person rows
    const sectionHeaderRows = [];

    const writeSection = (title, ids) => {
      if (ids.length === 0) return;
      // Section header row (single cell in column A, no merges)
      const sectionRow = sheet.addRow([title]);
      sectionRow.getCell(1).font = { bold: true };
      sectionRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      // ensure the rest of the row has empty cells so borders apply uniformly later
      for (let c = 2; c <= sheet.columnCount; c++) {
        const cell = sectionRow.getCell(c);
        if (cell.value === undefined) cell.value = '';
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
      // Remember this row so we don't re-center it later
      sectionHeaderRows.push(sectionRow.number);

      // Person rows
      for (const pid of ids) {
        const p = peopleMap.get(pid) || { name: pid, company: "", role: "" };
        const rowVals = [p.name, p.company, p.role];
        for (const date of allDates) {
          const meals = (pivot[pid] && pivot[pid][date]) ? pivot[pid][date] : {};
          for (const { abb } of sortedSlots) {
            rowVals.push(meals[abb] || "");
          }
        }
        sheet.addRow(rowVals);
      }
    };

    writeSection('Accommodated', accommodatedIds);
    writeSection('Others', otherIds);

    // Row for totals
    const lastRowNumber = sheet.lastRow.number + 1;
    const totalRow = sheet.getRow(lastRowNumber);
    totalRow.getCell(1).value = "TOTAL";
    totalRow.getCell(1).font = { bold: true };
    totalRow.getCell(1).alignment = { vertical: "middle" };
    totalRow.getCell(2).value = "";
    totalRow.getCell(2).alignment = { vertical: "middle" };
    totalRow.getCell(3).value = "";
    totalRow.getCell(3).alignment = { vertical: "middle" };

    const startRow = 9; // First row of data (was 8, now 9 due to desc row and meal abb row)
    const endRow = sheet.lastRow.number - 1; // Last data row before total
    const colCount = sheet.columnCount;

    for (let col = 4; col <= colCount; col++) {
      const colLetter = sheet.getColumn(col).letter;
      totalRow.getCell(col).value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
      totalRow.getCell(col).font = { bold: true };
      totalRow.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
    }

    totalRow.commit();

    // Center-align all data and total cells except for columns A, B, C.
    // Skip section header rows so their merged cells remain left-aligned.
    for (let rowNum = 9; rowNum <= sheet.lastRow.number; rowNum++) {
      if (sectionHeaderRows.includes(rowNum)) continue;
      const row = sheet.getRow(rowNum);
      for (let col = 4; col <= sheet.columnCount; col++) {
        const cell = row.getCell(col);
        cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
      }
    }

    // Borders: horizontal below row 7 and below last data row, and vertical left borders for slot groupings
    // Get last data row (before totals)
    const lastDataRow = sheet.lastRow.number - 1;
    // Horizontal border below row 7 (meal abbreviations)
    const borderStyle = { style: 'thin', color: { argb: 'FF000000' } };
    const keyBorderStyle = { style: 'medium', color: { argb: 'FF000000' } };
    // Below row 7
    const mealAbbRowNum = 7;
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
    // Slot groupings start at column 4 (D), then every (number of slots) columns after
    const slotCount = sortedSlots.length;
    const totalDates = allDates.length;
    const slotStartCols = [];
    let colIdx = 4;
    for (let i = 0; i < totalDates; i++) {
      slotStartCols.push(colIdx);
      colIdx += slotCount;
    }
    // For each group start col, apply left border from row 5 to totals row
    const firstRowWithSlots = 5; // start at the date header row to avoid protruding borders above dates
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
    // - The first three columns ("Name", "Company", "Role") are made wider (20 units) to accommodate longer text.
    // - All other columns (meal abbreviations) are narrower (3 units) since they only hold small numeric values.
    sheet.columns.forEach((col, idx) => {
      col.width = idx < 3 ? 20 : 3;
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

    // Build HTML snapshot
    // 1) Friendly date labels
    const dateLabels = allDates.map(d => format(parseISO(d), "EEE d MMM"));
    // 2) Header rows
    const groupHeaderCells = [
      '<th class="sticky name" style="text-align:left;">Name</th>',
      '<th class="sticky company" style="text-align:left;">Company</th>',
      '<th class="sticky role" style="text-align:left;">Role</th>'
    ];
    for (const lbl of dateLabels) {
      groupHeaderCells.push(`<th class="group-header group" colspan="${sortedSlots.length}" >${esc(lbl)}</th>`);
    }
    const slotHeaderCells = ['<th></th>','<th></th>','<th></th>'];
    for (let i = 0; i < allDates.length; i++) {
      for (const { abb } of sortedSlots) slotHeaderCells.push(`<th class="slot-header slot">${esc(abb)}</th>`);
    }
    // 2b) Description headers (one cell per date, spanning all that date's slots)
    const descHeaderCells = ['<th></th>','<th></th>','<th></th>'];
    for (const d of allDates) {
      const desc = descByIso.get(d) || '';
      descHeaderCells.push(`<th class="date-desc" colspan="${sortedSlots.length}"><span class="desc-text">${esc(desc)}</span></th>`);
    }
    // 3) Body rows
    const bodyRows = [];
    const renderPeople = (title, ids) => {
      if (ids.length === 0) return;
      const colspan = 3 + (allDates.length * sortedSlots.length);
      const sectionClass = (title && title.toLowerCase() === 'accommodated')
        ? 'section-header accommodated'
        : 'section-header others';
      bodyRows.push(`<tr class="section-row"><td class="${sectionClass}" colspan="${colspan}">${esc(title)}</td></tr>`);
      for (const pid of ids) {
        const p = peopleMap.get(pid) || { name: pid, company: '', role: '' };
        const cells = [
          `<td class="left">${esc(p.name)}</td>`,
          `<td class="left">${esc(p.company)}</td>`,
          `<td class="left">${esc(p.role)}</td>`
        ];
        for (const date of allDates) {
          const meals = (pivot[pid] && pivot[pid][date]) ? pivot[pid][date] : {};
          for (const { abb } of sortedSlots) {
            const v = meals[abb] ?? '';
            cells.push(`<td class="meal-num num">${esc(v)}</td>`);
          }
        }
        bodyRows.push(`<tr>${cells.join('')}</tr>`);
      }
    };
    renderPeople('Accommodated', accommodatedIds);
    renderPeople('Others', otherIds);
    // 4) Totals row
    const totalsCells = ['<td class="total left meal-total-label" colspan="3">TOTAL</td>'];
    for (let i = 0; i < allDates.length; i++) {
      for (const { abb } of sortedSlots) {
        let sum = 0;
        for (const pid of allPersonIds) {
          const meals = (pivot[pid] && pivot[pid][allDates[i]]) ? pivot[pid][allDates[i]] : {};
          const v = meals[abb];
          if (typeof v === 'number') sum += v;
        }
        totalsCells.push(`<td class="total num meal-total">${sum}</td>`);
      }
    }
    const totalsRowHtml = `<tr class="totals-row">${totalsCells.join('')}</tr>`;
    // 5) Key table
    const keyRowsHtml = sortedSlots.map(s =>
      `<tr class="key-row"><td class="key-meal">${esc(s.name)}</td><td class="key-abb">${esc(s.abb)}</td><td class="key-loc loc">${esc(s.location || '')}</td></tr>`
    ).join('');
    // 6) CSS and HTML doc (CSS can be external in project but is embedded into the final HTML)
    const generatedAtText = format(new Date(), "EEEE d MMM yyyy, h:mm a");

    // Strictly load external CSS; throw if missing or empty to keep output consistent
    const fs = await import("fs");
    const cssPath = join(__dirname, "mealsPivot.css");
    if (!fs.existsSync(cssPath)) {
      throw new Error(`CSS file not found at ${cssPath}`);
    }
    let htmlCss = (await fs.promises.readFile(cssPath, "utf8")).trim();
    if (!htmlCss) {
      throw new Error(`CSS file at ${cssPath} is empty`);
    }
    // Ensure .left is styled with left alignment
    if (!/\.left\s*\{[^}]*text-align\s*:\s*left/i.test(htmlCss)) {
      htmlCss += '\n.left { text-align: left; }\n';
    }

    // Strictly load external HTML template; throw if missing/empty
    const htmlPath = join(__dirname, "mealsPivot.html");
    if (!fs.existsSync(htmlPath)) {
      throw new Error(`HTML template not found at ${htmlPath}`);
    }
    let htmlTemplate = (await fs.promises.readFile(htmlPath, "utf8")).trim();
    if (!htmlTemplate) {
      throw new Error(`HTML template at ${htmlPath} is empty`);
    }

    // Replace placeholders (Excel URL filled later per target)
    let htmlPrepared = htmlTemplate
      .replace("/* {{CSS}} */", htmlCss)
      // Replace *all* occurrences of EVENT_NAME (title tag and H1)
      .replace(/{{EVENT_NAME}}/g, esc(eventName))
      // Plain text for generated-at; alignment handled in CSS
      .replace("{{GENERATED_AT}}", esc(generatedAtText))
      .replace("{{GROUP_HEADERS}}", groupHeaderCells.join(''))
      .replace("{{SLOT_HEADERS}}", slotHeaderCells.join(''))
      .replace("{{DESC_HEADERS}}", descHeaderCells.join(''))
      .replace("{{BODY_ROWS}}", bodyRows.join(''))
      .replace("{{TOTALS_ROW}}", totalsRowHtml)
      .replace("{{KEY_ROWS}}", keyRowsHtml);

    // Save outputs
    const ts = Date.now();
    const xlsxFileName = `${eventName}_Catering_${ts}.xlsx`;
    const htmlFileName = `${eventName}_Catering_${ts}.html`;

    const isRunningLocally = process.env.FUNCTIONS_EMULATOR === 'true';

    if (isRunningLocally) {
      const fs = await import('fs');
      const localDir = '/Users/apndavies/Coding/Flair Schedules/output';
      if (!fs.existsSync(localDir)) fs.mkdirSync(localDir);
      const localXlsxPath = `${localDir}/${xlsxFileName}`;
      const localHtmlPath = `${localDir}/${htmlFileName}`;

      await workbook.xlsx.writeFile(localXlsxPath);
      // For the HTML we don’t yet know the Excel URL; use a file path or leave the placeholder
      const htmlWithLocalLink = htmlPrepared.replace("{{EXCEL_URL}}", xlsxFileName);
      await fs.promises.writeFile(localHtmlPath, htmlWithLocalLink, 'utf8');

      console.log(`✅ Files saved locally at ${localDir}`);
      return res.json({ status: "success", localXlsxPath, localHtmlPath });
    } else {
      // Cloud: write to tmp then upload both files to Storage and make public
      const tmpXlsx = join(tmpdir(), xlsxFileName);
      await workbook.xlsx.writeFile(tmpXlsx);

      const bucket = getStorage().bucket();
      const xlsxDest = `meals/${xlsxFileName}`;
      await bucket.upload(tmpXlsx, { destination: xlsxDest, metadata: { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
      await bucket.file(xlsxDest).makePublic();
      const excelUrl = `https://storage.googleapis.com/${bucket.name}/${xlsxDest}`;

      // Now upload HTML with the real Excel URL embedded
      const tmpHtml = join(tmpdir(), htmlFileName);
      const fs = await import('fs');
      const htmlWithUrl = htmlPrepared.replace("{{EXCEL_URL}}", excelUrl);
      await fs.promises.writeFile(tmpHtml, htmlWithUrl, 'utf8');
      const htmlDest = `meals/${htmlFileName}`;
      await bucket.upload(tmpHtml, { destination: htmlDest, metadata: { contentType: 'text/html' } });
      await bucket.file(htmlDest).makePublic();
      const htmlUrl = `https://storage.googleapis.com/${bucket.name}/${htmlDest}`;

      return res.json({ status: 'success', fileUrl: excelUrl, htmlUrl });
    }

  } catch (err) {
    console.error("❌ Error in mealsPivotHandler:", err);
    return res.status(500).json({ error: "Server error", details: err.message });
  }
}