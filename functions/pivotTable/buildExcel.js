// functions/pivotTable/buildExcel.js
import ExcelJS from "exceljs";
import { format } from "date-fns";

/**
 * Build the Excel and write it to outputPath.
 *
 * @param {Object} opts
 * @param {string} opts.outputPath
 * @param {string} opts.eventName
 * @param {string[]} opts.allDates           // ISO strings, ascending
 * @param {string[]} opts.dateLabels         // friendly labels (EEE d MMM)
 * @param {Array<{slot:number, abb:string, name:string, location?:string}>} opts.sortedSlots
 * @param {Map<string,string>} opts.descByIso
 * @param {Map<string,{name:string,company:string,role:string}>} opts.peopleMap
 * @param {Object} opts.pivot               // personId -> dateISO -> { [abb]: qty }
 * @param {string[]} opts.accommodatedIds
 * @param {string[]} opts.otherIds
 * @param {number} opts.descFontSize
 */
export async function buildExcel({
  outputPath,
  eventName,
  allDates,
  dateLabels,
  sortedSlots,
  descByIso,
  peopleMap,
  pivot,
  accommodatedIds,
  otherIds,
  descFontSize = 10
}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Meals");

  // Light grey fill for Key section
  const greyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };

  // Title / subtitle / generated-at / spacer
  sheet.addRow([eventName]).font = { bold: true };
  sheet.addRow(["Catering Grid"]);
  const generatedAt = format(new Date(), "EEEE d MMM yyyy, h:mm a");
  const genRow = sheet.addRow([`Generated ${generatedAt}`]);
  genRow.font = { bold: true };
  sheet.addRow([]);

  // Row 5: Name/Company/Role + merged date headers
  const headerRowValues = ["Name", "Company", "Role"];
  const mergeRanges = [];
  let col = 4;
  for (const lbl of dateLabels) {
    const start = col;
    const end = col + sortedSlots.length - 1;
    headerRowValues.push(lbl);
    for (let i = 1; i < sortedSlots.length; i++) headerRowValues.push("");
    mergeRanges.push([start, end]);
    col += sortedSlots.length;
  }
  // Totals column header
  headerRowValues.push("Total");
  const headerRowNum = 5;
  headerRowValues.forEach((v, idx) => {
    const cell = sheet.getCell(headerRowNum, idx + 1);
    cell.value = v;
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  for (const [start, end] of mergeRanges) {
    sheet.mergeCells(5, start, 5, end);
  }

  // Row 6: description row (merged per-date, centered, wrap)
  const totalCols = 3 + allDates.length * sortedSlots.length + 1;
  const descRow = sheet.addRow(new Array(totalCols).fill(""));
  const descRowNum = descRow.number;

  let c = 4;
  for (const iso of allDates) {
    const start = c;
    const end = c + sortedSlots.length - 1;
    sheet.mergeCells(descRowNum, start, descRowNum, end);
    const tl = sheet.getCell(descRowNum, start);
    tl.value = (descByIso.get(iso) || "");
    tl.font = { size: descFontSize };
    tl.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    for (let cc = start + 1; cc <= end; cc++) {
      const cell = sheet.getCell(descRowNum, cc);
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.font = { size: descFontSize };
    }
    c = end + 1;
  }
  descRow.height = 42;

  // Row 7: meal abbreviations
  const slotRow = ["", "", ""];
  for (let i = 0; i < allDates.length; i++) {
    for (const s of sortedSlots) slotRow.push(s.abb);
  }
  slotRow.push("Total");
  const slotRowObj = sheet.addRow(slotRow);
  slotRowObj.font = { bold: true };
  slotRowObj.alignment = { vertical: "middle", horizontal: "center" };

  // Section helper
  const sectionHeaderRows = [];
  const addSection = (title, ids) => {
    if (!ids.length) return;
    const r = sheet.addRow([title]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    for (let col = 2; col <= sheet.columnCount; col++) {
      const cell = r.getCell(col);
      if (cell.value === undefined) cell.value = '';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    }
    sectionHeaderRows.push(r.number);

    for (const pid of ids) {
      const p = peopleMap.get(pid) || { name: pid, company: "", role: "" };
      const vals = [p.name, p.company, p.role];
      for (const d of allDates) {
        const meals = (pivot[pid] && pivot[pid][d]) ? pivot[pid][d] : {};
        for (const { abb } of sortedSlots) vals.push(meals[abb] || "");
      }
      {
        // Append per-person Total column with SUM across meal cells
        const firstDataCol = 4;
        const lastDataCol = 3 + allDates.length * sortedSlots.length;
        const totalColIdx = lastDataCol + 1;
        const row = sheet.addRow(vals.concat([""]));
        const firstLetter = sheet.getColumn(firstDataCol).letter;
        const lastLetter = sheet.getColumn(lastDataCol).letter;
        row.getCell(totalColIdx).value = { formula: `SUM(${firstLetter}${row.number}:${lastLetter}${row.number})` };
        row.getCell(totalColIdx).font = { bold: true };
        row.getCell(totalColIdx).alignment = { horizontal: "center", vertical: "middle" };
      }
    }
  };

  addSection("Accommodated", accommodatedIds);
  addSection("Others",        otherIds);

  // Totals row (SUM down each meal column)
  const lastBeforeTotal = sheet.lastRow.number;
  const totalRow = sheet.getRow(lastBeforeTotal + 1);
  totalRow.getCell(1).value = "TOTAL";
  totalRow.getCell(1).font = { bold: true };
  totalRow.getCell(1).alignment = { vertical: "middle" };
  totalRow.getCell(2).value = "";
  totalRow.getCell(3).value = "";

  const startDataRow = 9; // first person/section row after desc+slot rows
  const endDataRow = sheet.lastRow.number - 1;
  for (let colIdx = 4; colIdx <= sheet.columnCount; colIdx++) {
    const colLetter = sheet.getColumn(colIdx).letter;
    totalRow.getCell(colIdx).value = { formula: `SUM(${colLetter}${startDataRow}:${colLetter}${endDataRow})` };
    totalRow.getCell(colIdx).font = { bold: true };
    totalRow.getCell(colIdx).alignment = { horizontal: "center", vertical: "middle" };
  }
  totalRow.commit();

  // Center all data cells except first 3 cols; skip section headers
  for (let r = 9; r <= sheet.lastRow.number; r++) {
    if (sectionHeaderRows.includes(r)) continue;
    const row = sheet.getRow(r);
    for (let colIdx = 4; colIdx <= sheet.columnCount; colIdx++) {
      const cell = row.getCell(colIdx);
      cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
    }
  }

  // Borders (no protruding above date headers)
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const mealAbbRowNum = 7;
  const mealAbbRow = sheet.getRow(mealAbbRowNum);
  for (let colIdx = 1; colIdx <= sheet.columnCount; colIdx++) {
    mealAbbRow.getCell(colIdx).border = { ...(mealAbbRow.getCell(colIdx).border || {}), bottom: thin };
  }
  const lastDataRowNum = sheet.lastRow.number - 1;
  const lastDataRow = sheet.getRow(lastDataRowNum);
  for (let colIdx = 1; colIdx <= sheet.columnCount; colIdx++) {
    lastDataRow.getCell(colIdx).border = { ...(lastDataRow.getCell(colIdx).border || {}), bottom: thin };
  }

  // Vertical left border for each date group (from row 5 downwards)
  const firstRowWithSlots = 5;
  const lastRowWithTotals = sheet.lastRow.number;
  let startCol = 4;
  for (let i = 0; i < allDates.length; i++) {
    for (let r = firstRowWithSlots; r <= lastRowWithTotals; r++) {
      const cell = sheet.getRow(r).getCell(startCol);
      cell.border = { ...(cell.border || {}), left: thin };
    }
    startCol += sortedSlots.length;
  }

  // Right border for Totals column
  const totalColIdx = sheet.columnCount;
  for (let r = firstRowWithSlots; r <= lastRowWithTotals; r++) {
    const cell = sheet.getRow(r).getCell(totalColIdx);
    cell.border = { ...(cell.border || {}), right: { style: 'medium', color: { argb: 'FF000000' } } };
  }

  // Left border for Totals column
  for (let r = firstRowWithSlots; r <= lastRowWithTotals; r++) {
    const cell = sheet.getRow(r).getCell(totalColIdx);
    cell.border = { ...(cell.border || {}), left: { style: 'medium', color: { argb: 'FF000000' } }, ...(cell.border || {}) };
  }

  // Outer border
  const medium = { style: 'medium', color: { argb: 'FF000000' } };
  const topRow = 5, bottomRow = sheet.lastRow.number, leftCol = 1, rightCol = sheet.columnCount;
  for (let colIdx = leftCol; colIdx <= rightCol; colIdx++) {
    sheet.getRow(topRow).getCell(colIdx).border = { ...(sheet.getRow(topRow).getCell(colIdx).border || {}), top: medium };
    sheet.getRow(bottomRow).getCell(colIdx).border = { ...(sheet.getRow(bottomRow).getCell(colIdx).border || {}), bottom: medium };
  }
  for (let r = topRow; r <= bottomRow; r++) {
    sheet.getRow(r).getCell(leftCol).border  = { ...(sheet.getRow(r).getCell(leftCol).border || {}),  left: medium };
    sheet.getRow(r).getCell(rightCol).border = { ...(sheet.getRow(r).getCell(rightCol).border || {}), right: medium };
  }

  // Column widths
  sheet.columns.forEach((col, idx) => {
    if (idx < 3) {
      col.width = 20;
    } else if (idx === sheet.columns.length - 1) {
      col.width = 5; // Totals column width
    } else {
      col.width = 3;
    }
  });

  // Key table
  const keyStartRow = sheet.lastRow.number + 2;
  sheet.getRow(keyStartRow).getCell(1).value = "Key";
  sheet.getRow(keyStartRow).getCell(1).font = { bold: true };
  for (let colIdx = 1; colIdx <= 12; colIdx++) {
    sheet.getRow(keyStartRow).getCell(colIdx).fill = greyFill;
  }

  const keyHeader = sheet.getRow(keyStartRow + 1);
  keyHeader.getCell(1).value = "Meal";
  keyHeader.getCell(2).value = "Abbreviation";
  keyHeader.getCell(3).value = "Location";
  for (let colIdx = 1; colIdx <= 12; colIdx++) keyHeader.getCell(colIdx).fill = greyFill;
  keyHeader.getCell(1).font = { italic: true };
  keyHeader.getCell(2).font = { italic: true };
  keyHeader.getCell(3).font = { italic: true };

  sortedSlots.forEach((slot, i) => {
    const r = sheet.getRow(keyStartRow + 2 + i);
    r.getCell(1).value = slot.name;
    r.getCell(2).value = slot.abb;
    r.getCell(3).value = slot.location;
    [1,2,3].forEach(ci => r.getCell(ci).fill = greyFill);
    sheet.mergeCells(keyStartRow + 2 + i, 3, keyStartRow + 2 + i, 12);
  });

  // Border around key
  const keyTop = keyStartRow;
  const keyBottom = keyStartRow + 1 + sortedSlots.length;
  const keyLeft = 1, keyRight = 12;
  for (let colIdx = keyLeft; colIdx <= keyRight; colIdx++) {
    sheet.getRow(keyTop).getCell(colIdx).border = { ...(sheet.getRow(keyTop).getCell(colIdx).border || {}), top: medium };
    sheet.getRow(keyBottom).getCell(colIdx).border = { ...(sheet.getRow(keyBottom).getCell(colIdx).border || {}), bottom: medium };
  }
  for (let r = keyTop; r <= keyBottom; r++) {
    sheet.getRow(r).getCell(keyLeft).border  = { ...(sheet.getRow(r).getCell(keyLeft).border || {}),  left: medium };
    sheet.getRow(r).getCell(keyRight).border = { ...(sheet.getRow(r).getCell(keyRight).border || {}), right: medium };
  }

  await workbook.xlsx.writeFile(outputPath);
}