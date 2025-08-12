// functions/pivotTable/buildHtml.js

/**
 * Build the HTML snapshot string.
 *
 * @param {Object} opts
 * @param {string} opts.eventName
 * @param {string[]} opts.dateLabels
 * @param {string[]} opts.allDates           // ISO strings
 * @param {Array<{slot:number, abb:string, name:string, location?:string}>} opts.sortedSlots
 * @param {Map<string,string>} opts.descByIso
 * @param {Map<string,{name:string,company:string,role:string}>} opts.peopleMap
 * @param {Object} opts.pivot               // personId -> dateISO -> { [abb]: qty }
 * @param {string[]} opts.accommodatedIds
 * @param {string[]} opts.otherIds
 * @param {string} opts.generatedAtText
 * @param {string} opts.excelHref           // href to XLSX (filename locally or public URL in cloud)
 * @param {string} opts.cssPath
 * @param {string} opts.htmlTemplatePath
 * @returns {Promise<string>}
 */
export async function buildHtml({
  eventName,
  dateLabels,
  allDates,
  sortedSlots,
  descByIso,
  peopleMap,
  pivot,
  accommodatedIds,
  otherIds,
  generatedAtText,
  excelHref,
  cssPath,
  htmlTemplatePath
}) {
  const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  // Group headers
  const groupHeaderCells = [
    '<th class="sticky name" style="text-align:left;">Name</th>',
    '<th class="sticky company" style="text-align:left;">Company</th>',
    '<th class="sticky role" style="text-align:left;">Role</th>'
  ];
  for (const lbl of dateLabels) {
    groupHeaderCells.push(`<th class="group-header group" colspan="${sortedSlots.length}" >${esc(lbl)}</th>`);
  }
  // Totals column header
  groupHeaderCells.push('<th class="group-header total-col">Total</th>');

  // Slot headers (B/L/D/O/TC...)
  const slotHeaderCells = ['<th></th>','<th></th>','<th></th>'];
  for (let i = 0; i < allDates.length; i++) {
    for (let j = 0; j < sortedSlots.length; j++) {
      const { abb } = sortedSlots[j];
      const extraClass = (j === 0) ? ' first-slot' : '';
      slotHeaderCells.push(`<th class="slot-header slot${extraClass}">${esc(abb)}</th>`);
    }
  }
  // Totals column (slot header row)
  slotHeaderCells.push('<th class="slot-header total-col">Total</th>');

  // Description headers (one per date, spanning all that date's slots)
  const descHeaderCells = ['<th></th>','<th></th>','<th></th>'];
  for (const iso of allDates) {
    const desc = descByIso.get(iso) || '';
    // No first-slot needed for description row since it spans all slots for each date
    descHeaderCells.push(`<th class="date-desc" colspan="${sortedSlots.length}"><span class="desc-text">${esc(desc)}</span></th>`);
  }
  // Extra empty cell for per-person Total column
  descHeaderCells.push('<th class="total-col"></th>');

  // Body
  const bodyRows = [];
  const renderPeople = (title, ids) => {
    if (!ids.length) return;
    const sectionClass = (title && title.toLowerCase() === 'accommodated')
      ? 'section-header accommodated'
      : 'section-header others';
    // Render section header row with 3 <td>s for the first columns, then <td>s for all slots across all dates with appropriate classes
    const cells = [
      `<td class="${sectionClass}">${esc(title)}</td>`,
      `<td class="meal-num"></td>`,
      `<td class="meal-num"></td>`
    ];
    for (let i = 0; i < allDates.length; i++) {
      for (let j = 0; j < sortedSlots.length; j++) {
        const extraClass = (j === 0) ? ' num first-slot' : ' num';
        cells.push(`<td class="meal-num${extraClass}"></td>`);
      }
    }
    // Totals column cell (blank in section header)
    cells.push('<td class="meal-num total-col"></td>');
    bodyRows.push(`<tr class="section-row">${cells.join('')}</tr>`);
    for (const pid of ids) {
      const p = peopleMap.get(pid) || { name: pid, company: '', role: '' };
      const cells = [
        `<td class="left">${esc(p.name)}</td>`,
        `<td class="left">${esc(p.company)}</td>`,
        `<td class="left">${esc(p.role)}</td>`
      ];
      let rowTotal = 0;
      for (let i = 0; i < allDates.length; i++) {
        const iso = allDates[i];
        const meals = (pivot[pid] && pivot[pid][iso]) ? pivot[pid][iso] : {};
        for (let j = 0; j < sortedSlots.length; j++) {
          const { abb } = sortedSlots[j];
          const extraClass = (j === 0) ? ' first-slot' : '';
          const v = meals[abb];
          if (typeof v === 'number') rowTotal += v;
          cells.push(`<td class="meal-num num${extraClass}">${esc(v ?? '')}</td>`);
        }
      }
      // Per-person total cell at far right
      cells.push(`<td class="meal-num num total-col">${rowTotal}</td>`);
      bodyRows.push(`<tr>${cells.join('')}</tr>`);
    }
  };
  renderPeople('Accommodated', accommodatedIds);
  renderPeople('Others',        otherIds);

  // Totals row
  const allIds = [...accommodatedIds, ...otherIds];
  const totalsCells = ['<td class="total left meal-total-label" colspan="3">TOTAL</td>'];
  let grandTotal = 0; // sum of all per-date/slot totals (should match sum of per-person totals column)
  for (let i = 0; i < allDates.length; i++) {
    const iso = allDates[i];
    for (let j = 0; j < sortedSlots.length; j++) {
      const { abb } = sortedSlots[j];
      let sum = 0;
      for (const pid of allIds) {
        const meals = (pivot[pid] && pivot[pid][iso]) ? pivot[pid][iso] : {};
        const v = meals[abb];
        if (typeof v === 'number') sum += v;
      }
      grandTotal += sum;
      const extraClass = (j === 0) ? ' first-slot' : '';
      totalsCells.push(`<td class="total num meal-total${extraClass}">${sum}</td>`);
    }
  }
  // Bottom-right grand total (aligns with per-person Total column)
  totalsCells.push(`<td class="total num total-col grand-total">${grandTotal}</td>`);
  const totalsRowHtml = `<tr class="totals-row">${totalsCells.join('')}</tr>`;

  // Key table
  const keyRowsHtml = sortedSlots.map(s =>
    `<tr class="key-row"><td class="key-meal">${esc(s.name)}</td><td class="key-abb">${esc(s.abb)}</td><td class="key-loc loc">${esc(s.location || '')}</td></tr>`
  ).join('');

  // Load CSS + template strictly
  const fs = await import("fs");
  if (!fs.existsSync(cssPath)) throw new Error(`CSS file not found at ${cssPath}`);
  if (!fs.existsSync(htmlTemplatePath)) throw new Error(`HTML template not found at ${htmlTemplatePath}`);
  let htmlCss = (await fs.promises.readFile(cssPath, "utf8")).trim();
  if (!htmlCss) throw new Error(`CSS file at ${cssPath} is empty`);
  if (!/\.left\s*\{[^}]*text-align\s*:\s*left/i.test(htmlCss)) {
    htmlCss += '\n.left { text-align: left; }\n';
  }
  let template = (await fs.promises.readFile(htmlTemplatePath, "utf8")).trim();
  if (!template) throw new Error(`HTML template at ${htmlTemplatePath} is empty`);

  // Fill template
  const htmlPrepared = template
    .replace('/* {{CSS}} */', htmlCss)
    .replace(/{{EVENT_NAME}}/g, esc(eventName))
    .replace('{{GENERATED_AT}}', esc(generatedAtText))
    .replace('{{GROUP_HEADERS}}', groupHeaderCells.join(''))
    .replace('{{SLOT_HEADERS}}', slotHeaderCells.join(''))
    .replace('{{DESC_HEADERS}}', descHeaderCells.join(''))
    .replace('{{BODY_ROWS}}', bodyRows.join(''))
    .replace('{{TOTALS_ROW}}', totalsRowHtml)
    .replace('{{KEY_ROWS}}', keyRowsHtml)
    .replace('{{EXCEL_URL}}', esc(excelHref));

  return htmlPrepared;
}