// mealsPivot.js (refactored)
import { tmpdir } from "os";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { parseISO, format } from "date-fns";
import { getStorage } from "firebase-admin/storage";

// NEW: delegate builders
import { buildExcel } from "./buildExcel.js";
import { buildHtml } from "./buildHtml.js";

// Font size for the Excel description row (adjustable)
const DESC_FONT_SIZE = 10;

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);


export async function mealsPivotHandler(req, res) {
  try {
    const { eventName = "Event", dates = [], slots = [], names = [], tags = [], data = [] } = req.body;

    // --- Basic validation of inbound JSON ---
    if (!Array.isArray(slots) || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid or missing 'slots' or 'data'" });
    }
    if (!Array.isArray(names) || !Array.isArray(tags)) {
      return res.status(400).json({ error: "Invalid or missing 'names' or 'tags'" });
    }
    if (!Array.isArray(dates) || dates.length === 0) {
      return res.status(400).json({ error: "Missing 'dates' array in payload" });
    }

    // --- Sort slots and map slot keys to abbreviations ---
    const sortedSlots = [...slots].sort((a, b) => a.slot - b.slot);
    const slotMap = new Map(); // e.g., slot1 -> B
    for (const { slot, abb } of sortedSlots) slotMap.set(`slot${slot}`, abb);

    // --- Merge people from names & tags into a single directory ---
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
    // Ensure any IDs only present in data are also represented
    for (const r of data) {
      const pid = r?.name; // the id lives in data.name
      if (pid && !peopleMap.has(pid)) {
        peopleMap.set(pid, { name: pid, company: "", role: "" });
      }
    }
    const allPersonIds = Array.from(peopleMap.keys());

    // --- Normalize/Sort dates and build description maps ---
    const normalized = dates
      .filter(d => d && d.date)
      .map(d => ({ date: new Date(d.date).toISOString(), description: d.description || "" }));
    normalized.sort((a, b) => new Date(a.date) - new Date(b.date));

    const allDates = normalized.map(d => d.date); // ISO strings
    const dateLabels = allDates.map(d => format(parseISO(d), "EEE d MMM"));

    const descByIso = new Map(normalized.map(d => [d.date, d.description || ""]));

    // --- Build pivot: personId -> date -> { abb: qty } + track accommodation ---
    const pivot = {};
    const accommodatedByPerson = new Map();

    for (const row of data) {
      const personId = row.name; // person/tag id
      const date = row.Date;
      if (!pivot[personId]) pivot[personId] = {};
      if (!pivot[personId][date]) pivot[personId][date] = {};

      if (row.accommodated === true) accommodatedByPerson.set(personId, true);

      for (const [slotKey, abb] of slotMap.entries()) {
        const qty = row[slotKey];
        if (typeof qty === 'number' && qty > 0) {
          pivot[personId][date][abb] = qty;
        }
      }
    }

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

    // --- Filenames and environment ---
    const ts = Date.now();
    const xlsxFileName = `${eventName}_Catering_${ts}.xlsx`;
    const htmlFileName = `${eventName}_Catering_${ts}.html`;

    const generatedAtText = format(new Date(), "EEEE d MMM yyyy, h:mm a");
    const isRunningLocally = process.env.FUNCTIONS_EMULATOR === 'true';

    // --- 1) Build Excel and persist to disk ---
    let localXlsxPath;
    if (isRunningLocally) {
      // Save to local /output folder for convenience
      const fs = await import('fs');
      const localDir = '/Users/apndavies/Coding/Flair Schedules/output';
      if (!fs.existsSync(localDir)) fs.mkdirSync(localDir, { recursive: true });
      localXlsxPath = `${localDir}/${xlsxFileName}`;
    } else {
      localXlsxPath = join(tmpdir(), xlsxFileName);
    }

    await buildExcel({
      outputPath: localXlsxPath,
      eventName,
      allDates,
      dateLabels,
      sortedSlots,
      descByIso,
      peopleMap,
      pivot,
      accommodatedIds,
      otherIds,
      descFontSize: DESC_FONT_SIZE
    });

    // --- 2) Upload Excel if in cloud to obtain a public URL ---
    let excelHrefForHtml = xlsxFileName; // local case: link to file name in /output alongside HTML
    let excelUrl = null;

    if (!isRunningLocally) {
      const bucket = getStorage().bucket();
      const xlsxDest = `meals/${xlsxFileName}`;
      await bucket.upload(localXlsxPath, {
        destination: xlsxDest,
        metadata: { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
      });
      await bucket.file(xlsxDest).makePublic();
      excelUrl = `https://storage.googleapis.com/${bucket.name}/${xlsxDest}`;
      excelHrefForHtml = excelUrl;
    }

    // --- 3) Build HTML (inject Excel link) and persist ---
    let localHtmlPath;
    if (isRunningLocally) {
      const fs = await import('fs');
      const localDir = '/Users/apndavies/Coding/Flair Schedules/output';
      if (!fs.existsSync(localDir)) fs.mkdirSync(localDir, { recursive: true });
      localHtmlPath = `${localDir}/${htmlFileName}`;
    } else {
      localHtmlPath = join(tmpdir(), htmlFileName);
    }

    const htmlString = await buildHtml({
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
      excelHref: excelHrefForHtml,
      // Allow builder to resolve css/template relative to its own directory
      cssPath: join(__dirname, 'mealsPivot.css'),
      htmlTemplatePath: join(__dirname, 'mealsPivot.html')
    });

    {
      const fs = await import('fs');
      await fs.promises.writeFile(localHtmlPath, htmlString, 'utf8');
    }

    // --- 4) Upload HTML in cloud and return links ---
    if (!isRunningLocally) {
      const bucket = getStorage().bucket();
      const htmlDest = `meals/${htmlFileName}`;
      await bucket.upload(localHtmlPath, { destination: htmlDest, metadata: { contentType: 'text/html' } });
      await bucket.file(htmlDest).makePublic();
      const htmlUrl = `https://storage.googleapis.com/${bucket.name}/${htmlDest}`;
      return res.json({ status: 'success', fileUrl: excelUrl, htmlUrl });
    }

    // Local success payload
    return res.json({ status: 'success', localXlsxPath, localHtmlPath });

  } catch (err) {
    console.error("❌ Error in mealsPivotHandler:", err);
    return res.status(500).json({ error: "Server error", details: err.message });
  }
}