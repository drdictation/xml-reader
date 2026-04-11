import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import XLSX from "xlsx-js-style";

// ─── Configuration ──────────────────────────────────────────────────────────
const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const DEFAULT_CUTOFF_DATE = "2026-04-13"; 
const DEFAULT_TIMEZONE = "Australia/Melbourne";
const OUTPUT_DIR = "output";
const XLSX_FILE = "scope-appointments.xlsx";

// Locations to filter for in the <reason> tag
const TARGET_LOCATIONS = [
  "Hobsons Bay DPU",
  "St Vin DPU",
  "Freemasons DPU"
];

// ─── Colour palette ────────────────────────────────────────────────────────
const COLOURS = {
  headerBg: "1B3A5C",
  headerFont: "FFFFFF",
  dateHighlight: "E8F4FD",
  patientInfoBg: "FFF8E7",
  contactBg: "F0FFF0",
  apptDetailBg: "FFF5F5",
  procedureBg: "E1F5FE", // Light blue for procedures
  altRowBg: "F7F9FC",
  borderColour: "B0C4DE",
};

async function main() {
  const inputDir = path.resolve(process.cwd(), process.argv[2] || DEFAULT_XML_DIR);
  const cutoffDate = process.argv[3] || DEFAULT_CUTOFF_DATE;
  const outputDir = path.resolve(process.cwd(), OUTPUT_DIR);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoffDate)) {
    throw new Error(`Invalid cutoff date "${cutoffDate}". Use YYYY-MM-DD.`);
  }

  console.log(`Scanning: ${inputDir}`);
  console.log(`Looking for scope appointments after: ${cutoffDate}`);
  console.log(`Target Locations: ${TARGET_LOCATIONS.join(", ")}`);

  const appointments = [];
  let fileCount = 0;
  let patientFileCount = 0;

  // ─── Process files one-by-one using a generator for low memory ───────────
  for await (const filePath of walkDirectory(inputDir)) {
    fileCount++;
    if (fileCount % 500 === 0) {
      console.log(`  Processed ${fileCount} files... (${appointments.length} appointments found)`);
      if (global.gc) global.gc();
    }

    if (!filePath.toLowerCase().endsWith(".xml")) continue;

    try {
      const xml = await fsp.readFile(filePath, "utf8");
      if (xml.includes("<patient_summary>")) {
        patientFileCount++;
        const found = extractScopeAppointmentsFromXml(xml, filePath, cutoffDate);
        if (found.length > 0) {
          appointments.push(...found);
        }
      }
    } catch (err) {
      console.error(`  Error reading ${filePath}: ${err.message}`);
    }
  }

  // Final sort by date
  appointments.sort((left, right) => {
    if (left.startKey !== right.startKey) return left.startKey.localeCompare(right.startKey);
    return left.patientName.localeCompare(right.patientName);
  });

  await fsp.mkdir(outputDir, { recursive: true });
  const xlsxPath = path.join(outputDir, XLSX_FILE);

  console.log(`\nGenerating Excel file...`);
  buildExcel(appointments, xlsxPath);

  console.log(`Finished.`);
  console.log(`  Total files scanned  : ${fileCount}`);
  console.log(`  Patient files found  : ${patientFileCount}`);
  console.log(`  Scope appointments   : ${appointments.length}`);
  console.log(`  Excel saved to       : ${xlsxPath}`);
}

// ─── Memory-Efficient Directory Walker ──────────────────────────────────────
async function* walkDirectory(dir) {
  try {
    const list = await fsp.readdir(dir, { withFileTypes: true });
    for (const entry of list) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        yield* walkDirectory(fullPath);
      } else {
        yield fullPath;
      }
    }
  } catch (err) {
    console.error(`  Warning: Could not read directory ${dir}: ${err.message}`);
  }
}

// ─── XML Parsing ────────────────────────────────────────────────────────────
function extractScopeAppointmentsFromXml(xml, filePath, cutoffDate) {
  const pf = extractFirstTagFields(xml, "patient");
  const patientName = getPatientName(pf) || path.basename(filePath, ".xml");
  const appointmentBlocks = extractTagBlocks(xml, "appt");
  const scopeAppointments = [];

  for (const block of appointmentBlocks) {
    const fields = extractDirectChildFields(block);
    const startDate = fields.startdate || fields.apptstartdate;
    const startTime = normalizeTime(fields.starttime || fields.apptstarttime);
    const reason = dec(fields.reason || fields.apptreason || "");
    const note = dec(fields.note || fields.apptname || fields.name || "");

    // 1. Check Cutoff Date
    if (!startDate || !startTime || startDate < cutoffDate) continue;

    // 2. Filter by Location (Reason)
    const isTargetLocation = TARGET_LOCATIONS.some(loc => reason.toLowerCase().includes(loc.toLowerCase()));
    if (!isTargetLocation) continue;

    // 3. Identify Case-Insensitive Procedures
    const hasGastro = /gastroscopy/i.test(note) || /gastroscopy/i.test(reason);
    const hasColo = /colonoscopy/i.test(note) || /colonoscopy/i.test(reason);
    
    let procedureType = "Other Scope";
    if (hasGastro && hasColo) procedureType = "Gastroscopy + Colonoscopy";
    else if (hasGastro) procedureType = "Gastroscopy";
    else if (hasColo) procedureType = "Colonoscopy";

    // 4. Ignore cancelled/DNA
    if (normalizeBoolean(fields.cancelled) || normalizeBoolean(fields.dna)) continue;

    const dateTimeKey = `${startDate}T${startTime}`;

    scopeAppointments.push({
      patientName,
      dateOfBirth: pf.dob || "",
      mobilePhone: dec(pf.mobilephone || pf.homephone || ""),
      email: dec(pf.emailaddress || ""),
      suburb: dec(pf.suburb || ""),
      procedureType,
      location: reason,
      note: note,
      startDate,
      startTime,
      durationMinutes: Math.round(parseDurationSeconds(fields.apptduration) / 60),
      referringDoctor: dec(pf.referringdoctor || ""),
      referralDate: pf.referraldate || "",
      startKey: dateTimeKey,
      sourceFile: path.basename(filePath),
    });
  }
  return scopeAppointments;
}

// ─── Helper Functions ───────────────────────────────────────────────────────
function extractFirstTagFields(xml, tagName) {
  const blocks = extractTagBlocks(xml, tagName);
  return blocks.length > 0 ? extractDirectChildFields(blocks[0]) : {};
}

function extractTagBlocks(xml, tagName) {
  const pattern = new RegExp(`<${tagName}>([\\s\\S]*?)<\\/${tagName}>`, "g");
  const blocks = [];
  for (const match of xml.matchAll(pattern)) blocks.push(match[1]);
  return blocks;
}

function extractDirectChildFields(block) {
  const fields = {};
  const pattern = /<([a-zA-Z0-9_]+)>([\s\S]*?)<\/\1>/g;
  for (const match of block.matchAll(pattern)) {
    const tagName = match[1];
    const rawValue = match[2];
    if (/<[a-zA-Z]/.test(rawValue)) continue;
    fields[tagName] = dec(rawValue.trim());
  }
  return fields;
}

function getPatientName(fields) {
  if (fields.fullname) return dec(fields.fullname);
  return [fields.firstname, fields.surname].filter(Boolean).map(dec).join(" ").trim();
}

function normalizeTime(value) {
  if (!value) return "";
  const trimmed = value.trim();
  const timePart = trimmed.includes("T") ? trimmed.split("T")[1] : trimmed;
  if (!/^\d{2}:\d{2}(:\d{2})?$/.test(timePart)) return "";
  return timePart.slice(0, 5);
}

function parseDurationSeconds(value) {
  const parsed = parseInt(value || "", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function normalizeBoolean(value) {
  return /^(true|yes|y|1)$/i.test((value || "").trim());
}

function dec(value) {
  return value
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}

function formatDateForDisplay(isoDate) {
  if (!isoDate || isoDate.startsWith("0000")) return "";
  const [y, m, d] = isoDate.split("-");
  return `${d}/${m}/${y}`;
}

function formatTimeForDisplay(time24) {
  if (!time24) return "";
  const [h, m] = time24.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;
  return `${h12}:${String(m).padStart(2, "0")} ${ampm}`;
}

// ─── Excel Generation ───────────────────────────────────────────────────────
function buildExcel(appointments, outputPath) {
  const wb = XLSX.utils.book_new();
  const data = [
    ["SCOPE PROCEDURES LIST — " + DEFAULT_CUTOFF_DATE + " ONWARDS"],
    [`Generated: ${new Date().toLocaleString("en-AU")}`, "", "", "", "", "", "", "", "", "", ""],
    [],
    [
      "DATE", "TIME", "PROCEDURE TYPE", "PATIENT NAME", "DOB", 
      "MOBILE PHONE", "EMAIL", "LOCATION (DPU)", "REFERRING DOCTOR", 
      "NOTE / DETAILS"
    ]
  ];

  for (const a of appointments) {
    data.push([
      formatDateForDisplay(a.startDate),
      formatTimeForDisplay(a.startTime),
      a.procedureType,
      a.patientName,
      formatDateForDisplay(a.dateOfBirth),
      a.mobilePhone,
      a.email,
      a.location,
      a.referringDoctor,
      a.note
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  applyStyles(ws, data);
  XLSX.utils.book_append_sheet(wb, ws, "Scope Appointments");
  XLSX.writeFile(wb, outputPath);
}

function applyStyles(ws, data) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  ws["!cols"] = [
    { wch: 14 }, { wch: 12 }, { wch: 25 }, { wch: 28 }, { wch: 14 },
    { wch: 16 }, { wch: 28 }, { wch: 20 }, { wch: 25 }, { wch: 50 }
  ];

  // Header styles
  styleCell(ws, "A1", { font: { bold: true, sz: 16, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= 9; c++) {
    styleCell(ws, XLSX.utils.encode_cell({ r: 3, c }), {
      fill: { fgColor: { rgb: COLOURS.headerBg } },
      font: { color: { rgb: COLOURS.headerFont }, bold: true },
      border: thinB(),
      alignment: { horizontal: "center" }
    });
  }

  // Row styles
  for (let r = 4; r <= range.e.r; r++) {
    const isAlt = (r % 2) === 0;
    const bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
    for (let c = 0; c <= 9; c++) {
      let cellBg = bg;
      if (c === 0 || c === 1) cellBg = COLOURS.dateHighlight;
      if (c === 2) cellBg = COLOURS.procedureBg;
      
      styleCell(ws, XLSX.utils.encode_cell({ r, c }), {
        fill: { fgColor: { rgb: cellBg } },
        font: { name: "Calibri", bold: (c === 2 || c === 3) },
        border: thinB(),
        alignment: { vertical: "center", horizontal: [0, 1, 4].includes(c) ? "center" : "left" }
      });
    }
  }
}

function styleCell(sheet, addr, style) { if (sheet[addr]) sheet[addr].s = style; }
function thinB() { return { top: { style: "thin", color: { rgb: COLOURS.borderColour } }, bottom: { style: "thin", color: { rgb: COLOURS.borderColour } }, left: { style: "thin", color: { rgb: COLOURS.borderColour } }, right: { style: "thin", color: { rgb: COLOURS.borderColour } } }; }

main().catch((error) => { console.error(error.message); process.exitCode = 1; });
