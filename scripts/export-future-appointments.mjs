import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import XLSX from "xlsx-js-style";

// ─── Configuration ──────────────────────────────────────────────────────────
const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const DEFAULT_CUTOFF_DATE = new Date().toISOString().slice(0, 10); // Today
const DEFAULT_TIMEZONE = "Australia/Melbourne";
const OUTPUT_DIR = "output";
const CALENDAR_FILE = "future-appointments.ics";
const CSV_FILE = "future-appointments.csv";
const XLSX_FILE = "future-appointments.xlsx";

// Only include appointments with this provider (case-insensitive partial match)
// Set to "" or null to include ALL providers
const PROVIDER_FILTER = "chamara basnayake";

// ─── Colour palette ────────────────────────────────────────────────────────
const COLOURS = {
  headerBg: "1B3A5C",
  headerFont: "FFFFFF",
  dateHighlight: "E8F4FD",
  patientInfoBg: "FFF8E7",
  contactBg: "F0FFF0",
  apptDetailBg: "FFF5F5",
  referralOk: "E8F5E9",
  referralExpiring: "FFF3E0",
  referralExpired: "FFEBEE",
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
  console.log(`Looking for appointments after: ${cutoffDate}`);

  const appointments = [];
  let fileCount = 0;
  let patientFileCount = 0;

  // ─── Process files one-by-one using a generator for low memory ───────────
  for await (const filePath of walkDirectory(inputDir)) {
    fileCount++;
    if (fileCount % 500 === 0) {
      console.log(`  Processed ${fileCount} files... (${appointments.length} appointments found)`);
      // Optional: Manually trigger GC if user runs with --expose-gc
      if (global.gc) global.gc();
    }

    if (!filePath.toLowerCase().endsWith(".xml")) continue;

    // Read and check if it's a patient file
    try {
      const xml = await fsp.readFile(filePath, "utf8");
      if (xml.includes("<patient_summary>")) {
        patientFileCount++;
        const found = extractFutureAppointmentsFromXml(xml, filePath, cutoffDate);
        if (found.length > 0) {
          appointments.push(...found);
        }
      }
    } catch (err) {
      console.error(`  Error reading ${filePath}: ${err.message}`);
    }
  }

  // Final sort
  appointments.sort((left, right) => {
    if (left.startKey !== right.startKey) return left.startKey.localeCompare(right.startKey);
    return left.patientName.localeCompare(right.patientName);
  });

  await fsp.mkdir(outputDir, { recursive: true });

  const calendarPath = path.join(outputDir, CALENDAR_FILE);
  const csvPath = path.join(outputDir, CSV_FILE);
  const xlsxPath = path.join(outputDir, XLSX_FILE);

  console.log(`\nGenerating output files...`);
  await fsp.writeFile(calendarPath, buildIcsCalendar(appointments), "utf8");
  await fsp.writeFile(csvPath, buildCsv(appointments), "utf8");
  buildExcel(appointments, xlsxPath);

  console.log(`Finished.`);
  console.log(`  Total files scanned  : ${fileCount}`);
  console.log(`  Patient files found  : ${patientFileCount}`);
  console.log(`  Future appointments  : ${appointments.length}`);
  console.log(`  Excel saved to       : ${xlsxPath}`);

  if (appointments.length > 0 && appointments.length < 50) {
    console.log("\nRecent findings:");
    for (const a of appointments.slice(0, 10)) {
      console.log(`  • ${a.startDate} ${a.startTime} — ${a.patientName}`);
    }
  }
}

// ─── Memory-Efficient Directory Walker ──────────────────────────────────────

async function* walkDirectory(dir) {
  const list = await fsp.readdir(dir, { withFileTypes: true });
  for (const entry of list) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      yield* walkDirectory(fullPath);
    } else {
      yield fullPath;
    }
  }
}

// ─── XML Parsing ────────────────────────────────────────────────────────────

function extractFutureAppointmentsFromXml(xml, filePath, cutoffDate) {
  const pf = extractFirstTagFields(xml, "patient");
  const patientName = getPatientName(pf) || path.basename(filePath, ".xml");
  const appointmentBlocks = extractTagBlocks(xml, "appt");
  const futureAppointments = [];

  const referralDate = pf.referraldate || "";
  const referralDurationMonths = parseInt(pf.referralduration || "0", 10);
  const referralExpiry = calculateReferralExpiry(referralDate, referralDurationMonths);
  const referralStatus = getReferralStatus(referralExpiry);

  for (const block of appointmentBlocks) {
    const fields = extractDirectChildFields(block);
    const startDate = fields.startdate || fields.apptstartdate;
    const startTime = normalizeTime(fields.starttime || fields.apptstarttime);

    if (!startDate || !startTime || startDate <= cutoffDate) continue;
    if (normalizeBoolean(fields.cancelled) || normalizeBoolean(fields.dna)) continue;

    // Filter by provider — only include appointments with Dr Chamara Basnayake
    const providerName = dec(fields.providername || fields.apptprovidername || "");
    if (PROVIDER_FILTER && !providerName.toLowerCase().includes(PROVIDER_FILTER.toLowerCase())) continue;

    const durationSeconds = parseDurationSeconds(fields.apptduration);
    const dateTimeKey = `${startDate}T${startTime}`;

    const apptReferralStatus = referralExpiry && startDate > referralExpiry ? "EXPIRED" : referralStatus;

    futureAppointments.push({
      appointmentId: fields.id || fields.uuid || `${path.basename(filePath)}:${dateTimeKey}`,
      patientName,
      dateOfBirth: pf.dob || "",
      mobilePhone: dec(pf.mobilephone || pf.homephone || ""),
      email: dec(pf.emailaddress || ""),
      suburb: dec(pf.suburb || ""),
      referringDoctor: dec(pf.referringdoctor || ""),
      referrerProviderNum: dec(pf.referrerprovidernum || ""),
      referralDate,
      referralDurationMonths,
      referralExpiry,
      referralStatus: apptReferralStatus,
      usualGP: dec(pf.usualgp || ""),
      accountType: dec(pf.accounttype || ""),
      providerName: dec(fields.providername || fields.apptprovidername || ""),
      reason: dec(fields.reason || fields.apptreason || ""),
      note: dec(fields.note || fields.apptname || fields.name || ""),
      startDate,
      startTime,
      durationSeconds,
      durationMinutes: Math.round(durationSeconds / 60),
      startKey: dateTimeKey,
      sourceFile: path.basename(filePath),
    });
  }
  return futureAppointments;
}

// ... rest of the helper functions from the previous implementation ...
// (I will repeat them here to ensure the file is complete and working)

function calculateReferralExpiry(referralDate, durationMonths) {
  if (!referralDate || referralDate.startsWith("0000") || !durationMonths) return "";
  const [y, m, d] = referralDate.split("-").map(Number);
  const expiry = new Date(y, m - 1 + durationMonths, d);
  return `${expiry.getFullYear()}-${pad(expiry.getMonth() + 1)}-${pad(expiry.getDate())}`;
}

function getReferralStatus(expiryDate) {
  if (!expiryDate) return "NO REFERRAL";
  const today = new Date().toISOString().slice(0, 10);
  if (expiryDate < today) return "EXPIRED";
  const [y, m, d] = expiryDate.split("-").map(Number);
  const expiryMs = new Date(y, m - 1, d).getTime();
  const todayMs = new Date().setHours(0, 0, 0, 0);
  const daysLeft = Math.round((expiryMs - todayMs) / (1000 * 60 * 60 * 24));
  if (daysLeft <= 30) return "EXPIRING SOON";
  return "VALID";
}

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

function getDayOfWeek(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-").map(Number);
  return ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][new Date(y, m - 1, d).getDay()];
}

// ═══════════════════════════════════════════════════════════════════════════
//  EXCEL GENERATION (unchanged logic, just ensuring consistency)
// ═══════════════════════════════════════════════════════════════════════════

const MAIN_COL_COUNT = 16;

function buildExcel(appointments, outputPath) {
  const wb = XLSX.utils.book_new();
  const mainData = buildMainSheetData(appointments);
  const mainSheet = XLSX.utils.aoa_to_sheet(mainData);
  applyMainSheetStyles(mainSheet, mainData);
  XLSX.utils.book_append_sheet(wb, mainSheet, "Future Appointments");

  const summaryData = buildSummaryByDate(appointments);
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  applySummaryStyles(summarySheet, summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Daily Summary");

  const contactData = buildContactList(appointments);
  const contactSheet = XLSX.utils.aoa_to_sheet(contactData);
  applyContactStyles(contactSheet, contactData);
  XLSX.utils.book_append_sheet(wb, contactSheet, "Patient Contacts & Referrals");
  XLSX.writeFile(wb, outputPath);
}

function buildMainSheetData(appointments) {
  const rows = [];
  const emptyRow = new Array(MAIN_COL_COUNT).fill("");
  rows.push(["FUTURE APPOINTMENTS — PRACTICE TRANSITION", ...new Array(MAIN_COL_COUNT - 1).fill("")]);
  const sRow = [...emptyRow];
  sRow[0] = `Generated: ${new Date().toLocaleDateString("en-AU")} at ${new Date().toLocaleTimeString("en-AU")}`;
  sRow[5] = `Total: ${appointments.length} appointments`;
  rows.push(sRow);
  rows.push(emptyRow);
  rows.push([
    "#", "DAY", "DATE", "TIME", "DURATION",
    "PATIENT NAME", "DOB", "MOBILE PHONE", "EMAIL",
    "REASON", "PROVIDER",
    "REFERRING DOCTOR", "REFERRAL DATE", "REFERRAL EXPIRY", "REFERRAL STATUS",
    "ACCOUNT TYPE",
  ]);
  let rowNum = 0;
  let currentDate = "";
  for (const a of appointments) {
    if (a.startDate !== currentDate) { if (currentDate !== "") rows.push(emptyRow); currentDate = a.startDate; }
    rowNum++;
    rows.push([
      rowNum, getDayOfWeek(a.startDate), formatDateForDisplay(a.startDate), formatTimeForDisplay(a.startTime), `${a.durationMinutes} min`,
      a.patientName, formatDateForDisplay(a.dateOfBirth), a.mobilePhone, a.email, a.reason || "—", a.providerName,
      a.referringDoctor, formatDateForDisplay(a.referralDate), formatDateForDisplay(a.referralExpiry), a.referralStatus, a.accountType,
    ]);
  }
  rows.push(emptyRow);
  const fRow = [...emptyRow]; fRow[0] = `END OF LIST — ${appointments.length} total`; rows.push(fRow);
  return rows;
}

function buildSummaryByDate(appointments) {
  const rows = [["DAILY APPOINTMENT SUMMARY"], [], ["DATE", "DAY", "NUMBER OF PATIENTS", "FIRST APPOINTMENT", "LAST APPOINTMENT"]];
  const byDate = {};
  for (const a of appointments) { if (!byDate[a.startDate]) byDate[a.startDate] = []; byDate[a.startDate].push(a); }
  for (const date of Object.keys(byDate).sort()) {
    const dayAppts = byDate[date]; const times = dayAppts.map((a) => a.startTime).sort();
    rows.push([formatDateForDisplay(date), getDayOfWeek(date), dayAppts.length, formatTimeForDisplay(times[0]), formatTimeForDisplay(times[times.length - 1])]);
  }
  rows.push([], [`TOTAL: ${appointments.length} appointments across ${Object.keys(byDate).length} day(s)`]);
  return rows;
}

function buildContactList(appointments) {
  const rows = [["PATIENT CONTACTS & REFERRAL STATUS"], [], ["PATIENT NAME", "DATE OF BIRTH", "MOBILE PHONE", "EMAIL", "SUBURB", "USUAL GP", "REFERRING DOCTOR", "PROVIDER NO.", "REFERRAL DATE", "REFERRAL EXPIRY", "⚠ STATUS", "NEXT APPOINTMENT"]];
  const seen = new Map();
  for (const a of appointments) { if (!seen.has(a.patientName)) seen.set(a.patientName, a); }
  const patients = Array.from(seen.values()).sort((a,b)=>a.patientName.localeCompare(b.patientName));
  for (const p of patients) {
    rows.push([p.patientName, formatDateForDisplay(p.dateOfBirth), p.mobilePhone, p.email, p.suburb, p.usualGP, p.referringDoctor, p.referrerProviderNum, formatDateForDisplay(p.referralDate), formatDateForDisplay(p.referralExpiry), p.referralStatus, `${formatDateForDisplay(p.startDate)} ${formatTimeForDisplay(p.startTime)}`]);
  }
  rows.push([], [`TOTAL: ${patients.length} unique patient(s)`]);
  return rows;
}

// ─── Styles (shortened version of existing) ────────────────────────────────

function applyMainSheetStyles(sheet, data) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  sheet["!cols"] = [{ wch: 5 }, { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 28 }, { wch: 14 }, { wch: 16 }, { wch: 28 }, { wch: 20 }, { wch: 25 }, { wch: 25 }, { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 14 }];
  sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 15 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } }, { s: { r: 1, c: 5 }, e: { r: 1, c: 15 } }];
  styleCell(sheet, "A1", { font: { bold: true, sz: 16, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= 15; c++) styleCell(sheet, XLSX.utils.encode_cell({ r: 3, c }), headerStyle());

  for (let r = 4; r <= range.e.r; r++) {
    const rowData = data[r]; if (!rowData || rowData.every(v => v === "")) continue;
    const isAlt = (r % 2) === 0; const status = rowData[14];
    for (let c = 0; c <= 15; c++) {
      let bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
      if (c >= 1 && c <= 4) bg = COLOURS.dateHighlight;
      else if (c >= 5 && c <= 6) bg = COLOURS.patientInfoBg;
      else if (c >= 7 && c <= 8) bg = COLOURS.contactBg;
      else if (c === 9) bg = COLOURS.apptDetailBg;
      else if (c === 10) bg = COLOURS.contactBg;
      if (c >= 11 && c <= 14) {
        if (status === "EXPIRED") bg = COLOURS.referralExpired;
        else if (status === "EXPIRING SOON") bg = COLOURS.referralExpiring;
        else if (status === "VALID") bg = COLOURS.referralOk;
      }
      styleCell(sheet, XLSX.utils.encode_cell({ r, c }), cellStyle(bg, c === 5 || c === 14, [0,1,2,3,4,12,13,14,15].includes(c), status));
    }
  }
}

function applySummaryStyles(sheet, data) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
  styleCell(sheet, "A1", { font: { bold: true, sz: 14, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= 4; c++) styleCell(sheet, XLSX.utils.encode_cell({ r: 2, c }), headerStyle());
  for (let r = 3; r <= range.e.r; r++) {
    const isAlt = (r % 2) === 0;
    for (let c = 0; c <= 4; c++) styleCell(sheet, XLSX.utils.encode_cell({ r, c }), cellStyle(isAlt ? COLOURS.altRowBg : "FFFFFF", c === 2, true));
  }
}

function applyContactStyles(sheet, data) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }];
  styleCell(sheet, "A1", { font: { bold: true, sz: 14, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= 11; c++) styleCell(sheet, XLSX.utils.encode_cell({ r: 2, c }), headerStyle());
  for (let r = 3; r <= range.e.r; r++) {
    const isAlt = (r % 2) === 0; const status = data[r][10];
    for (let c = 0; c <= 11; c++) {
      let bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
      if (c === 0) bg = COLOURS.patientInfoBg;
      if (c === 2 || c === 3) bg = COLOURS.contactBg;
      if (c === 11) bg = COLOURS.dateHighlight;
      if (c >= 6 && c <= 10) {
        if (status === "EXPIRED") bg = COLOURS.referralExpired;
        else if (status === "EXPIRING SOON") bg = COLOURS.referralExpiring;
        else if (status === "VALID") bg = COLOURS.referralOk;
      }
      styleCell(sheet, XLSX.utils.encode_cell({ r, c }), cellStyle(bg, c === 0 || c === 10, false, status));
    }
  }
}

function styleCell(sheet, addr, style) { if (sheet[addr]) sheet[addr].s = style; }
function headerStyle() { return { fill: { fgColor: { rgb: COLOURS.headerBg } }, font: { color: { rgb: COLOURS.headerFont }, bold: true }, border: thinB(), alignment: { horizontal: "center" } }; }
function cellStyle(bg, bold, center, status) {
  let fc = "333333";
  if (status === "EXPIRED") fc = "C62828";
  else if (status === "EXPIRING SOON") fc = "E65100";
  else if (status === "VALID") fc = "2E7D32";
  return { fill: { fgColor: { rgb: bg } }, font: { bold, color: { rgb: fc }, name: "Calibri" }, border: thinB(), alignment: { vertical: "center", horizontal: center ? "center" : "left" } };
}
function thinB() { return { top: { style: "thin", color: { rgb: COLOURS.borderColour } }, bottom: { style: "thin", color: { rgb: COLOURS.borderColour } }, left: { style: "thin", color: { rgb: COLOURS.borderColour } }, right: { style: "thin", color: { rgb: COLOURS.borderColour } } }; }

// ─── ICS & CSV (unchanged) ──────────────────────────────────────────────────

function buildIcsCalendar(appointments) {
  const lines = [ "BEGIN:VCALENDAR", "VERSION:2.0", "METHOD:PUBLISH", "X-WR-CALNAME:Genie Future Appointments" ];
  for (const a of appointments) {
    lines.push("BEGIN:VEVENT", `SUMMARY:${a.patientName}`, `DTSTART;TZID=${DEFAULT_TIMEZONE}:${a.startDate.replace(/-/g,"")}T${a.startTime.replace(":","")}00`, `DESCRIPTION:${a.reason}`, "END:VEVENT");
  }
  lines.push("END:VCALENDAR"); return lines.join("\r\n");
}

function buildCsv(appointments) {
  const header = [ "patient_name", "date_of_birth", "mobile_phone", "email", "appointment_date", "appointment_time", "reason", "referring_doctor", "referral_expiry", "referral_status" ];
  const rows = appointments.map((a) => [ a.patientName, a.dateOfBirth, a.mobilePhone, a.email, a.startDate, a.startTime, a.reason, a.referringDoctor, a.referralExpiry, a.referralStatus ]);
  return [header, ...rows].map(r => r.map(v => `"${String(v ?? "").replace(/"/g,'""')}"`).join(",")).join("\n");
}

function pad(v) { return String(v).padStart(2, "0"); }

main().catch((error) => { console.error(error.message); process.exitCode = 1; });
