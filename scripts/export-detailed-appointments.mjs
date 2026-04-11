import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx-js-style";

// ─── Configuration ──────────────────────────────────────────────────────────
const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const DEFAULT_CUTOFF_DATE = new Date().toISOString().slice(0, 10);
const DEFAULT_TIMEZONE = "Australia/Melbourne";
const OUTPUT_DIR = "output";
const CALENDAR_FILE = "future-appointments-detailed.ics";
const CSV_FILE = "future-appointments-detailed.csv";
const XLSX_FILE = "future-appointments-detailed.xlsx";

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
  doctorBg: "EDE7F6",       // Light purple for doctor details
  addressBg: "E3F2FD",      // Light blue for address
  scopeBg: "FFF9C4",        // Yellow for scope/procedure
  billingBg: "E8F5E9",      // Green for billing codes
  altRowBg: "F7F9FC",
  borderColour: "B0C4DE",
};

// V8 Memory Leak Fix: Detach sliced strings from their multi-megabyte parent string
function copyStr(str) {
  if (!str) return "";
  return Buffer.from(str, "utf8").toString("utf8");
}

function main() {
  const inputDir = path.resolve(process.cwd(), process.argv[2] || DEFAULT_XML_DIR);
  const cutoffDate = process.argv[3] || DEFAULT_CUTOFF_DATE;
  const outputDir = path.resolve(process.cwd(), OUTPUT_DIR);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoffDate)) {
    throw new Error(`Invalid cutoff date "${cutoffDate}". Use YYYY-MM-DD.`);
  }

  console.log(`\n======================================================`);
  console.log(`STARTING DETAILED XML EXTRACTION (Memory Optimized)`);
  console.log(`Scanning: ${inputDir}`);
  console.log(`Looking for appointments after: ${cutoffDate}`);
  console.log(`======================================================\n`);

  // Get all XML files
  const allFiles = fs.readdirSync(inputDir, { recursive: true })
    .filter(f => f.toLowerCase().endsWith(".xml"))
    .map(f => path.join(inputDir, f));
  console.log(`Found ${allFiles.length} XML files.\n`);

  const checkBuffer = Buffer.alloc(16384);

  // ─── Pass 1: Build Address Book Map ──────────────────────────────────────
  console.log(`Pass 1: Building address book map...`);
  const addressBookMap = new Map();
  let addressBookFileCount = 0;

  for (let i = 0; i < allFiles.length; i++) {
    if (i > 0 && i % 200 === 0) {
      console.log(`  Processed ${i}/${allFiles.length} files...`);
      if (global.gc) global.gc();
    }
    try {
      // Fast check before reading massive string
      const fd = fs.openSync(allFiles[i], "r");
      fs.readSync(fd, checkBuffer, 0, 16384, 0);
      fs.closeSync(fd);
      if (!checkBuffer.toString("utf8").includes("<addressbook")) continue;

      let xml = fs.readFileSync(allFiles[i], "utf8");
      if (xml.indexOf("<addressbook_list>") !== -1) {
        addressBookFileCount++;
        for (const block of extractTagBlocks(xml, "addressbook")) {
          const f = extractDirectChildFields(block);
          if (f.id) {
            addressBookMap.set(f.id, {
              fullName: getFullName(f),
              clinic: f.clinic || "",
              address: [f.address1, f.address2, f.suburb, f.state, f.postcode].filter(Boolean).join(", "),
              phone: f.workphone || f.mobile || "",
              fax: f.fax || "",
              email: f.emailaddress || "",
              providerNum: f.providernum || "",
            });
          }
        }
      }
      xml = null; // Free parent string
    } catch (err) { /* skip */ }
  }
  console.log(`  Found ${addressBookMap.size} entries in ${addressBookFileCount} files.\n`);

  // ─── Pass 2: Extract Appointments ────────────────────────────────────────
  console.log(`Pass 2: Extracting appointments...`);
  const appointments = [];
  let patientFileCount = 0;

  for (let i = 0; i < allFiles.length; i++) {
    if (i > 0 && i % 400 === 0) {
      console.log(`  Processed ${i}/${allFiles.length}... (${appointments.length} isolated appointments)`);
      if (global.gc) global.gc(); // Keep memory lean
    }
    try {
      const fd = fs.openSync(allFiles[i], "r");
      fs.readSync(fd, checkBuffer, 0, 16384, 0);
      fs.closeSync(fd);
      if (!checkBuffer.toString("utf8").includes("<patient_summary")) continue;

      let xml = fs.readFileSync(allFiles[i], "utf8");
      if (xml.indexOf("<patient_summary>") !== -1) {
        patientFileCount++;
        const found = extractFutureAppointmentsFromXml(xml, allFiles[i], cutoffDate, addressBookMap);
        if (found.length > 0) appointments.push(...found);
      }
      xml = null; // Free parent string
    } catch (err) {
      console.error(`  Error reading ${allFiles[i]}: ${err.message}`);
    }
  }

  // Final sort
  appointments.sort((a, b) => {
    if (a.startKey !== b.startKey) return a.startKey.localeCompare(b.startKey);
    return a.patientName.localeCompare(b.patientName);
  });

  fs.mkdirSync(outputDir, { recursive: true });
  const calendarPath = path.join(outputDir, CALENDAR_FILE);
  const csvPath = path.join(outputDir, CSV_FILE);
  const xlsxPath = path.join(outputDir, XLSX_FILE);

  console.log(`\nGenerating output files...`);
  fs.writeFileSync(calendarPath, buildIcsCalendar(appointments), "utf8");
  fs.writeFileSync(csvPath, buildCsv(appointments), "utf8");
  buildExcel(appointments, xlsxPath);

  console.log(`\n======================================================`);
  console.log(`Finished Successfully!`);
  console.log(`  Total files scanned  : ${allFiles.length}`);
  console.log(`  Patient files found  : ${patientFileCount}`);
  console.log(`  Future appointments  : ${appointments.length}`);
  console.log(`  Excel saved to       : ${xlsxPath}`);
  console.log(`======================================================\n`);
}

// ─── XML Parsing ────────────────────────────────────────────────────────────

function extractFutureAppointmentsFromXml(xml, filePath, cutoffDate, addressBookMap) {
  const patientBlocks = extractTagBlocks(xml, "patient");
  if (patientBlocks.length === 0) return [];
  const pf = extractDirectChildFields(patientBlocks[0]);
  const patientName = getPatientName(pf) || copyStr(path.basename(filePath, ".xml"));
  const appointmentBlocks = extractTagBlocks(xml, "appt");
  const futureAppointments = [];

  const referralDate = pf.referraldate || "";
  const referralDurationMonths = parseInt(pf.referralduration || "0", 10);
  const referralExpiry = calculateReferralExpiry(referralDate, referralDurationMonths);
  const referralStatus = getReferralStatus(referralExpiry);

  // ─── Patient full address ───────────────────────────────────────────
  const fullAddress = [pf.addressline1, pf.addressline2, pf.suburb, pf.state, pf.postcode]
    .filter(Boolean).map(dec).join(", ");

  // ─── Referring doctor details from address book ─────────────────────
  const refId = pf.referrer_ab_id_fk || "";
  const refDetails = addressBookMap.get(refId) || {
    fullName: dec(pf.referringdoctor || ""),
    clinic: "", address: "", phone: "", fax: "", email: "",
    providerNum: dec(pf.referrerprovidernum || ""),
  };

  for (const block of appointmentBlocks) {
    const fields = extractDirectChildFields(block);
    const startDate = fields.startdate || fields.apptstartdate;
    const startTime = normalizeTime(fields.starttime || fields.apptstarttime);

    if (!startDate || !startTime || startDate <= cutoffDate) continue;
    if (normalizeBoolean(fields.cancelled) || normalizeBoolean(fields.dna)) continue;

    const providerName = dec(fields.providername || fields.apptprovidername || "");
    if (PROVIDER_FILTER && !providerName.toLowerCase().includes(PROVIDER_FILTER.toLowerCase())) continue;

    const durationSeconds = parseDurationSeconds(fields.apptduration);
    const dateTimeKey = `${startDate}T${startTime}`;
    const apptReferralStatus = referralExpiry && startDate > referralExpiry ? "EXPIRED" : referralStatus;
    const rawNote = dec(fields.note || fields.apptname || fields.name || "");
    const rawReason = dec(fields.reason || fields.apptreason || "");

    futureAppointments.push({
      appointmentId: fields.id || fields.uuid || copyStr(`${path.basename(filePath)}:${dateTimeKey}`),
      patientName,
      dateOfBirth: pf.dob || "",
      mobilePhone: dec(pf.mobilephone || pf.homephone || ""),
      email: dec(pf.emailaddress || ""),
      fullAddress,
      suburb: dec(pf.suburb || ""),

      referringDoctor: refDetails.fullName,
      referrerProviderNum: refDetails.providerNum,
      referralDate,
      referralDurationMonths,
      referralExpiry,
      referralStatus: apptReferralStatus,

      refClinic: refDetails.clinic,
      refAddress: refDetails.address,
      refPhone: refDetails.phone,
      refFax: refDetails.fax,
      refEmail: refDetails.email,

      usualGP: dec(pf.usualgp || ""),
      accountType: dec(pf.accounttype || ""),
      providerName,
      reason: rawReason,
      note: rawNote,
      procedureType: parseProcedureType(rawNote, rawReason),  // NEW
      billingCodes: parseBillingCodes(rawNote),                // NEW
      startDate,
      startTime,
      durationSeconds,
      durationMinutes: Math.round(durationSeconds / 60),
      startKey: copyStr(dateTimeKey),
      sourceFile: copyStr(path.basename(filePath)),
    });
  }
  return futureAppointments;
}

// ─── Procedure & Billing Parsers (NEW) ─────────────────────────────────────

// Detects procedure type from the note and reason fields
function parseProcedureType(note, reason) {
  const text = (note + " " + reason).toLowerCase();
  const hasGastro = /gastroscopy|gastroscop/i.test(text);
  const hasColo   = /colonoscopy|colonoscop/i.test(text);
  if (hasGastro && hasColo) return "Gastroscopy + Colonoscopy";
  if (hasGastro) return "Gastroscopy";
  if (hasColo)   return "Colonoscopy";
  return ""; // Regular outpatient appointment
}

// Extracts MBS/billing codes from the note field (e.g. 91825, 132, 133, 110, 116)
function parseBillingCodes(note) {
  // Known patterns: standalone numbers >= 3 digits that look like MBS codes
  // Usually at start of note: "91825 $190.00" or "132  $450.00" or "116 BB"
  const matches = note.match(/\b(9\d{4}|1[0-9]{2,4}|2\d{3,4}|3\d{4})\b/g);
  if (!matches) return "";
  // Deduplicate and filter out year-like numbers
  const codes = [...new Set(matches)].filter(c => {
    const n = parseInt(c, 10);
    return n < 100000; // exclude long numeric strings
  });
  return codes.join(", ");
}

// ─── Helper Functions ───────────────────────────────────────────────────────

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

// Extract subsets without Regex for speed, but returns sliced strings
function extractTagBlocks(xml, tagName) {
  const startTag = `<${tagName}>`;
  const endTag = `</${tagName}>`;
  const blocks = [];
  let pos = 0;
  while (true) {
    const s = xml.indexOf(startTag, pos);
    if (s === -1) break;
    const e = xml.indexOf(endTag, s);
    if (e === -1) break;
    blocks.push(xml.substring(s + startTag.length, e));
    pos = e + endTag.length;
  }
  return blocks;
}

// Crucial Memory Fix: Forces all dictionary strings to be deep copied
// freeing the 50MB parent strings for Garbage Collection.
function extractDirectChildFields(block) {
  const fields = {};
  const pattern = /<([a-zA-Z0-9_]+)>([\s\S]*?)<\/\1>/g;
  for (const match of block.matchAll(pattern)) {
    const tagName = copyStr(match[1]);
    const rawValue = match[2];
    if (rawValue.indexOf("<") !== -1) continue;
    fields[tagName] = dec(copyStr(rawValue.trim()));
  }
  return fields;
}

function getPatientName(fields) {
  if (fields.fullname) return dec(fields.fullname);
  return [fields.firstname, fields.surname].filter(Boolean).map(dec).join(" ").trim();
}

function getFullName(fields) {
  if (fields.fullname) return dec(fields.fullname);
  return [fields.title, fields.firstname, fields.surname].filter(Boolean).map(dec).join(" ").trim();
}

function normalizeTime(value) {
  if (!value) return "";
  const trimmed = value.trim();
  const timePart = trimmed.includes("T") ? trimmed.split("T")[1] : trimmed;
  if (!/^\d{2}:\d{2}(:\d{2})?$/.test(timePart)) return "";
  return copyStr(timePart.slice(0, 5));
}

function parseDurationSeconds(value) {
  const parsed = parseInt(value || "", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function normalizeBoolean(value) {
  return /^(true|yes|y|1)$/i.test((value || "").trim());
}

function dec(value) {
  if (!value) return "";
  return copyStr(value
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'"));
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

function pad(v) { return String(v).padStart(2, "0"); }

// ═══════════════════════════════════════════════════════════════════════════
//  EXCEL GENERATION — Same beautiful styling, expanded columns
// ═══════════════════════════════════════════════════════════════════════════

const MAIN_COL_COUNT = 23; // Added PROCEDURE TYPE (col 10) and BILLING CODE (col 11)

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
  rows.push(["FUTURE APPOINTMENTS — DETAILED EXPORT", ...new Array(MAIN_COL_COUNT - 1).fill("")]);
  const sRow = [...emptyRow];
  sRow[0] = `Generated: ${new Date().toLocaleDateString("en-AU")} at ${new Date().toLocaleTimeString("en-AU")}`;
  sRow[5] = `Total: ${appointments.length} appointments`;
  rows.push(sRow);
  rows.push(emptyRow);
  rows.push([
    "#", "DAY", "DATE", "TIME", "DURATION",
    "PATIENT NAME", "DOB", "MOBILE PHONE", "EMAIL",
    "PATIENT ADDRESS",
    "⚕ PROCEDURE TYPE", "💳 BILLING CODE",                 // NEW
    "REASON", "PROVIDER",
    "REFERRING DOCTOR", "REF CLINIC", "REF ADDRESS",
    "REF PHONE", "REF FAX",
    "REFERRAL DATE", "REFERRAL EXPIRY", "REFERRAL STATUS",
    "ACCOUNT TYPE",
  ]);
  let rowNum = 0;
  let currentDate = "";
  for (const a of appointments) {
    if (a.startDate !== currentDate) { if (currentDate !== "") rows.push(emptyRow); currentDate = a.startDate; }
    rowNum++;
    rows.push([
      rowNum, getDayOfWeek(a.startDate), formatDateForDisplay(a.startDate), formatTimeForDisplay(a.startTime), `${a.durationMinutes} min`,
      a.patientName, formatDateForDisplay(a.dateOfBirth), a.mobilePhone, a.email,
      a.fullAddress,
      a.procedureType || "", a.billingCodes || "",           // NEW
      a.reason || "—", a.providerName,
      a.referringDoctor, a.refClinic, a.refAddress,
      a.refPhone, a.refFax,
      formatDateForDisplay(a.referralDate), formatDateForDisplay(a.referralExpiry), a.referralStatus,
      a.accountType,
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
  const rows = [
    ["PATIENT CONTACTS & REFERRAL STATUS"],
    [],
    [
      "PATIENT NAME", "DATE OF BIRTH", "MOBILE PHONE", "EMAIL",
      "FULL ADDRESS",                                             // NEW
      "SUBURB", "USUAL GP",
      "REFERRING DOCTOR", "REF CLINIC", "REF ADDRESS",           // NEW
      "REF PHONE", "REF FAX",                                    // NEW
      "PROVIDER NO.", "REFERRAL DATE", "REFERRAL EXPIRY",
      "⚠ STATUS", "NEXT APPOINTMENT",
    ],
  ];
  const seen = new Map();
  for (const a of appointments) { if (!seen.has(a.patientName)) seen.set(a.patientName, a); }
  const patients = Array.from(seen.values()).sort((a, b) => a.patientName.localeCompare(b.patientName));
  for (const p of patients) {
    rows.push([
      p.patientName, formatDateForDisplay(p.dateOfBirth), p.mobilePhone, p.email,
      p.fullAddress,                                              // NEW
      p.suburb, p.usualGP,
      p.referringDoctor, p.refClinic, p.refAddress,               // NEW
      p.refPhone, p.refFax,                                       // NEW
      p.referrerProviderNum, formatDateForDisplay(p.referralDate), formatDateForDisplay(p.referralExpiry),
      p.referralStatus,
      `${formatDateForDisplay(p.startDate)} ${formatTimeForDisplay(p.startTime)}`,
    ]);
  }
  rows.push([], [`TOTAL: ${patients.length} unique patient(s)`]);
  return rows;
}

// ─── Styles ────────────────────────────────────────────────────────────────

function applyMainSheetStyles(sheet, data) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const lastCol = MAIN_COL_COUNT - 1; // 22
  sheet["!cols"] = [
    { wch: 5 },  { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 10 },  // #, Day, Date, Time, Duration
    { wch: 28 }, { wch: 14 }, { wch: 16 }, { wch: 28 },                // Name, DOB, Phone, Email
    { wch: 40 },                                                        // Patient Address
    { wch: 28 }, { wch: 18 },                                          // Procedure Type, Billing Code (NEW)
    { wch: 20 }, { wch: 25 },                                          // Reason, Provider
    { wch: 25 }, { wch: 25 }, { wch: 40 },                             // Ref Doctor, Clinic, Address
    { wch: 16 }, { wch: 16 },                                          // Ref Phone, Fax
    { wch: 14 }, { wch: 16 }, { wch: 18 },                             // Referral Date, Expiry, Status
    { wch: 14 },                                                        // Account Type
  ];
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
    { s: { r: 1, c: 5 }, e: { r: 1, c: lastCol } },
  ];
  styleCell(sheet, "A1", { font: { bold: true, sz: 16, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= lastCol; c++) styleCell(sheet, XLSX.utils.encode_cell({ r: 3, c }), headerStyle());

  for (let r = 4; r <= range.e.r; r++) {
    const rowData = data[r]; if (!rowData || rowData.every(v => v === "")) continue;
    const isAlt = (r % 2) === 0;
    const status = rowData[21]; // Referral Status is now col 21
    const procedureType = rowData[10] || "";
    for (let c = 0; c <= lastCol; c++) {
      let bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
      if (c >= 1 && c <= 4) bg = COLOURS.dateHighlight;         // Day/Date/Time/Duration
      else if (c >= 5 && c <= 6) bg = COLOURS.patientInfoBg;    // Name/DOB
      else if (c >= 7 && c <= 8) bg = COLOURS.contactBg;        // Phone/Email
      else if (c === 9) bg = COLOURS.addressBg;                  // Patient Address
      else if (c === 10) bg = procedureType ? COLOURS.scopeBg : COLOURS.altRowBg; // Procedure (yellow if scope)
      else if (c === 11) bg = COLOURS.billingBg;                 // Billing Code
      else if (c === 12) bg = COLOURS.apptDetailBg;              // Reason
      else if (c === 13) bg = COLOURS.contactBg;                 // Provider
      else if (c >= 14 && c <= 18) bg = COLOURS.doctorBg;       // Doctor details
      if (c >= 19 && c <= 21) {                                  // Referral status columns
        if (status === "EXPIRED") bg = COLOURS.referralExpired;
        else if (status === "EXPIRING SOON") bg = COLOURS.referralExpiring;
        else if (status === "VALID") bg = COLOURS.referralOk;
      }
      const isBold = c === 5 || c === 10 || c === 14 || c === 21;
      const isCenter = [0,1,2,3,4,11,19,20,21,22].includes(c);
      styleCell(sheet, XLSX.utils.encode_cell({ r, c }), cellStyle(bg, isBold, isCenter, status));
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
  const lastCol = 16;
  sheet["!cols"] = [
    { wch: 28 }, { wch: 14 }, { wch: 16 }, { wch: 28 },   // Name, DOB, Phone, Email
    { wch: 40 },                                            // Full Address (NEW)
    { wch: 16 }, { wch: 25 },                               // Suburb, GP
    { wch: 25 }, { wch: 25 }, { wch: 40 },                  // Ref Doctor, Clinic, Address (NEW)
    { wch: 16 }, { wch: 16 },                               // Ref Phone, Fax (NEW)
    { wch: 14 }, { wch: 14 }, { wch: 16 },                  // Prov No, Ref Date, Expiry
    { wch: 18 }, { wch: 20 },                               // Status, Next Appt
  ];
  sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } }];
  styleCell(sheet, "A1", { font: { bold: true, sz: 14, color: { rgb: COLOURS.headerBg } } });
  for (let c = 0; c <= lastCol; c++) styleCell(sheet, XLSX.utils.encode_cell({ r: 2, c }), headerStyle());

  for (let r = 3; r <= range.e.r; r++) {
    const isAlt = (r % 2) === 0;
    const status = data[r] ? data[r][15] : "";
    for (let c = 0; c <= lastCol; c++) {
      let bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
      if (c === 0) bg = COLOURS.patientInfoBg;             // Name
      if (c === 2 || c === 3) bg = COLOURS.contactBg;      // Phone/Email
      if (c === 4) bg = COLOURS.addressBg;                  // Full Address (NEW)
      if (c >= 7 && c <= 11) bg = COLOURS.doctorBg;        // Doctor details (NEW)
      if (c === 16) bg = COLOURS.dateHighlight;             // Next Appt
      if (c >= 12 && c <= 15) {                             // Referral columns
        if (status === "EXPIRED") bg = COLOURS.referralExpired;
        else if (status === "EXPIRING SOON") bg = COLOURS.referralExpiring;
        else if (status === "VALID") bg = COLOURS.referralOk;
      }
      styleCell(sheet, XLSX.utils.encode_cell({ r, c }), cellStyle(bg, c === 0 || c === 7 || c === 15, false, status));
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

// ─── ICS & CSV ──────────────────────────────────────────────────────────────

function buildIcsCalendar(appointments) {
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "METHOD:PUBLISH", "X-WR-CALNAME:Genie Future Appointments"];
  for (const a of appointments) {
    lines.push("BEGIN:VEVENT", `SUMMARY:${a.patientName}`, `DTSTART;TZID=${DEFAULT_TIMEZONE}:${a.startDate.replace(/-/g, "")}T${a.startTime.replace(":", "")}00`, `DESCRIPTION:${a.reason}`, "END:VEVENT");
  }
  lines.push("END:VCALENDAR"); return lines.join("\r\n");
}

function buildCsv(appointments) {
  const header = [
    "patient_name", "date_of_birth", "mobile_phone", "email", "full_address",
    "appointment_date", "appointment_time",
    "procedure_type", "billing_codes",                    // NEW
    "reason",
    "referring_doctor", "ref_clinic", "ref_address", "ref_phone", "ref_fax",
    "referral_expiry", "referral_status",
  ];
  const rows = appointments.map((a) => [
    a.patientName, a.dateOfBirth, a.mobilePhone, a.email, a.fullAddress,
    a.startDate, a.startTime,
    a.procedureType || "", a.billingCodes || "",           // NEW
    a.reason,
    a.referringDoctor, a.refClinic, a.refAddress, a.refPhone, a.refFax,
    a.referralExpiry, a.referralStatus,
  ]);
  return [header, ...rows].map(r => r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
}

main();
