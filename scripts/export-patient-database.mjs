import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx-js-style";

// ─── Configuration ──────────────────────────────────────────────────────────
const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const OUTPUT_DIR = "output";
const CSV_FILE = "patient-database.csv";
const XLSX_FILE = "patient-database.xlsx";

// ─── Colour palette ────────────────────────────────────────────────────────
const COLOURS = {
  headerBg: "1B3A5C",
  headerFont: "FFFFFF",
  dateHighlight: "E8F4FD",
  patientInfoBg: "FFF8E7",
  contactBg: "F0FFF0",
  referralOk: "E8F5E9",
  referralExpiring: "FFF3E0",
  referralExpired: "FFEBEE",
  doctorBg: "EDE7F6",       // Light purple for doctor details
  addressBg: "E3F2FD",      // Light blue for address
  altRowBg: "F7F9FC",
  borderColour: "B0C4DE",
};

// V8 Memory Fix: Detach sliced strings from their multi-megabyte parent string
function copyStr(str) {
  if (!str) return "";
  return Buffer.from(str, "utf8").toString("utf8");
}

function dec(value) {
  if (!value) return "";
  return copyStr(value
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'"));
}

function main() {
  const inputDir = path.resolve(process.cwd(), process.argv[2] || DEFAULT_XML_DIR);
  const outputDir = path.resolve(process.cwd(), OUTPUT_DIR);

  console.log(`\n======================================================`);
  console.log(`STARTING PATIENT DATABASE EXPORT`);
  console.log(`Scanning: ${inputDir}`);
  console.log(`======================================================\n`);

  if (!fs.existsSync(inputDir)) {
      console.error(`Error: Directory not found: ${inputDir}`);
      process.exit(1);
  }

  // Get all XML files
  const allFiles = fs.readdirSync(inputDir, { recursive: true })
    .filter(f => f.toLowerCase().endsWith(".xml"))
    .map(f => path.join(inputDir, f));
  console.log(`Found ${allFiles.length} XML files.\n`);

  const checkBuffer = Buffer.alloc(16384);

  // ─── Pass 1: Build Address Book Map (for Referrer details) ────────────────
  console.log(`Pass 1: Building address book map...`);
  const addressBookMap = new Map();
  let addressBookFileCount = 0;

  for (let i = 0; i < allFiles.length; i++) {
    if (i > 0 && i % 400 === 0) {
      console.log(`  Processed ${i}/${allFiles.length}...`);
      if (global.gc) global.gc();
    }
    try {
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
      xml = null;
    } catch (err) { /* skip */ }
  }
  console.log(`  Found ${addressBookMap.size} address book entries.\n`);

  // ─── Pass 2: Extract Patient Data ────────────────────────────────────────
  console.log(`Pass 2: Extracting patient details...`);
  const patients = [];
  let patientFileCount = 0;

  for (let i = 0; i < allFiles.length; i++) {
    if (i > 0 && i % 400 === 0) {
      console.log(`  Processed ${i}/${allFiles.length}... (${patients.length} patients found)`);
      if (global.gc) global.gc();
    }
    try {
      const fd = fs.openSync(allFiles[i], "r");
      fs.readSync(fd, checkBuffer, 0, 16384, 0);
      fs.closeSync(fd);
      if (!checkBuffer.toString("utf8").includes("<patient_summary")) continue;

      let xml = fs.readFileSync(allFiles[i], "utf8");
      if (xml.indexOf("<patient_summary>") !== -1) {
        patientFileCount++;
        const pData = extractPatientFromXml(xml, allFiles[i], addressBookMap);
        if (pData) patients.push(pData);
      }
      xml = null;
    } catch (err) {
      console.error(`  Error reading ${allFiles[i]}: ${err.message}`);
    }
  }

  // Sort alphabetically
  patients.sort((a, b) => a.patientName.localeCompare(b.patientName));

  fs.mkdirSync(outputDir, { recursive: true });
  const csvPath = path.join(outputDir, CSV_FILE);
  const xlsxPath = path.join(outputDir, XLSX_FILE);

  console.log(`\nGenerating output files...`);
  fs.writeFileSync(csvPath, buildCsv(patients), "utf8");
  buildExcel(patients, xlsxPath);

  console.log(`\n======================================================`);
  console.log(`Finished Successfully!`);
  console.log(`  Total files scanned  : ${allFiles.length}`);
  console.log(`  Patient records found: ${patients.length}`);
  console.log(`  Excel saved to       : ${xlsxPath}`);
  console.log(`  CSV saved to         : ${csvPath}`);
  console.log(`======================================================\n`);
}

function extractPatientFromXml(xml, filePath, addressBookMap) {
  const patientBlocks = extractTagBlocks(xml, "patient");
  if (patientBlocks.length === 0) return null;
  const pf = extractDirectChildFields(patientBlocks[0]);

  const patientName = getPatientName(pf) || copyStr(path.basename(filePath, ".xml"));
  
  const referralDate = pf.referraldate || "";
  const referralDurationMonths = parseInt(pf.referralduration || "0", 10);
  const referralExpiry = calculateReferralExpiry(referralDate, referralDurationMonths);
  const referralStatus = getReferralStatus(referralExpiry);

  const fullAddress = [pf.addressline1, pf.addressline2, pf.suburb, pf.state, pf.postcode]
    .filter(Boolean).map(dec).join(", ");

  const refId = pf.referrer_ab_id_fk || "";
  const refDetails = addressBookMap.get(refId) || {
    fullName: dec(pf.referringdoctor || ""),
    clinic: "", address: "", phone: "", fax: "", email: "",
    providerNum: dec(pf.referrerprovidernum || ""),
  };

  return {
    patientName,
    dateOfBirth: pf.dob || "",
    sex: pf.sex || "",
    mobilePhone: dec(pf.mobilephone || ""),
    homePhone: dec(pf.homephone || ""),
    email: dec(pf.emailaddress || ""),
    fullAddress,
    suburb: dec(pf.suburb || ""),
    
    // Clinical History
    lastSeenDate: pf.lastseendate || "",
    lastSeenBy: dec(pf.lastseenby || ""),
    lastOpvDate: pf.lastopvdate || "",
    usualGP: dec(pf.usualgp || ""),
    accountType: dec(pf.accounttype || ""),
    healthFund: dec(pf.healthfundname || ""),
    
    // Referral Details
    referringDoctor: refDetails.fullName,
    refClinic: refDetails.clinic,
    refAddress: refDetails.address,
    refPhone: refDetails.phone,
    refFax: refDetails.fax,
    refEmail: refDetails.email,
    referralDate,
    referralExpiry,
    referralStatus,
    
    sourceFile: copyStr(path.basename(filePath)),
  };
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

function formatDateForDisplay(isoDate) {
  if (!isoDate || isoDate.startsWith("0000")) return "";
  const [y, m, d] = isoDate.split("-");
  return `${d}/${m}/${y}`;
}

function pad(v) { return String(v).padStart(2, "0"); }

// ─── Excel Generation ───────────────────────────────────────────────────────

function buildExcel(patients, outputPath) {
  const wb = XLSX.utils.book_new();
  const header = [
    "PATIENT NAME", "DOB", "SEX", "MOBILE PHONE", "HOME PHONE", "EMAIL",
    "FULL ADDRESS", "SUBURB",
    "LAST SEEN DATE", "LAST SEEN BY", "USUAL GP", "ACCOUNT TYPE", "HEALTH FUND",
    "REFERRING DOCTOR", "REF CLINIC", "REF ADDRESS", "REF PHONE", "REF EMAIL",
    "REFERRAL DATE", "REFERRAL EXPIRY", "REFERRAL STATUS"
  ];

  const rows = [
    ["PATIENT DATABASE EXPORT — ALL PRIOR PATIENTS"],
    [`Generated: ${new Date().toLocaleDateString("en-AU")} at ${new Date().toLocaleTimeString("en-AU")}`, "", "", "", `Total: ${patients.length} patients`],
    [],
    header
  ];

  for (const p of patients) {
    rows.push([
      p.patientName, formatDateForDisplay(p.dateOfBirth), p.sex, p.mobilePhone, p.homePhone, p.email,
      p.fullAddress, p.suburb,
      formatDateForDisplay(p.lastSeenDate), p.lastSeenBy, p.usualGP, p.accountType, p.healthFund,
      p.referringDoctor, p.refClinic, p.refAddress, p.refPhone, p.refEmail,
      formatDateForDisplay(p.referralDate), formatDateForDisplay(p.referralExpiry), p.referralStatus
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  applyStyles(ws, rows);
  XLSX.utils.book_append_sheet(wb, ws, "Patient Database");
  XLSX.writeFile(wb, outputPath);
}

function applyStyles(ws, rows) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  ws["!cols"] = [
    { wch: 28 }, { wch: 14 }, { wch: 8 }, { wch: 16 }, { wch: 16 }, { wch: 28 }, // Demographics
    { wch: 40 }, { wch: 18 },                                                 // Address
    { wch: 16 }, { wch: 25 }, { wch: 25 }, { wch: 14 }, { wch: 20 },          // Clinical
    { wch: 25 }, { wch: 25 }, { wch: 40 }, { wch: 16 }, { wch: 28 },          // Ref Doctor
    { wch: 14 }, { wch: 16 }, { wch: 18 }                                      // Referral
  ];

  // Title
  styleCell(ws, "A1", { font: { bold: true, sz: 16, color: { rgb: COLOURS.headerBg } } });
  
  // Headers
  for (let c = 0; c <= 20; c++) {
    styleCell(ws, XLSX.utils.encode_cell({ r: 3, c }), {
      fill: { fgColor: { rgb: COLOURS.headerBg } },
      font: { color: { rgb: COLOURS.headerFont }, bold: true },
      border: thinB(),
      alignment: { horizontal: "center" }
    });
  }

  // Row styling
  for (let r = 4; r <= range.e.r; r++) {
    const isAlt = (r % 2) === 0;
    const status = rows[r] ? rows[r][20] : "";
    for (let c = 0; c <= 20; c++) {
      let bg = isAlt ? COLOURS.altRowBg : "FFFFFF";
      
      // Demographics highlight
      if (c >= 0 && c <= 2) bg = COLOURS.patientInfoBg;
      else if (c >= 3 && c <= 5) bg = COLOURS.contactBg;
      else if (c >= 6 && c <= 7) bg = COLOURS.addressBg;
      // Clinical highlight
      else if (c >= 8 && c <= 12) bg = "F0F4C3"; // Light Lime
      // Referral highlights
      else if (c >= 13 && c <= 17) bg = COLOURS.doctorBg;
      else if (c >= 18 && c <= 20) {
        if (status === "EXPIRED") bg = COLOURS.referralExpired;
        else if (status === "EXPIRING SOON") bg = COLOURS.referralExpiring;
        else if (status === "VALID") bg = COLOURS.referralOk;
      }

      styleCell(ws, XLSX.utils.encode_cell({ r, c }), {
        fill: { fgColor: { rgb: bg } },
        font: { name: "Calibri", bold: (c === 0 || c === 13 || c === 20) },
        border: thinB(),
        alignment: { vertical: "center", horizontal: [1, 2, 8, 18, 19, 20].includes(c) ? "center" : "left" }
      });
    }
  }
}

function styleCell(sheet, addr, style) { if (sheet[addr]) sheet[addr].s = style; }
function thinB() { return { top: { style: "thin", color: { rgb: COLOURS.borderColour } }, bottom: { style: "thin", color: { rgb: COLOURS.borderColour } }, left: { style: "thin", color: { rgb: COLOURS.borderColour } }, right: { style: "thin", color: { rgb: COLOURS.borderColour } } }; }

function buildCsv(patients) {
  const header = [
    "patient_name", "date_of_birth", "sex", "mobile_phone", "home_phone", "email",
    "full_address", "suburb",
    "last_seen_date", "last_seen_by", "usual_gp", "account_type", "health_fund",
    "referring_doctor", "ref_clinic", "ref_address", "ref_phone", "ref_email",
    "referral_date", "referral_expiry", "referral_status"
  ];
  const rows = patients.map((p) => [
    p.patientName, p.dateOfBirth, p.sex, p.mobilePhone, p.homePhone, p.email,
    p.fullAddress, p.suburb,
    p.lastSeenDate, p.lastSeenBy, p.usualGP, p.accountType, p.healthFund,
    p.referringDoctor, p.refClinic, p.refAddress, p.refPhone, p.refEmail,
    p.referralDate, p.referralExpiry, p.referralStatus
  ]);
  return [header, ...rows].map(r => r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
}

main();
