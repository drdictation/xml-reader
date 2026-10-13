import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import XLSX from "xlsx-js-style";
import "dotenv/config";

const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const OUTPUT_DIR = path.resolve("output");

function decode(value = "") {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"').replace(/&apos;|&#39;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([\da-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

function blocks(xml, tag) {
  return [...xml.matchAll(new RegExp("<" + tag + "(?:\\s[^>]*)?>([\\s\\S]*?)</" + tag + "\\s*>", "gi"))].map(m => m[1]);
}

function fields(xml) {
  const result = {};
  for (const m of xml.matchAll(/<([\w.-]+)(?:\s[^>]*)?>([^<>]*)<\/\1\s*>/g)) result[m[1].toLowerCase()] = decode(m[2].trim());
  return result;
}

function parseDate(value) {
  const s = String(value || "").trim();
  if (!s || /^0{4}[-/]0{2}[-/]0{2}/.test(s)) return "";
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + "-" + m[2].padStart(2, "0") + "-" + m[3].padStart(2, "0");
  m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (m) return m[3] + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0");
  return "";
}

function addMonths(iso, count) {
  const [y, m, d] = iso.split("-").map(Number);
  const date = new Date(Date.UTC(y, m - 1 + count, 1));
  const last = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate();
  date.setUTCDate(Math.min(d, last));
  return date.toISOString().slice(0, 10);
}

function extract(xml, source) {
  const patient = fields(blocks(xml, "patient")[0] || "");
  const name = patient.fullname || [patient.title, patient.firstname, patient.middlename, patient.surname].filter(Boolean).join(" ").trim();
  if (!name) return [];
  const rows = [];
  for (const block of blocks(xml, "recall")) {
    const recall = fields(block);
    rows.push({
      patient_name: name,
      recall_date: parseDate(recall.duedate),
      procedure: recall.reason || "",
      xml_file: source,
    });
  }
  return rows;
}

async function* walk(dir) {
  for (const entry of await fsp.readdir(dir, { withFileTypes: true })) {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* walk(file);
    else if (entry.isFile() && entry.name.toLowerCase().endsWith(".xml")) yield file;
  }
}

async function listDriveFiles(drive, folderId) {
  const files = [];
  let pageToken;
  do {
    const r = await drive.files.list({
      q: "'" + folderId + "' in parents and name contains '.xml' and trashed = false",
      fields: "nextPageToken,files(id,name)", pageSize: 1000, pageToken,
    });
    files.push(...(r.data.files || []));
    pageToken = r.data.nextPageToken;
  } while (pageToken);
  return files;
}

async function main() {
  const rows = [];
  let files = 0, patientFiles = 0, errors = 0;
  const localDir = process.argv[3] || (fs.existsSync(DEFAULT_XML_DIR) ? DEFAULT_XML_DIR : "");

  if (localDir && fs.existsSync(localDir)) {
    console.log("Scanning local XML files in " + localDir);
    for await (const file of walk(localDir)) {
      files++;
      try {
        const xml = await fsp.readFile(file, "utf8");
        if (!xml.includes("<patient_summary")) continue;
        patientFiles++;
        if (xml.includes("<recall>")) rows.push(...extract(xml, path.basename(file)));
      } catch (e) { errors++; console.error("Could not read " + file + ": " + e.message); }
    }
  } else if (process.env.GOOGLE_DRIVE_FOLDER_ID && process.env.GOOGLE_SERVICE_ACCOUNT_KEY) {
    const { google } = await import("googleapis");
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY),
      scopes: ["https://www.googleapis.com/auth/drive.readonly"],
    });
    const drive = google.drive({ version: "v3", auth });
    console.log("Scanning configured Google Drive patient folder");
    const driveFiles = await listDriveFiles(drive, process.env.GOOGLE_DRIVE_FOLDER_ID);
    files = driveFiles.length;
    let next = 0;
    async function worker() {
      while (true) {
        const i = next++;
        if (i >= driveFiles.length) return;
        const f = driveFiles[i];
        try {
          const r = await drive.files.get({ fileId: f.id, alt: "media" }, { responseType: "text" });
          const xml = String(r.data);
          if (xml.includes("<patient_summary")) {
            patientFiles++;
            if (xml.includes("<recall>")) rows.push(...extract(xml, f.name));
          }
        } catch (e) { errors++; console.error("Could not read " + f.name + ": " + e.message); }
        if ((i + 1) % 100 === 0) console.log("  Read " + (i + 1) + "/" + driveFiles.length + " XML files");
      }
    }
    await Promise.all(Array.from({ length: 5 }, worker));
  } else {
    throw new Error("No local XML folder or Google Drive credentials found. Pass a local folder, or configure GOOGLE_DRIVE_FOLDER_ID and GOOGLE_SERVICE_ACCOUNT_KEY in .env.");
  }

  await fsp.mkdir(OUTPUT_DIR, { recursive: true });
  const headers = ["patient_name", "recall_date", "procedure", "xml_file"];
  rows.sort((a, b) => (a.recall_date || "9999").localeCompare(b.recall_date || "9999") || a.patient_name.localeCompare(b.patient_name));
  const makeCsv = [headers, ...rows.map(r => headers.map(h => r[h] || ""))]
    .map(row => row.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(",")).join("\r\n");
  const csvPath = path.join(OUTPUT_DIR, "all-patient-recalls.csv");
  const xlsxPath = path.join(OUTPUT_DIR, "all-patient-recalls.xlsx");
  await fsp.writeFile(csvPath, makeCsv, "utf8");
  const widths = [{ wch: 30 }, { wch: 16 }, { wch: 42 }, { wch: 24 }];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows.map(r => headers.map(h => r[h] || ""))]);
  ws["!cols"] = widths;
  XLSX.utils.book_append_sheet(wb, ws, "All recall entries");
  XLSX.writeFile(wb, xlsxPath);
  console.log("Scanned " + files + " files (" + patientFiles + " patient XML files); recall entries: " + rows.length + "; read errors: " + errors);
  console.log("Excel: " + xlsxPath);
  console.log("CSV: " + csvPath);
  console.log("Excel: " + xlsxPath);
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
