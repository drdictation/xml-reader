import { google } from "googleapis";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import "dotenv/config";

// ─── Configuration ──────────────────────────────────────────────────────────
const FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
const JSON_KEY_STR = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

if (!FOLDER_ID || !JSON_KEY_STR) {
  console.error("Missing GOOGLE_DRIVE_FOLDER_ID or GOOGLE_SERVICE_ACCOUNT_KEY environment variables.");
  process.exit(1);
}

const credentials = JSON.parse(JSON_KEY_STR);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/drive.readonly"],
});

const drive = google.drive({ version: "v3", auth });

// ─── Extraction Logic (Mirrored from build-patient-index.mjs) ────────────────

function copyStr(str) {
  if (!str) return "";
  return Buffer.from(str, "utf8").toString("utf8");
}

function dec(value) {
  if (!value) return "";
  return copyStr(
    value
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'"),
  );
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

async function main() {
  console.log(`\n══════════════════════════════════════════════════`);
  console.log(`  SYNCING GOOGLE DRIVE PATIENT INDEX`);
  console.log(`  Folder ID: ${FOLDER_ID}`);
  console.log(`══════════════════════════════════════════════════\n`);

  try {
    // 1. Gather all XML files from Drive
    let allFiles = [];
    let pageToken = null;
    do {
      const res = await drive.files.list({
        q: `'${FOLDER_ID}' in parents and name contains '.xml' and trashed = false`,
        fields: "nextPageToken, files(id, name)",
        pageSize: 1000,
        pageToken,
      });
      allFiles = allFiles.concat(res.data.files || []);
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`Found ${allFiles.length} XML files on Drive.\n`);

    const patients = [];
    let patientCount = 0;

    for (let i = 0; i < allFiles.length; i++) {
      if (i > 0 && i % 100 === 0) {
        console.log(`  Processed ${i}/${allFiles.length} files...`);
      }

      try {
        // Read just the beginning of the file for efficiency (if possible) or full content
        // Drive API 'media' alt doesn't support Range header directly easily via this client, 
        // so we fetch the content. For 1600 files, this is okay for a sync script.
        const res = await drive.files.get({
          fileId: allFiles[i].id,
          alt: "media"
        }, { responseType: "text" });

        let xml = res.data;
        if (!xml.includes("<patient_summary")) continue;

        const patientBlocks = extractTagBlocks(xml, "patient");
        if (patientBlocks.length === 0) continue;

        const pf = extractDirectChildFields(patientBlocks[0]);

        const fullname =
          pf.fullname
            ? dec(pf.fullname)
            : [pf.firstname, pf.surname].filter(Boolean).map(dec).join(" ").trim();

        if (!fullname) continue;

        patientCount++;
        patients.push({
          name: fullname,
          surname: dec(pf.surname || ""),
          firstname: dec(pf.firstname || ""),
          dob: pf.dob || "",
          sex: pf.sex || "",
          urn: dec(pf.urn || pf.externalid || pf.medicarenum || ""),
          phone: dec(pf.mobilephone || pf.homephone || ""),
          suburb: dec(pf.suburb || ""),
          file: copyStr(allFiles[i].name),
          filePath: copyStr(allFiles[i].id), // Maps to fileId in the cloud version
        });
      } catch (err) {
        console.error(`  Error processing ${allFiles[i].name}:`, err.message);
      }
    }

    // Sort alphabetically
    patients.sort((a, b) => {
      const surnameCompare = a.surname.localeCompare(b.surname);
      if (surnameCompare !== 0) return surnameCompare;
      return a.firstname.localeCompare(b.firstname);
    });

    const indexContent = JSON.stringify(patients, null, 2);

    // 2. Save index locally so it can be committed to Git and bundled by Vercel
    const OUTPUT_FILE = path.resolve(process.cwd(), "output", "patient-index.json");
    fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
    fs.writeFileSync(OUTPUT_FILE, indexContent, "utf8");

    console.log(`\n══════════════════════════════════════════════════`);
    console.log(`  DONE!`);
    console.log(`  Total XML files processed : ${allFiles.length}`);
    console.log(`  Patient records found     : ${patients.length}`);
    console.log(`  Index saved to            : ${OUTPUT_FILE}`);
    console.log(``);
    console.log(`  NEXT STEPS:`);
    console.log(`  1. Commit output/patient-index.json to Git`);
    console.log(`  2. Push to GitHub — Vercel will auto-deploy`);
    console.log(`══════════════════════════════════════════════════\n`);

  } catch (err) {
    console.error("FATAL ERROR:", err);
    process.exit(1);
  }
}

main();
