import fs from "node:fs";
import path from "node:path";

// ─── Configuration ──────────────────────────────────────────────────────────
const DEFAULT_XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards";
const OUTPUT_FILE = path.resolve(process.cwd(), "output", "patient-index.json");

// V8 memory fix: deep-copy sliced strings so the multi-MB parent can be GC'd
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

function main() {
  const inputDir = path.resolve(process.argv[2] || DEFAULT_XML_DIR);
  console.log(`\n══════════════════════════════════════════════════`);
  console.log(`  BUILDING PATIENT INDEX`);
  console.log(`  Scanning: ${inputDir}`);
  console.log(`══════════════════════════════════════════════════\n`);

  // Gather all XML files
  const allFiles = fs
    .readdirSync(inputDir, { recursive: true })
    .filter((f) => f.toLowerCase().endsWith(".xml"))
    .map((f) => path.join(inputDir, f));

  console.log(`Found ${allFiles.length} XML files.\n`);

  const checkBuffer = Buffer.alloc(16384);
  const patients = [];
  let patientCount = 0;

  for (let i = 0; i < allFiles.length; i++) {
    if (i > 0 && i % 500 === 0) {
      console.log(`  Scanned ${i}/${allFiles.length} files (${patientCount} patients found)...`);
      if (global.gc) global.gc();
    }

    try {
      // Quick check: only read files that look like patient summaries
      const fd = fs.openSync(allFiles[i], "r");
      fs.readSync(fd, checkBuffer, 0, 16384, 0);
      fs.closeSync(fd);
      if (!checkBuffer.toString("utf8").includes("<patient_summary")) continue;

      let xml = fs.readFileSync(allFiles[i], "utf8");
      if (xml.indexOf("<patient_summary>") === -1) {
        xml = null;
        continue;
      }

      const patientBlocks = extractTagBlocks(xml, "patient");
      if (patientBlocks.length === 0) {
        xml = null;
        continue;
      }

      const pf = extractDirectChildFields(patientBlocks[0]);
      xml = null; // Free the large string

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
        file: copyStr(path.basename(allFiles[i])),
        filePath: copyStr(allFiles[i]),
      });
    } catch (err) {
      // skip unreadable files
    }
  }

  // Sort alphabetically by surname, then firstname
  patients.sort((a, b) => {
    const surnameCompare = a.surname.localeCompare(b.surname);
    if (surnameCompare !== 0) return surnameCompare;
    return a.firstname.localeCompare(b.firstname);
  });

  // Write out
  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(patients, null, 2), "utf8");

  console.log(`\n══════════════════════════════════════════════════`);
  console.log(`  DONE!`);
  console.log(`  Total XML files scanned : ${allFiles.length}`);
  console.log(`  Patient records found   : ${patients.length}`);
  console.log(`  Index saved to          : ${OUTPUT_FILE}`);
  console.log(`══════════════════════════════════════════════════\n`);

  // Show a preview
  console.log(`First 10 entries:`);
  for (const p of patients.slice(0, 10)) {
    console.log(`  • ${p.name} (DOB: ${p.dob || "N/A"}, URN: ${p.urn || "N/A"}) → ${p.file}`);
  }
}

main();
