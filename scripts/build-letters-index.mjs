import { google } from "googleapis";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import "dotenv/config";

// ─── Configuration ──────────────────────────────────────────────────────────
const FOLDER_ID = process.env.GOOGLE_DRIVE_LETTERS_FOLDER_ID;
const JSON_KEY_STR = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

if (!FOLDER_ID || !JSON_KEY_STR) {
  console.error("Missing GOOGLE_DRIVE_LETTERS_FOLDER_ID or GOOGLE_SERVICE_ACCOUNT_KEY environment variables.");
  process.exit(1);
}

const credentials = JSON.parse(JSON_KEY_STR);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/drive.readonly"],
});

const drive = google.drive({ version: "v3", auth });

async function main() {
  console.log(`\n══════════════════════════════════════════════════`);
  console.log(`  BUILDING LETTERS INDEX FROM GOOGLE DRIVE`);
  console.log(`  Folder ID: ${FOLDER_ID}`);
  console.log(`══════════════════════════════════════════════════\n`);

  try {
    // 1. Gather all files from the letters folder
    let allFiles = [];
    let pageToken = null;
    do {
      const res = await drive.files.list({
        q: `'${FOLDER_ID}' in parents and trashed = false`,
        fields: "nextPageToken, files(id, name, mimeType)",
        pageSize: 1000,
        pageToken,
      });
      allFiles = allFiles.concat(res.data.files || []);
      pageToken = res.data.nextPageToken;
    } while (pageToken);

    console.log(`Found ${allFiles.length} files in the letters folder.\n`);

    const lettersIndex = {};
    let matchedCount = 0;

    for (const file of allFiles) {
      // Extract numeric prefix (Genie Letter ID)
      // Example: "18632_Dr Laura Shobbrook (1).docx" -> "18632"
      const match = file.name.match(/^(\d+)_/);
      if (match) {
        const letterId = match[1];
        lettersIndex[letterId] = {
          fileId: file.id,
          fileName: file.name,
          mimeType: file.mimeType,
        };
        matchedCount++;
      }
    }

    console.log(`Successfully indexed ${matchedCount} letters.\n`);

    // 2. Save index locally
    const OUTPUT_FILE = path.resolve(process.cwd(), "output", "letters-index.json");
    fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
    fs.writeFileSync(OUTPUT_FILE, JSON.stringify(lettersIndex, null, 2), "utf8");

    console.log(`\n══════════════════════════════════════════════════`);
    console.log(`  DONE!`);
    console.log(`  Total files scanned   : ${allFiles.length}`);
    console.log(`  Letters indexed       : ${matchedCount}`);
    console.log(`  Index saved to        : ${OUTPUT_FILE}`);
    console.log(`══════════════════════════════════════════════════\n`);

  } catch (err) {
    if (err.message.includes("File not found")) {
      console.error("\nERROR: Could not access the folder. Have you shared it with the service account?");
      console.error(`Service Account: ${credentials.client_email}`);
    } else {
      console.error("FATAL ERROR:", err);
    }
    process.exit(1);
  }
}

main();
