import { google } from "googleapis";

const SCOPES = ["https://www.googleapis.com/auth/drive.readonly"];

/**
 * Get Google Drive client using Service Account credentials from environment variables.
 */
export async function getDriveClient() {
  const jsonKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!jsonKey) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set.");
  }

  const credentials = JSON.parse(jsonKey);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: SCOPES,
  });

  return google.drive({ version: "v3", auth });
}

/**
 * List files in a specific Google Drive folder.
 */
export async function listFiles(folderId: string) {
  const drive = await getDriveClient();
  const response = await drive.files.list({
    q: `'${folderId}' in parents and trashed = false`,
    fields: "files(id, name, mimeType)",
  });

  return response.data.files || [];
}

/**
 * Read text content from a Google Drive file.
 */
export async function getFileContent(fileId: string) {
  const drive = await getDriveClient();
  const response = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "text" }
  );

  return response.data as string;
}

/**
 * Read binary content from a Google Drive file.
 */
export async function getFileBinary(fileId: string) {
  const drive = await getDriveClient();
  const response = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "arraybuffer" }
  );

  return Buffer.from(response.data as ArrayBuffer);
}

/**
 * Export a Google Drive file (like a Word doc) as a PDF.
 */
export async function exportFileAsPdf(fileId: string) {
  const drive = await getDriveClient();
  const response = await drive.files.export(
    { fileId, mimeType: "application/pdf" },
    { responseType: "arraybuffer" }
  );

  return Buffer.from(response.data as ArrayBuffer);
}

/**
 * Specialized function to find and read the patient-index.json from a folder.
 */
export async function getPatientIndexFromDrive(folderId: string) {
  const drive = await getDriveClient();
  const response = await drive.files.list({
    q: `'${folderId}' in parents and name = 'patient-index.json' and trashed = false`,
    fields: "files(id, name)",
    pageSize: 1,
  });

  const files = response.data.files;
  if (!files || files.length === 0) {
    return null;
  }

  const content = await getFileContent(files[0].id!);
  return JSON.parse(content);
}
