export type FieldMap = Record<string, string>;

export type ParsedRecord = {
  id: string;
  title: string;
  subtitle?: string;
  fields: FieldMap;
  primaryDate?: string;
  sortTimestamp: number;
  isFuture?: boolean;
  preview?: EmbeddedPreview;
  exportReadability?: {
    status: "readable" | "metadata_only";
    label: string;
    detail: string;
  };
};

export type EmbeddedPreview =
  | { type: "html"; content: string; sourceField: string }
  | { type: "text"; content: string; sourceField: string }
  | {
      type: "binary";
      mimeType: string;
      base64: string;
      fileName?: string;
      sourceField: string;
      displayMode: "pdf" | "image" | "download";
    };

export type ParsedSection = {
  key: string;
  label: string;
  description: string;
  records: ParsedRecord[];
};

export type ParsedGenieXml = {
  fileName: string;
  rawXml: string;
  exportInfo: FieldMap;
  patient: FieldMap;
  sections: ParsedSection[];
};

type SectionDefinition = {
  key: string;
  label: string;
  description: string;
  titleFields: string[];
  subtitleFields: string[];
  dateFields: string[];
};

const SECTION_DEFINITIONS: Record<string, SectionDefinition> = {
  consult_list: {
    key: "consult_list",
    label: "Consults",
    description: "Clinical notes and consult records.",
    titleFields: ["description", "diagnosis", "doctorname", "consultdate"],
    subtitleFields: ["history", "plan", "examination", "creator"],
    dateFields: ["consultdate", "consulttime", "datecreated", "lastupdated"],
  },
  downloadedresult_list: {
    key: "downloadedresult_list",
    label: "Results",
    description: "Downloaded pathology and imaging results.",
    titleFields: ["test", "reportdate", "providername"],
    subtitleFields: ["result", "specimen", "status", "resultstatus"],
    dateFields: ["reportdate", "receiveddate", "requestdate", "datecreated", "lastupdated"],
  },
  outgoingletter_list: {
    key: "outgoingletter_list",
    label: "Letters",
    description: "Outgoing clinical letters and embedded documents.",
    titleFields: ["title", "description", "recipientname", "letterdate"],
    subtitleFields: ["documenttype", "doctorname", "recipientname"],
    dateFields: ["letterdate", "datecreated", "modificationdate", "lastupdated"],
  },
  referral_list: {
    key: "referral_list",
    label: "Referrals",
    description: "Referral history and referrer details.",
    titleFields: ["referraldate", "providername", "title"],
    subtitleFields: ["firstname", "surname", "reason"],
    dateFields: ["referraldate", "datecreated", "lastupdated"],
  },
  scriptarchive_list: {
    key: "scriptarchive_list",
    label: "Scripts",
    description: "Archived scripts and medication history.",
    titleFields: ["medication", "scriptdate", "datecreated"],
    subtitleFields: ["dose", "quantity", "repeats", "doctorname"],
    dateFields: ["scriptdate", "datecreated", "lastupdated"],
  },
  prescription_list: {
    key: "prescription_list",
    label: "Prescriptions",
    description: "Active prescription records.",
    titleFields: ["medication", "datecreated", "doctorname"],
    subtitleFields: ["dose", "quantity", "instructions"],
    dateFields: ["datecreated", "lastupdated"],
  },
  prescriptionhistory_list: {
    key: "prescriptionhistory_list",
    label: "Prescription History",
    description: "Prescription change history.",
    titleFields: ["medication", "datecreated", "doctorname"],
    subtitleFields: ["dose", "quantity", "instructions"],
    dateFields: ["datecreated", "lastupdated"],
  },
  appt_list: {
    key: "appt_list",
    label: "Appointments",
    description: "Patient appointment records.",
    titleFields: ["apptreason", "reason", "apptprovidername", "providername", "apptstartdate", "startdate"],
    subtitleFields: ["apptstarttime", "starttime", "apptroomname", "eventtype"],
    dateFields: ["startdatetime", "startdate", "apptstartdate", "apptstarttime", "starttime", "datecreated", "lastupdated"],
  },
  appttimingevent_list: {
    key: "appttimingevent_list",
    label: "Appointment Timing",
    description: "Appointment timing events and workflow steps.",
    titleFields: ["startdate", "starttime", "eventtype"],
    subtitleFields: ["providername", "reason", "durationseconds"],
    dateFields: ["startdatetime", "startdate", "starttime", "datecreated", "lastupdated"],
  },
  task_list: {
    key: "task_list",
    label: "Tasks",
    description: "Tasks and follow-up actions.",
    titleFields: ["note", "taskfor", "taskdate"],
    subtitleFields: ["creator", "completed", "urgentfg"],
    dateFields: ["taskdate", "datecompleted", "datecreated", "lastupdated"],
  },
  smslog_list: {
    key: "smslog_list",
    label: "SMS Log",
    description: "SMS messages and responses.",
    titleFields: ["recipientname", "message", "sentdate"],
    subtitleFields: ["message", "reply", "prfusr_name"],
    dateFields: ["sentdate", "senttime", "replydate", "replytime", "datecreated", "lastupdated"],
  },
  graphic_list: {
    key: "graphic_list",
    label: "Documents",
    description: "Attached graphics and scanned items.",
    titleFields: ["description", "realname", "imagedate"],
    subtitleFields: ["windocumenttype", "documenttype", "reviewby"],
    dateFields: ["imagedate", "consultdate", "datecreated", "modificationdate", "lastupdated"],
  },
  measurement_list: {
    key: "measurement_list",
    label: "Measurements",
    description: "Recorded measurements and observations.",
    titleFields: ["datecreated", "title", "creator"],
    subtitleFields: ["value", "unit", "note"],
    dateFields: ["datecreated", "lastupdated"],
  },
  patientclinical_list: {
    key: "patientclinical_list",
    label: "Clinical",
    description: "Clinical flags and coded patient data.",
    titleFields: ["title", "datecreated", "creator"],
    subtitleFields: ["note", "value", "description"],
    dateFields: ["datecreated", "lastupdated"],
  },
  consultationproblem_list: {
    key: "consultationproblem_list",
    label: "Problems",
    description: "Consultation problem list.",
    titleFields: ["description", "datecreated", "creator"],
    subtitleFields: ["note", "status", "title"],
    dateFields: ["datecreated", "lastupdated"],
  },
  procedures_list: {
    key: "procedures_list",
    label: "Procedures",
    description: "Procedural history.",
    titleFields: ["title", "datecreated", "doctorname"],
    subtitleFields: ["description", "note"],
    dateFields: ["datecreated", "lastupdated"],
  },
  customform_list: {
    key: "customform_list",
    label: "Custom Forms",
    description: "Custom forms included in the export.",
    titleFields: ["title", "datecreated", "creator"],
    subtitleFields: ["description", "note"],
    dateFields: ["datecreated", "lastupdated"],
  },
  addressbook_list: {
    key: "addressbook_list",
    label: "Address Book",
    description: "Linked providers and contacts.",
    titleFields: ["title", "firstname", "surname"],
    subtitleFields: ["providername", "speciality", "email"],
    dateFields: ["datecreated", "lastupdated"],
  },
  interestedparty_list: {
    key: "interestedparty_list",
    label: "Interested Parties",
    description: "Related contacts and interested parties.",
    titleFields: ["title", "firstname", "surname"],
    subtitleFields: ["description", "email", "mobile"],
    dateFields: ["datecreated", "lastupdated"],
  },
  hiaudit_list: {
    key: "hiaudit_list",
    label: "HI Audit",
    description: "Healthcare identifier audit records.",
    titleFields: ["datecreated", "creator", "status"],
    subtitleFields: ["note", "description"],
    dateFields: ["datecreated", "lastupdated"],
  },
  hoimcclaim_list: {
    key: "hoimcclaim_list",
    label: "Claims",
    description: "Claim-related records exported from Genie.",
    titleFields: ["datecreated", "creator", "status"],
    subtitleFields: ["description", "note"],
    dateFields: ["datecreated", "lastupdated"],
  },
};

const IGNORE_EMPTY = new Set(["createdat", "gcs_createdat"]);
const FALLBACK_DATE_FIELDS = [
  "date_time",
  "consultdate",
  "reportdate",
  "letterdate",
  "referraldate",
  "imagedate",
  "startdatetime",
  "apptstartdate",
  "taskdate",
  "sentdate",
  "replydate",
  "requestdate",
  "receiveddate",
  "datecompleted",
  "datecreated",
  "modificationdate",
  "lastupdated",
  "gcs_updatedat",
];

export function parseGenieXml(xml: string, fileName: string): ParsedGenieXml {
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "text/xml");
  const parseErrors = document.getElementsByTagName("parsererror");

  if (parseErrors.length > 0) {
    throw new Error("This file could not be parsed as XML.");
  }

  const root = document.documentElement;

  if (!root || root.tagName !== "patient_summary") {
    throw new Error("This does not look like a Genie patient_summary XML export.");
  }

  const exportInfo = getSingleRecord(root, "export_info");
  const patient = getSingleRecord(root, "patient");
  const sections = Array.from(root.children)
    .filter((child) => child.tagName.endsWith("_list"))
    .map((section) => createSection(section))
    .filter((section): section is ParsedSection => section.records.length > 0);

  return {
    fileName,
    rawXml: xml,
    exportInfo,
    patient,
    sections,
  };
}

function createSection(sectionElement: Element): ParsedSection {
  const definition =
    SECTION_DEFINITIONS[sectionElement.tagName] ?? {
      key: sectionElement.tagName,
      label: toTitle(sectionElement.tagName.replace(/_list$/, "")),
      description: "Additional Genie export records.",
      titleFields: ["title", "description", "datecreated"],
      subtitleFields: ["note", "creator"],
      dateFields: FALLBACK_DATE_FIELDS,
    };

  const nowMs = Date.now();
  const isApptSection = sectionElement.tagName === "appt_list";

  const records = Array.from(sectionElement.children)
    .map((record, index) => {
      const fields = elementToFieldMap(record);
      const { primaryDate, sortTimestamp } = getRecordDateInfo(fields, definition.dateFields);
      const preview = getEmbeddedPreview(fields);
      const title = pickFirstValue(fields, definition.titleFields) ?? `${definition.label} ${index + 1}`;
      const subtitle = pickFirstValue(fields, definition.subtitleFields);
      const isFuture = isApptSection && sortTimestamp > nowMs;

      return {
        id: fields.id || fields.uuid || `${sectionElement.tagName}-${index + 1}`,
        title,
        subtitle,
        fields,
        primaryDate,
        sortTimestamp,
        isFuture,
        preview,
        exportReadability: getExportReadability(sectionElement.tagName, fields, preview),
      };
    })
    .sort((left, right) => {
      // For appointments: future first (ascending), then past (descending)
      if (isApptSection) {
        if (left.isFuture && !right.isFuture) return -1;
        if (!left.isFuture && right.isFuture) return 1;
        if (left.isFuture && right.isFuture) {
          // Both future: soonest first
          return left.sortTimestamp - right.sortTimestamp;
        }
      }
      // Default: newest first
      if (left.sortTimestamp !== right.sortTimestamp) {
        return right.sortTimestamp - left.sortTimestamp;
      }
      return left.title.localeCompare(right.title);
    });

  return {
    key: definition.key,
    label: definition.label,
    description: definition.description,
    records,
  };
}

function getSingleRecord(root: Element, tagName: string): FieldMap {
  const element = root.getElementsByTagName(tagName)[0];
  return element ? elementToFieldMap(element) : {};
}

function elementToFieldMap(element: Element): FieldMap {
  const fields: FieldMap = {};

  Array.from(element.children).forEach((child) => {
    if (hasElementChildren(child)) {
      fields[child.tagName] = flattenNestedElement(child);
      return;
    }

    const value = normalizeValue(child.textContent ?? "");

    if (!value && IGNORE_EMPTY.has(child.tagName)) {
      return;
    }

    fields[child.tagName] = value;
  });

  return fields;
}

function flattenNestedElement(element: Element): string {
  return Array.from(element.children)
    .map((child) => {
      const value = hasElementChildren(child)
        ? flattenNestedElement(child)
        : normalizeValue(child.textContent ?? "");
      return value ? `${toTitle(child.tagName)}: ${value}` : "";
    })
    .filter(Boolean)
    .join("\n");
}

function hasElementChildren(element: Element): boolean {
  return Array.from(element.children).length > 0;
}

function normalizeValue(value: string): string {
  return value.replace(/\r\n/g, "\n").trim();
}

function pickFirstValue(fields: FieldMap, keys: string[]): string | undefined {
  for (const key of keys) {
    const value = fields[key];
    if (value && value !== "0000-00-00" && value !== "0000-00-00T00:00:00") {
      return value;
    }
  }

  return undefined;
}

function getEmbeddedPreview(fields: FieldMap): EmbeddedPreview | undefined {
  const candidates = ["document_content", "referralcontent_", "graphic"];

  for (const sourceField of candidates) {
    const source = fields[sourceField];

    if (!source) {
      continue;
    }

    const trimmed = source.trim();
    const binaryFileName = fields.realname || fields.description || fields.title;

    if (looksLikeHtml(trimmed)) {
      return { type: "html", content: trimmed, sourceField };
    }

    if (trimmed.startsWith("data:")) {
      const parsedDataUri = parseDataUri(trimmed);

      if (parsedDataUri) {
        return {
          type: "binary",
          mimeType: parsedDataUri.mimeType,
          base64: parsedDataUri.base64,
          fileName: binaryFileName,
          sourceField,
          displayMode: getDisplayMode(parsedDataUri.mimeType),
        };
      }
    }

    const inferredMimeType = inferMimeType(trimmed, binaryFileName, fields.windocumenttype);

    if (inferredMimeType && looksLikeBase64(trimmed)) {
      return {
        type: "binary",
        mimeType: inferredMimeType,
        base64: collapseWhitespace(trimmed),
        fileName: binaryFileName,
        sourceField,
        displayMode: getDisplayMode(inferredMimeType),
      };
    }

    if (trimmed.length <= 5000) {
      return { type: "text", content: trimmed, sourceField };
    }
  }

  return undefined;
}

function looksLikeHtml(value: string): boolean {
  return /^<(html|div|p|span|table|body|section|article)\b/i.test(value);
}

function getExportReadability(
  sectionKey: string,
  fields: FieldMap,
  preview?: EmbeddedPreview,
): ParsedRecord["exportReadability"] | undefined {
  if (sectionKey !== "outgoingletter_list") {
    return undefined;
  }

  const hasEmbeddedLetterBody =
    Boolean(fields.referralcontent_?.trim()) &&
    fields.referralcontent_.trim() !== "<![CDATA[]]>" &&
    (preview?.sourceField === "referralcontent_" || fields.referralcontent_.trim().length > 0);

  if (hasEmbeddedLetterBody) {
    return {
      status: "readable",
      label: "Readable in export",
      detail: "This letter body is embedded in the XML and can be shown in the viewer.",
    };
  }

  return {
    status: "metadata_only",
    label: "Not readable in export",
    detail:
      "This XML contains the letter metadata only. The actual letter body is not present in this export record.",
  };
}

function getRecordDateInfo(fields: FieldMap, preferredKeys: string[]): {
  primaryDate?: string;
  sortTimestamp: number;
} {
  const dateKeys = [...preferredKeys, ...FALLBACK_DATE_FIELDS];

  for (const key of dateKeys) {
    const value = fields[key];

    if (!value) {
      continue;
    }

    const parsedDate = parseDateValue(value, key);

    if (parsedDate) {
      return {
        primaryDate: value,
        sortTimestamp: parsedDate.getTime(),
      };
    }
  }

  return { sortTimestamp: 0 };
}

export function formatAustralianDate(value?: string, fieldName?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const parsedDate = parseDateValue(value, fieldName);

  if (!parsedDate) {
    return undefined;
  }

  const hasTime = /T\d{2}:\d{2}/.test(value) || /\d{14}/.test(value) || /time/i.test(fieldName ?? "");

  return new Intl.DateTimeFormat("en-AU", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    ...(hasTime
      ? {
          hour: "2-digit",
          minute: "2-digit",
          hour12: false,
        }
      : {}),
  }).format(parsedDate);
}

export function isDateField(fieldName: string): boolean {
  return /(date|dob|dod|expiry|updatedat|verified|time)/i.test(fieldName);
}

export function isBinaryLikeField(fieldName: string, value: string, preview?: EmbeddedPreview): boolean {
  if (preview && preview.sourceField === fieldName) {
    return true;
  }

  return ["document_content", "referralcontent_", "graphic"].includes(fieldName) && looksLikeBase64(value);
}

function parseDateValue(value: string, fieldName?: string): Date | undefined {
  const trimmed = value.trim();

  if (!trimmed || trimmed.startsWith("0000-00-00")) {
    return undefined;
  }

  if (/^\d{14}$/.test(trimmed)) {
    const year = Number(trimmed.slice(0, 4));
    const month = Number(trimmed.slice(4, 6));
    const day = Number(trimmed.slice(6, 8));
    const hour = Number(trimmed.slice(8, 10));
    const minute = Number(trimmed.slice(10, 12));
    const second = Number(trimmed.slice(12, 14));
    return new Date(year, month - 1, day, hour, minute, second);
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const [year, month, day] = trimmed.split("-").map(Number);
    return new Date(year, month - 1, day);
  }

  if (/^\d{4}-\d{2}-\d{2}T/.test(trimmed)) {
    const parsed = new Date(trimmed);
    return Number.isNaN(parsed.getTime()) ? undefined : parsed;
  }

  if (/time/i.test(fieldName ?? "") && /^\d{2}:\d{2}/.test(trimmed)) {
    const [hour, minute] = trimmed.split(":").map(Number);
    const today = new Date();
    today.setHours(hour, minute, 0, 0);
    return today;
  }

  return undefined;
}

function parseDataUri(value: string): { mimeType: string; base64: string } | undefined {
  const match = value.match(/^data:([^;]+);base64,([\s\S]+)$/);

  if (!match) {
    return undefined;
  }

  return {
    mimeType: match[1],
    base64: collapseWhitespace(match[2]),
  };
}

function inferMimeType(value: string, fileName?: string, documentType?: string): string | undefined {
  const compactValue = collapseWhitespace(value);
  const lowerFileName = fileName?.toLowerCase() ?? "";
  const lowerDocumentType = documentType?.toLowerCase() ?? "";

  if (compactValue.startsWith("JVBERi0") || lowerFileName.endsWith(".pdf") || lowerDocumentType === "pdf") {
    return "application/pdf";
  }

  if (compactValue.startsWith("/9j/") || lowerFileName.endsWith(".jpg") || lowerFileName.endsWith(".jpeg")) {
    return "image/jpeg";
  }

  if (compactValue.startsWith("iVBORw0KGgo") || lowerFileName.endsWith(".png")) {
    return "image/png";
  }

  if (compactValue.startsWith("R0lGOD") || lowerFileName.endsWith(".gif")) {
    return "image/gif";
  }

  return undefined;
}

function getDisplayMode(mimeType: string): "pdf" | "image" | "download" {
  if (mimeType === "application/pdf") {
    return "pdf";
  }

  if (mimeType.startsWith("image/")) {
    return "image";
  }

  return "download";
}

function looksLikeBase64(value: string): boolean {
  const compactValue = collapseWhitespace(value);

  return compactValue.length > 120 && /^[A-Za-z0-9+/=]+$/.test(compactValue);
}

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, "");
}

function toTitle(value: string): string {
  return value
    .replace(/_/g, " ")
    .replace(/\b\w/g, (character) => character.toUpperCase());
}
