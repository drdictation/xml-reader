"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  formatAustralianDate,
  isBinaryLikeField,
  isDateField,
  parseGenieXml,
  type ParsedGenieXml,
  type ParsedRecord,
  type EmbeddedPreview,
} from "../lib/genie-xml";
import {
  getLetterDownload,
  getLetterHtml,
  getLetterPdf,
  getLettersIndex,
  getPatientIndex,
  readBackupFile,
  type LetterIndexEntry,
  type PatientIndexEntry,
} from "../app/actions";

const PATIENT_FIELDS = [
  "fullname",
  "firstname",
  "surname",
  "dob",
  "sex",
  "emailaddress",
  "mobilephone",
  "homephone",
  "addressline1",
  "suburb",
  "state",
  "postcode",
  "medicarenum",
  "medicarerefnum",
  "ihi",
  "usualgp",
  "referringdoctor",
  "lastseendate",
  "accounttype",
  "healthfundname",
  "healthfundnum",
  "memo",
];

const EXPORT_FIELDS = ["date_time", "genie_app_version", "genie_instance_id"];

const FIELD_PRIORITY = [
  "consultdate",
  "reportdate",
  "letterdate",
  "referraldate",
  "imagedate",
  "datecreated",
  "description",
  "note",
  "message",
  "reply",
  "history",
  "examination",
  "plan",
  "diagnosis",
  "document_content",
  "result",
  "graphic",
];

const SECTION_PRIMARY_FIELDS: Record<string, string[]> = {
  consult_list: ["doctorname", "creator", "diagnosis", "history", "plan", "examination"],
  downloadedresult_list: ["test", "providername", "resultstatus", "status", "specimen", "result"],
  outgoingletter_list: ["addresseename", "creator", "documenttype", "readytosend", "reviewed"],
  referral_list: ["providername", "firstname", "surname", "reason"],
  scriptarchive_list: ["medication", "dose", "qty", "repeat", "creator", "note"],
  prescription_list: ["medication", "dose", "qty", "repeat", "creator", "note"],
  prescriptionhistory_list: ["medication", "dose", "qty", "repeat", "creator", "note"],
  appt_list: ["apptstarttime", "apptprovidername", "apptreason", "apptroomname", "eventtype"],
  appttimingevent_list: ["starttime", "providername", "reason", "eventtype", "durationseconds"],
  task_list: ["taskfor", "creator", "completed", "urgentfg", "note"],
  smslog_list: ["prfusr_name", "message", "reply", "mobile"],
  graphic_list: ["description", "realname", "windocumenttype", "reviewed", "reviewby"],
  measurement_list: ["value", "unit", "note", "creator"],
  patientclinical_list: ["value", "description", "note", "creator"],
  consultationproblem_list: ["status", "note", "creator"],
  procedures_list: ["description", "doctorname", "note"],
  customform_list: ["description", "note", "creator"],
  addressbook_list: ["providername", "speciality", "email", "mobile"],
  interestedparty_list: ["description", "email", "mobile"],
  hiaudit_list: ["status", "description", "note", "creator"],
  hoimcclaim_list: ["status", "description", "note", "creator"],
};

const TECHNICAL_FIELD_PATTERN =
  /(^id$|uuid|_id_fk$|^pt_id_fk$|^ab_id_fk$|^apt_id_fk$|^prfusr_id_fk$|^graphic_id_fk$|^pathname$|^createdat$|^gcs_|^lastupdated$|^lastupdatedby$|^externalid$|^sync$|^archived$|^deleted$)/i;

export function XmlReaderApp() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [data, setData] = useState<ParsedGenieXml | null>(null);
  const [activeSectionKey, setActiveSectionKey] = useState<string>("patient");
  const [query, setQuery] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [patientIndex, setPatientIndex] = useState<PatientIndexEntry[]>([]);
  const [lettersIndex, setLettersIndex] = useState<Record<string, LetterIndexEntry>>({});
  const [patientSearch, setPatientSearch] = useState("");
  const [isLoadingPatient, setIsLoadingPatient] = useState(false);
  const [activePatientFile, setActivePatientFile] = useState<string | null>(null);

  useEffect(() => {
    getPatientIndex().then(setPatientIndex);
    getLettersIndex().then(setLettersIndex);
  }, []);

  const filteredPatients = useMemo(() => {
    const q = patientSearch.trim().toLowerCase();
    if (!q) return patientIndex.slice(0, 50); // show first 50 alphabetically
    return patientIndex.filter((p) =>
      p.name.toLowerCase().includes(q) ||
      p.urn.toLowerCase().includes(q) ||
      p.dob.includes(q) ||
      p.suburb.toLowerCase().includes(q)
    ).slice(0, 100);
  }, [patientIndex, patientSearch]);

  async function openPatientFile(entry: PatientIndexEntry) {
    setIsLoadingPatient(true);
    setActivePatientFile(entry.filePath);
    try {
      setError(null);
      const { name, content } = await readBackupFile(entry.filePath);
      const parsed = parseGenieXml(content, name);
      setData(parsed);
      setActiveSectionKey("patient");
      setQuery("");
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Could not open file.");
    } finally {
      setIsLoadingPatient(false);
    }
  }

  const activeSection = useMemo(() => {
    if (!data) {
      return null;
    }

    if (activeSectionKey === "patient") {
      return null;
    }

    if (activeSectionKey === "raw") {
      return null;
    }

    return data.sections.find((section) => section.key === activeSectionKey) ?? data.sections[0] ?? null;
  }, [activeSectionKey, data]);

  const filteredRecords = useMemo(() => {
    if (!activeSection) {
      return [];
    }

    const trimmedQuery = query.trim().toLowerCase();

    if (!trimmedQuery) {
      return activeSection.records;
    }

    return activeSection.records.filter((record) =>
      [record.title, record.subtitle ?? "", ...Object.values(record.fields)]
        .join("\n")
        .toLowerCase()
        .includes(trimmedQuery),
    );
  }, [activeSection, query]);

  function openPicker() {
    inputRef.current?.click();
  }

  async function handleFile(file: File) {
    try {
      setError(null);
      const text = await file.text();
      const parsed = parseGenieXml(text, file.name);
      setData(parsed);
      setActiveSectionKey("patient");
      setQuery("");
    } catch (caughtError) {
      const message =
        caughtError instanceof Error ? caughtError.message : "This file could not be opened.";
      setError(message);
      setData(null);
    }
  }

  async function onFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) {
      await handleFile(file);
    }
    event.target.value = "";
  }

  async function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) {
      await handleFile(file);
    }
  }

  const patientName = data?.patient.fullname || [data?.patient.firstname, data?.patient.surname].filter(Boolean).join(" ") || "Patient record";

  return (
    <main className="page-shell">
      <div className="page-frame">
        <section className="hero">
          <div className="hero-topline">
            <div>
              <p className="eyebrow">Genie XML Reader</p>
              <h1>Chamara Old CMG Notes</h1>
            </div>
            <div className="privacy-pill">Local browser parsing only</div>
          </div>
          <p>
            This viewer opens one Genie <code>patient_summary</code> XML export at a time,
            keeps the file in your browser, and presents the contents in a structured read-only layout.
          </p>
        </section>

        <div className="grid">
          <aside className="panel sidebar">
            <div className="sidebar-section">
              <div
                className={`dropzone ${isDragging ? "is-active" : ""}`}
                onDragEnter={(event) => {
                  event.preventDefault();
                  setIsDragging(true);
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  setIsDragging(true);
                }}
                onDragLeave={(event) => {
                  event.preventDefault();
                  setIsDragging(false);
                }}
                onDrop={onDrop}
              >
                <div>
                  <p className="label">Open XML File</p>
                  <strong>Drag a Genie XML file here</strong>
                  <p className="file-name">
                    Or choose a file manually. Nothing is uploaded anywhere in v1.
                  </p>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button className="button" onClick={openPicker} type="button">
                    Choose File
                  </button>
                  {data ? (
                    <button
                      className="button secondary"
                      onClick={() => {
                        setData(null);
                        setActiveSectionKey("patient");
                        setQuery("");
                        setError(null);
                      }}
                      type="button"
                    >
                      Clear
                    </button>
                  ) : null}
                </div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xml,text/xml,application/xml"
                  style={{ display: "none" }}
                  onChange={onFileChange}
                />
                {data ? <div className="file-name">{data.fileName}</div> : null}
              </div>
            </div>

            <div className="sidebar-section">
              <p className="label">Patient Directory ({patientIndex.length} patients)</p>
              <input
                className="search-input"
                type="search"
                placeholder="Search by name, URN, DOB, suburb..."
                value={patientSearch}
                onChange={(event) => setPatientSearch(event.target.value)}
                style={{ marginBottom: 8 }}
              />
              {patientSearch && (
                <p className="file-name" style={{ marginBottom: 8 }}>
                  {filteredPatients.length} result{filteredPatients.length !== 1 ? "s" : ""}
                  {filteredPatients.length === 100 ? " (showing max 100)" : ""}
                </p>
              )}
              <div className="tab-list" style={{ maxHeight: "400px", overflowY: "auto", border: "1px solid var(--border)", borderRadius: "var(--radius-m)", padding: "4px" }}>
                {filteredPatients.length === 0 && (
                  <div className="empty-state" style={{ padding: "12px", fontSize: "13px" }}>
                    {patientIndex.length === 0
                      ? "No index found. Run: npm run build-index"
                      : "No patients matched your search."}
                  </div>
                )}
                {filteredPatients.map((entry) => (
                  <button
                    key={entry.filePath}
                    className={`tab-button ${activePatientFile === entry.filePath ? "active" : ""}`}
                    type="button"
                    style={{ padding: "8px 12px", textAlign: "left", display: "block", width: "100%" }}
                    onClick={() => openPatientFile(entry)}
                    disabled={isLoadingPatient}
                  >
                    <div style={{ fontWeight: 600, fontSize: "13px" }}>
                      {entry.name}
                    </div>
                    <div className="tab-subtitle" style={{ fontSize: "11px", display: "flex", gap: "8px", flexWrap: "wrap" }}>
                      {entry.dob && <span>DOB: {entry.dob}</span>}
                      {entry.urn && <span>URN: {entry.urn}</span>}
                      {entry.suburb && <span>{entry.suburb}</span>}
                    </div>
                  </button>
                ))}
              </div>
            </div>

            {error ? (
              <div className="sidebar-section">
                <div className="warning">{error}</div>
              </div>
            ) : null}

            <div className="sidebar-section">
              <div className="stack">
                <div>
                  <p className="label">Find In Current View</p>
                  <input
                    className="search-input"
                    type="search"
                    placeholder="Search fields, notes, letters..."
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    disabled={!data || activeSectionKey === "patient" || activeSectionKey === "raw"}
                  />
                </div>

                {data ? (
                  <div className="summary-grid">
                    <div className="summary-card">
                      <span className="label">Loaded File</span>
                      <strong>{data.fileName}</strong>
                    </div>
                    <div className="summary-card">
                      <span className="label">Sections</span>
                      <strong>{data.sections.length + 2}</strong>
                    </div>
                    <div className="summary-card">
                      <span className="label">Records</span>
                      <strong>{data.sections.reduce((sum, section) => sum + section.records.length, 0)}</strong>
                    </div>
                  </div>
                ) : null}
              </div>
            </div>

            {data ? (
              <div className="sidebar-section">
                <p className="label">Sections</p>
                <div className="tab-list">
                  <TabButton
                    label="Patient"
                    subtitle="Demographics and core record details"
                    count={1}
                    isActive={activeSectionKey === "patient"}
                    onClick={() => setActiveSectionKey("patient")}
                  />
                  {data.sections.map((section) => (
                    <TabButton
                      key={section.key}
                      label={section.label}
                      subtitle={section.description}
                      count={section.records.length}
                      isActive={activeSectionKey === section.key}
                      onClick={() => setActiveSectionKey(section.key)}
                    />
                  ))}
                  <TabButton
                    label="Raw XML"
                    subtitle="Fallback view of the original file"
                    count={0}
                    isActive={activeSectionKey === "raw"}
                    onClick={() => setActiveSectionKey("raw")}
                  />
                </div>
              </div>
            ) : null}
          </aside>

          <section className="panel">
            {!data ? (
              <div className="content-section">
                <div className="empty-state">
                  Open one Genie XML file to begin. The first version is intentionally read-only and local-only.
                </div>
              </div>
            ) : activeSectionKey === "patient" ? (
              <>
                <div className="content-section">
                  <div className="content-header">
                    <div>
                      <p className="label">Patient Record</p>
                      <h2 className="patient-name">{patientName}</h2>
                    </div>
                    <div className="patient-meta">
                      <div className="meta-pill">{safeValue(data.patient.dob, "DOB unavailable", "dob")}</div>
                      <div className="meta-pill">{safeValue(data.patient.sex, "Sex unavailable")}</div>
                      <div className="meta-pill">{safeValue(data.patient.lastseendate, "No last seen date", "lastseendate")}</div>
                    </div>
                  </div>
                </div>

                <div className="content-section">
                  <p className="label">Core Patient Details</p>
                  <dl className="info-grid">
                    {PATIENT_FIELDS.map((field) => (
                      <InfoCard
                        key={field}
                        label={field}
                        value={data.patient[field]}
                      />
                    ))}
                  </dl>
                </div>

                <div className="content-section">
                  <p className="label">Export Metadata</p>
                  <dl className="info-grid">
                    {EXPORT_FIELDS.map((field) => (
                      <InfoCard
                        key={field}
                        label={field}
                        value={data.exportInfo[field]}
                      />
                    ))}
                  </dl>
                </div>
              </>
            ) : activeSectionKey === "raw" ? (
              <div className="content-section">
                <p className="label">Raw XML</p>
                <pre className="raw-xml">{data.rawXml}</pre>
              </div>
            ) : (
              <>
                <div className="content-section">
                  <div className="content-header">
                    <div>
                      <p className="label">{activeSection?.label}</p>
                      <h2 className="patient-name">{activeSection?.records.length ?? 0} records</h2>
                    </div>
                    <div className="patient-meta">
                      {activeSectionKey === "appt_list" ? (
                        <>
                          <div className="meta-pill" style={{ background: "#E8F5E9", color: "#2E7D32", fontWeight: 600 }}>
                            {activeSection?.records.filter((r) => r.isFuture).length ?? 0} upcoming
                          </div>
                          <div className="meta-pill">Future first, then past</div>
                        </>
                      ) : (
                        <div className="meta-pill">Newest first</div>
                      )}
                      <div className="meta-pill">{activeSection?.description}</div>
                      {query ? <div className="meta-pill">{filteredRecords.length} match(es)</div> : null}
                    </div>
                  </div>
                </div>

                <div className="content-section">
                      {activeSection && filteredRecords.length > 0 ? (
                    <div className="records-grid">
                      {filteredRecords.map((record) => (
                        <RecordCard
                          key={record.id}
                          record={record}
                          sectionKey={activeSection.key}
                          lettersIndex={lettersIndex}
                        />
                      ))}
                    </div>
                  ) : (
                    <div className="empty-state">
                      {query
                        ? "No records matched the current search."
                        : "This section exists in the file, but there are no records to display."}
                    </div>
                  )}
                </div>
              </>
            )}
          </section>
        </div>
      </div>
    </main>
  );
}

function TabButton({
  label,
  subtitle,
  count,
  isActive,
  onClick,
}: {
  label: string;
  subtitle: string;
  count: number;
  isActive: boolean;
  onClick: () => void;
}) {
  return (
    <button className={`tab-button ${isActive ? "active" : ""}`} type="button" onClick={onClick}>
      <div className="tab-title">
        <span>{label}</span>
        <span>{count > 0 ? count : ""}</span>
      </div>
      <div className="tab-subtitle">{subtitle}</div>
    </button>
  );
}

function InfoCard({ label, value }: { label: string; value?: string }) {
  return (
    <div className="info-card">
      <dt>{toLabel(label)}</dt>
      <dd>{safeValue(value, "Not provided", label)}</dd>
    </div>
  );
}

function RecordCard({
  record,
  sectionKey,
  lettersIndex,
}: {
  record: ParsedRecord;
  sectionKey: string;
  lettersIndex: Record<string, LetterIndexEntry>;
}) {
  const [isOpen, setIsOpen] = useState(record.isFuture === true);
  const isFutureAppt = record.isFuture === true;
  const primaryFieldOrder = SECTION_PRIMARY_FIELDS[sectionKey] ?? [];
  const allEntries = Object.entries(record.fields)
    .filter(([key, value]) => value && !isBinaryLikeField(key, value, record.preview))
    .sort(([leftKey], [rightKey]) => fieldRank(leftKey) - fieldRank(rightKey));
  const primaryEntries = [
    ...primaryFieldOrder
      .map((fieldName) => [fieldName, record.fields[fieldName]] as const)
      .filter((entry): entry is readonly [string, string] => Boolean(entry[1])),
    ...allEntries.filter(
      ([key, value]) =>
        !primaryFieldOrder.includes(key) &&
        !TECHNICAL_FIELD_PATTERN.test(key) &&
        !isLongBlockField(key, value),
    ),
  ].slice(0, 6);
  const compactMeta = primaryEntries.filter(([key]) => !isNarrativeField(key)).slice(0, 3);
  const technicalEntries = allEntries.filter(
    ([key]) => !primaryEntries.some(([primaryKey]) => primaryKey === key),
  );
  const narrativeEntry =
    primaryEntries.find(([key]) => isNarrativeField(key)) ??
    allEntries.find(([key]) => isNarrativeField(key));

  return (
    <article className="record-card" style={isFutureAppt ? { borderLeft: "3px solid #2E7D32", background: "#F1FDF2" } : undefined}>
      <header className="record-header">
        <div>
          {isFutureAppt ? (
            <span style={{ display: "inline-block", background: "#2E7D32", color: "#fff", fontSize: "10px", fontWeight: 700, padding: "2px 7px", borderRadius: "4px", letterSpacing: "0.05em", marginBottom: "4px" }}>
              ✓ UPCOMING
            </span>
          ) : null}
          <h3 className="record-title">{record.title}</h3>
          {record.subtitle ? <p className="record-subtitle">{record.subtitle}</p> : null}
          {narrativeEntry ? (
            <p className="record-summary">{truncateText(formatFieldValue(narrativeEntry[0], narrativeEntry[1]), 220)}</p>
          ) : null}
          {compactMeta.length > 0 ? (
            <div className="record-meta-row">
              {record.exportReadability ? (
                <span className={`meta-pill ${record.exportReadability.status === "metadata_only" ? "meta-pill-warning" : ""}`}>
                  {record.exportReadability.label}
                </span>
              ) : null}
              {compactMeta.map(([key, value]) => (
                <span className="meta-pill" key={key}>
                  {toLabel(key)}: {formatFieldValue(key, value)}
                </span>
              ))}
            </div>
          ) : null}
        </div>
        <div className="record-actions">
          {record.primaryDate ? (
            <div className="record-date-pill">{safeValue(record.primaryDate, undefined, "record_date")}</div>
          ) : null}
          <button className="button secondary" type="button" onClick={() => setIsOpen((current) => !current)}>
            {isOpen ? "Hide details" : "Open"}
          </button>
        </div>
      </header>

      {isOpen ? (
        <>
          {primaryEntries.length > 0 ? (
            <div className="field-grid">
              {primaryEntries.map(([key, value]) => (
                <div className="field" key={key}>
                  <div className="field-name">{toLabel(key)}</div>
                  <div className="field-value">{formatFieldValue(key, value)}</div>
                </div>
              ))}
            </div>
          ) : null}

          {record.exportReadability ? (
            <div className={`empty-state ${record.exportReadability.status === "metadata_only" ? "empty-state-warning" : ""}`}>
              {record.exportReadability.detail}
            </div>
          ) : null}

          {record.preview ? <EmbeddedPreviewSection record={record} /> : null}
          {sectionKey === "outgoingletter_list" ? (
            <LinkedLetterPreviewSection record={record} lettersIndex={lettersIndex} />
          ) : null}

          {technicalEntries.length > 0 ? (
            <details className="details-block">
              <summary>Show all fields</summary>
              <div className="field-grid">
                {technicalEntries.map(([key, value]) => (
                  <div className="field" key={key}>
                    <div className="field-name">{toLabel(key)}</div>
                    <div className="field-value">{formatFieldValue(key, value)}</div>
                  </div>
                ))}
              </div>
            </details>
          ) : null}
        </>
      ) : null}
    </article>
  );
}

function EmbeddedPreviewSection({ record }: { record: ParsedRecord }) {
  const preview = record.preview;
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const previewUrl = useObjectUrl(isPreviewOpen ? preview : undefined);

  if (!preview) {
    return null;
  }

  return (
    <div className="preview-block">
      <p className="label">Embedded Preview</p>
      {preview.type === "html" ? (
        <iframe className="preview-frame" srcDoc={preview.content} title={record.title} />
      ) : null}
      {preview.type === "text" ? (
        <pre className="raw-xml">{preview.content}</pre>
      ) : null}
      {preview.type === "binary" ? (
        <>
          <div className="preview-actions">
            <button className="button secondary" type="button" onClick={() => setIsPreviewOpen((current) => !current)}>
              {isPreviewOpen ? "Hide preview" : "Preview document"}
            </button>
            {previewUrl ? (
              <a className="button secondary" href={previewUrl} target="_blank" rel="noreferrer">
                Open {preview.displayMode === "pdf" ? "PDF" : preview.displayMode === "image" ? "Image" : "File"}
              </a>
            ) : null}
            {previewUrl ? (
              <a
                className="button secondary"
                href={previewUrl}
                download={preview.fileName || `${record.id}.${getFileExtension(preview.mimeType)}`}
              >
                Download
              </a>
            ) : null}
            {preview.fileName ? <span className="file-name">{preview.fileName}</span> : null}
          </div>
          {isPreviewOpen && preview.displayMode === "image" && previewUrl ? (
            <img className="preview-image" src={previewUrl} alt={record.title} />
          ) : null}
          {isPreviewOpen && preview.displayMode === "pdf" && previewUrl ? (
            <iframe className="preview-frame" src={previewUrl} title={record.title} />
          ) : null}
          {isPreviewOpen && preview.displayMode === "download" ? (
            <div className="empty-state">
              Embedded file detected. Preview is not available for this format yet, but the file can still be opened or downloaded.
            </div>
          ) : null}
        </>
      ) : null}
    </div>
  );
}

function LinkedLetterPreviewSection({
  record,
  lettersIndex,
}: {
  record: ParsedRecord;
  lettersIndex: Record<string, LetterIndexEntry>;
}) {
  const linkedLetterId = findLinkedLetterId(record.fields, lettersIndex);
  const linkedLetter = linkedLetterId ? lettersIndex[linkedLetterId] : undefined;
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<EmbeddedPreview>();
  const previewUrl = useObjectUrl(isPreviewOpen ? preview : undefined);

  if (!linkedLetter || record.preview) {
    return null;
  }

  const letter = linkedLetter;

  async function loadPreview() {
    if (isLoading) {
      return;
    }

    if (preview) {
      setIsPreviewOpen((current) => !current);
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      if (letter.mimeType === "application/pdf") {
        const base64 = await getLetterDownload(letter.fileId);
        setPreview({
          type: "binary",
          mimeType: "application/pdf",
          base64,
          fileName: letter.fileName,
          sourceField: "google_drive_letter",
          displayMode: "pdf",
        });
      } else if (isWordDocument(letter)) {
        const content = await getLetterHtml(letter.fileId, letter.fileName);
        setPreview({
          type: "html",
          content,
          sourceField: "google_drive_letter",
        });
      } else {
        const base64 = await getLetterPdf(letter.fileId);
        setPreview({
          type: "binary",
          mimeType: "application/pdf",
          base64,
          fileName: letter.fileName.replace(/\.[^.]+$/, ".pdf"),
          sourceField: "google_drive_letter",
          displayMode: "pdf",
        });
      }

      setIsPreviewOpen(true);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Could not load linked letter preview.");
    } finally {
      setIsLoading(false);
    }
  }

  async function downloadLetter() {
    try {
      const base64 = await getLetterDownload(letter.fileId);
      downloadBase64File(base64, letter.mimeType, letter.fileName);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Could not download the linked letter.");
    }
  }

  return (
    <div className="preview-block">
      <p className="label">Linked Letter</p>
      <div className="preview-actions">
        <button className="button secondary" type="button" onClick={loadPreview} disabled={isLoading}>
          {isLoading ? "Loading..." : isPreviewOpen ? "Hide letter" : "View letter"}
        </button>
        <button className="button secondary" type="button" onClick={downloadLetter}>
          Download original
        </button>
        {previewUrl ? (
          <a className="button secondary" href={previewUrl} target="_blank" rel="noreferrer">
            Open PDF
          </a>
        ) : null}
        {previewUrl ? (
          <a
            className="button secondary"
            href={previewUrl}
            download={preview?.fileName || `${record.id}.pdf`}
          >
            Download PDF
          </a>
        ) : null}
        <span className="file-name">{letter.fileName}</span>
      </div>
      {error ? <div className="warning">{error}</div> : null}
      {isPreviewOpen && preview?.type === "html" ? (
        <iframe className="preview-frame" srcDoc={preview.content} title={record.title} />
      ) : null}
      {isPreviewOpen && previewUrl ? (
        <iframe className="preview-frame" src={previewUrl} title={record.title} />
      ) : null}
    </div>
  );
}

function useObjectUrl(preview?: EmbeddedPreview) {
  const [previewUrl, setPreviewUrl] = useState<string>();

  useEffect(() => {
    if (!preview || preview.type !== "binary") {
      setPreviewUrl(undefined);
      return;
    }

    const binaryString = window.atob(preview.base64);
    const bytes = Uint8Array.from(binaryString, (character) => character.charCodeAt(0));
    const blob = new Blob([bytes], { type: preview.mimeType });
    const nextPreviewUrl = URL.createObjectURL(blob);
    setPreviewUrl(nextPreviewUrl);

    return () => {
      URL.revokeObjectURL(nextPreviewUrl);
    };
  }, [preview]);

  return previewUrl;
}

function safeValue(value?: string, fallback = "Not provided", fieldName?: string) {
  if (!value || value === "0000-00-00" || value === "0000-00-00T00:00:00") {
    return fallback;
  }

  return formatFieldValue(fieldName, value) || fallback;
}

function formatFieldValue(fieldName: string | undefined, value: string) {
  if (fieldName && isDateField(fieldName)) {
    return formatAustralianDate(value, fieldName) ?? value;
  }

  if (value === "true") {
    return "Yes";
  }

  if (value === "false") {
    return "No";
  }

  return value;
}

function toLabel(value: string) {
  return value
    .replace(/_/g, " ")
    .replace(/\b\w/g, (character) => character.toUpperCase());
}

function fieldRank(key: string) {
  const foundIndex = FIELD_PRIORITY.indexOf(key);
  return foundIndex === -1 ? FIELD_PRIORITY.length + key.charCodeAt(0) : foundIndex;
}

function getFileExtension(mimeType: string) {
  if (mimeType === "application/pdf") {
    return "pdf";
  }

  if (mimeType === "image/jpeg") {
    return "jpg";
  }

  if (mimeType === "image/png") {
    return "png";
  }

  if (mimeType === "image/gif") {
    return "gif";
  }

  return "bin";
}

function isWordDocument(letter: LetterIndexEntry) {
  return (
    letter.mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    letter.mimeType === "application/msword" ||
    /\.docx?$/i.test(letter.fileName)
  );
}

function downloadBase64File(base64: string, mimeType: string, fileName: string) {
  const binaryString = window.atob(base64);
  const bytes = Uint8Array.from(binaryString, (character) => character.charCodeAt(0));
  const blob = new Blob([bytes], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function findLinkedLetterId(
  fields: Record<string, string>,
  lettersIndex: Record<string, LetterIndexEntry>,
): string | undefined {
  const preferredKeys = [
    "letterid",
    "letter_id",
    "outgoingletterid",
    "outgoingletter_id",
    "documentid",
    "document_id",
    "graphicid",
    "graphic_id",
    "id",
  ];

  for (const key of preferredKeys) {
    const value = fields[key]?.trim();
    if (value && lettersIndex[value]) {
      return value;
    }
  }

  for (const value of Object.values(fields)) {
    if (!value) {
      continue;
    }

    const matches = value.match(/\b\d{4,}\b/g);
    if (!matches) {
      continue;
    }

    const matchedId = matches.find((candidate) => lettersIndex[candidate]);
    if (matchedId) {
      return matchedId;
    }
  }

  return undefined;
}

function isNarrativeField(fieldName: string) {
  return ["message", "note", "history", "plan", "examination", "result", "description", "reason"].includes(fieldName);
}

function isLongBlockField(fieldName: string, value: string) {
  return isNarrativeField(fieldName) && value.length > 160;
}

function truncateText(value: string, length: number) {
  if (value.length <= length) {
    return value;
  }

  return `${value.slice(0, length).trimEnd()}...`;
}
