"use client";

import React, { useState } from "react";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";

export default function PublicEndoscopyPage() {
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [dob, setDob] = useState("");
  const [phone, setPhone] = useState("");
  const [urn, setUrn] = useState("");
  const [referralDate, setReferralDate] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  });

  const [location, setLocation] = useState<"mains" | "svop">("mains");
  const [procedure, setProcedure] = useState<"colonoscopy" | "gastroscopy" | "gascolon">("colonoscopy");

  const [hasDiabetes, setHasDiabetes] = useState<"no" | "yes" | "">("");
  const [diabetesMeds, setDiabetesMeds] = useState("");

  const [hasBloodThinners, setHasBloodThinners] = useState<"no" | "yes" | "">("");
  const [bloodThinnersMeds, setBloodThinnersMeds] = useState("");

  const [isGenerating, setIsGenerating] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ text: string; type: "success" | "error" } | null>(null);

  function formatDate(isoDate: string): string {
    if (!isoDate) return "";
    const p = isoDate.split("-");
    if (p.length === 3) return `${p[2]}/${p[1]}/${p[0]}`;
    return isoDate;
  }

  async function handleGenerate(e: React.FormEvent) {
    e.preventDefault();
    setStatusMessage(null);

    // Validation
    if (!firstName.trim() || !lastName.trim()) {
      setStatusMessage({ text: "Please enter patient first and last name.", type: "error" });
      return;
    }
    if (!dob) {
      setStatusMessage({ text: "Please enter date of birth.", type: "error" });
      return;
    }
    if (!phone.trim()) {
      setStatusMessage({ text: "Please enter mobile / phone number.", type: "error" });
      return;
    }
    if (!referralDate) {
      setStatusMessage({ text: "Please enter referral date.", type: "error" });
      return;
    }
    if (!hasDiabetes) {
      setStatusMessage({ text: "Please declare whether patient is taking diabetes medications (Yes or No).", type: "error" });
      return;
    }
    if (hasDiabetes === "yes" && !diabetesMeds.trim()) {
      setStatusMessage({ text: "Please enter the diabetes medications and plan.", type: "error" });
      return;
    }
    if (!hasBloodThinners) {
      setStatusMessage({ text: "Please declare whether patient is taking blood thinners (Yes or No).", type: "error" });
      return;
    }
    if (hasBloodThinners === "yes" && !bloodThinnersMeds.trim()) {
      setStatusMessage({ text: "Please enter the blood thinner details and plan.", type: "error" });
      return;
    }

    setIsGenerating(true);

    try {
      // Determine template filenames
      const page1Map = {
        gastroscopy: "1st_gas.pdf",
        colonoscopy: "1st_colon.pdf",
        gascolon: "1st_gascolon.pdf",
      };

      const page2Map = {
        mains: {
          gastroscopy: "2nd_mains_gas.pdf",
          colonoscopy: "2nd_mains_colon.pdf",
          gascolon: "2nd_mains_gascolon.pdf",
        },
        svop: {
          gastroscopy: "2nd_svop_gas.pdf",
          colonoscopy: "2nd_svop_colon.pdf",
          gascolon: "2nd_svop_gas_and_colon.pdf",
        },
      };

      const p1File = page1Map[procedure];
      const p2File = page2Map[location][procedure];

      // Fetch template PDFs
      const [p1Res, p2Res] = await Promise.all([
        fetch(`/templates/endoscopy/${p1File}`),
        fetch(`/templates/endoscopy/${p2File}`),
      ]);

      if (!p1Res.ok || !p2Res.ok) {
        throw new Error("Could not load PDF templates from server.");
      }

      const [p1Bytes, p2Bytes] = await Promise.all([
        p1Res.arrayBuffer(),
        p2Res.arrayBuffer(),
      ]);

      const doc = await PDFDocument.load(p1Bytes);
      const doc2 = await PDFDocument.load(p2Bytes);
      const [page2Copied] = await doc.copyPages(doc2, [0]);
      doc.addPage(page2Copied);

      const font = await doc.embedFont(StandardFonts.Helvetica);
      const page1 = doc.getPage(0);
      const page2 = doc.getPage(1);
      const H = 842;

      const fullName = `${firstName.trim()} ${lastName.trim()}`;
      const formattedDob = formatDate(dob);
      const formattedDate = formatDate(referralDate);

      // Helper to draw sticker
      const drawSticker = (page: typeof page1) => {
        const x = 440;
        const y = H - 140;
        const w = 135;
        const h = 104;

        page.drawRectangle({
          x,
          y,
          width: w,
          height: h,
          color: rgb(1, 1, 1),
          borderColor: rgb(0.85, 0.85, 0.85),
          borderWidth: 0.8,
        });

        let curY = y + h - 14;
        page.drawText(`Name: ${fullName}`, { x: x + 6, y: curY, size: 8.5, font });
        curY -= 12;
        page.drawText(`DOB: ${formattedDob}${urn.trim() ? "   URN: " + urn.trim() : ""}`, {
          x: x + 6,
          y: curY,
          size: 8.5,
          font,
        });
        curY -= 12;
        if (phone.trim()) {
          page.drawText(`Mob: ${phone.trim()}`, { x: x + 6, y: curY, size: 8.5, font });
        }
      };

      // 1. Page 1 Sticker
      drawSticker(page1);

      // 2. Page 1 Referring doctor date
      page1.drawRectangle({
        x: 355,
        y: H - 192,
        width: 95,
        height: 17,
        color: rgb(1, 1, 1),
      });
      page1.drawText(formattedDate, { x: 358, y: H - 188, size: 9, font });

      // 3. Page 1 Diabetes
      const diabText = hasDiabetes === "yes" ? diabetesMeds.trim() : "Nil / No diabetes medications";
      page1.drawRectangle({
        x: 80,
        y: H - 746,
        width: 268,
        height: 31,
        color: rgb(1, 1, 1),
        borderColor: rgb(0.85, 0.85, 0.85),
        borderWidth: 0.8,
      });
      page1.drawText(diabText.slice(0, 75), { x: 84, y: H - 732, size: 8, font });

      // 4. Page 1 Blood thinners
      const bloodText = hasBloodThinners === "yes" ? bloodThinnersMeds.trim() : "Nil / No blood thinners";
      page1.drawRectangle({
        x: 362,
        y: H - 796,
        width: 216,
        height: 103,
        color: rgb(1, 1, 1),
        borderColor: rgb(0.85, 0.85, 0.85),
        borderWidth: 0.8,
      });
      page1.drawText(bloodText.slice(0, 140), { x: 367, y: H - 710, size: 8, font });

      // 5. Page 2 Sticker
      drawSticker(page2);

      // 6. Page 2 Consent name and date
      page2.drawText(fullName, { x: 110, y: H - 694, size: 9.5, font });
      page2.drawText(formattedDate, { x: 455, y: H - 752, size: 9.5, font });

      const pdfBytes = await doc.save();
      const blob = new Blob([pdfBytes as Uint8Array<ArrayBuffer>], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);

      const safeName = fullName.replace(/[^a-zA-Z0-9_-]/g, "_");
      const filename = `${safeName}_SVHM_endoscopy_${referralDate}.pdf`;

      // Trigger instant download
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      setStatusMessage({ text: `Success! Form generated and downloaded: ${filename}`, type: "success" });
    } catch (err: unknown) {
      console.error(err);
      setStatusMessage({
        text: err instanceof Error ? err.message : "Failed to generate PDF.",
        type: "error",
      });
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <div style={{ maxWidth: 720, margin: "2rem auto", padding: "0 1rem", fontFamily: "sans-serif" }}>
      <div style={{ background: "#ffffff", borderRadius: 12, padding: "2rem", boxShadow: "0 4px 20px rgba(0,0,0,0.08)", border: "1px solid #e5e7eb" }}>
        <h1 style={{ margin: "0 0 0.5rem 0", fontSize: "1.6rem", color: "#111827" }}>
          St Vincent&apos;s Public Endoscopy Referral
        </h1>
        <p style={{ margin: "0 0 1.5rem 0", color: "#6b7280", fontSize: "0.95rem" }}>
          Fill in the required patient details below. All questions are compulsory except the UR number.
        </p>

        <form onSubmit={handleGenerate}>
          {/* Section: Patient Demographics */}
          <fieldset style={{ border: "1px solid #e5e7eb", borderRadius: 8, padding: "1.25rem", margin: "0 0 1.5rem 0" }}>
            <legend style={{ fontWeight: 600, color: "#374151", padding: "0 0.5rem" }}>Patient Details</legend>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem", marginBottom: "1rem" }}>
              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  First Name <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <input
                  type="text"
                  required
                  value={firstName}
                  onChange={(e) => setFirstName(e.target.value)}
                  placeholder="e.g. John"
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>

              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Last Name <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <input
                  type="text"
                  required
                  value={lastName}
                  onChange={(e) => setLastName(e.target.value)}
                  placeholder="e.g. Doe"
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem", marginBottom: "1rem" }}>
              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Date of Birth <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <input
                  type="date"
                  required
                  value={dob}
                  onChange={(e) => setDob(e.target.value)}
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>

              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Mobile / Phone <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <input
                  type="tel"
                  required
                  value={phone}
                  onChange={(e) => setPhone(e.target.value)}
                  placeholder="e.g. 0412 345 678"
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  UR Number <span style={{ color: "#6b7280", fontWeight: 400 }}>(Optional)</span>
                </label>
                <input
                  type="text"
                  value={urn}
                  onChange={(e) => setUrn(e.target.value)}
                  placeholder="e.g. 123456"
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>

              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Referral Date <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <input
                  type="date"
                  required
                  value={referralDate}
                  onChange={(e) => setReferralDate(e.target.value)}
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>
            </div>
          </fieldset>

          {/* Section: Procedure & Location */}
          <fieldset style={{ border: "1px solid #e5e7eb", borderRadius: 8, padding: "1.25rem", margin: "0 0 1.5rem 0" }}>
            <legend style={{ fontWeight: 600, color: "#374151", padding: "0 0.5rem" }}>Location & Procedure</legend>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Location <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <select
                  value={location}
                  onChange={(e) => setLocation(e.target.value as "mains" | "svop")}
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                >
                  <option value="mains">Mains</option>
                  <option value="svop">SVOP</option>
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Procedure <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <select
                  value={procedure}
                  onChange={(e) => setProcedure(e.target.value as "colonoscopy" | "gastroscopy" | "gascolon")}
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                >
                  <option value="colonoscopy">Colonoscopy</option>
                  <option value="gastroscopy">Gastroscopy</option>
                  <option value="gascolon">Gastroscopy & Colonoscopy</option>
                </select>
              </div>
            </div>
          </fieldset>

          {/* Section: Diabetes Medications */}
          <fieldset style={{ border: "1px solid #e5e7eb", borderRadius: 8, padding: "1.25rem", margin: "0 0 1.5rem 0" }}>
            <legend style={{ fontWeight: 600, color: "#374151", padding: "0 0.5rem" }}>
              Diabetes Medications <span style={{ color: "#dc2626" }}>*</span>
            </legend>
            <p style={{ margin: "0 0 0.75rem 0", fontSize: "0.85rem", color: "#4b5563" }}>
              Are you currently taking any diabetes medications?
            </p>

            <div style={{ display: "flex", gap: "1.5rem", marginBottom: hasDiabetes === "yes" ? "1rem" : 0 }}>
              <label style={{ display: "flex", alignItems: "center", gap: "0.4rem", cursor: "pointer", fontSize: "0.95rem" }}>
                <input
                  type="radio"
                  name="diabetes"
                  value="no"
                  checked={hasDiabetes === "no"}
                  onChange={() => setHasDiabetes("no")}
                  required
                />
                No
              </label>

              <label style={{ display: "flex", alignItems: "center", gap: "0.4rem", cursor: "pointer", fontSize: "0.95rem" }}>
                <input
                  type="radio"
                  name="diabetes"
                  value="yes"
                  checked={hasDiabetes === "yes"}
                  onChange={() => setHasDiabetes("yes")}
                  required
                />
                Yes
              </label>
            </div>

            {hasDiabetes === "yes" && (
              <div style={{ marginTop: "0.75rem" }}>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Specify Diabetes Medications & Instructions <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <textarea
                  required
                  rows={2}
                  value={diabetesMeds}
                  onChange={(e) => setDiabetesMeds(e.target.value)}
                  placeholder="e.g. Metformin 500mg bd. Withhold SGLT2i 3 days prior."
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>
            )}
          </fieldset>

          {/* Section: Blood Thinners */}
          <fieldset style={{ border: "1px solid #e5e7eb", borderRadius: 8, padding: "1.25rem", margin: "0 0 1.5rem 0" }}>
            <legend style={{ fontWeight: 600, color: "#374151", padding: "0 0.5rem" }}>
              Blood Thinners / Anticoagulants <span style={{ color: "#dc2626" }}>*</span>
            </legend>
            <p style={{ margin: "0 0 0.75rem 0", fontSize: "0.85rem", color: "#4b5563" }}>
              Are you currently taking any blood thinners or anticoagulants?
            </p>

            <div style={{ display: "flex", gap: "1.5rem", marginBottom: hasBloodThinners === "yes" ? "1rem" : 0 }}>
              <label style={{ display: "flex", alignItems: "center", gap: "0.4rem", cursor: "pointer", fontSize: "0.95rem" }}>
                <input
                  type="radio"
                  name="bloodThinners"
                  value="no"
                  checked={hasBloodThinners === "no"}
                  onChange={() => setHasBloodThinners("no")}
                  required
                />
                No
              </label>

              <label style={{ display: "flex", alignItems: "center", gap: "0.4rem", cursor: "pointer", fontSize: "0.95rem" }}>
                <input
                  type="radio"
                  name="bloodThinners"
                  value="yes"
                  checked={hasBloodThinners === "yes"}
                  onChange={() => setHasBloodThinners("yes")}
                  required
                />
                Yes
              </label>
            </div>

            {hasBloodThinners === "yes" && (
              <div style={{ marginTop: "0.75rem" }}>
                <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.3rem" }}>
                  Specify Blood Thinner & Withholding Instructions <span style={{ color: "#dc2626" }}>*</span>
                </label>
                <textarea
                  required
                  rows={2}
                  value={bloodThinnersMeds}
                  onChange={(e) => setBloodThinnersMeds(e.target.value)}
                  placeholder="e.g. Withhold Apixaban 48 hours prior to procedure."
                  style={{ width: "100%", padding: "0.6rem", borderRadius: 6, border: "1px solid #d1d5db" }}
                />
              </div>
            )}
          </fieldset>

          {/* Status Message */}
          {statusMessage && (
            <div
              style={{
                padding: "0.85rem",
                borderRadius: 6,
                marginBottom: "1.25rem",
                fontSize: "0.9rem",
                backgroundColor: statusMessage.type === "success" ? "#ecfdf5" : "#fef2f2",
                color: statusMessage.type === "success" ? "#065f46" : "#991b1b",
                border: `1px solid ${statusMessage.type === "success" ? "#a7f3d0" : "#fecaca"}`,
              }}
            >
              {statusMessage.text}
            </div>
          )}

          {/* Submit Button */}
          <button
            type="submit"
            disabled={isGenerating}
            style={{
              width: "100%",
              padding: "0.85rem",
              background: isGenerating ? "#9ca3af" : "#0d5c63",
              color: "#ffffff",
              border: "none",
              borderRadius: 8,
              fontSize: "1rem",
              fontWeight: 600,
              cursor: isGenerating ? "not-allowed" : "pointer",
              transition: "background 0.2s",
            }}
          >
            {isGenerating ? "Generating PDF..." : "Generate & Download Referral PDF"}
          </button>
        </form>
      </div>
    </div>
  );
}
