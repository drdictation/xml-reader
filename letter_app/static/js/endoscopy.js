let endoPatientsData = [];
let endoPatientSelectInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function setEndoFieldValue(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
}

function getSelectedEndoPatient() {
    const name = endoPatientSelectInstance
        ? endoPatientSelectInstance.getValue()
        : document.getElementById("endo-patient-select").value;
    return endoPatientsData.find((patient) => patient.name === name) || null;
}

function applyEndoPatientToForm(patient) {
    if (!patient) return;
    setEndoFieldValue("endo-name", patient.name || "");
    setEndoFieldValue("endo-dob", patient.dob || "");
    setEndoFieldValue("endo-phone", patient.mobile_phone || patient.phone || patient.home_phone || "");
    setEndoFieldValue("endo-urn", patient.urn || "");
    setEndoFieldValue("endo-medicare", patient.medicare_number || patient.medicare || "");
}

async function loadEndoDatabase() {
    const loader = document.getElementById("loading-indicator");
    if (loader) loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        endoPatientsData = data.patients || [];

        const patientSelect = document.getElementById("endo-patient-select");
        patientSelect.innerHTML = '<option value="">Select Patient...</option>';
        endoPatientsData.forEach((patient) => {
            const option = document.createElement("option");
            option.value = patient.name;
            option.textContent = `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            patientSelect.appendChild(option);
        });

        endoPatientSelectInstance = new TomSelect("#endo-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyEndoPatientToForm(getSelectedEndoPatient()),
        });
    } catch (err) {
        console.error("Error loading database:", err);
        if (loader) loader.innerText = "Error loading database";
        return;
    }

    if (loader) loader.classList.add("hidden");
    setEndoFieldValue("endo-date", todayIsoDate());
}

function collectEndoPayload() {
    return {
        patient: {
            name: document.getElementById("endo-name").value.trim(),
            dob: document.getElementById("endo-dob").value,
            urn: document.getElementById("endo-urn").value.trim(),
            phone: document.getElementById("endo-phone").value.trim(),
            mobile_phone: document.getElementById("endo-phone").value.trim(),
            medicare: document.getElementById("endo-medicare").value.trim(),
            medicare_number: document.getElementById("endo-medicare").value.trim(),
        },
        location: document.getElementById("endo-location").value,
        procedure: document.getElementById("endo-procedure").value,
        diabetes_plan: document.getElementById("endo-diabetes").value.trim(),
        blood_thinners_plan: document.getElementById("endo-blood-thinners").value.trim(),
        request_date: document.getElementById("endo-date").value,
    };
}

async function generateEndoscopy() {
    const button = document.getElementById("generate-endo-btn");
    const status = document.getElementById("endo-status");
    const payload = collectEndoPayload();

    status.classList.add("hidden");
    status.textContent = "";

    if (!payload.patient.name) {
        alert("Please enter the patient name.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating & Opening Outlook...";

    try {
        const res = await fetch("/api/generate-endoscopy", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        const result = await res.json();
        if (!res.ok || result.error) {
            throw new Error(result.error || "Generation failed.");
        }

        status.textContent = result.message || "Endoscopy referral generated. Check Outlook.";
        if (result.url) {
            status.innerHTML += ` <a href="${result.url}" target="_blank" style="color: inherit; text-decoration: underline; margin-left: 8px;">View PDF</a>`;
        }
        status.classList.remove("hidden", "is-error");
        status.classList.add("is-success");
    } catch (err) {
        status.textContent = err.message || "Generation failed.";
        status.classList.remove("hidden", "is-success");
        status.classList.add("is-error");
    } finally {
        button.disabled = false;
        button.innerText = "Generate & Open Outlook";
    }
}

function showOcrStatus(message, type) {
    const statusText = document.getElementById("ocr-status-text");
    if (!statusText) return;
    statusText.className = "ocr-status";
    statusText.classList.remove("hidden");
    statusText.innerText = message;
    if (type === "loading") {
        statusText.classList.add("is-loading");
    } else if (type === "success") {
        statusText.classList.add("is-success");
    } else if (type === "error") {
        statusText.classList.add("is-error");
    }
}

function handleOcrResponse(data) {
    if (data.error) {
        showOcrStatus(`Error: ${data.error} ${data.details ? `(${data.details.join(', ')})` : ''}`, "error");
        return;
    }

    if (data.name) setEndoFieldValue("endo-name", data.name);
    if (data.dob) setEndoFieldValue("endo-dob", data.dob);
    if (data.urn) setEndoFieldValue("endo-urn", data.urn);
    if (data.medicare) setEndoFieldValue("endo-medicare", data.medicare);
    if (data.phone) setEndoFieldValue("endo-phone", data.phone);

    showOcrStatus(`Imported patient details successfully via ${data.provider}!`, "success");
}

async function startOcrCapture() {
    showOcrStatus("Please drag a box over demographics on your screen...", "loading");
    try {
        const res = await fetch("/api/ocr-capture");
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Capture failed: ${err.message || err}`, "error");
    }
}

async function uploadOcrImage(file) {
    showOcrStatus("Processing image and extracting text...", "loading");
    const formData = new FormData();
    formData.append("image", file);

    try {
        const res = await fetch("/api/ocr-upload", {
            method: "POST",
            body: formData,
        });
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Upload failed: ${err.message || err}`, "error");
    }
}

async function uploadPastedText(text) {
    showOcrStatus("Processing referral text with AI...", "loading");
    try {
        const res = await fetch("/api/ocr-text", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: text }),
        });
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Extraction failed: ${err.message || err}`, "error");
    }
}

function initOcrAutofill() {
    const captureBtn = document.getElementById("ocr-capture-btn");
    const dropzone = document.getElementById("ocr-dropzone");
    const fileInput = document.getElementById("ocr-file-input");

    if (captureBtn) {
        captureBtn.addEventListener("click", startOcrCapture);
    }

    if (dropzone && fileInput) {
        dropzone.addEventListener("click", () => fileInput.click());

        fileInput.addEventListener("change", (e) => {
            if (e.target.files && e.target.files[0]) {
                uploadOcrImage(e.target.files[0]);
            }
        });

        ["dragenter", "dragover"].forEach((eventName) => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add("dragover");
            }, false);
        });

        ["dragleave", "drop"].forEach((eventName) => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.remove("dragover");
            }, false);
        });

        dropzone.addEventListener("drop", (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files && files[0]) {
                uploadOcrImage(files[0]);
            } else {
                const text = dt.getData("text");
                if (text && text.trim().length > 10) {
                    uploadPastedText(text.trim());
                }
            }
        });
    }

    document.addEventListener("paste", (e) => {
        const items = (e.clipboardData || e.originalEvent.clipboardData).items;
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.indexOf("image") === 0) {
                const file = items[i].getAsFile();
                uploadOcrImage(file);
                return;
            }
        }

        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : "";
        if (activeTag === "input" || activeTag === "textarea") {
            return;
        }

        const pastedText = (e.clipboardData || e.originalEvent.clipboardData).getData("text");
        if (pastedText && pastedText.trim().length > 10) {
            uploadPastedText(pastedText.trim());
        }
    });
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadEndoDatabase();
    initOcrAutofill();
    const button = document.getElementById("generate-endo-btn");
    if (button) {
        button.disabled = false;
        button.addEventListener("click", generateEndoscopy);
    }
});
