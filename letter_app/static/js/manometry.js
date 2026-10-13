let manometryPatientsData = [];
let manometryDoctorsData = [];
let manometryPatientSelectInstance = null;
let manometryReferrerSelectInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function setFieldValue(id, value) {
    document.getElementById(id).value = value || "";
}

function getSelectedPatient() {
    const name = manometryPatientSelectInstance
        ? manometryPatientSelectInstance.getValue()
        : document.getElementById("manometry-patient-select").value;
    return manometryPatientsData.find((patient) => patient.name === name) || null;
}

function getSelectedReferrer() {
    const name = manometryReferrerSelectInstance
        ? manometryReferrerSelectInstance.getValue()
        : document.getElementById("manometry-referrer-select").value;
    return manometryDoctorsData.find((doctor) => doctor.name === name) || null;
}

function applyPatientToForm(patient) {
    if (!patient) {
        return;
    }

    setFieldValue("manometry-name", patient.name || "");
    setFieldValue("manometry-dob", patient.dob || "");
    setFieldValue("manometry-urn", patient.urn || "");
    setFieldValue("manometry-address", patient.full_address || patient.address || "");
    setFieldValue("manometry-phone", patient.home_phone || patient.mobile_phone || patient.phone || "");
    setFieldValue("manometry-email", patient.email || "");

    if (!document.getElementById("manometry-date").value) {
        setFieldValue("manometry-date", todayIsoDate());
    }

    const referrerName = patient.referrer || patient.referring_doctor || "";
    if (referrerName && manometryReferrerSelectInstance) {
        manometryReferrerSelectInstance.setValue(referrerName, true);
        applyReferrerToForm(getSelectedReferrer() || { name: referrerName });
    }
}

function applyReferrerToForm(referrer) {
    if (!referrer) {
        return;
    }

    setFieldValue("referrer-name", referrer.name || "");
    setFieldValue("referrer-address", referrer.clinic || referrer.address || "");
}

function populateSelect(selectEl, items, labelFn) {
    selectEl.innerHTML = '<option value="">Select...</option>';
    items.forEach((item) => {
        const option = document.createElement("option");
        option.value = item.name;
        option.textContent = labelFn(item);
        selectEl.appendChild(option);
    });
}

async function loadDatabase() {
    const loader = document.getElementById("loading-indicator");
    loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        manometryPatientsData = data.patients || [];
        manometryDoctorsData = data.doctors || [];

        const patientSelect = document.getElementById("manometry-patient-select");
        populateSelect(patientSelect, manometryPatientsData, (patient) =>
            `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`
        );

        const referrerSelect = document.getElementById("manometry-referrer-select");
        populateSelect(referrerSelect, manometryDoctorsData, (doctor) =>
            `${doctor.name}${doctor.clinic ? ` - ${doctor.clinic}` : ""}`
        );

        manometryPatientSelectInstance = new TomSelect("#manometry-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyPatientToForm(getSelectedPatient()),
        });

        manometryReferrerSelectInstance = new TomSelect("#manometry-referrer-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyReferrerToForm(getSelectedReferrer()),
        });
    } catch (err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }

    loader.classList.add("hidden");
    setFieldValue("manometry-date", todayIsoDate());
}

function collectTests() {
    const tests = [];
    if (document.getElementById("test-oesophageal-manometry").checked) {
        tests.push("oesophageal_manometry");
    }
    if (document.getElementById("test-24-hour-ph-study").checked) {
        tests.push("24_hour_ph_study");
    }
    if (document.getElementById("test-24-ph-impedance").checked) {
        tests.push("24_ph_impedance");
    }
    return tests;
}

function collectIndications() {
    const indications = [];
    if (document.getElementById("ind-dysphagia").checked) indications.push("dysphagia");
    if (document.getElementById("ind-poem").checked) indications.push("consideration_for_poem");
    if (document.getElementById("ind-spasm").checked) indications.push("hx_of_oesophageal_spasm");
    if (document.getElementById("ind-heartburn").checked) indications.push("heart_burn");
    if (document.getElementById("ind-erosive").checked) indications.push("erosive_oesophagitis_despite_ppi");
    if (document.getElementById("ind-achalasia").checked) indications.push("hx_of_achalasia");
    if (document.getElementById("ind-regurgitation").checked) indications.push("regurgitation");
    if (document.getElementById("ind-fundoplication").checked) indications.push("consideration_for_fundoplication");
    if (document.getElementById("ind-motility").checked) indications.push("hx_of_motility_disorder");
    if (document.getElementById("ind-chest-pain").checked) indications.push("atypical_chest_pain");
    if (document.getElementById("ind-hellers").checked) indications.push("consideration_for_hellers_myotomy");
    if (document.getElementById("ind-post-fundo").checked) indications.push("post_fundoplication");
    return indications;
}

function collectPayload() {
    return {
        patient: {
            name: document.getElementById("manometry-name").value.trim(),
            dob: document.getElementById("manometry-dob").value,
            urn: document.getElementById("manometry-urn").value.trim(),
            address: document.getElementById("manometry-address").value.trim(),
            full_address: document.getElementById("manometry-address").value.trim(),
            phone: document.getElementById("manometry-phone").value.trim(),
            home_phone: document.getElementById("manometry-phone").value.trim(),
            email: document.getElementById("manometry-email").value.trim(),
        },
        request_date: document.getElementById("manometry-date").value,
        tests: collectTests(),
        on_ppi: document.getElementById("manometry-ppi").value,
        indications: collectIndications(),
        other_indication: document.getElementById("other-indication").value.trim(),
        clinical_details: document.getElementById("clinical-details").value.trim(),
        last_gastroscopy: document.getElementById("last-gastroscopy").value.trim(),
        gastroscopy_results: document.getElementById("gastroscopy-results").value.trim(),
        surgery_details: document.getElementById("surgery-details").value.trim(),
        medications: document.getElementById("medications").value.trim(),
        referrer: {
            name: document.getElementById("referrer-name").value.trim(),
            address: document.getElementById("referrer-address").value.trim(),
            clinic: document.getElementById("referrer-address").value.trim(),
            phone: document.getElementById("referrer-phone").value.trim(),
            department: document.getElementById("referrer-department").value.trim(),
            provider_number: document.getElementById("referrer-provider-number").value.trim(),
            fax: document.getElementById("referrer-fax").value.trim(),
        },
    };
}

async function generateManometry() {
    const button = document.getElementById("generate-manometry-btn");
    const status = document.getElementById("manometry-status");
    const payload = collectPayload();

    status.classList.add("hidden");
    status.textContent = "";

    if (!payload.patient.name) {
        alert("Please enter the patient name.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating...";

    try {
        const res = await fetch("/api/generate-manometry", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        const result = await res.json();
        if (!res.ok || !result.success) {
            throw new Error(result.error || "Server error");
        }

        status.textContent = "Manometry referral generated. Opening document...";
        status.classList.remove("hidden", "is-error");
        status.classList.add("is-success");
        if (result.url) {
            window.open(result.url, "_blank", "noopener");
        }
    } catch (err) {
        status.textContent = err.message || "Generation failed.";
        status.classList.remove("hidden", "is-success");
        status.classList.add("is-error");
    } finally {
        button.disabled = false;
        button.innerText = "Generate Manometry DOCX";
    }
}

function showOcrStatus(message, type) {
    const statusText = document.getElementById("ocr-status-text");
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

    if (data.name) setFieldValue("manometry-name", data.name);
    if (data.dob) setFieldValue("manometry-dob", data.dob);
    if (data.urn) setFieldValue("manometry-urn", data.urn);
    if (data.phone) setFieldValue("manometry-phone", data.phone);
    if (data.email) setFieldValue("manometry-email", data.email);
    if (data.address) setFieldValue("manometry-address", data.address);

    const patientName = (data.name || "").trim();
    if (patientName) {
        const patient = {
            name: patientName,
            dob: data.dob || "",
            urn: data.urn || "",
            medicare_number: data.medicare || "",
            medicare: data.medicare || "",
            private_health: data.private_health || "",
            phone: data.phone || "",
            home_phone: data.phone || "",
            email: data.email || "",
            address: data.address || "",
            full_address: data.address || "",
        };
        const existingIndex = manometryPatientsData.findIndex((item) => item.name === patientName);
        if (existingIndex >= 0) {
            manometryPatientsData[existingIndex] = { ...manometryPatientsData[existingIndex], ...patient };
        } else {
            manometryPatientsData.push(patient);
        }

        if (manometryPatientSelectInstance) {
            const label = `${patientName}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            const option = { value: patientName, text: label };
            if (manometryPatientSelectInstance.options[patientName]) {
                manometryPatientSelectInstance.updateOption(patientName, option);
            } else {
                manometryPatientSelectInstance.addOption(option);
            }
            manometryPatientSelectInstance.setValue(patientName, true);
        }
    }

    showOcrStatus(`Imported patient demographics successfully via ${data.provider || "OCR"}.`, "success");
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
            body: JSON.stringify({ text: text })
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

        fileInput.addEventListener("change", (event) => {
            if (event.target.files && event.target.files[0]) {
                uploadOcrImage(event.target.files[0]);
            }
        });

        ["dragenter", "dragover"].forEach((eventName) => {
            dropzone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dropzone.classList.add("dragover");
            }, false);
        });

        ["dragleave", "drop"].forEach((eventName) => {
            dropzone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dropzone.classList.remove("dragover");
            }, false);
        });

        dropzone.addEventListener("drop", (event) => {
            const dt = event.dataTransfer;
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

    document.addEventListener("paste", (event) => {
        const items = (event.clipboardData || event.originalEvent.clipboardData).items;
        for (let index = 0; index < items.length; index += 1) {
            if (items[index].type.indexOf("image") === 0) {
                const file = items[index].getAsFile();
                uploadOcrImage(file);
                return;
            }
        }

        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : "";
        if (activeTag === "input" || activeTag === "textarea") {
            return;
        }

        const pastedText = (event.clipboardData || event.originalEvent.clipboardData).getData("text");
        if (pastedText && pastedText.trim().length > 10) {
            uploadPastedText(pastedText.trim());
        }
    });
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadDatabase();
    initOcrAutofill();
    const button = document.getElementById("generate-manometry-btn");
    button.disabled = false;
    button.addEventListener("click", generateManometry);
});
