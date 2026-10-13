let radiologyPatientsData = [];
let radiologyPatientSelectInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function getRadiologySelectedPatient() {
    const name = radiologyPatientSelectInstance
        ? radiologyPatientSelectInstance.getValue()
        : document.getElementById("radiology-patient-select").value;
    return radiologyPatientsData.find((patient) => patient.name === name) || null;
}

function setFieldValue(id, value) {
    document.getElementById(id).value = value || "";
}

function applyPatientToForm(patient) {
    if (!patient) {
        return;
    }

    setFieldValue("rad-name", patient.name || "");
    setFieldValue("rad-dob", patient.dob || "");
    setFieldValue("rad-address", patient.full_address || patient.address || "");
    setFieldValue("rad-phone", patient.home_phone || patient.mobile_phone || patient.phone || "");
    setFieldValue("rad-medicare", patient.medicare_number || patient.medicare || "");

    if (!document.getElementById("rad-date").value) {
        setFieldValue("rad-date", todayIsoDate());
    }
}

async function loadRadiologyDatabase() {
    const loader = document.getElementById("loading-indicator");
    loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        radiologyPatientsData = data.patients || [];

        const patientSelect = document.getElementById("radiology-patient-select");
        patientSelect.innerHTML = '<option value="">Select Patient...</option>';
        radiologyPatientsData.forEach((patient) => {
            const option = document.createElement("option");
            option.value = patient.name;
            option.textContent = `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            patientSelect.appendChild(option);
        });

        radiologyPatientSelectInstance = new TomSelect("#radiology-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyPatientToForm(getRadiologySelectedPatient()),
        });
    } catch (err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }

    loader.classList.add("hidden");
    setFieldValue("rad-date", todayIsoDate());
}

function collectRadiologyPayload() {
    return {
        patient: {
            name: document.getElementById("rad-name").value.trim(),
            dob: document.getElementById("rad-dob").value,
            address: document.getElementById("rad-address").value.trim(),
            full_address: document.getElementById("rad-address").value.trim(),
            phone: document.getElementById("rad-phone").value.trim(),
            home_phone: document.getElementById("rad-phone").value.trim(),
            medicare_number: document.getElementById("rad-medicare").value.trim(),
            medicare: document.getElementById("rad-medicare").value.trim(),
        },
        request_for: document.getElementById("rad-request-for").value.trim(),
        clinical_details: document.getElementById("rad-clinical-details").value.trim(),
        request_date: document.getElementById("rad-date").value,
        patient_category: document.getElementById("rad-category").value,
    };
}

async function generateRadiology() {
    const button = document.getElementById("generate-radiology-btn");
    const status = document.getElementById("radiology-status");
    const payload = collectRadiologyPayload();

    status.classList.add("hidden");
    status.textContent = "";

    if (!payload.patient.name) {
        alert("Please enter the patient name.");
        return;
    }

    if (!payload.request_for) {
        alert("Please enter what the radiology request is for.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating...";

    try {
        const res = await fetch("/api/generate-radiology", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        const result = await res.json();
        if (!res.ok || !result.success) {
            throw new Error(result.error || "Server error");
        }

        status.textContent = "Radiology request generated. Opening PDF preview...";
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
        button.innerText = "Generate Radiology PDF";
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

function clearOcrStatus() {
    const statusText = document.getElementById("ocr-status-text");
    if (!statusText) return;
    statusText.className = "ocr-status hidden";
    statusText.innerText = "";
}

function handleOcrResponse(data) {
    if (data.error) {
        showOcrStatus(`Error: ${data.error} ${data.details ? `(${data.details.join(', ')})` : ''}`, "error");
        return;
    }

    if (data.name) setFieldValue("rad-name", data.name);
    if (data.dob) setFieldValue("rad-dob", data.dob);
    if (data.medicare) setFieldValue("rad-medicare", data.medicare);
    if (data.phone) setFieldValue("rad-phone", data.phone);
    if (data.address) setFieldValue("rad-address", data.address);
    if (data.request_for) {
        const reqEl = document.getElementById("rad-request-for");
        if (reqEl && !reqEl.value.trim()) {
            reqEl.value = data.request_for;
        }
    }
    if (data.clinical_details) {
        const clinEl = document.getElementById("rad-clinical-details");
        if (clinEl && !clinEl.value.trim()) {
            clinEl.value = data.clinical_details;
        }
    }

    const patientName = (data.name || "").trim();
    if (patientName) {
        const patient = {
            name: patientName,
            dob: data.dob || "",
            medicare_number: data.medicare || "",
            medicare: data.medicare || "",
            phone: data.phone || "",
            home_phone: data.phone || "",
            address: data.address || "",
            full_address: data.address || "",
        };
        const existingIndex = radiologyPatientsData.findIndex(
            (item) => item.name && item.name.toLowerCase() === patientName.toLowerCase()
        );
        if (existingIndex >= 0) {
            radiologyPatientsData[existingIndex] = { ...radiologyPatientsData[existingIndex], ...patient };
        } else {
            radiologyPatientsData.push(patient);
        }

        if (radiologyPatientSelectInstance) {
            const label = `${patientName}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            const option = { value: patientName, text: label };
            if (radiologyPatientSelectInstance.options[patientName]) {
                radiologyPatientSelectInstance.updateOption(patientName, option);
            } else {
                radiologyPatientSelectInstance.addOption(option);
            }
            radiologyPatientSelectInstance.setValue(patientName, true);
        }
    }

    showOcrStatus(`Imported patient details successfully via ${data.provider || "AI"}!`, "success");
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
            dropzone.addEventListener(
                eventName,
                (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dropzone.classList.add("dragover");
                },
                false
            );
        });

        ["dragleave", "drop"].forEach((eventName) => {
            dropzone.addEventListener(
                eventName,
                (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dropzone.classList.remove("dragover");
                },
                false
            );
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
    await loadRadiologyDatabase();
    initOcrAutofill();
    const button = document.getElementById("generate-radiology-btn");
    button.disabled = false;
    button.addEventListener("click", generateRadiology);
});
