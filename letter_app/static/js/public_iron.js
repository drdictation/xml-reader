let ironPatientsData = [];
let ironPatientSelectInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function setIronFieldValue(id, value) {
    document.getElementById(id).value = value || "";
}

function getSelectedIronPatient() {
    const name = ironPatientSelectInstance
        ? ironPatientSelectInstance.getValue()
        : document.getElementById("iron-patient-select").value;
    return ironPatientsData.find((patient) => patient.name === name) || null;
}

function applyIronPatientToForm(patient) {
    if (!patient) {
        return;
    }

    setIronFieldValue("iron-name", patient.name || "");
    setIronFieldValue("iron-dob", patient.dob || "");
    setIronFieldValue("iron-address", patient.full_address || patient.address || "");
    setIronFieldValue("iron-phone", patient.home_phone || patient.mobile_phone || patient.phone || "");
    setIronFieldValue("iron-email", patient.email || "");
    setIronFieldValue("iron-medicare", patient.medicare_number || patient.medicare || "");
    setIronFieldValue("iron-private-health", patient.private_health || "");
}

async function loadIronDatabase() {
    const loader = document.getElementById("loading-indicator");
    loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        ironPatientsData = data.patients || [];

        const patientSelect = document.getElementById("iron-patient-select");
        patientSelect.innerHTML = '<option value="">Select Patient...</option>';
        ironPatientsData.forEach((patient) => {
            const option = document.createElement("option");
            option.value = patient.name;
            option.textContent = `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            patientSelect.appendChild(option);
        });

        ironPatientSelectInstance = new TomSelect("#iron-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyIronPatientToForm(getSelectedIronPatient()),
        });
    } catch (err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }

    loader.classList.add("hidden");
    setIronFieldValue("iron-date", todayIsoDate());
}

function collectIronPayload() {
    return {
        patient: {
            name: document.getElementById("iron-name").value.trim(),
            dob: document.getElementById("iron-dob").value,
            urn: document.getElementById("iron-urn").value.trim(),
            medicare_number: document.getElementById("iron-medicare").value.trim(),
            medicare: document.getElementById("iron-medicare").value.trim(),
            private_health: document.getElementById("iron-private-health").value.trim(),
            address: document.getElementById("iron-address").value.trim(),
            full_address: document.getElementById("iron-address").value.trim(),
            phone: document.getElementById("iron-phone").value.trim(),
            home_phone: document.getElementById("iron-phone").value.trim(),
            email: document.getElementById("iron-email").value.trim(),
        },
        request_date: document.getElementById("iron-date").value,
    };
}

async function generatePublicIron() {
    const button = document.getElementById("generate-iron-btn");
    const status = document.getElementById("iron-status");
    const payload = collectIronPayload();

    status.classList.add("hidden");
    status.textContent = "";

    if (!payload.patient.name) {
        alert("Please enter the patient name.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating...";

    try {
        const res = await fetch("/api/generate-public-iron", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        const result = await res.json();
        if (!res.ok || !result.success) {
            throw new Error(result.error || "Server error");
        }

        status.textContent = result.message || "Public Iron request generated. Check Outlook.";
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
    
    if (data.name) setIronFieldValue("iron-name", data.name);
    if (data.dob) setIronFieldValue("iron-dob", data.dob);
    if (data.urn) setIronFieldValue("iron-urn", data.urn);
    if (data.medicare) setIronFieldValue("iron-medicare", data.medicare);
    if (data.private_health) setIronFieldValue("iron-private-health", data.private_health);
    if (data.phone) setIronFieldValue("iron-phone", data.phone);
    if (data.email) setIronFieldValue("iron-email", data.email);
    if (data.address) setIronFieldValue("iron-address", data.address);
    
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
            body: formData
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
        
        fileInput.addEventListener("change", (e) => {
            if (e.target.files && e.target.files[0]) {
                uploadOcrImage(e.target.files[0]);
            }
        });
        
        ["dragenter", "dragover"].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add("dragover");
            }, false);
        });
        
        ["dragleave", "drop"].forEach(eventName => {
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
    await loadIronDatabase();
    initOcrAutofill();
    const button = document.getElementById("generate-iron-btn");
    button.disabled = false;
    button.addEventListener("click", generatePublicIron);
});
