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

document.addEventListener("DOMContentLoaded", async () => {
    await loadRadiologyDatabase();
    const button = document.getElementById("generate-radiology-btn");
    button.disabled = false;
    button.addEventListener("click", generateRadiology);
});
