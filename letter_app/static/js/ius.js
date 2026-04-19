let iusPatientsData = [];
let iusPatientSelectInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function setIusFieldValue(id, value) {
    document.getElementById(id).value = value || "";
}

function getSelectedIusPatient() {
    const name = iusPatientSelectInstance
        ? iusPatientSelectInstance.getValue()
        : document.getElementById("ius-patient-select").value;
    return iusPatientsData.find((patient) => patient.name === name) || null;
}

function applyIusPatientToForm(patient) {
    if (!patient) {
        return;
    }

    setIusFieldValue("ius-name", patient.name || "");
    setIusFieldValue("ius-dob", patient.dob || "");
    setIusFieldValue("ius-address", patient.full_address || patient.address || "");
    setIusFieldValue("ius-phone", patient.home_phone || patient.mobile_phone || patient.phone || "");
    setIusFieldValue("ius-email", patient.email || "");
    setIusFieldValue("ius-medicare", patient.medicare_number || patient.medicare || "");
}

async function loadIusDatabase() {
    const loader = document.getElementById("loading-indicator");
    loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        iusPatientsData = data.patients || [];

        const patientSelect = document.getElementById("ius-patient-select");
        patientSelect.innerHTML = '<option value="">Select Patient...</option>';
        iusPatientsData.forEach((patient) => {
            const option = document.createElement("option");
            option.value = patient.name;
            option.textContent = `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            patientSelect.appendChild(option);
        });

        iusPatientSelectInstance = new TomSelect("#ius-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: () => applyIusPatientToForm(getSelectedIusPatient()),
        });
    } catch (err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }

    loader.classList.add("hidden");
    setIusFieldValue("ius-date", todayIsoDate());
}

function collectIusReasons() {
    const reasons = [];
    if (document.getElementById("reason-symptoms").checked) {
        reasons.push("symptoms");
    }
    if (document.getElementById("reason-ibd").checked) {
        reasons.push("ibd");
    }
    if (document.getElementById("reason-incomplete-colonoscopy").checked) {
        reasons.push("incomplete_colonoscopy");
    }
    return reasons;
}

function collectIusPayload() {
    return {
        patient: {
            name: document.getElementById("ius-name").value.trim(),
            dob: document.getElementById("ius-dob").value,
            urn: document.getElementById("ius-urn").value.trim(),
            medicare_number: document.getElementById("ius-medicare").value.trim(),
            medicare: document.getElementById("ius-medicare").value.trim(),
            address: document.getElementById("ius-address").value.trim(),
            full_address: document.getElementById("ius-address").value.trim(),
            phone: document.getElementById("ius-phone").value.trim(),
            home_phone: document.getElementById("ius-phone").value.trim(),
            email: document.getElementById("ius-email").value.trim(),
        },
        request_date: document.getElementById("ius-date").value,
        urgent_referral: document.getElementById("ius-urgent").value,
        reasons: collectIusReasons(),
        medical_history: document.getElementById("ius-medical-history").value.trim(),
        abnormal_markers: document.getElementById("ius-abnormal-markers").value.trim(),
        abnormal_markers_result: document.getElementById("ius-abnormal-markers-result").value.trim(),
        faecal_calprotectin: document.getElementById("ius-calprotectin").value.trim(),
        faecal_calprotectin_result: document.getElementById("ius-calprotectin-result").value.trim(),
        pregnant: document.getElementById("ius-pregnant").value,
        gestation_weeks: document.getElementById("ius-gestation-weeks").value.trim(),
    };
}

async function generateIus() {
    const button = document.getElementById("generate-ius-btn");
    const status = document.getElementById("ius-status");
    const payload = collectIusPayload();

    status.classList.add("hidden");
    status.textContent = "";

    if (!payload.patient.name) {
        alert("Please enter the patient name.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating...";

    try {
        const res = await fetch("/api/generate-ius", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });

        const result = await res.json();
        if (!res.ok || !result.success) {
            throw new Error(result.error || "Server error");
        }

        status.textContent = "IUS request generated. Opening PDF preview...";
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
        button.innerText = "Generate IUS PDF";
    }
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadIusDatabase();
    const button = document.getElementById("generate-ius-btn");
    button.disabled = false;
    button.addEventListener("click", generateIus);
});
