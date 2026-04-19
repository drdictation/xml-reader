let patientsData = [];
let doctorsData = [];
let patientSelectInstance = null;
let copyReportsInstance = null;

function todayIsoDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function formatDisplayDate(value) {
    if (!value) return "-";
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    if (!match) return value;
    return `${match[3]}/${match[2]}/${match[1]}`;
}

function summaryValue(id, value) {
    const element = document.getElementById(id);
    element.textContent = value || "-";
}

function getSelectedPatient() {
    const name = patientSelectInstance ? patientSelectInstance.getValue() : document.getElementById("pathology-patient-select").value;
    return patientsData.find((patient) => patient.name === name) || null;
}

function updatePatientSummary() {
    const patient = getSelectedPatient();
    const requestDate = document.getElementById("request-date");

    if (!patient) {
        summaryValue("summary-name", "No patient selected");
        summaryValue("summary-dob", "-");
        summaryValue("summary-sex", "-");
        summaryValue("summary-phone", "-");
        summaryValue("summary-address", "-");
        summaryValue("summary-referrer", "-");
        summaryValue("summary-referral-date", "-");
        summaryValue("summary-referral-status", "-");
        if (!requestDate.value) {
            requestDate.value = todayIsoDate();
        }
        return;
    }

    summaryValue("summary-name", patient.name);
    summaryValue("summary-dob", formatDisplayDate(patient.dob));
    summaryValue("summary-sex", patient.sex || "-");
    summaryValue("summary-phone", patient.home_phone || patient.mobile_phone || patient.phone || "-");
    summaryValue("summary-address", patient.full_address || patient.address || "-");
    summaryValue("summary-referrer", patient.referring_doctor || "-");
    summaryValue("summary-referral-date", formatDisplayDate(patient.referral_date));
    summaryValue("summary-referral-status", patient.referral_status || "-");

    if (patient.referral_date) {
        requestDate.value = patient.referral_date;
    } else if (!requestDate.value) {
        requestDate.value = todayIsoDate();
    }

    if (copyReportsInstance && patient.referring_doctor) {
        copyReportsInstance.setValue(patient.referring_doctor, true);
    }
}

async function loadDatabase() {
    const loader = document.getElementById("loading-indicator");
    loader.classList.remove("hidden");

    try {
        const res = await fetch("/api/data");
        const data = await res.json();
        patientsData = data.patients || [];
        doctorsData = data.doctors || [];

        const patientSelect = document.getElementById("pathology-patient-select");
        patientSelect.innerHTML = '<option value="">Select Patient...</option>';
        patientsData.forEach((patient) => {
            const option = document.createElement("option");
            option.value = patient.name;
            option.textContent = `${patient.name}${patient.dob ? ` (DOB: ${patient.dob})` : ""}`;
            patientSelect.appendChild(option);
        });

        const copySelect = document.getElementById("copy-reports-select");
        copySelect.innerHTML = '<option value="">Select Doctor...</option>';
        doctorsData.forEach((doctor) => {
            const option = document.createElement("option");
            option.value = doctor.name;
            option.textContent = doctor.clinic ? `${doctor.name} - ${doctor.clinic}` : doctor.name;
            copySelect.appendChild(option);
        });

        patientSelectInstance = new TomSelect("#pathology-patient-select", {
            create: false,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
            onChange: updatePatientSummary,
        });

        copyReportsInstance = new TomSelect("#copy-reports-select", {
            create: true,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50,
        });
    } catch (err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }

    loader.classList.add("hidden");
    document.getElementById("request-date").value = todayIsoDate();
    updatePatientSummary();
}

async function generatePathology() {
    const button = document.getElementById("generate-pathology-btn");
    const status = document.getElementById("pathology-status");
    const patient = getSelectedPatient();
    const testsRequired = document.getElementById("tests-required").value.trim();
    const clinicalNotes = document.getElementById("clinical-notes").value.trim();
    const requestDate = document.getElementById("request-date").value;
    const copyReportsTo = copyReportsInstance ? copyReportsInstance.getValue() : document.getElementById("copy-reports-select").value;
    const copyDoctor = doctorsData.find((doctor) => doctor.name === copyReportsTo) || { name: copyReportsTo };

    status.classList.add("hidden");
    status.textContent = "";

    if (!patient) {
        alert("Please select a patient.");
        return;
    }

    if (!testsRequired) {
        alert("Please enter the requested tests.");
        return;
    }

    button.disabled = true;
    button.innerText = "Generating...";

    try {
        const res = await fetch("/api/generate-pathology", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                patient,
                tests_required: testsRequired,
                clinical_notes: clinicalNotes,
                request_date: requestDate,
                copy_reports_to: copyReportsTo,
                copy_doctor: copyDoctor,
            }),
        });

        const result = await res.json();
        if (!res.ok || !result.success) {
            throw new Error(result.error || "Server error");
        }

        status.textContent = "Pathology request generated. Opening PDF preview...";
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
        button.innerText = "Generate Pathology PDF";
    }
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadDatabase();
    const button = document.getElementById("generate-pathology-btn");
    button.disabled = false;
    button.addEventListener("click", generatePathology);
});
