let patientsData = [];
let doctorsData = [];

function fmt(cmd) {
    document.execCommand(cmd, false, null);
    document.getElementById('body-text').focus();
}

function fmtPathology(cmd) {
    document.execCommand(cmd, false, null);
    document.getElementById('pathology-text').focus();
}

document.addEventListener('DOMContentLoaded', async () => {
    await Promise.all([loadTemplates(), loadDatabase()]);
    
    const btn = document.getElementById('generate-btn');
    btn.disabled = false;
    btn.addEventListener('click', generateLetter);
});

async function loadTemplates() {
    try {
        const res = await fetch('/api/templates');
        const templates = await res.json();
        
        const templateSelect = document.getElementById('template-select');
        let html = '<option value="">Select template...</option>';
        templates.forEach(t => {
            html += `<option value="${t}">${t}</option>`;
        });
        templateSelect.innerHTML = html;
        const ts = new TomSelect('#template-select', {create: false});

        // Auto-select test.pdf if it exists
        const defaultTemplate = templates.find(t => t.toLowerCase() === 'test.pdf');
        if (defaultTemplate) {
            ts.setValue(defaultTemplate);
        }
        
    } catch(err) {
        console.error("Error loading templates:", err);
    }
}

async function loadDatabase() {
    const loader = document.getElementById('loading-indicator');
    loader.classList.remove('hidden');
    try {
        const res = await fetch('/api/data');
        const data = await res.json();
        patientsData = data.patients || [];
        doctorsData = data.doctors || [];
        
        const patientSelect = document.getElementById('patient-select');
        let patHtml = '<option value="">Select Patient...</option>';
        patientsData.forEach(p => {
            patHtml += `<option value="${p.name}">${p.name} ${p.dob ? '(DOB: '+p.dob+')' : ''}</option>`;
        });
        patientSelect.innerHTML = patHtml;
        new TomSelect('#patient-select', {
            create: true,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50
        });
        
        const refSelect = document.getElementById('referrer-select');
        const ccSelect = document.getElementById('cc-select');
        let docHtml = '<option value="">Select Doctor...</option>';
        let ccHtml = '';
        
        doctorsData.forEach(d => {
            const display = `${d.name} ${d.clinic ? ' - '+d.clinic : ''}`;
            docHtml += `<option value="${d.name}">${display}</option>`;
            ccHtml += `<option value="${d.name}">${display}</option>`;
        });
        
        refSelect.innerHTML = docHtml;
        ccSelect.innerHTML = ccHtml;
        
        const tsCC = new TomSelect('#cc-select', {
            create: true,
            plugins: ['remove_button'],
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50
        });
        window._tsCC = tsCC;
        window._tsRef = new TomSelect('#referrer-select', {
            create: true,
            sortField: { field: "text", direction: "asc" },
            maxOptions: 50
        });
        
    } catch(err) {
        console.error("Error loading database:", err);
        loader.innerText = "Error loading database";
        return;
    }
    loader.classList.add('hidden');
}

async function generateLetter() {
    const btn = document.getElementById('generate-btn');
    let step = 'init';
    try {
        step = 'template';
        const template = document.getElementById('template-select').value;

        step = 'patient';
        const patientName = document.getElementById('patient-select').value;

        step = 'referrer';
        let referrerName = '';
        try {
            referrerName = window._tsRef ? window._tsRef.getValue() : document.getElementById('referrer-select').value;
        } catch(e) { referrerName = document.getElementById('referrer-select').value; }

        step = 'cc';
        let ccNames = [];
        try {
            ccNames = window._tsCC ? window._tsCC.getValue() : [];
            if (!Array.isArray(ccNames)) ccNames = ccNames ? [ccNames] : [];
        } catch(e) { ccNames = []; }

        step = 'body';
        const bodyEl = document.getElementById('body-text');
        const bodyText = bodyEl.innerHTML || '';
        const bodyPlain = (bodyEl.textContent || bodyEl.innerText || '').trim();

        step = 'pathology';
        const pathologyEl = document.getElementById('pathology-text');
        const pathologyText = pathologyEl ? (pathologyEl.innerHTML || '') : '';

        step = 'validate';
        if (!template || !patientName || !bodyPlain) {
            alert("Please select a template, a patient, and enter letter content.");
            return;
        }

        step = 'lookup';
        const patient = patientsData.find(p => p.name === patientName) || { name: patientName };
        const referrer = doctorsData.find(d => d.name === referrerName) || { name: referrerName };
        const cc_doctors = ccNames.map(n => doctorsData.find(d => d.name === n) || { name: n });

        step = 'fetch';
        btn.disabled = true;
        btn.innerText = "Generating...";

        const res = await fetch('/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ template, patient, referrer, cc_doctors, body_text: bodyText, pathology_text: pathologyText })
        });

        step = 'response';
        const result = await res.json();
        if (res.ok && result.success) {
            btn.innerText = "Done! Check Outlook";
            setTimeout(() => {
                btn.innerText = "Generate & Open Outlook";
                btn.disabled = false;
            }, 3000);
        } else {
            throw new Error(result.error || "Server error");
        }
    } catch(err) {
        alert("Generation failed at [" + step + "]: " + err.message);
        btn.disabled = false;
        btn.innerText = "Generate & Open Outlook";
    }
}
