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
    initOcrAutofill();
    
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
        window._tsPatient = new TomSelect('#patient-select', {
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

function showOcrStatus(message, type) {
    const statusText = document.getElementById('ocr-status-text');
    statusText.className = 'ocr-status';
    statusText.classList.remove('hidden');
    statusText.innerText = message;
    if (type === 'loading') {
        statusText.classList.add('is-loading');
    } else if (type === 'success') {
        statusText.classList.add('is-success');
    } else if (type === 'error') {
        statusText.classList.add('is-error');
    }
}

function upsertOcrPatient(data) {
    const patientName = (data.name || '').trim();
    if (!patientName) {
        showOcrStatus('No patient name found in the captured demographics.', 'error');
        return;
    }

    const patient = {
        name: patientName,
        dob: data.dob || '',
        urn: data.urn || '',
        medicare_number: data.medicare || '',
        medicare: data.medicare || '',
        private_health: data.private_health || '',
        phone: data.phone || '',
        home_phone: data.phone || '',
        email: data.email || '',
        address: data.address || '',
        full_address: data.address || '',
    };

    const existingIndex = patientsData.findIndex((item) => item.name === patientName);
    if (existingIndex >= 0) {
        patientsData[existingIndex] = { ...patientsData[existingIndex], ...patient };
    } else {
        patientsData.push(patient);
    }

    const label = `${patientName}${patient.dob ? ` (DOB: ${patient.dob})` : ''}`;
    if (window._tsPatient) {
        const option = { value: patientName, text: label };
        if (window._tsPatient.options[patientName]) {
            window._tsPatient.updateOption(patientName, option);
        } else {
            window._tsPatient.addOption(option);
        }
        window._tsPatient.setValue(patientName, true);
    } else {
        const patientSelect = document.getElementById('patient-select');
        let option = Array.from(patientSelect.options).find((item) => item.value === patientName);
        if (!option) {
            option = new Option(label, patientName);
            patientSelect.add(option);
        }
        patientSelect.value = patientName;
    }

    showOcrStatus(`Imported patient demographics successfully via ${data.provider || 'OCR'}.`, 'success');
}

function handleOcrResponse(data) {
    if (data.error) {
        showOcrStatus(`Error: ${data.error} ${data.details ? `(${data.details.join(', ')})` : ''}`, 'error');
        return;
    }

    upsertOcrPatient(data);
}

async function startOcrCapture() {
    showOcrStatus('Please drag a box over demographics on your screen...', 'loading');
    try {
        const res = await fetch('/api/ocr-capture');
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Capture failed: ${err.message || err}`, 'error');
    }
}

async function uploadOcrImage(file) {
    showOcrStatus('Processing image and extracting text...', 'loading');
    const formData = new FormData();
    formData.append('image', file);

    try {
        const res = await fetch('/api/ocr-upload', {
            method: 'POST',
            body: formData,
        });
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Upload failed: ${err.message || err}`, 'error');
    }
}

async function uploadPastedText(text) {
    showOcrStatus('Processing referral text with AI...', 'loading');
    try {
        const res = await fetch('/api/ocr-text', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: text })
        });
        const data = await res.json();
        handleOcrResponse(data);
    } catch (err) {
        showOcrStatus(`Extraction failed: ${err.message || err}`, 'error');
    }
}

function initOcrAutofill() {
    const captureBtn = document.getElementById('ocr-capture-btn');
    const dropzone = document.getElementById('ocr-dropzone');
    const fileInput = document.getElementById('ocr-file-input');

    if (captureBtn) {
        captureBtn.addEventListener('click', startOcrCapture);
    }

    if (dropzone && fileInput) {
        dropzone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (event) => {
            if (event.target.files && event.target.files[0]) {
                uploadOcrImage(event.target.files[0]);
            }
        });

        ['dragenter', 'dragover'].forEach((eventName) => {
            dropzone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dropzone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach((eventName) => {
            dropzone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dropzone.classList.remove('dragover');
            }, false);
        });

        dropzone.addEventListener('drop', (event) => {
            const dt = event.dataTransfer;
            const files = dt.files;
            if (files && files[0]) {
                uploadOcrImage(files[0]);
            } else {
                const text = dt.getData('text');
                if (text && text.trim().length > 10) {
                    uploadPastedText(text.trim());
                }
            }
        });
    }

    document.addEventListener('paste', (event) => {
        const items = (event.clipboardData || event.originalEvent.clipboardData).items;
        for (let index = 0; index < items.length; index += 1) {
            if (items[index].type.indexOf('image') === 0) {
                const file = items[index].getAsFile();
                uploadOcrImage(file);
                return;
            }
        }

        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
        const isEditable = document.activeElement && (document.activeElement.isContentEditable || activeTag === 'input' || activeTag === 'textarea');
        if (isEditable) {
            return;
        }

        const pastedText = (event.clipboardData || event.originalEvent.clipboardData).getData('text');
        if (pastedText && pastedText.trim().length > 10) {
            uploadPastedText(pastedText.trim());
        }
    });
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
