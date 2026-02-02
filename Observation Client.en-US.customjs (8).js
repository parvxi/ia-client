/**
 * PETROLUBE OBSERVATION CLIENT - BULLETPROOF VERSION
 * Multiple fallbacks to ensure mainContent shows
 */

// =============================================
// CONFIGURATION
// =============================================
const CONFIG = {
    OBSERVATIONS_API: '/_api/cr650_ia_observations',
    UPDATES_API: '/_api/cr650_iaclientupdates',
    FLOW_URL: 'https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/82cacb5cdfca45bf9a2543ddcb222f9d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RItBspZoJsn9HzqxR85unIfaEgFmZ4xnPRA4y86D0hc', // ✅ ADD THIS
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
    ALLOWED_FILE_TYPES: ['.pdf', '.doc', '.docx', '.xls', '.xlsx'],
    RISK_RATINGS: {
        1: { text: 'Critical', class: 'risk-critical' },
        2: { text: 'High', class: 'risk-high' },
        3: { text: 'Moderate', class: 'risk-moderate' },
        4: { text: 'Low', class: 'risk-low' }
    }
};

// =============================================
// STATE
// =============================================
let observationId = null;
let observationData = null;
let uploadedFiles = [];

// =============================================
// CRITICAL: SHOW MAIN CONTENT (MULTIPLE METHODS)
// =============================================
function forceShowMainContent() {
    console.log('🔥 FORCING MAIN CONTENT TO SHOW...');

    const main = document.getElementById('mainContent');
    if (!main) {
        console.error('❌ mainContent not found!');
        return false;
    }

    // Method 1: Add CSS class
    main.classList.add('show');

    // Method 2: Force inline styles (backup)
    main.style.display = 'block';
    main.style.visibility = 'visible';
    main.style.opacity = '1';

    // Method 3: Remove any hiding attributes
    main.removeAttribute('hidden');

    console.log('✅ MAIN CONTENT FORCED VISIBLE');
    console.log('   - Classes:', main.className);
    console.log('   - Display:', main.style.display);

    return true;
}

function hideLoadingScreen() {
    console.log('🗑️ Removing loading screen...');
    const loading = document.getElementById('loadingScreen');
    if (loading) {
        loading.remove();
        console.log('✅ Loading screen removed');
    }
}

// =============================================
// UTILITY FUNCTIONS
// =============================================
function formatDate(dateString) {
    if (!dateString) return '-';
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function getDaysOverdue(dueDate) {
    if (!dueDate) return 0;
    const today = new Date();
    const due = new Date(dueDate);
    const diff = Math.ceil((today - due) / (1000 * 60 * 60 * 24));
    return diff > 0 ? diff : 0;
}

function getUrlParam(param) {
    const params = new URLSearchParams(window.location.search);
    return params.get(param);
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

function showElement(id) {
    const el = document.getElementById(id);
    if (el) el.style.display = 'flex';
}

function hideElement(id) {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
}

// =============================================
// API FUNCTIONS
// =============================================
async function getToken() {
    if (typeof shell !== 'undefined' && shell.getTokenDeferred) {
        try {
            return await new Promise((resolve) => {
                const timeout = setTimeout(() => resolve(null), 5000);
                shell.getTokenDeferred()
                    .done(token => {
                        clearTimeout(timeout);
                        resolve(token);
                    })
                    .fail(() => {
                        clearTimeout(timeout);
                        resolve(null);
                    });
            });
        } catch (e) {
            console.warn('Shell token error:', e);
        }
    }

    const field = document.getElementById('__RequestVerificationToken');
    if (field && field.value) return field.value;

    return null;
}

async function apiGet(url) {
    console.log('API GET:', url);
    const response = await fetch(url, {
        method: 'GET',
        headers: { 'Accept': 'application/json' }
    });

    if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    return await response.json();
}

async function apiPost(url, data) {
    console.log('API POST:', url);
    const token = await getToken();

    const headers = {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    };

    if (token) {
        headers['__RequestVerificationToken'] = token;
    }

    const response = await fetch(url, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(data)
    });

    if (!response.ok) {
        const errorText = await response.text();
        console.error('API POST error response:', errorText);
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    // ✅ Dataverse often returns 204 No Content
    if (response.status === 204 || response.headers.get('Content-Length') === '0') {
        return null;
    }

    return await response.json();
}

async function apiPatch(url, data) {
    console.log('API PATCH:', url);
    const token = await getToken();

    const headers = {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    };

    if (token) {
        headers['__RequestVerificationToken'] = token;
    }

    const response = await fetch(url, {
        method: 'PATCH',
        headers: headers,
        body: JSON.stringify(data)
    });

    if (!response.ok) {
        const errorText = await response.text();
        console.error('API PATCH error response:', errorText);
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    // PATCH returns 204 — success
    return true;
}


// =============================================
// FILE UPLOAD VIA POWER AUTOMATE
// =============================================

// Convert file to Base64
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1]; // Remove data:type prefix
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// Upload files via Power Automate
async function uploadFilesViaFlow(observationId, uploadedBy, files) {
    if (files.length === 0) return;

    console.log(`📤 Uploading ${files.length} file(s) via Power Automate...`);

    try {
        // Convert all files to Base64
        const filesData = await Promise.all(
            files.map(async (file) => ({
                fileName: file.name,
                content: await fileToBase64(file)
            }))
        );

        // Prepare payload for Power Automate
        const payload = {
            observationId: observationId,
            observationName: observationData.cr650_name || `OBS_${observationId}`,
            uploadedBy: uploadedBy,
            files: filesData
        };

        console.log(`📦 Payload prepared with ${filesData.length} files`);

        // Call Power Automate flow
        const response = await fetch(CONFIG.FLOW_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error('❌ Flow response error:', errorText);
            throw new Error(`File upload failed: ${response.status} - ${response.statusText}`);
        }

        const result = await response.json();
        console.log('✅ Files uploaded successfully:', result);
        return result;

    } catch (error) {
        console.error('❌ File upload error:', error);
        throw error;
    }
}

// =============================================
// FILE UPLOAD HANDLING
// =============================================
function validateFile(file) {
    if (file.size > CONFIG.MAX_FILE_SIZE) {
        return { valid: false, error: `File "${file.name}" exceeds 10MB limit` };
    }

    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!CONFIG.ALLOWED_FILE_TYPES.includes(ext)) {
        return { valid: false, error: `File type "${ext}" not allowed. Use PDF, Word, or Excel files.` };
    }

    return { valid: true };
}

function addFileToList(file) {
    const validation = validateFile(file);
    if (!validation.valid) {
        alert(validation.error);
        return;
    }

    if (uploadedFiles.find(f => f.name === file.name)) {
        alert(`File "${file.name}" is already added`);
        return;
    }

    uploadedFiles.push(file);
    renderFileList();
}

function removeFile(fileName) {
    uploadedFiles = uploadedFiles.filter(f => f.name !== fileName);
    renderFileList();
}

function renderFileList() {
    const fileList = document.getElementById('fileList');
    if (!fileList) return;

    if (uploadedFiles.length === 0) {
        fileList.innerHTML = '';
        return;
    }

    fileList.innerHTML = uploadedFiles.map(file => `
        <div class="file-item">
            <div class="file-icon">
                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/>
                    <polyline points="13 2 13 9 20 9"/>
                </svg>
            </div>
            <div class="file-info">
                <div class="file-name">${file.name}</div>
                <div class="file-size">${formatFileSize(file.size)}</div>
            </div>
            <button type="button" class="file-remove" onclick="removeFile('${file.name}')">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `).join('');
}

function setupFileUpload() {
    const fileInput = document.getElementById('fileInput');
    const uploadArea = document.getElementById('fileUploadArea');

    if (!fileInput || !uploadArea) return;

    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        files.forEach(file => addFileToList(file));
        fileInput.value = '';
    });

    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');

        const files = Array.from(e.dataTransfer.files);
        files.forEach(file => addFileToList(file));
    });
}

window.removeFile = removeFile;

// =============================================
// RENDER OBSERVATION
// =============================================
function renderObservation(obs) {
    console.log('🎨 Rendering observation:', obs);

    let name = 'Valued Partner';
    if (obs.cr650_personresponsible) {
        name = obs.cr650_personresponsible;
    } else if (obs.cr650_email && obs.cr650_email.includes('@')) {
        name = obs.cr650_email.split('@')[0];
        name = name.charAt(0).toUpperCase() + name.slice(1);
    }
    setText('recipientName', name);

    const risk = CONFIG.RISK_RATINGS[obs.cr650_riskrating] || CONFIG.RISK_RATINGS[3];
    const badge = document.getElementById('riskBadge');
    if (badge) {
        badge.textContent = risk.text;
        badge.className = 'card-badge ' + risk.class;
    }

    setText('auditName', obs.cr650_auditname || '-');
    setText('auditDate', formatDate(obs.cr650_auditreportdate));
    setText('observation', obs.cr650_observation || '-');
    setText('details', obs.cr650_details || 'No additional details provided.');

    const mgmtResp = document.getElementById('managementResponse');
    if (mgmtResp) {
        if (obs.cr650_managementresponse && obs.cr650_managementresponse.trim()) {
            mgmtResp.textContent = obs.cr650_managementresponse;
        } else {
            mgmtResp.textContent = 'No management response recorded yet. Please provide your initial response in the update form below.';
            mgmtResp.style.fontStyle = 'italic';
            mgmtResp.style.color = '#737373';
        }
    }

    setText('dueDate', formatDate(obs.cr650_duedate));

    const daysOverdue = obs.cr650_daysoverdue || getDaysOverdue(obs.cr650_duedate);
    const overdueEl = document.getElementById('overdueIndicator');
    const overdueText = document.getElementById('overdueText');
    if (overdueEl && overdueText && daysOverdue > 0) {
        overdueText.textContent = `${daysOverdue} day${daysOverdue > 1 ? 's' : ''} overdue`;
        overdueEl.style.display = 'inline-flex';
    }

    const dateInput = document.getElementById('revisedDueDate');
    if (dateInput) {
        dateInput.min = new Date().toISOString().split('T')[0];
        if (obs.cr650_duedate) {
            dateInput.value = obs.cr650_duedate.split('T')[0];
        }
    }

    console.log('✅ Rendered successfully');
}

// =============================================
// FORM HANDLING
// =============================================
function setupForm() {
    const feedback = document.getElementById('revisedFeedback');
    const charCount = document.getElementById('charCount');
    if (feedback && charCount) {
        feedback.addEventListener('input', () => {
            charCount.textContent = feedback.value.length.toLocaleString();
        });
    }

    const comments = document.getElementById('clientComments');
    const commentsCount = document.getElementById('commentsCharCount');
    if (comments && commentsCount) {
        comments.addEventListener('input', () => {
            commentsCount.textContent = comments.value.length.toLocaleString();
        });
    }

    const cancelBtn = document.getElementById('cancelBtn');
    if (cancelBtn) {
        cancelBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to cancel? Your changes will be lost.')) {
                window.close();  // ✅ UPDATED
            }
        });
    }

    setupFileUpload();

    const form = document.getElementById('clientUpdateForm');
    if (form) {
        form.addEventListener('submit', handleSubmit);
    }

    console.log('✅ Form setup complete');
}

async function handleSubmit(e) {
    e.preventDefault();
    console.log('📝 Form submitted');

    const form = document.getElementById('clientUpdateForm');
    const submitBtn = document.getElementById('submitBtn');

    if (!form.checkValidity()) {
        alert('Please fill in all required fields.');
        return;
    }

    const formData = new FormData(form);

    let submittedBy = 'Client User';
    if (observationData.cr650_personresponsible) {
        submittedBy = observationData.cr650_personresponsible;
    } else if (observationData.cr650_email) {
        submittedBy = observationData.cr650_email;
    }

    const data = {
        'cr650_revisedmanagementfeedback': formData.get('revisedFeedback'),
        'cr650_revisedduedate': formData.get('revisedDueDate'),
        'cr650_clientcomments': formData.get('clientComments') || '',
        'cr650_submitteddate': new Date().toISOString(),
        'cr650_submittedby': submittedBy,
        'cr650_updatestatus': 1,
        'cr650_Observation@odata.bind': `/cr650_ia_observations(${observationId})`
    };

    console.log('📤 Submitting data:', data);

    if (uploadedFiles.length > 0) {
        console.log(`📎 ${uploadedFiles.length} file(s) to upload:`, uploadedFiles.map(f => f.name));
    }

    try {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `
            <div style="width: 20px; height: 20px; border: 2px solid white; border-top-color: transparent; border-radius: 50%; animation: spin 1s linear infinite;"></div>
            Submitting...
        `;

        // Step 1: Create client update record
        console.log('Creating client update record...');
        const updateResponse = await apiPost(CONFIG.UPDATES_API, data);
        console.log('✅ Client update created:', updateResponse);

        // Step 2: Upload files via Power Automate (if any)
        if (uploadedFiles.length > 0) {
            console.log('📤 Starting file upload process...');
            try {
                await uploadFilesViaFlow(observationId, submittedBy, uploadedFiles);
                console.log('✅ All files uploaded successfully');
            } catch (fileError) {
                console.error('⚠️ File upload failed:', fileError);
                // Don't fail the entire submission - just warn the user
                alert('Your update was submitted, but there was an issue uploading files:\n\n' + fileError.message + '\n\nPlease contact the Internal Audit team to manually upload your files.');
            }
        }

        // Step 3: Update observation status
        console.log('Updating observation status...');
        await apiPatch(`${CONFIG.OBSERVATIONS_API}(${observationId})`, {
            'cr650_status': 3
        });
        console.log('✅ Observation status updated');

        hideElement('mainContent');
        showElement('successMessage');

        // Auto-redirect to dashboard after 3 seconds
        setTimeout(() => {
            const userEmail = observationData.cr650_email || '';
            window.location.href = '/Client-Observation-Dashboard/?email=' + encodeURIComponent(userEmail);
        }, 3000);
        console.log('🎉 Submission complete!');

    } catch (error) {
        console.error('❌ Submit error:', error);

        let errorMsg = 'Failed to submit your update. ';

        if (error.message.includes('400')) {
            errorMsg += 'Invalid data format. Please check all fields and try again.';
        } else if (error.message.includes('401') || error.message.includes('403')) {
            errorMsg += 'Session expired. Please refresh the page and try again.';
        } else if (error.message.includes('404')) {
            errorMsg += 'Observation not found. The link may be invalid.';
        } else {
            errorMsg += 'Please try again or contact the Internal Audit team for assistance.';
        }

        alert(errorMsg + '\n\nError details: ' + error.message);

        submitBtn.disabled = false;
        submitBtn.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="20 6 9 17 4 12"/>
            </svg>
            Submit Update
        `;
    }
}

function showError(message) {
    hideLoadingScreen();

    const errorMessageEl = document.getElementById('errorMessage');
    if (errorMessageEl) {
        errorMessageEl.textContent = message;
    }

    showElement('errorScreen');
}

// =============================================
// INITIALIZATION
// =============================================

// Check if value looks like a GUID
function isGuid(value) {
    if (!value) return false;
    const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
    return guidRegex.test(value);
}

// Fetch observation by GUID or by reference name
async function fetchObservation(idOrName) {
    // If it's a GUID, use direct lookup
    if (isGuid(idOrName)) {
        console.log('📡 Fetching by GUID:', idOrName);
        return await apiGet(`${CONFIG.OBSERVATIONS_API}(${idOrName})`);
    }

    // Otherwise, search by cr650_name (reference like IA--0001)
    console.log('📡 Fetching by reference name:', idOrName);
    const result = await apiGet(`${CONFIG.OBSERVATIONS_API}?$filter=cr650_name eq '${encodeURIComponent(idOrName)}'`);

    if (!result.value || result.value.length === 0) {
        throw new Error(`Observation "${idOrName}" not found`);
    }

    // Return the first matching observation
    return result.value[0];
}

async function initialize() {
    console.log('🚀 Initializing...');

    try {
        let idParam = getUrlParam('id') || getUrlParam('observationId');
        console.log('📝 Observation ID/Reference:', idParam);

        if (!idParam) {
            throw new Error('No observation ID provided in URL');
        }

        console.log('📡 Loading observation data...');
        observationData = await fetchObservation(idParam);

        // Store the actual GUID for API calls
        observationId = observationData.cr650_ia_observationid;
        console.log('✅ Observation loaded:', observationData);
        console.log('📝 Resolved GUID:', observationId);

        console.log('🎨 Rendering observation...');
        renderObservation(observationData);

        console.log('⚙️ Setting up form...');
        setupForm();

        // CRITICAL: Show content with multiple methods
        hideLoadingScreen();
        forceShowMainContent();

        console.log('✅ Initialization complete!');

    } catch (error) {
        console.error('❌ Initialization failed:', error);
        showError(
            error.message.includes('No observation ID')
                ? 'No observation ID was provided. Please use the link from your email.'
                : 'Unable to load the observation. Please check the link or contact the Internal Audit team.'
        );
    }
}

// =============================================
// START APPLICATION
// =============================================
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
} else {
    initialize();
}