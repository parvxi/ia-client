/**
 * PETROLUBE OBSERVATION TRACKER - COMPLETE ENHANCED VERSION
 * All 27+ fields with professional UI/UX, auto-calculations, and conditional logic
 * Version: 2.0
 */

// ================================
// CONFIGURATION
// ================================
const CONFIG = {
    DATAVERSE_ENDPOINT: '/_api/cr650_ia_observations',
    CLIENT_UPDATES_ENDPOINT: '/_api/cr650_iaclientupdates',
    PAGE_SIZE: 20,
    DASHBOARD_URL: '/observation-tracker-dashboard/',

    // Field mappings - UPDATED per user requirements
    RISK_MAP: {
        1: 'Critical',
        2: 'High',
        3: 'Moderate',
        4: 'Low'
    },

    // UPDATED Status mapping per user requirements
    STATUS_MAP: {
        1: 'In Progress',
        2: 'Overdue',
        3: 'Closed'
    },

    // UPDATED Aging mapping per user requirements
    AGING_MAP: {
        1: 'Not due',
        2: '0-6M',
        3: '6M-1Y',
        4: '1Y-2Y',
        5: 'Above 2Y'
    }
};

let agingTouched = false;

// ================================
// CSRF TOKEN HELPER
// ================================
async function getAntiForgeryToken() {
    if (typeof shell !== 'undefined' && typeof shell.getTokenDeferred === 'function') {
        return new Promise((resolve, reject) => {
            shell.getTokenDeferred()
                .done(token => resolve(token))
                .fail(() => resolve(null));
        });
    }
    const tokenField = document.getElementById('__RequestVerificationToken');
    if (tokenField && tokenField.value) {
        return tokenField.value;
    }
    return null;
}

async function safeFetch(url, options = {}) {
    const token = await getAntiForgeryToken();
    const headers = {
        'Accept': 'application/json',
        ...options.headers
    };
    if (token) {
        headers['__RequestVerificationToken'] = token;
    }
    return fetch(url, {
        ...options,
        headers
    });
}

// ================================
// STATE MANAGEMENT
// ================================
const AppState = {
    allObservations: [],
    filteredObservations: [],
    currentPage: 1,
    pageSize: CONFIG.PAGE_SIZE,
    sortBy: 'cr650_duedate',
    sortOrder: 'desc',
    filters: {
        search: '',
        year: '',
        status: '',
        risk: ''
    },
    panelMode: null,
    currentObservationId: null,
    currentObservation: null,
    latestClientUpdate: null,
    observationHistory: [],
    activeTab: 'edit',
    isDirty: false,
    isLoading: false
};

// ================================
// DATA LOADING
// ================================
async function loadObservations() {
    try {
        AppState.isLoading = true;
        showLoadingState();

        const response = await fetch(
            `${CONFIG.DATAVERSE_ENDPOINT}?$top=1000&$orderby=${AppState.sortBy} ${AppState.sortOrder === 'desc' ? 'desc' : 'asc'}`,
            {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        AppState.allObservations = data.value || [];
        const updatesRes = await fetch('/_api/cr650_iaclientupdates?$select=_cr650_observation_value');
        const updatesData = await updatesRes.json();

        AppState.clientUpdatedIds = new Set(
            updatesData.value.map(u => u._cr650_observation_value)
        );

        AppState.filteredObservations = [...AppState.allObservations];

        renderTable();
        updateStatistics();
        updatePagination();

    } catch (error) {
        console.error('Error loading observations:', error);
        showErrorState('Failed to load observations. Please refresh the page.');
    } finally {
        AppState.isLoading = false;
    }
}

async function loadObservationDetails(observationId) {
    try {
        const response = await fetch(
            `${CONFIG.DATAVERSE_ENDPOINT}(${observationId})`,
            {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const observation = await response.json();
        AppState.currentObservation = observation;
        const docs = await loadObservationDocuments(observationId);
        await loadLatestClientUpdate(observationId);
        await loadObservationHistory(observationId);
        return observation;

    } catch (error) {
        console.error('Error loading observation details:', error);
        alert('Failed to load observation details.');
        return null;
    }


}

async function loadLatestClientUpdate(observationId) {
    try {
        const response = await fetch(
            `${CONFIG.CLIENT_UPDATES_ENDPOINT}?$filter=_cr650_observation_value eq '${observationId}'&$orderby=cr650_submitteddate desc&$top=1`,
            {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        AppState.latestClientUpdate = data.value && data.value.length > 0 ? data.value[0] : null;

    } catch (error) {
        console.error('Error loading client update:', error);
        AppState.latestClientUpdate = null;
    }
}

async function loadObservationHistory(observationId) {
    try {
        const clientUpdatesResponse = await fetch(
            `${CONFIG.CLIENT_UPDATES_ENDPOINT}?$filter=_cr650_observation_value eq '${observationId}'&$orderby=cr650_submitteddate desc`,
            {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            }
        );

        const clientUpdates = clientUpdatesResponse.ok
            ? (await clientUpdatesResponse.json()).value || []
            : [];

        const timeline = [];

        if (AppState.currentObservation.createdon) {
            timeline.push({
                type: 'observation_created',
                actor: 'Internal Audit Team',
                actorRole: 'Auditor',
                date: AppState.currentObservation.createdon,
                action: 'Created observation record',
                description: `Initial draft created from ${AppState.currentObservation.cr650_auditname || 'audit report'}.`
            });
        }

        clientUpdates.forEach(update => {
            timeline.push({
                type: 'client_update',
                actor: update.cr650_submittedby || 'Client',
                actorRole: 'Client',
                date: update.cr650_submitteddate,
                action: 'Submitted revised feedback',
                description: update.cr650_clientcomments || 'Updated management response and requested extension.'
            });
        });

        if (AppState.currentObservation.modifiedon && AppState.currentObservation.modifiedon !== AppState.currentObservation.createdon) {
            timeline.push({
                type: 'status_change',
                actor: 'Internal Audit Team',
                actorRole: 'Auditor',
                date: AppState.currentObservation.modifiedon,
                action: 'Updated observation',
                description: `Status changed to ${CONFIG.STATUS_MAP[AppState.currentObservation.cr650_status] || 'Unknown'}.`
            });
        }

        timeline.sort((a, b) => new Date(b.date) - new Date(a.date));
        AppState.observationHistory = timeline;

    } catch (error) {
        console.error('Error loading observation history:', error);
        AppState.observationHistory = [];
    }
}

async function loadObservationDocuments(observationId) {

    const url = `/_api/cr650_ia_documentses?$filter=_cr650_observation_value eq '${observationId}'`;

    const response = await fetch(url, {
        method: 'GET',
        headers: { 'Accept': 'application/json' }
    });

    if (!response.ok) {
        console.error('Failed to load documents');
        return [];
    }

    const data = await response.json();

    return data.value.map(d => ({
        name: d.cr650_documentname,
        url: d.cr650_sharepointurl,
        uploadedBy: d.cr650_uploadedby,
        createdOn: d.createdon
    }));
}


function renderDocuments(docs) {

    const list = document.getElementById('documentsList');

    if (!list) {
        console.warn('documentsList not found (panel not ready yet)');
        return;
    }

    if (!docs.length) {
        list.innerHTML = '<p class="text-muted">No documents uploaded.</p>';
        return;
    }

    list.innerHTML = docs.map(doc => `
        <div class="document-row">
            📄 <a href="https://petrolubegroup.sharepoint.com/sites/IAAutomationPortal${doc.url}" 
                target="_blank">
                ${doc.name}
                </a>
            <div class="doc-info">
                Uploaded by ${doc.uploadedBy || 'Client'} • 
                ${new Date(doc.createdOn).toLocaleString()}
            </div>
        </div>
    `).join('');
}



// ================================
// TABLE RENDERING
// ================================
function renderTable() {
    const tbody = document.getElementById('observationsTableBody');
    if (!tbody) return;

    const start = (AppState.currentPage - 1) * AppState.pageSize;
    const end = start + AppState.pageSize;
    const pageData = AppState.filteredObservations.slice(start, end);

    if (pageData.length === 0) {
        tbody.innerHTML = `
            <tr class="empty-row">
                <td colspan="8">
                    <div class="empty-state">
                        <div class="empty-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none"
                                stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                                <rect width="8" height="4" x="8" y="2" rx="1" ry="1" />
                                <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
                                <path d="M9 12h6" />
                                <path d="M9 16h6" />
                            </svg>
                        </div>
                        <h3 class="empty-title">No observations found</h3>
                        <p class="empty-message">
                            ${AppState.filters.search || AppState.filters.year || AppState.filters.status || AppState.filters.risk
                ? 'Try adjusting your filters to find observations.'
                : 'Start by adding your first observation.'}
                        </p>
                        ${!AppState.filters.search && !AppState.filters.year && !AppState.filters.status && !AppState.filters.risk
                ? `<button class="btn btn-primary btn-empty-action" data-action="create-first">
                                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                                    stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M5 12h14" />
                                    <path d="M12 5v14" />
                                </svg>
                                <span>Add First Observation</span>
                            </button>`
                : ''}
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = pageData.map(obs => `
        <tr class="data-row" data-obs-id="${obs.cr650_ia_observationid}">
            <td class="col-period">
                <div class="period-cell">
                    <span class="period-year">${obs.cr650_year || ''}</span>
                    <span class="period-quarter">${obs.cr650_quarter || ''}</span>
                </div>
            </td>
            <td class="col-audit">
                <span class="audit-name">${escapeHtml(obs.cr650_auditname || '')}</span>
            </td>
            <td class="col-observation">
                <div class="observation-preview">${escapeHtml(truncate(obs.cr650_observation || '', 120))}</div>
            </td>
            <td class="col-risk">
                ${getRiskBadge(obs.cr650_riskrating)}
            </td>
            <td class="col-status">
                ${getStatusBadge(obs.cr650_status)}
                ${hasClientUpdate(obs.cr650_ia_observationid) 
                    ? `<div class="client-update-indicator">🔔 Client updated</div>` 
                    : ''}
            </td>

            <td class="col-responsible">
                <span>${escapeHtml(obs.cr650_personresponsible || '-')}</span>
            </td>
            <td class="col-due">
                <span class="due-date">${formatDate(obs.cr650_duedate)}</span>
                ${obs.cr650_daysoverdue && obs.cr650_daysoverdue > 0
            ? `<span class="overdue-badge" style="background: #EF4444; color: white; padding: 2px 8px; border-radius: 12px; font-size: 0.75rem; margin-left: 8px;">
                        ${obs.cr650_daysoverdue}d overdue
                    </span>`
            : ''}
            </td>
            <td class="col-actions">
                <div class="action-buttons">
                    <button class="btn-icon btn-edit" data-action="edit" title="Edit observation">
                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" />
                            <path d="m15 5 4 4" />
                        </svg>
                    </button>
                    <button class="btn-icon btn-delete" data-action="delete" title="Delete observation">
                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M3 6h18" />
                            <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6" />
                            <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2" />
                        </svg>
                    </button>
                </div>
            </td>
        </tr>
    `).join('');

    attachTableEventListeners();
}

function attachTableEventListeners() {
    const tbody = document.getElementById('observationsTableBody');
    if (!tbody) return;
    tbody.removeEventListener('click', handleTableClick);
    tbody.addEventListener('click', handleTableClick);
}

function handleTableClick(e) {
    const actionBtn = e.target.closest('[data-action]');
    if (!actionBtn) return;

    const row = actionBtn.closest('tr');
    const obsId = row?.dataset.obsId;
    if (!obsId) return;

    const action = actionBtn.dataset.action;

    if (action === 'edit') {
        openPanel('detail', obsId);
    } else if (action === 'delete') {
        deleteObservation(obsId);
    }
}

// ================================
// PANEL MANAGEMENT - COMPLETE FORM
// ================================
async function openPanel(mode, observationId = null) {
    agingTouched = false;
    AppState.panelMode = mode;
    AppState.currentObservationId = observationId;
    AppState.activeTab = 'edit';

    const panelOverlay = document.getElementById('panelOverlay');
    const panelContent = document.getElementById('panelContent');

    if (!panelOverlay || !panelContent) return;

    if (mode === 'create') {
        panelContent.innerHTML = renderCreateForm();
        setupFormEventListeners();
    } else if (mode === 'detail' && observationId) {
        panelContent.innerHTML = '<div class="loading-panel">Loading...</div>';
        const observation = await loadObservationDetails(observationId);
        if (observation) {
            panelContent.innerHTML = renderDetailPanel(observation);
            setupFormEventListeners();

            // 👉 NOW documents container exists — safe to render
            const docs = await loadObservationDocuments(observationId);
            renderDocuments(docs);
        }

    }

    panelOverlay.classList.add('active');
    // FIX: Removed - this was blocking panel scrolling
}

function closePanel() {
    const panelOverlay = document.getElementById('panelOverlay');
    if (!panelOverlay) return;

    if (AppState.isDirty) {
        if (!confirm('You have unsaved changes. Are you sure you want to close?')) {
            return;
        }
    }

    panelOverlay.classList.remove('active');
    // FIX: Removed - no longer needed
    AppState.panelMode = null;
    AppState.currentObservationId = null;
    AppState.isDirty = false;
}

function switchTab(tabName) {
    AppState.activeTab = tabName;
    const panelContent = document.getElementById('panelContent');
    if (!panelContent || !AppState.currentObservation) return;
    panelContent.innerHTML = renderDetailPanel(AppState.currentObservation);
    setupFormEventListeners();
}

// ================================
// FORM RENDERING - ALL 27+ FIELDS
// ================================
function renderCreateForm() {
    const today = new Date().toISOString().split('T')[0];
    const currentYear = new Date().getFullYear();

    return `
        <div class="panel-wrapper">
            <div class="panel-header">
                <div>
                    <h2 class="panel-title">Create New Observation</h2>
                    <p class="panel-subtitle">Complete all required fields to create a comprehensive audit observation record</p>
                </div>
                <button class="btn-icon-lg" onclick="closePanel()" title="Close panel">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M18 6 6 18" />
                        <path d="m6 6 12 12" />
                    </svg>
                </button>
            </div>

            <div class="panel-tabs">
                <button class="tab-btn active" data-tab="edit">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" />
                        <path d="m15 5 4 4" />
                    </svg>
                    Observation Details
                </button>
            </div>

            <div class="panel-body">
                <form id="observationForm" class="observation-form">
                    ${renderCompleteFormFields({}, 'create', today, currentYear)}
                </form>
            </div>

            <div class="panel-footer">
                <button type="button" class="btn btn-ghost" onclick="closePanel()">Cancel</button>
                <div class="footer-actions">
                    <button type="button" class="btn btn-secondary" onclick="saveObservation(true)">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
                            <polyline points="17 21 17 13 7 13 7 21" />
                            <polyline points="7 3 7 8 15 8" />
                        </svg>
                        Save as Draft
                    </button>
                    <button type="button" class="btn btn-primary" onclick="saveObservation(false)">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="m9 12 2 2 4-4" />
                            <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z" />
                        </svg>
                        Create Observation
                    </button>
                </div>
            </div>
        </div>
    `;
}

function renderDetailPanel(obs) {
    const clientUpdate = AppState.latestClientUpdate;
    const today = new Date().toISOString().split('T')[0];
    const currentYear = new Date().getFullYear();

    if (AppState.activeTab === 'history') {
        return renderHistoryTab(obs);
    }

    return `
        <div class="panel-wrapper">
            <div class="panel-header">
                <div>
                    <h2 class="panel-title">Edit Observation</h2>
                    <p class="panel-subtitle">ID: ${obs.cr650_name || obs.cr650_ia_observationid}</p>
                </div>
                <button class="btn-icon-lg" onclick="closePanel()" title="Close panel">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M18 6 6 18" />
                        <path d="m6 6 12 12" />
                    </svg>
                </button>
            </div>

            <div class="panel-tabs">
                <button class="tab-btn ${AppState.activeTab === 'edit' ? 'active' : ''}" onclick="switchTab('edit')">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" />
                        <path d="m15 5 4 4" />
                    </svg>
                    Edit Details
                </button>
                <button class="tab-btn ${AppState.activeTab === 'history' ? 'active' : ''}" onclick="switchTab('history')">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10"></circle>
                        <polyline points="12 6 12 12 16 14"></polyline>
                    </svg>
                    History (${AppState.observationHistory.length})
                </button>
            </div>

            <div class="panel-body">
                <form id="observationForm" class="observation-form">
                    ${renderCompleteFormFields(obs, 'edit', today, currentYear)}
                </form>
            </div>

            <div class="panel-footer">
                <button type="button" class="btn btn-ghost" onclick="closePanel()">Cancel</button>
                <div class="footer-actions">
                    <button type="button" class="btn btn-danger" onclick="deleteObservation('${obs.cr650_ia_observationid}')">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M3 6h18" />
                            <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6" />
                            <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2" />
                        </svg>
                        Delete
                    </button>
                    <button type="button" class="btn btn-primary" onclick="saveObservation(false)">
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="m9 12 2 2 4-4" />
                            <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z" />
                        </svg>
                        Save Changes
                    </button>
                </div>
            </div>
        </div>
    `;
}

/**
 * COMPREHENSIVE FORM FIELDS - ALL 27+ FIELDS
 * Organized in 5 logical sections with professional UX
 */
function renderCompleteFormFields(obs, mode, today, currentYear) {
    const isCreate = mode === 'create';
    const isClosed = obs.cr650_status === 3; // Status = Closed

    return `
        <!-- SECTION 1: AUDIT CONTEXT -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <rect width="18" height="18" x="3" y="4" rx="2" ry="2"></rect>
                        <line x1="16" x2="16" y1="2" y2="6"></line>
                        <line x1="8" x2="8" y1="2" y2="6"></line>
                        <line x1="3" x2="21" y1="10" y2="10"></line>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Audit Context</h3>
                    <p class="section-subtitle">Period, company, and audit information</p>
                </div>
            </div>

            <div class="form-row form-row-3">
                <div class="form-group">
                    <label class="form-label required">Year</label>
                    <input type="number" name="year" class="form-control" 
                        value="${obs.cr650_year || currentYear}" 
                        min="2020" max="2030" required
                        placeholder="e.g., ${currentYear}" />
                    <span class="field-hint">Audit year</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Quarter</label>
                    <select name="quarter" class="form-control" required>
                        <option value="">Select Quarter</option>
                        <option value="Q1" ${obs.cr650_quarter === 'Q1' ? 'selected' : ''}>Q1 (Jan-Mar)</option>
                        <option value="Q2" ${obs.cr650_quarter === 'Q2' ? 'selected' : ''}>Q2 (Apr-Jun)</option>
                        <option value="Q3" ${obs.cr650_quarter === 'Q3' ? 'selected' : ''}>Q3 (Jul-Sep)</option>
                        <option value="Q4" ${obs.cr650_quarter === 'Q4' ? 'selected' : ''}>Q4 (Oct-Dec)</option>
                    </select>
                    <span class="field-hint">Fiscal quarter</span>
                </div>
                <div class="form-group">
                    <label class="form-label">Month</label>
                    <select name="month" class="form-control">
                        <option value="">Select Month</option>
                        ${generateMonthOptions(obs.cr650_month)}
                    </select>
                    <span class="field-hint">Specific month (optional)</span>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Company Name</label>
                    <input type="text" name="companyName" class="form-control" 
                        value="${escapeHtml(obs.cr650_companyname || '')}" 
                        required
                        placeholder="e.g., Petrolube Ltd." />
                    <span class="field-hint">Legal entity name</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Region</label>
                    <input type="text" name="region" class="form-control" 
                        value="${escapeHtml(obs.cr650_region || '')}" 
                        required
                        placeholder="e.g., Middle East, Europe" />
                    <span class="field-hint">Geographic region</span>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Audit Name</label>
                    <input type="text" name="auditName" class="form-control" 
                        value="${escapeHtml(obs.cr650_auditname || '')}" 
                        required
                        placeholder="e.g., Operations Efficiency Audit 2025" />
                    <span class="field-hint">Full audit title</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Audit Report Date</label>
                    <input type="date" name="auditReportDate" class="form-control" 
                        value="${obs.cr650_auditreportdate ? obs.cr650_auditreportdate.split('T')[0] : ''}" 
                        required />
                    <span class="field-hint">Date audit report was issued</span>
                </div>
            </div>
        </div>

        <!-- SECTION 2: OBSERVATION DETAILS -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"></path>
                        <path d="M14 2v4a2 2 0 0 0 2 2h4"></path>
                        <path d="M9 13h6"></path>
                        <path d="M9 17h6"></path>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Observation Details</h3>
                    <p class="section-subtitle">Finding description, risk assessment, and management response</p>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Observation Type</label>
                    <select name="observationType" class="form-control" required>
                        <option value="">Select Type</option>
                        <option value="New" ${obs.cr650_observationtype === 'New' ? 'selected' : ''}>New</option>
                        <option value="Repeat" ${obs.cr650_observationtype === 'Repeat' ? 'selected' : ''}>Repeat</option>
                        <option value="Follow-up" ${obs.cr650_observationtype === 'Follow-up' ? 'selected' : ''}>Follow-up</option>
                    </select>
                    <span class="field-hint">Category of the finding</span>
                </div>
            </div>

            <div class="form-group">
                <label class="form-label required">Observation Description</label>
                <textarea name="observation" class="form-control" rows="4" required
                    placeholder="Provide a clear, concise description of what was observed during the audit...">${escapeHtml(obs.cr650_observation || '')}</textarea>
                <span class="field-hint">Main audit finding (be specific and objective)</span>
            </div>

            <div class="form-group">
                <label class="form-label required">Detailed Description</label>
                <textarea name="details" class="form-control" rows="4" required
                    placeholder="Expand on the observation with additional context, evidence, and potential impact...">${escapeHtml(obs.cr650_details || '')}</textarea>
                <span class="field-hint">Supporting details, context, and evidence</span>
            </div>

            <div class="form-group">
                <label class="form-label required">Risk Rating</label>
                <div class="risk-selector">
                    ${generateRiskSelector(obs.cr650_riskrating)}
                </div>
                <span class="field-hint">Assess the severity and potential impact</span>
            </div>

            <div class="form-group">
                <label class="form-label required">Management Response / Action Plan</label>
                <textarea name="managementResponse" class="form-control" rows="5" required
                    placeholder="Describe the corrective actions, responsible parties, and timelines agreed with management...">${escapeHtml(obs.cr650_managementresponse || '')}</textarea>
                <span class="field-hint">Client's proposed remediation plan</span>
            </div>
        </div>

        <!-- SECTION 3: RESPONSIBILITY & CONTACTS -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path>
                        <circle cx="9" cy="7" r="4"></circle>
                        <path d="M22 21v-2a4 4 0 0 0-3-3.87"></path>
                        <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Responsibility & Contacts</h3>
                    <p class="section-subtitle">Accountable parties and key contacts</p>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Department Responsible</label>
                    <input type="text" name="departmentResponsible" class="form-control" 
                        value="${escapeHtml(obs.cr650_departmentresponsible || '')}" 
                        required
                        placeholder="e.g., Finance, Operations, IT" />
                    <span class="field-hint">Primary department accountable</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Head of Department</label>
                    <input type="text" name="headOfDepartment" class="form-control" 
                        value="${escapeHtml(obs.cr650_headofdepartemt || '')}" 
                        required
                        placeholder="e.g., John Smith" />
                    <span class="field-hint">Department head name</span>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Person Responsible</label>
                    <input type="text" name="personResponsible" class="form-control" 
                        value="${escapeHtml(obs.cr650_personresponsible || '')}" 
                        required
                        placeholder="e.g., Jane Doe" />
                    <span class="field-hint">Primary contact for remediation</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Email</label>
                    <input type="email" name="email" class="form-control" 
                        value="${escapeHtml(obs.cr650_email || '')}" 
                        required
                        placeholder="e.g., jane.doe@petrolube.com" />
                    <span class="field-hint">Contact email address</span>
                </div>
            </div>

            <div class="form-group">
                <label class="form-label">Support Person</label>
                <input type="text" name="supportPerson" class="form-control" 
                    value="${escapeHtml(obs.cr650_supportperson || '')}" 
                    placeholder="e.g., Additional team member assisting" />
                <span class="field-hint">Secondary contact (optional)</span>
            </div>
        </div>

        <!-- SECTION 4: TIMELINE & STATUS TRACKING -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10"></circle>
                        <polyline points="12 6 12 12 16 14"></polyline>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Timeline & Status Tracking</h3>
                    <p class="section-subtitle">Due dates, aging, and current status</p>
                </div>
            </div>

            <div class="form-row form-row-3">
                <div class="form-group">
                    <label class="form-label required">Due Date</label>
                    <input type="date" name="dueDate" class="form-control" 
                        value="${obs.cr650_duedate ? obs.cr650_duedate.split('T')[0] : ''}" 
                        required
                        onchange="calculateDaysOverdue()" />
                    <span class="field-hint">Target completion date</span>
                </div>
                <div class="form-group">
                    <label class="form-label">Days Overdue</label>
                    <input type="number" id="daysOverdueField" name="daysOverdue" class="form-control" 
                        value="${obs.cr650_daysoverdue || 0}" 
                        readonly 
                        style="background: #f3f4f6;" />
                    <span class="field-hint">Auto-calculated</span>
                </div>
                <div class="form-group">
                    <label class="form-label required">Aging</label>
                    <select name="aging" class="form-control" required>
                        <option value="">Select Aging</option>
                        <option value="1" ${obs.cr650_aging === 1 ? 'selected' : ''}>Not due</option>
                        <option value="2" ${obs.cr650_aging === 2 ? 'selected' : ''}>0-6M</option>
                        <option value="3" ${obs.cr650_aging === 3 ? 'selected' : ''}>6M-1Y</option>
                        <option value="4" ${obs.cr650_aging === 4 ? 'selected' : ''}>1Y-2Y</option>
                        <option value="5" ${obs.cr650_aging === 5 ? 'selected' : ''}>Above 2Y</option>
                    </select>
                    <span class="field-hint">Time bucket classification</span>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label required">Status</label>
                    <select name="status" id="statusField" class="form-control" required onchange="toggleClosureFields()">
                        <option value="">Select Status</option>
                        <option value="1" ${obs.cr650_status === 1 ? 'selected' : ''}>In Progress</option>
                        <option value="2" ${obs.cr650_status === 2 ? 'selected' : ''}>Overdue</option>
                        <option value="3" ${obs.cr650_status === 3 ? 'selected' : ''}>Closed</option>
                    </select>
                    <span class="field-hint">Current observation status</span>
                </div>
                <div class="form-group" id="dateClosedGroup" style="display: ${isClosed ? 'block' : 'none'};">
                    <label class="form-label ${isClosed ? 'required' : ''}">Date Closed</label>
                    <input type="date" name="dateClosed" class="form-control" 
                        value="${obs.cr650_dateclosed ? obs.cr650_dateclosed.split('T')[0] : ''}" 
                        ${isClosed ? 'required' : ''} />
                    <span class="field-hint">Date observation was closed</span>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label class="form-label">Last Communication Date</label>
                    <input type="date" name="lastCommunicationDate" class="form-control" 
                        value="${obs.cr650_lastcommunicationdate ? obs.cr650_lastcommunicationdate.split('T')[0] : ''}" 
                        placeholder="Date of last follow-up" />
                    <span class="field-hint">Most recent contact date</span>
                </div>
                <div class="form-group">
                    <label class="form-label">Last Person Communicated With</label>
                    <input type="text" name="lastPersonCommunicated" class="form-control" 
                        value="${escapeHtml(obs.cr650_lastpersoncommunicatedwith || '')}" 
                        placeholder="e.g., John Smith" />
                    <span class="field-hint">Contact person from last interaction</span>
                </div>
            </div>
        </div>

        <!-- SECTION 5: FOLLOW-UP & CLOSURE -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z"></path>
                        <path d="m9 12 2 2 4-4"></path>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Follow-up & Closure</h3>
                    <p class="section-subtitle">Internal audit notes and closure documentation</p>
                </div>
            </div>

            <div class="form-group">
                <label class="form-label">IA Work / Internal Notes</label>
                <textarea name="iaWork" class="form-control" rows="3"
                    placeholder="Internal audit follow-up activities, validation work, or additional notes...">${escapeHtml(obs.cr650_iawork || '')}</textarea>
                <span class="field-hint">For internal audit team use</span>
            </div>

            <div class="form-group" id="closingRemarksGroup" style="display: ${isClosed ? 'block' : 'none'};">
                <label class="form-label ${isClosed ? 'required' : ''}">Closing Remarks</label>
                <textarea name="closingRemarks" class="form-control" rows="3" ${isClosed ? 'required' : ''}
                    placeholder="Summarize how the observation was resolved and any final notes...">${escapeHtml(obs.cr650_closingremarks || '')}</textarea>
                <span class="field-hint">Required when status is Closed</span>
            </div>

            <div class="form-group">
                <label class="form-label">Latest Revised MAP</label>
                <textarea name="latestRevisedMap" class="form-control" rows="3"
                    placeholder="Most recent management action plan updates or revisions...">${escapeHtml(obs.cr650_latestrevisedmap || '')}</textarea>
                <span class="field-hint">Updated action plan from client</span>
            </div>
        </div>

        <!-- SUPPORTING DOCUMENTS -->
        <div class="form-section">
            <div class="section-header">
                <div class="section-icon">📁</div>
                <div>
                    <h3 class="section-title">Supporting Documents</h3>
                    <p class="section-subtitle">Files uploaded for this observation</p>
                </div>
            </div>

            <div id="documentsList" class="documents-list">
                <!-- Filled by JS -->
            </div>
        </div>


        <!-- CLIENT UPDATE INFO (Read-only if exists) -->
        ${!isCreate && AppState.latestClientUpdate ? `
        <div class="form-section client-update-section">
            <div class="section-header">
                <div class="section-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                        <polyline points="17 10 12 15 7 10"></polyline>
                        <line x1="12" x2="12" y1="15" y2="3"></line>
                    </svg>
                </div>
                <div>
                    <h3 class="section-title">Pending Client Response</h3>
                    <p class="section-subtitle">Review and accept or reject the client's submission</p>
                </div>
            </div>

            <div class="client-update-card">
                <div class="form-group">
                    <label class="form-label">Revised Management Feedback</label>
                    <div class="readonly-field">${escapeHtml(AppState.latestClientUpdate.cr650_revisedmanagementfeedback || '-')}</div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label class="form-label">Revised Due Date</label>
                        <div class="readonly-field">${formatDate(AppState.latestClientUpdate.cr650_revisedduedate)}</div>
                    </div>
                    <div class="form-group">
                        <label class="form-label">Submitted Date</label>
                        <div class="readonly-field">${formatDate(AppState.latestClientUpdate.cr650_submitteddate)}</div>
                    </div>
                </div>
                <div class="form-group">
                    <label class="form-label">Client Comments</label>
                    <div class="readonly-field">${escapeHtml(AppState.latestClientUpdate.cr650_clientcomments || '-')}</div>
                </div>

                <!-- Accept/Reject Actions (only show if observation is not already Closed) -->
                ${obs.cr650_status !== 3 ? `
                <div class="client-response-actions" style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #e5e7eb; display: flex; gap: 12px;">
                    <button type="button" class="btn btn-success" onclick="acceptClientResponse('${obs.cr650_ia_observationid}')" style="background: #10B981; color: white; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px; font-weight: 500;">
                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="20 6 9 17 4 12"/>
                        </svg>
                        Accept & Close
                    </button>
                    <button type="button" class="btn btn-warning" onclick="rejectClientResponse('${obs.cr650_ia_observationid}')" style="background: #F59E0B; color: white; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px; font-weight: 500;">
                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
                            <path d="M3 3v5h5"/>
                        </svg>
                        Reject & Send Back
                    </button>
                </div>
                ` : `
                <div style="margin-top: 20px; padding: 12px; background: #D1FAE5; border-radius: 8px; color: #065F46;">
                    This observation has been closed.
                </div>
                `}
            </div>
        </div>
        ` : ''}
    `;
}

function renderHistoryTab(obs) {
    return `
        <div class="panel-wrapper">
            <div class="panel-header">
                <div>
                    <h2 class="panel-title">Observation History</h2>
                    <p class="panel-subtitle">Complete activity timeline</p>
                </div>
                <button class="btn-icon-lg" onclick="closePanel()" title="Close panel">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M18 6 6 18" />
                        <path d="m6 6 12 12" />
                    </svg>
                </button>
            </div>

            <div class="panel-tabs">
                <button class="tab-btn" onclick="switchTab('edit')">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" />
                        <path d="m15 5 4 4" />
                    </svg>
                    Edit Details
                </button>
                <button class="tab-btn active" onclick="switchTab('history')">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10"></circle>
                        <polyline points="12 6 12 12 16 14"></polyline>
                    </svg>
                    History (${AppState.observationHistory.length})
                </button>
            </div>

            <div class="panel-body">
                <div class="timeline">
                    ${AppState.observationHistory.length > 0 ? AppState.observationHistory.map(item => `
                        <div class="timeline-item timeline-${item.type}">
                            <div class="timeline-marker"></div>
                            <div class="timeline-content">
                                <div class="timeline-header">
                                    <div>
                                        <span class="timeline-actor">${item.actor}</span>
                                        <span class="timeline-role">${item.actorRole}</span>
                                    </div>
                                    <span class="timeline-date">${formatDateTime(item.date)}</span>
                                </div>
                                <div class="timeline-action">${item.action}</div>
                                <div class="timeline-description">${escapeHtml(item.description)}</div>
                            </div>
                        </div>
                    `).join('') : `
                        <div class="empty-timeline">
                            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none"
                                stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                                <circle cx="12" cy="12" r="10"></circle>
                                <polyline points="12 6 12 12 16 14"></polyline>
                            </svg>
                            <p>No history available</p>
                        </div>
                    `}
                </div>
            </div>

            <div class="panel-footer">
                <button type="button" class="btn btn-ghost" onclick="closePanel()">Close</button>
            </div>
        </div>
    `;
}

// ================================
// FORM UTILITIES & CALCULATIONS
// ================================

function hasClientUpdate(obsId) {
    return AppState.clientUpdatedIds?.has(obsId);
}


function calculateAging(daysOverdue) {
    if (daysOverdue <= 0) return 1;      // Not due
    if (daysOverdue <= 180) return 2;    // 0–6M
    if (daysOverdue <= 365) return 3;    // 6M–1Y
    if (daysOverdue <= 730) return 4;    // 1Y–2Y
    return 5;                            // Above 2Y
}


/**
 * Calculate days overdue based on due date
 */
function calculateDaysOverdue() {
    const dueDateField = document.querySelector('input[name="dueDate"]');
    const daysOverdueField = document.getElementById('daysOverdueField');
    const agingField = document.querySelector('[name="aging"]');

    if (!dueDateField || !daysOverdueField || !agingField) return;

    const dueDate = new Date(dueDateField.value);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dueDate.setHours(0, 0, 0, 0);

    const diffTime = today - dueDate;
    const daysOverdue = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    const finalDays = daysOverdue > 0 ? daysOverdue : 0;

    daysOverdueField.value = finalDays;

    // ✅ AUTO-SET AGING
    if (!agingTouched) {
        agingField.value = calculateAging(finalDays).toString();
    }
}

function toggleClosureFields() {
    const statusField = document.getElementById('statusField');
    const agingField = document.querySelector('[name="aging"]');
    const dateClosedGroup = document.getElementById('dateClosedGroup');
    const closingRemarksGroup = document.getElementById('closingRemarksGroup');
    const dateClosedInput = document.querySelector('input[name="dateClosed"]');
    const closingRemarksInput = document.querySelector('textarea[name="closingRemarks"]');

    if (!statusField) return;

    const isClosed = statusField.value === '3';

    // 🔹 UI logic
    if (dateClosedGroup) {
        dateClosedGroup.style.display = isClosed ? 'block' : 'none';
        if (dateClosedInput) {
            dateClosedInput.required = isClosed;
            if (isClosed && !dateClosedInput.value) {
                dateClosedInput.value = new Date().toISOString().split('T')[0];
            }
        }
    }

    if (closingRemarksGroup) {
        closingRemarksGroup.style.display = isClosed ? 'block' : 'none';
        if (closingRemarksInput) {
            closingRemarksInput.required = isClosed;
        }
    }

    // 🔹 Aging rule
    if (isClosed && agingField) {
        agingField.value = '1'; // Not due
    }
}


/**
 * Setup form event listeners
 */
function setupFormEventListeners() {
    const form = document.getElementById('observationForm');
    if (!form) return;

    // Mark form as dirty when any field changes
    form.addEventListener('change', () => {
        AppState.isDirty = true;
    });

    // Calculate days overdue when due date changes
    const dueDateField = form.querySelector('input[name="dueDate"]');
    if (dueDateField) {
        dueDateField.addEventListener('change', calculateDaysOverdue);
        // Initial calculation
        calculateDaysOverdue();
    }

    // Toggle closure fields when status changes
    const statusField = form.querySelector('select[name="status"]');
    if (statusField) {
        statusField.addEventListener('change', toggleClosureFields);
    }

    const agingField = form.querySelector('select[name="aging"]');
    if (agingField) {
        agingField.addEventListener('change', () => {
            agingTouched = true;
        });
    }

}

// ================================
// SAVE & DELETE OPERATIONS
// ================================
async function saveObservation(isDraft = false) {
    const form = document.getElementById('observationForm');
    if (!form) {
        console.error('❌ Form not found!');
        return;
    }

    if (!form.checkValidity()) {
        form.reportValidity();
        return;
    }

    console.log('📝 Starting save process...');

    const formData = new FormData(form);
    const data = {};

    try {
        // ================= REQUIRED FIELDS =================

        const year = formData.get('year');
        if (!year) throw new Error('Year is required');
        data.cr650_year = formData.get('year').toString();

        const quarter = formData.get('quarter');
        if (!quarter) throw new Error('Quarter is required');
        data.cr650_quarter = quarter;

        const companyName = formData.get('companyName');
        if (!companyName) throw new Error('Company Name is required');
        data.cr650_companyname = companyName;

        const region = formData.get('region');
        if (!region) throw new Error('Region is required');
        data.cr650_region = region;

        const auditName = formData.get('auditName');
        if (!auditName) throw new Error('Audit Name is required');
        data.cr650_auditname = auditName;

        const auditReportDate = formData.get('auditReportDate');
        if (!auditReportDate) throw new Error('Audit Report Date is required');
        data.cr650_auditreportdate = new Date(auditReportDate).toISOString();

        const observationType = formData.get('observationType');
        if (!observationType) throw new Error('Observation Type is required');
        data.cr650_observationtype = observationType;

        const observation = formData.get('observation');
        if (!observation) throw new Error('Observation is required');
        data.cr650_observation = observation;

        const details = formData.get('details');
        if (!details) throw new Error('Details is required');
        data.cr650_details = details;

        const riskRating = formData.get('riskRating');
        if (!riskRating) throw new Error('Risk Rating is required');
        data.cr650_riskrating = parseInt(riskRating);

        const managementResponse = formData.get('managementResponse');
        if (!managementResponse) throw new Error('Management Response is required');
        data.cr650_managementresponse = managementResponse;

        const headOfDepartment = formData.get('headOfDepartment');
        if (!headOfDepartment) throw new Error('Head of Department is required');
        data.cr650_headofdepartemt = headOfDepartment;

        const departmentResponsible = formData.get('departmentResponsible');
        if (!departmentResponsible) throw new Error('Department Responsible is required');
        data.cr650_departmentresponsible = departmentResponsible;

        const personResponsible = formData.get('personResponsible');
        if (!personResponsible) throw new Error('Person Responsible is required');
        data.cr650_personresponsible = personResponsible;

        const email = formData.get('email');
        if (!email) throw new Error('Email is required');
        data.cr650_email = email;

        const dueDate = formData.get('dueDate');
        if (!dueDate) throw new Error('Due Date is required');
        data.cr650_duedate = new Date(dueDate).toISOString();

        const aging = formData.get('aging');
        if (!aging) throw new Error('Aging is required');
        data.cr650_aging = parseInt(aging);

        const status = formData.get('status');
        if (!status) throw new Error('Status is required');
        data.cr650_status = parseInt(status);

        // ================= OPTIONAL (ONLY IF FILLED) =================

        const month = formData.get('month');
        if (month) {
            data.cr650_month = formData.get('month');
        }

        const supportPerson = formData.get('supportPerson');
        if (supportPerson) {
            data.cr650_supportperson = supportPerson;
        }

        const daysOverdue = formData.get('daysOverdue');
        if (daysOverdue !== null && daysOverdue !== '') {
            data.cr650_daysoverdue = parseInt(daysOverdue);
        }

        const lastCommunicationDate = formData.get('lastCommunicationDate');
        if (lastCommunicationDate) {
            data.cr650_lastcommunicationdate = new Date(lastCommunicationDate).toISOString();
        }

        const lastPersonCommunicated = formData.get('lastPersonCommunicated');
        if (lastPersonCommunicated) {
            data.cr650_lastpersoncommunicatedwith = lastPersonCommunicated;
        }

        const iaWork = formData.get('iaWork');
        if (iaWork) {
            data.cr650_iawork = iaWork;
        }

        const latestRevisedMap = formData.get('latestRevisedMap');
        if (latestRevisedMap) {
            data.cr650_latestrevisedmap = latestRevisedMap;
        }

        // ================= CONDITIONAL CLOSURE FIELDS =================

        if (data.cr650_status === 3) { // Closed

            const dateClosed = formData.get('dateClosed');
            if (!dateClosed) {
                throw new Error('Date Closed is required when Status is Closed');
            }
            data.cr650_dateclosed = new Date(dateClosed).toISOString();

            const closingRemarks = formData.get('closingRemarks');
            if (!closingRemarks) {
                throw new Error('Closing Remarks is required when Status is Closed');
            }
            data.cr650_closingremarks = closingRemarks;
        }

    } catch (validationError) {
        console.error('❌ Validation Error:', validationError.message);
        alert(validationError.message);
        return;
    }

    // ================= SEND TO DATAVERSE =================

    console.log('📤 Final payload:', JSON.stringify(data, null, 2));

    try {
        let response;

        if (AppState.panelMode === 'create') {
            response = await safeFetch(CONFIG.DATAVERSE_ENDPOINT, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
        } else {
            response = await safeFetch(
                `${CONFIG.DATAVERSE_ENDPOINT}(${AppState.currentObservationId})`,
                {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                }
            );
        }

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText);
        }

        alert(AppState.panelMode === 'create'
            ? 'Observation created successfully!'
            : 'Observation updated successfully!'
        );

        AppState.isDirty = false;
        closePanel();
        await loadObservations();

    } catch (error) {
        console.error('❌ Save failed:', error);
        alert('Failed to save observation. Check console for details.');
    }
}


async function deleteObservation(observationId) {
    if (!confirm('Are you sure you want to delete this observation? This action cannot be undone.')) {
        return;
    }

    try {
        const response = await safeFetch(
            `${CONFIG.DATAVERSE_ENDPOINT}(${observationId})`,
            {
                method: 'DELETE'
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        alert('Observation deleted successfully!');
        closePanel();
        await loadObservations();

    } catch (error) {
        console.error('Error deleting observation:', error);
        alert('Failed to delete observation.');
    }
}

// ================================
// FILTERS & SEARCH
// ================================

function applyFilters() {
    const { search, year, status, risk } = AppState.filters;

    AppState.filteredObservations = AppState.allObservations.filter(obs => {
        // Search filter
        if (search) {
            const searchLower = search.toLowerCase();
            const matchesSearch =
                (obs.cr650_observation || '').toLowerCase().includes(searchLower) ||
                (obs.cr650_auditname || '').toLowerCase().includes(searchLower) ||
                (obs.cr650_departmentresponsible || '').toLowerCase().includes(searchLower) ||
                (obs.cr650_personresponsible || '').toLowerCase().includes(searchLower);

            if (!matchesSearch) return false;
        }

        // Year filter
        if (year && obs.cr650_year?.toString() !== year) {
            return false;
        }

        // Status filter
        if (status && obs.cr650_status?.toString() !== status) {
            return false;
        }

        // Risk filter
        if (risk && obs.cr650_riskrating?.toString() !== risk) {
            return false;
        }

        return true;
    });

    AppState.currentPage = 1; // Reset to first page
    renderTable();
    updateStatistics();
    updatePagination();
}

function clearFilters() {
    AppState.filters = {
        search: '',
        year: '',
        status: '',
        risk: ''
    };

    document.getElementById('searchInput').value = '';
    document.getElementById('yearFilter').value = '';
    document.getElementById('statusFilter').value = '';
    document.getElementById('riskFilter').value = '';

    applyFilters();
}

// ================================
// PAGINATION
// ================================

function updatePagination() {
    const totalRecords = AppState.filteredObservations.length;
    const totalPages = Math.ceil(totalRecords / AppState.pageSize);
    const start = (AppState.currentPage - 1) * AppState.pageSize + 1;
    const end = Math.min(AppState.currentPage * AppState.pageSize, totalRecords);

    const pageInfo = document.getElementById('pageInfo');
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');

    if (pageInfo) {
        pageInfo.textContent = totalRecords > 0
            ? `${start}-${end} of ${totalRecords}`
            : '0-0 of 0';
    }

    if (prevBtn) {
        prevBtn.disabled = AppState.currentPage === 1;
    }

    if (nextBtn) {
        nextBtn.disabled = AppState.currentPage >= totalPages || totalRecords === 0;
    }
}

function previousPage() {
    if (AppState.currentPage > 1) {
        AppState.currentPage--;
        renderTable();
        updatePagination();
    }
}

function nextPage() {
    const totalPages = Math.ceil(AppState.filteredObservations.length / AppState.pageSize);
    if (AppState.currentPage < totalPages) {
        AppState.currentPage++;
        renderTable();
        updatePagination();
    }
}

// ================================
// STATISTICS
// ================================

function updateStatistics() {
    const stats = {
        total: AppState.allObservations.length,
        overdue: AppState.allObservations.filter(obs => obs.cr650_status === 2).length,
        inProgress: AppState.allObservations.filter(obs => obs.cr650_status === 1).length,
        closed: AppState.allObservations.filter(obs => obs.cr650_status === 3).length
    };

    document.getElementById('statTotal').textContent = stats.total;
    document.getElementById('statOverdue').textContent = stats.overdue;
    document.getElementById('statProgress').textContent = stats.inProgress;
    document.getElementById('statClosed').textContent = stats.closed;
}

// ================================
// UI HELPERS
// ================================

function getRiskBadge(riskValue) {
    const risk = CONFIG.RISK_MAP[riskValue];
    if (!risk) return '<span class="badge">-</span>';

    const riskClass = {
        'Critical': 'risk-critical',
        'High': 'risk-high',
        'Moderate': 'risk-moderate',
        'Low': 'risk-low'
    }[risk];

    return `<span class="badge ${riskClass}">${risk}</span>`;
}

function getStatusBadge(statusValue) {
    const status = CONFIG.STATUS_MAP[statusValue];
    if (!status) return '<span class="badge">-</span>';

    const statusClass = {
        'In Progress': 'status-progress',
        'Overdue': 'status-overdue',
        'Closed': 'status-closed'
    }[status];

    return `<span class="badge ${statusClass}">${status}</span>`;
}

function showLoadingState() {
    const tbody = document.getElementById('observationsTableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr class="loading-row">
                <td colspan="8">
                    <div class="loading-state">
                        <div class="spinner"></div>
                        <p>Loading observations...</p>
                    </div>
                </td>
            </tr>
        `;
    }
}

function showErrorState(message) {
    const tbody = document.getElementById('observationsTableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr class="error-row">
                <td colspan="8">
                    <div class="error-state">
                        <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="12" x2="12" y1="8" y2="12"></line>
                            <line x1="12" x2="12.01" y1="16" y2="16"></line>
                        </svg>
                        <h3>Error</h3>
                        <p>${message}</p>
                        <button class="btn btn-primary" data-action="retry-load">Retry</button>
                    </div>
                </td>
            </tr>
        `;
    }
}

// ================================
// EXPORT FUNCTIONALITY
// ================================

async function exportToExcel() {
    const observations = AppState.filteredObservations;
    if (observations.length === 0) {
        alert('No observations to export.');
        return;
    }

    const headers = [
        'Year', 'Quarter', 'Month', 'Company Name', 'Region', 'Audit Name',
        'Observation Type', 'Observation', 'Risk Rating', 'Details',
        'Management Response', 'Head of Department', 'Department Responsible',
        'Person Responsible', 'Email', 'Support Person', 'Audit Report Date',
        'Due Date', 'Days Overdue', 'Aging', 'Date Closed', 'Status',
        'Last Communication Date', 'Last Person Communicated', 'IA Work',
        'Closing Remarks', 'Latest Revised MAP'
    ];

    const rows = observations.map(obs => [
        obs.cr650_year || '',
        obs.cr650_quarter || '',
        obs.cr650_month || '',
        obs.cr650_companyname || '',
        obs.cr650_region || '',
        obs.cr650_auditname || '',
        obs.cr650_observationtype || '',
        obs.cr650_observation || '',
        CONFIG.RISK_MAP[obs.cr650_riskrating] || '',
        obs.cr650_details || '',
        obs.cr650_managementresponse || '',
        obs.cr650_headofdepartemt || '',
        obs.cr650_departmentresponsible || '',
        obs.cr650_personresponsible || '',
        obs.cr650_email || '',
        obs.cr650_supportperson || '',
        formatDate(obs.cr650_auditreportdate),
        formatDate(obs.cr650_duedate),
        obs.cr650_daysoverdue || 0,
        CONFIG.AGING_MAP[obs.cr650_aging] || '',
        formatDate(obs.cr650_dateclosed),
        CONFIG.STATUS_MAP[obs.cr650_status] || '',
        formatDate(obs.cr650_lastcommunicationdate),
        obs.cr650_lastpersoncommunicatedwith || '',
        obs.cr650_iawork || '',
        obs.cr650_closingremarks || '',
        obs.cr650_latestrevisedmap || ''
    ]);

    const csv = [
        headers.join(','),
        ...rows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
    ].join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `observations_complete_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ================================
// UTILITY FUNCTIONS
// ================================

function formatDate(dateString) {
    if (!dateString) return '-';
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}

function formatDateTime(dateString) {
    if (!dateString) return '-';
    const date = new Date(dateString);
    const today = new Date();
    const isToday = date.toDateString() === today.toDateString();

    if (isToday) {
        return `Today, ${date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })}`;
    }

    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    }) + ' – ' + date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
}

function truncate(text, maxLength) {
    if (!text || text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '...';
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function generateMonthOptions(selectedMonth) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];

    return months.map(month =>
        `<option value="${month}" ${selectedMonth === month ? 'selected' : ''}>
      ${month}
     </option>`
    ).join('');
}


function generateRiskSelector(selectedRisk) {
    const risks = [
        { value: 1, label: 'Critical', class: 'risk-critical' },
        { value: 2, label: 'High', class: 'risk-high' },
        { value: 3, label: 'Moderate', class: 'risk-moderate' },
        { value: 4, label: 'Low', class: 'risk-low' }
    ];

    return risks.map(risk => `
        <input type="radio" name="riskRating" id="risk-${risk.value}" class="risk-option" value="${risk.value}" ${selectedRisk === risk.value ? 'checked' : ''} required />
        <label for="risk-${risk.value}" class="risk-label ${risk.class}">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none"
                stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/>
                <path d="M12 9v4"/><path d="M12 17h.01"/>
            </svg>
            ${risk.label}
        </label>
    `).join('');
}

// ================================
// CLIENT RESPONSE ACCEPT/REJECT
// ================================

/**
 * Accept client response and close the observation
 * - Sets status to Closed (3)
 * - Sets date closed to today
 * - Prompts for closing remarks
 */
async function acceptClientResponse(observationId) {
    const closingRemarks = prompt('Enter closing remarks for this observation:');
    if (closingRemarks === null) return; // User cancelled

    if (!closingRemarks.trim()) {
        alert('Closing remarks are required when accepting a client response.');
        return;
    }

    try {
        const today = new Date().toISOString();

        const updateData = {
            'cr650_status': 3, // Closed
            'cr650_dateclosed': today,
            'cr650_closingremarks': closingRemarks,
            'cr650_aging': 1 // Not due (closed observations)
        };

        // If there's a revised due date from the client, apply it
        if (AppState.latestClientUpdate?.cr650_revisedduedate) {
            updateData.cr650_duedate = AppState.latestClientUpdate.cr650_revisedduedate;
        }

        // If there's revised management feedback, update the Latest Revised MAP
        if (AppState.latestClientUpdate?.cr650_revisedmanagementfeedback) {
            updateData.cr650_latestrevisedmap = AppState.latestClientUpdate.cr650_revisedmanagementfeedback;
        }

        const response = await safeFetch(
            `${CONFIG.DATAVERSE_ENDPOINT}(${observationId})`,
            {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(updateData)
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        alert('Client response accepted. Observation has been closed.');
        closePanel();
        await loadObservations();

    } catch (error) {
        console.error('Error accepting client response:', error);
        alert('Failed to accept client response. Please try again.');
    }
}

/**
 * Reject client response and send back to client
 * - Keeps/sets status to In Progress (1)
 * - Client must revise and resubmit
 */
async function rejectClientResponse(observationId) {
    const reason = prompt('Enter reason for rejection (this will be added to IA Work notes):');
    if (reason === null) return; // User cancelled

    if (!reason.trim()) {
        alert('Please provide a reason for rejection.');
        return;
    }

    try {
        const today = new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });

        // Build the rejection note
        const rejectionNote = `[${today}] Client response rejected: ${reason}`;

        // Append to existing IA Work notes
        let iaWork = AppState.currentObservation?.cr650_iawork || '';
        if (iaWork) {
            iaWork = iaWork + '\n\n' + rejectionNote;
        } else {
            iaWork = rejectionNote;
        }

        const updateData = {
            'cr650_status': 1, // In Progress
            'cr650_iawork': iaWork,
            'cr650_lastcommunicationdate': new Date().toISOString()
        };

        const response = await safeFetch(
            `${CONFIG.DATAVERSE_ENDPOINT}(${observationId})`,
            {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(updateData)
            }
        );

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        alert('Client response rejected. The observation remains In Progress and the client will need to revise their submission.');
        closePanel();
        await loadObservations();

    } catch (error) {
        console.error('Error rejecting client response:', error);
        alert('Failed to reject client response. Please try again.');
    }
}

// ================================
// EVENT LISTENERS
// ================================

function initializeEventListeners() {
    // Search input (with debounce)
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', debounce((e) => {
            AppState.filters.search = e.target.value;
            applyFilters();
        }, 300));
    }

    // Filter dropdowns
    const yearFilter = document.getElementById('yearFilter');
    if (yearFilter) {
        yearFilter.addEventListener('change', (e) => {
            AppState.filters.year = e.target.value;
            applyFilters();
        });
    }

    const statusFilter = document.getElementById('statusFilter');
    if (statusFilter) {
        statusFilter.addEventListener('change', (e) => {
            AppState.filters.status = e.target.value;
            applyFilters();
        });
    }

    const riskFilter = document.getElementById('riskFilter');
    if (riskFilter) {
        riskFilter.addEventListener('change', (e) => {
            AppState.filters.risk = e.target.value;
            applyFilters();
        });
    }

    // Clear filters button
    const clearBtn = document.getElementById('clearFiltersBtn');
    if (clearBtn) {
        clearBtn.addEventListener('click', clearFilters);
    }

    // Pagination
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    if (prevBtn) prevBtn.addEventListener('click', previousPage);
    if (nextBtn) nextBtn.addEventListener('click', nextPage);

    // Dashboard button
    const dashboardBtn = document.getElementById('dashboardBtn');
    if (dashboardBtn) {
        dashboardBtn.addEventListener('click', () => {
            window.location.href = CONFIG.DASHBOARD_URL;
        });
    }

    // Export button
    const exportBtn = document.getElementById('exportBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportToExcel);
    }

    // Add observation button
    const addBtn = document.getElementById('addObservationBtn');
    if (addBtn) {
        addBtn.addEventListener('click', () => openPanel('create'));
    }

    // Close panel on overlay click
    const overlay = document.getElementById('panelOverlay');
    if (overlay) {
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) closePanel();
        });
    }

    // ESC key to close panel
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            const overlay = document.getElementById('panelOverlay');
            if (overlay && overlay.classList.contains('active')) {
                closePanel();
            }
        }
    });

    document.addEventListener('click', (e) => {
        const target = e.target.closest('button[data-action]');
        if (!target) return;

        const action = target.dataset.action;

        if (action === 'create-first') {
            openPanel('create');
        } else if (action === 'retry-load') {
            loadObservations();
        }
    });

    attachTableEventListeners();
}

// ================================
// INITIALIZATION
// ================================

async function init() {
    console.log('Initializing Petrolube Observation Tracker (Complete Version)...');
    initializeEventListeners();
    await loadObservations();
    console.log('Observation Tracker initialized successfully');
}

// Start app when DOM is ready
document.addEventListener('DOMContentLoaded', init);

// Global functions (for inline onclick handlers)
window.openPanel = openPanel;
window.closePanel = closePanel;
window.switchTab = switchTab;
window.saveObservation = saveObservation;
window.deleteObservation = deleteObservation;
window.calculateDaysOverdue = calculateDaysOverdue;
window.toggleClosureFields = toggleClosureFields;
window.acceptClientResponse = acceptClientResponse;
window.rejectClientResponse = rejectClientResponse;