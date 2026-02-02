/**
 * Client Portal - Internal Audit
 * BULLETPROOF VERSION with enhanced error handling
 */

(function() {
    'use strict';

    // =============================================
    // CONFIGURATION
    // =============================================
    const CONFIG = {
        OBSERVATIONS_API: '/_api/cr650_ia_observations',
        OBSERVATION_PAGE_URL: '/Observation-Client/',
        RISK_RATINGS: {
            1: { label: 'Critical', class: 'risk-critical' },
            2: { label: 'High', class: 'risk-high' },
            3: { label: 'Moderate', class: 'risk-moderate' },
            4: { label: 'Low', class: 'risk-low' }
        },
        STATUS_CODES: {
            OPEN: 1,
            IN_PROGRESS: 2,
            COMPLETED: 3,
            CLOSED: 4
        },
        DATE_FORMAT_OPTIONS: {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        }
    };

    // =============================================
    // STATE MANAGEMENT
    // =============================================
    let state = {
        observations: [],
        filteredObservations: [],
        currentView: 'cards',
        urlParams: {},
        filters: {
            dueDate: '',
            status: '',
            audit: '',
            search: ''
        },
        userEmail: '',
        userName: ''
    };

    // =============================================
    // DOM ELEMENTS
    // =============================================
    const elements = {};

    function cacheElements() {
        console.log('🔍 Caching elements...');
        
        // Required elements
        const required = [
            'loadingOverlay', 'mainContent', 'greetingName', 'welcomeBanner', 
            'welcomeText', 'closeBanner', 'filterNotice', 'filterNoticeText', 
            'clearFilters', 'totalCount', 'overdueCount', 'pendingCount', 
            'completedCount', 'filterDueDate', 'filterStatus', 'filterAudit', 
            'filterSearch', 'resetFilters', 'viewCards', 'viewTable', 
            'observationsCards', 'observationsTable', 'observationsTableBody', 
            'resultsCount', 'emptyState', 'emptyResetFilters', 'listLoading', 
            'detailModal', 'detailModalBody', 'closeDetailModal', 'closeDetailBtn', 
            'updateObservationBtn', 'errorModal', 'errorMessage', 'retryButton'
        ];

        const missing = [];
        
        for (const id of required) {
            elements[id] = document.getElementById(id);
            if (!elements[id]) {
                missing.push(id);
                console.warn(`⚠️ Element not found: #${id}`);
            } else {
                console.log(`✓ Found: #${id}`);
            }
        }

        // Optional element (may not exist if header removed)
        elements.userName = document.getElementById('userName');
        if (!elements.userName) {
            console.log('ℹ️ Optional element userName not found (expected if header removed)');
        }

        if (missing.length > 0) {
            console.error('❌ Missing required elements:', missing);
            throw new Error(`Missing required DOM elements: ${missing.join(', ')}`);
        }

        console.log('✅ All elements cached successfully');
    }

    // =============================================
    // SAFE ELEMENT OPERATIONS
    // =============================================
    const safe = {
        setText(element, text) {
            if (element && typeof element.textContent !== 'undefined') {
                try {
                    element.textContent = text;
                    return true;
                } catch (e) {
                    console.error('Error setting text:', e);
                    return false;
                }
            }
            return false;
        },
        
        setHTML(element, html) {
            if (element && typeof element.innerHTML !== 'undefined') {
                try {
                    element.innerHTML = html;
                    return true;
                } catch (e) {
                    console.error('Error setting HTML:', e);
                    return false;
                }
            }
            return false;
        },
        
        setDisplay(element, value) {
            if (element && element.style) {
                try {
                    element.style.display = value;
                    return true;
                } catch (e) {
                    console.error('Error setting display:', e);
                    return false;
                }
            }
            return false;
        },
        
        addClass(element, className) {
            if (element && element.classList) {
                try {
                    element.classList.add(className);
                    return true;
                } catch (e) {
                    console.error('Error adding class:', e);
                    return false;
                }
            }
            return false;
        },
        
        removeClass(element, className) {
            if (element && element.classList) {
                try {
                    element.classList.remove(className);
                    return true;
                } catch (e) {
                    console.error('Error removing class:', e);
                    return false;
                }
            }
            return false;
        }
    };

    // =============================================
    // URL PARAMETER HANDLING
    // =============================================
    function parseUrlParams() {
        const params = new URLSearchParams(window.location.search);
        state.urlParams = {
            email: params.get('email'),
            dueDate: params.get('dueDate'),
            status: params.get('status')
        };
    }

    function showUrlFilterNotice() {
        if (Object.values(state.urlParams).some(v => v)) {
            const notices = [];
            if (state.urlParams.email) notices.push(`Email: ${state.urlParams.email}`);
            if (state.urlParams.dueDate) notices.push(`Due Date: ${formatDate(new Date(state.urlParams.dueDate))}`);
            if (state.urlParams.status) notices.push(`Status: ${capitalizeFirst(state.urlParams.status)}`);
            
            safe.setText(elements.filterNoticeText, 'Filtered by: ' + notices.join(' | '));
            safe.setDisplay(elements.filterNotice, 'flex');
        }
    }

    // =============================================
    // DATA FETCHING
    // =============================================
    async function fetchObservations() {
        let url = CONFIG.OBSERVATIONS_API + '?$orderby=cr650_duedate asc';
        
        const filters = [];
        if (state.urlParams.email) {
            filters.push(`cr650_email eq '${encodeURIComponent(state.urlParams.email)}'`);
        }
        if (state.urlParams.dueDate) {
            filters.push(`cr650_duedate eq ${state.urlParams.dueDate}`);
        }
        if (state.urlParams.status) {
            const statusCode = getStatusCodeFromString(state.urlParams.status);
            if (statusCode) {
                filters.push(`cr650_status eq ${statusCode}`);
            }
        }
        
        if (filters.length > 0) {
            url += '&$filter=' + filters.join(' and ');
        }

        const response = await fetch(url);
        if (!response.ok) {
            throw new Error('Failed to fetch observations');
        }
        
        const data = await response.json();
        return data.value || [];
    }

    function getStatusCodeFromString(statusStr) {
        const statusMap = {
            'open': CONFIG.STATUS_CODES.OPEN,
            'in_progress': CONFIG.STATUS_CODES.IN_PROGRESS,
            'completed': CONFIG.STATUS_CODES.COMPLETED,
            'closed': CONFIG.STATUS_CODES.CLOSED,
            'overdue': CONFIG.STATUS_CODES.OPEN,
            'pending': CONFIG.STATUS_CODES.IN_PROGRESS
        };
        return statusMap[statusStr.toLowerCase()];
    }

    function transformObservations(rawObservations) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return rawObservations.map(obs => {
            let dueDate = null;
            if (obs.cr650_duedate) {
                dueDate = new Date(obs.cr650_duedate);
            }

            let status = 'pending';
            if (obs.cr650_status === CONFIG.STATUS_CODES.COMPLETED || 
                obs.cr650_status === CONFIG.STATUS_CODES.CLOSED) {
                status = 'completed';
            } else if (dueDate && dueDate < today) {
                status = 'overdue';
            }

            return {
                id: obs.cr650_ia_observationid,
                reference: obs.cr650_name || 'N/A',
                observation: obs.cr650_observation || '',
                details: obs.cr650_details || '',
                auditName: obs.cr650_auditname || 'N/A',
                auditDate: obs.cr650_auditreportdate ? new Date(obs.cr650_auditreportdate) : null,
                responsiblePerson: obs.cr650_personresponsible || 'N/A',
                email: obs.cr650_email || '',
                dueDate: dueDate,
                dueDateISO: obs.cr650_duedate ? obs.cr650_duedate.split('T')[0] : '',
                daysOverdue: obs.cr650_daysoverdue || 0,
                riskRating: obs.cr650_riskrating || 4,
                managementResponse: obs.cr650_managementresponse || '',
                statusCode: obs.cr650_status,
                status: status
            };
        });
    }

    function extractUserInfo() {
        if (state.observations.length > 0) {
            const first = state.observations[0];
            state.userName = first.responsiblePerson;
            state.userEmail = first.email;
        }
    }

    // =============================================
    // FILTER MANAGEMENT
    // =============================================
    function populateFilterOptions() {
        const dueDates = [...new Set(state.observations
            .filter(o => o.dueDateISO)
            .map(o => o.dueDateISO))]
            .sort();

        safe.setHTML(elements.filterDueDate, '<option value="">All Due Dates</option>');
        dueDates.forEach(date => {
            const option = document.createElement('option');
            option.value = date;
            option.textContent = formatDate(new Date(date));
            elements.filterDueDate.appendChild(option);
        });

        const audits = [...new Set(state.observations
            .filter(o => o.auditName && o.auditName !== 'N/A')
            .map(o => o.auditName))]
            .sort();

        safe.setHTML(elements.filterAudit, '<option value="">All Audits</option>');
        audits.forEach(audit => {
            const option = document.createElement('option');
            option.value = audit;
            option.textContent = audit;
            elements.filterAudit.appendChild(option);
        });

        if (state.urlParams.dueDate && elements.filterDueDate) {
            elements.filterDueDate.value = state.urlParams.dueDate;
            state.filters.dueDate = state.urlParams.dueDate;
        }

        if (state.urlParams.status && elements.filterStatus) {
            elements.filterStatus.value = state.urlParams.status;
            state.filters.status = state.urlParams.status;
        }
    }

    function applyFilters() {
        let filtered = [...state.observations];

        if (state.filters.dueDate) {
            filtered = filtered.filter(o => o.dueDateISO === state.filters.dueDate);
        }

        if (state.filters.status) {
            filtered = filtered.filter(o => o.status === state.filters.status);
        }

        if (state.filters.audit) {
            filtered = filtered.filter(o => o.auditName === state.filters.audit);
        }

        if (state.filters.search) {
            const searchLower = state.filters.search.toLowerCase();
            filtered = filtered.filter(o => 
                o.reference.toLowerCase().includes(searchLower) ||
                o.observation.toLowerCase().includes(searchLower) ||
                o.auditName.toLowerCase().includes(searchLower) ||
                o.responsiblePerson.toLowerCase().includes(searchLower)
            );
        }

        state.filteredObservations = filtered;
        return filtered;
    }

    // =============================================
    // STATISTICS
    // =============================================
    function updateStats() {
        const total = state.observations.length;
        const overdue = state.observations.filter(o => o.status === 'overdue').length;
        const pending = state.observations.filter(o => o.status === 'pending').length;
        const completed = state.observations.filter(o => o.status === 'completed').length;

        animateNumber(elements.totalCount, total);
        animateNumber(elements.overdueCount, overdue);
        animateNumber(elements.pendingCount, pending);
        animateNumber(elements.completedCount, completed);
    }

    function animateNumber(element, target) {
        if (!element) {
            console.warn('animateNumber: element is null');
            return;
        }
        
        const duration = 500;
        const start = parseInt(element.textContent) || 0;
        const increment = (target - start) / (duration / 16);
        let current = start;

        const animate = () => {
            current += increment;
            if ((increment > 0 && current >= target) || (increment < 0 && current <= target)) {
                safe.setText(element, target);
            } else {
                safe.setText(element, Math.round(current));
                requestAnimationFrame(animate);
            }
        };

        animate();
    }

    // =============================================
    // RENDERING
    // =============================================
    function renderObservations() {
        const observations = applyFilters();
        const count = observations.length;
        
        safe.setText(elements.resultsCount, count + ' observation' + (count !== 1 ? 's' : ''));

        if (count === 0) {
            safe.setDisplay(elements.emptyState, 'flex');
            safe.setDisplay(elements.observationsCards, 'none');
            safe.setDisplay(elements.observationsTable, 'none');
            return;
        }

        safe.setDisplay(elements.emptyState, 'none');

        if (state.currentView === 'cards') {
            renderCardView(observations);
            safe.setDisplay(elements.observationsCards, 'grid');
            safe.setDisplay(elements.observationsTable, 'none');
        } else {
            renderTableView(observations);
            safe.setDisplay(elements.observationsCards, 'none');
            safe.setDisplay(elements.observationsTable, 'block');
        }
    }

    function renderCardView(observations) {
        const html = observations.map(obs => `
            <div class="observation-card ${obs.status}" data-id="${obs.id}">
                <div class="card-header-row">
                    <span class="obs-reference">${escapeHtml(obs.reference)}</span>
                    <span class="risk-badge ${CONFIG.RISK_RATINGS[obs.riskRating]?.class || 'risk-low'}">
                        ${CONFIG.RISK_RATINGS[obs.riskRating]?.label || 'Low'}
                    </span>
                </div>
                <div class="obs-content">
                    <p class="obs-text">${escapeHtml(truncateText(obs.observation, 120))}</p>
                </div>
                <div class="obs-meta">
                    <div class="meta-item">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                        </svg>
                        <span>${escapeHtml(obs.auditName)}</span>
                    </div>
                    <div class="meta-item">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                            <line x1="16" y1="2" x2="16" y2="6"/>
                            <line x1="8" y1="2" x2="8" y2="6"/>
                            <line x1="3" y1="10" x2="21" y2="10"/>
                        </svg>
                        <span class="${obs.status === 'overdue' ? 'text-danger' : ''}">
                            ${obs.dueDate ? formatDate(obs.dueDate) : 'No due date'}
                        </span>
                    </div>
                </div>
                <div class="obs-footer">
                    <span class="status-badge status-${obs.status}">
                        ${getStatusIcon(obs.status)}
                        ${capitalizeFirst(obs.status)}
                    </span>
                    <div class="card-actions">
                        <button class="btn-icon view-detail-btn" data-id="${obs.id}" title="View Details">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                                <circle cx="12" cy="12" r="3"/>
                            </svg>
                        </button>
                        <a href="${CONFIG.OBSERVATION_PAGE_URL}?id=${obs.id}" class="btn-icon btn-primary-icon" title="Update">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                            </svg>
                        </a>
                    </div>
                </div>
            </div>
        `).join('');
        
        safe.setHTML(elements.observationsCards, html);
    }

    function renderTableView(observations) {
        const html = observations.map(obs => `
            <tr class="${obs.status}" data-id="${obs.id}">
                <td><span class="table-reference">${escapeHtml(obs.reference)}</span></td>
                <td><span class="table-observation">${escapeHtml(truncateText(obs.observation, 80))}</span></td>
                <td>${escapeHtml(obs.auditName)}</td>
                <td class="${obs.status === 'overdue' ? 'text-danger' : ''}">${obs.dueDate ? formatDate(obs.dueDate) : 'N/A'}</td>
                <td><span class="risk-badge-small ${CONFIG.RISK_RATINGS[obs.riskRating]?.class || 'risk-low'}">${CONFIG.RISK_RATINGS[obs.riskRating]?.label || 'Low'}</span></td>
                <td><span class="status-badge-small status-${obs.status}">${capitalizeFirst(obs.status)}</span></td>
                <td>
                    <div class="table-actions">
                        <button class="btn-icon-small view-detail-btn" data-id="${obs.id}" title="View">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                                <circle cx="12" cy="12" r="3"/>
                            </svg>
                        </button>
                        <a href="${CONFIG.OBSERVATION_PAGE_URL}?id=${obs.id}" class="btn-icon-small btn-primary-icon" title="Update">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                            </svg>
                        </a>
                    </div>
                </td>
            </tr>
        `).join('');
        
        safe.setHTML(elements.observationsTableBody, html);
    }

    function getStatusIcon(status) {
        switch (status) {
            case 'overdue':
                return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
            case 'pending':
                return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>';
            case 'completed':
                return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>';
            default:
                return '';
        }
    }

    // =============================================
    // DETAIL MODAL
    // =============================================
    window.viewObservation = function(id) {
        console.log('👁️ Viewing observation:', id);
        const obs = state.observations.find(o => o.id === id);
        if (!obs) {
            console.error('Observation not found:', id);
            return;
        }

        const detailHTML = `
            <div class="detail-grid">
                <div class="detail-section">
                    <h3 class="detail-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <circle cx="12" cy="12" r="10"/>
                            <line x1="12" y1="16" x2="12" y2="12"/>
                            <line x1="12" y1="8" x2="12.01" y2="8"/>
                        </svg>
                        Observation Information
                    </h3>
                    <div class="detail-row">
                        <span class="detail-label">Reference</span>
                        <span class="detail-value">${escapeHtml(obs.reference)}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Risk Rating</span>
                        <span class="risk-badge ${CONFIG.RISK_RATINGS[obs.riskRating]?.class || 'risk-low'}">
                            ${CONFIG.RISK_RATINGS[obs.riskRating]?.label || 'Low'}
                        </span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Status</span>
                        <span class="status-badge status-${obs.status}">
                            ${getStatusIcon(obs.status)}
                            ${capitalizeFirst(obs.status)}
                        </span>
                    </div>
                </div>
                <div class="detail-section">
                    <h3 class="detail-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                        </svg>
                        Audit Details
                    </h3>
                    <div class="detail-row">
                        <span class="detail-label">Audit Name</span>
                        <span class="detail-value">${escapeHtml(obs.auditName)}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Audit Date</span>
                        <span class="detail-value">${obs.auditDate ? formatDate(obs.auditDate) : 'N/A'}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Due Date</span>
                        <span class="detail-value ${obs.status === 'overdue' ? 'text-danger' : ''}">${obs.dueDate ? formatDate(obs.dueDate) : 'N/A'}</span>
                    </div>
                </div>
                <div class="detail-section full-width">
                    <h3 class="detail-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="16" y1="13" x2="8" y2="13"/>
                            <line x1="16" y1="17" x2="8" y2="17"/>
                        </svg>
                        Observation
                    </h3>
                    <div class="detail-box">${escapeHtml(obs.observation)}</div>
                </div>
                <div class="detail-section full-width">
                    <h3 class="detail-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                        </svg>
                        Details
                    </h3>
                    <div class="detail-box">${escapeHtml(obs.details || 'No additional details provided.')}</div>
                </div>
                <div class="detail-section full-width">
                    <h3 class="detail-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
                        </svg>
                        Management Response
                    </h3>
                    <div class="detail-box">${escapeHtml(obs.managementResponse || 'No management response yet.')}</div>
                </div>
            </div>
        `;

        safe.setHTML(elements.detailModalBody, detailHTML);

        if (elements.updateObservationBtn) {
            elements.updateObservationBtn.href = CONFIG.OBSERVATION_PAGE_URL + '?id=' + obs.id;
        }

        safe.setDisplay(elements.detailModal, 'flex');
    };

    function closeDetailModal() {
        safe.setDisplay(elements.detailModal, 'none');
    }

    // =============================================
    // UI STATE
    // =============================================
    function updateUI() {
        console.log('🎨 Updating UI...');
        
        try {
            // Update user info (optional - may not exist)
            if (elements.userName) {
                safe.setText(elements.userName, state.userName || 'User');
                console.log('✓ Updated userName');
            }
            
            if (elements.greetingName) {
                const firstName = getFirstName(state.userName);
                safe.setText(elements.greetingName, firstName);
                console.log('✓ Updated greetingName to:', firstName);
            } else {
                console.warn('⚠️ greetingName element not found');
            }

            // Update welcome text based on observations
            if (elements.welcomeText) {
                const overdue = state.observations.filter(o => o.status === 'overdue').length;
                if (overdue > 0) {
                    safe.setText(elements.welcomeText, `You have ${overdue} overdue observation${overdue !== 1 ? 's' : ''} requiring immediate attention. Please review and update them as soon as possible.`);
                } else {
                    safe.setText(elements.welcomeText, 'You have observations assigned to you. Use the filters below to navigate through your assignments and submit updates.');
                }
                console.log('✓ Updated welcomeText');
            } else {
                console.warn('⚠️ welcomeText element not found');
            }
            
            console.log('✅ UI update complete');
        } catch (error) {
            console.error('❌ Error in updateUI:', error);
        }
    }

    function showLoading(show) {
        if (show) {
            safe.setDisplay(elements.loadingOverlay, 'flex');
            safe.removeClass(elements.mainContent, 'show');
        } else {
            safe.setDisplay(elements.loadingOverlay, 'none');
            safe.addClass(elements.mainContent, 'show');
        }
    }

    function showError(message) {
        safe.setText(elements.errorMessage, message);
        safe.setDisplay(elements.errorModal, 'flex');
    }

    function hideError() {
        safe.setDisplay(elements.errorModal, 'none');
    }

    // =============================================
    // EVENT HANDLERS
    // =============================================
    function setupEventListeners() {
        console.log('🎯 Setting up event listeners...');

        if (elements.closeBanner) {
            elements.closeBanner.addEventListener('click', () => {
                safe.setDisplay(elements.welcomeBanner, 'none');
            });
        }

        if (elements.clearFilters) {
            elements.clearFilters.addEventListener('click', () => {
                window.history.replaceState({}, '', window.location.pathname);
                state.urlParams = {};
                resetAllFilters();
            });
        }

        if (elements.filterDueDate) {
            elements.filterDueDate.addEventListener('change', (e) => {
                state.filters.dueDate = e.target.value;
                renderObservations();
            });
        }

        if (elements.filterStatus) {
            elements.filterStatus.addEventListener('change', (e) => {
                state.filters.status = e.target.value;
                renderObservations();
            });
        }

        if (elements.filterAudit) {
            elements.filterAudit.addEventListener('change', (e) => {
                state.filters.audit = e.target.value;
                renderObservations();
            });
        }

        if (elements.filterSearch) {
            elements.filterSearch.addEventListener('input', debounce((e) => {
                state.filters.search = e.target.value;
                renderObservations();
            }, 300));
        }

        if (elements.resetFilters) {
            elements.resetFilters.addEventListener('click', resetAllFilters);
        }
        if (elements.emptyResetFilters) {
            elements.emptyResetFilters.addEventListener('click', resetAllFilters);
        }

        if (elements.viewCards) {
            elements.viewCards.addEventListener('click', () => {
                state.currentView = 'cards';
                safe.addClass(elements.viewCards, 'active');
                if (elements.viewTable) safe.removeClass(elements.viewTable, 'active');
                renderObservations();
            });
        }

        if (elements.viewTable) {
            elements.viewTable.addEventListener('click', () => {
                state.currentView = 'table';
                safe.addClass(elements.viewTable, 'active');
                if (elements.viewCards) safe.removeClass(elements.viewCards, 'active');
                renderObservations();
            });
        }

        // Event delegation for view buttons (CSP compliant)
        document.addEventListener('click', (e) => {
            const viewBtn = e.target.closest('.view-detail-btn');
            if (viewBtn) {
                e.preventDefault();
                const obsId = viewBtn.getAttribute('data-id');
                if (obsId) {
                    viewObservation(obsId);
                }
            }
        });

        if (elements.closeDetailModal) {
            elements.closeDetailModal.addEventListener('click', closeDetailModal);
        }
        if (elements.closeDetailBtn) {
            elements.closeDetailBtn.addEventListener('click', closeDetailModal);
        }
        if (elements.detailModal) {
            elements.detailModal.addEventListener('click', (e) => {
                if (e.target === elements.detailModal) {
                    closeDetailModal();
                }
            });
        }

        if (elements.retryButton) {
            elements.retryButton.addEventListener('click', () => {
                hideError();
                init();
            });
        }

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                closeDetailModal();
                hideError();
            }
        });

        console.log('✅ Event listeners set up');
    }

    function resetAllFilters() {
        state.filters = {
            dueDate: '',
            status: '',
            audit: '',
            search: ''
        };

        if (elements.filterDueDate) elements.filterDueDate.value = '';
        if (elements.filterStatus) elements.filterStatus.value = '';
        if (elements.filterAudit) elements.filterAudit.value = '';
        if (elements.filterSearch) elements.filterSearch.value = '';
        safe.setDisplay(elements.filterNotice, 'none');

        renderObservations();
    }

    // =============================================
    // UTILITY FUNCTIONS
    // =============================================
    function formatDate(date) {
        if (!date || !(date instanceof Date) || isNaN(date)) return 'N/A';
        return date.toLocaleDateString('en-US', CONFIG.DATE_FORMAT_OPTIONS);
    }

    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function truncateText(text, maxLength) {
        if (!text) return '';
        if (text.length <= maxLength) return text;
        return text.substring(0, maxLength) + '...';
    }

    function capitalizeFirst(str) {
        if (!str) return '';
        return str.charAt(0).toUpperCase() + str.slice(1);
    }

    function getFirstName(fullName) {
        if (!fullName) return 'User';
        return fullName.split(' ')[0];
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

    // =============================================
    // INITIALIZATION
    // =============================================
    async function init() {
        console.log('🚀 Initializing portal...');
        
        try {
            cacheElements();
            parseUrlParams();
            // showUrlFilterNotice(); // Hidden per user request
            setupEventListeners();
            showLoading(true);

            const rawObservations = await fetchObservations();
            state.observations = transformObservations(rawObservations);
            extractUserInfo();

            updateUI();
            updateStats();
            populateFilterOptions();
            renderObservations();

            showLoading(false);
            console.log('✅ Portal initialized successfully');
        } catch (error) {
            console.error('❌ Error initializing portal:', error);
            showLoading(false);
            showError('An error occurred while loading your observations. Please refresh the page or contact support.');
        }
    }

    // Wait for DOM to be fully loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        // DOM is already loaded, wait a tiny bit for Power Pages to finish
        setTimeout(init, 100);
    }

})();