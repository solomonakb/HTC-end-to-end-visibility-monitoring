// app.js

document.addEventListener('DOMContentLoaded', init);

const API_BASE = '/api';

// State
let currentTab = 'tab-all-events';
let currentPage = 1;
let currentFilters = {
    fleet: '',
    event_type: '',
    date_from: '',
    date_to: '',
    search: '',
    bom: '',
    part_group: ''
};
let cachedResolutions = [];

function init() {
    setupEventListeners();
    loadDashboard();
    loadFleetTypes();
    loadEvents(1);
    
    // Set default dates if needed, or leave blank
    const savedTheme = localStorage.getItem('htc-theme') || 'dark';
    document.body.setAttribute('data-theme', savedTheme);
    const icon = document.getElementById('theme-icon');
    if(icon) icon.textContent = savedTheme === 'dark' ? '🌙' : '☀️';
}

function setupEventListeners() {
    // Tabs
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const targetId = e.target.getAttribute('data-target');
            switchTab(targetId);
        });
    });

    // KPI Cards
    const cardTotal = document.getElementById('card-total');
    if (cardTotal) cardTotal.addEventListener('click', () => { switchTab('tab-all-events'); });
    
    const cardInstall = document.getElementById('card-install');
    if (cardInstall) cardInstall.addEventListener('click', () => { 
        document.getElementById('filter-event-type').value = 'INSTALL';
        document.getElementById('btn-apply-filters').click();
        switchTab('tab-all-events'); 
    });
    
    const cardRemove = document.getElementById('card-remove');
    if (cardRemove) cardRemove.addEventListener('click', () => { 
        document.getElementById('filter-event-type').value = 'REMOVE';
        document.getElementById('btn-apply-filters').click();
        switchTab('tab-all-events'); 
    });
    
    const cardXxx = document.getElementById('card-xxx');
    if (cardXxx) cardXxx.addEventListener('click', () => { switchTab('tab-alert-b'); });
    
    const cardEmpty = document.getElementById('card-empty');
    if (cardEmpty) cardEmpty.addEventListener('click', () => { switchTab('tab-alert-c'); });

    // Filters
    document.getElementById('btn-apply-filters').addEventListener('click', () => {
        updateFiltersFromUI();
        loadDashboard();
        if(currentTab === 'tab-all-events') loadEvents(1);
        else if(currentTab === 'tab-alert-a') loadAlertA();
        else if(currentTab === 'tab-alert-b') loadAlertB();
        else if(currentTab === 'tab-alert-c') loadAlertC();
        else if(currentTab === 'tab-resolutions') loadResolutions();
    });

    // Enter key to apply filters
    const filterInputs = document.querySelectorAll('.filters-bar input, .filters-bar select');
    filterInputs.forEach(input => {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault(); // prevent any form submission
                document.getElementById('btn-apply-filters').click();
            }
        });
    });

    // Clear filters
    const clearBtn = document.getElementById('btn-clear-filters');
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            document.getElementById('fleet-filter').value = '';
            document.getElementById('event-type-filter').value = '';
            document.getElementById('date-from').value = '';
            document.getElementById('date-to').value = '';
            document.getElementById('search-input').value = '';
            document.getElementById('bom-filter').value = '';
            document.getElementById('part-group-filter').value = '';
            document.getElementById('btn-apply-filters').click();
        });
    }

    // Export
    document.getElementById('btn-export-csv').addEventListener('click', exportCSV);

    // Fetch from network share
    const fetchBtn = document.getElementById('btn-fetch-reports');
    if (fetchBtn) {
        fetchBtn.addEventListener('click', fetchReports);
    }

    // Resolution Modal
    document.getElementById('btn-new-resolution').addEventListener('click', () => {
        document.getElementById('resolution-form').reset();
        document.getElementById('res-id').value = '';
        document.getElementById('resolution-modal').classList.remove('hidden');
    });

    document.getElementById('btn-cancel-resolution').addEventListener('click', () => {
        document.getElementById('resolution-modal').classList.add('hidden');
    });

    document.getElementById('resolution-form').addEventListener('submit', (e) => {
        e.preventDefault();
        submitResolution();
    });

    const themeBtn = document.getElementById('btn-theme-toggle');
    if(themeBtn) themeBtn.addEventListener('click', toggleTheme);
}

function toggleTheme() {
    const body = document.body;
    const current = body.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    body.setAttribute('data-theme', next);
    localStorage.setItem('htc-theme', next);
    const icon = document.getElementById('theme-icon');
    if(icon) icon.textContent = next === 'dark' ? '🌙' : '☀️';
}

function updateFiltersFromUI() {
    currentFilters.fleet = document.getElementById('fleet-filter').value;
    currentFilters.event_type = document.getElementById('event-type-filter').value;
    currentFilters.date_from = document.getElementById('date-from').value;
    currentFilters.date_to = document.getElementById('date-to').value;
    currentFilters.search = document.getElementById('search-input').value;
    currentFilters.bom = document.getElementById('bom-filter').value;
    currentFilters.part_group = document.getElementById('part-group-filter').value;
}

function buildQueryString(params = {}) {
    const query = new URLSearchParams({
        ...currentFilters,
        ...params
    });
    // Remove empty params
    for (const [key, value] of Array.from(query.entries())) {
        if (!value) query.delete(key);
    }
    return query.toString();
}

function switchTab(tabId) {
    // Update active class on buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.getAttribute('data-target') === tabId);
    });

    // Update active class on content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('hidden', content.id !== tabId);
    });

    currentTab = tabId;

    // Load data specific to the tab
    if (tabId === 'tab-all-events') loadEvents(1);
    else if (tabId === 'tab-alert-a') loadAlertA();
    else if (tabId === 'tab-alert-b') loadAlertB();
    else if (tabId === 'tab-alert-c') loadAlertC();
    else if (tabId === 'tab-resolutions') loadResolutions();
}

async function fetchAPI(endpoint, options = {}) {
    try {
        const res = await fetch(API_BASE + endpoint, options);
        if(!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        const contentType = res.headers.get('content-type');
        if(contentType && contentType.includes('application/json')) {
            return await res.json();
        }
        return await res.text();
    } catch (e) {
        console.error("API Fetch Error:", e);
        showToast("Error communicating with server", "error");
        return null;
    }
}

async function loadDashboard() {
    const qs = buildQueryString();
    const data = await fetchAPI(`/dashboard?${qs}`);
    if(data) {
        animateCounter(document.getElementById('kpi-total'), data.total_events || 0);
        animateCounter(document.getElementById('kpi-install'), data.install_count || 0);
        animateCounter(document.getElementById('kpi-remove'), data.remove_count || 0);
        animateCounter(document.getElementById('kpi-xxx'), data.xxx_sn_count || 0);
        animateCounter(document.getElementById('kpi-empty'), data.empty_slots_count || 0);
    }
}

async function loadFleetTypes() {
    const data = await fetchAPI('/fleet-types');
    if(data && data.fleets) {
        const select = document.getElementById('fleet-filter');
        select.innerHTML = '<option value="">All Fleets</option>';
        data.fleets.forEach(fleet => {
            if(fleet) {
                const opt = document.createElement('option');
                opt.value = fleet;
                opt.textContent = fleet;
                select.appendChild(opt);
            }
        });
    }

}

async function loadEvents(page = 1) {
    currentPage = page;
    const qs = buildQueryString({ page, per_page: 50 });
    const data = await fetchAPI(`/events?${qs}`);
    
    if(data && data.events) {
        renderEventsTable(data.events, 'events-tbody');
        renderPagination(data.total, data.page, data.per_page);
    }
}

async function loadAlertA() {
    const qs = buildQueryString();
    const data = await fetchAPI(`/alerts/install-remove?${qs}`);
    if(data && data.events) {
        let events = data.events;
        if (currentFilters.search) {
            const s = currentFilters.search.toLowerCase();
            events = events.filter(e => 
                (e.aircraft && e.aircraft.toLowerCase().includes(s)) ||
                (e.part_no && e.part_no.toLowerCase().includes(s)) ||
                (e.config_slot_code && e.config_slot_code.toLowerCase().includes(s))
            );
        }
        document.getElementById('alert-a-summary').innerHTML = `<p>Found <strong>${events.length}</strong> installation/removal events matching criteria.</p>`;
        renderEventsTable(events, 'alert-a-tbody');
    }
}

async function loadAlertB() {
    const qs = buildQueryString();
    const data = await fetchAPI(`/alerts/xxx-sn?${qs}`);
    if(data && data.alerts) {
        document.getElementById('alert-b-summary').innerHTML = `<p>Found <strong>${data.total}</strong> critical placeholder S/N issues.</p>`;
        renderAlertBTable(data.alerts, 'alert-b-tbody');
    }
}

async function loadAlertC() {
    const qs = buildQueryString();
    const data = await fetchAPI(`/alerts/mmc?${qs}`);
    if(data && data.alerts) {
        document.getElementById('alert-c-summary').innerHTML = `<p>Found <strong>${data.total}</strong> missing mandatory components (<strong>${data.critical_count}</strong> critical, <strong>${data.warning_count}</strong> warnings).</p>`;
        renderAlertCTable(data.alerts, 'alert-c-tbody');
    }
}

async function loadResolutions() {
    const data = await fetchAPI('/resolutions');
    if(data && data.resolutions) {
        cachedResolutions = data.resolutions;
        renderResolutionsTable(data.resolutions, 'resolutions-tbody');
    }
}

function renderEventsTable(events, tbodyId) {
    const tbody = document.getElementById(tbodyId);
    tbody.innerHTML = '';
    
    if(events.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" style="text-align: center;">No events found</td></tr>';
        return;
    }

    events.forEach(evt => {
        const tr = document.createElement('tr');
        
        let typeBadge = '';
        if(evt.event_type === 'INSTALL') typeBadge = '<span class="badge badge-install">INSTALL</span>';
        else if(evt.event_type === 'REMOVE') typeBadge = '<span class="badge badge-remove">REMOVE</span>';
        else typeBadge = `<span class="badge">${evt.event_type}</span>`;

        let snDisplay = evt.serial_number;
        if(evt.has_xxx_sn) {
            snDisplay = `<span class="badge badge-critical">${evt.serial_number}</span>`;
        }

        tr.innerHTML = `
            <td>${formatDate(evt.event_dt)}</td>
            <td>${typeBadge}</td>
            <td>${evt.aircraft || '-'}</td>
            <td>${evt.config_slot_code || '-'}</td>
            <td>${evt.config_slot_name || '-'}</td>
            <td>${evt.part_no || '-'}</td>
            <td>${evt.part_desc || '-'}</td>
            <td>${snDisplay}</td>
            <td>${evt.status_cd || '-'}</td>
            <td>${evt.barcode || '-'}</td>
            <td>${evt.performed_by_username || evt.performed_by_user || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderAlertBTable(events, tbodyId) {
    const tbody = document.getElementById(tbodyId);
    tbody.innerHTML = '';
    
    if(events.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align: center;">No critical alerts found</td></tr>';
        return;
    }

    events.forEach(evt => {
        const tr = document.createElement('tr');
        
        const deadlineInfo = calculateDeadline(evt.event_dt);
        let deadlineHtml = '';
        if(deadlineInfo.overdue) {
            deadlineHtml = `<span style="color: var(--color-danger); font-weight: bold;">OVERDUE (${Math.abs(deadlineInfo.hours)}h)</span>`;
        } else {
            deadlineHtml = `<span style="color: var(--color-secondary);">${deadlineInfo.hours}h remaining</span>`;
        }

        tr.innerHTML = `
            <td><span class="badge badge-critical">CRITICAL</span></td>
            <td>${formatDate(evt.event_dt)}</td>
            <td>${evt.aircraft || '-'}</td>
            <td>${evt.config_slot_code || '-'}</td>
            <td>${evt.part_no || '-'}</td>
            <td>${evt.part_desc || '-'}</td>
            <td><span class="badge badge-critical">${evt.serial_number}</span></td>
            <td>${deadlineHtml}</td>
            <td>PENDING</td>
            <td>
                <button class="btn btn-primary btn-sm" onclick="resolveAlertB('${evt.barcode}', '${evt.aircraft}', '${evt.config_slot_code}', '${evt.part_no}', '${evt.serial_number}', '${evt.event_dt}')">Resolve</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderAlertCTable(events, tbodyId) {
    const tbody = document.getElementById(tbodyId);
    tbody.innerHTML = '';
    
    if(events.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center;">No MMC alerts found</td></tr>';
        return;
    }

    events.forEach(evt => {
        const tr = document.createElement('tr');
        
        let severityBadge = '';
        if(evt.mmc_severity === 'CRITICAL') severityBadge = '<span class="badge badge-critical">CRITICAL</span>';
        else if(evt.mmc_severity === 'WARNING') severityBadge = '<span class="badge" style="background-color: var(--color-warning); color: #000;">WARNING</span>';
        
        let statusBadge = '';
        if(evt.reinstall_status === 'OVERDUE') statusBadge = '<span class="badge badge-remove">OVERDUE</span>';
        else if(evt.reinstall_status === 'PENDING') statusBadge = '<span class="badge badge-warning">PENDING</span>';
        else if(evt.reinstall_status === 'REINSTALLED_LATE') statusBadge = '<span class="badge badge-install">REINSTALLED (LATE)</span>';

        tr.innerHTML = `
            <td>${formatDate(evt.event_dt)}</td>
            <td>${evt.aircraft || '-'}</td>
            <td>${evt.config_slot_code || '-'}</td>
            <td>${evt.part_no || '-'}</td>
            <td>${evt.part_group_name || '-'}</td>
            <td>${evt.barcode || '-'}</td>
            <td>${evt.performed_by_user || '-'}</td>
            <td style="font-weight: bold; ${evt.days_since_removal >= 7 ? 'color: var(--color-danger);' : ''}">${evt.days_since_removal} days</td>
            <td>${evt.reinstall_date ? formatDate(evt.reinstall_date) : '-'}</td>
            <td>${evt.reinstall_sn || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderResolutionsTable(resolutions, tbodyId) {
    const tbody = document.getElementById(tbodyId);
    tbody.innerHTML = '';
    
    if(resolutions.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" style="text-align: center;">No resolutions logged</td></tr>';
        return;
    }

    resolutions.forEach(res => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${res.id}</td>
            <td>${formatDate(res.alert_date || res.created_at)}</td>
            <td>${res.aircraft || '-'}</td>
            <td>${res.config_slot || '-'}</td>
            <td>${res.part_no || '-'}</td>
            <td>${res.original_sn || '-'}</td>
            <td>${res.resolved_sn || '-'}</td>
            <td>${res.engineer_responsible || '-'}</td>
            <td>${res.resolution_date ? formatDate(res.resolution_date) : '-'}</td>
            <td><span class="badge ${res.status === 'RESOLVED' ? 'badge-install' : 'badge-remove'}">${res.status}</span></td>
            <td>${res.notes || '-'}</td>
            <td>
                <button class="btn btn-secondary btn-sm" onclick="editResolution(${res.id})">Edit</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

window.resolveAlertB = function(barcode, aircraft, configSlot, partNo, originalSn, eventDate) {
    document.getElementById('resolution-form').reset();
    document.getElementById('res-id').value = '';
    document.getElementById('res-barcode').value = barcode;
    document.getElementById('res-aircraft').value = aircraft;
    document.getElementById('res-config-slot').value = configSlot;
    document.getElementById('res-part-no').value = partNo;
    document.getElementById('res-original-sn').value = originalSn;
    
    let dateInput = document.getElementById('res-date');
    if (eventDate) {
        try {
            const d = new Date(eventDate);
            if (!isNaN(d.getTime())) {
                dateInput.value = d.toISOString().split('T')[0];
            }
        } catch(e){}
    }
    
    document.getElementById('resolution-modal').classList.remove('hidden');
};

window.editResolution = function(id) {
    const res = cachedResolutions.find(r => r.id === id);
    if (!res) { showToast('Resolution not found', 'error'); return; }
    document.getElementById('res-id').value = res.id;
    document.getElementById('res-barcode').value = res.barcode || res.event_barcode || '';
    document.getElementById('res-aircraft').value = res.aircraft || '';
    document.getElementById('res-config-slot').value = res.config_slot || '';
    document.getElementById('res-part-no').value = res.part_no || '';
    document.getElementById('res-original-sn').value = res.original_sn || '';
    document.getElementById('res-resolved-sn').value = res.resolved_sn || '';
    document.getElementById('res-engineer').value = res.engineer_responsible || res.engineer || '';
    document.getElementById('res-date').value = res.resolution_date ? res.resolution_date.split('T')[0] : '';
    document.getElementById('res-status').value = res.status || 'PENDING';
    document.getElementById('res-notes').value = res.notes || '';
    document.getElementById('resolution-modal').classList.remove('hidden');
}

async function submitResolution() {
    const id = document.getElementById('res-id').value;
    const payload = {
        barcode: document.getElementById('res-barcode').value,
        aircraft: document.getElementById('res-aircraft').value,
        config_slot: document.getElementById('res-config-slot').value,
        part_no: document.getElementById('res-part-no').value,
        original_sn: document.getElementById('res-original-sn').value,
        resolved_sn: document.getElementById('res-resolved-sn').value,
        engineer: document.getElementById('res-engineer').value,
        resolution_date: document.getElementById('res-date').value,
        status: document.getElementById('res-status').value,
        notes: document.getElementById('res-notes').value
    };

    let url = '/resolutions';
    let method = 'POST';
    if(id) {
        url = `/resolutions/${id}`;
        method = 'PUT';
    }

    try {
        const res = await fetch(API_BASE + url, {
            method: method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        const data = await res.json();
        if(res.ok && data.success) {
            showToast("Resolution saved successfully", "success");
            document.getElementById('resolution-modal').classList.add('hidden');
            if(currentTab === 'tab-resolutions') loadResolutions();
            if(currentTab === 'tab-alert-b') loadAlertB();
        } else {
            showToast(data.error || data.message || "Failed to save resolution", "error");
        }
    } catch (e) {
        showToast("Error saving resolution", "error");
    }
}

async function fetchReports() {
    const overlay = document.getElementById('upload-overlay');
    overlay.classList.remove('hidden');
    
    try {
        const res = await fetch(API_BASE + '/fetch-reports', { method: 'POST' });
        const data = await res.json();
        
        if(res.ok && data.success) {
            const msg = data.files_loaded > 0 
                ? `Loaded ${data.files_loaded} new report(s) with ${data.total_records} records from network share.`
                : 'No new report files found on network share.';
            showToast(msg, data.files_loaded > 0 ? 'success' : 'warning');
            if(data.files_loaded > 0) {
                loadDashboard();
                if(currentTab === 'tab-all-events') loadEvents(1);
                else if(currentTab === 'tab-alert-a') loadAlertA();
                else if(currentTab === 'tab-alert-b') loadAlertB();
            }
        } else {
            showToast(data.error || 'Failed to fetch reports from network share', 'error');
        }
    } catch(e) {
        showToast('Error connecting to network share', 'error');
    } finally {
        overlay.classList.add('hidden');
    }
}

function renderPagination(total, page, perPage) {
    const container = document.getElementById('pagination');
    container.innerHTML = '';
    
    if(total === 0) return;

    const totalPages = Math.ceil(total / perPage);
    
    const prevBtn = document.createElement('button');
    prevBtn.className = 'page-btn';
    prevBtn.innerHTML = '&lt;';
    prevBtn.disabled = page === 1;
    prevBtn.onclick = () => loadEvents(page - 1);
    
    const info = document.createElement('span');
    info.className = 'page-info';
    info.textContent = `Page ${page} of ${totalPages} (${total} total)`;
    
    const nextBtn = document.createElement('button');
    nextBtn.className = 'page-btn';
    nextBtn.innerHTML = '&gt;';
    nextBtn.disabled = page === totalPages;
    nextBtn.onclick = () => loadEvents(page + 1);

    container.appendChild(prevBtn);
    container.appendChild(info);
    container.appendChild(nextBtn);
}


function exportCSV() {
    const qs = buildQueryString();
    window.location.href = API_BASE + `/export?${qs}`;
}

// Utils
function formatDate(dateStr) {
    if(!dateStr) return '';
    try {
        const d = new Date(dateStr);
        if(isNaN(d.getTime())) return dateStr;
        return d.toLocaleString('en-US', { 
            year: 'numeric', month: 'short', day: 'numeric',
            hour: '2-digit', minute:'2-digit'
        });
    } catch(e) {
        return dateStr;
    }
}

function calculateDeadline(eventDateStr) {
    if(!eventDateStr) return { overdue: false, hours: 0 };
    const eventDate = new Date(eventDateStr);
    const now = new Date();
    
    // 24 hours deadline
    const deadline = new Date(eventDate.getTime() + (24 * 60 * 60 * 1000));
    const diffMs = deadline - now;
    const diffHours = Math.round(diffMs / (1000 * 60 * 60));
    
    if(diffHours < 0) {
        return { overdue: true, hours: diffHours };
    }
    return { overdue: false, hours: diffHours };
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let icon = '✅';
    if(type === 'error') icon = '❌';
    if(type === 'warning') icon = '⚠️';
    
    toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function animateCounter(element, target) {
    if(!element) return;
    const duration = 1500; // ms
    const start = parseInt(element.textContent.replace(/,/g, '')) || 0;
    const change = target - start;
    const startTime = performance.now();
    
    function step(timestamp) {
        const progress = Math.min((timestamp - startTime) / duration, 1);
        // easeOutQuart
        const ease = 1 - Math.pow(1 - progress, 4);
        const current = Math.floor(start + (change * ease));
        
        element.textContent = current.toLocaleString();
        
        if (progress < 1) {
            requestAnimationFrame(step);
        } else {
            element.textContent = target.toLocaleString();
        }
    }
    
    requestAnimationFrame(step);
}
