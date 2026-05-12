// ===== Content Calendar Pro =====
(function () {
    'use strict';

    // State
    let state = {
        workbook: null,
        fileName: '',
        sheets: {},        // { name: { headers: [], rows: [] } }
        activeSheet: '',
        mapping: {},       // { date, time, title, account, probability, network, url, notes, status }
        events: [],
        filters: {},
        calendar: null,
        currentView: 'dayGridMonth',
        conflicts: {},
        theme: 'dark',
        tooltipEl: null,
        toastTimer: null,
    };

    // ===== CONSTANTS =====
    const DATE_KEYWORDS = ['fecha', 'date', 'dia', 'día', 'day', 'when', 'inicio', 'start'];
    const TIME_KEYWORDS = ['hora', 'time', 'horario', 'hour', 'clock', 'hora exacta'];
    const TITLE_KEYWORDS = ['titulo', 'título', 'title', 'nombre', 'name', 'video', 'creativo', 'contenido', 'content', 'tema', 'subject'];
    const ACCOUNT_KEYWORDS = ['cuenta', 'account', 'perfil', 'profile', 'canal', 'channel'];
    const PROB_KEYWORDS = ['probabilidad', 'probability', 'prioridad', 'priority', 'nivel'];
    const NET_KEYWORDS = ['red', 'redes', 'network', 'plataforma', 'platform', 'canal'];
    const URL_KEYWORDS = ['url', 'link', 'enlace', 'fuente', 'source', 'href'];
    const NOTES_KEYWORDS = ['notas', 'notes', 'observaciones', 'comentarios', 'comments', 'descripcion', 'description'];
    const STATUS_KEYWORDS = ['estado', 'status', 'situacion', 'fase'];
    const GROUP_KEYWORDS = ['grupo', 'group', 'similitud', 'categoria', 'category', 'tipo', 'type'];
    const OBJECTIVE_KEYWORDS = ['objetivo', 'objective', 'meta', 'goal', 'hipotesis', 'hypothesis'];

    const NETWORK_COLORS = {
        tiktok: '#010101', instagram: '#e1306c', facebook: '#1877f2',
        x: '#000', twitter: '#1da1f2', youtube: '#ff0000', snapchat: '#fffc00',
    };
    const NETWORK_ICONS = {
        tiktok: 'fab fa-tiktok', instagram: 'fab fa-instagram', facebook: 'fab fa-facebook-f',
        x: 'fab fa-x-twitter', twitter: 'fab fa-x-twitter', youtube: 'fab fa-youtube', snapchat: 'fab fa-snapchat',
    };

    const ACCOUNT_COLORS = {
        'cuenta a': { bg: '#3b82f6', cls: 'event-cuenta-a' },
        'cuenta b': { bg: '#ec4899', cls: 'event-cuenta-b' },
        a: { bg: '#3b82f6', cls: 'event-cuenta-a' },
        b: { bg: '#ec4899', cls: 'event-cuenta-b' },
    };

    // ===== INIT =====
    document.addEventListener('DOMContentLoaded', init);

    function init() {
        setupTheme();
        setupUpload();
        setupTopBar();
        setupModals();
        setupFAB();
        ensureFloatingUI();
    }

    // ===== FILE UPLOAD =====
    function setupUpload() {
        const dropZone = document.getElementById('drop-zone');
        const fileInput = document.getElementById('file-input');

        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
        });
        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });
    }

    function handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(ext)) {
            alert('Formato no soportado. Usa .xlsx, .xls o .csv');
            return;
        }
        state.fileName = file.name;
        showLoading('Procesando ' + file.name + '...');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                console.log('[CCP] Parsing workbook...');
                parseWorkbook(e.target.result);
                console.log('[CCP] Parsed OK, sheets:', Object.keys(state.sheets));
                hideLoading();
                console.log('[CCP] Switching to app...');
                switchToApp();
                console.log('[CCP] App ready');
            } catch (err) {
                hideLoading();
                alert('Error: ' + err.message + '\n\nStack: ' + err.stack);
                console.error('[CCP] Error:', err);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function parseWorkbook(data) {
        const wb = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
        state.workbook = wb;
        state.sheets = {};

        wb.SheetNames.forEach((name) => {
            const ws = wb.Sheets[name];
            const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
            if (json.length < 2) {
                state.sheets[name] = { headers: json[0] || [], rows: [], rawRows: [] };
                return;
            }
            // Find header row (first row with >= 3 non-empty cells)
            let headerIdx = 0;
            for (let i = 0; i < Math.min(5, json.length); i++) {
                const filled = json[i].filter(c => c !== '' && c != null).length;
                if (filled >= 3) { headerIdx = i; break; }
            }
            const headers = json[headerIdx].map(h => h != null ? String(h).trim() : '');
            const rows = [];
            const rawJson = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
            for (let i = headerIdx + 1; i < json.length; i++) {
                const row = json[i];
                const rawRow = rawJson[i] || [];
                if (row.every(c => c === '' || c == null)) continue;
                const obj = {};
                headers.forEach((h, idx) => {
                    obj[h] = row[idx] != null ? String(row[idx]).trim() : '';
                    obj['__raw_' + h] = rawRow[idx];
                });
                rows.push(obj);
            }
            state.sheets[name] = { headers, rows, rawRows: rawJson.slice(headerIdx + 1) };
        });
    }

    // ===== COLUMN AUTO-DETECTION =====
    function autoDetectMapping(headers) {
        const mapping = {};
        const lower = headers.map(h => h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''));

        function findCol(keywords) {
            for (const kw of keywords) {
                const idx = lower.findIndex((h, i) => {
                    if (Object.values(mapping).includes(headers[i])) return false;
                    return h.includes(kw);
                });
                if (idx !== -1) return headers[idx];
            }
            return '';
        }

        mapping.date = findCol(DATE_KEYWORDS);
        mapping.time = findCol(TIME_KEYWORDS);
        mapping.title = findCol(TITLE_KEYWORDS);
        mapping.account = findCol(ACCOUNT_KEYWORDS);
        mapping.probability = findCol(PROB_KEYWORDS);
        mapping.network = findCol(NET_KEYWORDS);
        mapping.url = findCol(URL_KEYWORDS);
        mapping.notes = findCol(NOTES_KEYWORDS);
        mapping.status = findCol(STATUS_KEYWORDS);
        mapping.group = findCol(GROUP_KEYWORDS);
        mapping.objective = findCol(OBJECTIVE_KEYWORDS);

        return mapping;
    }

    // ===== SWITCH TO APP =====
    function switchToApp() {
        document.getElementById('upload-screen').classList.remove('active');
        const appScreen = document.getElementById('app-screen');
        appScreen.classList.add('active');
        appScreen.style.display = 'flex';
        appScreen.style.flexDirection = 'column';
        appScreen.style.minHeight = '100vh';

        document.getElementById('file-name-display').textContent = state.fileName;

        // Build sheet tabs
        buildSheetTabs();

        // Pick best sheet (the one with most rows and a date column)
        let bestSheet = '';
        let bestScore = 0;
        Object.entries(state.sheets).forEach(([name, data]) => {
            const mapping = autoDetectMapping(data.headers);
            const score = (mapping.date ? data.rows.length * 2 : 0) + data.rows.length;
            if (score > bestScore) { bestScore = score; bestSheet = name; }
        });
        if (!bestSheet) bestSheet = Object.keys(state.sheets)[0];

        selectSheet(bestSheet);
    }

    function buildSheetTabs() {
        const container = document.getElementById('sheet-tabs');
        container.innerHTML = '';
        Object.entries(state.sheets).forEach(([name, data]) => {
            const btn = document.createElement('button');
            btn.className = 'sheet-tab';
            btn.dataset.sheet = name;
            btn.innerHTML = name + '<span class="tab-count">' + data.rows.length + '</span>';
            btn.addEventListener('click', () => selectSheet(name));
            container.appendChild(btn);
        });
    }

    function selectSheet(name) {
        try {
            console.log('[CCP] selectSheet:', name);
            state.activeSheet = name;
            const data = state.sheets[name];
            state.mapping = autoDetectMapping(data.headers);
            console.log('[CCP] Mapping:', JSON.stringify(state.mapping));
            state.filters = {};

            document.querySelectorAll('.sheet-tab').forEach(t => t.classList.toggle('active', t.dataset.sheet === name));
            document.getElementById('sheet-badge').textContent = name;

            buildEvents();
            console.log('[CCP] Events built:', state.events.length);
            renderStats();
            console.log('[CCP] Stats rendered');
            renderCalendar();
            console.log('[CCP] Calendar rendered');
            renderTable();
            console.log('[CCP] Table rendered');
            buildFilterControls();
            console.log('[CCP] Filters built');
        } catch (err) {
            alert('Error en selectSheet: ' + err.message + '\n\n' + err.stack);
            console.error('[CCP] selectSheet error:', err);
        }
    }

    // ===== BUILD EVENTS =====
    function buildEvents() {
        const data = state.sheets[state.activeSheet];
        const m = state.mapping;
        state.events = [];

        if (!m.date) return;

        data.rows.forEach((row, idx) => {
            const dateVal = row[m.date];
            if (!dateVal) return;

            const parsed = parseDate(dateVal, row['__raw_' + m.date]);
            if (!parsed) return;

            let startStr = parsed;
            if (m.time && row[m.time]) {
                const time = parseTime(row[m.time]);
                if (time) startStr = parsed + 'T' + time;
            }

            const title = m.title ? row[m.title] : ('Evento ' + (idx + 1));
            const account = m.account ? row[m.account] : '';
            const prob = m.probability ? row[m.probability] : '';
            const network = m.network ? row[m.network] : '';
            const url = m.url ? row[m.url] : '';
            const notes = m.notes ? row[m.notes] : '';
            const status = m.status ? row[m.status] : '';
            const group = m.group ? row[m.group] : '';
            const objective = m.objective ? row[m.objective] : '';

            const acctKey = account.toLowerCase().trim();
            const acctColor = ACCOUNT_COLORS[acctKey] || null;
            const probLower = prob.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

            let className = 'event-default';
            if (acctColor) className = acctColor.cls;
            if (probLower.includes('alta')) className += ' event-prob-alta';
            else if (probLower.includes('media')) className += ' event-prob-media';
            else if (probLower.includes('baja')) className += ' event-prob-baja';

            state.events.push({
                id: idx,
                title: title || 'Sin titulo',
                start: startStr,
                allDay: !startStr.includes('T'),
                className: className,
                extendedProps: { account, probability: prob, network, url, notes, status, group, objective, row, rowIdx: idx, baseClassName: className },
            });
        });

        updateConflictState();
    }

    function parseDate(val, rawVal) {
        // Handle Excel serial date
        if (typeof rawVal === 'number' && rawVal > 30000 && rawVal < 70000) {
            const d = excelDateToJS(rawVal);
            return formatDateISO(d);
        }
        // Handle Date objects
        if (rawVal instanceof Date && !isNaN(rawVal)) return formatDateISO(rawVal);
        // Handle string dates
        if (typeof val === 'string') {
            // Try ISO
            const isoMatch = val.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
            if (isoMatch) return isoMatch[1] + '-' + pad(isoMatch[2]) + '-' + pad(isoMatch[3]);
            // Try dd/mm/yyyy
            const dmyMatch = val.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
            if (dmyMatch) return dmyMatch[3] + '-' + pad(dmyMatch[2]) + '-' + pad(dmyMatch[1]);
            // Try date string
            const d = new Date(val);
            if (!isNaN(d) && d.getFullYear() > 2000) return formatDateISO(d);
        }
        return null;
    }

    function parseTime(val) {
        if (!val) return null;
        const s = String(val);
        // HH:MM or H:MM
        const match = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        if (match) return pad(match[1]) + ':' + match[2] + ':' + (match[3] || '00');
        // Decimal hours (0.75 = 18:00)
        const num = parseFloat(s);
        if (!isNaN(num) && num >= 0 && num < 1) {
            const totalMin = Math.round(num * 24 * 60);
            return pad(Math.floor(totalMin / 60)) + ':' + pad(totalMin % 60) + ':00';
        }
        return null;
    }

    function excelDateToJS(serial) {
        const utc_days = Math.floor(serial - 25569);
        const d = new Date(utc_days * 86400 * 1000);
        return d;
    }
    function formatDateISO(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
    function pad(n) { return String(n).padStart(2, '0'); }

    // ===== RENDER CALENDAR =====
    function renderCalendar() {
        const container = document.getElementById('calendar');
        container.innerHTML = '';

        if (state.currentView === 'timeline') {
            document.getElementById('calendar-container').style.display = 'none';
            document.getElementById('timeline-container').style.display = 'block';
            renderTimeline();
            return;
        }

        document.getElementById('calendar-container').style.display = 'block';
        document.getElementById('timeline-container').style.display = 'none';

        const filteredEvents = getFilteredEvents();
        const dayLoadMap = buildDayLoadMap(filteredEvents);

        state.calendar = new FullCalendar.Calendar(container, {
            initialView: state.currentView,
            locale: 'es',
            headerToolbar: { left: 'prev,next today', center: 'title', right: '' },
            events: filteredEvents,
            editable: true,
            eventClick: (info) => showEventDetail(info.event),
            eventDrop: (info) => handleEventDrop(info),
            eventMouseEnter: (info) => showEventTooltip(info.event, info.jsEvent),
            eventMouseLeave: () => hideEventTooltip(),
            eventContent: (arg) => renderEventContent(arg),
            eventDidMount: (info) => {
                info.el.addEventListener('mousemove', moveEventTooltip);
            },
            dayCellDidMount: (info) => decorateDayCell(info, dayLoadMap),
            height: 'auto',
            firstDay: 1,
            nowIndicator: true,
            slotMinTime: '06:00:00',
            slotMaxTime: '23:00:00',
            dayMaxEvents: 4,
            moreLinkText: (n) => '+' + n + ' mas',
            initialDate: findInitialDate(),
            buttonText: { today: 'Hoy' },
        });

        state.calendar.render();
    }

    function findInitialDate() {
        if (state.events.length === 0) return new Date().toISOString().slice(0, 10);
        const dates = state.events.map(e => e.start).filter(Boolean).sort();
        return dates[0].slice(0, 10);
    }

    function renderEventContent(arg) {
        const props = arg.event.extendedProps;
        const time = arg.event.start ? arg.event.start.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }) : '';
        const title = arg.event.title;

        const div = document.createElement('div');
        div.className = 'event-custom';

        let html = '';
        if (time && !arg.event.allDay) html += '<span class="event-time">' + time + '</span>';
        html += '<span class="event-name">' + escapeHtml(title) + '</span>';

        // Network icons
        if (props.network) {
            const nets = detectNetworks(props.network);
            if (nets.length) {
                html += '<span class="event-networks">';
                nets.forEach(n => { html += '<i class="' + (NETWORK_ICONS[n] || 'fas fa-globe') + '"></i>'; });
                html += '</span>';
            }
        }

        if (props.hasConflict) {
            html += '<span class="event-warning" title="Conflicto de horario"><i class="fas fa-triangle-exclamation"></i></span>';
        }

        div.innerHTML = html;
        return { domNodes: [div] };
    }

    function detectNetworks(val) {
        const lower = val.toLowerCase();
        const found = [];
        ['tiktok', 'instagram', 'facebook', 'youtube', 'snapchat', 'x', 'twitter'].forEach(n => {
            if (lower.includes(n)) found.push(n);
        });
        // Check for standalone "X"
        if (!found.includes('x') && /\bx\b/i.test(val)) found.push('x');
        return found;
    }

    // ===== TIMELINE VIEW =====
    function renderTimeline() {
        const container = document.getElementById('timeline-container');
        const events = getFilteredEvents().sort((a, b) => (a.start || '').localeCompare(b.start || ''));

        // Group by date
        const groups = {};
        events.forEach(e => {
            const dateKey = (e.start || '').slice(0, 10);
            if (!groups[dateKey]) groups[dateKey] = [];
            groups[dateKey].push(e);
        });

        let html = '<div class="timeline">';
        Object.entries(groups).sort(([a], [b]) => a.localeCompare(b)).forEach(([dateKey, evts]) => {
            const d = new Date(dateKey + 'T12:00:00');
            const dayName = d.toLocaleDateString('es-MX', { weekday: 'long', day: 'numeric', month: 'long' });
            html += '<div class="timeline-day">';
            html += '<div class="timeline-date">' + dayName + '</div>';

            evts.forEach(evt => {
                const props = evt.extendedProps;
                const time = evt.start && evt.start.includes('T') ? evt.start.split('T')[1].slice(0, 5) : '--:--';
                const tags = buildTags(props);

                html += '<div class="timeline-event" data-idx="' + evt.id + '">';
                html += '<div class="timeline-time">' + time + '</div>';
                html += '<div class="timeline-info">';
                html += '<h4>' + escapeHtml(evt.title) + '</h4>';
                if (props.objective) html += '<p>' + escapeHtml(props.objective.slice(0, 100)) + '</p>';
                html += '<div class="timeline-tags">' + tags + '</div>';
                html += '</div></div>';
            });
            html += '</div>';
        });
        html += '</div>';
        container.innerHTML = html;

        // Click handlers
        container.querySelectorAll('.timeline-event').forEach(el => {
            el.addEventListener('click', () => {
                const idx = parseInt(el.dataset.idx);
                const evt = findStateEventById(idx);
                if (evt) showEventDetailFromData(evt);
            });
        });
    }

    function buildTags(props) {
        let html = '';
        if (props.account) {
            const key = props.account.toLowerCase().includes('a') ? 'a' : 'b';
            html += '<span class="tag tag-' + key + '">' + escapeHtml(props.account) + '</span>';
        }
        if (props.probability) {
            const p = props.probability.toLowerCase();
            const cls = p.includes('alta') ? 'alta' : p.includes('media') ? 'media' : 'baja';
            html += '<span class="tag tag-' + cls + '">' + escapeHtml(props.probability) + '</span>';
        }
        if (props.network) {
            const nets = detectNetworks(props.network);
            nets.forEach(n => { html += '<span class="tag tag-net"><i class="' + (NETWORK_ICONS[n] || '') + '"></i> ' + n + '</span>'; });
        }
        if (props.status) html += '<span class="tag tag-net">' + escapeHtml(props.status) + '</span>';
        return html;
    }

    // ===== STATS =====
    function renderStats() {
        const events = getFilteredEvents();
        const row = document.getElementById('stats-row');

        if (events.length === 0) { row.innerHTML = ''; return; }

        const uniqueDates = new Set(events.map(e => (e.start || '').slice(0, 10)));
        const accounts = {};
        const probs = {};
        const networks = {};
        events.forEach(e => {
            const p = e.extendedProps;
            if (p.account) accounts[p.account] = (accounts[p.account] || 0) + 1;
            if (p.probability) probs[p.probability] = (probs[p.probability] || 0) + 1;
            if (p.network) {
                detectNetworks(p.network).forEach(n => { networks[n] = (networks[n] || 0) + 1; });
            }
        });

        const statCards = [
            { icon: 'fas fa-calendar-check', color: 'purple', value: events.length, label: 'Contenidos totales' },
            { icon: 'fas fa-calendar-day', color: 'blue', value: uniqueDates.size, label: 'Dias programados' },
            { icon: 'fas fa-users', color: 'pink', value: Object.keys(accounts).length, label: 'Cuentas' },
            { icon: 'fas fa-share-alt', color: 'cyan', value: Object.keys(networks).length, label: 'Redes sociales' },
        ];

        // Add probability breakdown
        if (Object.keys(probs).length) {
            const topProb = Object.entries(probs).sort((a, b) => b[1] - a[1])[0];
            statCards.push({ icon: 'fas fa-bullseye', color: 'green', value: topProb[1], label: topProb[0] + ' probabilidad' });
        }

        // Date range
        const sortedDates = [...uniqueDates].sort();
        if (sortedDates.length >= 2) {
            const start = new Date(sortedDates[0] + 'T12:00:00');
            const end = new Date(sortedDates[sortedDates.length - 1] + 'T12:00:00');
            const weeks = Math.ceil((end - start) / (7 * 86400000)) || 1;
            statCards.push({ icon: 'fas fa-clock', color: 'orange', value: weeks, label: 'Semanas de cobertura' });
        }

        row.innerHTML = statCards.map(s => `
            <div class="stat-card">
                <div class="stat-icon ${s.color}"><i class="${s.icon}"></i></div>
                <div class="stat-info"><h4>${s.value}</h4><p>${s.label}</p></div>
            </div>
        `).join('');
    }

    // ===== FILTERS =====
    function buildFilterControls() {
        const data = state.sheets[state.activeSheet];
        const m = state.mapping;
        const panel = document.getElementById('filter-controls');
        panel.innerHTML = '';

        const filterDefs = [
            { key: 'account', label: 'Cuenta', field: m.account },
            { key: 'probability', label: 'Probabilidad', field: m.probability },
            { key: 'network', label: 'Red', field: m.network },
            { key: 'status', label: 'Estado', field: m.status },
            { key: 'group', label: 'Grupo', field: m.group },
        ];

        filterDefs.forEach(def => {
            if (!def.field) return;
            const values = [...new Set(data.rows.map(r => r[def.field]).filter(Boolean))].sort();
            if (values.length <= 1) return;

            const group = document.createElement('div');
            group.className = 'filter-group';
            group.innerHTML = '<label>' + def.label + '</label><div class="filter-options"></div>';
            const opts = group.querySelector('.filter-options');

            values.forEach(v => {
                const chip = document.createElement('button');
                chip.className = 'filter-chip';
                chip.textContent = v;
                chip.addEventListener('click', () => {
                    chip.classList.toggle('active');
                    updateFilters();
                });
                opts.appendChild(chip);
            });

            panel.appendChild(group);
        });

        // Reset button
        const resetBtn = document.createElement('button');
        resetBtn.className = 'filter-reset';
        resetBtn.textContent = 'Limpiar filtros';
        resetBtn.addEventListener('click', () => {
            panel.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
            updateFilters();
        });
        panel.appendChild(resetBtn);
    }

    function updateFilters() {
        const panel = document.getElementById('filter-controls');
        state.filters = {};
        panel.querySelectorAll('.filter-group').forEach(group => {
            const key = group.querySelector('label').textContent.toLowerCase();
            const active = [...group.querySelectorAll('.filter-chip.active')].map(c => c.textContent);
            if (active.length) state.filters[key] = active;
        });
        renderCalendar();
        renderStats();
    }

    function getFilteredEvents() {
        if (Object.keys(state.filters).length === 0) return state.events;

        return state.events.filter(evt => {
            const p = evt.extendedProps;
            for (const [key, values] of Object.entries(state.filters)) {
                let val = '';
                if (key === 'cuenta') val = p.account;
                else if (key === 'probabilidad') val = p.probability;
                else if (key === 'red') val = p.network;
                else if (key === 'estado') val = p.status;
                else if (key === 'grupo') val = p.group;
                if (val && !values.includes(val)) return false;
            }
            return true;
        });
    }

    // ===== TABLE =====
    function renderTable() {
        const data = state.sheets[state.activeSheet];
        const table = document.getElementById('data-table');
        if (!data || !data.headers.length) { table.innerHTML = '<tr><td>Sin datos</td></tr>'; return; }

        let html = '<thead><tr>';
        data.headers.forEach(h => { html += '<th>' + escapeHtml(h) + '</th>'; });
        html += '</tr></thead><tbody>';

        data.rows.forEach(row => {
            html += '<tr>';
            data.headers.forEach(h => {
                let val = row[h] || '';
                if (typeof val === 'string' && val.match(/^https?:\/\//)) {
                    val = formatUrlCell(val);
                } else {
                    val = escapeHtml(val.slice(0, 80));
                }
                html += '<td>' + val + '</td>';
            });
            html += '</tr>';
        });
        html += '</tbody>';
        table.innerHTML = html;

        // Search
        const searchInput = document.getElementById('table-search');
        searchInput.oninput = () => {
            const q = searchInput.value.toLowerCase();
            table.querySelectorAll('tbody tr').forEach(tr => {
                tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
            });
        };
    }

    // ===== EVENT DETAIL =====
    function showEventDetail(calEvent) {
        const props = calEvent.extendedProps;
        const modal = document.getElementById('event-modal');
        document.getElementById('modal-title').textContent = calEvent.title;

        const body = document.getElementById('modal-body');
        const details = [];
        const copyText = buildCopyText(calEvent);

        if (calEvent.start) details.push(['Fecha', calEvent.start.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })]);
        if (calEvent.start && !calEvent.allDay) details.push(['Hora', calEvent.start.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' })]);
        if (props.account) details.push(['Cuenta', props.account]);
        if (props.probability) details.push(['Probabilidad', props.probability]);
        if (props.network) details.push(['Redes', props.network]);
        if (props.group) details.push(['Grupo', props.group]);
        if (props.objective) details.push(['Objetivo', props.objective]);
        if (props.status) details.push(['Estado', props.status]);
        if (props.url) details.push(['URL', formatUrlPreview(props.url)]);
        if (props.notes) details.push(['Notas', props.notes]);

        // Show all other fields from row
        const m = state.mapping;
        const mappedFields = new Set(Object.values(m).filter(Boolean));
        const data = state.sheets[state.activeSheet];
        if (props.row) {
            data.headers.forEach(h => {
                if (mappedFields.has(h)) return;
                if (h.startsWith('__raw_')) return;
                const v = props.row[h];
                if (!v) return;
                if (typeof v === 'string' && v.match(/^https?:\/\//)) {
                    details.push([h, formatUrlPreview(v)]);
                } else {
                    details.push([h, v]);
                }
            });
        }

        body.innerHTML = (props.hasConflict ? `
            <div class="detail-alert">
                <i class="fas fa-triangle-exclamation"></i>
                <div>
                    <strong>Conflicto detectado</strong>
                    <span>Hay ${props.conflictCount} contenidos programados para este mismo horario.</span>
                </div>
            </div>
        ` : '') + `
            <div class="detail-actions">
                <button type="button" class="copy-btn" id="copy-event-btn">
                    <i class="fas fa-copy"></i>
                    <span>Copiar resumen</span>
                </button>
            </div>
        ` + details.map(([label, val]) => {
            const isUrl = label === 'URL' || (typeof val === 'string' && val.match(/^<a |^<div class="url-/));
            return `
            <div class="detail-row">
                <div class="detail-label">${escapeHtml(label)}</div>
                <div class="detail-value">${isUrl ? val : escapeHtml(String(val))}</div>
            </div>`;
        }).join('');

        const copyBtn = document.getElementById('copy-event-btn');
        if (copyBtn) {
            copyBtn.addEventListener('click', async () => {
                const copied = await copyTextToClipboard(copyText);
                showToast(copied ? 'Resumen copiado al portapapeles.' : 'No se pudo copiar el resumen.');
            });
        }

        modal.style.display = 'flex';
    }

    function showEventDetailFromData(evt) {
        // Simulate a calendar event for the detail view
        const fakeEvent = {
            title: evt.title,
            start: evt.start ? new Date(evt.start) : null,
            allDay: evt.allDay,
            extendedProps: evt.extendedProps,
        };
        showEventDetail(fakeEvent);
    }

    // ===== TOP BAR =====
    function setupTopBar() {
        // View toggle
        document.querySelectorAll('.view-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                const view = btn.dataset.view;
                if (view === 'timeline') {
                    state.currentView = 'timeline';
                    renderCalendar();
                } else {
                    state.currentView = view;
                    renderCalendar();
                }
            });
        });

        // Filters panel
        document.getElementById('btn-filters').addEventListener('click', () => togglePanel('filters-panel'));
        document.getElementById('close-filters').addEventListener('click', () => togglePanel('filters-panel', false));

        // Stats panel
        document.getElementById('btn-stats').addEventListener('click', () => {
            renderStatsPanel();
            togglePanel('stats-panel');
        });
        document.getElementById('close-stats').addEventListener('click', () => togglePanel('stats-panel', false));

        // Export
        document.getElementById('btn-export').addEventListener('click', exportCalendar);

        // Theme toggle
        document.getElementById('btn-theme-toggle').addEventListener('click', () => {
            applyTheme(state.theme === 'dark' ? 'light' : 'dark');
        });

        // New file
        document.getElementById('btn-new-file').addEventListener('click', () => {
            document.getElementById('app-screen').classList.remove('active');
            document.getElementById('app-screen').style.display = 'none';
            document.getElementById('upload-screen').classList.add('active');
            state = { ...state, workbook: null, sheets: {}, events: [], activeSheet: '' };
        });

        applyTheme(state.theme, false);
    }

    function togglePanel(id, forceState) {
        const panel = document.getElementById(id);
        const visible = forceState !== undefined ? forceState : panel.style.display === 'none';
        // Close other panels
        document.querySelectorAll('.side-panel').forEach(p => p.style.display = 'none');
        panel.style.display = visible ? 'flex' : 'none';
    }

    // ===== STATS PANEL =====
    function renderStatsPanel() {
        const content = document.getElementById('stats-content');
        const events = getFilteredEvents();
        const m = state.mapping;
        let html = '';

        // By account
        if (m.account) {
            const counts = {};
            events.forEach(e => { const a = e.extendedProps.account || 'Sin cuenta'; counts[a] = (counts[a] || 0) + 1; });
            html += renderBarSection('Por Cuenta', counts, ['#3b82f6', '#ec4899', '#a855f7', '#f59e0b']);
        }

        // By probability
        if (m.probability) {
            const counts = {};
            events.forEach(e => { const p = e.extendedProps.probability || 'Sin dato'; counts[p] = (counts[p] || 0) + 1; });
            html += renderBarSection('Por Probabilidad', counts, ['#22c55e', '#f59e0b', '#ef4444', '#6366f1']);
        }

        // By day of week
        const dayCount = { Lun: 0, Mar: 0, Mie: 0, Jue: 0, Vie: 0, Sab: 0, Dom: 0 };
        const dayNames = ['Dom', 'Lun', 'Mar', 'Mie', 'Jue', 'Vie', 'Sab'];
        events.forEach(e => {
            if (e.start) {
                const d = new Date(e.start);
                if (!isNaN(d)) dayCount[dayNames[d.getDay()]]++;
            }
        });
        html += renderBarSection('Por Dia de la Semana', dayCount, ['#06b6d4']);

        // By network
        if (m.network) {
            const counts = {};
            events.forEach(e => {
                detectNetworks(e.extendedProps.network || '').forEach(n => { counts[n] = (counts[n] || 0) + 1; });
            });
            if (Object.keys(counts).length) html += renderBarSection('Por Red Social', counts, ['#e1306c', '#1877f2', '#010101', '#ff0000', '#fffc00']);
        }

        // By group
        if (m.group) {
            const counts = {};
            events.forEach(e => { const g = e.extendedProps.group || 'Sin grupo'; counts[g] = (counts[g] || 0) + 1; });
            html += renderBarSection('Por Grupo', counts, ['#6366f1', '#3b82f6', '#06b6d4', '#22c55e', '#f59e0b']);
        }

        content.innerHTML = html || '<p style="color:var(--text-dim)">No hay datos suficientes para estadisticas.</p>';
    }

    function renderBarSection(title, counts, colors) {
        const total = Object.values(counts).reduce((a, b) => a + b, 0);
        if (total === 0) return '';
        let html = '<div class="stat-section"><h4>' + title + '</h4>';
        Object.entries(counts).sort((a, b) => b[1] - a[1]).forEach(([label, count], i) => {
            const pct = Math.round((count / total) * 100);
            const color = colors[i % colors.length];
            html += `
                <div class="stat-bar-group">
                    <div class="stat-bar-label"><span>${escapeHtml(label)}</span><span>${count} (${pct}%)</span></div>
                    <div class="stat-bar"><div class="stat-bar-fill" style="width:${pct}%;background:${color}"></div></div>
                </div>`;
        });
        html += '</div>';
        return html;
    }

    // ===== EXPORT =====
    function exportCalendar() {
        const events = getFilteredEvents();
        // Export as ICS
        let ics = 'BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//ContentCalendarPro//EN\n';
        events.forEach(evt => {
            const start = (evt.start || '').replace(/[-:]/g, '').replace('T', 'T');
            ics += 'BEGIN:VEVENT\n';
            ics += 'DTSTART:' + start + '\n';
            ics += 'SUMMARY:' + (evt.title || '') + '\n';
            if (evt.extendedProps.notes) ics += 'DESCRIPTION:' + evt.extendedProps.notes.replace(/\n/g, '\\n') + '\n';
            if (evt.extendedProps.url) ics += 'URL:' + evt.extendedProps.url + '\n';
            ics += 'END:VEVENT\n';
        });
        ics += 'END:VCALENDAR';

        downloadFile(ics, 'calendario_contenido.ics', 'text/calendar');
    }

    function downloadFile(content, name, type) {
        const blob = new Blob([content], { type });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ===== MODALS =====
    function setupModals() {
        document.querySelectorAll('.modal-overlay, .close-modal').forEach(el => {
            el.addEventListener('click', () => {
                el.closest('.modal').style.display = 'none';
            });
        });
    }

    // ===== FAB =====
    function setupFAB() {
        const fab = document.getElementById('fab-table');
        const tableContainer = document.getElementById('table-container');
        fab.addEventListener('click', () => {
            const visible = tableContainer.style.display !== 'none';
            tableContainer.style.display = visible ? 'none' : 'block';
            fab.classList.toggle('active', !visible);
        });
    }

    // ===== LOADING =====
    function showLoading(text) {
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay';
        overlay.id = 'loading';
        overlay.innerHTML = '<div class="spinner"></div><div class="loading-text">' + escapeHtml(text) + '</div>';
        document.body.appendChild(overlay);
    }
    function hideLoading() {
        const el = document.getElementById('loading');
        if (el) el.remove();
    }

    function setupTheme() {
        let savedTheme = 'dark';
        try {
            savedTheme = localStorage.getItem('content-calendar-theme') || 'dark';
        } catch (err) {
            savedTheme = 'dark';
        }
        applyTheme(savedTheme, false);
    }

    function applyTheme(theme, persist = true) {
        state.theme = theme === 'light' ? 'light' : 'dark';
        document.body.classList.toggle('theme-light', state.theme === 'light');

        const button = document.getElementById('btn-theme-toggle');
        if (button) {
            const icon = button.querySelector('i');
            const label = button.querySelector('span');
            button.title = state.theme === 'light' ? 'Cambiar a modo oscuro' : 'Cambiar a modo claro';
            if (icon) icon.className = state.theme === 'light' ? 'fas fa-moon' : 'fas fa-sun';
            if (label) label.textContent = state.theme === 'light' ? 'Dark' : 'Light';
        }

        if (persist) {
            try {
                localStorage.setItem('content-calendar-theme', state.theme);
            } catch (err) {
                // Ignore storage failures.
            }
        }
    }

    function ensureFloatingUI() {
        if (!state.tooltipEl) {
            const tooltip = document.createElement('div');
            tooltip.className = 'event-tooltip';
            tooltip.id = 'event-tooltip';
            document.body.appendChild(tooltip);
            state.tooltipEl = tooltip;
        }
        if (!document.getElementById('app-toast')) {
            const toast = document.createElement('div');
            toast.className = 'app-toast';
            toast.id = 'app-toast';
            document.body.appendChild(toast);
        }
    }

    function buildDayLoadMap(events) {
        return events.reduce((acc, evt) => {
            const key = getEventDateKey(evt.start);
            if (!key) return acc;
            acc[key] = (acc[key] || 0) + 1;
            return acc;
        }, {});
    }

    function decorateDayCell(info, dayLoadMap) {
        if (state.currentView !== 'dayGridMonth') return;
        const dateKey = formatDateISO(info.date);
        const count = dayLoadMap[dateKey] || 0;
        const heatClass = 'heat-level-' + Math.min(count, 4);
        info.el.classList.add(heatClass);
        if (count > 0) {
            const badge = document.createElement('span');
            badge.className = 'heat-count';
            badge.textContent = count + 'x';
            info.el.appendChild(badge);
        }
    }

    function handleEventDrop(info) {
        syncDroppedEvent(info.event);
        updateConflictState();
        renderCalendar();
        renderStats();
        renderTable();
        if (document.getElementById('stats-panel').style.display !== 'none') renderStatsPanel();
        hideEventTooltip();
        showToast('Contenido reprogramado en el calendario.');
    }

    function syncDroppedEvent(eventApi) {
        const stateEvent = findStateEventById(parseInt(eventApi.id, 10));
        if (!stateEvent) return;

        const dateKey = eventApi.start ? formatDateISO(eventApi.start) : '';
        const timeKey = !eventApi.allDay && eventApi.start
            ? pad(eventApi.start.getHours()) + ':' + pad(eventApi.start.getMinutes()) + ':00'
            : '';

        stateEvent.start = dateKey ? (timeKey ? dateKey + 'T' + timeKey : dateKey) : stateEvent.start;
        stateEvent.allDay = eventApi.allDay;

        const row = state.sheets[state.activeSheet] && state.sheets[state.activeSheet].rows[stateEvent.extendedProps.rowIdx];
        if (row) {
            if (state.mapping.date) {
                row[state.mapping.date] = dateKey;
                row['__raw_' + state.mapping.date] = dateKey;
            }
            if (state.mapping.time) {
                row[state.mapping.time] = timeKey ? timeKey.slice(0, 5) : '';
                row['__raw_' + state.mapping.time] = timeKey ? timeKey.slice(0, 5) : '';
            }
            stateEvent.extendedProps.row = row;
        }
    }

    function updateConflictState() {
        const conflicts = {};

        state.events.forEach(evt => {
            const key = getConflictKey(evt.start, evt.allDay);
            if (!key) return;
            if (!conflicts[key]) conflicts[key] = [];
            conflicts[key].push(evt.id);
        });

        state.conflicts = Object.fromEntries(Object.entries(conflicts).filter(([, ids]) => ids.length > 1));

        state.events.forEach(evt => {
            const key = getConflictKey(evt.start, evt.allDay);
            const conflictCount = key && state.conflicts[key] ? state.conflicts[key].length : 0;
            evt.extendedProps.hasConflict = conflictCount > 1;
            evt.extendedProps.conflictCount = conflictCount;
            evt.className = composeEventClassName(evt.extendedProps.baseClassName, evt.extendedProps.hasConflict);
        });
    }

    function composeEventClassName(baseClassName, hasConflict) {
        const tokens = String(baseClassName || '')
            .split(/\s+/)
            .filter(Boolean)
            .filter((token) => token !== 'event-conflict');
        if (hasConflict) tokens.push('event-conflict');
        return tokens.join(' ');
    }

    function getConflictKey(start, allDay) {
        if (!start || allDay || !String(start).includes('T')) return '';
        const [date, time] = String(start).split('T');
        if (!time) return '';
        return date + ' ' + time.slice(0, 5);
    }

    function findStateEventById(id) {
        return state.events.find((evt) => evt.id === id);
    }

    function getEventDateKey(start) {
        return start ? String(start).slice(0, 10) : '';
    }

    function showEventTooltip(event, jsEvent) {
        ensureFloatingUI();
        const tooltip = state.tooltipEl;
        if (!tooltip) return;

        const props = event.extendedProps || {};
        const time = event.start && !event.allDay
            ? event.start.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' })
            : 'Todo el dia';
        const networks = detectNetworks(props.network || '');

        tooltip.innerHTML = `
            <div class="event-tooltip-title">${escapeHtml(event.title)}</div>
            <div class="event-tooltip-meta">
                <div class="event-tooltip-row"><span class="event-tooltip-label">Cuenta</span><span>${escapeHtml(props.account || 'Sin cuenta')}</span></div>
                <div class="event-tooltip-row"><span class="event-tooltip-label">Hora</span><span>${escapeHtml(time)}</span></div>
                <div class="event-tooltip-row"><span class="event-tooltip-label">Prob.</span><span>${escapeHtml(props.probability || 'Sin dato')}</span></div>
                ${props.hasConflict ? `<div class="event-tooltip-row"><span class="event-tooltip-label">Warning</span><span>${props.conflictCount} en conflicto</span></div>` : ''}
            </div>
            ${networks.length ? `<div class="event-tooltip-networks">${networks.map((network) => `<span class="tooltip-pill"><i class="${NETWORK_ICONS[network] || 'fas fa-globe'}"></i>${escapeHtml(network)}</span>`).join('')}</div>` : ''}
        `;
        tooltip.classList.add('visible');
        positionTooltip(jsEvent);
    }

    function moveEventTooltip(jsEvent) {
        if (!state.tooltipEl || !state.tooltipEl.classList.contains('visible')) return;
        positionTooltip(jsEvent);
    }

    function positionTooltip(jsEvent) {
        const tooltip = state.tooltipEl;
        if (!tooltip || !jsEvent) return;
        const offset = 18;
        const width = tooltip.offsetWidth || 240;
        const height = tooltip.offsetHeight || 120;
        const maxLeft = window.innerWidth - width - 12;
        const maxTop = window.innerHeight - height - 12;
        const left = Math.min(jsEvent.clientX + offset, maxLeft);
        const top = Math.min(jsEvent.clientY + offset, maxTop);
        tooltip.style.left = Math.max(12, left) + 'px';
        tooltip.style.top = Math.max(12, top) + 'px';
    }

    function hideEventTooltip() {
        if (state.tooltipEl) state.tooltipEl.classList.remove('visible');
    }

    function buildCopyText(calEvent) {
        if (!calEvent.start) return calEvent.title;
        const dateLabel = calEvent.start.toLocaleDateString('es-MX', { year: 'numeric', month: '2-digit', day: '2-digit' });
        const timeLabel = calEvent.allDay ? 'Todo el dia' : calEvent.start.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
        return calEvent.title + ' | ' + dateLabel + ' | ' + timeLabel;
    }

    async function copyTextToClipboard(text) {
        try {
            await navigator.clipboard.writeText(text);
            return true;
        } catch (err) {
            const fallback = document.createElement('textarea');
            fallback.value = text;
            fallback.setAttribute('readonly', '');
            fallback.style.position = 'fixed';
            fallback.style.opacity = '0';
            document.body.appendChild(fallback);
            fallback.select();
            const copied = document.execCommand('copy');
            document.body.removeChild(fallback);
            return copied;
        }
    }

    function showToast(message) {
        const toast = document.getElementById('app-toast');
        if (!toast) return;
        toast.textContent = message;
        toast.classList.add('visible');
        if (state.toastTimer) window.clearTimeout(state.toastTimer);
        state.toastTimer = window.setTimeout(() => {
            toast.classList.remove('visible');
        }, 2200);
    }

    // ===== URL HELPERS =====
    function detectUrlPlatform(url) {
        const u = url.toLowerCase();
        if (u.includes('instagram.com')) return { name: 'Instagram', icon: 'fab fa-instagram', color: '#e1306c' };
        if (u.includes('facebook.com')) return { name: 'Facebook', icon: 'fab fa-facebook-f', color: '#1877f2' };
        if (u.includes('tiktok.com')) return { name: 'TikTok', icon: 'fab fa-tiktok', color: '#010101' };
        if (u.includes('youtube.com') || u.includes('youtu.be')) return { name: 'YouTube', icon: 'fab fa-youtube', color: '#ff0000' };
        if (u.includes('twitter.com') || u.includes('x.com')) return { name: 'X', icon: 'fab fa-x-twitter', color: '#000' };
        if (u.includes('snapchat.com')) return { name: 'Snapchat', icon: 'fab fa-snapchat', color: '#fffc00' };
        return { name: 'Web', icon: 'fas fa-globe', color: '#6366f1' };
    }

    function extractUrlPath(url) {
        try {
            const u = new URL(url);
            const path = u.pathname.replace(/\/$/, '');
            const parts = path.split('/').filter(Boolean);
            if (parts.length > 2) return '/' + parts.slice(-2).join('/');
            return path || u.hostname;
        } catch (e) { return url.slice(0, 40); }
    }

    function formatUrlCell(url) {
        const platform = detectUrlPlatform(url);
        const path = extractUrlPath(url);
        return '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener" class="url-cell" title="' + escapeHtml(url) + '">' +
            '<i class="' + platform.icon + '" style="color:' + platform.color + ';margin-right:4px"></i>' +
            '<span>' + escapeHtml(platform.name) + '</span>' +
            '<span class="url-path">' + escapeHtml(path) + '</span>' +
            '</a>';
    }

    function formatUrlPreview(url) {
        const platform = detectUrlPlatform(url);
        const igMatch = url.match(/instagram\.com\/(p|reels?|reel)\/([A-Za-z0-9_-]+)/);
        const embedHtml = igMatch
            ? '<div class="url-embed"><iframe src="https://www.instagram.com/' + igMatch[1] + '/' + igMatch[2] + '/embed" ' +
              'width="100%" height="380" frameborder="0" scrolling="no" allowtransparency="true" loading="lazy"></iframe></div>'
            : '';

        return '<div class="url-preview">' +
            '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener" class="url-preview-link">' +
            '<i class="' + platform.icon + '" style="color:' + platform.color + '"></i> ' +
            '<span>' + escapeHtml(platform.name) + '</span>' +
            '<span class="url-full">' + escapeHtml(url) + '</span>' +
            '<i class="fas fa-external-link-alt url-external"></i>' +
            '</a>' + embedHtml + '</div>';
    }

    // ===== UTILS =====
    function escapeHtml(s) {
        if (!s) return '';
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
})();
