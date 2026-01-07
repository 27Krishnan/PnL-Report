document.addEventListener('DOMContentLoaded', () => {
    const tableBody = document.getElementById('table-body');
    const addRowBtn = document.getElementById('add-row-btn');

    // Modal Elements
    const manageModal = document.getElementById('manage-modal');
    const manageModalTitle = manageModal.querySelector('h2');
    const closeModalBtn = document.getElementById('close-modal-btn');
    const managementList = document.getElementById('owner-management-list'); // Keeping ID for CSS, but content is dynamic

    // State
    const DATA_DEFAULTS = {
        owners: ['Owner 1', 'Owner 2'],
        types: ['Intraday', 'Delivery', 'F&O', 'Currency', 'Commodity']
    };

    let appData = {
        owners: [],
        types: []
    };

    // UI State
    let globalDropdown = null;
    let activeInput = null;
    let activeSourceKey = null; // 'owners' or 'types'
    let managingSourceKey = null; // For the modal
    let hideTimeout = null; // To manage blur race conditions
    let ownerChart = null; // Chart Instance
    let typeChart = null; // Type Chart Instance
    let dailyChart = null; // Daily Chart Instance

    // Settings State
    let settingsGASUrl = localStorage.getItem('gas_url') || '';
    let settingsAutoSync = localStorage.getItem('auto_sync') === 'true';
    let syncTimeout = null;

    // --- Initialization ---
    // If we have saved rows, they will be loaded. If not, we might want to add one blank row.
    loadAllData();
    // Use timeout to ensure UI is ready or just check if table empty
    if (tableBody.children.length === 0) {
        addRow();
    }
    createGlobalDropdown();

    // View Toggle Logic
    // View Toggle Logic
    const btnViewTable = document.getElementById('btn-view-table');
    const btnViewDashboard = document.getElementById('btn-view-dashboard');
    const viewTable = document.getElementById('view-table');
    const viewDashboard = document.getElementById('view-dashboard');

    function switchView(viewName) {
        // Reset all
        [viewTable, viewDashboard].forEach(el => el.classList.add('hidden'));
        [btnViewTable, btnViewDashboard].forEach(el => el.classList.remove('active'));

        if (viewName === 'table') {
            viewTable.classList.remove('hidden');
            btnViewTable.classList.add('active');
        } else if (viewName === 'dashboard') {
            viewDashboard.classList.remove('hidden');
            btnViewDashboard.classList.add('active');
            updateDashboard(); // Refresh stats when opening
        }
    }

    btnViewTable.addEventListener('click', () => switchView('table'));
    btnViewDashboard.addEventListener('click', () => switchView('dashboard'));


    addRowBtn.addEventListener('click', () => {
        addRow();
    });

    // --- Dashboard Logic ---
    function updateDashboard() {
        const trs = tableBody.querySelectorAll('tr');
        let totalPL = 0;
        let totalTrades = 0;
        let wins = 0;
        let grossWin = 0;
        let grossLoss = 0;

        trs.forEach(tr => {
            // Find P/L input
            const inputs = tr.querySelectorAll('input');
            // Based on index: 4 is P/L
            if (inputs.length > 4) {
                const plVal = parseFloat(inputs[4].value);
                if (!isNaN(plVal)) {
                    totalTrades++;
                    totalPL += plVal;
                    if (plVal > 0) {
                        wins++;
                        grossWin += plVal;
                    } else if (plVal < 0) {
                        grossLoss += Math.abs(plVal);
                    }
                }
            }
        });

        // Update DOM
        const elNetPL = document.getElementById('dash-net-pl');
        const elTrades = document.getElementById('dash-total-trades');
        const elWinRate = document.getElementById('dash-win-rate');
        const elProfitFactor = document.getElementById('dash-profit-factor');

        if (elNetPL) {
            elNetPL.textContent = totalPL.toFixed(2);
            elNetPL.className = 'metric-value ' + (totalPL >= 0 ? 'positive' : 'negative');
        }

        if (elTrades) elTrades.textContent = totalTrades;

        if (elWinRate) {
            const rate = totalTrades > 0 ? ((wins / totalTrades) * 100).toFixed(1) : '0.0';
            elWinRate.textContent = rate + '%';
        }

        if (elProfitFactor) {
            const pf = grossLoss > 0 ? (grossWin / grossLoss).toFixed(2) : (grossWin > 0 ? '∞' : '0.00');
            elProfitFactor.textContent = pf;
        }

        // Update Report Table
        if (typeof updateOwnerReport === 'function') {
            updateOwnerReport();
        }
        if (typeof updateTypeReport === 'function') {
            updateTypeReport();
        }
        if (typeof updateDailyChart === 'function') {
            updateDailyChart();
        }
        if (typeof updateCumulativePL === 'function') {
            updateCumulativePL();
        }
        if (typeof updateDetailedReport === 'function') {
            updateDetailedReport();
        }
    }

    // --- Owner P&L Report Logic ---
    function updateOwnerReport() {
        const trs = tableBody.querySelectorAll('tr');
        const ownerStats = {};

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // inputs[1] is Owner 
            if (inputs.length >= 5) {
                const owner = inputs[1].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (owner && !isNaN(pl)) {
                    if (!ownerStats[owner]) {
                        ownerStats[owner] = { pl: 0, trades: 0 };
                    }
                    ownerStats[owner].pl += pl;
                    ownerStats[owner].trades += 1;
                }
            }
        });

        // Convert to array and sort by P/L descending
        const sortedOwners = Object.keys(ownerStats).map(key => ({
            name: key,
            ...ownerStats[key]
        })).sort((a, b) => b.pl - a.pl);

        // Render
        const tbody = document.getElementById('owner-pl-body');
        if (!tbody) return;
        tbody.innerHTML = '';

        // Helper to generate consistent colors - USING SHARED FUNCTION

        sortedOwners.forEach((stat, index) => {
            const tr = document.createElement('tr');
            const color = getColorForString(stat.name);

            // Enable Drill Down
            tr.style.cursor = 'pointer';
            tr.title = 'Click to view trades';
            tr.onclick = () => showDrillDown('Owner', stat.name);

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td style="color: ${color}; font-weight: 500;">${stat.name}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'}">${stat.pl.toFixed(2)}</td>
                <td>${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });

        // Update Chart
        if (typeof updateChart === 'function') {
            updateChart(sortedOwners);
        }
    }

    // --- Type Performance Report Logic ---
    function updateTypeReport() {
        const trs = tableBody.querySelectorAll('tr');
        const typeStats = {};

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // inputs[2] is Type (Date=0, Owner=1, Type=2...)
            if (inputs.length >= 5) {
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (type && !isNaN(pl)) {
                    if (!typeStats[type]) {
                        typeStats[type] = { pl: 0, trades: 0 };
                    }
                    typeStats[type].pl += pl;
                    typeStats[type].trades += 1;
                }
            }
        });

        // Convert to array and sort by P/L descending
        const sortedTypes = Object.keys(typeStats).map(key => ({
            name: key,
            ...typeStats[key]
        })).sort((a, b) => b.pl - a.pl);

        // Render
        const tbody = document.getElementById('type-pl-body');
        if (!tbody) return;
        tbody.innerHTML = '';

        sortedTypes.forEach((stat, index) => {
            const tr = document.createElement('tr');

            // Enable Drill Down
            tr.style.cursor = 'pointer';
            tr.title = 'Click to view trades';
            tr.onclick = () => showDrillDown('Type', stat.name);

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${stat.name}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'}">${stat.pl.toFixed(2)}</td>
                <td>${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });

        // Update Type Chart
        if (typeof updateTypeChart === 'function') {
            updateTypeChart(sortedTypes);
        }
    }
    // --- Detailed Report Logic ---
    function updateDetailedReport() {
        // Aggregate by Owner + Type
        const trs = document.getElementById('table-body').querySelectorAll('tr');
        const stats = {};

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // inputs[1] = Owner, inputs[2] = Type, inputs[4] = P/L
            if (inputs.length >= 5) {
                const owner = inputs[1].value.trim();
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (owner && type && !isNaN(pl)) {
                    const key = `${owner}|${type}`;
                    if (!stats[key]) {
                        stats[key] = { owner, type, pl: 0, trades: 0 };
                    }
                    stats[key].pl += pl;
                    stats[key].trades += 1;
                }
            }
        });

        // Convert and sort
        const sortedStats = Object.values(stats).sort((a, b) => b.pl - a.pl);

        // Render
        const tbody = document.getElementById('detailed-pl-body');
        if (!tbody) return;
        tbody.innerHTML = '';

        sortedStats.forEach((stat, index) => {
            const tr = document.createElement('tr');
            const color = getColorForString(stat.owner);

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td style="color: ${color}; font-weight: 500;">${stat.owner}</td>
                <td>${stat.type}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'}">${stat.pl.toFixed(2)}</td>
                <td>${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    // --- Chart Logic --- //
    // Generic Chart Config Creator
    function createChartConfig(label) {
        return {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: label,
                    data: [],
                    backgroundColor: [],
                    borderColor: 'rgba(0,0,0,0.1)',
                    borderWidth: 1,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255, 255, 255, 0.1)' },
                        ticks: { color: '#8b949e' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#8b949e' }
                    }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function (context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                if (context.parsed.y !== null) label += context.parsed.y.toFixed(2);
                                return label;
                            }
                        }
                    }
                }
            }
        };
    }

    function initChart() {
        const ctxOwner = document.getElementById('ownerChart').getContext('2d');
        ownerChart = new Chart(ctxOwner, createChartConfig('Net P/L'));

        const ctxType = document.getElementById('typeChart').getContext('2d');
        typeChart = new Chart(ctxType, createChartConfig('Net P/L'));

        const ctxDaily = document.getElementById('dailyChart').getContext('2d');
        dailyChart = new Chart(ctxDaily, createChartConfig('Daily P/L'));
    }

    function updateChart(ownersData) {
        if (!ownerChart) initChart();
        updateChartData(ownerChart, ownersData);
    }

    function updateTypeChart(typesData) {
        if (!typeChart) initChart();
        updateChartData(typeChart, typesData);
    }

    // --- Daily Chart Logic ---
    function updateDailyChart() {
        if (!dailyChart) initChart();

        // Aggregate by Date
        const trs = document.getElementById('table-body').querySelectorAll('tr');
        const dailyStats = {};

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // inputs[3] is Exit Date
            if (inputs.length >= 5) {
                const date = inputs[3].value;
                const pl = parseFloat(inputs[4].value);

                if (date && !isNaN(pl)) {
                    if (!dailyStats[date]) {
                        dailyStats[date] = 0;
                    }
                    dailyStats[date] += pl;
                }
            }
        });

        // Sort by Date
        const sortedDates = Object.keys(dailyStats).sort((a, b) => new Date(a) - new Date(b));

        const data = sortedDates.map(date => dailyStats[date]);

        // Colors: Green for >= 0, Red for < 0
        const colors = data.map(val => val >= 0 ? 'rgba(16, 185, 129, 0.7)' : 'rgba(239, 68, 68, 0.7)');

        dailyChart.data.labels = sortedDates;
        dailyChart.data.datasets[0].data = data;
        dailyChart.data.datasets[0].backgroundColor = colors;
        dailyChart.data.datasets[0].borderColor = colors.map(c => c.replace('0.7', '1'));
        dailyChart.update();
    }

    function updateChartData(chartInstance, dataArray) {
        const labels = dataArray.map(d => d.name);
        const data = dataArray.map(d => d.pl);
        const colors = dataArray.map(d => getColorForString(d.name));

        chartInstance.data.labels = labels;
        chartInstance.data.datasets[0].data = data;
        chartInstance.data.datasets[0].backgroundColor = colors;
        chartInstance.update();
    }

    // --- Modal Logic ---
    closeModalBtn.addEventListener('click', closeModal);
    manageModal.addEventListener('click', (e) => {
        if (e.target === manageModal) closeModal();
    });

    function openModal(sourceKey) {
        managingSourceKey = sourceKey;
        const titleMap = { owners: 'Manage Owners', types: 'Manage Types' };
        manageModalTitle.textContent = titleMap[sourceKey] || 'Manage Items';
        renderManageList();
        manageModal.classList.remove('hidden');
    }

    function closeModal() {
        manageModal.classList.add('hidden');
        managingSourceKey = null;
    }

    function renderManageList(filterText = '') {
        managementList.innerHTML = '';
        let list = appData[managingSourceKey] || [];

        // Filter
        if (filterText) {
            const lowerFilter = filterText.toLowerCase();
            list = list.filter(item => item.toLowerCase().includes(lowerFilter));
        }

        if (list.length === 0) {
            managementList.innerHTML = '<li style="text-align:center; color: var(--text-secondary); padding: 1rem;">No items found</li>';
            return;
        }

        list.forEach(item => {
            const li = document.createElement('li');
            li.className = 'owner-item';

            const span = document.createElement('span');
            span.className = 'owner-name';
            span.textContent = item;

            const delBtn = document.createElement('button');
            delBtn.className = 'delete-btn';
            delBtn.style.opacity = '1';
            delBtn.title = 'Delete Item';
            delBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"></path>
                </svg>
            `;
            delBtn.onclick = () => deleteItem(item);

            li.appendChild(span);
            li.appendChild(delBtn);
            managementList.appendChild(li);
        });
    }

    function deleteItem(itemToDelete) {
        if (confirm(`Delete "${itemToDelete}"?`)) {
            appData[managingSourceKey] = appData[managingSourceKey].filter(i => i !== itemToDelete);
            saveData(managingSourceKey);
            renderManageList();
        }
    }

    // --- Data Logic ---
    function loadAllData() {
        // Load Dropdown Options
        ['owners', 'types'].forEach(key => {
            let data = JSON.parse(localStorage.getItem(key));
            if (!data) {
                data = DATA_DEFAULTS[key];
            }
            data = data.filter(i => i && i.trim().length > 0);
            data = [...new Set(data)];
            appData[key] = data;
            saveData(key);
        });

        // Load Table Rows
        loadTableRows();

        // Initial Dashboard Update to populate stats and charts
        updateDashboard();
    }

    function saveData(key) {
        if (appData[key]) {
            localStorage.setItem(key, JSON.stringify(appData[key]));
        }
    }
    // New: Save Table Rows
    function saveTableRows() {
        const rows = [];
        const trs = tableBody.querySelectorAll('tr');
        trs.forEach(tr => {
            // Get inputs
            // 0: S.No (text), 1: Date, 2: Owner, 3: Type, 4: Exit, 5: PL, 6: Remark
            const inputs = tr.querySelectorAll('input, select');
            // We expect: Date, Owner, Type, Exit, PL, Remark (6 inputs if standard)
            // But let's be robust.
            // Indices in querySelectorAll might vary if disabled/hidden, but here likely consistent.
            // inputs[0]=Date, inputs[1]=Owner, inputs[2]=Type, inputs[3]=Exit, inputs[4]=PL, inputs[5]=Remark
            if (inputs.length >= 6) {
                rows.push({
                    date: inputs[0].value,
                    owner: inputs[1].value,
                    type: inputs[2].value,
                    exitDate: inputs[3].value,
                    pl: inputs[4].value,
                    remark: inputs[5].value,
                    lastEdited: tr.dataset.lastEdited || null,
                    lastEditedMsg: tr.dataset.lastEditedMsg || null
                });
            }
        });
        localStorage.setItem('pl_report_rows', JSON.stringify(rows));

        // Also trigger dashboard update if it exists
        if (typeof updateDashboard === 'function') {
            updateDashboard();
        }

        // Auto Sync Logic
        if (settingsAutoSync && settingsGASUrl) {
            clearTimeout(syncTimeout);
            syncTimeout = setTimeout(() => {
                syncToGoogleSheets(rows);
            }, 2000); // 2 second debounce
        }
    }

    // --- Google Sheets Sync Logic ---
    function syncToGoogleSheets(rowsData) {
        if (!settingsGASUrl) return;

        const btnSave = document.getElementById('btn-save-settings'); // For visual feedback if open
        if (btnSave) btnSave.textContent = 'Syncing...';

        fetch(settingsGASUrl, {
            method: 'POST',
            mode: 'no-cors', // Standard for GAS Web App calls (opaque response)
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ rows: rowsData })
        })
            .then(() => {
                console.log('Synced to Google Sheets');
                if (btnSave) btnSave.textContent = 'Save & Sync';
            })
            .catch(err => {
                console.error('Sync Error:', err);
                if (btnSave) btnSave.textContent = 'Error (See Console)';
            });
    }

    function loadTableRows() {
        const savedRows = JSON.parse(localStorage.getItem('pl_report_rows'));
        if (savedRows && Array.isArray(savedRows)) {
            // Clear current initial row if it was added by default and empty? 
            // Actually 'addRow()' is called at init. Let's clear tableBody first if we have data.
            if (savedRows.length > 0) {
                tableBody.innerHTML = '';
                savedRows.forEach(row => {
                    addRow(row); // Pass data to addRow
                });
            }
        }
    }

    function addItemIfNew(key, newItem) {
        const trimmed = newItem ? newItem.trim() : '';
        if (trimmed && !appData[key].includes(trimmed)) {
            appData[key].push(trimmed);
            saveData(key);
        }
    }

    // --- Dropdown Logic ---
    function createGlobalDropdown() {
        globalDropdown = document.createElement('div');
        globalDropdown.className = 'dropdown-list';
        globalDropdown.style.position = 'fixed';
        globalDropdown.style.display = 'none';
        document.body.appendChild(globalDropdown);

        // Prevent input blur when clicking on the dropdown container (scrollbar, footer, etc.)
        globalDropdown.addEventListener('mousedown', (e) => {
            e.preventDefault();
        });

        // Window events
        window.addEventListener('resize', hideDropdown);
        window.addEventListener('scroll', (e) => {
            // Ignore scroll events coming from the dropdown itself
            if (e.target === globalDropdown || globalDropdown.contains(e.target)) {
                return;
            }
            hideDropdown();
        }, true);
    }

    function showDropdown(input) {
        // Clear any pending hide (e.g. from blurring previous field)
        if (hideTimeout) {
            clearTimeout(hideTimeout);
            hideTimeout = null;
        }

        activeInput = input;
        activeSourceKey = input.dataset.source;

        if (!activeSourceKey || !appData[activeSourceKey]) return;

        updateDropdownContent(input.value);
        positionDropdown(input);
        globalDropdown.style.display = 'block';
    }

    function hideDropdown() {
        if (globalDropdown) {
            globalDropdown.style.display = 'none';
        }
    }

    function positionDropdown(input) {
        if (!input || !globalDropdown) return;
        const rect = input.getBoundingClientRect();
        globalDropdown.style.top = `${rect.bottom + 4}px`;
        globalDropdown.style.left = `${rect.left}px`;
        globalDropdown.style.width = `${rect.width}px`;
    }

    function updateDropdownContent(filterText) {
        globalDropdown.innerHTML = '';

        const list = appData[activeSourceKey] || [];
        const filtered = list.filter(i => i.toLowerCase().includes(filterText.toLowerCase()));

        // Items
        filtered.forEach(item => {
            const div = document.createElement('div');
            div.className = 'dropdown-item';

            const span = document.createElement('span');
            span.className = 'owner-text';
            span.textContent = item;
            div.appendChild(span);

            div.addEventListener('mousedown', (e) => {
                e.preventDefault();
                if (activeInput) {
                    activeInput.value = item;
                    addItemIfNew(activeSourceKey, item);
                    hideDropdown();
                    // Optional: auto-advance
                }
            });

            globalDropdown.appendChild(div);
        });

        // Footer
        const footer = document.createElement('div');
        footer.className = 'dropdown-footer';

        const editBtn = document.createElement('button');
        editBtn.className = 'btn-edit-list';
        editBtn.title = `Manage ${activeSourceKey}`;
        editBtn.innerHTML = `
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M12 20h9"></path>
                <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
            </svg>
        `;

        editBtn.addEventListener('mousedown', (e) => {
            e.preventDefault();
            hideDropdown();
            openModal(activeSourceKey);
        });

        footer.appendChild(editBtn);
        globalDropdown.appendChild(footer);
    }

    // --- Table & Cell Logic ---
    function addRow(data = null) {
        const tr = document.createElement('tr');

        // S.No
        const snoTd = document.createElement('td');
        snoTd.className = 'sno-cell';
        tr.appendChild(snoTd);

        // Date
        const dateTd = createCell('date');
        const dateInput = dateTd.querySelector('input');
        if (data && data.date) {
            dateInput.value = data.date;
        } else {
            // Set default to today if new row
            const today = new Date();
            const yyyy = today.getFullYear();
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            const dd = String(today.getDate()).padStart(2, '0');
            dateInput.value = `${yyyy}-${mm}-${dd}`;
        }
        addSaveTrigger(dateInput);
        tr.appendChild(dateTd);

        // Owner (Smart Input)
        const ownerTd = createSmartCell('owners', 'Owner', data ? data.owner : '');
        tr.appendChild(ownerTd);

        // Type (Smart Input)
        const typeTd = createSmartCell('types', 'Type', data ? data.type : '');
        tr.appendChild(typeTd);

        // Exit Date
        const exitDateTd = createCell('date');
        if (data && data.exitDate) exitDateTd.querySelector('input').value = data.exitDate;
        addSaveTrigger(exitDateTd.querySelector('input'));
        tr.appendChild(exitDateTd);

        // P/L
        const plTd = createCell('number', '0.00');
        plTd.classList.add('pl-cell');
        const plInput = plTd.querySelector('input');
        if (data && data.pl) {
            plInput.value = data.pl;
            // Trigger color update manually
            updatePLColor({ target: plInput });
        }
        plInput.addEventListener('input', updatePLColor);
        addSaveTrigger(plInput);

        // --- Auto-Date & Validation ---
        const exitDateInput = exitDateTd.querySelector('input');

        const validateDate = (input) => {
            if (!input.value) input.classList.add('missing-date');
            else input.classList.remove('missing-date');
        };

        // Bind Validation
        [dateInput, exitDateInput].forEach(input => {
            input.addEventListener('input', () => validateDate(input));
            input.addEventListener('blur', () => validateDate(input));
            validateDate(input); // Initial state
        });

        // Auto-Fill Exit Date on P/L Entry
        plInput.addEventListener('input', () => {
            if (!exitDateInput.value) {
                const today = new Date().toISOString().split('T')[0];
                exitDateInput.value = today;
                validateDate(exitDateInput);
                saveTableRows();
            }
        });

        // Enter Navigation for P/L
        plInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                // Move to next row
                const nextTr = tr.nextElementSibling;
                if (nextTr) {
                    const firstInput = nextTr.querySelector('input');
                    if (firstInput) firstInput.focus();
                } else {
                    // Create new row and focus it
                    addRow();
                }
            }
        });

        tr.appendChild(plTd);

        // Cumulative P/L (Read Only)
        const cumTd = document.createElement('td');
        cumTd.className = 'cum-cell';
        cumTd.style.fontWeight = '500';
        cumTd.textContent = '--';
        tr.appendChild(cumTd);

        // REMARK
        const remarkTd = createCell('text', 'Remarks');
        if (data && data.remark) remarkTd.querySelector('input').value = data.remark;
        addSaveTrigger(remarkTd.querySelector('input'));
        tr.appendChild(remarkTd);

        // Actions
        const actionsTd = document.createElement('td');
        actionsTd.className = 'actions-col';
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'delete-btn';
        deleteBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"></path>
            </svg>
        `;
        deleteBtn.onclick = () => {
            requestDelete(tr);
        };
        actionsTd.appendChild(deleteBtn);
        tr.appendChild(actionsTd);

        // Restore Edited Timestamp
        // Restore Edited Timestamp
        if (data) {
            if (data.lastEditedMsg) {
                tr.dataset.lastEditedMsg = data.lastEditedMsg;
                tr.dataset.lastEdited = data.lastEdited;
                if (typeof renderEditTimestamp === 'function') renderEditTimestamp(tr, data.lastEditedMsg);
            } else if (data.lastEdited) {
                // Legacy fallback
                const date = new Date(data.lastEdited);
                const timestamp = date.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit' }) + ' ' +
                    date.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
                tr.dataset.lastEdited = data.lastEdited;
                if (typeof renderEditTimestamp === 'function') renderEditTimestamp(tr, 'Edited: ' + timestamp);
            }
        }

        tableBody.appendChild(tr);
        updateSerialNumbers();

        // Save on creation IF it's a manual add (not loading)
        // Actually, simplest is to just save whenever the serial numbers update? 
        // Or if we passed no data, it's a new row, so save.
        if (!data) {
            saveTableRows();
            // Scroll to center for better visibility
            requestAnimationFrame(() => {
                tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
            });
        }

        // Focus on first input (Date) - Child index 1 now (0 is S.No)
        const focusInput = tr.children[1].querySelector('input');
        if (!data && focusInput) focusInput.focus();
    }

    function addSaveTrigger(input) {
        input.addEventListener('input', (e) => {
            if (typeof updateRowTimestamp === 'function') updateRowTimestamp(e.target);
            saveTableRows();
        });
        input.addEventListener('change', (e) => {
            if (typeof updateRowTimestamp === 'function') updateRowTimestamp(e.target);
            saveTableRows();
        });
    }

    function updateSerialNumbers() {
        const rows = tableBody.querySelectorAll('tr');
        rows.forEach((row, index) => {
            const cell = row.querySelector('.sno-cell') || row.firstElementChild;
            if (cell) cell.textContent = index + 1;
        });

        // Also update cumulative whenever serial numbers (order) change
        if (typeof updateCumulativePL === 'function') {
            updateCumulativePL();
        }
    }

    // Update Cumulative P/L Column
    function updateCumulativePL() {
        const rows = tableBody.querySelectorAll('tr');
        let runningTotal = 0;

        rows.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // P/L is inputs[4] (5th input: Date, Owner, Type, Exit, P/L)
            if (inputs.length >= 5) {
                const plVal = parseFloat(inputs[4].value);
                const cumCell = tr.querySelector('.cum-cell');

                if (cumCell) {
                    if (!isNaN(plVal)) {
                        runningTotal += plVal;
                        cumCell.textContent = runningTotal.toFixed(2);
                        cumCell.className = 'cum-cell ' + (runningTotal >= 0 ? 'positive' : 'negative');
                        // Add color style directly for specificity
                        cumCell.style.color = runningTotal >= 0 ? 'var(--success-color)' : 'var(--danger-color)';
                    } else {
                        cumCell.textContent = '--';
                        cumCell.style.color = 'var(--text-secondary)';
                    }
                }
            }
        });
    }

    // Helper to generate consistent colors
    function getColorForString(str) {
        if (!str) return 'var(--text-primary)';
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            hash = str.charCodeAt(i) + ((hash << 5) - hash);
        }
        // Generate HSL: Hue derived from hash, Saturation 70%, Lightness 60% (readable on dark)
        const h = Math.abs(hash % 360);
        return `hsl(${h}, 70%, 60%)`;
    }

    function createSmartCell(sourceKey, placeholder, initialValue = '') {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input smart-input';
        input.dataset.source = sourceKey;
        input.placeholder = placeholder;
        if (initialValue) {
            input.value = initialValue;
            if (sourceKey === 'owners') {
                input.style.color = getColorForString(initialValue);
                input.style.fontWeight = '500';
            }
        }

        addSaveTrigger(input);

        // Event Listeners
        input.addEventListener('focus', () => showDropdown(input));

        // Re-open on click if closed (e.g. if user clicked away then clicked back)
        input.addEventListener('click', () => {
            if (globalDropdown.style.display === 'none') {
                showDropdown(input);
            }
        });

        input.addEventListener('input', (e) => {
            const val = input.value;

            // Dynamic Color for Owners
            if (sourceKey === 'owners') {
                input.style.color = getColorForString(val);
                input.style.fontWeight = '500';
            }

            showDropdown(input);

            // Inline Autocomplete
            // 1. Not deleting
            // 2. Cursor is at the end (don't mess up editing in middle)
            // 3. Value has length
            if (e.inputType !== 'deleteContentBackward' &&
                e.inputType !== 'deleteContentForward' &&
                val.length > 0 &&
                input.selectionStart === val.length) {

                const list = appData[sourceKey] || [];
                const match = list.find(item => item.toLowerCase().startsWith(val.toLowerCase()));

                // Only autocomplete if we found a match AND it adds characters
                if (match && match.length > val.length) {
                    const originalLen = val.length;
                    // Preserve casing of what user typed? No, usually autocomplete suggests the stored case
                    // But we must be careful. match is the full string.
                    input.value = match;
                    input.setSelectionRange(originalLen, match.length);
                }
            }
        });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                // Autocomplete top match
                const firstItem = globalDropdown.querySelector('.dropdown-item .owner-text');
                if (firstItem && globalDropdown.style.display !== 'none') {
                    input.value = firstItem.textContent;
                    if (sourceKey === 'owners') {
                        input.style.color = getColorForString(input.value);
                    }
                }
                if (input.value) addItemIfNew(sourceKey, input.value);
                hideDropdown();

                // Move focus
                const nextInput = td.nextElementSibling ? td.nextElementSibling.querySelector('input, select') : null;
                if (nextInput) nextInput.focus();
            }
        });

        input.addEventListener('blur', () => {
            hideTimeout = setTimeout(() => {
                hideDropdown();
                if (input.value) addItemIfNew(sourceKey, input.value);
            }, 100);
        });

        td.appendChild(input);
        return td;
    }

    // --- Filter Logic ---
    const filterInputs = document.querySelectorAll('.filter-input');

    filterInputs.forEach(input => {
        input.addEventListener('input', filterTable);
    });

    function filterTable() {
        const trs = tableBody.querySelectorAll('tr');
        const filters = {};

        // Collect active filters
        filterInputs.forEach(input => {
            const colIndex = input.dataset.col; // 1-based index from HTML
            const value = input.value.toLowerCase().trim();
            if (value) {
                filters[colIndex] = value;
            }
        });

        trs.forEach(tr => {
            let isVisible = true;
            const cells = tr.querySelectorAll('td');

            // Iterate over active filters
            for (const [colIndex, filterValue] of Object.entries(filters)) {
                // Adjust index: Data columns start at index 1 (Date) in cells array if S.No is 0
                // My HTML filter inputs map: data-col="1" -> Date. 
                // In `tr`, cells[0] is S.No, cells[1] is Date.
                // So cells[colIndex] matches perfectly if S.No is 0.

                const cell = cells[colIndex];
                if (!cell) continue;

                // Special handling for Cumulative (Col 6) which is textContent, not input
                if (parseInt(colIndex) === 6) {
                    const text = cell.textContent.trim().toLowerCase();
                    if (!text.includes(filterValue)) {
                        isVisible = false;
                        break;
                    }
                    continue;
                }

                const cellValue = getCellValue(cell).toLowerCase();

                // Simple includes matching
                if (!cellValue.includes(filterValue)) {
                    isVisible = false;
                    break;
                }
            }

            tr.style.display = isVisible ? '' : 'none';
        });
    }


    // --- Sorting Logic ---
    const headers = document.querySelectorAll('th.sortable');
    let currentSort = { col: null, dir: 'asc' };

    headers.forEach(th => {
        th.addEventListener('click', () => {
            const colIndex = parseInt(th.dataset.colIndex);
            sortTable(colIndex, th);
        });
    });

    function sortTable(colIndex, th) {
        const rowsArray = Array.from(tableBody.querySelectorAll('tr'));

        // Determine direction
        let dir = 'asc';
        if (currentSort.col === colIndex && currentSort.dir === 'asc') {
            dir = 'desc';
        }
        currentSort = { col: colIndex, dir: dir };

        // Update UI Icons
        headers.forEach(h => h.classList.remove('asc', 'desc'));
        th.classList.add(dir);

        rowsArray.sort((rowA, rowB) => {
            const cellA = rowA.children[colIndex];
            const cellB = rowB.children[colIndex];

            const valA = getCellValue(cellA);
            const valB = getCellValue(cellB);

            // Compare
            return compareValues(valA, valB, colIndex, dir);
        });

        // Re-append sorted rows
        tableBody.innerHTML = '';
        rowsArray.forEach(row => tableBody.appendChild(row));

        // Update Serial Numbers (but keep sorted order? No, usually S.No stays 1,2,3... relative to view)
        // If user wants to see original order text, they look at ID. 
        // Typically in excel sorting changes the row position but Row Number stays left.
        updateSerialNumbers();
    }

    function getCellValue(cell) {
        if (!cell) return '';
        const input = cell.querySelector('input, select');
        return input ? input.value : cell.textContent.trim();
    }

    function compareValues(a, b, colIndex, dir) {
        // Date Columns (index 1 & 4 usually)
        if (colIndex === 1 || colIndex === 4) {
            const dateA = parseDate(a);
            const dateB = parseDate(b);
            return dir === 'asc' ? dateA - dateB : dateB - dateA;
        }

        // Number Columns (P/L is index 5)
        // Check if string looks like number?
        const numA = parseFloat(a);
        const numB = parseFloat(b);

        if (!isNaN(numA) && !isNaN(numB) && colIndex === 5) {
            return dir === 'asc' ? numA - numB : numB - numA;
        }

        // String Default
        const valA = a.toLowerCase();
        const valB = b.toLowerCase();

        if (valA < valB) return dir === 'asc' ? -1 : 1;
        if (valA > valB) return dir === 'asc' ? 1 : -1;
        return 0;
    }

    function parseDate(dateStr) {
        // Expecting YYYY-MM-DD from input[type="date"]
        if (!dateStr) return -1;
        const d = new Date(dateStr);
        return d.getTime();
    }

    function createCell(type, placeholder = '') {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = type;
        input.className = 'cell-input';
        if (placeholder) {
            input.placeholder = placeholder;
        }
        td.appendChild(input);
        return td;
    }

    function updatePLColor(e) {
        const input = e.target;
        const value = parseFloat(input.value);
        const cell = input.parentElement;
        cell.classList.remove('positive', 'negative');
        if (!isNaN(value)) {
            if (value > 0) {
                cell.classList.add('positive');
            } else if (value < 0) {
                cell.classList.add('negative');
            }
        }
    }

    // --- Modal Add Logic ---
    const modalAddBtn = document.getElementById('modal-add-btn');
    const modalAddInput = document.getElementById('modal-add-input');

    function handleModalAdd() {
        const val = modalAddInput.value;
        if (val && managingSourceKey) {
            addItemIfNew(managingSourceKey, val);
            modalAddInput.value = '';
            // Refresh list
            if (typeof renderManageList === 'function') {
                renderManageList();
            }
        }
    }

    if (modalAddBtn) {
        modalAddBtn.addEventListener('click', handleModalAdd);
    }

    if (modalAddInput) {
        modalAddInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') handleModalAdd();
        });
        // Search/Filter capability
        modalAddInput.addEventListener('input', (e) => {
            if (typeof renderManageList === 'function') {
                renderManageList(e.target.value);
            }
        });
    }

    // --- Settings UI Logic ---
    const settingsModal = document.getElementById('settings-modal');
    const inputGasUrl = document.getElementById('setting-gas-url');
    const checkAutoSync = document.getElementById('setting-auto-sync');
    const settingsBtn = document.getElementById('btn-settings');
    const closeSettingsBtn = document.getElementById('close-settings-btn');
    const btnSaveSettings = document.getElementById('btn-save-settings');
    const btnTestSync = document.getElementById('btn-test-sync');

    if (settingsBtn) {
        settingsBtn.addEventListener('click', () => {
            inputGasUrl.value = settingsGASUrl;
            checkAutoSync.checked = settingsAutoSync;
            settingsModal.classList.remove('hidden');
        });
    }

    if (closeSettingsBtn) {
        closeSettingsBtn.addEventListener('click', () => {
            settingsModal.classList.add('hidden');
        });
    }

    if (settingsModal) {
        settingsModal.addEventListener('click', (e) => {
            if (e.target === settingsModal) settingsModal.classList.add('hidden');
        });
    }

    if (btnSaveSettings) {
        btnSaveSettings.addEventListener('click', () => {
            const url = inputGasUrl.value.trim();
            const auto = checkAutoSync.checked;

            settingsGASUrl = url;
            settingsAutoSync = auto;

            localStorage.setItem('gas_url', url);
            localStorage.setItem('auto_sync', auto);

            btnSaveSettings.textContent = 'Saved!';
            setTimeout(() => btnSaveSettings.textContent = 'Save & Sync', 1000);

            const rows = JSON.parse(localStorage.getItem('pl_report_rows') || '[]');
            syncToGoogleSheets(rows);
        });
    }

    if (btnTestSync) {
        btnTestSync.addEventListener('click', () => {
            const rows = JSON.parse(localStorage.getItem('pl_report_rows') || '[]');
            if (rows.length === 0) {
                alert('Table is empty. Add data to test sync.');
                return;
            }
            const oldUrl = settingsGASUrl;
            settingsGASUrl = inputGasUrl.value.trim();

            if (!settingsGASUrl) {
                alert('Please enter a valid URL');
                settingsGASUrl = oldUrl;
                return;
            }

            syncToGoogleSheets(rows);
            settingsGASUrl = oldUrl;
        });
    }

    // --- Table Enhancements Logic (Recycle Bin, Undo, Confirm) ---
    let pendingDeleteTr = null;
    let recycleBin = JSON.parse(localStorage.getItem('pl_report_recycle_bin') || '[]');
    let undoTimeout = null;
    let lastDeletedData = null; // For Undo

    // Modals & UI
    const deleteConfirmModal = document.getElementById('delete-confirm-modal');
    const btnConfirmDelete = document.getElementById('btn-confirm-delete');
    const btnCancelDelete = document.getElementById('btn-cancel-delete');

    const recycleBinModal = document.getElementById('recycle-bin-modal');
    const btnRecycleBin = document.getElementById('btn-recycle-bin');
    const closeRecycleBtn = document.getElementById('close-recycle-btn');
    const btnEmptyBin = document.getElementById('btn-empty-bin');
    const recycleBinBody = document.getElementById('recycle-bin-body');
    const recycleEmptyState = document.getElementById('recycle-empty-state');

    const undoToast = document.getElementById('undo-toast');
    const btnUndo = document.getElementById('btn-undo');

    // Functions
    window.requestDelete = function (tr) {
        pendingDeleteTr = tr;
        deleteConfirmModal.classList.remove('hidden');
    };

    if (btnConfirmDelete) {
        btnConfirmDelete.addEventListener('click', () => {
            if (pendingDeleteTr) {
                moveToRecycleBin(pendingDeleteTr);
                pendingDeleteTr.remove();
                updateSerialNumbers();
                saveTableRows();
                deleteConfirmModal.classList.add('hidden');
                pendingDeleteTr = null;

                showUndoToast();
            }
        });
    }

    if (btnCancelDelete) {
        btnCancelDelete.addEventListener('click', () => {
            deleteConfirmModal.classList.add('hidden');
            pendingDeleteTr = null;
        });
    }

    function moveToRecycleBin(tr) {
        const inputs = tr.querySelectorAll('input, select');
        // Extract data
        const rowData = {
            date: inputs[0].value,
            owner: inputs[1].value,
            type: inputs[2].value,
            exitDate: inputs[3].value,
            pl: inputs[4].value,
            remark: inputs[5].value,
            deletedAt: new Date().toISOString()
        };

        lastDeletedData = rowData; // For immediate undo
        recycleBin.unshift(rowData); // Add to top
        localStorage.setItem('pl_report_recycle_bin', JSON.stringify(recycleBin));
    }

    function showUndoToast() {
        if (!undoToast) return;
        undoToast.style.transform = 'translateX(-50%) translateY(0)'; // Show
    }

    const btnCloseToast = document.getElementById('btn-close-toast');
    if (btnCloseToast) {
        btnCloseToast.addEventListener('click', () => {
            if (undoToast) undoToast.style.transform = 'translateX(-50%) translateY(100px)';
        });
    }

    if (btnUndo) {
        btnUndo.addEventListener('click', () => {
            if (lastDeletedData) {
                // Restore
                const restoredRow = addRow(lastDeletedData);
                saveTableRows();

                // Remove from recycle bin (since we just undid the add)
                recycleBin.shift();
                localStorage.setItem('pl_report_recycle_bin', JSON.stringify(recycleBin));

                if (undoToast) undoToast.style.transform = 'translateX(-50%) translateY(100px)'; // Hide
            }
        });
    }

    // Recycle Bin UI
    if (btnRecycleBin) {
        btnRecycleBin.addEventListener('click', () => {
            renderRecycleBin();
            recycleBinModal.classList.remove('hidden');
        });
    }

    if (closeRecycleBtn) {
        closeRecycleBtn.addEventListener('click', () => recycleBinModal.classList.add('hidden'));
    }

    if (btnEmptyBin) {
        btnEmptyBin.addEventListener('click', () => {
            if (confirm('Permanently delete all items in Recycle Bin?')) {
                recycleBin = [];
                localStorage.setItem('pl_report_recycle_bin', JSON.stringify(recycleBin));
                renderRecycleBin();
            }
        });
    }

    function renderRecycleBin() {
        if (!recycleBinBody) return;
        recycleBinBody.innerHTML = '';
        if (recycleBin.length === 0) {
            recycleEmptyState.style.display = 'block';
            return;
        }
        recycleEmptyState.style.display = 'none';

        recycleBin.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${item.date}</td>
                <td style="padding: 12px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${item.owner}</td>
                <td style="padding: 12px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${item.type}</td>
                <td style="padding: 12px; border-bottom: 1px solid var(--border-color); text-align: right; font-weight: 500; color: ${parseFloat(item.pl) >= 0 ? 'var(--success-color)' : 'var(--danger-color)'}">${item.pl}</td>
                <td style="padding: 12px; border-bottom: 1px solid var(--border-color); text-align: right;">
                    <button class="btn-restore" data-index="${index}" style="margin-right: 8px; background: none; border: none; color: var(--accent-color); cursor: pointer;">Restore</button>
                    <button class="btn-forever" data-index="${index}" style="background: none; border: none; color: var(--danger-color); cursor: pointer;">Delete</button>
                </td>
            `;
            recycleBinBody.appendChild(tr);
        });

        // Event Delegation for buttons
        recycleBinBody.querySelectorAll('.btn-restore').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.target.dataset.index;
                const item = recycleBin[idx];
                addRow(item);
                saveTableRows();

                recycleBin.splice(idx, 1);
                localStorage.setItem('pl_report_recycle_bin', JSON.stringify(recycleBin));
                renderRecycleBin();
            });
        });

        recycleBinBody.querySelectorAll('.btn-forever').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.target.dataset.index;
                if (confirm('Delete this item forever?')) {
                    recycleBin.splice(idx, 1);
                    localStorage.setItem('pl_report_recycle_bin', JSON.stringify(recycleBin));
                    renderRecycleBin();
                }
            });
        });
    }

    // --- Row Edit Timestamp Logic ---
    function updateRowTimestamp(input) {
        const tr = input.closest('tr');
        if (!tr) return;

        const now = new Date();
        const timeStr = now.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit' }) + ' ' +
            now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });

        let colName = '';
        const td = input.closest('td');
        if (td) {
            const index = td.cellIndex;
            if (index === 1) colName = 'Date';
            else if (index === 2) colName = 'Owner';
            else if (index === 3) colName = 'Type';
            else if (index === 4) colName = 'Exit';
            else if (index === 5) colName = 'P/L';
            else if (index === 7) colName = 'Rem.';
        }

        const msg = colName ? `${colName}: ${timeStr}` : `Edited: ${timeStr}`;

        tr.dataset.lastEdited = now.toISOString();
        tr.dataset.lastEditedMsg = msg;
        renderEditTimestamp(tr, msg);
    }

    function renderEditTimestamp(tr, text) {
        let editInfo = tr.querySelector('.edit-info');
        if (!editInfo) {
            const actionsTd = tr.querySelector('.actions-col');
            if (actionsTd) {
                editInfo = document.createElement('span');
                editInfo.className = 'edit-info';
                actionsTd.appendChild(editInfo);
            }
        }
        if (editInfo) editInfo.textContent = text.startsWith('Edited') ? text : 'Edited: ' + text;
    }

    function addSaveAndEditTrigger(input) {
        input.addEventListener('input', (e) => {
            updateRowTimestamp(e.target);
            saveTableRows();
        });

        input.addEventListener('change', (e) => {
            updateRowTimestamp(e.target);
            saveTableRows();
        });
    }

    // --- Drill Down Logic ---
    const detailsModal = document.getElementById('details-modal');
    const detailsModalTitle = document.getElementById('details-modal-title');
    const closeDetailsModal = document.getElementById('close-details-modal');
    const detailsTableBody = document.getElementById('details-table-body');

    if (closeDetailsModal) {
        closeDetailsModal.addEventListener('click', () => {
            if (detailsModal) detailsModal.classList.add('hidden');
        });
    }
    if (detailsModal) {
        detailsModal.addEventListener('click', (e) => {
            if (e.target === detailsModal) detailsModal.classList.add('hidden');
        });
    }

    function showDrillDown(filterType, filterValue) {
        if (!detailsModal) return;
        if (detailsModalTitle) detailsModalTitle.textContent = `${filterType}: ${filterValue} - Trades`;
        if (detailsTableBody) detailsTableBody.innerHTML = '';

        const trs = document.getElementById('table-body').querySelectorAll('tr');
        let count = 0;

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length < 5) return;

            const dateVal = inputs[0].value || '--';
            const ownerVal = inputs[1].value.trim();
            const typeVal = inputs[2].value.trim();
            const exitVal = inputs[3].value || '--';
            const plVal = parseFloat(inputs[4].value);

            let match = false;
            if (filterType === 'Owner' && ownerVal === filterValue) match = true;
            if (filterType === 'Type' && typeVal === filterValue) match = true;

            if (match) {
                const row = document.createElement('tr');
                const plClass = !isNaN(plVal) ? (plVal >= 0 ? 'positive' : 'negative') : '';
                const plDisplay = !isNaN(plVal) ? plVal.toFixed(2) : '0.00';

                row.innerHTML = `
                    <td style="padding: 10px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${dateVal}</td>
                    <td style="padding: 10px; border-bottom: 1px solid var(--border-color); color: ${getColorForString(ownerVal)}; font-weight: 500;">${ownerVal}</td>
                    <td style="padding: 10px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${typeVal}</td>
                    <td style="padding: 10px; border-bottom: 1px solid var(--border-color); color: var(--text-primary);">${exitVal}</td>
                    <td style="padding: 10px; text-align: right; border-bottom: 1px solid var(--border-color); font-family: 'JetBrains Mono', monospace;" class="${plClass}">${plDisplay}</td>
                 `;
                if (detailsTableBody) detailsTableBody.appendChild(row);
                count++;
            }
        });

        if (count === 0 && detailsTableBody) {
            detailsTableBody.innerHTML = '<tr><td colspan="4" style="padding: 20px; text-align: center; color: var(--text-secondary);">No trades found</td></tr>';
        }

        detailsModal.classList.remove('hidden');
    }

});
