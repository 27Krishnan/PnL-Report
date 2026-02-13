document.addEventListener('DOMContentLoaded', () => {
    const tableBody = document.getElementById('table-body');
    const addRowBtn = document.getElementById('add-row-btn');

    // View Toggle Elements
    const btnViewTable = document.getElementById('btn-view-table');
    const btnViewDashboard = document.getElementById('btn-view-dashboard');
    const btnViewMonthly = document.getElementById('btn-view-monthly');
    const btnViewPortfolio = document.getElementById('btn-view-portfolio');
    const viewTable = document.getElementById('view-table');
    const viewDashboard = document.getElementById('view-dashboard');
    const viewMonthly = document.getElementById('view-monthly');
    const viewPortfolio = document.getElementById('view-portfolio');

    // Portfolio Elements
    const portfolioBody = document.getElementById('portfolio-body');
    const addPortfolioRowBtn = document.getElementById('add-portfolio-row-btn');
    const portfolioTotalNet = document.getElementById('portfolio-total-net');

    // Modal Elements
    const manageModal = document.getElementById('manage-modal');
    const manageModalTitle = manageModal.querySelector('h2');
    const closeModalBtn = document.getElementById('close-modal-btn');
    const managementList = document.getElementById('owner-management-list');

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
    let monthlyTrendChart = null; // Monthly Trend Chart Instance
    let monthlyDailyChart = null; // Monthly Daily Trend Chart Instance

    // Settings State
    let settingsGASUrl = localStorage.getItem('gas_url') || '';
    let settingsAutoSync = localStorage.getItem('auto_sync') === 'true';
    let syncTimeout = null;

    // --- UI Setup ---
    createGlobalDropdown();

    // --- View Toggle Logic ---
    function switchView(viewName) {
        // Reset all
        [viewTable, viewDashboard, viewMonthly, viewPortfolio].forEach(el => el.classList.add('hidden'));
        [btnViewTable, btnViewDashboard, btnViewMonthly, btnViewPortfolio].forEach(el => el.classList.remove('active'));

        try {
            if (viewName === 'table') {
                viewTable.classList.remove('hidden');
                btnViewTable.classList.add('active');
            } else if (viewName === 'dashboard') {
                viewDashboard.classList.remove('hidden');
                btnViewDashboard.classList.add('active');
                updateDashboard(); // Refresh stats when opening
            } else if (viewName === 'monthly') {
                viewMonthly.classList.remove('hidden');
                btnViewMonthly.classList.add('active');
                updateMonthlyReport(); // Refresh monthly stats when opening
            } else if (viewName === 'portfolio') {
                viewPortfolio.classList.remove('hidden');
                btnViewPortfolio.classList.add('active');
                loadPortfolioRows(); // Refresh portfolio when opening
            }
        } catch (err) {
            console.error(`Error switching to ${viewName}:`, err);
        }
    }

    btnViewTable.addEventListener('click', () => switchView('table'));
    btnViewDashboard.addEventListener('click', () => switchView('dashboard'));
    btnViewMonthly.addEventListener('click', () => switchView('monthly'));
    btnViewPortfolio.addEventListener('click', () => switchView('portfolio'));

    // Monthly View Chart Toggle
    const btnToggleMonthlyTrend = document.getElementById('btn-toggle-monthly-trend');
    const btnToggleDailyTrend = document.getElementById('btn-toggle-daily-trend');
    const monthlyTrendCanvas = document.getElementById('monthlyTrendChart');
    const monthlyDailyCanvas = document.getElementById('monthlyDailyChart');

    if (btnToggleMonthlyTrend && btnToggleDailyTrend) {
        btnToggleMonthlyTrend.addEventListener('click', () => {
            btnToggleMonthlyTrend.classList.add('active');
            btnToggleDailyTrend.classList.remove('active');
            monthlyTrendCanvas.classList.remove('hidden');
            monthlyDailyCanvas.classList.add('hidden');
        });

        btnToggleDailyTrend.addEventListener('click', () => {
            btnToggleDailyTrend.classList.add('active');
            btnToggleMonthlyTrend.classList.remove('active');
            monthlyDailyCanvas.classList.remove('hidden');
            monthlyTrendCanvas.classList.add('hidden');
            updateMonthlyDailyChart(); // Refresh daily data for monthly view
        });
    }

    // Include Current Month Toggle
    const toggleIncludeCurrent = document.getElementById('toggle-include-current');
    if (toggleIncludeCurrent) {
        toggleIncludeCurrent.addEventListener('change', () => {
            updateMonthlyReport();
        });
    }

    // Monthly Overall Filters
    const setupMonthlyFilters = () => {
        ['filter-monthly-owner', 'filter-monthly-type', 'filter-monthly-det-owner', 'filter-monthly-det-type'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('input', () => {
                    if (id.includes('owner') && !id.includes('det')) updateMonthlyOwnerReport();
                    else if (id.includes('type') && !id.includes('det')) updateMonthlyTypeReport();
                    else updateMonthlyDetailedReport();
                });
            }
        });
    };
    setupMonthlyFilters();

    // Minimize Past Months Toggle
    const toggleMinimizePast = document.getElementById('toggle-minimize-past');
    if (toggleMinimizePast) {
        // Load initial state
        const savedState = localStorage.getItem('minimize_past_months') === 'true';
        toggleMinimizePast.checked = savedState;

        toggleMinimizePast.addEventListener('change', () => {
            localStorage.setItem('minimize_past_months', toggleMinimizePast.checked);
            filterTable();
        });
    }

    // Calendar Range Selector
    const rangeBtns = document.querySelectorAll('.range-btn:not(#calendar-apply-custom):not(#calendar-clear-custom)');
    const customPicker = document.getElementById('calendar-custom-picker');
    const applyCustomBtn = document.getElementById('calendar-apply-custom');
    const clearCustomBtn = document.getElementById('calendar-clear-custom');

    rangeBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            rangeBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            const range = btn.getAttribute('data-range');
            if (range === 'custom') {
                customPicker.classList.remove('hidden');
                // Set defaults if empty
                const startInput = document.getElementById('calendar-start-date');
                const endInput = document.getElementById('calendar-end-date');
                if (!startInput.value || !endInput.value) {
                    const now = new Date();
                    const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 2, 1);
                    startInput.value = threeMonthsAgo.toISOString().split('T')[0];
                    endInput.value = now.toISOString().split('T')[0];
                }
                updateMonthlyCalendar();
            } else {
                customPicker.classList.add('hidden');
                updateMonthlyCalendar();
            }
        });
    });

    if (applyCustomBtn) {
        applyCustomBtn.addEventListener('click', () => {
            updateMonthlyCalendar();
        });
    }

    if (clearCustomBtn) {
        clearCustomBtn.addEventListener('click', () => {
            // Reset to default (Year)
            rangeBtns.forEach(b => b.classList.remove('active'));
            const yearBtn = document.querySelector('.range-btn[data-range="1y"]');
            if (yearBtn) yearBtn.classList.add('active');

            customPicker.classList.add('hidden');
            updateMonthlyCalendar();
        });
    }


    addRowBtn.addEventListener('click', () => {
        addRow();
    });

    if (addPortfolioRowBtn) {
        addPortfolioRowBtn.addEventListener('click', () => {
            addPortfolioRow();
        });
    }

    // --- Portfolio Logic ---
    function addPortfolioRow(data = null) {
        const tr = document.createElement('tr');

        // snoTd
        const snoTd = document.createElement('td');
        snoTd.className = 'sno-cell';
        tr.appendChild(snoTd);

        // Name
        const nameTd = createCell('text', 'Portfolio Name');
        if (data && data.name) nameTd.querySelector('input').value = data.name;
        tr.appendChild(nameTd);

        // Start Date
        const startDateTd = createCell('date');
        if (data && data.startDate) startDateTd.querySelector('input').value = data.startDate;
        else startDateTd.querySelector('input').value = new Date().toISOString().split('T')[0];
        tr.appendChild(startDateTd);

        // End Date
        const endDateTd = createCell('date');
        if (data && data.endDate) endDateTd.querySelector('input').value = data.endDate;
        tr.appendChild(endDateTd);

        // Fund
        const fundTd = createCell('number', '0.00');
        if (data && data.fund) fundTd.querySelector('input').value = data.fund;
        tr.appendChild(fundTd);

        // Handling Charges
        const chargesTd = createCell('number', '0.00');
        if (data && data.charges) chargesTd.querySelector('input').value = data.charges;
        tr.appendChild(chargesTd);

        // Monthly Profit
        const profitTd = createCell('number', '0.00');
        if (data && data.profit) profitTd.querySelector('input').value = data.profit;
        tr.appendChild(profitTd);

        // Profit Sharing
        const sharingTd = createCell('number', '0.00');
        if (data && data.sharing) sharingTd.querySelector('input').value = data.sharing;
        tr.appendChild(sharingTd);

        // Net Profit (Read Only)
        const netTd = document.createElement('td');
        netTd.className = 'net-cell';
        netTd.style.fontWeight = '600';
        netTd.textContent = '0.00';
        tr.appendChild(netTd);

        // Remark
        const remarkTd = createCell('text', 'Remarks');
        if (data && data.remark) remarkTd.querySelector('input').value = data.remark;
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
            if (confirm('Delete this portfolio entry?')) {
                tr.remove();
                savePortfolioRows();
                updatePortfolioSerialNumbers();
                updatePortfolioTotal(); // Fix: Update total when row is deleted
            }
        };
        actionsTd.appendChild(deleteBtn);
        tr.appendChild(actionsTd);

        // Listeners for calculation
        const inputs = tr.querySelectorAll('input');
        inputs.forEach(input => {
            input.addEventListener('input', () => {
                updatePortfolioRowCalculation(tr);
                savePortfolioRows();
            });
        });

        portfolioBody.appendChild(tr);
        updatePortfolioRowCalculation(tr);
        updatePortfolioSerialNumbers();

        if (!data) {
            savePortfolioRows();
            startDateTd.querySelector('input').focus({ preventScroll: true });
        }
    }

    function updatePortfolioRowCalculation(tr) {
        const inputs = tr.querySelectorAll('input');
        const fund = parseFloat(inputs[3].value) || 0;
        const charges = parseFloat(inputs[4].value) || 0;
        const profit = parseFloat(inputs[5].value) || 0;
        const sharing = parseFloat(inputs[6].value) || 0;

        const netProfit = profit - charges - sharing;
        const netTd = tr.querySelector('.net-cell');
        if (netTd) {
            netTd.textContent = netProfit.toFixed(2);
            netTd.className = 'net-cell ' + (netProfit >= 0 ? 'positive' : 'negative');
            netTd.style.color = netProfit >= 0 ? 'var(--success-color)' : 'var(--danger-color)';
        }
        updatePortfolioTotal();
    }

    function updatePortfolioTotal() {
        let totalNet = 0;
        const rows = portfolioBody.querySelectorAll('tr');
        rows.forEach(tr => {
            const netTd = tr.querySelector('.net-cell');
            if (netTd) {
                totalNet += parseFloat(netTd.textContent) || 0;
            }
        });
        if (portfolioTotalNet) {
            portfolioTotalNet.textContent = totalNet.toFixed(2);
            portfolioTotalNet.className = 'value ' + (totalNet >= 0 ? 'positive' : 'negative');
        }
    }

    function updatePortfolioSerialNumbers() {
        const rows = portfolioBody.querySelectorAll('tr');
        rows.forEach((row, i) => {
            const cell = row.querySelector('.sno-cell');
            if (cell) cell.textContent = i + 1;
        });
    }

    function savePortfolioRows() {
        const rows = portfolioBody.querySelectorAll('tr');
        const data = Array.from(rows).map(tr => {
            const inputs = tr.querySelectorAll('input');
            return {
                name: inputs[0].value,
                startDate: inputs[1].value,
                endDate: inputs[2].value,
                fund: inputs[3].value,
                charges: inputs[4].value,
                profit: inputs[5].value,
                sharing: inputs[6].value,
                remark: inputs[7].value
            };
        });
        localStorage.setItem('portfolio_rows', JSON.stringify(data));
    }

    function loadPortfolioRows() {
        const saved = localStorage.getItem('portfolio_rows');
        if (saved) {
            portfolioBody.innerHTML = '';
            const data = JSON.parse(saved);
            data.forEach(item => addPortfolioRow(item));
        } else if (portfolioBody.children.length === 0) {
            addPortfolioRow();
        }
    }

    // --- Dashboard Logic ---
    function updateDashboard() {
        const trs = tableBody.querySelectorAll('tr');
        let totalPL = 0;
        let totalTrades = 0;
        let wins = 0;
        let grossWin = 0;
        let grossLoss = 0;

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        const allTradesCurrentMonth = [];

        trs.forEach(tr => {
            try {
                const inputs = tr.querySelectorAll('input');
                if (inputs.length > 4) {
                    const exitDateStr = inputs[3].value;
                    const plVal = parseFloat(inputs[4].value);

                    if (exitDateStr && !isNaN(plVal)) {
                        const exitDate = parseRobustDate(exitDateStr);
                        if (exitDate && exitDate.getFullYear() === currentYear && exitDate.getMonth() === currentMonth) {
                            totalTrades++;
                            totalPL += plVal;

                            // For advanced stats
                            allTradesCurrentMonth.push({ date: exitDate, pl: plVal });

                            if (plVal > 0) {
                                wins++;
                                grossWin += plVal;
                            } else if (plVal < 0) {
                                grossLoss += Math.abs(plVal);
                            }
                        }
                    }
                }
            } catch (e) {
                console.error("Error processing row in updateDashboard:", e);
            }
        });

        // Advanced Metrics Calculation
        allTradesCurrentMonth.sort((a, b) => a.date - b.date);
        let maxDD = 0, peak = 0, current = 0;
        allTradesCurrentMonth.forEach(t => {
            current += t.pl;
            if (current > peak) peak = current;
            const dd = peak - current;
            if (dd > maxDD) maxDD = dd;
        });

        const winRate = totalTrades > 0 ? wins / totalTrades : 0;
        const lossRate = totalTrades > 0 ? (totalTrades - wins) / totalTrades : 0;
        const avgWin = wins > 0 ? grossWin / wins : 0;
        const avgLoss = (totalTrades - wins) > 0 ? grossLoss / (totalTrades - wins) : 0;
        const expectancy = (winRate * avgWin) - (lossRate * avgLoss);
        const rrRatio = avgLoss > 0 ? avgWin / avgLoss : (avgWin > 0 ? 100 : 0);
        const recoveryFactor = maxDD > 0 ? totalPL / maxDD : (totalPL > 0 ? 100 : 0);

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

        const elExpectancy = document.getElementById('dash-expectancy');
        const elRecoveryFactor = document.getElementById('dash-recovery-factor');
        const elRR = document.getElementById('dash-rr-ratio');

        if (elExpectancy) {
            elExpectancy.textContent = expectancy.toFixed(2);
            elExpectancy.className = 'metric-value ' + (expectancy >= 0 ? 'positive' : 'negative');
        }

        if (elRecoveryFactor) {
            elRecoveryFactor.textContent = recoveryFactor.toFixed(2);
            elRecoveryFactor.className = 'metric-value ' + (recoveryFactor >= 1 ? 'positive' : 'negative');
        }

        if (elRR) {
            elRR.textContent = rrRatio.toFixed(2);
            elRR.className = 'metric-value ' + (rrRatio >= 1 ? 'positive' : 'negative');
        }

        // Update Report Table
        try { if (typeof updateOwnerReport === 'function') updateOwnerReport(); } catch (e) { console.error(e); }
        try { if (typeof updateTypeReport === 'function') updateTypeReport(); } catch (e) { console.error(e); }
        try { if (typeof updateDailyChart === 'function') updateDailyChart(); } catch (e) { console.error(e); }
        try { if (typeof updateCumulativePL === 'function') updateCumulativePL(); } catch (e) { console.error(e); }
        try { if (typeof updateDetailedReport === 'function') updateDetailedReport(); } catch (e) { console.error(e); }

        // Auto-update monthly view if visible
        if (typeof updateMonthlyReport === 'function' && viewMonthly && !viewMonthly.classList.contains('hidden')) {
            updateMonthlyReport();
        }
    }

    function updateMonthlyReport() {
        if (!tableBody) return;
        const trs = tableBody.querySelectorAll('tr');
        const monthlyStats = {};

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // Index 3 is Exit Date, 4 is P/L
            if (inputs.length > 4) {
                const exitDate = inputs[3].value;
                const plVal = parseFloat(inputs[4].value);

                if (exitDate && !isNaN(plVal)) {
                    const dateObj = parseRobustDate(exitDate);
                    if (!dateObj) return;

                    const year = dateObj.getFullYear();
                    const month = dateObj.getMonth() + 1;
                    const monthKey = `${year}-${String(month).padStart(2, '0')}`;
                    const monthLabel = dateObj.toLocaleString('default', { month: 'long', year: 'numeric' });

                    if (!monthlyStats[monthKey]) {
                        monthlyStats[monthKey] = {
                            label: monthLabel,
                            pl: 0,
                            trades: 0,
                            wins: 0,
                            grossWin: 0,
                            grossLoss: 0,
                            days: {},
                            owners: {},
                            types: {}
                        };
                    }

                    monthlyStats[monthKey].pl += plVal;
                    monthlyStats[monthKey].trades++;
                    if (plVal > 0) {
                        monthlyStats[monthKey].wins++;
                        monthlyStats[monthKey].grossWin += plVal;
                    } else if (plVal < 0) {
                        monthlyStats[monthKey].grossLoss += Math.abs(plVal);
                    }

                    // Detailed aggregates for drill-down
                    const dayKey = exitDate;
                    monthlyStats[monthKey].days[dayKey] = (monthlyStats[monthKey].days[dayKey] || 0) + plVal;

                    const owner = inputs[1].value.trim() || 'Unknown';
                    if (!monthlyStats[monthKey].owners[owner]) monthlyStats[monthKey].owners[owner] = { pl: 0, trades: 0 };
                    monthlyStats[monthKey].owners[owner].pl += plVal;
                    monthlyStats[monthKey].owners[owner].trades++;

                    const type = inputs[2].value.trim() || 'Unknown';
                    if (!monthlyStats[monthKey].types[type]) monthlyStats[monthKey].types[type] = { pl: 0, trades: 0 };
                    monthlyStats[monthKey].types[type].pl += plVal;
                    monthlyStats[monthKey].types[type].trades++;
                }
            }
        });

        const sortedMonthKeys = Object.keys(monthlyStats).sort().reverse();
        const tbody = document.getElementById('monthly-pl-body');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (sortedMonthKeys.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; padding: 20px; color: var(--text-secondary);">No monthly data available</td></tr>';
            return;
        }

        sortedMonthKeys.forEach(key => {
            const stat = monthlyStats[key];
            const winRate = stat.trades > 0 ? (stat.wins / stat.trades * 100).toFixed(1) : '0.0';
            const profitFactor = stat.grossLoss === 0 ? (stat.grossWin > 0 ? 'MAX' : '0.00') : (stat.grossWin / stat.grossLoss).toFixed(2);

            const tr = document.createElement('tr');
            tr.className = 'month-row';
            tr.title = 'Click to view breakdown (Days, Owners, Types)';
            tr.onclick = () => toggleMonthlyDetails(key, tr, stat);

            tr.innerHTML = `
            <td class="col-month"><span class="chevron-icon">▶</span> ${stat.label}</td>
            <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
            <td class="text-right col-trades">${stat.trades}</td>
            <td class="text-right col-win">${winRate}%</td>
            <td class="text-right col-pf">${profitFactor}</td>
        `;
            tbody.appendChild(tr);
        });

        // --- Update Monthly Dashboard Stats ---
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        // Filter keys for Summary Cards and Trend Chart based on toggle
        const reportMonthKeys = sortedMonthKeys.filter(key => {
            if (!includeCurrent && key === currentMonthKey) return false;
            return true;
        });

        const totalNetPL = reportMonthKeys.reduce((sum, key) => sum + monthlyStats[key].pl, 0);
        const avgPL = reportMonthKeys.length > 0 ? totalNetPL / reportMonthKeys.length : 0;

        let bestMonth = { key: null, pl: -Infinity };
        let lowestMonth = { key: null, pl: Infinity };

        reportMonthKeys.forEach(key => {
            const pl = monthlyStats[key].pl;
            if (pl > bestMonth.pl) bestMonth = { key, pl, label: monthlyStats[key].label };
            if (pl < lowestMonth.pl) lowestMonth = { key, pl, label: monthlyStats[key].label };
        });

        const elTotal = document.getElementById('monthly-total-pl');
        const elAvg = document.getElementById('monthly-avg-pl');
        const elBestVal = document.getElementById('monthly-best-val');
        const elBestLabel = document.getElementById('monthly-best-label');
        const elLowestVal = document.getElementById('monthly-lowest-val');
        const elLowestLabel = document.getElementById('monthly-lowest-label');

        if (elTotal) {
            elTotal.textContent = totalNetPL.toFixed(2);
            elTotal.className = 'metric-value ' + (totalNetPL >= 0 ? 'positive' : 'negative');
        }
        if (elAvg) {
            elAvg.textContent = avgPL.toFixed(2);
            elAvg.className = 'metric-value ' + (avgPL >= 0 ? 'positive' : 'negative');
        }
        if (elBestVal && bestMonth.key) {
            elBestVal.textContent = bestMonth.pl.toFixed(2);
            elBestVal.className = 'metric-value ' + (bestMonth.pl >= 0 ? 'positive' : 'negative');
            elBestLabel.textContent = bestMonth.label;
        } else if (elBestVal) {
            elBestVal.textContent = '0.00';
            elBestLabel.textContent = '--';
            elBestVal.className = 'metric-value';
        }

        if (elLowestVal && lowestMonth.key) {
            elLowestVal.textContent = lowestMonth.pl.toFixed(2);
            elLowestVal.className = 'metric-value ' + (lowestMonth.pl >= 0 ? 'positive' : 'negative');
            elLowestLabel.textContent = lowestMonth.label;
        } else if (elLowestVal) {
            elLowestVal.textContent = '0.00';
            elLowestLabel.textContent = '--';
            elLowestVal.className = 'metric-value';
        }

        // --- Update Advanced Trade Metrics ---
        const allTrades = [];
        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length > 4) {
                const entryDate = inputs[0].value;
                const owner = inputs[1].value;
                const type = inputs[2].value;
                const exitDateStr = inputs[3].value;
                const pl = parseFloat(inputs[4].value);
                if (exitDateStr && !isNaN(pl)) {
                    const dateObj = parseRobustDate(exitDateStr);
                    if (dateObj) {
                        const monthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                        if (includeCurrent || monthKey !== currentMonthKey) {
                            allTrades.push({ date: dateObj, entryDate, owner, type, pl });
                        }
                    }
                }
            }
        });

        // Sort by date ascending for streak/drawdown
        allTrades.sort((a, b) => a.date - b.date);

        let grossWinCombined = 0, grossLossCombined = 0, winCount = 0, lossCount = 0;
        let bigWinTrade = null, bigLossTrade = null;
        let bestStreak = 0, worstStreak = 0;
        let currentWinStreak = 0, currentLossStreak = 0;
        let bestStreakStart = null, bestStreakEnd = null;
        let worstStreakStart = null, worstStreakEnd = null;
        let tempWinStart = null, tempLossStart = null;

        let maxDrawdown = 0, peakPL = 0, currentPL = 0, drawdownDate = null;

        allTrades.forEach(t => {
            currentPL += t.pl;
            if (currentPL > peakPL) {
                peakPL = currentPL;
            }
            const dd = peakPL - currentPL;
            if (dd > maxDrawdown) {
                maxDrawdown = dd;
                drawdownDate = t.date.toISOString().split('T')[0];
            }

            if (t.pl > 0) {
                grossWinCombined += t.pl;
                winCount++;

                if (currentWinStreak === 0) tempWinStart = t.date.toISOString().split('T')[0];
                currentWinStreak++;
                currentLossStreak = 0;

                if (currentWinStreak > bestStreak) {
                    bestStreak = currentWinStreak;
                    bestStreakStart = tempWinStart;
                    bestStreakEnd = t.date.toISOString().split('T')[0];
                }
                if (!bigWinTrade || t.pl > bigWinTrade.pl) bigWinTrade = t;
            } else if (t.pl < 0) {
                const absPL = Math.abs(t.pl);
                grossLossCombined += absPL;
                lossCount++;

                if (currentLossStreak === 0) tempLossStart = t.date.toISOString().split('T')[0];
                currentLossStreak++;
                currentWinStreak = 0;

                if (currentLossStreak > worstStreak) {
                    worstStreak = currentLossStreak;
                    worstStreakStart = tempLossStart;
                    worstStreakEnd = t.date.toISOString().split('T')[0];
                }
                if (!bigLossTrade || t.pl < bigLossTrade.pl) bigLossTrade = t;
            } else {
                currentWinStreak = 0;
                currentLossStreak = 0;
            }
        });

        const totalTradesCount = allTrades.length;
        const avgTrade = totalTradesCount > 0 ? (totalNetPL / totalTradesCount) : 0;
        const avgWin = winCount > 0 ? (grossWinCombined / winCount) : 0;
        const avgLoss = lossCount > 0 ? (grossLossCombined / lossCount) : 0;
        const expectancy = totalTradesCount > 0 ? ((winCount / totalTradesCount) * avgWin) - ((lossCount / totalTradesCount) * avgLoss) : 0;
        const recoveryFactor = maxDrawdown > 0 ? totalNetPL / maxDrawdown : (totalNetPL > 0 ? 100 : 0);
        const rrRatio = avgLoss > 0 ? (avgWin / avgLoss) : (avgWin > 0 ? 100 : 0);

        // Update UI Elements
        const elAvgTrade = document.getElementById('monthly-avg-trade');
        const elExpectancy = document.getElementById('monthly-expectancy');
        const elRecoveryFactor = document.getElementById('monthly-recovery-factor');
        const elRR = document.getElementById('monthly-rr-ratio');
        const elDD = document.getElementById('monthly-max-drawdown');
        const elDDDate = document.getElementById('monthly-max-drawdown-date');

        const elBestStreak = document.getElementById('monthly-best-streak');
        const elBestStreakRange = document.getElementById('monthly-best-streak-range');

        const elWorstStreak = document.getElementById('monthly-worst-streak');
        const elWorstStreakRange = document.getElementById('monthly-worst-streak-range');

        const elBigWin = document.getElementById('monthly-biggest-win');
        const elBigWinDate = document.getElementById('monthly-biggest-win-date');
        const elBigWinOwner = document.getElementById('monthly-biggest-win-owner');
        const elBigWinType = document.getElementById('monthly-biggest-win-type');

        const elBigLoss = document.getElementById('monthly-biggest-loss');
        const elBigLossDate = document.getElementById('monthly-biggest-loss-date');
        const elBigLossOwner = document.getElementById('monthly-biggest-loss-owner');
        const elBigLossType = document.getElementById('monthly-biggest-loss-type');

        if (elAvgTrade) {
            elAvgTrade.textContent = avgTrade.toFixed(2);
            elAvgTrade.className = 'metric-value ' + (avgTrade >= 0 ? 'positive' : 'negative');
        }
        if (elExpectancy) {
            elExpectancy.textContent = expectancy.toFixed(2);
            elExpectancy.className = 'metric-value ' + (expectancy >= 0 ? 'positive' : 'negative');
        }
        if (elRecoveryFactor) {
            elRecoveryFactor.textContent = recoveryFactor.toFixed(2);
            elRecoveryFactor.className = 'metric-value ' + (recoveryFactor >= 1 ? 'positive' : 'negative');
        }
        if (elRR) {
            elRR.textContent = rrRatio.toFixed(2);
            elRR.className = 'metric-value ' + (rrRatio >= 1 ? 'positive' : 'negative');
        }

        if (elDD) {
            elDD.textContent = (-maxDrawdown).toFixed(2);
            elDDDate.textContent = drawdownDate || '--';
        }

        if (elBestStreak) {
            elBestStreak.textContent = bestStreak;
            elBestStreakRange.textContent = bestStreak > 0 ? `${bestStreakStart} to ${bestStreakEnd}` : '--';
        }

        if (elWorstStreak) {
            elWorstStreak.textContent = worstStreak;
            elWorstStreakRange.textContent = worstStreak > 0 ? `${worstStreakStart} to ${worstStreakEnd}` : '--';
        }

        if (elBigWin && bigWinTrade) {
            elBigWin.textContent = bigWinTrade.pl.toFixed(2);
            elBigWinDate.textContent = bigWinTrade.date.toISOString().split('T')[0];
            elBigWinOwner.textContent = bigWinTrade.owner || '--';
            elBigWinType.textContent = bigWinTrade.type || '--';
        } else if (elBigWin) {
            elBigWin.textContent = '0.00';
            elBigWinDate.textContent = '--';
            elBigWinOwner.textContent = '--';
            elBigWinType.textContent = '--';
        }

        if (elBigLoss && bigLossTrade) {
            elBigLoss.textContent = bigLossTrade.pl.toFixed(2);
            elBigLossDate.textContent = bigLossTrade.date.toISOString().split('T')[0];
            elBigLossOwner.textContent = bigLossTrade.owner || '--';
            elBigLossType.textContent = bigLossTrade.type || '--';
        } else if (elBigLoss) {
            elBigLoss.textContent = '0.00';
            elBigLossDate.textContent = '--';
            elBigLossOwner.textContent = '--';
            elBigLossType.textContent = '--';
        }

        // --- Update Monthly Trend Chart ---
        updateMonthlyTrendChart(monthlyStats, reportMonthKeys);
        updateMonthlyDailyChart();

        // --- Update Overall Performance Reports (Monthly View Version) ---
        updateMonthlyOwnerReport();
        updateMonthlyTypeReport();
        updateMonthlyDetailedReport();
        updateMonthlyCalendar();
    }

    function toggleMonthlyDetails(monthKey, rowEl, stats) {
        const nextRow = rowEl.nextElementSibling;
        const isExpanded = rowEl.classList.contains('expanded');

        if (isExpanded) {
            rowEl.classList.remove('expanded');
            if (nextRow && nextRow.classList.contains('monthly-details-row')) {
                nextRow.classList.add('hidden');
            }
        } else {
            rowEl.classList.add('expanded');
            let detailsRow;
            if (nextRow && nextRow.classList.contains('monthly-details-row')) {
                detailsRow = nextRow;
                detailsRow.classList.remove('hidden');
            } else {
                detailsRow = document.createElement('tr');
                detailsRow.className = 'monthly-details-row';

                // Sort Days descending and Owners/Types by P&L descending
                const sortedDays = Object.keys(stats.days).sort().reverse();
                const sortedOwners = Object.keys(stats.owners).sort((a, b) => stats.owners[b].pl - stats.owners[a].pl);
                const sortedTypes = Object.keys(stats.types).sort((a, b) => stats.types[b].pl - stats.types[a].pl);

                detailsRow.innerHTML = `
                    <td colspan="5">
                        <div class="monthly-details-container">
                            <div class="monthly-details-section">
                                <h4>Days Wise Breakdown</h4>
                                <table class="monthly-sub-table">
                                    <thead><tr><th>Date</th><th style="text-align: right;">P&L</th></tr></thead>
                                    <tbody>
                                        ${sortedDays.map(date => `
                                            <tr>
                                                <td>${date}</td>
                                                <td style="text-align: right;" class="pl-text ${stats.days[date] >= 0 ? 'positive' : 'negative'}">
                                                    ${stats.days[date].toFixed(2)}
                                                </td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                            <div class="monthly-details-section">
                                <h4>Owner Breakdown</h4>
                                <table class="monthly-sub-table">
                                    <thead><tr><th>Owner</th><th style="text-align: right;">P&L</th></tr></thead>
                                    <tbody>
                                        ${sortedOwners.map(owner => `
                                            <tr>
                                                <td style="color: ${getColorForString(owner)}; font-weight: 500;">${owner}</td>
                                                <td style="text-align: right;" class="pl-text ${stats.owners[owner].pl >= 0 ? 'positive' : 'negative'}">
                                                    ${stats.owners[owner].pl.toFixed(2)}
                                                </td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                            <div class="monthly-details-section">
                                <h4>Type Breakdown</h4>
                                <table class="monthly-sub-table">
                                    <thead><tr><th>Type</th><th style="text-align: right;">P&L</th></tr></thead>
                                    <tbody>
                                        ${sortedTypes.map(type => `
                                            <tr>
                                                <td>${type}</td>
                                                <td style="text-align: right;" class="pl-text ${stats.types[type].pl >= 0 ? 'positive' : 'negative'}">
                                                    ${stats.types[type].pl.toFixed(2)}
                                                </td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                                <div style="margin-top: 20px; text-align: center;">
                                    <button class="btn-primary" style="padding: 8px 16px; font-size: 0.8rem;" onclick="showDrillDown('Month', '${monthKey}')">
                                        View All Trades
                                    </button>
                                </div>
                            </div>
                        </div>
                    </td>
                `;
                rowEl.parentNode.insertBefore(detailsRow, nextRow);
            }
        }
    }

    function updateMonthlyOwnerReport() {
        const trs = tableBody.querySelectorAll('tr');
        const ownerStats = {};
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const dateStr = inputs[0].value; // Entry Date or Exit Date? Let's use Exit Date for consistency with Monthly aggregation
                const exitDateStr = inputs[3].value;
                const owner = inputs[1].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (owner && !isNaN(pl)) {
                    // Filter logic
                    if (!includeCurrent && exitDateStr) {
                        const dateObj = parseRobustDate(exitDateStr);
                        if (dateObj) {
                            const monthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                            if (monthKey === currentMonthKey) return;
                        }
                    }

                    if (!ownerStats[owner]) ownerStats[owner] = { pl: 0, trades: 0 };
                    ownerStats[owner].pl += pl;
                    ownerStats[owner].trades += 1;
                }
            }
        });
        const sorted = Object.keys(ownerStats).map(k => ({ name: k, ...ownerStats[k] })).sort((a, b) => b.pl - a.pl);
        const tbody = document.getElementById('monthly-overall-owner-body');
        const filterVal = (document.getElementById('filter-monthly-owner')?.value || '').toLowerCase();

        if (!tbody) return;
        tbody.innerHTML = '';

        sorted.filter(s => s.name.toLowerCase().includes(filterVal)).forEach((stat, i) => {
            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.onclick = () => showDrillDown('Owner', stat.name);
            tr.innerHTML = `
                <td class="col-rank">${i + 1}</td>
                <td style="color: ${getColorForString(stat.name)}; font-weight: 500;">${stat.name}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
                <td class="text-right col-trades">${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function updateMonthlyTypeReport() {
        const trs = tableBody.querySelectorAll('tr');
        const typeStats = {};
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (type && !isNaN(pl)) {
                    // Filter logic
                    if (!includeCurrent && exitDateStr) {
                        const dateObj = parseRobustDate(exitDateStr);
                        if (dateObj) {
                            const monthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                            if (monthKey === currentMonthKey) return;
                        }
                    }

                    if (!typeStats[type]) typeStats[type] = { pl: 0, trades: 0 };
                    typeStats[type].pl += pl;
                    typeStats[type].trades += 1;
                }
            }
        });
        const sorted = Object.keys(typeStats).map(k => ({ name: k, ...typeStats[k] })).sort((a, b) => b.pl - a.pl);
        const tbody = document.getElementById('monthly-overall-type-body');
        const filterVal = (document.getElementById('filter-monthly-type')?.value || '').toLowerCase();

        if (!tbody) return;
        tbody.innerHTML = '';

        sorted.filter(s => s.name.toLowerCase().includes(filterVal)).forEach((stat, i) => {
            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.onclick = () => showDrillDown('Type', stat.name);
            tr.innerHTML = `
                <td class="col-rank">${i + 1}</td>
                <td>${stat.name}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
                <td class="text-right col-trades">${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function updateMonthlyDetailedReport() {
        const trs = tableBody.querySelectorAll('tr');
        const stats = {};
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const owner = inputs[1].value.trim();
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (owner && type && !isNaN(pl)) {
                    // Filter logic
                    if (!includeCurrent && exitDateStr) {
                        const dateObj = parseRobustDate(exitDateStr);
                        if (dateObj) {
                            const monthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                            if (monthKey === currentMonthKey) return;
                        }
                    }

                    const key = `${owner}|${type}`;
                    if (!stats[key]) stats[key] = { owner, type, pl: 0, trades: 0 };
                    stats[key].pl += pl;
                    stats[key].trades += 1;
                }
            }
        });
        const sorted = Object.values(stats).sort((a, b) => b.pl - a.pl);
        const tbody = document.getElementById('monthly-overall-detailed-body');
        const filterOwner = (document.getElementById('filter-monthly-det-owner')?.value || '').toLowerCase();
        const filterType = (document.getElementById('filter-monthly-det-type')?.value || '').toLowerCase();

        if (!tbody) return;
        tbody.innerHTML = '';

        sorted.filter(s =>
            s.owner.toLowerCase().includes(filterOwner) &&
            s.type.toLowerCase().includes(filterType)
        ).forEach((stat, i) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="col-rank">${i + 1}</td>
                <td style="color: ${getColorForString(stat.owner)}; font-weight: 500;">${stat.owner}</td>
                <td>${stat.type}</td>
                <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
                <td class="text-right col-trades">${stat.trades}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function updateMonthlyTrendChart(monthlyStats, sortedMonthKeys) {
        if (!monthlyTrendChart) initChart();
        const chronKeys = [...sortedMonthKeys].sort(); // Chronological order for trend

        const labels = chronKeys.map(key => monthlyStats[key].label);
        const data = chronKeys.map(key => monthlyStats[key].pl);
        const colors = data.map(val => val >= 0 ? 'rgba(16, 185, 129, 0.7)' : 'rgba(239, 68, 68, 0.7)');

        // Store keys for drill down
        chronKeys.forEach((key, i) => {
            monthlyTrendChart.canvas.dataset[`monthKey_${i}`] = key;
        });

        monthlyTrendChart.data.labels = labels;
        monthlyTrendChart.data.datasets[0].data = data;
        monthlyTrendChart.data.datasets[0].backgroundColor = colors;
        monthlyTrendChart.data.datasets[0].borderColor = colors.map(c => c.replace('0.7', '1'));
        monthlyTrendChart.update();
    }

    // --- Owner P&L Report Logic ---
    function updateOwnerReport() {
        const trs = tableBody.querySelectorAll('tr');
        const ownerStats = {};

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            // inputs[1] is Owner 
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const owner = inputs[1].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (exitDateStr && owner && !isNaN(pl)) {
                    const exitDate = parseRobustDate(exitDateStr);
                    if (exitDate && exitDate.getFullYear() === currentYear && exitDate.getMonth() === currentMonth) {
                        if (!ownerStats[owner]) {
                            ownerStats[owner] = { pl: 0, trades: 0, wins: 0, grossWin: 0, grossLoss: 0 };
                        }
                        ownerStats[owner].pl += pl;
                        ownerStats[owner].trades += 1;
                        if (pl > 0) {
                            ownerStats[owner].wins += 1;
                            ownerStats[owner].grossWin += pl;
                        } else if (pl < 0) {
                            ownerStats[owner].grossLoss += Math.abs(pl);
                        }
                    }
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

            const winRate = stat.trades > 0 ? (stat.wins / stat.trades * 100).toFixed(1) : '0.0';
            const avgWin = stat.wins > 0 ? stat.grossWin / stat.wins : 0;
            const losses = stat.trades - stat.wins;
            const avgLoss = losses > 0 ? stat.grossLoss / losses : 0;
            const rr = avgLoss > 0 ? (avgWin / avgLoss).toFixed(2) : (avgWin > 0 ? 'MAX' : '0.00');

            // Enable Drill Down
            tr.style.cursor = 'pointer';
            tr.title = 'Click to view trades';
            const mKey = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;
            tr.onclick = () => showDrillDown('Owner', stat.name, mKey);

            tr.innerHTML = `
            <td class="col-rank">${index + 1}</td>
            <td style="color: ${color}; font-weight: 500;">${stat.name}</td>
            <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
            <td class="text-right col-trades">${stat.trades}</td>
            <td class="text-right col-win">${winRate}%</td>
            <td class="text-right col-rr">${rr}</td>
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

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (exitDateStr && type && !isNaN(pl)) {
                    const exitDate = parseRobustDate(exitDateStr);
                    if (exitDate && exitDate.getFullYear() === currentYear && exitDate.getMonth() === currentMonth) {
                        if (!typeStats[type]) {
                            typeStats[type] = { pl: 0, trades: 0, wins: 0, grossWin: 0, grossLoss: 0 };
                        }
                        typeStats[type].pl += pl;
                        typeStats[type].trades += 1;
                        if (pl > 0) {
                            typeStats[type].wins += 1;
                            typeStats[type].grossWin += pl;
                        } else if (pl < 0) {
                            typeStats[type].grossLoss += Math.abs(pl);
                        }
                    }
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

            const winRate = stat.trades > 0 ? (stat.wins / stat.trades * 100).toFixed(1) : '0.0';
            const avgWin = stat.wins > 0 ? stat.grossWin / stat.wins : 0;
            const losses = stat.trades - stat.wins;
            const avgLoss = losses > 0 ? stat.grossLoss / losses : 0;
            const rr = avgLoss > 0 ? (avgWin / avgLoss).toFixed(2) : (avgWin > 0 ? 'MAX' : '0.00');

            // Enable Drill Down
            tr.style.cursor = 'pointer';
            tr.title = 'Click to view trades';
            const mKey = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;
            tr.onclick = () => showDrillDown('Type', stat.name, mKey);

            tr.innerHTML = `
            <td class="col-rank">${index + 1}</td>
            <td>${stat.name}</td>
            <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
            <td class="text-right col-trades">${stat.trades}</td>
            <td class="text-right col-win">${winRate}%</td>
            <td class="text-right col-rr">${rr}</td>
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

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const owner = inputs[1].value.trim();
                const type = inputs[2].value.trim();
                const pl = parseFloat(inputs[4].value);

                if (exitDateStr && owner && type && !isNaN(pl)) {
                    const exitDate = parseRobustDate(exitDateStr);
                    if (exitDate && exitDate.getFullYear() === currentYear && exitDate.getMonth() === currentMonth) {
                        const key = `${owner}|${type}`;
                        if (!stats[key]) {
                            stats[key] = { owner, type, pl: 0, trades: 0 };
                        }
                        stats[key].pl += pl;
                        stats[key].trades += 1;
                    }
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

            // Enable Drill Down
            tr.style.cursor = 'pointer';
            tr.title = 'Click to view trades';
            const mKey = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;
            tr.onclick = () => showDrillDown('Owner|Type', `${stat.owner}|${stat.type}`, mKey);

            tr.innerHTML = `
            <td class="col-rank">${index + 1}</td>
            <td style="color: ${color}; font-weight: 500;">${stat.owner}</td>
            <td>${stat.type}</td>
            <td class="pl-text ${stat.pl >= 0 ? 'positive' : 'negative'} text-right col-pl">${stat.pl.toFixed(2)}</td>
            <td class="text-right col-trades">${stat.trades}</td>
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
                },
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const chart = elements[0].element.$context.chart;
                        const label = chart.data.labels[index];
                        const datasetLabel = chart.data.datasets[0].label;

                        if (datasetLabel === 'Daily P/L') {
                            showDrillDown('Date', label);
                        } else if (datasetLabel === 'Monthly P/L Trend') {
                            // Map chronological label back to monthKey (YYYY-MM)
                            const monthKey = chart.canvas.dataset[`monthKey_${index}`];
                            if (monthKey) showDrillDown('Month', monthKey);
                        } else if (datasetLabel === 'Net P/L') {
                            const now = new Date();
                            const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
                            if (chart === ownerChart) {
                                showDrillDown('Owner', label, currentMonthKey);
                            } else if (chart === typeChart) {
                                showDrillDown('Type', label, currentMonthKey);
                            }
                        }
                    }
                },
                onHover: (e, elements) => {
                    e.native.target.style.cursor = elements.length ? 'pointer' : 'default';
                }
            }
        };
    }

    function initChart() {
        if (typeof Chart === 'undefined') {
            console.warn("Chart.js dependency not found. Charts will be disabled.");
            return;
        }
        try {
            const elOwner = document.getElementById('ownerChart');
            const elType = document.getElementById('typeChart');
            const elDaily = document.getElementById('dailyChart');
            const elMonthly = document.getElementById('monthlyTrendChart');
            const elMonthlyDaily = document.getElementById('monthlyDailyChart');

            if (elOwner) {
                const ctxOwner = elOwner.getContext('2d');
                ownerChart = new Chart(ctxOwner, createChartConfig('Net P/L'));
            }
            if (elType) {
                const ctxType = elType.getContext('2d');
                typeChart = new Chart(ctxType, createChartConfig('Net P/L'));
            }
            if (elDaily) {
                const ctxDaily = elDaily.getContext('2d');
                dailyChart = new Chart(ctxDaily, createChartConfig('Daily P/L'));
            }
            if (elMonthly) {
                const ctxMonthly = elMonthly.getContext('2d');
                monthlyTrendChart = new Chart(ctxMonthly, createChartConfig('Monthly P/L Trend'));
            }
            if (elMonthlyDaily) {
                const ctxMonthlyDaily = elMonthlyDaily.getContext('2d');
                monthlyDailyChart = new Chart(ctxMonthlyDaily, createChartConfig('Daily P/L'));
            }
        } catch (e) {
            console.error("Error initializing charts:", e);
        }
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

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const pl = parseFloat(inputs[4].value);

                if (exitDateStr && !isNaN(pl)) {
                    const exitDate = parseRobustDate(exitDateStr);
                    if (exitDate && exitDate.getFullYear() === currentYear && exitDate.getMonth() === currentMonth) {
                        const date = exitDateStr; // Use the raw string for label
                        if (!dailyStats[date]) {
                            dailyStats[date] = 0;
                        }
                        dailyStats[date] += pl;
                    }
                }
            }
        });

        // Sort by Date
        const sortedDates = Object.keys(dailyStats).sort((a, b) => {
            const da = parseRobustDate(a);
            const db = parseRobustDate(b);
            const ta = da ? da.getTime() : 0;
            const tb = db ? db.getTime() : 0;
            return ta - tb;
        });

        const data = sortedDates.map(date => dailyStats[date]);

        // Colors: Green for >= 0, Red for < 0
        const colors = data.map(val => val >= 0 ? 'rgba(16, 185, 129, 0.7)' : 'rgba(239, 68, 68, 0.7)');

        dailyChart.data.labels = sortedDates;
        dailyChart.data.datasets[0].data = data;
        dailyChart.data.datasets[0].backgroundColor = colors;
        dailyChart.data.datasets[0].borderColor = colors.map(c => c.replace('0.7', '1'));
        dailyChart.update();
    }

    // --- Monthly View Daily Chart Logic ---
    function updateMonthlyDailyChart() {
        if (!monthlyDailyChart) initChart();

        const trs = document.getElementById('table-body').querySelectorAll('tr');
        const dailyStats = {};
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const pl = parseFloat(inputs[4].value);

                if (exitDateStr && !isNaN(pl)) {
                    const exitDate = parseRobustDate(exitDateStr);
                    if (exitDate) {
                        const monthKey = `${exitDate.getFullYear()}-${String(exitDate.getMonth() + 1).padStart(2, '0')}`;
                        if (includeCurrent || monthKey !== currentMonthKey) {
                            const date = exitDateStr;
                            if (!dailyStats[date]) dailyStats[date] = 0;
                            dailyStats[date] += pl;
                        }
                    }
                }
            }
        });

        // Sort for Chronology
        const sortedDates = Object.keys(dailyStats).sort((a, b) => {
            const da = parseRobustDate(a);
            const db = parseRobustDate(b);
            return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
        });

        const data = sortedDates.map(d => dailyStats[d]);
        const colors = data.map(val => val >= 0 ? 'rgba(16, 185, 129, 0.7)' : 'rgba(239, 68, 68, 0.7)');

        monthlyDailyChart.data.labels = sortedDates;
        monthlyDailyChart.data.datasets[0].data = data;
        monthlyDailyChart.data.datasets[0].backgroundColor = colors;
        monthlyDailyChart.data.datasets[0].borderColor = colors.map(c => c.replace('0.7', '1'));
        monthlyDailyChart.update();
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

        // Apply filters (including minimization) on load
        filterTable();

        // Update Calendar
        updateMonthlyCalendar();

        // Initial scroll to bottom (last row)
        requestAnimationFrame(() => {
            const lastRow = tableBody.lastElementChild;
            if (lastRow) {
                lastRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        });
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

    // --- Google Sheets Sync Logic (CORS-FREE BRIDGES) ---

    // 1. PUSH (Save) using Hidden Form to bypass CORS from file://
    function syncToGoogleSheets(rowsData) {
        if (!settingsGASUrl) return;

        // Visual feedback
        const btnSave = document.getElementById('btn-save-settings');
        if (btnSave) btnSave.textContent = 'Syncing...';

        // Show status on main screen too if possible
        let statusEl = document.getElementById('sync-status-indicator');
        if (!statusEl) {
            statusEl = document.createElement('div');
            statusEl.id = 'sync-status-indicator';
            statusEl.style.cssText = 'position:fixed;bottom:10px;right:10px;font-size:10px;color:#888;z-index:9999;';
            document.body.appendChild(statusEl);
        }
        statusEl.textContent = 'Last Sync: Sending...';

        // Create or reuse hidden bridge
        let bridge = document.getElementById('gas-post-bridge');
        if (!bridge) {
            bridge = document.createElement('div');
            bridge.id = 'gas-post-bridge';
            bridge.style.display = 'none';
            bridge.innerHTML = `
                <iframe name="gas-target"></iframe>
                <form method="POST" target="gas-target">
                    <textarea name="data"></textarea>
                </form>
            `;
            document.body.appendChild(bridge);
        }

        const form = bridge.querySelector('form');
        const textarea = form.querySelector('textarea');

        // Prepare simplified data package
        const dataPackage = {
            rows: rowsData,
            owners: appData.owners,
            types: appData.types,
            timestamp: new Date().toISOString()
        };

        textarea.value = JSON.stringify(dataPackage);
        form.action = settingsGASUrl;
        form.submit();

        // Assumption of success
        setTimeout(() => {
            if (btnSave) btnSave.textContent = 'Save & Sync';
            statusEl.textContent = 'Last Sync: ' + new Date().toLocaleTimeString();
        }, 1500);
    }

    // 2. PULL (Restore) using Script Tag (JSONP) to bypass CORS
    window.handleSyncRestore = function (data) {
        const btnPull = document.getElementById('btn-pull-data');

        if (data && data.status === 'success') {
            if (data.rows) localStorage.setItem('pl_report_rows', JSON.stringify(data.rows));
            if (data.portfolioRows) localStorage.setItem('portfolio_rows', JSON.stringify(data.portfolioRows));
            if (data.owners) localStorage.setItem('owners', JSON.stringify(data.owners));
            if (data.types) localStorage.setItem('types', JSON.stringify(data.types));

            alert('Data restored successfully! The page will now reload.');
            window.location.reload();
        } else {
            const msg = data.message || 'The script returned an error.';
            alert('Script Error: ' + msg);
            if (btnPull) {
                btnPull.textContent = 'Sync From Sheets (Restore)';
                btnPull.disabled = false;
            }
        }
    };

    function pullFromGoogleSheets() {
        const inputUrl = document.getElementById('setting-gas-url');
        let url = (inputUrl && inputUrl.value) ? inputUrl.value.trim() : settingsGASUrl;

        if (!url || !url.includes('script.google.com')) {
            alert('Please enter a valid Google Apps Script URL.');
            return;
        }

        // Auto-fix /exec
        if (!url.endsWith('/exec')) {
            url = url.split('?')[0];
            if (!url.endsWith('/')) url += '/';
            url += 'exec';
        }

        const btnPull = document.getElementById('btn-pull-data');
        if (btnPull) {
            btnPull.textContent = 'Pulling...';
            btnPull.disabled = true;
        }

        // Cleanup old script
        const oldScript = document.getElementById('gas-pull-script');
        if (oldScript) oldScript.remove();

        // Create Script Tag Injection (JSONP)
        const script = document.createElement('script');
        script.id = 'gas-pull-script';
        // Append action and callback parameters
        const sep = url.includes('?') ? '&' : '?';
        script.src = `${url}${sep}action=getData&callback=handleSyncRestore&cachebust=${Date.now()}`;

        script.onerror = () => {
            alert('Connection Failed. Please check your internet or Google Script deployment.');
            if (btnPull) {
                btnPull.textContent = 'Sync From Sheets (Restore)';
                btnPull.disabled = false;
            }
        };

        document.body.appendChild(script);
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
        }

        // Focus on first input (Date) - Child index 1 now (0 is S.No)
        const focusInput = tr.children[1].querySelector('input');
        if (!data && focusInput) focusInput.focus({ preventScroll: true });
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
        let visibleIndex = 1;
        rows.forEach((row) => {
            const cell = row.querySelector('.sno-cell') || row.firstElementChild;
            if (row.style.display !== 'none') {
                if (cell) cell.textContent = visibleIndex++;
            } else {
                if (cell) cell.textContent = '--';
            }
        });

        // Also update cumulative whenever serial numbers (order) change
        if (typeof updateCumulativePL === 'function') {
            updateCumulativePL();
        }
    }

    // Update Cumulative P/L Column (VISUAL ORDER: Follows top-to-bottom lines on screen)
    function updateCumulativePL() {
        const rows = tableBody.querySelectorAll('tr');
        let runningTotal = 0;

        rows.forEach(tr => {
            const cumCell = tr.querySelector('.cum-cell');
            if (cumCell) {
                // If the row is visible, add to running total
                if (tr.style.display !== 'none') {
                    const inputs = tr.querySelectorAll('input');
                    if (inputs.length >= 5) {
                        const plVal = parseFloat(inputs[4].value) || 0;
                        runningTotal += plVal;

                        cumCell.textContent = runningTotal.toFixed(2);
                        cumCell.className = 'cum-cell ' + (runningTotal >= 0 ? 'positive' : 'negative');
                        cumCell.style.color = runningTotal >= 0 ? 'var(--success-color)' : 'var(--danger-color)';
                    }
                } else {
                    // Hidden rows show nothing in cumulative
                    cumCell.textContent = '--';
                    cumCell.style.color = 'var(--text-secondary)';
                }
            }
        });

        // Also update the overall table summary
        updateTableSummary();
    }

    function updateTableSummary() {
        const rows = tableBody.querySelectorAll('tr');
        let visibleTotalPL = 0;

        rows.forEach(tr => {
            if (tr.style.display !== 'none') {
                const inputs = tr.querySelectorAll('input');
                if (inputs.length >= 5) {
                    const plVal = parseFloat(inputs[4].value) || 0;
                    visibleTotalPL += plVal;
                }
            }
        });

        const elTotalPL = document.getElementById('table-total-pl');
        if (elTotalPL) {
            elTotalPL.textContent = visibleTotalPL.toFixed(2);
            elTotalPL.className = 'value ' + (visibleTotalPL >= 0 ? 'positive' : 'negative');
        }
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

        const toggleMinimizePast = document.getElementById('toggle-minimize-past');
        const shouldMinimize = toggleMinimizePast ? toggleMinimizePast.checked : false;

        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        trs.forEach(tr => {
            try {
                let isVisible = true;
                const cells = tr.querySelectorAll('td');

                // 1. Minimize Past Months Filter
                if (shouldMinimize) {
                    const entryDateInput = cells[1]?.querySelector('input');
                    const exitDateInput = cells[4]?.querySelector('input');
                    const plInput = cells[5]?.querySelector('input'); // Index 5 is P/L

                    const entryDate = entryDateInput ? parseRobustDate(entryDateInput.value) : null;
                    const exitDate = exitDateInput ? parseRobustDate(exitDateInput.value) : null;
                    const plVal = plInput ? parseFloat(plInput.value) : NaN;

                    const currentMonthKey = currentYear * 12 + currentMonth;

                    // Decision logic:
                    // If Exit Date exists and is in the past -> Hide
                    if (exitDate) {
                        const exitMonthKey = exitDate.getFullYear() * 12 + exitDate.getMonth();
                        if (exitMonthKey < currentMonthKey) isVisible = false;
                    }
                    // If NO Exit Date but Entry Date is in the past, and NO P/L -> Hide (it's an old empty row)
                    else if (entryDate) {
                        const entryMonthKey = entryDate.getFullYear() * 12 + entryDate.getMonth();
                        if (entryMonthKey < currentMonthKey && isNaN(plVal)) {
                            isVisible = false;
                        }
                    }
                }

                // 2. Column Filters (only apply if still visible)
                if (isVisible) {
                    // Iterate over active filters
                    for (const [colIndex, filterValue] of Object.entries(filters)) {
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
                }

                tr.style.display = isVisible ? '' : 'none';
            } catch (e) {
                console.error("Error in filterTable row iteration:", e);
            }
        });

        // RE-CALCULATE Cumulative and Total after filtering
        updateCumulativePL();
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

    // --- Helper Functions (Hoisted manually to top of block) ---
    function parseRobustDate(dateStr) {
        if (!dateStr || typeof dateStr !== 'string') return null;
        dateStr = dateStr.trim();
        if (!dateStr) return null;

        // Handle YYYY-MM-DD (Standard ISO)
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
            const d = new Date(dateStr);
            return isNaN(d.getTime()) ? null : d;
        }

        // Handle DD-MM-YYYY or DD/MM/YYYY
        const parts = dateStr.split(/[-/]/);
        if (parts.length === 3) {
            // Case 1: YYYY is at the end (parts[2]) -> DD-MM-YYYY
            if (parts[2].length === 4) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10);
                const year = parseInt(parts[2], 10);
                const d = new Date(year, month - 1, day);
                return isNaN(d.getTime()) ? null : d;
            }
            // Case 2: YYYY is at the beginning (parts[0]) -> YYYY-MM-DD
            if (parts[0].length === 4) {
                const year = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10);
                const day = parseInt(parts[2], 10);
                const d = new Date(year, month - 1, day);
                return isNaN(d.getTime()) ? null : d;
            }
        }

        const d = new Date(dateStr);
        return isNaN(d.getTime()) ? null : d;
    }

    function getCellValue(cell) {
        if (!cell) return '';
        const input = cell.querySelector('input, select');
        return input ? input.value : cell.textContent.trim();
    }

    function updatePLColor(e) {
        const input = e.target;
        if (!input) return;
        const value = parseFloat(input.value);
        const cell = input.parentElement;
        if (cell) {
            cell.classList.remove('positive', 'negative');
            if (!isNaN(value)) {
                if (value > 0) cell.classList.add('positive');
                else if (value < 0) cell.classList.add('negative');
            }
        }
    }

    function compareValues(a, b, colIndex, dir) {
        // Date Columns (index 1 & 4 usually)
        if (colIndex === 1 || colIndex === 4) {
            const dateA = parseRobustDate(a);
            const dateB = parseRobustDate(b);
            const timeA = dateA ? dateA.getTime() : -1;
            const timeB = dateB ? dateB.getTime() : -1;
            return dir === 'asc' ? timeA - timeB : timeB - timeA;
        }

        // Number Columns (P/L is index 5)
        const numA = parseFloat(a);
        const numB = parseFloat(b);

        if (!isNaN(numA) && !isNaN(numB) && colIndex === 5) {
            return dir === 'asc' ? numA - numB : numB - numA;
        }

        // String Default
        const valA = (a || '').toLowerCase();
        const valB = (b || '').toLowerCase();

        if (valA < valB) return dir === 'asc' ? -1 : 1;
        if (valA > valB) return dir === 'asc' ? 1 : -1;
        return 0;
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

    const btnPullData = document.getElementById('btn-pull-data');
    if (btnPullData) {
        btnPullData.addEventListener('click', () => {
            if (confirm('This will replace your current local data with the data from Google Sheets. Continue?')) {
                pullFromGoogleSheets();
            }
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

    function showDrillDown(filterType, filterValue, monthKey = null) {
        if (!detailsModal) return;

        let title = `${filterType}: ${filterValue} - Trades`;
        if (monthKey) {
            // Human readable month (e.g. Feb 2026)
            const [y, m] = monthKey.split('-');
            const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const monthName = monthNames[parseInt(m) - 1] || m;
            title += ` (${monthName} ${y})`;
        }

        if (detailsModalTitle) detailsModalTitle.textContent = title;
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

            // Secondary Filter: If monthKey is provided, trade must be in that month
            if (monthKey) {
                if (exitVal === '--') return; // Skip open trades when scoped to a month
                const dateObj = parseRobustDate(exitVal);
                if (dateObj) {
                    const trMonthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                    if (trMonthKey !== monthKey) return; // Skip if month doesn't match
                } else {
                    return; // Skip if date is invalid but required
                }
            }

            if (filterType === 'Owner' && ownerVal === filterValue) match = true;
            if (filterType === 'Type' && typeVal === filterValue) match = true;
            if (filterType === 'Date' && exitVal === filterValue) match = true;
            if (filterType === 'Owner|Type') {
                const [targetOwner, targetType] = filterValue.split('|');
                if (ownerVal === targetOwner && typeVal === targetType) match = true;
            }
            if (filterType === 'Month') {
                const dateObj = parseRobustDate(exitVal);
                if (dateObj) {
                    const mKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                    if (mKey === filterValue) match = true;
                }
            }

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
            detailsTableBody.innerHTML = '<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--text-secondary);">No trades found</td></tr>';
        }

        detailsModal.classList.remove('hidden');
    }

    function updateMonthlyCalendar() {
        const container = document.getElementById('monthly-calendar-container');
        if (!container) return;
        container.innerHTML = '';

        const trs = tableBody.querySelectorAll('tr');
        const dailyPL = {};
        const includeCurrent = document.getElementById('toggle-include-current')?.checked !== false;
        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();

        // Determine current range
        const activeRangeBtn = document.querySelector('.range-btn.active');
        const range = activeRangeBtn ? activeRangeBtn.getAttribute('data-range') : '1y';

        // 1. Aggregate P/L by Day (YYYY-MM-DD)
        trs.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            if (inputs.length >= 5) {
                const exitDateStr = inputs[3].value;
                const pl = parseFloat(inputs[4].value) || 0;
                if (exitDateStr) {
                    const dateObj = parseRobustDate(exitDateStr);
                    if (dateObj) {
                        const dayKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
                        dailyPL[dayKey] = (dailyPL[dayKey] || 0) + pl;
                    }
                }
            }
        });

        // 2. Determine Month Range
        const monthKeys = [];
        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        // Determine precise limits for filtering dots (LOCAL TIME)
        let limitStart = null;
        let limitEnd = null;

        if (range === 'custom') {
            const startVal = document.getElementById('calendar-start-date').value;
            const endVal = document.getElementById('calendar-end-date').value;
            if (startVal && endVal) {
                const s = startVal.split('-');
                limitStart = new Date(parseInt(s[0]), parseInt(s[1]) - 1, parseInt(s[2]), 0, 0, 0, 0);
                const e = endVal.split('-');
                limitEnd = new Date(parseInt(e[0]), parseInt(e[1]) - 1, parseInt(e[2]), 23, 59, 59, 999);
            }
        } else if (range === '1w') {
            limitStart = new Date();
            limitStart.setDate(limitStart.getDate() - 6);
            limitStart.setHours(0, 0, 0, 0);
            limitEnd = new Date();
            limitEnd.setHours(23, 59, 59, 999);
        } else if (range === '15d') {
            limitStart = new Date();
            limitStart.setDate(limitStart.getDate() - 14);
            limitStart.setHours(0, 0, 0, 0);
            limitEnd = new Date();
            limitEnd.setHours(23, 59, 59, 999);
        }

        // 2. Determine Month Range

        if (range === 'custom') {
            if (limitStart && limitEnd) {
                // Parse dates and move to start of month for comparison
                let d = new Date(limitStart.getFullYear(), limitStart.getMonth(), 1);
                const endD = new Date(limitEnd.getFullYear(), limitEnd.getMonth(), 1);

                let count = 0;
                while (d <= endD && count < 60) {
                    monthKeys.push({
                        year: d.getFullYear(),
                        month: d.getMonth(),
                        label: `${monthNames[d.getMonth()]} ${String(d.getFullYear()).slice(-2)}`
                    });
                    d.setMonth(d.getMonth() + 1);
                    count++;
                }
            }
        } else {
            let monthsToDisplay = 12;
            if (range === '1w' || range === '15d' || range === '1m') monthsToDisplay = 1;
            else if (range === '3m') monthsToDisplay = 3;
            else if (range === '1y') monthsToDisplay = 12;

            // Start from current month and go back
            for (let i = 0; i < monthsToDisplay; i++) {
                const d = new Date(currentYear, currentMonth - i, 1);
                if (!includeCurrent && d.getFullYear() === currentYear && d.getMonth() === currentMonth) continue;
                monthKeys.unshift({
                    year: d.getFullYear(),
                    month: d.getMonth(),
                    label: `${monthNames[d.getMonth()]} ${String(d.getFullYear()).slice(-2)}`
                });
            }
        }

        let cumulativeTotal = 0;

        // 3. Render Each Month
        monthKeys.forEach(m => {
            const monthCol = document.createElement('div');
            monthCol.className = 'month-column';

            const label = document.createElement('div');
            label.className = 'month-label';
            label.textContent = m.label;
            monthCol.appendChild(label);

            // Weekday Headers
            const headers = document.createElement('div');
            headers.className = 'weekday-headers';
            ["M", "T", "W", "T", "F", "S", "S"].forEach(day => {
                const h = document.createElement('div');
                h.className = 'weekday-header';
                h.textContent = day;
                headers.appendChild(h);
            });
            monthCol.appendChild(headers);

            const grid = document.createElement('div');
            grid.className = 'days-dots-grid';

            // Alignment Padding
            const firstDay = new Date(m.year, m.month, 1);
            let dayOfWeek = firstDay.getDay(); // 0 (Sun) to 6 (Sat)
            let paddingCount = (dayOfWeek + 6) % 7; // Convert to Mon-Sun index (0-6)

            for (let i = 0; i < paddingCount; i++) {
                const emptyDot = document.createElement('div');
                emptyDot.className = 'calendar-dot empty';
                grid.appendChild(emptyDot);
            }

            let monthTotal = 0;
            const daysInMonth = new Date(m.year, m.month + 1, 0).getDate();

            for (let day = 1; day <= daysInMonth; day++) {
                const dayKey = `${m.year}-${String(m.month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                const pl = dailyPL[dayKey];
                const currentDate = new Date(m.year, m.month, day);

                const dot = document.createElement('div');
                dot.className = 'calendar-dot';

                // Check if day is within precise limits
                let isOutOfRange = false;
                if (limitStart && currentDate < limitStart) isOutOfRange = true;
                if (limitEnd && currentDate > limitEnd) isOutOfRange = true;

                if (isOutOfRange) {
                    dot.classList.add('empty');
                } else {
                    if (pl !== undefined) {
                        dot.classList.add(pl >= 0 ? 'profit' : 'loss');
                        dot.title = `${dayKey}: ₹${pl.toFixed(2)}`;
                        monthTotal += pl;
                    } else {
                        dot.title = `${dayKey}: No Trades`;
                    }
                }

                grid.appendChild(dot);
            }

            cumulativeTotal += monthTotal;
            monthCol.appendChild(grid);

            const totalEl = document.createElement('div');
            const absTotal = Math.abs(monthTotal).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            totalEl.className = 'month-total-pl ' + (monthTotal >= 0 ? 'positive' : 'negative');
            totalEl.textContent = `${monthTotal < 0 ? '- ' : ''}₹${absTotal}`;

            monthCol.appendChild(totalEl);
            container.appendChild(monthCol);
        });

        // Update Grand Total UI
        const grandTotalContainer = document.getElementById('calendar-grand-total-container');
        const grandTotalValue = document.getElementById('calendar-grand-total-value');
        if (grandTotalContainer && grandTotalValue) {
            grandTotalContainer.style.display = 'flex';
            const absGrand = Math.abs(cumulativeTotal).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            grandTotalValue.textContent = `${cumulativeTotal < 0 ? '- ' : ''}₹${absGrand}`;
            grandTotalValue.style.color = cumulativeTotal >= 0 ? 'var(--success-color)' : 'var(--danger-color)';
        }
    }

    // --- Final Execution ---
    try {
        loadAllData();
        if (tableBody && tableBody.children.length === 0) {
            addRow();
        }
    } catch (e) {
        console.error("Critical error during data initialization:", e);
    }
});

