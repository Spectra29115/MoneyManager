/* ============================================
   MoneyFlow — App Logic
   ============================================ */

(function () {
    'use strict';

    // ---- Constants ----
    const STORAGE_KEY = 'moneyflow_data';
    const UI_STATE_KEY = 'moneyflow_ui_state';

    let uiState = {
        collapsedDates: [],
        filterCategory: 'all',
        filterProfile: 'all'
    };

    let lastAddedId = null;

    const CATEGORIES = [
        { value: 'Food', label: '🍔 Food' },
        { value: 'Transport', label: '🚗 Transport' },
        { value: 'Shopping', label: '🛍️ Shopping' },
        { value: 'Bills', label: '📄 Bills' },
        { value: 'Entertainment', label: '🎮 Entertainment' },
        { value: 'Health', label: '💊 Health' },
        { value: 'Education', label: '📚 Education' },
        { value: 'Salary', label: '💼 Salary' },
        { value: 'Freelance', label: '💻 Freelance' },
        { value: 'Investment', label: '📊 Investment' },
        { value: 'Other', label: '📌 Other' },
    ];

    const CATEGORY_COLORS = {
        Food: '#f59e0b',
        Transport: '#3b82f6',
        Shopping: '#ec4899',
        Bills: '#8b5cf6',
        Entertainment: '#06b6d4',
        Health: '#10b981',
        Education: '#6366f1',
        Salary: '#22c55e',
        Freelance: '#14b8a6',
        Investment: '#a855f7',
        Other: '#64748b',
    };

    // ---- Helpers ----
    function generateId() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
    }

    function formatCurrency(amount) {
        const num = Number(amount) || 0;
        return '₹' + num.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }

    function formatLocalDate(d) {
        if (!(d instanceof Date) || isNaN(d.getTime())) return null;
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function todayStr() {
        return formatLocalDate(new Date());
    }

    function parseDate(str) {
        const d = new Date(str + 'T00:00:00');
        return isNaN(d.getTime()) ? new Date() : d;
    }

    function isThisWeek(dateStr) {
        const d = parseDate(dateStr);
        const now = new Date();
        const startOfWeek = new Date(now);
        startOfWeek.setDate(now.getDate() - now.getDay());
        startOfWeek.setHours(0, 0, 0, 0);
        return d >= startOfWeek;
    }

    function isThisMonth(dateStr) {
        const d = parseDate(dateStr);
        const now = new Date();
        return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }

    // ---- Transaction Manager ----
    class TransactionManager {
        constructor() {
            this.transactions = this.load();
        }

        load() {
            try {
                const data = localStorage.getItem(STORAGE_KEY);
                return data ? JSON.parse(data) : [];
            } catch {
                return [];
            }
        }

        save() {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(this.transactions));
        }

        add(tx) {
            this.transactions.push(tx);
            this.sortByDate();
            this.save();
        }

        update(id, fields) {
            const idx = this.transactions.findIndex(t => t.id === id);
            if (idx !== -1) {
                Object.assign(this.transactions[idx], fields);
                this.sortByDate();
                this.save();
            }
        }

        remove(id) {
            this.transactions = this.transactions.filter(t => t.id !== id);
            this.save();
        }

        clearAll() {
            this.transactions = [];
            this.save();
        }

        sortByDate() {
            this.transactions.sort((a, b) => {
                const da = parseDate(a.date);
                const db = parseDate(b.date);
                return da - db;
            });
        }

        getAll(categoryFilter, profileFilter) {
            let filtered = [...this.transactions];
            if (categoryFilter && categoryFilter !== 'all') {
                filtered = filtered.filter(t => t.category === categoryFilter);
            }
            if (profileFilter && profileFilter !== 'all') {
                filtered = filtered.filter(t => (t.profile || 'OVERALL') === profileFilter);
            }
            return filtered;
        }

        getStats() {
            let totalCredit = 0;
            let totalDebit = 0;
            let overallCredit = 0, overallDebit = 0;
            let suhasCredit = 0, suhasDebit = 0;
            let monthlyDebit = 0;

            this.transactions.forEach(t => {
                const credit = Number(t.credit) || 0;
                const debit = Number(t.debit) || 0;
                totalCredit += credit;
                totalDebit += debit;

                if (t.profile === 'SUHAS') {
                    suhasCredit += credit;
                    suhasDebit += debit;
                } else {
                    overallCredit += credit;
                    overallDebit += debit;
                }

                if (isThisMonth(t.date)) monthlyDebit += debit;
            });

            return {
                overallBalance: overallCredit - overallDebit,
                suhasBalance: suhasCredit - suhasDebit,
                totalBalance: totalCredit - totalDebit,
                totalCredit,
                totalDebit,
                monthlyDebit
            };
        }

        getCategoryTotals() {
            const map = {};
            this.transactions.forEach(t => {
                const debit = Number(t.debit) || 0;
                if (debit > 0 && t.category) {
                    map[t.category] = (map[t.category] || 0) + debit;
                }
            });
            return map;
        }

        getLast7DaysSpending() {
            const days = [];
            const now = new Date();
            for (let i = 6; i >= 0; i--) {
                const d = new Date(now);
                d.setDate(now.getDate() - i);
                const key = formatLocalDate(d);
                const label = d.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric' });
                days.push({ key, label, total: 0 });
            }

            this.transactions.forEach(t => {
                const debit = Number(t.debit) || 0;
                const day = days.find(d => d.key === t.date);
                if (day) day.total += debit;
            });

            return days;
        }
    }

    // ---- App Init ----
    const manager = new TransactionManager();
    let selectedIds = new Set();

    // DOM refs
    const ledgerBody = document.getElementById('ledgerBody');
    const emptyState = document.getElementById('emptyState');
    const addRowBtn = document.getElementById('addRowBtn');
    const filterCategory = document.getElementById('filterCategory');
    const filterProfile = document.getElementById('filterProfile');
    const currentMonthEl = document.getElementById('currentMonth');

    // Bulk action refs
    const selectAllCheck = document.getElementById('selectAll');
    const bulkActionsBar = document.getElementById('bulkActions');
    const selectedCountText = document.getElementById('selectedCount');
    const btnBulkDelete = document.getElementById('btnBulkDelete');

    // Stat elements
    const statOverall = document.getElementById('statOverall');
    const statSuhas = document.getElementById('statSuhas');
    const statTotal = document.getElementById('statTotal');
    const statCredited = document.getElementById('statCredited');
    const statDebited = document.getElementById('statDebited');
    const statMonthly = document.getElementById('statMonthly');

    // Chart canvases
    const pieCanvas = document.getElementById('pieChart');
    const barCanvas = document.getElementById('barChart');
    const pieLegend = document.getElementById('pieLegend');

    function setCurrentMonth() {
        const now = new Date();
        currentMonthEl.textContent = now.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
    }

    // ---- UI State Persistence ----
    function loadUIState() {
        try {
            const data = localStorage.getItem(UI_STATE_KEY);
            if (data) {
                const parsed = JSON.parse(data);
                uiState = { ...uiState, ...parsed };
                // Apply to DOM
                if (filterCategory) filterCategory.value = uiState.filterCategory;
                if (filterProfile) filterProfile.value = uiState.filterProfile;
            }
        } catch (e) {}
    }

    function saveUIState() {
        if (filterCategory) uiState.filterCategory = filterCategory.value;
        if (filterProfile) uiState.filterProfile = filterProfile.value;
        localStorage.setItem(UI_STATE_KEY, JSON.stringify(uiState));
    }

    // ---- Render everything ----
    function render() {
        renderTable();
        updateStats();
        drawPieChart();
        drawBarChart();
    }

    // ---- Render spreadsheet table ----
    function renderTable() {
        const catFilter = filterCategory.value;
        const profFilter = filterProfile.value;
        const txs = manager.getAll(catFilter, profFilter);
        ledgerBody.innerHTML = '';

        if (txs.length === 0) {
            emptyState.classList.add('visible');
            return;
        }

        emptyState.classList.remove('visible');

        // We need to compute running balance separately for OVERALL and SUHAS
        const allTxs = manager.getAll();
        const balanceMap = {};
        let rbOverall = 0;
        let rbSuhas = 0;

        allTxs.forEach(t => {
            const net = (Number(t.credit) || 0) - (Number(t.debit) || 0);
            if (t.profile === 'SUHAS') {
                rbSuhas += net;
                balanceMap[t.id] = rbSuhas;
            } else {
                rbOverall += net;
                balanceMap[t.id] = rbOverall;
            }
        });

        const grouped = {};
        txs.forEach(tx => {
            if (!grouped[tx.date]) grouped[tx.date] = [];
            grouped[tx.date].push(tx);
        });

        // Parse dates to sort them latest first
        const sortedDates = Object.keys(grouped).sort((a,b) => new Date(b) - new Date(a));

        let rowIdx = 1;

        sortedDates.forEach(date => {
            const groupTxs = grouped[date];
            let gCredit = 0;
            let gDebit = 0;
            groupTxs.forEach(t => {
                gCredit += (Number(t.credit) || 0);
                gDebit += (Number(t.debit) || 0);
            });

            // The closing balance for the day is the balance of the last transaction in that group
            // (transactions are added chronologically, so the last one in the array is the latest)
            const lastTx = groupTxs[groupTxs.length - 1];
            const closingBal = balanceMap[lastTx.id] || 0;

            let displayDate = date;
            const dObj = new Date(date + 'T00:00:00');
            if (!isNaN(dObj.getTime())) {
                displayDate = dObj.toLocaleDateString('en-IN', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' });
            }

            const headerRow = document.createElement('tr');
            headerRow.className = 'date-group-header';

            const isCollapsed = uiState.collapsedDates.includes(date);
            const icon = isCollapsed ? '▶' : '▼';

            headerRow.innerHTML = `
                <td colspan="6" class="group-title-cell">
                    <button class="btn-collapse" data-date="${date}">${icon}</button>
                    <strong>${displayDate}</strong>
                    <span class="group-count">(${groupTxs.length} item${groupTxs.length !== 1 ? 's' : ''})</span>
                </td>
                <td class="cell-credit group-credit">${gCredit > 0 ? '+' : ''}${formatCurrency(gCredit)}</td>
                <td class="cell-debit group-debit">${gDebit > 0 ? '-' : ''}${formatCurrency(gDebit)}</td>
                <td class="cell-balance group-balance">
                    <span class="balance-label">Closing:</span>
                    <span class="${closingBal >= 0 ? 'positive' : 'negative'}">${formatCurrency(closingBal)}</span>
                </td>
                <td class="col-actions"></td>
            `;

            headerRow.querySelector('.btn-collapse').addEventListener('click', (e) => {
                const currentlyCollapsed = uiState.collapsedDates.includes(date);
                if (currentlyCollapsed) {
                    uiState.collapsedDates = uiState.collapsedDates.filter(d => d !== date);
                    e.target.textContent = '▼';
                } else {
                    uiState.collapsedDates.push(date);
                    e.target.textContent = '▶';
                }

                const children = ledgerBody.querySelectorAll(`.date-child-${date}`);
                children.forEach(c => {
                    c.style.display = currentlyCollapsed ? '' : 'none';
                });

                saveUIState();
            });

            ledgerBody.appendChild(headerRow);

            // Hybrid Internal Sorting:
            // 1. The entry matching lastAddedId always goes to the top.
            // 2. All other entries follow their original ascending (oldest-first) order.
            const sortedGroupTxs = [...groupTxs].sort((a, b) => {
                if (a.id === lastAddedId) return -1;
                if (b.id === lastAddedId) return 1;
                return 0; // Maintain original ascending order from manager.transactions
            });

            sortedGroupTxs.forEach(tx => {
                const row = document.createElement('tr');
                row.dataset.id = tx.id;
                row.className = `date-child-${date}`;
                if (isCollapsed) row.style.display = 'none';

                const balance = balanceMap[tx.id] || 0;
                const balanceClass = balance >= 0 ? 'positive' : 'negative';

                // Build category options
                let categoryOptions = '<option value="">Select</option>';
                CATEGORIES.forEach(c => {
                    const selected = c.value === tx.category ? 'selected' : '';
                    categoryOptions += `<option value="${c.value}" ${selected}>${c.label}</option>`;
                });

                // Build profile options
                const profOptions = `
                    <option value="OVERALL" ${tx.profile !== 'SUHAS' ? 'selected' : ''}>🏦 OVERALL</option>
                    <option value="SUHAS" ${tx.profile === 'SUHAS' ? 'selected' : ''}>👤 SUHAS</option>
                `;

                if (selectedIds.has(tx.id)) {
                    row.classList.add('selected-row');
                }

                row.innerHTML = `
                    <td class="cell-check">
                        <input type="checkbox" class="row-checkbox" value="${tx.id}" ${selectedIds.has(tx.id) ? 'checked' : ''}>
                    </td>
                    <td class="cell-sno">${rowIdx++}</td>
                    <td><input type="date" class="date-input" value="${tx.date}" data-field="date"></td>
                    <td><select class="profile-select" data-field="profile">${profOptions}</select></td>
                    <td contenteditable="true" data-field="description" data-placeholder="Description">${tx.description || ''}</td>
                    <td><select class="category-select" data-field="category">${categoryOptions}</select></td>
                    <td contenteditable="true" class="cell-credit" data-field="credit" data-placeholder="0">${tx.credit || ''}</td>
                    <td contenteditable="true" class="cell-debit" data-field="debit" data-placeholder="0">${tx.debit || ''}</td>
                    <td class="cell-balance ${balanceClass}">${formatCurrency(balance)}</td>
                    <td><button class="btn-delete" title="Delete">🗑️</button></td>
                `;

                // Event: inline edit blur
                row.querySelectorAll('[contenteditable="true"]').forEach(cell => {
                    cell.addEventListener('blur', () => handleCellEdit(tx.id, cell));
                    cell.addEventListener('keydown', (e) => {
                        if (e.key === 'Enter') {
                            e.preventDefault();
                            cell.blur();
                        }
                    });
                });

                // Event: date change
                row.querySelector('.date-input').addEventListener('change', (e) => {
                    manager.update(tx.id, { date: e.target.value });
                    render();
                    showToast('Date updated');
                });

                // Event: profile change
                row.querySelector('.profile-select').addEventListener('change', (e) => {
                    manager.update(tx.id, { profile: e.target.value });
                    render();
                    showToast('Profile updated');
                });

                // Event: category change
                row.querySelector('.category-select').addEventListener('change', (e) => {
                    manager.update(tx.id, { category: e.target.value });
                    render();
                    showToast('Category updated');
                });

                // Event: delete
                row.querySelector('.btn-delete').addEventListener('click', () => {
                    row.style.animation = 'rowSlideIn 0.3s ease-out reverse';
                    setTimeout(() => {
                        selectedIds.delete(tx.id);
                        manager.remove(tx.id);
                        render();
                        showToast('Entry deleted');
                    }, 280);
                });

                // Event: row checkbox
                row.querySelector('.row-checkbox').addEventListener('change', (e) => {
                    if (e.target.checked) {
                        selectedIds.add(tx.id);
                        row.classList.add('selected-row');
                    } else {
                        selectedIds.delete(tx.id);
                        row.classList.remove('selected-row');
                    }
                    updateBulkUI();
                });

                ledgerBody.appendChild(row);
            });
        });

        // Update Select All Checkbox state
        selectAllCheck.checked = txs.length > 0 && selectedIds.size === txs.length;
        updateBulkUI();
    }

    function updateBulkUI() {
        if (selectedIds.size > 0) {
            bulkActionsBar.classList.remove('hidden');
            selectedCountText.textContent = `${selectedIds.size} selected`;
        } else {
            bulkActionsBar.classList.add('hidden');
            selectAllCheck.checked = false;
        }
    }

    function handleCellEdit(id, cell) {
        const field = cell.dataset.field;
        let value = cell.textContent.trim();

        // For numeric fields, strip non-numeric chars
        if (field === 'credit' || field === 'debit') {
            value = value.replace(/[^0-9.]/g, '');
            const num = parseFloat(value);
            cell.textContent = isNaN(num) ? '' : num;
            value = isNaN(num) ? 0 : num;
        }

        manager.update(id, { [field]: value });
        render();
    }

    // ---- Update stats ----
    function updateStats() {
        const stats = manager.getStats();
        statOverall.textContent = formatCurrency(stats.overallBalance);
        statSuhas.textContent = formatCurrency(stats.suhasBalance);
        statTotal.textContent = formatCurrency(stats.totalBalance);
        statCredited.textContent = formatCurrency(stats.totalCredit);
        statDebited.textContent = formatCurrency(stats.totalDebit);
        statMonthly.textContent = formatCurrency(stats.monthlyDebit);
    }

    // ---- Pie Chart (Category Breakdown) ----
    function drawPieChart() {
        const ctx = pieCanvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const size = 300;
        pieCanvas.width = size * dpr;
        pieCanvas.height = size * dpr;
        pieCanvas.style.width = size + 'px';
        pieCanvas.style.height = size + 'px';
        ctx.scale(dpr, dpr);

        const catTotals = manager.getCategoryTotals();
        const entries = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);
        const total = entries.reduce((s, e) => s + e[1], 0);

        ctx.clearRect(0, 0, size, size);

        if (total === 0) {
            // Draw empty state
            ctx.beginPath();
            ctx.arc(size / 2, size / 2, 100, 0, Math.PI * 2);
            ctx.strokeStyle = 'rgba(255,255,255,0.06)';
            ctx.lineWidth = 40;
            ctx.stroke();

            ctx.fillStyle = '#64748b';
            ctx.font = '500 14px Inter, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText('No spending data', size / 2, size / 2 + 5);

            pieLegend.innerHTML = '';
            return;
        }

        const cx = size / 2;
        const cy = size / 2;
        const radius = 100;
        const innerRadius = 60;
        let startAngle = -Math.PI / 2;

        entries.forEach(([cat, val]) => {
            const sliceAngle = (val / total) * Math.PI * 2;
            const color = CATEGORY_COLORS[cat] || '#64748b';

            ctx.beginPath();
            ctx.arc(cx, cy, radius, startAngle, startAngle + sliceAngle);
            ctx.arc(cx, cy, innerRadius, startAngle + sliceAngle, startAngle, true);
            ctx.closePath();
            ctx.fillStyle = color;
            ctx.fill();

            startAngle += sliceAngle;
        });

        // Center text
        ctx.fillStyle = '#f1f5f9';
        ctx.font = '700 18px Inter, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(formatCurrency(total), cx, cy - 6);
        ctx.fillStyle = '#64748b';
        ctx.font = '500 11px Inter, sans-serif';
        ctx.fillText('Total Spent', cx, cy + 14);

        // Legend
        pieLegend.innerHTML = entries.map(([cat, val]) => {
            const pct = ((val / total) * 100).toFixed(1);
            const color = CATEGORY_COLORS[cat] || '#64748b';
            const emoji = CATEGORIES.find(c => c.value === cat)?.label.split(' ')[0] || '📌';
            return `<div class="pie-legend-item">
                <span class="pie-legend-color" style="background:${color}"></span>
                ${emoji} ${cat} — ${formatCurrency(val)} (${pct}%)
            </div>`;
        }).join('');
    }

    // ---- Bar Chart (Last 7 Days) ----
    function drawBarChart() {
        const ctx = barCanvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const width = 500;
        const height = 300;
        barCanvas.width = width * dpr;
        barCanvas.height = height * dpr;
        barCanvas.style.width = width + 'px';
        barCanvas.style.height = height + 'px';
        ctx.scale(dpr, dpr);

        const days = manager.getLast7DaysSpending();
        const maxVal = Math.max(...days.map(d => d.total), 1);

        ctx.clearRect(0, 0, width, height);

        const padding = { top: 20, right: 20, bottom: 50, left: 60 };
        const chartW = width - padding.left - padding.right;
        const chartH = height - padding.top - padding.bottom;
        const barWidth = chartW / days.length * 0.6;
        const gap = chartW / days.length;

        // Y-axis grid lines
        const gridLines = 5;
        for (let i = 0; i <= gridLines; i++) {
            const y = padding.top + (chartH / gridLines) * i;
            ctx.beginPath();
            ctx.moveTo(padding.left, y);
            ctx.lineTo(width - padding.right, y);
            ctx.strokeStyle = 'rgba(255, 255, 255, 0.05)';
            ctx.lineWidth = 1;
            ctx.stroke();

            // Y-axis labels
            const val = maxVal - (maxVal / gridLines) * i;
            ctx.fillStyle = '#64748b';
            ctx.font = '500 10px Inter, sans-serif';
            ctx.textAlign = 'right';
            ctx.textBaseline = 'middle';
            ctx.fillText(formatCurrency(val), padding.left - 8, y);
        }

        // Bars
        days.forEach((day, i) => {
            const x = padding.left + gap * i + (gap - barWidth) / 2;
            const barH = (day.total / maxVal) * chartH;
            const y = padding.top + chartH - barH;

            // Bar gradient
            const grad = ctx.createLinearGradient(x, y, x, padding.top + chartH);
            grad.addColorStop(0, '#3b82f6');
            grad.addColorStop(1, 'rgba(59, 130, 246, 0.2)');

            ctx.beginPath();
            const r = 4;
            ctx.moveTo(x + r, y);
            ctx.lineTo(x + barWidth - r, y);
            ctx.quadraticCurveTo(x + barWidth, y, x + barWidth, y + r);
            ctx.lineTo(x + barWidth, padding.top + chartH);
            ctx.lineTo(x, padding.top + chartH);
            ctx.lineTo(x, y + r);
            ctx.quadraticCurveTo(x, y, x + r, y);
            ctx.closePath();
            ctx.fillStyle = grad;
            ctx.fill();

            // Value on top
            if (day.total > 0) {
                ctx.fillStyle = '#94a3b8';
                ctx.font = '600 9px Inter, sans-serif';
                ctx.textAlign = 'center';
                ctx.fillText(formatCurrency(day.total), x + barWidth / 2, y - 8);
            }

            // Day label
            ctx.fillStyle = '#94a3b8';
            ctx.font = '500 10px Inter, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText(day.label, x + barWidth / 2, padding.top + chartH + 20);
        });
    }

    // ---- Add new row ----
    function addNewRow() {
        const tx = {
            id: generateId(),
            date: todayStr(),
            profile: filterProfile.value !== 'all' ? filterProfile.value : 'OVERALL',
            description: '',
            category: '',
            credit: '',
            debit: '',
        };
        lastAddedId = tx.id;
        manager.add(tx);
        render();

        // Scroll and focus the specific newly added row by its ID
        setTimeout(() => {
            const newRow = ledgerBody.querySelector(`tr[data-id="${tx.id}"]`);
            if (newRow) {
                newRow.classList.add('row-new');
                newRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                const descCell = newRow.querySelector('[data-field="description"]');
                if (descCell) descCell.focus();
            }
        }, 50);

        showToast('New entry added — fill in the details!');
    }

    // ---- Toast notification ----
    function showToast(message, isError = false) {
        // Remove existing toasts
        document.querySelectorAll('.toast').forEach(t => t.remove());

        const toast = document.createElement('div');
        toast.className = 'toast' + (isError ? ' toast-error' : '');
        toast.textContent = message;
        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(20px)';
            toast.style.transition = '0.3s ease-out';
            setTimeout(() => toast.remove(), 300);
        }, 2500);
    }

    // ---- Event Listeners ----
    addRowBtn.addEventListener('click', addNewRow);

    // ---- Delete All ----
    const deleteAllBtn = document.getElementById('deleteAllBtn');
    deleteAllBtn.addEventListener('click', () => {
        const count = manager.getAll().length;
        if (count === 0) {
            showToast('Nothing to delete!', true);
            return;
        }
        if (confirm(`Are you sure you want to delete all ${count} entries? This cannot be undone.`)) {
            manager.clearAll();
            render();
            showToast(`Deleted all ${count} entries.`);
        }
    });

    filterCategory.addEventListener('change', () => {
        selectedIds.clear();
        saveUIState();
        render();
    });
    filterProfile.addEventListener('change', () => {
        selectedIds.clear();
        saveUIState();
        render();
    });

    selectAllCheck.addEventListener('change', (e) => {
        const isChecked = e.target.checked;
        const currentTxs = manager.getAll(filterCategory.value, filterProfile.value);
        if (isChecked) {
            currentTxs.forEach(tx => selectedIds.add(tx.id));
        } else {
            currentTxs.forEach(tx => selectedIds.delete(tx.id));
        }
        render();
    });

    btnBulkDelete.addEventListener('click', () => {
        if (selectedIds.size === 0) return;
        if (confirm(`Delete ${selectedIds.size} selected entries?`)) {
            selectedIds.forEach(id => manager.remove(id));
            selectedIds.clear();
            render();
            showToast('Deleted selected entries.');
        }
    });

    // ---- Keyboard shortcut: Ctrl+N to add row ----
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'n') {
            e.preventDefault();
            addNewRow();
        }
    });

    // ==================================================
    //  EXCEL IMPORT FEATURE
    // ==================================================

    const importBtn = document.getElementById('importBtn');
    const fileInput = document.getElementById('fileInput');
    const importModal = document.getElementById('importModal');
    const modalClose = document.getElementById('modalClose');
    const modalCancel = document.getElementById('modalCancel');
    const modalImport = document.getElementById('modalImport');
    const mappingGrid = document.getElementById('mappingGrid');
    const previewHead = document.getElementById('previewHead');
    const previewBody = document.getElementById('previewBody');
    const previewCount = document.getElementById('previewCount');

    let importedSheetData = []; // raw rows from Excel
    let importedHeaders = [];   // column headers from sheet

    // Our fields to map
    const IMPORT_FIELDS = [
        { key: 'date', label: 'Date', keywords: ['date', 'dated', 'day', 'time', 'txn date', 'transaction date', 'trans date', 'value date'] },
        { key: 'profile', label: 'Profile', keywords: ['profile', 'account', 'bank', 'savings', 'wallet'] },
        { key: 'description', label: 'Description', keywords: ['description', 'desc', 'narration', 'particular', 'particulars', 'details', 'remark', 'remarks', 'memo', 'note', 'notes', 'transaction'] },
        { key: 'category', label: 'Category', keywords: ['category', 'cat', 'type', 'group', 'head'] },
        { key: 'credit', label: 'Credit (₹)', keywords: ['credit', 'credited', 'deposit', 'income', 'cr', 'amount credited', 'money in', 'inflow'] },
        { key: 'debit', label: 'Debit (₹)', keywords: ['debit', 'debited', 'withdrawal', 'expense', 'dr', 'amount debited', 'money out', 'outflow', 'spent'] },
    ];

    // Open file picker
    importBtn.addEventListener('click', () => {
        fileInput.value = '';
        fileInput.click();
    });

    // Handle file selection
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const isPdf = file.name.toLowerCase().endsWith('.pdf');

        if (isPdf) {
            handlePdfFile(file);
        } else {
            handleExcelFile(file);
        }
    });

    // ---- Excel/CSV handler ----
    function handleExcelFile(file) {
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                if (rawData.length < 2) {
                    showToast('The spreadsheet appears to be empty or has no data rows.', true);
                    return;
                }

                importedHeaders = rawData[0].map(h => String(h).trim());
                importedSheetData = rawData.slice(1).filter(row => row.some(cell => cell !== ''));

                if (importedSheetData.length === 0) {
                    showToast('No data rows found in the spreadsheet.', true);
                    return;
                }

                buildMappingUI();
                showPreview();
                openImportModal();

            } catch (err) {
                console.error('Import error:', err);
                showToast('Failed to read file. Make sure it\'s a valid .xlsx, .xls or .csv file.', true);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // ---- PDF handler ----
    async function handlePdfFile(file) {
        try {
            showToast('📄 Reading PDF... please wait.');

            // Wait for pdf.js to be ready
            if (!window.pdfjsLib) {
                await window.pdfjsReady;
            }
            const pdfjsLib = window.pdfjsLib;

            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            // Extract text from all pages
            let allTextItems = [];
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                allTextItems.push(...content.items);
            }

            // Strategy: Group text items by their Y position to form rows,
            // then sort items within each row by X position to form columns.
            const rows = groupTextItemsIntoRows(allTextItems);

            if (rows.length < 2) {
                showToast('Could not extract a table from this PDF. Try Excel/CSV instead.', true);
                return;
            }

            // Detect the header row (the row that best matches our field keywords)
            const { headerIdx, headers } = detectHeaderRow(rows);

            if (headerIdx === -1) {
                // Fallback: use first row as header
                importedHeaders = rows[0].map(cell => String(cell).trim());
                importedSheetData = rows.slice(1).filter(row => row.some(cell => cell !== ''));
            } else {
                importedHeaders = headers;
                importedSheetData = rows.slice(headerIdx + 1).filter(row => row.some(cell => cell !== ''));
            }

            // Normalize column count
            const colCount = importedHeaders.length;
            importedSheetData = importedSheetData.map(row => {
                const normalized = [];
                for (let i = 0; i < colCount; i++) {
                    normalized.push(row[i] !== undefined ? row[i] : '');
                }
                return normalized;
            });

            if (importedSheetData.length === 0) {
                showToast('No data rows found in the PDF.', true);
                return;
            }

            buildMappingUI();
            showPreview();
            openImportModal();
            showToast(`📄 PDF parsed — ${importedSheetData.length} rows found.`);

        } catch (err) {
            console.error('PDF import error:', err);
            showToast('Failed to parse PDF. The file may be scanned/image-based or corrupted.', true);
        }
    }

    // Group PDF text items by Y-position into rows, then sort by X within each row
    function groupTextItemsIntoRows(items) {
        if (items.length === 0) return [];

        // Each item has: str, transform[4]=x, transform[5]=y
        const yThreshold = 5; // items within 5 units of Y are same row
        const buckets = [];

        items.forEach(item => {
            const y = Math.round(item.transform[5]);
            const x = item.transform[4];
            const text = item.str;

            // Find existing bucket
            let bucket = buckets.find(b => Math.abs(b.y - y) < yThreshold);
            if (!bucket) {
                bucket = { y, cells: [] };
                buckets.push(bucket);
            }
            bucket.cells.push({ x, text });
        });

        // Sort buckets top-to-bottom (higher Y = higher on page in PDF coords)
        buckets.sort((a, b) => b.y - a.y);

        // Within each bucket, sort by X left-to-right, then merge close items
        const rows = buckets.map(bucket => {
            bucket.cells.sort((a, b) => a.x - b.x);
            return mergeCloseCells(bucket.cells);
        });

        // Filter out rows that look empty
        return rows.filter(row => row.some(cell => cell.trim() !== ''));
    }

    // Merge cells that are very close on X-axis (likely same column, split text)
    function mergeCloseCells(cells) {
        if (cells.length === 0) return [];

        const merged = [{ x: cells[0].x, text: cells[0].text }];
        const xGapThreshold = 15; // merge if gap < 15 units

        for (let i = 1; i < cells.length; i++) {
            const prev = merged[merged.length - 1];
            const curr = cells[i];

            if (curr.x - prev.x < xGapThreshold && curr.text.trim()) {
                // Same column — append text
                prev.text += ' ' + curr.text;
            } else {
                merged.push({ x: curr.x, text: curr.text });
            }
        }

        return merged.map(c => c.text.trim());
    }

    // Detect which row is the header by scoring against our known field keywords
    function detectHeaderRow(rows) {
        const allKeywords = IMPORT_FIELDS.flatMap(f => f.keywords);
        let bestIdx = -1;
        let bestScore = 0;

        const searchLimit = Math.min(rows.length, 10); // only check first 10 rows

        for (let i = 0; i < searchLimit; i++) {
            const row = rows[i];
            let score = 0;
            row.forEach(cell => {
                const lower = String(cell).toLowerCase().trim();
                if (allKeywords.some(kw => lower.includes(kw) || kw.includes(lower))) {
                    score++;
                }
            });
            if (score > bestScore) {
                bestScore = score;
                bestIdx = i;
            }
        }

        // Require at least 2 keyword matches to call it a header
        if (bestScore >= 2) {
            return { headerIdx: bestIdx, headers: rows[bestIdx].map(c => String(c).trim()) };
        }
        return { headerIdx: -1, headers: [] };
    }

    function buildMappingUI() {
        mappingGrid.innerHTML = '';

        IMPORT_FIELDS.forEach(field => {
            const item = document.createElement('div');
            item.className = 'mapping-item';

            // Auto-detect best matching column
            const autoMatch = autoDetectColumn(field.keywords);

            let options = `<option value="-1">— Skip —</option>`;
            importedHeaders.forEach((h, i) => {
                const selected = i === autoMatch ? 'selected' : '';
                options += `<option value="${i}" ${selected}>${h}</option>`;
            });

            item.innerHTML = `
                <label>${field.label}</label>
                <select data-field="${field.key}">${options}</select>
            `;

            mappingGrid.appendChild(item);
        });
    }

    function autoDetectColumn(keywords) {
        let bestIdx = -1;
        let bestScore = 0;

        importedHeaders.forEach((header, idx) => {
            const h = header.toLowerCase().trim();
            for (const kw of keywords) {
                if (h === kw) {
                    // Exact match → highest priority
                    if (10 > bestScore) {
                        bestScore = 10;
                        bestIdx = idx;
                    }
                } else if (h.includes(kw) || kw.includes(h)) {
                    if (5 > bestScore) {
                        bestScore = 5;
                        bestIdx = idx;
                    }
                }
            }
        });

        return bestIdx;
    }

    function showPreview() {
        // Show first 5 rows as preview
        const maxRows = 5;
        previewCount.textContent = importedSheetData.length;

        previewHead.innerHTML = importedHeaders.map(h => `<th>${h}</th>`).join('');

        previewBody.innerHTML = importedSheetData.slice(0, maxRows).map(row => {
            return '<tr>' + importedHeaders.map((_, i) => {
                let val = row[i] !== undefined ? row[i] : '';
                // Format dates nicely
                if (val instanceof Date) {
                    val = formatLocalDate(val);
                }
                return `<td>${String(val)}</td>`;
            }).join('') + '</tr>';
        }).join('');
    }

    function openImportModal() {
        importModal.classList.add('active');
    }

    function closeImportModal() {
        importModal.classList.remove('active');
        importedSheetData = [];
        importedHeaders = [];
    }

    modalClose.addEventListener('click', closeImportModal);
    modalCancel.addEventListener('click', closeImportModal);

    // Close modal on overlay click
    importModal.addEventListener('click', (e) => {
        if (e.target === importModal) closeImportModal();
    });

    // Escape key to close
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && importModal.classList.contains('active')) {
            closeImportModal();
        }
    });

    // ---- Execute Import ----
    modalImport.addEventListener('click', () => {
        // Read mappings
        const mappings = {};
        mappingGrid.querySelectorAll('select').forEach(sel => {
            const fieldKey = sel.dataset.field;
            const colIdx = parseInt(sel.value, 10);
            if (colIdx >= 0) {
                mappings[fieldKey] = colIdx;
            }
        });

        // Validate: at least credit or debit should be mapped
        if (mappings.credit === undefined && mappings.debit === undefined) {
            showToast('Please map at least a Credit or Debit column.', true);
            return;
        }

        let imported = 0;

        importedSheetData.forEach(row => {
            const tx = {
                id: generateId(),
                date: todayStr(),
                profile: 'OVERALL',
                description: '',
                category: '',
                credit: '',
                debit: '',
            };

            // Map date
            if (mappings.date !== undefined) {
                let rawDate = row[mappings.date];
                if (rawDate instanceof Date) {
                    tx.date = formatLocalDate(rawDate);
                } else {
                    const parsed = parseFuzzyDate(String(rawDate));
                    if (parsed) tx.date = parsed;
                }
            }

            // Map profile
            if (mappings.profile !== undefined) {
                const rawProf = String(row[mappings.profile] || '').trim().toLowerCase();
                if (rawProf.includes('suhas') || rawProf.includes('save') || rawProf.includes('savings')) {
                    tx.profile = 'SUHAS';
                }
            }

            // Map description
            if (mappings.description !== undefined) {
                tx.description = String(row[mappings.description] || '').trim();
            }

            // Map category
            if (mappings.category !== undefined) {
                const rawCat = String(row[mappings.category] || '').trim();
                // Try matching to existing categories
                const match = CATEGORIES.find(c => c.value.toLowerCase() === rawCat.toLowerCase());
                tx.category = match ? match.value : 'Other';
            }

            // Map credit
            if (mappings.credit !== undefined) {
                const val = parseNumberValue(row[mappings.credit]);
                if (val > 0) tx.credit = val;
            }

            // Map debit
            if (mappings.debit !== undefined) {
                const val = parseNumberValue(row[mappings.debit]);
                if (val > 0) tx.debit = val;
            }

            // Only import if there's some meaningful data
            if (tx.credit || tx.debit || tx.description) {
                manager.add(tx);
                imported++;
            }
        });

        closeImportModal();
        render();
        showToast(`✅ Successfully imported ${imported} transaction${imported !== 1 ? 's' : ''}!`);
    });

    // Helper: parse various date formats
    function parseFuzzyDate(str) {
        if (!str) return null;
        str = str.trim();

        // Try ISO format first (YYYY-MM-DD)
        if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
            const d = new Date(str + 'T00:00:00');
            if (!isNaN(d.getTime())) return formatLocalDate(d);
        }

        // DD/MM/YYYY or DD-MM-YYYY
        const dmyMatch = str.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
        if (dmyMatch) {
            let [, dd, mm, yyyy] = dmyMatch;
            if (yyyy.length === 2) yyyy = '20' + yyyy;
            const d = new Date(`${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}T00:00:00`);
            if (!isNaN(d.getTime())) return formatLocalDate(d);
        }

        // MM/DD/YYYY (fallback)
        const mdyMatch = str.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
        if (mdyMatch) {
            let [, mm, dd, yyyy] = mdyMatch;
            const d = new Date(`${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}T00:00:00`);
            if (!isNaN(d.getTime())) return formatLocalDate(d);
        }

        // Try native Date parse as last resort
        const d = new Date(str);
        if (!isNaN(d.getTime())) return formatLocalDate(d);

        return null;
    }

    // Helper: extract number from various formats
    function parseNumberValue(val) {
        if (typeof val === 'number') return Math.abs(val);
        const str = String(val).replace(/[₹,\s]/g, '').trim();
        const num = parseFloat(str);
        return isNaN(num) ? 0 : Math.abs(num);
    }

    // ---- Init ----
    setCurrentMonth();

    loadUIState();
    render();

    // Redraw charts on resize
    let resizeTimer;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            drawPieChart();
            drawBarChart();
        }, 250);
    });
})();
