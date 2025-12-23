// Data Structure
let budgetConfig = {
    bps: { team: 0, connectivity: 0 },
    tl: { team: 0, connectivity: 0 },
    tm: { team: 0, connectivity: 0 },
    currentMonth: ''
};

let members = [];
let expenses = [];
let attrition = [];

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    loadData();
    initializeUI();
    updateDashboard();
    renderAllTables();
    populateFilters();
    setCurrentMonth();
});

// Load data from localStorage
function loadData() {
    const savedConfig = localStorage.getItem('budgetConfig');
    const savedMembers = localStorage.getItem('members');
    const savedExpenses = localStorage.getItem('expenses');
    const savedAttrition = localStorage.getItem('attrition');
    
    if (savedConfig) budgetConfig = JSON.parse(savedConfig);
    if (savedMembers) members = JSON.parse(savedMembers);
    if (savedExpenses) expenses = JSON.parse(savedExpenses);
    if (savedAttrition) attrition = JSON.parse(savedAttrition);
}

// Save data to localStorage
function saveData() {
    localStorage.setItem('budgetConfig', JSON.stringify(budgetConfig));
    localStorage.setItem('members', JSON.stringify(members));
    localStorage.setItem('expenses', JSON.stringify(expenses));
    localStorage.setItem('attrition', JSON.stringify(attrition));
}

// Set current month
function setCurrentMonth() {
    const today = new Date();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    const currentMonthStr = `${year}-${month}`;
    
    document.getElementById('currentMonth').value = budgetConfig.currentMonth || currentMonthStr;
    if (!budgetConfig.currentMonth) {
        budgetConfig.currentMonth = currentMonthStr;
        saveData();
    }
}

// Initialize UI
function initializeUI() {
    // Tab switching
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', function() {
            const tabName = this.dataset.tab;
            switchTab(tabName);
        });
    });
    
    // Configuration modal
    document.getElementById('configBtn').addEventListener('click', openConfigModal);
    document.getElementById('saveConfigBtn').addEventListener('click', saveConfiguration);
    
    // Member modal
    document.getElementById('addMemberBtn').addEventListener('click', openMemberModal);
    document.getElementById('saveMemberBtn').addEventListener('click', saveMember);
    
    // Expense modal
    document.getElementById('addExpenseBtn').addEventListener('click', openExpenseModal);
    document.getElementById('saveExpenseBtn').addEventListener('click', saveExpense);
    
    // Attrition modal
    document.getElementById('addAttritionBtn').addEventListener('click', openAttritionModal);
    document.getElementById('saveAttritionBtn').addEventListener('click', saveAttritionEntry);
    
    // Export button
    document.getElementById('exportBtn').addEventListener('click', exportReport);
    
    // Excel upload (hidden)
    document.getElementById('uploadTrigger').addEventListener('click', function() {
        document.getElementById('excelUpload').click();
    });
    document.getElementById('excelUpload').addEventListener('change', handleExcelUpload);
    
    // Spend Theme button
    document.getElementById('spendThemeBtn').addEventListener('click', openSpendThemeModal);
    
    // Close modals
    document.querySelectorAll('.close').forEach(closeBtn => {
        closeBtn.addEventListener('click', function() {
            this.closest('.modal').classList.remove('active');
        });
    });
    
    // Close modal on outside click
    document.querySelectorAll('.modal').forEach(modal => {
        modal.addEventListener('click', function(e) {
            if (e.target === this) {
                this.classList.remove('active');
            }
        });
    });
    
    // Filters
    document.getElementById('roleFilter').addEventListener('change', renderMembersTable);
    document.getElementById('expenseTypeFilter').addEventListener('change', renderExpensesTable);
    document.getElementById('expenseMonthFilter').addEventListener('change', renderExpensesTable);
    document.getElementById('expenseMemberFilter').addEventListener('change', renderExpensesTable);
}

// Switch tabs
function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(`${tabName}Tab`).classList.add('active');
}

// Configuration Modal
function openConfigModal() {
    document.getElementById('bpsTeamBudget').value = budgetConfig.bps.team;
    document.getElementById('bpsConnectivityBudget').value = budgetConfig.bps.connectivity;
    document.getElementById('tlTeamBudget').value = budgetConfig.tl.team;
    document.getElementById('tlConnectivityBudget').value = budgetConfig.tl.connectivity;
    document.getElementById('tmTeamBudget').value = budgetConfig.tm.team;
    document.getElementById('tmConnectivityBudget').value = budgetConfig.tm.connectivity;
    
    document.getElementById('configModal').classList.add('active');
}

function saveConfiguration() {
    budgetConfig.bps.team = parseFloat(document.getElementById('bpsTeamBudget').value) || 0;
    budgetConfig.bps.connectivity = parseFloat(document.getElementById('bpsConnectivityBudget').value) || 0;
    budgetConfig.tl.team = parseFloat(document.getElementById('tlTeamBudget').value) || 0;
    budgetConfig.tl.connectivity = parseFloat(document.getElementById('tlConnectivityBudget').value) || 0;
    budgetConfig.tm.team = parseFloat(document.getElementById('tmTeamBudget').value) || 0;
    budgetConfig.tm.connectivity = parseFloat(document.getElementById('tmConnectivityBudget').value) || 0;
    budgetConfig.currentMonth = document.getElementById('currentMonth').value;
    
    saveData();
    updateDashboard();
    renderMembersTable();
    document.getElementById('configModal').classList.remove('active');
}

// Member Management
let editingMemberId = null;

function openMemberModal(memberId = null) {
    editingMemberId = memberId;
    
    if (memberId) {
        const member = members.find(m => m.id === memberId);
        document.getElementById('memberModalTitle').textContent = 'Edit Team Member';
        document.getElementById('memberName').value = member.name;
        document.getElementById('memberRole').value = member.role;
        document.getElementById('memberTL').value = member.tl || '';
    } else {
        document.getElementById('memberModalTitle').textContent = 'Add Team Member';
        document.getElementById('memberName').value = '';
        document.getElementById('memberRole').value = 'BPS';
        document.getElementById('memberTL').value = '';
    }
    
    document.getElementById('memberModal').classList.add('active');
}

function saveMember() {
    const name = document.getElementById('memberName').value.trim();
    const role = document.getElementById('memberRole').value;
    const tl = document.getElementById('memberTL').value.trim();
    
    if (!name) {
        alert('Please enter member name');
        return;
    }
    
    if (editingMemberId) {
        const member = members.find(m => m.id === editingMemberId);
        member.name = name;
        member.role = role;
        member.tl = tl;
        member.teamBudget = budgetConfig[role.toLowerCase()].team;
        member.connectivityBudget = budgetConfig[role.toLowerCase()].connectivity;
    } else {
        members.push({
            id: Date.now().toString(),
            name: name,
            role: role,
            tl: tl,
            teamBudget: budgetConfig[role.toLowerCase()].team,
            connectivityBudget: budgetConfig[role.toLowerCase()].connectivity,
            status: 'active'
        });
    }
    
    saveData();
    updateDashboard();
    renderMembersTable();
    populateFilters();
    document.getElementById('memberModal').classList.remove('active');
}

function deleteMember(memberId) {
    if (confirm('Are you sure you want to delete this member?')) {
        members = members.filter(m => m.id !== memberId);
        expenses = expenses.filter(e => e.memberId !== memberId);
        attrition = attrition.filter(a => a.memberId !== memberId);
        
        saveData();
        updateDashboard();
        renderAllTables();
        populateFilters();
    }
}

// Expense Management
let editingExpenseId = null;

function openExpenseModal(expenseId = null) {
    editingExpenseId = expenseId;
    populateMemberDropdown('expenseMember');
    
    if (expenseId) {
        const expense = expenses.find(e => e.id === expenseId);
        document.getElementById('expenseName').value = expense.name;
        document.getElementById('expenseAmount').value = expense.amount;
        document.getElementById('expenseEvent').value = expense.event;
        document.getElementById('expenseCategory').value = expense.category;
        document.getElementById('expenseMember').value = expense.memberId || '';
        document.getElementById('expenseDate').value = expense.date;
    } else {
        document.getElementById('expenseName').value = '';
        document.getElementById('expenseAmount').value = '';
        document.getElementById('expenseEvent').value = '';
        document.getElementById('expenseCategory').value = 'team';
        document.getElementById('expenseMember').value = '';
        
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('expenseDate').value = today;
    }
    
    document.getElementById('expenseModal').classList.add('active');
}

function saveExpense() {
    const name = document.getElementById('expenseName').value.trim();
    const amount = parseFloat(document.getElementById('expenseAmount').value);
    const event = document.getElementById('expenseEvent').value.trim();
    const category = document.getElementById('expenseCategory').value;
    const memberId = document.getElementById('expenseMember').value;
    const date = document.getElementById('expenseDate').value;
    
    if (!name || !amount || !event || !date) {
        alert('Please fill in all required fields');
        return;
    }
    
    if (editingExpenseId) {
        const expense = expenses.find(e => e.id === editingExpenseId);
        expense.name = name;
        expense.amount = amount;
        expense.event = event;
        expense.category = category;
        expense.memberId = memberId;
        expense.date = date;
    } else {
        expenses.push({
            id: Date.now().toString(),
            name: name,
            amount: amount,
            event: event,
            category: category,
            memberId: memberId,
            date: date
        });
    }
    
    saveData();
    updateDashboard();
    renderExpensesTable();
    document.getElementById('expenseModal').classList.remove('active');
}

function deleteExpense(expenseId) {
    if (confirm('Are you sure you want to delete this expense?')) {
        expenses = expenses.filter(e => e.id !== expenseId);
        saveData();
        updateDashboard();
        renderExpensesTable();
    }
}

// Attrition Management
function openAttritionModal() {
    populateMemberDropdown('attritionMember', true);
    
    const today = new Date();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    document.getElementById('attritionMonth').value = `${year}-${month}`;
    
    document.getElementById('attritionModal').classList.add('active');
}

function saveAttritionEntry() {
    const memberId = document.getElementById('attritionMember').value;
    const exitMonth = document.getElementById('attritionMonth').value;
    
    if (!memberId || !exitMonth) {
        alert('Please select member and exit month');
        return;
    }
    
    // Check if already exists
    if (attrition.find(a => a.memberId === memberId)) {
        alert('This member already has an exit entry');
        return;
    }
    
    const member = members.find(m => m.id === memberId);
    member.status = 'exited';
    
    attrition.push({
        id: Date.now().toString(),
        memberId: memberId,
        exitMonth: exitMonth
    });
    
    saveData();
    updateDashboard();
    renderAllTables();
    populateFilters();
    document.getElementById('attritionModal').classList.remove('active');
}

function deleteAttritionEntry(attritionId) {
    if (confirm('Are you sure you want to remove this exit entry?')) {
        const entry = attrition.find(a => a.id === attritionId);
        const member = members.find(m => m.id === entry.memberId);
        member.status = 'active';
        
        attrition = attrition.filter(a => a.id !== attritionId);
        
        saveData();
        updateDashboard();
        renderAllTables();
    }
}

// Dashboard Updates
function updateDashboard() {
    const totalTeamBudget = calculateTotalBudget('team');
    const totalConnectivityBudget = calculateTotalBudget('connectivity');
    
    const teamSpent = calculateSpent('team');
    const connectivitySpent = calculateSpent('connectivity');
    
    const teamRemaining = totalTeamBudget - teamSpent;
    const connectivityRemaining = totalConnectivityBudget - connectivitySpent;
    
    const teamUtilization = totalTeamBudget > 0 ? (teamSpent / totalTeamBudget * 100) : 0;
    const connectivityUtilization = totalConnectivityBudget > 0 ? (connectivitySpent / totalConnectivityBudget * 100) : 0;
    
    // Update Team Budget
    document.getElementById('totalTeamBudget').textContent = formatCurrency(totalTeamBudget);
    document.getElementById('teamBudgetSpent').textContent = formatCurrency(teamSpent);
    document.getElementById('teamBudgetRemaining').textContent = formatCurrency(teamRemaining);
    document.getElementById('teamBudgetProgress').style.width = teamUtilization + '%';
    document.getElementById('teamBudgetUtilization').textContent = teamUtilization.toFixed(1) + '% utilized';
    
    // Update Connectivity Budget
    document.getElementById('totalConnectivityBudget').textContent = formatCurrency(totalConnectivityBudget);
    document.getElementById('connectivityBudgetSpent').textContent = formatCurrency(connectivitySpent);
    document.getElementById('connectivityBudgetRemaining').textContent = formatCurrency(connectivityRemaining);
    document.getElementById('connectivityBudgetProgress').style.width = connectivityUtilization + '%';
    document.getElementById('connectivityBudgetUtilization').textContent = connectivityUtilization.toFixed(1) + '% utilized';
}

function calculateTotalBudget(category) {
    let total = 0;
    
    members.forEach(member => {
        if (member.status === 'active') {
            const budget = category === 'team' ? member.teamBudget : member.connectivityBudget;
            total += budget;
        } else {
            // Calculate pro-rated budget for exited members
            const exitEntry = attrition.find(a => a.memberId === member.id);
            if (exitEntry) {
                const exitDate = new Date(exitEntry.exitMonth + '-01');
                const currentDate = new Date(budgetConfig.currentMonth + '-01');
                
                if (exitDate >= currentDate) {
                    const budget = category === 'team' ? member.teamBudget : member.connectivityBudget;
                    total += budget;
                }
            }
        }
    });
    
    return total;
}

function calculateSpent(category) {
    return expenses
        .filter(e => e.category === category)
        .reduce((sum, e) => sum + e.amount, 0);
}

// Render Tables
function renderAllTables() {
    renderMembersTable();
    renderExpensesTable();
    renderAttritionTable();
}

function renderMembersTable() {
    const roleFilter = document.getElementById('roleFilter').value;
    const tbody = document.getElementById('membersTableBody');
    
    let filteredMembers = members;
    if (roleFilter !== 'all') {
        filteredMembers = members.filter(m => m.role === roleFilter);
    }
    
    if (filteredMembers.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state"><div class="empty-state-text">No members found</div><div class="empty-state-subtext">Add team members to get started</div></td></tr>';
        return;
    }
    
    tbody.innerHTML = filteredMembers.map(member => `
        <tr>
            <td>${member.name}</td>
            <td><span class="status-badge ${member.role === 'BPS' ? 'category-team' : member.role === 'TL' ? 'category-connectivity' : 'status-badge'}">${member.role}</span></td>
            <td>${member.tl || '-'}</td>
            <td>${formatCurrency(member.teamBudget)}</td>
            <td>${formatCurrency(member.connectivityBudget)}</td>
            <td><span class="status-badge ${member.status === 'active' ? 'status-active' : 'status-exited'}">${member.status}</span></td>
            <td>
                <div class="action-btns">
                    <button class="action-btn" onclick="openMemberModal('${member.id}')">Edit</button>
                    <button class="action-btn" onclick="deleteMember('${member.id}')">Delete</button>
                </div>
            </td>
        </tr>
    `).join('');
}

function renderExpensesTable() {
    const typeFilter = document.getElementById('expenseTypeFilter').value;
    const monthFilter = document.getElementById('expenseMonthFilter').value;
    const memberFilter = document.getElementById('expenseMemberFilter').value;
    const tbody = document.getElementById('expensesTableBody');
    
    let filteredExpenses = expenses;
    
    if (typeFilter !== 'all') {
        filteredExpenses = filteredExpenses.filter(e => e.category === typeFilter);
    }
    
    if (monthFilter !== 'all') {
        filteredExpenses = filteredExpenses.filter(e => {
            const expenseMonth = e.date.substring(0, 7);
            return expenseMonth === monthFilter;
        });
    }
    
    if (memberFilter !== 'all') {
        filteredExpenses = filteredExpenses.filter(e => e.memberId === memberFilter);
    }
    
    if (filteredExpenses.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state"><div class="empty-state-text">No expenses found</div><div class="empty-state-subtext">Add expenses to track spending</div></td></tr>';
        return;
    }
    
    tbody.innerHTML = filteredExpenses.map(expense => {
        const member = expense.memberId ? members.find(m => m.id === expense.memberId) : null;
        const memberName = member ? member.name : 'General';
        
        return `
            <tr>
                <td>${formatDate(expense.date)}</td>
                <td>${expense.name}</td>
                <td>${formatCurrency(expense.amount)}</td>
                <td>${expense.event}</td>
                <td><span class="category-badge category-${expense.category}">${expense.category === 'team' ? 'Team Budget' : 'Connectivity Budget'}</span></td>
                <td>${memberName}</td>
                <td>
                    <div class="action-btns">
                        <button class="action-btn" onclick="openExpenseModal('${expense.id}')">Edit</button>
                        <button class="action-btn" onclick="deleteExpense('${expense.id}')">Delete</button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderAttritionTable() {
    const tbody = document.getElementById('attritionTableBody');
    
    if (attrition.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state"><div class="empty-state-text">No exits recorded</div><div class="empty-state-subtext">Track team member attrition here</div></td></tr>';
        return;
    }
    
    tbody.innerHTML = attrition.map(entry => {
        const member = members.find(m => m.id === entry.memberId);
        const budgetImpact = member.teamBudget + member.connectivityBudget;
        
        return `
            <tr>
                <td>${member.name}</td>
                <td><span class="status-badge category-${member.role === 'BPS' ? 'team' : 'connectivity'}">${member.role}</span></td>
                <td>${formatMonth(entry.exitMonth)}</td>
                <td>${formatCurrency(budgetImpact)}</td>
                <td>
                    <div class="action-btns">
                        <button class="action-btn" onclick="deleteAttritionEntry('${entry.id}')">Remove</button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// Populate Filters
function populateFilters() {
    // Populate month filter for expenses
    const months = [...new Set(expenses.map(e => e.date.substring(0, 7)))].sort().reverse();
    const monthFilter = document.getElementById('expenseMonthFilter');
    monthFilter.innerHTML = '<option value="all">All Months</option>' + 
        months.map(m => `<option value="${m}">${formatMonth(m)}</option>`).join('');
    
    // Populate member filter for expenses
    populateMemberDropdown('expenseMemberFilter', false, true);
}

function populateMemberDropdown(elementId, activeOnly = false, includeAll = false) {
    const select = document.getElementById(elementId);
    let filteredMembers = activeOnly ? members.filter(m => m.status === 'active') : members;
    
    let options = includeAll ? '<option value="all">All Members</option>' : '<option value="">Select member (optional)</option>';
    options += filteredMembers.map(m => `<option value="${m.id}">${m.name} (${m.role})</option>`).join('');
    
    select.innerHTML = options;
}

// Export Report
function exportReport() {
    const totalTeamBudget = calculateTotalBudget('team');
    const totalConnectivityBudget = calculateTotalBudget('connectivity');
    const teamSpent = calculateSpent('team');
    const connectivitySpent = calculateSpent('connectivity');
    
    let report = 'BUDGET TRACKING REPORT\n';
    report += '='.repeat(80) + '\n\n';
    
    report += 'BUDGET OVERVIEW\n';
    report += '-'.repeat(80) + '\n';
    report += `Team Budget: ${formatCurrency(totalTeamBudget)} | Spent: ${formatCurrency(teamSpent)} | Remaining: ${formatCurrency(totalTeamBudget - teamSpent)}\n`;
    report += `Connectivity Budget: ${formatCurrency(totalConnectivityBudget)} | Spent: ${formatCurrency(connectivitySpent)} | Remaining: ${formatCurrency(totalConnectivityBudget - connectivitySpent)}\n\n`;
    
    report += 'TEAM MEMBERS\n';
    report += '-'.repeat(80) + '\n';
    members.forEach(m => {
        report += `${m.name} (${m.role}) - Team: ${formatCurrency(m.teamBudget)} | Connectivity: ${formatCurrency(m.connectivityBudget)} | Status: ${m.status}\n`;
    });
    
    report += '\n\nEXPENSES\n';
    report += '-'.repeat(80) + '\n';
    expenses.forEach(e => {
        const member = e.memberId ? members.find(m => m.id === e.memberId)?.name : 'General';
        report += `${formatDate(e.date)} | ${e.name} | ${formatCurrency(e.amount)} | ${e.event} | ${e.category} | ${member}\n`;
    });
    
    if (attrition.length > 0) {
        report += '\n\nATTRITION\n';
        report += '-'.repeat(80) + '\n';
        attrition.forEach(a => {
            const member = members.find(m => m.id === a.memberId);
            report += `${member.name} (${member.role}) - Exit Month: ${formatMonth(a.exitMonth)}\n`;
        });
    }
    
    // Download as text file
    const blob = new Blob([report], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `budget-report-${new Date().toISOString().split('T')[0]}.txt`;
    a.click();
    URL.revokeObjectURL(url);
}

// Excel Upload Handler
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            parseExcelData(workbook);
            
            // Show success message
            alert('Excel data imported successfully!');
            
            // Refresh all views
            updateDashboard();
            renderAllTables();
            populateFilters();
            
            // Reset file input
            event.target.value = '';
        } catch (error) {
            console.error('Error parsing Excel:', error);
            alert('Error importing Excel file. Please check the format.');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseExcelData(workbook) {
    // Parse 2025 sheet for team members
    if (workbook.SheetNames.includes('2025')) {
        const sheet = workbook.Sheets['2025'];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        // Clear existing members
        members = [];
        
        // Parse members starting from row 3 (index 2)
        for (let i = 2; i < data.length; i++) {
            const row = data[i];
            const name = row[4]; // Column E - Name
            const role = row[5]; // Column F - Role
            const tl = row[6];   // Column G - TL
            
            if (!name || name === '' || typeof name !== 'string') continue;
            if (name.toLowerCase().includes('total') || name.toLowerCase().includes('budget')) continue;
            
            // Calculate team budget from monthly data (columns H-S for months)
            let monthlyBudget = 0;
            let monthCount = 0;
            for (let col = 7; col <= 18; col++) { // H to S
                const val = row[col];
                if (val && val !== '-' && typeof val === 'number' && val > 0) {
                    monthlyBudget += val;
                    monthCount++;
                }
            }
            
            // Default budgets if not in data
            let teamBudget = 12000; // Default per person annual
            let connectivityBudget = 4233; // Default per person annual (83*51)
            
            // Map role
            let mappedRole = 'BPS';
            if (role && typeof role === 'string') {
                if (role.toUpperCase().includes('TL') || role.toUpperCase().includes('TEAM LEADER')) {
                    mappedRole = 'TL';
                } else if (role.toUpperCase().includes('TM') || role.toUpperCase().includes('TEAM MANAGER')) {
                    mappedRole = 'TM';
                } else if (role.toUpperCase() === 'PC' || role.toUpperCase() === 'BPS') {
                    mappedRole = 'BPS';
                }
            }
            
            members.push({
                id: Date.now().toString() + '_' + i,
                name: name.trim(),
                role: mappedRole,
                tl: tl || '',
                teamBudget: teamBudget,
                connectivityBudget: connectivityBudget,
                status: 'active'
            });
        }
    }
    
    // Parse Expenses sheet
    if (workbook.SheetNames.includes('Expenses')) {
        const sheet = workbook.Sheets['Expenses'];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        // Clear existing expenses
        expenses = [];
        
        // Parse expense columns - this is complex due to the multi-column layout
        // We'll look for patterns like "BPS", "Amount" headers and parse accordingly
        parseExpenseColumns(data);
    }
    
    // Parse TL Connect Budget sheet
    if (workbook.SheetNames.includes('TL Connect Budget')) {
        const sheet = workbook.Sheets['TL Connect Budget'];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        // Update TL budgets based on this data
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const tlName = row[0];
            const budget = row[5]; // Budget Total column
            
            if (tlName && typeof tlName === 'string' && budget && typeof budget === 'number') {
                // Find TL in members and update budget
                const tl = members.find(m => m.name === tlName && m.role === 'TL');
                if (tl) {
                    tl.teamBudget = budget;
                }
            }
        }
    }
    
    saveData();
}

function parseExpenseColumns(data) {
    // Look for expense headers in row 2 (index 1)
    if (data.length < 3) return;
    
    const headerRow = data[1];
    const dataStartRow = 3; // Row 4 in Excel (index 3)
    
    // Find expense columns by looking for patterns
    for (let col = 0; col < headerRow.length; col++) {
        const header = headerRow[col];
        if (header && typeof header === 'string' && header.toLowerCase().includes('expense')) {
            // Found an expense column, parse it
            const expenseName = header;
            
            // Look for total in row 1
            const totalRow = data[0];
            const total = totalRow[col + 1]; // Amount column is usually next
            
            // Determine month from the header
            let month = 'JAN';
            if (header.toLowerCase().includes('feb')) month = 'FEB';
            else if (header.toLowerCase().includes('mar')) month = 'MAR';
            else if (header.toLowerCase().includes('apr')) month = 'APR';
            else if (header.toLowerCase().includes('may')) month = 'MAY';
            else if (header.toLowerCase().includes('jun')) month = 'JUN';
            else if (header.toLowerCase().includes('jul')) month = 'JUL';
            else if (header.toLowerCase().includes('aug')) month = 'AUG';
            else if (header.toLowerCase().includes('sep')) month = 'SEP';
            else if (header.toLowerCase().includes('oct')) month = 'OCT';
            else if (header.toLowerCase().includes('nov')) month = 'NOV';
            else if (header.toLowerCase().includes('dec')) month = 'DEC';
            
            // Determine category from expense name
            let category = 'team';
            if (header.toLowerCase().includes('connectivity')) {
                category = 'connectivity';
            }
            
            // Only add if we have a valid total
            if (total && typeof total === 'number' && total > 0) {
                const date = getDateFromMonth(month);
                
                expenses.push({
                    id: Date.now().toString() + '_' + col,
                    name: expenseName.replace(' - Expense', '').trim(),
                    amount: total,
                    event: expenseName.replace(' - Expense', '').trim(),
                    category: category,
                    memberId: '',
                    date: date
                });
            }
        }
    }
}

function getDateFromMonth(month) {
    const monthMap = {
        'JAN': '2025-01-15',
        'FEB': '2025-02-15',
        'MAR': '2025-03-15',
        'APR': '2025-04-15',
        'MAY': '2025-05-15',
        'JUN': '2025-06-15',
        'JUL': '2025-07-15',
        'AUG': '2025-08-15',
        'SEP': '2025-09-15',
        'OCT': '2025-10-15',
        'NOV': '2025-11-15',
        'DEC': '2025-12-15'
    };
    return monthMap[month] || '2025-01-15';
}

// Spend Theme Visualization
let spendChart = null;

function openSpendThemeModal() {
    populateTLFilter();
    renderSpendChart();
    document.getElementById('spendThemeModal').classList.add('active');
}

function populateTLFilter() {
    // Get unique TLs from members
    const tls = [...new Set(members.filter(m => m.role === 'TL').map(m => m.name))];
    const select = document.getElementById('chartTLFilter');
    
    let options = '<option value="all">All Team Leaders</option>';
    options += tls.map(tl => `<option value="${tl}">${tl}</option>`).join('');
    
    select.innerHTML = options;
    
    // Add change listener
    select.onchange = renderSpendChart;
    document.getElementById('chartMonthRangeFilter').onchange = renderSpendChart;
}

function renderSpendChart() {
    const tlFilter = document.getElementById('chartTLFilter').value;
    const rangeFilter = document.getElementById('chartMonthRangeFilter').value;
    
    // Get monthly expense data
    const monthlyData = getMonthlyExpenseData(tlFilter, rangeFilter);
    
    // Destroy existing chart
    if (spendChart) {
        spendChart.destroy();
    }
    
    // Create new chart
    const ctx = document.getElementById('spendChart').getContext('2d');
    spendChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: monthlyData.labels,
            datasets: [
                {
                    label: 'Team Budget',
                    data: monthlyData.teamBudget,
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 5,
                    pointHoverRadius: 7
                },
                {
                    label: 'Connectivity Budget',
                    data: monthlyData.connectivityBudget,
                    borderColor: '#f59e0b',
                    backgroundColor: 'rgba(245, 158, 11, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 5,
                    pointHoverRadius: 7
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 13,
                            weight: 600
                        },
                        color: '#1f2937',
                        padding: 15
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleFont: {
                        size: 14,
                        weight: 600
                    },
                    bodyFont: {
                        size: 13
                    },
                    padding: 12,
                    displayColors: true,
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ₹' + context.parsed.y.toLocaleString('en-IN');
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: '#e5e7eb'
                    },
                    ticks: {
                        font: {
                            size: 12
                        },
                        color: '#6b7280',
                        callback: function(value) {
                            return '₹' + (value / 1000) + 'K';
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        font: {
                            size: 12,
                            weight: 600
                        },
                        color: '#1f2937'
                    }
                }
            }
        }
    });
    
    // Update summary
    updateMonthlySummary(monthlyData);
}

function getMonthlyExpenseData(tlFilter, rangeFilter) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthNums = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    
    // Filter months based on range
    let filteredMonths = months;
    let filteredMonthNums = monthNums;
    
    if (rangeFilter === 'q1') {
        filteredMonths = months.slice(0, 3);
        filteredMonthNums = monthNums.slice(0, 3);
    } else if (rangeFilter === 'q2') {
        filteredMonths = months.slice(3, 6);
        filteredMonthNums = monthNums.slice(3, 6);
    } else if (rangeFilter === 'q3') {
        filteredMonths = months.slice(6, 9);
        filteredMonthNums = monthNums.slice(6, 9);
    } else if (rangeFilter === 'q4') {
        filteredMonths = months.slice(9, 12);
        filteredMonthNums = monthNums.slice(9, 12);
    }
    
    const teamBudgetData = [];
    const connectivityBudgetData = [];
    
    filteredMonthNums.forEach(monthNum => {
        let teamTotal = 0;
        let connectivityTotal = 0;
        
        expenses.forEach(expense => {
            const expenseMonth = expense.date.substring(5, 7);
            
            if (expenseMonth === monthNum) {
                // Check if expense matches TL filter
                if (tlFilter !== 'all') {
                    // Find member associated with expense
                    if (expense.memberId) {
                        const member = members.find(m => m.id === expense.memberId);
                        if (member && member.tl !== tlFilter) {
                            return; // Skip this expense
                        }
                    } else {
                        // Check if expense name contains TL name
                        if (!expense.name.toLowerCase().includes(tlFilter.toLowerCase())) {
                            return;
                        }
                    }
                }
                
                if (expense.category === 'team') {
                    teamTotal += expense.amount;
                } else {
                    connectivityTotal += expense.amount;
                }
            }
        });
        
        teamBudgetData.push(teamTotal);
        connectivityBudgetData.push(connectivityTotal);
    });
    
    return {
        labels: filteredMonths,
        teamBudget: teamBudgetData,
        connectivityBudget: connectivityBudgetData,
        months: filteredMonthNums
    };
}

function updateMonthlySummary(data) {
    const summaryDiv = document.getElementById('monthlySummary');
    
    let html = '';
    data.labels.forEach((month, index) => {
        const teamAmount = data.teamBudget[index];
        const connectivityAmount = data.connectivityBudget[index];
        const total = teamAmount + connectivityAmount;
        
        html += `
            <div style="background: white; padding: 15px; border-radius: 6px; border-left: 4px solid #3b82f6;">
                <div style="font-size: 14px; font-weight: 700; color: #1f2937; margin-bottom: 10px;">${month}</div>
                <div style="font-size: 12px; color: #6b7280; margin-bottom: 5px;">
                    Team: <span style="font-weight: 600; color: #1f2937;">₹${teamAmount.toLocaleString('en-IN')}</span>
                </div>
                <div style="font-size: 12px; color: #6b7280; margin-bottom: 5px;">
                    Connectivity: <span style="font-weight: 600; color: #1f2937;">₹${connectivityAmount.toLocaleString('en-IN')}</span>
                </div>
                <div style="font-size: 13px; font-weight: 700; color: #3b82f6; margin-top: 8px; padding-top: 8px; border-top: 1px solid #e5e7eb;">
                    Total: ₹${total.toLocaleString('en-IN')}
                </div>
            </div>
        `;
    });
    
    summaryDiv.innerHTML = html || '<div style="text-align: center; color: #6b7280;">No expense data for selected period</div>';
}

// Utility Functions
function formatCurrency(amount) {
    return '₹' + amount.toLocaleString('en-IN');
}

function formatDate(dateStr) {
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

function formatMonth(monthStr) {
    const [year, month] = monthStr.split('-');
    const date = new Date(year, month - 1);
    return date.toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
}
