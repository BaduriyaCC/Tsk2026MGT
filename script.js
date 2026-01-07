/**
 * Task Management System 2026
 * Baduriya Central College - Mawanella
 * 
 * Main Application Logic
 */

// --- STATE MANAGEMENT ---
const App = {
    data: {
        staff: [],
        tasks: []
    },
    config: {
        collegeName: "Baduriya Central College",
        subheading: "Mawanella"
    },

    // Initialize Application
    init() {
        this.loadData();
        this.setupEventListeners();
        this.renderDashboard();
        this.updateCurrentDate();

        // Render initial view
        this.renderStaffTable();
        this.renderTaskTable(); // Setup for later
        this.populateStaffDropdowns();
    },

    // --- DATA HANDLING ---
    loadData() {
        const storedStaff = localStorage.getItem('tms_staff');
        const storedTasks = localStorage.getItem('tms_tasks');

        if (storedStaff) this.data.staff = JSON.parse(storedStaff);
        if (storedTasks) this.data.tasks = JSON.parse(storedTasks);
    },

    saveData() {
        localStorage.setItem('tms_staff', JSON.stringify(this.data.staff));
        localStorage.setItem('tms_tasks', JSON.stringify(this.data.tasks));
        this.renderDashboard(); // Update stats whenever data changes
    },

    // --- UI HELPERS ---
    updateCurrentDate() {
        const dateEl = document.getElementById('current-date');
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        dateEl.textContent = new Date().toLocaleDateString('en-US', options);

        // Set default date for task input
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('task-date').value = today;
    },

    refreshAll() {
        this.renderStaffTable();
        this.renderTaskTable();
        this.renderDashboard();
        this.populateStaffDropdowns();
    },

    // --- DASHBOARD ---
    renderDashboard() {
        // Update counters
        document.getElementById('stat-total-staff').textContent = this.data.staff.length;

        const pending = this.data.tasks.filter(t => t.status === 'pending').length;
        const completed = this.data.tasks.filter(t => t.status === 'completed').length;

        document.getElementById('stat-pending-tasks').textContent = pending;
        document.getElementById('stat-completed-tasks').textContent = completed;

        // Render Chart
        this.renderChart(pending, completed);
    },

    renderChart(pending, completed) {
        const ctx = document.getElementById('overviewChart');
        if (!window.Chart) {
            ctx.parentNode.innerHTML = '<p style="text-align:center; padding: 2rem; color: #64748B;">Charts library not loaded (Offline Mode without dependencies). Stats are displayed above.</p>';
            return;
        }

        // Destroy existing chart if any
        if (window.myChart) window.myChart.destroy();

        window.myChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Pending', 'Completed'],
                datasets: [{
                    data: [pending, completed],
                    backgroundColor: ['#F59E0B', '#10B981'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom'
                    }
                },
                cutout: '70%'
            }
        });
    },

    // --- STAFF MANAGEMENT ---
    renderStaffTable() {
        const tbody = document.querySelector('#staff-table tbody');
        const search = document.getElementById('staff-search').value.toLowerCase();

        tbody.innerHTML = '';

        const filtered = this.data.staff.filter(s =>
            s.name.toLowerCase().includes(search) ||
            s.no.toLowerCase().includes(search)
        );

        filtered.forEach(staff => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${staff.no}</strong></td>
                <td>${staff.name}</td>
                <td>
                    <button class="icon-btn" onclick="App.editStaff('${staff.id}')"><i class="ph ph-pencil-simple"></i></button>
                    <button class="icon-btn delete" onclick="App.deleteStaff('${staff.id}')"><i class="ph ph-trash"></i></button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    },

    addStaff(staff) {
        this.data.staff.push({ ...staff, id: Date.now().toString() });
        this.saveData();
        this.refreshAll();
    },

    updateStaff(id, updatedData) {
        const index = this.data.staff.findIndex(s => s.id === id);
        if (index > -1) {
            this.data.staff[index] = { ...this.data.staff[index], ...updatedData };
            this.saveData();
            this.refreshAll();
        }
    },

    deleteStaff(id) {
        if (confirm('Are you sure you want to delete this staff member?')) {
            this.data.staff = this.data.staff.filter(s => s.id !== id);
            // Optional: Delete associated tasks? keeping them for now but maybe warn.
            this.saveData();
            this.refreshAll();
        }
    },

    editStaff(id) {
        const staff = this.data.staff.find(s => s.id === id);
        if (staff) {
            document.getElementById('staff-id-hidden').value = staff.id;
            document.getElementById('staff-no').value = staff.no;
            document.getElementById('staff-name').value = staff.name;
            document.getElementById('modal-staff-title').textContent = 'Edit Staff';
            document.getElementById('modal-staff').classList.remove('hidden');
            document.getElementById('modal-overlay').classList.remove('hidden');
        }
    },

    populateStaffDropdowns() {
        const select = document.getElementById('task-staff-select');
        const reportSelect = document.getElementById('report-staff-select');

        // Keep first option
        select.innerHTML = '<option value="">Select Staff...</option>';
        reportSelect.innerHTML = '<option value="all">All Staff</option>';

        this.data.staff.forEach(s => {
            const opt = `<option value="${s.id}">${s.no} - ${s.name}</option>`;
            select.innerHTML += opt;
            reportSelect.innerHTML += opt;
        });
    },

    handleStaffUpload(file) {
        if (!window.XLSX) {
            alert('Excel library (SheetJS) not loaded.');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            // Assuming Row 1 is header. Format: [Staff No, Staff Name]
            let count = 0;
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row[0] && row[1]) {
                    this.data.staff.push({
                        id: Date.now().toString() + i, // Unique ID
                        no: row[0].toString(),
                        name: row[1].toString()
                    });
                    count++;
                }
            }
            this.saveData();
            this.refreshAll();
            alert(`Succesfully imported ${count} staff profiles.`);
        };
        reader.readAsArrayBuffer(file);
    },

    // --- TASK MANAGEMENT ---
    renderTaskTable() {
        const tbody = document.querySelector('#task-table tbody');
        const search = document.getElementById('task-search').value.toLowerCase();

        tbody.innerHTML = '';

        // Join with staff data for display
        const enrichedTasks = this.data.tasks.map(t => {
            const staff = this.data.staff.find(s => s.id === t.staffId);
            return { ...t, staffName: staff ? staff.name : 'Unknown' };
        });

        const filtered = enrichedTasks.filter(t =>
            t.title.toLowerCase().includes(search) ||
            t.staffName.toLowerCase().includes(search)
        );

        filtered.forEach(task => {
            const tr = document.createElement('tr');
            const statusClass = task.status === 'completed' ? 'completed' : 'pending';
            const statusText = task.status === 'completed' ? 'Completed' : 'Pending';

            tr.innerHTML = `
                <td>${task.date}</td>
                <td>${task.staffName}</td>
                <td>${task.title}</td>
                <td>${task.dueDate || '-'}</td>
                <td><span class="badge ${statusClass}">${statusText}</span></td>
                <td>
                    <button class="icon-btn" title="Toggle Status" onclick="App.toggleTaskStatus('${task.id}')">
                        <i class="ph ${task.status === 'completed' ? 'ph-arrow-counter-clockwise' : 'ph-check'}"></i>
                    </button>
                    <button class="icon-btn" onclick="App.editTask('${task.id}')"><i class="ph ph-pencil-simple"></i></button>
                    <button class="icon-btn delete" onclick="App.deleteTask('${task.id}')"><i class="ph ph-trash"></i></button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    },

    addTask(task) {
        this.data.tasks.push({
            ...task,
            id: Date.now().toString(),
            status: 'pending'
        });
        this.saveData();
        this.refreshAll();
    },

    updateTask(id, updatedData) {
        const index = this.data.tasks.findIndex(t => t.id === id);
        if (index > -1) {
            this.data.tasks[index] = { ...this.data.tasks[index], ...updatedData };
            this.saveData();
            this.refreshAll();
        }
    },

    deleteTask(id) {
        if (confirm('Delete this task?')) {
            this.data.tasks = this.data.tasks.filter(t => t.id !== id);
            this.saveData();
            this.refreshAll();
        }
    },

    toggleTaskStatus(id) {
        const task = this.data.tasks.find(t => t.id === id);
        if (task) {
            task.status = task.status === 'pending' ? 'completed' : 'pending';
            this.saveData();
            this.refreshAll();
        }
    },

    editTask(id) {
        const task = this.data.tasks.find(t => t.id === id);
        if (task) {
            document.getElementById('task-id-hidden').value = task.id;
            document.getElementById('task-date').value = task.date;
            document.getElementById('task-staff-select').value = task.staffId;
            document.getElementById('task-title').value = task.title;
            document.getElementById('task-desc').value = task.description || '';
            document.getElementById('task-due-date').value = task.dueDate || '';

            document.getElementById('modal-task-title').textContent = 'Edit Task';
            document.getElementById('modal-task').classList.remove('hidden');
            document.getElementById('modal-overlay').classList.remove('hidden');
        }
    },

    // --- REPORTS ---
    generateReport() {
        if (!window.jspdf || !window.jspdf.jsPDF) {
            alert('PDF library not loaded.');
            return;
        }

        const staffId = document.getElementById('report-staff-select').value;
        let tasksToReport = this.data.tasks;
        let titleSuffix = "All Staff";

        if (staffId !== 'all') {
            tasksToReport = tasksToReport.filter(t => t.staffId === staffId);
            const staff = this.data.staff.find(s => s.id === staffId);
            if (staff) titleSuffix = staff.name;
        }

        // Prepare data for Table
        const tableBody = tasksToReport.map(t => {
            const staff = this.data.staff.find(s => s.id === t.staffId);
            return [
                staff ? staff.name : 'Unknown',
                t.date,
                t.title,
                t.status
            ];
        });

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        // Header
        doc.setFontSize(18);
        doc.text("Task Management System", 14, 22);
        doc.setFontSize(12);
        doc.setTextColor(100);
        doc.text(`${this.config.collegeName} - ${this.config.subheading}`, 14, 30);

        doc.setLineWidth(0.5);
        doc.line(14, 35, 196, 35);

        doc.setFontSize(11);
        doc.setTextColor(0);
        doc.text(`Report for: ${titleSuffix}`, 14, 45);
        doc.text(`Generated: ${new Date().toLocaleDateString()}`, 14, 50);

        // Table
        doc.autoTable({
            startY: 55,
            head: [['Staff Name', 'Date', 'Task Name', 'Status']],
            body: tableBody,
            theme: 'striped',
            headStyles: { fillColor: [15, 23, 42] }
        });

        doc.save(`Task_Report_${new Date().toISOString().split('T')[0]}.pdf`);
    },

    updateReportPreview() {
        const staffId = document.getElementById('report-staff-select').value;
        const tbody = document.querySelector('#report-table tbody');
        tbody.innerHTML = '';

        let tasks = this.data.tasks;
        if (staffId !== 'all') {
            tasks = tasks.filter(t => t.staffId === staffId);
        }

        tasks.forEach(t => {
            const staff = this.data.staff.find(s => s.id === t.staffId);
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${staff ? staff.name : 'Unknown'}</td>
                <td>${t.date}</td>
                <td>${t.title}</td>
                <td><span class="badge ${t.status}">${t.status}</span></td>
            `;
            tbody.appendChild(tr);
        });
    },

    // --- IMPORT/EXPORT ---
    exportData() {
        const dataStr = JSON.stringify(this.data);
        const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(dataStr);

        const exportFileDefaultName = 'tms_backup.json';

        const linkElement = document.createElement('a');
        linkElement.setAttribute('href', dataUri);
        linkElement.setAttribute('download', exportFileDefaultName);
        linkElement.click();
    },

    importData(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const imported = JSON.parse(e.target.result);
                if (imported.staff && imported.tasks) {
                    this.data = imported;
                    this.saveData();
                    this.refreshAll();
                    alert('Data imported successfully!');
                } else {
                    alert('Invalid file format');
                }
            } catch (err) {
                alert('Error parsing JSON');
            }
        };
        reader.readAsText(file);
    },

    // --- EVENT LISTENERS ---
    setupEventListeners() {
        // Navigation
        document.querySelectorAll('.nav-item').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');

                const viewName = btn.dataset.view;
                document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
                document.getElementById('view-' + viewName).classList.add('active');

                // Update header title
                document.getElementById('page-title').textContent = btn.querySelector('span').textContent;

                // Refresh analytics/tables if needed
                if (viewName === 'dashboard') this.renderDashboard();
                if (viewName === 'reports') this.updateReportPreview();
            });
        });

        // Staff CRUD
        document.getElementById('btn-add-staff').addEventListener('click', () => {
            document.getElementById('staff-form').reset();
            document.getElementById('staff-id-hidden').value = '';
            document.getElementById('modal-staff-title').textContent = 'Add Staff';
            document.getElementById('modal-staff').classList.remove('hidden');
            document.getElementById('modal-overlay').classList.remove('hidden');
        });

        document.getElementById('staff-form').addEventListener('submit', (e) => {
            e.preventDefault();
            const id = document.getElementById('staff-id-hidden').value;
            const staffData = {
                no: document.getElementById('staff-no').value,
                name: document.getElementById('staff-name').value
            };

            if (id) this.updateStaff(id, staffData);
            else this.addStaff(staffData);

            document.getElementById('modal-overlay').classList.add('hidden');
            document.getElementById('modal-staff').classList.add('hidden');
        });

        // Task CRUD
        document.getElementById('btn-add-task').addEventListener('click', () => {
            document.getElementById('task-form').reset();
            document.getElementById('task-id-hidden').value = '';

            // Set default date
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('task-date').value = today;

            document.getElementById('modal-task-title').textContent = 'Add Task';
            document.getElementById('modal-task').classList.remove('hidden');
            document.getElementById('modal-overlay').classList.remove('hidden');
        });

        document.getElementById('task-form').addEventListener('submit', (e) => {
            e.preventDefault();
            const id = document.getElementById('task-id-hidden').value;
            const taskData = {
                date: document.getElementById('task-date').value,
                staffId: document.getElementById('task-staff-select').value,
                title: document.getElementById('task-title').value,
                description: document.getElementById('task-desc').value,
                dueDate: document.getElementById('task-due-date').value
            };

            if (id) this.updateTask(id, taskData);
            else this.addTask(taskData);

            document.getElementById('modal-overlay').classList.add('hidden');
            document.getElementById('modal-task').classList.add('hidden');
        });

        // Search
        document.getElementById('staff-search').addEventListener('input', () => this.renderStaffTable());
        document.getElementById('task-search').addEventListener('input', () => this.renderTaskTable());

        // Report
        document.getElementById('report-staff-select').addEventListener('change', () => this.updateReportPreview());
        document.getElementById('btn-generate-report').addEventListener('click', () => this.generateReport());

        // Modals closing
        document.querySelectorAll('.close-modal').forEach(btn => {
            btn.addEventListener('click', () => {
                document.getElementById('modal-overlay').classList.add('hidden');
                document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
            });
        });

        // JSON Import
        document.getElementById('btn-import-data-click').addEventListener('click', () => {
            document.getElementById('json-import-input').click();
        });
        document.getElementById('json-import-input').addEventListener('change', (e) => {
            if (e.target.files.length > 0) this.importData(e.target.files[0]);
        });

        document.getElementById('btn-export-data').addEventListener('click', () => this.exportData());

        // Staff Excel Import
        document.getElementById('btn-import-staff').addEventListener('click', () => {
            document.getElementById('staff-upload-input').click();
        });
        document.getElementById('staff-upload-input').addEventListener('change', (e) => {
            if (e.target.files.length > 0) this.handleStaffUpload(e.target.files[0]);
        });

        // Sidebar Toggle
        const menuToggle = document.getElementById('menu-toggle');
        if (menuToggle) {
            menuToggle.addEventListener('click', () => {
                document.getElementById('sidebar').classList.toggle('open');
            });
        }

        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', (e) => {
            const sidebar = document.getElementById('sidebar');
            const toggle = document.getElementById('menu-toggle');
            if (window.innerWidth <= 768 &&
                sidebar && toggle &&
                !sidebar.contains(e.target) &&
                !toggle.contains(e.target) &&
                sidebar.classList.contains('open')) {
                sidebar.classList.remove('open');
            }
        });
    }
};

// Start the APP
window.addEventListener('DOMContentLoaded', () => {
    App.init();
});
