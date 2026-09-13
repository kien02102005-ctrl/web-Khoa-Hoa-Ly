<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WORKSPACE - Khoa Hóa Lý (Realtime KPI Master Sync)</title>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- FontAwesome Icon -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <!-- Firebase SDK (Modular v9/v10 Compat) -->
    <script src="https://www.gstatic.com/firebasejs/10.8.0/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/10.8.0/firebase-database-compat.js"></script>

    <!-- Thư viện docx & FileSaver để Xuất File Word chuẩn định dạng Docx -->
    <script src="https://unpkg.com/docx@8.5.0/build/index.umd.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

    <style>
        :root {
            --primary: #1e40af;
            --primary-hover: #1d4ed8;
            --sidebar-bg: #0f172a;
            --bg-light: #f1f5f9;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background-color: var(--bg-light);
            height: 100vh;
            overflow: hidden;
        }

        /* --- 1. GIAO DIỆN ĐĂNG NHẬP --- */
        #login-screen {
            position: fixed;
            top: 0; left: 0; width: 100vw; height: 100vh;
            background: linear-gradient(135deg, #1e3a8a, #3b82f6);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 9999;
        }

        .login-card {
            background: #fff;
            padding: 40px 30px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            width: 380px;
            text-align: center;
        }

        .login-card .icon {
            font-size: 48px;
            color: #2563eb;
            margin-bottom: 15px;
        }

        .login-card h2 {
            color: #1e293b;
            margin-bottom: 5px;
            text-transform: uppercase;
        }

        .login-card p {
            color: #64748b;
            font-size: 13px;
            margin-bottom: 25px;
        }

        .form-group {
            text-align: left;
            margin-bottom: 18px;
        }

        .form-group label {
            display: block;
            font-size: 13px;
            font-weight: 600;
            color: #334155;
            margin-bottom: 5px;
        }

        .form-group input, .form-group select {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            outline: none;
            font-size: 14px;
        }

        .form-group input:focus, .form-group select:focus {
            border-color: #2563eb;
        }

        .btn-login {
            width: 100%;
            padding: 12px;
            background-color: #2563eb;
            color: #fff;
            border: none;
            border-radius: 6px;
            font-weight: bold;
            font-size: 15px;
            cursor: pointer;
            transition: 0.2s;
        }

        .btn-login:hover {
            background-color: #1d4ed8;
        }

        /* --- 2. GIAO DIỆN CHÍNH (APP MAIN) --- */
        #app-screen {
            display: flex;
            height: 100vh;
        }

        .sidebar {
            width: 240px;
            background-color: var(--sidebar-bg);
            color: #fff;
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
        }

        .sidebar-brand {
            padding: 20px;
            font-size: 20px;
            font-weight: bold;
            display: flex;
            align-items: center;
            gap: 10px;
            border-bottom: 1px solid #334155;
        }

        .user-profile {
            padding: 15px 20px;
            border-bottom: 1px solid #334155;
        }

        .user-profile .name {
            font-weight: 600;
            font-size: 15px;
        }

        .user-profile .badge {
            display: inline-block;
            background-color: #ef4444;
            color: #fff;
            font-size: 10px;
            padding: 2px 8px;
            border-radius: 10px;
            margin-top: 4px;
        }

        .nav-list {
            list-style: none;
            padding: 15px 0;
        }

        .nav-item {
            padding: 12px 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            cursor: pointer;
            color: #94a3b8;
            transition: 0.2s;
            font-size: 14px;
        }

        .nav-item:hover, .nav-item.active {
            background-color: #2563eb;
            color: #fff;
        }

        .main-content {
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }

        .top-bar {
            background-color: #fff;
            padding: 15px 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #e2e8f0;
        }

        .top-bar h2 {
            font-size: 18px;
            color: #1e293b;
        }

        .top-actions {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        .admin-select-box {
            background: #f8fafc;
            border: 1px solid #cbd5e1;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 13px;
            display: none;
            align-items: center;
            gap: 8px;
        }

        .admin-select-box select {
            padding: 4px 8px;
            border-radius: 4px;
            border: 1px solid #cbd5e1;
            outline: none;
            font-weight: bold;
            color: #1e40af;
        }

        .btn-action {
            padding: 8px 14px;
            border-radius: 6px;
            border: 1px solid #cbd5e1;
            background: #fff;
            cursor: pointer;
            font-size: 13px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .btn-action.btn-danger {
            background-color: #ef4444;
            color: white;
            border: none;
        }

        .tab-content {
            flex: 1;
            padding: 20px;
            overflow-y: auto;
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
            padding: 10px;
        }

        .card {
            background: #ffffff;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
            border: 1px solid #f1f5f9;
            margin-bottom: 20px;
        }

        .card h3 {
            font-size: 18px;
            font-weight: 700;
            color: #334155;
            margin-bottom: 20px;
        }

        .chart-container {
            position: relative;
            height: 300px;
            width: 100%;
        }

        .task-input-bar {
            background: #fff;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
            align-items: center;
        }

        .task-input-bar input[type="text"],
        .task-input-bar select,
        .task-input-bar input[type="date"] {
            padding: 8px 12px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            outline: none;
            font-size: 14px;
        }

        .task-input-bar input[type="text"] { flex: 2; min-width: 200px; }
        .task-input-bar select { flex: 1; min-width: 120px; }
        .task-input-bar input[type="date"] { flex: 1; min-width: 140px; cursor: pointer; }

        .table-container {
            background: #fff;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 14px;
        }

        th, td {
            padding: 12px 16px;
            border-bottom: 1px solid #e2e8f0;
        }

        th {
            background-color: #f8fafc;
            color: #475569;
        }

        .inline-date-picker {
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            padding: 4px 6px;
            font-size: 13px;
            outline: none;
            cursor: pointer;
            background: #fff;
        }

        tr.row-overdue { background-color: #fef2f2 !important; }
        tr.row-overdue td { color: #991b1b; }

        .badge-overdue {
            background-color: #ef4444;
            color: #ffffff;
            font-size: 11px;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: bold;
            display: inline-flex;
            align-items: center;
            gap: 4px;
            margin-left: 6px;
        }

        .status-tag {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
        }

        .status-doing { background: #dbeafe; color: #1d4ed8; }
        .status-todo { background: #f3e8ff; color: #6b21a8; }
        .status-done { background: #dcfce7; color: #15803d; }

        .kanban-board {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            height: 100%;
        }

        .kanban-col {
            background: #e2e8f0;
            border-radius: 8px;
            padding: 15px;
            min-height: 400px;
        }

        .kanban-col-header {
            font-weight: bold;
            margin-bottom: 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .kanban-card {
            background: #fff;
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 10px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
            cursor: grab;
            border-left: 4px solid transparent;
        }

        .kanban-card.card-overdue {
            background-color: #fef2f2;
            border-left: 4px solid #ef4444;
        }

        .kanban-card:active { cursor: grabbing; }
        .kanban-card .id { color: #2563eb; font-size: 12px; font-weight: bold; }
        .kanban-card .title { font-size: 14px; font-weight: 600; margin: 5px 0 10px; }
        .kanban-card .meta { font-size: 12px; color: #64748b; display: flex; justify-content: space-between; align-items: center; }

        .kpi-section-title {
            background: #1e40af;
            color: #ffffff;
            padding: 10px 16px;
            font-size: 15px;
            font-weight: bold;
            border-radius: 6px 6px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .row-sub-header {
            background-color: #f1f5f9;
            font-weight: bold;
            color: #1e293b;
        }

        .kpi-table input[type="number"] {
            width: 75px;
            padding: 4px 6px;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            text-align: center;
            font-weight: 600;
        }

        .btn-sm {
            padding: 4px 8px;
            font-size: 11px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            margin-right: 2px;
        }
        .btn-edit { background-color: #f59e0b; color: white; }
        .btn-delete { background-color: #ef4444; color: white; }
        .btn-attach { background-color: #3b82f6; color: white; }
        .btn-download { background-color: #10b981; color: white; }

        .file-tag {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: #e2e8f0;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 11px;
            margin: 2px 0;
            color: #1e293b;
        }

        .file-tag a {
            color: #2563eb;
            text-decoration: none;
            font-weight: 600;
        }

        .file-tag a:hover {
            text-decoration: underline;
        }

        .kpi-target-bar {
            background: #e0f2fe;
            border: 1px solid #bae6fd;
            padding: 12px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 15px;
            flex-wrap: wrap;
        }

        .kpi-target-bar select {
            padding: 6px 12px;
            border-radius: 6px;
            border: 1px solid #0284c7;
            font-weight: bold;
            color: #0369a1;
            outline: none;
        }

        .kpi-month-selector-bar {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            padding: 12px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }

        .kpi-month-selector-bar select {
            padding: 6px 12px;
            border-radius: 6px;
            border: 1px solid #2563eb;
            font-weight: bold;
            color: #1e40af;
            outline: none;
            background: #f0f9ff;
        }

        .kpi-badge-type {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            margin-left: 10px;
        }
        .type-leader { background-color: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
        .type-staff { background-color: #e0e7ff; color: #3730a3; border: 1px solid #c7d2fe; }
        .type-cleaner { background-color: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }

        .kpi-total-box {
            display: flex;
            justify-content: flex-end;
            gap: 30px;
            margin-top: 20px;
            font-size: 16px;
            font-weight: bold;
            background: #f8fafc;
            padding: 15px 20px;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
        }

        .sync-status {
            font-size: 12px;
            padding: 4px 8px;
            border-radius: 4px;
            background: #dcfce7;
            color: #15803d;
            display: flex;
            align-items: center;
            gap: 5px;
        }
    </style>
</head>
<body>

    <!-- 1. MÀN HÌNH ĐĂNG NHẬP -->
    <div id="login-screen">
        <div class="login-card">
            <div class="icon"><i class="fa-solid fa-flask"></i></div>
            <h2>KHOA HÓA LÝ</h2>
            <p id="form-sub-title">Đăng nhập hệ thống quản trị công việc (Realtime)</p>
            
            <div style="display: flex; gap: 10px; margin-bottom: 20px; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">
                <button type="button" id="tab-login-btn" onclick="toggleAuthTab('login')" style="flex: 1; padding: 8px; border: none; background: none; font-weight: bold; color: #2563eb; border-bottom: 2px solid #2563eb; cursor: pointer;">ĐĂNG NHẬP</button>
                <button type="button" id="tab-register-btn" onclick="toggleAuthTab('register')" style="flex: 1; padding: 8px; border: none; background: none; font-weight: bold; color: #64748b; cursor: pointer;">ĐĂNG KÝ</button>
            </div>

            <form id="login-form" onsubmit="handleLogin(event)">
                <div class="form-group">
                    <label>Tên đăng nhập</label>
                    <input type="text" id="login-username" placeholder="Nhập tên đăng nhập..." required>
                </div>
                <div class="form-group">
                    <label>Mật khẩu</label>
                    <input type="password" id="login-password" placeholder="Nhập mật khẩu..." required>
                </div>
                <button type="submit" class="btn-login">ĐĂNG NHẬP</button>
            </form>

            <form id="register-form" onsubmit="handleRegister(event)" style="display: none;">
                <div class="form-group">
                    <label>Gmail</label>
                    <input type="email" id="reg-email" placeholder="example@gmail.com" required>
                </div>
                <div class="form-group">
                    <label>Tên đăng nhập</label>
                    <input type="text" id="reg-username" placeholder="Tạo tên đăng nhập..." required>
                </div>
                <div class="form-group">
                    <label>Mật khẩu</label>
                    <input type="password" id="reg-password" placeholder="Nhập mật khẩu..." required>
                </div>
                <div class="form-group">
                    <label>Xác nhận mật khẩu</label>
                    <input type="password" id="reg-confirm-password" placeholder="Nhập lại mật khẩu..." required>
                </div>
                <div class="form-group">
                    <label>Loại Bảng KPI áp dụng</label>
                    <select id="reg-kpi-type">
                        <option value="staff">Bảng KPI Dành cho Cán bộ / Nhân viên</option>
                        <option value="leader">Bảng KPI Dành cho Lãnh đạo</option>
                        <option value="cleaner">Bảng KPI Dành cho Lao Công</option>
                    </select>
                </div>
                <button type="submit" class="btn-login" style="background-color: #10b981;">ĐĂNG KÝ TÀI KHOẢN</button>
            </form>
        </div>
    </div>

    <!-- 2. MÀN HÌNH CHÍNH WEB APP -->
    <div id="app-screen" style="display: none;">
        <!-- Sidebar -->
        <div class="sidebar">
            <div class="sidebar-brand">
                <i class="fa-solid fa-shapes"></i> WORKSPACE
            </div>
            <div class="user-profile">
                <div class="name" id="user-display-name">Cán bộ</div>
                <span class="badge" id="user-role-badge">User</span>
            </div>
            <ul class="nav-list">
                <li class="nav-item active" onclick="switchTab('tong-quan', this)">
                    <i class="fa-solid fa-chart-pie"></i> Tổng quan
                </li>
                <li class="nav-item" onclick="switchTab('danh-sach', this)">
                    <i class="fa-solid fa-list-check"></i> Danh sách
                </li>
                <li class="nav-item" onclick="switchTab('kanban', this)">
                    <i class="fa-solid fa-table-columns"></i> Tiến độ công việc
                </li>
                <li class="nav-item" onclick="switchTab('gantt', this)">
                    <i class="fa-solid fa-bars-progress"></i> Sơ đồ Gantt
                </li>
                <li class="nav-item" onclick="switchTab('kpi', this)">
                    <i class="fa-solid fa-award"></i> Đánh giá KPI
                </li>
                <li class="nav-item" id="nav-admin-users" style="display: none;" onclick="switchTab('admin-users', this)">
                    <i class="fa-solid fa-users-gear"></i> Quản lý Users
                </li>
            </ul>
        </div>

        <!-- Main Content -->
        <div class="main-content">
            <!-- Top Bar -->
            <div class="top-bar">
                <h2 id="page-title">Dashboard Thống Kê</h2>
                <div class="top-actions">
                    <div class="sync-status"><i class="fa-solid fa-arrows-rotate fa-spin"></i> Đồng bộ Realtime</div>
                    <div class="admin-select-box" id="admin-user-selector">
                        <span><i class="fa-solid fa-user-pen"></i> Xem data công việc của:</span>
                        <select id="select-target-user" onchange="changeTargetUser(this.value)"></select>
                    </div>
                    <button class="btn-action btn-danger" onclick="logout()"><i class="fa-solid fa-power-off"></i> Đăng xuất</button>
                </div>
            </div>

            <!-- Tab 1: Tổng quan -->
            <div id="tab-tong-quan" class="tab-content active">
                <div id="admin-master-overview-banner" style="display:none; background: #eff6ff; border: 1px solid #bfdbfe; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: bold; color: #1e40af;">
                    <i class="fa-solid fa-circle-info"></i> Chế độ Admin: Đang tổng hợp dữ liệu toàn bộ tài khoản thường trong hệ thống.
                </div>
                <div class="dashboard-grid">
                    <div class="card">
                        <h3>Tỷ lệ Trạng thái</h3>
                        <div class="chart-container">
                            <canvas id="statusChart"></canvas>
                        </div>
                    </div>
                    <div class="card">
                        <h3>Mức độ Ưu tiên</h3>
                        <div class="chart-container">
                            <canvas id="priorityChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Tab 2: Danh sách -->
            <div id="tab-danh-sach" class="tab-content">
                <h3 style="font-size: 16px; color: #334155; margin-bottom: 12px;">Quản lý Công Việc Độc Lập</h3>
                
                <div class="task-input-bar">
                    <input type="text" id="newTaskName" placeholder="Nhập tên công việc mới..." />
                    <select id="newTaskPriority">
                        <option value="Bình thường">Bình thường</option>
                        <option value="Cao">Cao</option>
                        <option value="Thấp">Thấp</option>
                    </select>
                    <input type="date" id="newTaskDueDate" />
                    <button class="btn-login" style="width: auto; padding: 8px 18px;" onclick="addNewTask()"><i class="fa-solid fa-plus"></i> Thêm công việc</button>
                </div>

                <div class="table-container">
                    <table>
                        <thead>
                            <tr>
                                <th>Mã</th>
                                <th>Tên công việc</th>
                                <th>Người nhận</th>
                                <th>Mức ưu tiên</th>
                                <th>Trạng thái</th>
                                <th>Hạn chót</th>
                                <th>File đính kèm</th>
                                <th>Thao tác</th>
                            </tr>
                        </thead>
                        <tbody id="task-table-body"></tbody>
                    </table>
                </div>
            </div>

            <!-- Tab 3: Kanban / Tiến độ công việc -->
            <div id="tab-kanban" class="tab-content">
                <div class="kanban-board">
                    <div class="kanban-col" id="col-todo" ondragover="allowDrop(event)" ondrop="drop(event, 'Chưa làm')">
                        <div class="kanban-col-header">🌙 Chưa làm <span id="count-todo">0</span></div>
                        <div class="kanban-cards" id="cards-todo"></div>
                    </div>
                    <div class="kanban-col" id="col-doing" ondragover="allowDrop(event)" ondrop="drop(event, 'Đang làm')">
                        <div class="kanban-col-header">⌛ Đang làm <span id="count-doing">0</span></div>
                        <div class="kanban-cards" id="cards-doing"></div>
                    </div>
                    <div class="kanban-col" id="col-done" ondragover="allowDrop(event)" ondrop="drop(event, 'Hoàn thành')">
                        <div class="kanban-col-header">✔️ Hoàn thành <span id="count-done">0</span></div>
                        <div class="kanban-cards" id="cards-done"></div>
                    </div>
                </div>
            </div>

            <!-- Tab 4: Gantt -->
            <div id="tab-gantt" class="tab-content">
                <div class="card">
                    <h3>Lộ trình triển khai</h3>
                    <p style="color: #64748b; font-size: 14px;">(Sơ đồ tiến độ công việc cá nhân)</p>
                </div>
            </div>

            <!-- Tab 5: Đánh giá KPI -->
            <div id="tab-kpi" class="tab-content">
                <!-- Thanh Chọn Tháng KPI & Lịch Sử Lưu -->
                <div class="kpi-month-selector-bar">
                    <i class="fa-regular fa-calendar-check" style="font-size: 22px; color: #2563eb;"></i>
                    <span style="font-weight: bold; color: #1e293b;">KỲ ĐÁNH GIÁ KPI:</span>
                    <select id="select-kpi-month" onchange="changeKPIMonth(this.value)"></select>

                    <span style="font-size: 13px; color: #64748b; margin-left: auto;">
                        <i class="fa-solid fa-clock-rotate-left"></i> Dữ liệu các tháng trước được lưu trữ tự động
                    </span>
                </div>

                <div class="kpi-target-bar" id="kpi-admin-target-bar" style="display: none;">
                    <i class="fa-solid fa-user-check" style="font-size: 20px; color: #0284c7;"></i>
                    <span style="font-weight: 600; color: #0369a1;">Chọn tài khoản để chấm KPI:</span>
                    <select id="select-kpi-target-user" onchange="changeKPITargetUser(this.value)"></select>

                    <div style="margin-left: 20px; display: flex; align-items: center; gap: 8px;">
                        <span style="font-weight: 600; color: #0369a1;"><i class="fa-solid fa-arrow-right-arrow-left"></i> Chuyển Bảng KPI:</span>
                        <select id="select-kpi-type-change" onchange="adminChangeUserKPIType(this.value)">
                            <option value="staff">Bảng KPI Dành cho Cán bộ / Nhân viên</option>
                            <option value="leader">Bảng KPI Dành cho Lãnh đạo</option>
                            <option value="cleaner">Bảng KPI Dành cho Lao Công</option>
                        </select>
                    </div>
                </div>

                <div class="card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                        <h3 style="font-size: 18px; margin: 0; display: flex; align-items: center;">
                            PHIẾU ĐÁNH GIÁ KPI CỦA: <span id="kpi-target-name-display" style="color: #2563eb; text-transform: uppercase; margin-left: 8px;"></span>
                            <span id="kpi-type-badge-display" class="kpi-badge-type"></span>
                        </h3>
                        <!-- NÚT THÊM MỤC LỚN A, B, C... DÀNH CHO ADMIN -->
                        <button id="btn-add-main-section" class="btn-login" style="width: auto; background-color: #8b5cf6; padding: 8px 16px; display: none;" onclick="addMainSection()">
                            <i class="fa-solid fa-folder-plus"></i> Thêm Mục Lớn (D, E...)
                        </button>
                    </div>

                    <!-- CONTAINER CHỨA CÁC BẢNG MỤC LỚN (A, B, C, D...) ĐỘNG -->
                    <div id="kpi-sections-wrapper"></div>
                    
                    <!-- Bảng Tổng Điểm -->
                    <div class="kpi-total-box">
                        <div>
                            TỔNG ĐIỂM TỰ CHẤM: 
                            <span id="kpi-total-self" style="color: #2563eb;">0</span> / 
                            <span id="kpi-total-max" style="color: #64748b;">0</span>
                        </div>
                        <div style="border-left: 2px solid #cbd5e1; padding-left: 20px;">
                            TỔNG ĐIỂM ĐÁNH GIÁ: 
                            <span id="kpi-total-admin" style="color: #059669;">0</span>
                        </div>
                    </div>
                    
                    <div style="margin-top: 15px; text-align: right; font-size: 13px; color: #475569;">
                        <i class="fa-regular fa-clock" style="color: #0284c7;"></i> Lần cuối lưu điểm (<span id="selected-month-text" style="font-weight: bold; color: #0284c7;"></span>): <span id="kpi-last-time-saved" style="font-weight: bold; color: #1e293b;">Chưa có dữ liệu</span>
                    </div>

                    <div style="text-align: center; margin-top: 20px; display: flex; justify-content: center; gap: 15px; flex-wrap: wrap;">
                        <button id="btn-save-kpi-admin" class="btn-login" style="width: auto; padding: 10px 24px; background-color: #2563eb; display: none;" onclick="saveKPIRatingByAdmin()">
                            <i class="fa-solid fa-floppy-disk"></i> LƯU KẾT QUẢ ĐÁNH GIÁ THÁNG NÀY (ADMIN)
                        </button>
                        <button id="btn-save-kpi-user" class="btn-login" style="width: auto; padding: 10px 24px; background-color: #10b981; display: none;" onclick="saveKPIRatingByUser()">
                            <i class="fa-solid fa-user-check"></i> LƯU KẾT QUẢ TỰ CHẤM THÁNG NÀY
                        </button>
                        <!-- NÚT XUẤT FILE WORD THỰC THI CHUẨN ĐỊNH DẠNG -->
                        <button class="btn-login" onclick="exportKPIWord()" style="width: auto; background-color: #0284c7; padding: 10px 24px;">
                            <i class="fa-solid fa-file-word"></i> XUẤT FILE WORD
                        </button>
                    </div>
                </div>
            </div>

            <!-- Tab 6: Quản lý Users -->
            <div id="tab-admin-users" class="tab-content">
                <div class="card">
                    <h3>Danh Sách Tài Khoản Trong Hệ Thống (Online Sync)</h3>
                    <p style="color: #64748b; font-size: 13px; margin-bottom: 15px;">* Bạn có thể phân loại bảng đánh giá KPI cho từng tài khoản (Bao gồm cả Admin) tại cột "Loại Bảng KPI".</p>
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>STT</th>
                                    <th>Tên đăng nhập</th>
                                    <th>Email</th>
                                    <th>Loại Bảng KPI</th>
                                    <th>Thao tác</th>
                                </tr>
                            </thead>
                            <tbody id="user-management-body"></tbody>
                        </table>
                    </div>
                </div>
            </div>

        </div>
    </div>

    <!-- JAVASCRIPT XỬ LÝ DỮ LIỆU REALTIME & XUẤT FILE WORD THỰC THI CHUẨN -->
    <script>
        // CẤU HÌNH FIREBASE
        const firebaseConfig = {
            apiKey: "AiZaSyDvMPWqdqgTJ5SzMrhhV56EOpv95yow4VA",
            authDomain: "kien02102005-381b4.firebaseapp.com",
            projectId: "kien02102005-381b4",
            storageBucket: "kien02102005-381b4.firebasestorage.app",
            messagingSenderId: "554013353586",
            appId: "1:554013353586:web:8a64d8ba969104053404eb",
            measurementId: "G-TJE25HB0PQ",
            databaseURL: "https://kien02102005-381b4-default-rtdb.firebaseio.com"
        };

        firebase.initializeApp(firebaseConfig);
        const db = firebase.database();

        const ADMIN_USERNAME = 'linhnguyenxuan';
        const ADMIN_PASSWORD = '051214';

        let currentUser = '';          
        let targetUser = '';           
        let kpiTargetUser = '';        
        let selectedKpiMonth = ''; 
        let isAdmin = false;
        let lastKpiTimestamp = 'Chưa có dữ liệu';
        let isKpiSavedForCurrentMonth = false;

        let registeredUsers = [];
        let tasks = [];
        let allUsersTasksMap = {}; // Lưu trữ toàn bộ tasks của tất cả user phục vụ tổng quan admin
        let kpiDataList = [];
        let sectionMaxScores = { A: 50, B: 30, C: 20 };
        let sectionTitles = {
            A: "QUẢN LÝ & CHUYÊN MÔN",
            B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT",
            C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT"
        };

        let masterKPITemplates = {
            staff: { 
                sectionMaxScores: { A: 50, B: 30, C: 20 }, 
                sectionTitles: { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" },
                kpiDataList: [] 
            },
            leader: { 
                sectionMaxScores: { A: 50, B: 30, C: 20 }, 
                sectionTitles: { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" },
                kpiDataList: [] 
            },
            cleaner: { 
                sectionMaxScores: { A: 70, B: 30 }, 
                sectionTitles: { A: "CÔNG TÁC CHUYÊN MÔN & VỆ SINH", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT" },
                kpiDataList: [] 
            }
        };

        let statusChartInstance = null;
        let priorityChartInstance = null;

        listenRealtimeUsers();
        listenRealtimeTemplates();
        listenAllUsersTasks(); // Lắng nghe toàn bộ task hệ thống cho admin tổng quan

        function isValidNumber(val) {
            if (val === null || val === undefined) return false;
            let str = String(val).replace(',', '.').trim();
            if (str === '') return false;
            let num = Number(str);
            return !isNaN(num) && isFinite(num);
        }

        function parseFloatStrict(val) {
            if (!isValidNumber(val)) return 0;
            return parseFloat(String(val).replace(',', '.').trim());
        }

        function convertToDisplayDate(isoDateStr) {
            if (!isoDateStr) return '';
            if (isoDateStr.includes('/')) return isoDateStr;
            const parts = isoDateStr.split('-');
            if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
            return isoDateStr;
        }

        function convertToISODate(displayDateStr) {
            if (!displayDateStr) return '';
            if (displayDateStr.includes('-')) return displayDateStr;
            const parts = displayDateStr.split('/');
            if (parts.length === 3) return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            return displayDateStr;
        }

        const defaultStaffKPIStructure = [
            {
                id: 'sub_a1', section: 'A', code: 'I', title: 'Phẩm chất chính trị, phẩm chất đạo đức, văn hóa thực thi công vụ, nhiệm vụ và ý thức kỷ luật, kỷ cương trong thực thi công vụ, nhiệm vụ', maxScore: 10,
                items: [
                    { id: 'item_a1_1', title: 'Phẩm chất chính trị, phẩm chất đạo đức, văn hóa thực thi công vụ, nhiệm vụ', criteria: '', maxScore: 5, selfScore: 5, adminScore: 5 },
                    { id: 'item_a1_2', title: 'Ý thức kỷ luật, kỷ cương trong thực thi công vụ, nhiệm vụ', criteria: '', maxScore: 5, selfScore: 5, adminScore: 5 }
                ]
            },
            {
                id: 'sub_a2', section: 'A', code: 'II', title: 'Năng lực chuyên môn, nghiệp vụ theo yêu cầu của vị trí việc làm; khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao; tinh thần trách nhiệm trong thực thi công vụ, nhiệm vụ; thái độ phục vụ người dân, doanh nghiệp và khả năng phối hợp với đồng nghiệp', maxScore: 10,
                items: [
                    { id: 'item_a2_1', title: 'Năng lực chuyên môn, nghiệp vụ theo yêu cầu của vị trí việc làm', criteria: '', maxScore: 2.5, selfScore: 2.5, adminScore: 2.5 },
                    { id: 'item_a2_2', title: 'Khả năng đáp ứng yêu cầu thực thi nhiệm vụ được giao thường xuyên, đột xuất', criteria: '', maxScore: 2.5, selfScore: 2.5, adminScore: 2.5 },
                    { id: 'item_a2_3', title: 'Tinh thần trách nhiệm trong thực thi công vụ, nhiệm vụ', criteria: '', maxScore: 2.5, selfScore: 2.5, adminScore: 2.5 },
                    { id: 'item_a2_4', title: 'Thái độ phục vụ người dân, doanh nghiệp và khả năng phối hợp với đồng nghiệp', criteria: '', maxScore: 2.5, selfScore: 2.5, adminScore: 2.5 }
                ]
            }
        ];

        const defaultLeaderKPIStructure = [
            {
                id: 'sub_a1', section: 'A', code: 'I', title: 'Năng lực Lãnh đạo & Quản lý điều hành (Lãnh đạo)', maxScore: 30,
                items: [
                    { id: 'item_a1_1', title: 'Xây dựng kế hoạch và chỉ đạo thực hiện nhiệm vụ khoa', criteria: '100% chỉ tiêu năm đạt tiến độ', maxScore: 15, selfScore: 15, adminScore: 15 },
                    { id: 'item_a1_2', title: 'Quản lý nhân sự và phát triển đội ngũ', criteria: 'Không có cán bộ vi phạm kỷ luật', maxScore: 15, selfScore: 14, adminScore: 15 }
                ]
            }
        ];

        const defaultCleanerKPIStructure = [
            {
                id: 'sub_cleaner_a1', section: 'A', code: 'I', title: 'Công tác Vệ sinh & Môi trường làm việc', maxScore: 40,
                items: [
                    { id: 'item_cleaner_a1_1', title: 'Dọn dẹp vệ sinh khu vực được phân công (Phòng làm việc, hành lang, nhà vệ sinh)', criteria: 'Sạch sẽ, gọn gàng, đúng lịch trình', maxScore: 20, selfScore: 20, adminScore: 20 },
                    { id: 'item_cleaner_a1_2', title: 'Thu gom và phân loại rác thải đúng quy định', criteria: 'Không tồn đọng rác thải qua ngày', maxScore: 20, selfScore: 20, adminScore: 20 }
                ]
            },
            {
                id: 'sub_cleaner_a2', section: 'A', code: 'II', title: 'Bảo quản Vật tư & Thiết bị Vệ sinh', maxScore: 30,
                items: [
                    { id: 'item_cleaner_a2_1', title: 'Quản lý và sử dụng tiết kiệm dung dịch, hóa chất, dụng cụ vệ sinh', criteria: 'Không lãng phí, sử dụng đúng hướng dẫn', maxScore: 15, selfScore: 15, adminScore: 15 },
                    { id: 'item_cleaner_a2_2', title: 'Bảo quản và kiểm tra trang thiết bị làm việc', criteria: 'Bảo dưỡng tốt, báo cáo kịp thời hư hỏng', maxScore: 15, selfScore: 15, adminScore: 15 }
                ]
            },
            {
                id: 'sub_cleaner_b1', section: 'B', code: 'I', title: 'Chấp hành Kỷ luật & Thái độ làm việc', maxScore: 30,
                items: [
                    { id: 'item_cleaner_b1_1', title: 'Chấp hành thời gian làm việc và nội quy khoa/viện', criteria: 'Đúng giờ, không tự ý bỏ vị trí', maxScore: 15, selfScore: 15, adminScore: 15 },
                    { id: 'item_cleaner_b1_2', title: 'Thái độ giao tiếp, ứng xử với cán bộ và đồng nghiệp', criteria: 'Mực thước, hòa nhã, lịch sự', maxScore: 15, selfScore: 15, adminScore: 15 }
                ]
            }
        ];

        const defaultTasks = [
            { id: 'T001', name: 'Nghiên cứu tài liệu khoa học', status: 'Đang làm', date: '02/09/2026', priority: 'Cao', files: [] },
            { id: 'T002', name: 'Chuẩn bị hóa chất phòng thí nghiệm', status: 'Chưa làm', date: '30/09/2026', priority: 'Bình thường', files: [] },
            { id: 'T003', name: 'Viết báo cáo tổng kết tháng', status: 'Hoàn thành', date: '01/09/2026', priority: 'Thấp', files: [] }
        ];

        function initMonthSelector() {
            const monthSelect = document.getElementById('select-kpi-month');
            if (!monthSelect) return;
            monthSelect.innerHTML = '';
            const now = new Date();
            const currentYear = now.getFullYear();

            selectedKpiMonth = `${currentYear}-${String(now.getMonth() + 1).padStart(2, '0')}`;

            for (let y = currentYear; y >= currentYear - 1; y--) {
                for (let m = 12; m >= 1; m--) {
                    const monthKey = `${y}-${String(m).padStart(2, '0')}`;
                    const label = `Đánh giá KPI Tháng ${m}/${y}`;
                    const option = document.createElement('option');
                    option.value = monthKey;
                    option.innerText = label;
                    if (monthKey === selectedKpiMonth) option.selected = true;
                    monthSelect.appendChild(option);
                }
            }
            updateSelectedMonthText();
        }

        function updateSelectedMonthText() {
            if (!selectedKpiMonth) return;
            const parts = selectedKpiMonth.split('-');
            const el = document.getElementById('selected-month-text');
            if (el) el.innerText = `Tháng ${parts[1]}/${parts[0]}`;
        }

        function changeKPIMonth(newMonth) {
            selectedKpiMonth = newMonth;
            updateSelectedMonthText();
            listenRealtimeKPI();
        }

        window.addEventListener('DOMContentLoaded', () => {
            initMonthSelector();
            const dateInput = document.getElementById('newTaskDueDate');
            if (dateInput) {
                dateInput.value = new Date().toISOString().split('T')[0];
            }
        });

        function listenRealtimeTemplates() {
            db.ref('kpiTemplates').on('value', snapshot => {
                const data = snapshot.val();
                if (data) {
                    masterKPITemplates = data;
                    ['staff', 'leader', 'cleaner'].forEach(type => {
                        if (!masterKPITemplates[type]) {
                            masterKPITemplates[type] = {
                                sectionMaxScores: type === 'cleaner' ? { A: 70, B: 30 } : { A: 50, B: 30, C: 20 },
                                sectionTitles: type === 'cleaner' ? 
                                    { A: "CÔNG TÁC CHUYÊN MÔN & VỆ SINH", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT" } : 
                                    { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" },
                                kpiDataList: type === 'cleaner' ? defaultCleanerKPIStructure : type === 'leader' ? defaultLeaderKPIStructure : defaultStaffKPIStructure
                            };
                            db.ref(`kpiTemplates/${type}`).set(masterKPITemplates[type]);
                        }
                    });
                } else {
                    masterKPITemplates = {
                        staff: { 
                            sectionMaxScores: { A: 50, B: 30, C: 20 }, 
                            sectionTitles: { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" },
                            kpiDataList: defaultStaffKPIStructure 
                        },
                        leader: { 
                            sectionMaxScores: { A: 50, B: 30, C: 20 }, 
                            sectionTitles: { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" },
                            kpiDataList: defaultLeaderKPIStructure 
                        },
                        cleaner: { 
                            sectionMaxScores: { A: 70, B: 30 }, 
                            sectionTitles: { A: "CÔNG TÁC CHUYÊN MÔN & VỆ SINH", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT" },
                            kpiDataList: defaultCleanerKPIStructure 
                        }
                    };
                    db.ref('kpiTemplates').set(masterKPITemplates);
                }
                
                if (kpiTargetUser) {
                    loadAndMergeUserKPI(kpiTargetUser, selectedKpiMonth);
                }
            });
        }

        // HÀM TẢI USER REALTIME VÀ TỰ ĐỘNG CẬP NHẬT BẢNG QUẢN LÝ
        function listenRealtimeUsers() {
            db.ref('users').on('value', snapshot => {
                const data = snapshot.val();
                let loadedUsers = data ? Object.values(data) : [];
                
                const hasAdmin = loadedUsers.some(u => u.username === ADMIN_USERNAME);
                if (!hasAdmin) {
                    const adminObj = { email: 'admin@hoaly.edu.vn', username: ADMIN_USERNAME, password: ADMIN_PASSWORD, kpiType: 'leader' };
                    db.ref(`users/${ADMIN_USERNAME}`).set(adminObj);
                    loadedUsers.push(adminObj);
                }

                registeredUsers = loadedUsers;
                if (isAdmin) {
                    populateAdminUserSelector();
                    populateKPITargetSelector();
                    renderUserManagementTable();
                    renderDashboard(); // Cập nhật lại biểu đồ tổng quan khi user thay đổi
                }
            });
        }

        // Lắng nghe toàn bộ task của tất cả các user trên hệ thống
        function listenAllUsersTasks() {
            db.ref('tasks').on('value', snapshot => {
                const data = snapshot.val();
                if (data) {
                    allUsersTasksMap = data;
                } else {
                    allUsersTasksMap = {};
                }
                if (isAdmin) {
                    renderDashboard();
                }
            });
        }

        function listenRealtimeTasks() {
            if (!targetUser) return;
            db.ref(`tasks/${targetUser}`).on('value', snapshot => {
                const data = snapshot.val();
                if (data) {
                    tasks = Object.values(data);
                } else {
                    tasks = defaultTasks.map(t => ({ ...t, user: targetUser }));
                    saveUserData();
                }
                renderDashboard();
                renderTaskList();
                renderKanban();
            });
        }

        function listenRealtimeKPI() {
            if (!kpiTargetUser || !selectedKpiMonth) return;
            
            const targetUserInfo = registeredUsers.find(u => u.username === kpiTargetUser);
            if (targetUserInfo) {
                const sel = document.getElementById('select-kpi-type-change');
                if (sel) sel.value = targetUserInfo.kpiType || 'staff';
            }
            
            if (isAdmin) {
                document.getElementById('btn-save-kpi-admin').style.display = 'inline-block';
                document.getElementById('btn-add-main-section').style.display = 'inline-block';
                if (currentUser === kpiTargetUser) {
                    document.getElementById('btn-save-kpi-user').style.display = 'inline-block';
                } else {
                    document.getElementById('btn-save-kpi-user').style.display = 'none';
                }
            } else {
                document.getElementById('btn-save-kpi-admin').style.display = 'none';
                document.getElementById('btn-add-main-section').style.display = 'none';
                document.getElementById('btn-save-kpi-user').style.display = 'inline-block';
            }
            
            loadAndMergeUserKPI(kpiTargetUser, selectedKpiMonth);
        }

        function loadAndMergeUserKPI(username, monthKey) {
            const targetUserInfo = registeredUsers.find(u => u.username === username);
            const userKpiType = (targetUserInfo && targetUserInfo.kpiType) ? targetUserInfo.kpiType : 'staff';
            const currentTemplate = masterKPITemplates[userKpiType] || masterKPITemplates['staff'];

            db.ref(`kpi/${monthKey}/${username}`).on('value', snapshot => {
                const data = snapshot.val();
                sectionMaxScores = currentTemplate.sectionMaxScores || (userKpiType === 'cleaner' ? { A: 70, B: 30 } : { A: 50, B: 30, C: 20 });
                
                const defaultTitles = userKpiType === 'cleaner' ? 
                    { A: "CÔNG TÁC CHUYÊN MÔN & VỆ SINH", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT" } : 
                    { A: "QUẢN LÝ & CHUYÊN MÔN", B: "Ý THỨC CHẤP HÀNH & KỶ LUẬT", C: "NĂNG LỰC ĐỔI MỚI & NHIỆM VỤ ĐỘT XUẤT" };

                sectionTitles = currentTemplate.sectionTitles || defaultTitles;

                if (data && data.kpiDataList) {
                    lastKpiTimestamp = data.timestamp || 'Chưa ghi nhận';
                    isKpiSavedForCurrentMonth = true;
                    kpiDataList = mergeKPIWithTemplate(currentTemplate.kpiDataList, data.kpiDataList);
                } else {
                    lastKpiTimestamp = 'Chưa có dữ liệu';
                    isKpiSavedForCurrentMonth = false;
                    kpiDataList = JSON.parse(JSON.stringify(currentTemplate.kpiDataList));
                }
                renderKPITable();
            });
        }

        function mergeKPIWithTemplate(templateList, userList) {
            const templateCopy = JSON.parse(JSON.stringify(templateList || []));
            const userItemMap = {};

            if (userList) {
                userList.forEach(sub => {
                    if (sub.items) {
                        sub.items.forEach(item => {
                            userItemMap[item.id] = {
                                selfScore: item.selfScore,
                                adminScore: item.adminScore
                            };
                        });
                    }
                });
            }

            templateCopy.forEach(sub => {
                if (sub.items) {
                    sub.items.forEach(item => {
                        if (userItemMap[item.id]) {
                            item.selfScore = userItemMap[item.id].selfScore !== undefined ? userItemMap[item.id].selfScore : item.maxScore;
                            item.adminScore = userItemMap[item.id].adminScore !== undefined ? userItemMap[item.id].adminScore : item.maxScore;
                        }
                    });
                }
            });

            return templateCopy;
        }

        function saveUserData() {
            if (!targetUser) return;
            db.ref(`tasks/${targetUser}`).set(tasks);
        }

        function saveKPIRatingStorage(timestamp = null) {
            if (!kpiTargetUser || !selectedKpiMonth) return;
            let timeSaved = timestamp || lastKpiTimestamp;
            db.ref(`kpi/${selectedKpiMonth}/${kpiTargetUser}`).set({
                timestamp: timeSaved,
                kpiDataList: kpiDataList
            }, (err) => {
                if (!err) {
                    isKpiSavedForCurrentMonth = true;
                }
            });
        }

        function isTaskOverdue(dateStr, status) {
            if (status === 'Hoàn thành' || !dateStr) return false;
            let day, month, year;
            if (dateStr.includes('-')) {
                [year, month, day] = dateStr.split('-');
            } else if (dateStr.includes('/')) {
                [day, month, year] = dateStr.split('/');
            } else {
                return false;
            }
            const dueDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), 23, 59, 59);
            return dueDate < new Date();
        }

        function toggleAuthTab(tab) {
            const loginForm = document.getElementById('login-form');
            const registerForm = document.getElementById('register-form');
            const loginBtn = document.getElementById('tab-login-btn');
            const regBtn = document.getElementById('tab-register-btn');
            const subTitle = document.getElementById('form-sub-title');

            if (tab === 'login') {
                loginForm.style.display = 'block'; registerForm.style.display = 'none';
                loginBtn.style.color = '#2563eb'; loginBtn.style.borderBottom = '2px solid #2563eb';
                regBtn.style.color = '#64748b'; regBtn.style.borderBottom = 'none';
                subTitle.innerText = 'Đăng nhập hệ thống quản trị công việc (Realtime)';
            } else {
                loginForm.style.display = 'none'; registerForm.style.display = 'block';
                regBtn.style.color = '#10b981'; regBtn.style.borderBottom = '2px solid #10b981';
                loginBtn.style.color = '#64748b'; loginBtn.style.borderBottom = 'none';
                subTitle.innerText = 'Tạo tài khoản mới cho cán bộ';
            }
        }

        function handleRegister(e) {
            e.preventDefault();
            const email = document.getElementById('reg-email').value.trim();
            const username = document.getElementById('reg-username').value.trim().toLowerCase();
            const password = document.getElementById('reg-password').value;
            const confirmPassword = document.getElementById('reg-confirm-password').value;
            const kpiType = document.getElementById('reg-kpi-type').value;

            if (username === ADMIN_USERNAME) { alert('Tên đăng nhập trùng với Admin!'); return; }
            if (password !== confirmPassword) { alert('Mật khẩu không khớp!'); return; }
            if (registeredUsers.some(u => u.username === username)) { alert('Tên đăng nhập đã tồn tại!'); return; }

            const newUser = { email, username, password, kpiType: kpiType };
            db.ref(`users/${username}`).set(newUser, (err) => {
                if (!err) {
                    alert('Đăng ký tài khoản thành công!');
                    document.getElementById('register-form').reset();
                    toggleAuthTab('login');
                    document.getElementById('login-username').value = username;
                } else {
                    alert('Lỗi đăng ký: ' + err.message);
                }
            });
        }

        function handleLogin(e) {
            e.preventDefault();
            const usernameInput = document.getElementById('login-username').value.trim().toLowerCase();
            const passwordInput = document.getElementById('login-password').value;

            const validUser = registeredUsers.find(u => u.username === usernameInput && u.password === passwordInput);

            if (usernameInput === ADMIN_USERNAME && passwordInput === ADMIN_PASSWORD) {
                currentUser = ADMIN_USERNAME;
                isAdmin = true;
            } else if (validUser) {
                currentUser = validUser.username;
                isAdmin = false;
            } else {
                alert('Tên đăng nhập hoặc mật khẩu không đúng!');
                return;
            }

            targetUser = currentUser;
            kpiTargetUser = currentUser;

            document.getElementById('login-screen').style.display = 'none';
            document.getElementById('app-screen').style.display = 'flex';
            document.getElementById('user-display-name').innerText = currentUser;
            document.getElementById('user-role-badge').innerText = isAdmin ? 'Admin Root' : 'Cán bộ';

            if (isAdmin) {
                document.getElementById('admin-user-selector').style.display = 'flex';
                document.getElementById('kpi-admin-target-bar').style.display = 'flex';
                document.getElementById('nav-admin-users').style.display = 'flex';
                document.getElementById('admin-master-overview-banner').style.display = 'block';
                populateAdminUserSelector();
                populateKPITargetSelector();
                renderUserManagementTable();
            } else {
                document.getElementById('admin-user-selector').style.display = 'none';
                document.getElementById('kpi-admin-target-bar').style.display = 'none';
                document.getElementById('nav-admin-users').style.display = 'none';
                document.getElementById('admin-master-overview-banner').style.display = 'none';
            }

            listenRealtimeTasks();
            listenRealtimeKPI();
        }

        function logout() {
            currentUser = '';
            targetUser = '';
            kpiTargetUser = '';
            isAdmin = false;
            document.getElementById('app-screen').style.display = 'none';
            document.getElementById('login-screen').style.display = 'flex';
            document.getElementById('login-form').reset();
        }

        function populateAdminUserSelector() {
            const sel = document.getElementById('select-target-user');
            if (!sel) return;
            sel.innerHTML = '';
            registeredUsers.forEach(u => {
                const opt = document.createElement('option');
                opt.value = u.username;
                opt.innerText = u.username;
                if (u.username === targetUser) opt.selected = true;
                sel.appendChild(opt);
            });
        }

        function populateKPITargetSelector() {
            const sel = document.getElementById('select-kpi-target-user');
            if (!sel) return;
            sel.innerHTML = '';
            registeredUsers.forEach(u => {
                const opt = document.createElement('option');
                opt.value = u.username;
                opt.innerText = u.username;
                if (u.username === kpiTargetUser) opt.selected = true;
                sel.appendChild(opt);
            });
        }

        function changeTargetUser(val) {
            targetUser = val;
            listenRealtimeTasks();
        }

        function changeKPITargetUser(val) {
            kpiTargetUser = val;
            listenRealtimeKPI();
        }

        function adminChangeUserKPIType(newType) {
            if (!isAdmin || !kpiTargetUser) return;
            db.ref(`users/${kpiTargetUser}/kpiType`).set(newType, (err) => {
                if (!err) {
                    alert(`Đã đổi Bảng KPI của ${kpiTargetUser} thành ${newType.toUpperCase()}`);
                    listenRealtimeKPI();
                }
            });
        }

        function switchTab(tabId, el) {
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            
            document.getElementById(`tab-${tabId}`).classList.add('active');
            if (el) el.classList.add('active');
            
            const titles = {
                'tong-quan': 'Dashboard Thống Kê',
                'danh-sach': 'Danh Sách Công Việc',
                'kanban': 'Bảng Tiến Độ Công Việc',
                'gantt': 'Sơ Đồ Lộ Trình Gantt',
                'kpi': 'Đánh Giá KPI Cán Bộ',
                'admin-users': 'Quản Lý Hệ Thống Users'
            };
            document.getElementById('page-title').innerText = titles[tabId] || 'WORKSPACE';
            
            if (tabId === 'tong-quan') {
                renderDashboard();
            } else if (tabId === 'admin-users') {
                renderUserManagementTable();
            }
        }

        function renderDashboard() {
            const statusCounts = { 'Chưa làm': 0, 'Đang làm': 0, 'Hoàn thành': 0 };
            const priorityCounts = { 'Cao': 0, 'Bình thường': 0, 'Thấp': 0 };

            if (isAdmin) {
                // TỔNG HỢP TOÀN BỘ DỮ LIỆU CÔNG VIỆC CỦA CÁC TÀI KHOẢN THƯỜNG TRÊN WEB
                Object.keys(allUsersTasksMap).forEach(username => {
                    const userTasks = allUsersTasksMap[username];
                    if (userTasks) {
                        const taskList = Array.isArray(userTasks) ? userTasks : Object.values(userTasks);
                        taskList.forEach(t => {
                            if (statusCounts[t.status] !== undefined) statusCounts[t.status]++;
                            if (priorityCounts[t.priority] !== undefined) priorityCounts[t.priority]++;
                        });
                    }
                });
            } else {
                tasks.forEach(t => {
                    if (statusCounts[t.status] !== undefined) statusCounts[t.status]++;
                    if (priorityCounts[t.priority] !== undefined) priorityCounts[t.priority]++;
                });
            }

            const ctxStatus = document.getElementById('statusChart').getContext('2d');
            if (statusChartInstance) statusChartInstance.destroy();
            statusChartInstance = new Chart(ctxStatus, {
                type: 'doughnut',
                data: {
                    labels: Object.keys(statusCounts),
                    datasets: [{
                        data: Object.values(statusCounts),
                        backgroundColor: ['#a855f7', '#3b82f6', '#22c55e']
                    }]
                },
                options: { responsive: true, maintainAspectRatio: false }
            });

            const ctxPriority = document.getElementById('priorityChart').getContext('2d');
            if (priorityChartInstance) priorityChartInstance.destroy();
            priorityChartInstance = new Chart(ctxPriority, {
                type: 'bar',
                data: {
                    labels: Object.keys(priorityCounts),
                    datasets: [{
                        label: 'Số lượng công việc',
                        data: Object.values(priorityCounts),
                        backgroundColor: ['#ef4444', '#f59e0b', '#10b981']
                    }]
                },
                options: { responsive: true, maintainAspectRatio: false }
            });
        }

        function renderTaskList() {
            const tbody = document.getElementById('task-table-body');
            tbody.innerHTML = '';

            tasks.forEach(t => {
                const overdue = isTaskOverdue(t.date, t.status);
                const tr = document.createElement('tr');
                if (overdue) tr.classList.add('row-overdue');

                const statusClass = t.status === 'Đang làm' ? 'status-doing' : t.status === 'Chưa làm' ? 'status-todo' : 'status-done';

                // Xử lý hiển thị danh sách file đính kèm
                let fileListHTML = '';
                if (t.files && t.files.length > 0) {
                    fileListHTML = t.files.map((f, index) => `
                        <div class="file-tag">
                            <i class="fa-solid fa-paperclip"></i> 
                            <a href="${f.data}" download="${f.name}" title="Tải file về máy">${f.name}</a>
                            ${!isAdmin ? `<i class="fa-solid fa-xmark" style="cursor:pointer; color:#ef4444; margin-left:4px;" onclick="removeTaskFile('${t.id}', ${index})" title="Xóa file"></i>` : ''}
                        </div>
                    `).join('<br>');
                } else {
                    fileListHTML = '<span style="color:#94a3b8; font-size:12px;">Chưa có file</span>';
                }

                tr.innerHTML = `
                    <td><strong>${t.id}</strong></td>
                    <td>${t.name} ${overdue ? '<span class="badge-overdue"><i class="fa-solid fa-triangle-exclamation"></i> Quá hạn</span>' : ''}</td>
                    <td>${targetUser}</td>
                    <td>${t.priority}</td>
                    <td>
                        <select onchange="updateTaskStatus('${t.id}', this.value)" class="status-tag ${statusClass}">
                            <option value="Chưa làm" ${t.status === 'Chưa làm' ? 'selected' : ''}>Chưa làm</option>
                            <option value="Đang làm" ${t.status === 'Đang làm' ? 'selected' : ''}>Đang làm</option>
                            <option value="Hoàn thành" ${t.status === 'Hoàn thành' ? 'selected' : ''}>Hoàn thành</option>
                        </select>
                    </td>
                    <td>
                        <input type="date" value="${convertToISODate(t.date)}" class="inline-date-picker" onchange="updateTaskDate('${t.id}', this.value)" />
                    </td>
                    <td>
                        <div id="file-container-${t.id}">${fileListHTML}</div>
                    </td>
                    <td>
                        ${!isAdmin ? `
                            <input type="file" id="file-input-${t.id}" style="display:none;" onchange="uploadTaskFile('${t.id}', this)" />
                            <button class="btn-sm btn-attach" onclick="document.getElementById('file-input-${t.id}').click()"><i class="fa-solid fa-paperclip"></i> File đính kèm</button>
                        ` : ''}
                        <button class="btn-sm btn-delete" onclick="deleteTask('${t.id}')"><i class="fa-solid fa-trash"></i> Xóa</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        // XỬ LÝ UPLOAD FILE LÊN TÍCH HỢP CHO USER
        function uploadTaskFile(taskId, inputEl) {
            const file = inputEl.files[0];
            if (!file) return;

            // Kiểm tra kích thước file (khống chế tối đa 5MB để giữ tốc độ sync Database)
            if (file.size > 5 * 1024 * 1024) {
                alert('Dung lượng file vượt quá 5MB. Vui lòng chọn file có kích thước nhỏ hơn!');
                inputEl.value = '';
                return;
            }

            const reader = new FileReader();
            reader.onload = function(e) {
                const base64Data = e.target.result;
                const task = tasks.find(t => t.id === taskId);
                if (task) {
                    if (!task.files) task.files = [];
                    task.files.push({
                        name: file.name,
                        data: base64Data,
                        uploadedAt: new Date().toISOString()
                    });
                    saveUserData();
                    alert(`Đã đính kèm file "${file.name}" thành công!`);
                }
            };
            reader.readAsDataURL(file);
            inputEl.value = '';
        }

        // XÓA FILE ĐÍNH KÈM
        function removeTaskFile(taskId, fileIndex) {
            if (confirm('Bạn có chắc chắn muốn xóa file đính kèm này?')) {
                const task = tasks.find(t => t.id === taskId);
                if (task && task.files) {
                    task.files.splice(fileIndex, 1);
                    saveUserData();
                }
            }
        }

        function updateTaskStatus(id, newStatus) {
            const task = tasks.find(t => t.id === id);
            if (task) {
                task.status = newStatus;
                saveUserData();
            }
        }

        function updateTaskDate(id, isoDateStr) {
            const task = tasks.find(t => t.id === id);
            if (task) {
                task.date = convertToDisplayDate(isoDateStr);
                saveUserData();
            }
        }

        function addNewTask() {
            const name = document.getElementById('newTaskName').value.trim();
            const priority = document.getElementById('newTaskPriority').value;
            const dateISO = document.getElementById('newTaskDueDate').value;

            if (!name) { alert('Vui lòng nhập tên công việc!'); return; }

            const newId = 'T' + String(tasks.length + 1).padStart(3, '0');
            const newTask = {
                id: newId,
                name: name,
                priority: priority,
                status: 'Chưa làm',
                date: convertToDisplayDate(dateISO),
                files: []
            };

            tasks.push(newTask);
            saveUserData();
            document.getElementById('newTaskName').value = '';
        }

        function deleteTask(id) {
            if (confirm('Bạn có chắc chắn muốn xóa công việc này?')) {
                tasks = tasks.filter(t => t.id !== id);
                saveUserData();
            }
        }

        function renderKanban() {
            const cardsTodo = document.getElementById('cards-todo');
            const cardsDoing = document.getElementById('cards-doing');
            const cardsDone = document.getElementById('cards-done');

            cardsTodo.innerHTML = '';
            cardsDoing.innerHTML = '';
            cardsDone.innerHTML = '';

            let cTodo = 0, cDoing = 0, cDone = 0;

            tasks.forEach(t => {
                const overdue = isTaskOverdue(t.date, t.status);
                const card = document.createElement('div');
                card.className = `kanban-card ${overdue ? 'card-overdue' : ''}`;
                card.draggable = true;
                card.ondragstart = (e) => e.dataTransfer.setData('text/plain', t.id);

                let hasFilesBadge = (t.files && t.files.length > 0) ? `<span style="color:#2563eb; font-size:11px;"><i class="fa-solid fa-paperclip"></i> ${t.files.length}</span>` : '';

                card.innerHTML = `
                    <div class="id">${t.id} ${overdue ? '<span class="badge-overdue">Quá hạn</span>' : ''}</div>
                    <div class="title">${t.name}</div>
                    <div class="meta">
                        <span><i class="fa-regular fa-clock"></i> ${t.date || 'N/A'} ${hasFilesBadge}</span>
                        <span style="font-weight: bold;">${t.priority}</span>
                    </div>
                `;

                if (t.status === 'Chưa làm') { cardsTodo.appendChild(card); cTodo++; }
                else if (t.status === 'Đang làm') { cardsDoing.appendChild(card); cDoing++; }
                else if (t.status === 'Hoàn thành') { cardsDone.appendChild(card); cDone++; }
            });

            document.getElementById('count-todo').innerText = cTodo;
            document.getElementById('count-doing').innerText = cDoing;
            document.getElementById('count-done').innerText = cDone;
        }

        function allowDrop(e) { e.preventDefault(); }
        function drop(e, targetStatus) {
            e.preventDefault();
            const id = e.dataTransfer.getData('text/plain');
            const task = tasks.find(t => t.id === id);
            if (task) {
                task.status = targetStatus;
                saveUserData();
            }
        }

        // HÀM RENDER ĐỘNG TẤT CẢ CÁC MỤC LỚN (A, B, C, D...)
        function renderKPITable() {
            const targetUserInfo = registeredUsers.find(u => u.username === kpiTargetUser);
            const userKpiType = (targetUserInfo && targetUserInfo.kpiType) ? targetUserInfo.kpiType : 'staff';
            
            document.getElementById('kpi-target-name-display').innerText = kpiTargetUser || '';
            const badgeEl = document.getElementById('kpi-type-badge-display');
            
            const badgeMap = {
                leader: { text: 'BẢNG LÃNH ĐẠO', class: 'type-leader' },
                staff: { text: 'BẢNG CÁN BỘ', class: 'type-staff' },
                cleaner: { text: 'BẢNG LAO CÔNG', class: 'type-cleaner' }
            };
            
            badgeEl.innerText = badgeMap[userKpiType].text;
            badgeEl.className = `kpi-badge-type ${badgeMap[userKpiType].class}`;

            const wrapper = document.getElementById('kpi-sections-wrapper');
            wrapper.innerHTML = '';

            const bgColors = { A: '#1e40af', B: '#0d9488', C: '#7c3aed', D: '#b45309', E: '#4c1d95', F: '#0369a1' };

            Object.keys(sectionMaxScores).forEach(secKey => {
                const secTitle = sectionTitles[secKey] || '';
                const maxScore = sectionMaxScores[secKey] || 0;
                const headerBg = bgColors[secKey] || '#1e40af';

                const secDiv = document.createElement('div');
                secDiv.style.marginBottom = '25px';
                secDiv.id = `kpi-section-container-${secKey}`;

                secDiv.innerHTML = `
                    <div class="kpi-section-title" style="background-color: ${headerBg};">
                        <span>
                            ${secKey}. <span id="section-title-text-${secKey}">${secTitle}</span> (ĐIỂM TỐI ĐA: <span id="max-score-${secKey}-display" style="color: #fde047;">${maxScore}</span> ĐIỂM)
                            ${isAdmin ? `
                                <button class="btn-sm btn-edit" style="margin-left: 10px;" onclick="editSectionTitle('${secKey}')"><i class="fa-solid fa-pen-to-square"></i> Sửa Tên</button>
                                <button class="btn-sm btn-edit" style="margin-left: 5px;" onclick="editSectionMaxScore('${secKey}')"><i class="fa-solid fa-pen"></i> Sửa Điểm</button>
                                <button class="btn-sm btn-delete" style="margin-left: 5px;" onclick="deleteMainSection('${secKey}')"><i class="fa-solid fa-trash"></i> Xóa Mục Lớn</button>
                            ` : ''}
                        </span>
                        ${isAdmin ? `
                            <button class="btn-sm btn-login" style="width: auto; background-color: #10b981;" onclick="addSubSection('${secKey}')">
                                <i class="fa-solid fa-plus"></i> Thêm Mục La Mã (I, II...)
                            </button>
                        ` : ''}
                    </div>
                    <div class="table-container" style="border-radius: 0 0 8px 8px;">
                        <table class="kpi-table">
                            <thead>
                                <tr>
                                    <th style="width: 50px; text-align: center;">STT</th>
                                    <th>Nội dung Đánh giá</th>
                                    <th style="width: 250px;">Tiêu chí đánh giá</th>
                                    <th style="width: 100px; text-align: center;">Điểm tối đa</th>
                                    <th style="width: 100px; text-align: center;">Điểm tự chấm</th>
                                    <th style="width: 120px; text-align: center; background-color: #f0fdf4; color: #166534;">Điểm Đánh Giá</th>
                                    ${isAdmin ? `<th style="width: 160px; text-align: center;">Thao tác Admin</th>` : ''}
                                </tr>
                            </thead>
                            <tbody id="kpi-tbody-${secKey}"></tbody>
                        </table>
                    </div>
                `;

                wrapper.appendChild(secDiv);

                const tbody = document.getElementById(`kpi-tbody-${secKey}`);
                const sectionSubs = kpiDataList.filter(s => s.section === secKey);

                sectionSubs.forEach((sub) => {
                    const subTr = document.createElement('tr');
                    subTr.className = 'row-sub-header';
                    subTr.innerHTML = `
                        <td style="text-align: center;"><strong>${sub.code}</strong></td>
                        <td colspan="5"><strong>${sub.title}</strong></td>
                        ${isAdmin ? `<td style="text-align: center;">
                            <button class="btn-sm btn-login" style="background-color: #10b981;" onclick="addItemToSubSection('${sub.id}')"><i class="fa-solid fa-plus"></i> Con</button>
                            <button class="btn-sm btn-edit" onclick="editSubSectionTitle('${sub.id}')"><i class="fa-solid fa-pen"></i> Sửa</button>
                            <button class="btn-sm btn-delete" onclick="deleteSubSection('${sub.id}')"><i class="fa-solid fa-trash"></i> Xóa</button>
                        </td>` : ''}
                    `;
                    tbody.appendChild(subTr);

                    if (sub.items) {
                        sub.items.forEach((item, itemIdx) => {
                            const itemTr = document.createElement('tr');
                            itemTr.innerHTML = `
                                <td style="text-align: center;">${itemIdx + 1}</td>
                                <td>${item.title}</td>
                                <td>${item.criteria || ''}</td>
                                <td style="text-align: center;"><strong>${item.maxScore}</strong></td>
                                <td style="text-align: center;">
                                    <input type="number" step="0.5" max="${item.maxScore}" min="0" value="${item.selfScore}" onchange="updateKPIScore('${sub.id}', '${item.id}', 'selfScore', this.value)" ${(!isAdmin && currentUser !== kpiTargetUser) ? 'disabled' : ''}>
                                </td>
                                <td style="text-align: center; background-color: #f0fdf4;">
                                    <input type="number" step="0.5" max="${item.maxScore}" min="0" value="${item.adminScore}" onchange="updateKPIScore('${sub.id}', '${item.id}', 'adminScore', this.value)" ${!isAdmin ? 'disabled' : ''}>
                                </td>
                                ${isAdmin ? `<td style="text-align: center;">
                                    <button class="btn-sm btn-edit" onclick="editItem('${sub.id}', '${item.id}')"><i class="fa-solid fa-pen"></i> Sửa</button>
                                    <button class="btn-sm btn-delete" onclick="deleteItem('${sub.id}', '${item.id}')"><i class="fa-solid fa-trash"></i> Xóa</button>
                                </td>` : ''}
                            `;
                            tbody.appendChild(itemTr);
                        });
                    }
                });
            });

            calculateKPITotals();
            document.getElementById('kpi-last-time-saved').innerText = lastKpiTimestamp;
        }

        // TÍNH TỔNG ĐIỂM KPI THEO DỮ LIỆU ĐỘNG
        function calculateKPITotals() {
            let totalSelf = 0;
            let totalAdmin = 0;
            let totalMax = 0;

            Object.keys(sectionMaxScores).forEach(k => totalMax += parseFloatStrict(sectionMaxScores[k]));

            kpiDataList.forEach(sub => {
                if (sub.items) {
                    sub.items.forEach(item => {
                        totalSelf += parseFloatStrict(item.selfScore);
                        totalAdmin += parseFloatStrict(item.adminScore);
                    });
                }
            });

            document.getElementById('kpi-total-self').innerText = totalSelf.toFixed(1);
            document.getElementById('kpi-total-admin').innerText = totalAdmin.toFixed(1);
            document.getElementById('kpi-total-max').innerText = totalMax.toFixed(1);
        }

        // THÊM MỤC LỚN (A, B, C, D...) DÀNH CHO ADMIN
        function addMainSection() {
            if (!isAdmin) return;
            const secKeyInput = prompt("Nhập ký tự cho Mục Lớn mới (VD: D, E, F...):");
            if (!secKeyInput) return;
            const secKey = secKeyInput.trim().toUpperCase();

            if (sectionMaxScores[secKey] !== undefined) {
                alert(`Mục ${secKey} đã tồn tại!`);
                return;
            }

            const titleInput = prompt(`Nhập tiêu đề cho Mục Lớn ${secKey}:`);
            if (!titleInput) return;

            const scoreInput = prompt(`Nhập điểm tối đa cho Mục Lớn ${secKey}:`, "10");
            const maxScore = parseFloatStrict(scoreInput) || 10;

            sectionTitles[secKey] = titleInput.trim();
            sectionMaxScores[secKey] = maxScore;

            syncKPISettingToTemplate();
        }

        // XÓA MỤC LỚN DÀNH CHO ADMIN
        function deleteMainSection(secKey) {
            if (!isAdmin) return;
            if (confirm(`Bạn có chắc chắn muốn xóa Mục Lớn ${secKey} cùng tất cả tiêu chí bên trong?`)) {
                delete sectionTitles[secKey];
                delete sectionMaxScores[secKey];

                kpiDataList = kpiDataList.filter(s => s.section !== secKey);
                syncKPISettingToTemplate();
            }
        }

        function updateKPIScore(subId, itemId, scoreType, value) {
            const sub = kpiDataList.find(s => s.id === subId);
            if (sub && sub.items) {
                const item = sub.items.find(i => i.id === itemId);
                if (item) {
                    item[scoreType] = parseFloatStrict(value);
                    calculateKPITotals();
                }
            }
        }

        function saveKPIRatingByUser() {
            const nowStr = new Date().toLocaleString('vi-VN');
            lastKpiTimestamp = `${nowStr} (Bởi cá nhân tự chấm)`;
            saveKPIRatingStorage(lastKpiTimestamp);
            alert('Đã lưu kết quả tự chấm KPI thành công!');
        }

        function saveKPIRatingByAdmin() {
            const nowStr = new Date().toLocaleString('vi-VN');
            lastKpiTimestamp = `${nowStr} (Bởi Admin)`;
            saveKPIRatingStorage(lastKpiTimestamp);
            alert(`Đã lưu kết quả đánh giá KPI cho ${kpiTargetUser} thành công!`);
        }

        function editSectionTitle(secKey) {
            const current = sectionTitles[secKey] || '';
            const newTitle = prompt('Nhập tên mới cho Phần ' + secKey + ':', current);
            if (newTitle !== null && newTitle.trim() !== '') {
                sectionTitles[secKey] = newTitle.trim();
                syncKPISettingToTemplate();
            }
        }

        function editSectionMaxScore(secKey) {
            const current = sectionMaxScores[secKey] || 0;
            const newScore = prompt('Nhập Điểm tối đa mới cho Phần ' + secKey + ':', current);
            if (newScore !== null && isValidNumber(newScore)) {
                sectionMaxScores[secKey] = parseFloatStrict(newScore);
                syncKPISettingToTemplate();
            }
        }

        function addSubSection(secKey) {
            const code = prompt('Nhập mã số mục La Mã (I, II, III...):', 'I');
            if (!code) return;
            const title = prompt('Nhập tiêu đề mục La Mã mới:');
            if (!title) return;

            const newSub = {
                id: 'sub_' + Date.now(),
                section: secKey,
                code: code.trim(),
                title: title.trim(),
                maxScore: 10,
                items: []
            };

            kpiDataList.push(newSub);
            syncKPISettingToTemplate();
        }

        function editSubSectionTitle(subId) {
            const sub = kpiDataList.find(s => s.id === subId);
            if (!sub) return;
            const newTitle = prompt('Sửa tiêu đề mục La Mã:', sub.title);
            if (newTitle !== null && newTitle.trim() !== '') {
                sub.title = newTitle.trim();
                syncKPISettingToTemplate();
            }
        }

        function deleteSubSection(subId) {
            if (confirm('Bạn có chắc muốn xóa mục La Mã này và tất cả tiêu chí con bên trong?')) {
                kpiDataList = kpiDataList.filter(s => s.id !== subId);
                syncKPISettingToTemplate();
            }
        }

        function addItemToSubSection(subId) {
            const sub = kpiDataList.find(s => s.id === subId);
            if (!sub) return;
            const title = prompt('Nhập nội dung Đánh giá:');
            if (!title) return;
            const criteria = prompt('Nhập tiêu chí đánh giá (nếu có):', '');
            const maxScoreStr = prompt('Nhập điểm tối đa cho tiêu chí này:', '5');
            const maxScore = parseFloatStrict(maxScoreStr) || 5;

            if (!sub.items) sub.items = [];
            sub.items.push({
                id: 'item_' + Date.now(),
                title: title.trim(),
                criteria: criteria ? criteria.trim() : '',
                maxScore: maxScore,
                selfScore: maxScore,
                adminScore: maxScore
            });

            syncKPISettingToTemplate();
        }

        function editItem(subId, itemId) {
            const sub = kpiDataList.find(s => s.id === subId);
            if (!sub || !sub.items) return;
            const item = sub.items.find(i => i.id === itemId);
            if (!item) return;

            const newTitle = prompt('Sửa nội dung Đánh giá:', item.title);
            if (newTitle === null) return;
            const newCriteria = prompt('Sửa tiêu chí đánh giá:', item.criteria || '');
            if (newCriteria === null) return;
            const newMaxStr = prompt('Sửa điểm tối đa:', item.maxScore);
            if (newMaxStr === null) return;

            item.title = newTitle.trim();
            item.criteria = newCriteria.trim();
            item.maxScore = parseFloatStrict(newMaxStr);
            if (item.selfScore > item.maxScore) item.selfScore = item.maxScore;
            if (item.adminScore > item.maxScore) item.adminScore = item.maxScore;

            syncKPISettingToTemplate();
        }

        function deleteItem(subId, itemId) {
            const sub = kpiDataList.find(s => s.id === subId);
            if (!sub || !sub.items) return;
            if (confirm('Bạn có chắc chắn muốn xóa tiêu chí này?')) {
                sub.items = sub.items.filter(i => i.id !== itemId);
                syncKPISettingToTemplate();
            }
        }

        function syncKPISettingToTemplate() {
            const targetUserInfo = registeredUsers.find(u => u.username === kpiTargetUser);
            const userKpiType = (targetUserInfo && targetUserInfo.kpiType) ? targetUserInfo.kpiType : 'staff';
            
            masterKPITemplates[userKpiType].sectionTitles = sectionTitles;
            masterKPITemplates[userKpiType].sectionMaxScores = sectionMaxScores;
            masterKPITemplates[userKpiType].kpiDataList = kpiDataList;

            db.ref(`kpiTemplates/${userKpiType}`).set(masterKPITemplates[userKpiType]);
            renderKPITable();
        }

        // HÀM HIỂN THỊ DANH SÁCH USER TRÊN GIAO DIỆN QUẢN LÝ TÀI KHOẢN
        function renderUserManagementTable() {
            const tbody = document.getElementById('user-management-body');
            if (!tbody) return;
            tbody.innerHTML = '';

            registeredUsers.forEach((u, idx) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${idx + 1}</td>
                    <td><strong>${u.username}</strong> ${u.username === ADMIN_USERNAME ? '<span style="color:#ef4444; font-size:11px;">(Admin Root)</span>' : ''}</td>
                    <td>${u.email || 'N/A'}</td>
                    <td>
                        <select onchange="changeUserKPITypeInTable('${u.username}', this.value)" style="padding:4px 8px; border-radius:4px; border:1px solid #cbd5e1;">
                            <option value="staff" ${(u.kpiType || 'staff') === 'staff' ? 'selected' : ''}>Bảng Cán bộ / Nhân viên</option>
                            <option value="leader" ${u.kpiType === 'leader' ? 'selected' : ''}>Bảng Lãnh đạo</option>
                            <option value="cleaner" ${u.kpiType === 'cleaner' ? 'selected' : ''}>Bảng Lao Công</option>
                        </select>
                    </td>
                    <td>
                        ${u.username !== ADMIN_USERNAME ? `<button class="btn-sm btn-delete" onclick="deleteUserAccount('${u.username}')"><i class="fa-solid fa-user-xmark"></i> Xóa tài khoản</button>` : '<em>Không thể xóa</em>'}
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function changeUserKPITypeInTable(username, newType) {
            db.ref(`users/${username}/kpiType`).set(newType, (err) => {
                if (!err) {
                    alert(`Đã cập nhật Bảng KPI cho ${username} thành ${newType.toUpperCase()}`);
                }
            });
        }

        function deleteUserAccount(username) {
            if (confirm(`Bạn có chắc chắn muốn xóa vĩnh viễn tài khoản ${username} khỏi hệ thống?`)) {
                db.ref(`users/${username}`).remove();
                db.ref(`tasks/${username}`).remove();
            }
        }

        /* --- XUẤT FILE WORD XỬ LÝ ĐỘNG TẤT CẢ MỤC LỚN --- */
        async function exportKPIWord() {
            if (typeof docx === 'undefined') {
                alert('Hệ thống chưa tải xong thư viện docx. Vui lòng kiểm tra kết nối mạng và thử lại!');
                return;
            }

            const {
                Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, VerticalAlign, BorderStyle
            } = docx;

            const targetUserInfo = registeredUsers.find(u => u.username === kpiTargetUser);
            const userKpiType = (targetUserInfo && targetUserInfo.kpiType) ? targetUserInfo.kpiType : 'staff';
            
            const kpiTypeName = userKpiType === 'leader' ? 'Lãnh đạo' : (userKpiType === 'cleaner' ? 'Lao Công' : 'Cán bộ/Nhân viên');

            let totalSelf = 0, totalAdmin = 0, totalMax = 0;
            Object.keys(sectionMaxScores).forEach(k => totalMax += parseFloatStrict(sectionMaxScores[k]));

            kpiDataList.forEach(sub => {
                if (sub.items) {
                    sub.items.forEach(item => {
                        totalSelf += parseFloatStrict(item.selfScore);
                        totalAdmin += parseFloatStrict(item.adminScore);
                    });
                }
            });

            const tableHeaderRow = new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({
                        width: { size: 8, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STT", bold: true, font: "Times New Roman", size: 22 })] })]
                    }),
                    new TableCell({
                        width: { size: 56, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nội dung Đánh giá", bold: true, font: "Times New Roman", size: 22 })] })]
                    }),
                    new TableCell({
                        width: { size: 12, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Điểm tối đa", bold: true, font: "Times New Roman", size: 22 })] })]
                    }),
                    new TableCell({
                        width: { size: 12, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tự chấm", bold: true, font: "Times New Roman", size: 22 })] })]
                    }),
                    new TableCell({
                        width: { size: 12, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Đánh giá", bold: true, font: "Times New Roman", size: 22 })] })]
                    })
                ]
            });

            const tableRows = [tableHeaderRow];

            Object.keys(sectionMaxScores).forEach(secKey => {
                const secTitle = sectionTitles[secKey] || '';
                const maxSecScore = sectionMaxScores[secKey] || 0;

                const secRow = new TableRow({
                    children: [
                        new TableCell({
                            columnSpan: 5,
                            shading: { fill: "E2E8F0" },
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `${secKey}. ${secTitle.toUpperCase()} (ĐIỂM TỐI ĐA: ${maxSecScore} ĐIỂM)`,
                                            bold: true,
                                            font: "Times New Roman",
                                            size: 22
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                });
                tableRows.push(secRow);

                const sectionSubs = kpiDataList.filter(s => s.section === secKey);

                sectionSubs.forEach(sub => {
                    const subRow = new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 8, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sub.code, bold: true, font: "Times New Roman", size: 22 })] })]
                            }),
                            new TableCell({
                                columnSpan: 4,
                                children: [new Paragraph({ children: [new TextRun({ text: sub.title, bold: true, font: "Times New Roman", size: 22 })] })]
                            })
                        ]
                    });
                    tableRows.push(subRow);

                    if (sub.items) {
                        sub.items.forEach((item, idx) => {
                            const itemRow = new TableRow({
                                children: [
                                    new TableCell({
                                        width: { size: 8, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), font: "Times New Roman", size: 22 })] })]
                                    }),
                                    new TableCell({
                                        width: { size: 56, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ children: [new TextRun({ text: item.title, font: "Times New Roman", size: 22 })] })]
                                    }),
                                    new TableCell({
                                        width: { size: 12, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.maxScore), bold: true, font: "Times New Roman", size: 22 })] })]
                                    }),
                                    new TableCell({
                                        width: { size: 12, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.selfScore), font: "Times New Roman", size: 22 })] })]
                                    }),
                                    new TableCell({
                                        width: { size: 12, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.adminScore), font: "Times New Roman", size: 22 })] })]
                                    })
                                ]
                            });
                            tableRows.push(itemRow);
                        });
                    }
                });
            });

            const monthParts = selectedKpiMonth.split('-');
            const monthStr = `Tháng ${monthParts[1]} năm ${monthParts[0]}`;

            // Bảng Header 2 cột: Dòng Khoa Hóa lý căn giữa tương quan với TRUNG TÂM KSBT BẮC NINH
            const headerTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 45, type: WidthType.PERCENTAGE },
                                verticalAlign: VerticalAlign.TOP,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 50 },
                                        children: [
                                            new TextRun({ text: "TRUNG TÂM KSBT BẮC NINH", bold: true, font: "Times New Roman", size: 22 })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 200 },
                                        children: [
                                            new TextRun({ text: "KHOA HÓA LÝ", bold: true, font: "Times New Roman", size: 22 })
                                        ]
                                    })
                                ]
                            }),
                            new TableCell({
                                width: { size: 55, type: WidthType.PERCENTAGE },
                                verticalAlign: VerticalAlign.TOP,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 50 },
                                        children: [
                                            new TextRun({ text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold: true, font: "Times New Roman", size: 22 })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 200 },
                                        children: [
                                            new TextRun({ text: "Độc lập - Tự do - Hạnh phúc", bold: true, font: "Times New Roman", size: 22 })
                                        ]
                                    })
                                ]
                            })
                        ]
                    })
                ]
            });

            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            margin: { top: 1134, right: 1134, bottom: 1134, left: 1417 }
                        }
                    },
                    children: [
                        headerTable,
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 200, after: 100 },
                            children: [
                                new TextRun({ text: "PHIẾU THEO DÕI, ĐÁNH GIÁ VIÊN CHỨC", bold: true, font: "Times New Roman", size: 28 })
                            ]
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 300 },
                            children: [
                                new TextRun({ text: `(Kỳ đánh giá: ${monthStr})`, font: "Times New Roman", size: 24 })
                            ]
                        }),
                        new Paragraph({
                            spacing: { after: 100 },
                            children: [
                                new TextRun({ text: "Họ và tên cán bộ: ", bold: true, font: "Times New Roman", size: 24 }),
                                new TextRun({ text: kpiTargetUser.toUpperCase(), bold: true, font: "Times New Roman", size: 24 })
                            ]
                        }),
                        new Paragraph({
                            spacing: { after: 200 },
                            children: [
                                new TextRun({ text: "Chức vụ: ", bold: true, font: "Times New Roman", size: 24 }),
                                new TextRun({ text: kpiTypeName, font: "Times New Roman", size: 24 })
                            ]
                        }),
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: tableRows
                        }),
                        new Paragraph({
                            spacing: { before: 300, after: 400 },
                            children: [
                                new TextRun({ text: `TỔNG ĐIỂM ĐÁNH GIÁ: ${totalAdmin.toFixed(1)} / ${totalMax.toFixed(1)} ĐIỂM`, bold: true, font: "Times New Roman", size: 24 })
                            ]
                        }),
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            borders: {
                                top: { style: BorderStyle.NONE },
                                bottom: { style: BorderStyle.NONE },
                                left: { style: BorderStyle.NONE },
                                right: { style: BorderStyle.NONE },
                                insideHorizontal: { style: BorderStyle.NONE },
                                insideVertical: { style: BorderStyle.NONE }
                            },
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            width: { size: 50, type: WidthType.PERCENTAGE },
                                            children: [new Paragraph({ text: "" })]
                                        }),
                                        new TableCell({
                                            width: { size: 50, type: WidthType.PERCENTAGE },
                                            children: [
                                                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "XÁC NHẬN CỦA LÃNH ĐẠO KHOA, TRƯỞNG KHOA", bold: true, font: "Times New Roman", size: 22 })] }),
                                                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký và ghi rõ họ tên)", italic: true, font: "Times New Roman", size: 20 })] })
                                            ]
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }]
            });

            const blob = await Packer.toBlob(doc);
            const fileName = `Phieu_Danh_Gia_KPI_${kpiTargetUser}_${selectedKpiMonth}.docx`;
            saveAs(blob, fileName);
        }
    </script>
</body>
</html>
