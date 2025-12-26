// 映射配置頁面JavaScript
let mappingList = [];  // 新格式：列表形式，每筆為一個 mapping 記錄
let customers = [];    // 舊格式相容
let mappingData = {};  // 舊格式相容

// 分頁設定
let currentPage = 1;
const PAGE_SIZE = 10;  // 每頁顯示 10 筆

// 下拉選單選項定義
const SCHEDULE_OPTIONS = [
    { value: '', label: '請選擇' },
    { value: '禮拜一', label: '禮拜一' },
    { value: '禮拜二', label: '禮拜二' },
    { value: '禮拜三', label: '禮拜三' },
    { value: '禮拜四', label: '禮拜四' },
    { value: '禮拜五', label: '禮拜五' },
    { value: '禮拜六', label: '禮拜六' },
    { value: '禮拜日', label: '禮拜日' }
];

// ETD/ETA 週別選項（第一個下拉選單）
const WEEK_OPTIONS = [
    { value: '', label: '週別' },
    { value: '本週', label: '本週' },
    { value: '下週', label: '下週' },
    { value: '下下週', label: '下下週' },
    { value: '上週', label: '上週' }
];

// ETD/ETA 星期選項（第二個下拉選單）
const DAY_OPTIONS = [
    { value: '', label: '星期' },
    { value: '一', label: '一' },
    { value: '二', label: '二' },
    { value: '三', label: '三' },
    { value: '四', label: '四' },
    { value: '五', label: '五' },
    { value: '六', label: '六' },
    { value: '日', label: '日' }
];

// 解析 ETD/ETA 值為週別和星期
function parseWeekDay(value) {
    if (!value) return { week: '', day: '' };

    // 匹配格式如：本週一、下週二、下下週三、上週四
    const weekPatterns = ['下下週', '下週', '本週', '上週'];
    for (const week of weekPatterns) {
        if (value.startsWith(week)) {
            const day = value.substring(week.length);
            return { week, day };
        }
    }
    return { week: '', day: '' };
}

// 組合週別和星期為完整值
function combineWeekDay(week, day) {
    if (!week || !day) return '';
    return week + day;
}

// 生成 select 的 HTML
function generateSelectOptions(options, selectedValue) {
    return options.map(opt =>
        `<option value="${opt.value}" ${opt.value === selectedValue ? 'selected' : ''}>${opt.label}</option>`
    ).join('');
}

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    loadMappingData();
    initializeEventListeners();
});

// 事件監聽器
function initializeEventListeners() {
    document.getElementById('save-mapping-btn').addEventListener('click', saveMapping);
    document.getElementById('add-customer-btn').addEventListener('click', addNewCustomer);
}

// 載入映射數據
async function loadMappingData() {
    try {
        showLoading();

        const response = await fetch('/get_mapping_data');
        const result = await response.json();

        if (result.success) {
            // 檢查是新格式（list）還是舊格式
            if (result.format === 'list' && result.mapping_list) {
                // 新格式：直接使用列表
                mappingList = result.mapping_list;
                renderMappingTableList();
            } else {
                // 舊格式：轉換為列表格式
                customers = result.customers || [];
                mappingData = result.existing_mapping || {};
                // 轉換舊格式為新格式
                mappingList = convertOldFormatToList(customers, mappingData);
                renderMappingTableList();
            }
            hideLoading();

            // 顯示數據來源
            if (result.source === 'database') {
                showNotification(`已載入 ${mappingList.length} 筆映射資料`, 'success');
            } else if (result.source === 'mapping_table') {
                showNotification('已載入現有mapping表數據', 'success');
            } else {
                showNotification('已載入ERP文件客戶數據，請配置映射關係', 'info');
            }
        } else {
            showError(result.message);
        }
    } catch (error) {
        showError('載入數據失敗: ' + error.message);
    }
}

// 將舊格式轉換為列表格式
function convertOldFormatToList(customers, mappingData) {
    const list = [];
    customers.forEach(customer => {
        list.push({
            customer_name: customer,
            region: mappingData.regions ? (mappingData.regions[customer] || '') : '',
            schedule_breakpoint: mappingData.schedule_breakpoints ? (mappingData.schedule_breakpoints[customer] || '') : '',
            etd: mappingData.etd ? (mappingData.etd[customer] || '') : '',
            eta: mappingData.eta ? (mappingData.eta[customer] || '') : '',
            requires_transit: true  // 舊格式預設為需要在途
        });
    });
    return list;
}

// 計算總頁數
function getTotalPages() {
    return Math.ceil(mappingList.length / PAGE_SIZE);
}

// 取得當前頁的資料
function getCurrentPageData() {
    const startIndex = (currentPage - 1) * PAGE_SIZE;
    const endIndex = startIndex + PAGE_SIZE;
    return mappingList.slice(startIndex, endIndex).map((item, i) => ({
        ...item,
        originalIndex: startIndex + i  // 保存原始索引
    }));
}

// 渲染映射表格（新版列表格式，支援分頁）
function renderMappingTableList() {
    const tbody = document.getElementById('mapping-tbody');
    tbody.innerHTML = '';

    // 取得當前頁的資料
    const pageData = getCurrentPageData();

    pageData.forEach((item) => {
        const index = item.originalIndex;  // 使用原始索引
        const row = document.createElement('tr');
        row.setAttribute('data-row-index', index);

        // 解析 ETD 和 ETA 的週別和星期
        const etdParsed = parseWeekDay(item.etd || '');
        const etaParsed = parseWeekDay(item.eta || '');

        row.innerHTML = `
            <td><input type="text" data-row="${index}" data-field="customer" placeholder="輸入客戶簡稱" value="${item.customer_name || ''}" class="customer-name-input"></td>
            <td><input type="text" data-row="${index}" data-field="region" placeholder="輸入地區代碼" value="${item.region || ''}"></td>
            <td>
                <select data-row="${index}" data-field="schedule" class="mapping-select">
                    ${generateSelectOptions(SCHEDULE_OPTIONS, item.schedule_breakpoint || '')}
                </select>
            </td>
            <td>
                <div class="week-day-select">
                    <select data-row="${index}" data-field="etd-week" class="mapping-select week-select">
                        ${generateSelectOptions(WEEK_OPTIONS, etdParsed.week)}
                    </select>
                    <select data-row="${index}" data-field="etd-day" class="mapping-select day-select">
                        ${generateSelectOptions(DAY_OPTIONS, etdParsed.day)}
                    </select>
                </div>
            </td>
            <td>
                <div class="week-day-select">
                    <select data-row="${index}" data-field="eta-week" class="mapping-select week-select">
                        ${generateSelectOptions(WEEK_OPTIONS, etaParsed.week)}
                    </select>
                    <select data-row="${index}" data-field="eta-day" class="mapping-select day-select">
                        ${generateSelectOptions(DAY_OPTIONS, etaParsed.day)}
                    </select>
                </div>
            </td>
            <td class="toggle-cell">
                <label class="toggle-switch-mini">
                    <input type="checkbox" data-row="${index}" data-field="requires-transit" ${item.requires_transit !== false ? 'checked' : ''}>
                    <span class="toggle-slider-mini"></span>
                </label>
            </td>
            <td class="action-cell">
                <button class="btn btn-danger btn-sm delete-btn" onclick="deleteCustomerFromList(${index})" title="刪除此記錄">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });

    // 顯示表格
    document.getElementById('mapping-table-container').style.display = 'block';

    // 渲染分頁控件
    renderPagination();
}

// 渲染分頁控件
function renderPagination() {
    const totalPages = getTotalPages();
    let paginationContainer = document.getElementById('pagination-container');

    // 如果容器不存在，創建它
    if (!paginationContainer) {
        paginationContainer = document.createElement('div');
        paginationContainer.id = 'pagination-container';
        paginationContainer.className = 'pagination-container';
        const tableWrapper = document.querySelector('.table-wrapper');
        if (tableWrapper) {
            tableWrapper.after(paginationContainer);
        }
    }

    // 如果只有一頁或沒有資料，隱藏分頁
    if (totalPages <= 1) {
        paginationContainer.style.display = 'none';
        return;
    }

    paginationContainer.style.display = 'flex';

    // 生成分頁 HTML
    let paginationHTML = `
        <div class="pagination-info">
            共 ${mappingList.length} 筆記錄，第 ${currentPage} / ${totalPages} 頁
        </div>
        <div class="pagination-buttons">
            <button class="pagination-btn" onclick="goToPage(1)" ${currentPage === 1 ? 'disabled' : ''}>
                <i class="fas fa-angle-double-left"></i>
            </button>
            <button class="pagination-btn" onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>
                <i class="fas fa-angle-left"></i>
            </button>
    `;

    // 頁碼按鈕（最多顯示 5 個）
    const maxVisiblePages = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
        paginationHTML += `
            <button class="pagination-btn pagination-num ${i === currentPage ? 'active' : ''}" onclick="goToPage(${i})">
                ${i}
            </button>
        `;
    }

    paginationHTML += `
            <button class="pagination-btn" onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>
                <i class="fas fa-angle-right"></i>
            </button>
            <button class="pagination-btn" onclick="goToPage(${totalPages})" ${currentPage === totalPages ? 'disabled' : ''}>
                <i class="fas fa-angle-double-right"></i>
            </button>
        </div>
    `;

    paginationContainer.innerHTML = paginationHTML;
}

// 跳轉到指定頁
function goToPage(page) {
    const totalPages = getTotalPages();
    if (page < 1 || page > totalPages) return;

    // 保存當前頁面的編輯
    saveCurrentPageEdits();

    currentPage = page;
    renderMappingTableList();
}

// 保存當前頁面的編輯到 mappingList
function saveCurrentPageEdits() {
    const rows = document.querySelectorAll('#mapping-tbody tr');
    rows.forEach(row => {
        const index = parseInt(row.getAttribute('data-row-index'));
        if (index >= 0 && index < mappingList.length) {
            const customerInput = row.querySelector('input[data-field="customer"]');
            const regionInput = row.querySelector('input[data-field="region"]');
            const scheduleSelect = row.querySelector('select[data-field="schedule"]');
            const etdWeekSelect = row.querySelector('select[data-field="etd-week"]');
            const etdDaySelect = row.querySelector('select[data-field="etd-day"]');
            const etaWeekSelect = row.querySelector('select[data-field="eta-week"]');
            const etaDaySelect = row.querySelector('select[data-field="eta-day"]');
            const requiresTransitCheckbox = row.querySelector('input[data-field="requires-transit"]');

            mappingList[index].customer_name = customerInput ? customerInput.value.trim() : '';
            mappingList[index].region = regionInput ? regionInput.value.trim() : '';
            mappingList[index].schedule_breakpoint = scheduleSelect ? scheduleSelect.value : '';

            const etdWeek = etdWeekSelect ? etdWeekSelect.value : '';
            const etdDay = etdDaySelect ? etdDaySelect.value : '';
            const etaWeek = etaWeekSelect ? etaWeekSelect.value : '';
            const etaDay = etaDaySelect ? etaDaySelect.value : '';

            mappingList[index].etd = combineWeekDay(etdWeek, etdDay);
            mappingList[index].eta = combineWeekDay(etaWeek, etaDay);
            mappingList[index].requires_transit = requiresTransitCheckbox ? requiresTransitCheckbox.checked : true;
        }
    });
}

// 渲染映射表格（舊版相容）
function renderMappingTable() {
    // 轉換為列表格式後使用新版渲染
    mappingList = convertOldFormatToList(customers, mappingData);
    renderMappingTableList();
}

// 新增客戶（新版列表格式）
function addNewCustomer() {
    // 保存當前頁面的編輯
    saveCurrentPageEdits();

    // 新增一筆空的 mapping 記錄
    mappingList.push({
        customer_name: '',
        region: '',
        schedule_breakpoint: '',
        etd: '',
        eta: '',
        requires_transit: true  // 預設需要在途文件
    });

    // 跳轉到最後一頁
    currentPage = getTotalPages();
    renderMappingTableList();

    // 滾動到新增的行並聚焦
    const tbody = document.getElementById('mapping-tbody');
    const lastRow = tbody.lastElementChild;
    if (lastRow) {
        lastRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
        const customerInput = lastRow.querySelector('.customer-name-input');
        if (customerInput) {
            customerInput.focus();
            customerInput.select();
        }
    }

    showNotification('已新增一筆映射記錄，請填寫客戶簡稱和相關資訊', 'info');
}

// 待刪除的客戶索引（用於確認刪除）
let pendingDeleteIndex = null;

// 顯示刪除確認彈跳視窗（新版列表格式）
function showDeleteConfirmModalList(index) {
    const item = mappingList[index];
    const displayName = item.customer_name ? `${item.customer_name} - ${item.region}` : `記錄 #${index + 1}`;
    pendingDeleteIndex = index;

    // 建立 Modal HTML（如果不存在）
    let modal = document.getElementById('delete-confirm-modal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'delete-confirm-modal';
        modal.className = 'custom-modal-overlay';
        modal.innerHTML = `
            <div class="custom-modal">
                <div class="custom-modal-header">
                    <i class="fas fa-exclamation-triangle warning-icon"></i>
                    <h3>確認刪除</h3>
                </div>
                <div class="custom-modal-body">
                    <p id="delete-confirm-message"></p>
                    <p class="warning-text">此操作無法復原！</p>
                </div>
                <div class="custom-modal-footer">
                    <button class="btn btn-secondary" onclick="hideDeleteConfirmModal()">
                        <i class="fas fa-times"></i> 取消
                    </button>
                    <button class="btn btn-danger" onclick="confirmDeleteFromList()">
                        <i class="fas fa-trash"></i> 確認刪除
                    </button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }

    // 更新訊息
    document.getElementById('delete-confirm-message').textContent =
        `確定要刪除「${displayName}」嗎？`;

    // 顯示 Modal
    modal.style.display = 'flex';
    setTimeout(() => {
        modal.classList.add('show');
    }, 10);
}

// 隱藏刪除確認彈跳視窗
function hideDeleteConfirmModal() {
    const modal = document.getElementById('delete-confirm-modal');
    if (modal) {
        modal.classList.remove('show');
        setTimeout(() => {
            modal.style.display = 'none';
        }, 300);
    }
    pendingDeleteIndex = null;
}

// 確認刪除（新版列表格式）
function confirmDeleteFromList() {
    if (pendingDeleteIndex === null) return;

    const index = pendingDeleteIndex;
    const item = mappingList[index];
    const displayName = item.customer_name ? `${item.customer_name} - ${item.region}` : `記錄 #${index + 1}`;

    // 從列表中移除
    mappingList.splice(index, 1);

    // 如果刪除後當前頁沒有資料，則跳轉到前一頁
    const totalPages = getTotalPages();
    if (currentPage > totalPages && totalPages > 0) {
        currentPage = totalPages;
    }

    // 隱藏 Modal
    hideDeleteConfirmModal();

    // 重新渲染表格
    renderMappingTableList();
    showNotification(`已刪除「${displayName}」`, 'success');
}

// 刪除客戶（從列表，顯示確認視窗）
function deleteCustomerFromList(index) {
    showDeleteConfirmModalList(index);
}

// 舊版刪除函數（相容）
function deleteCustomer(index) {
    deleteCustomerFromList(index);
}

// 舊版確認刪除（相容）
function confirmDeleteCustomer() {
    confirmDeleteFromList();
}

// 保存映射配置（新版列表格式）
async function saveMapping() {
    try {
        // 先保存當前頁面的編輯到 mappingList
        saveCurrentPageEdits();

        // 使用完整的 mappingList 來驗證和保存
        // 用於檢查重複的 (customer_name, region) 組合
        const uniqueKeys = new Set();
        let hasDuplicate = false;
        let hasEmptyCustomer = false;
        let hasEmptyRegion = false;

        // 遍歷完整的 mappingList 進行驗證
        mappingList.forEach((item, index) => {
            const customerName = (item.customer_name || '').trim();
            const regionValue = (item.region || '').trim();

            // 檢查空的客戶名稱
            if (!customerName) {
                hasEmptyCustomer = true;
                return;
            }

            // 檢查空的地區代碼
            if (!regionValue) {
                hasEmptyRegion = true;
                return;
            }

            // 檢查重複的 (customer_name, region) 組合
            const uniqueKey = `${customerName}|${regionValue}`;
            if (uniqueKeys.has(uniqueKey)) {
                hasDuplicate = true;
                return;
            }
            uniqueKeys.add(uniqueKey);
        });

        // 準備要保存的資料
        const mappingListToSave = mappingList.filter(item => {
            const customerName = (item.customer_name || '').trim();
            const regionValue = (item.region || '').trim();
            return customerName && regionValue;
        }).map(item => ({
            customer_name: (item.customer_name || '').trim(),
            region: (item.region || '').trim(),
            schedule_breakpoint: item.schedule_breakpoint || '',
            etd: item.etd || '',
            eta: item.eta || '',
            requires_transit: item.requires_transit !== false  // 預設為 true
        }));

        // 驗證
        if (hasEmptyCustomer) {
            showNotification('客戶簡稱不能為空，請檢查後再保存', 'error');
            return;
        }

        if (hasEmptyRegion) {
            showNotification('客戶需求地區不能為空，請檢查後再保存', 'error');
            return;
        }

        if (hasDuplicate) {
            showNotification('存在重複的（客戶簡稱 + 地區）組合，請檢查後再保存', 'error');
            return;
        }

        // 發送到服務器（使用新的列表格式 API）
        const response = await fetch('/save_mapping_list', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ mapping_list: mappingListToSave })
        });

        const result = await response.json();

        if (result.success) {
            showNotification('映射配置保存成功！', 'success');
            // 延遲關閉窗口，讓用戶看到成功消息
            setTimeout(() => {
                window.close();
            }, 1500);
        } else {
            showNotification('保存失敗: ' + result.message, 'error');
        }
    } catch (error) {
        showNotification('保存失敗: ' + error.message, 'error');
    }
}

// 顯示載入狀態
function showLoading() {
    document.getElementById('loading-state').style.display = 'block';
    document.getElementById('mapping-table-container').style.display = 'none';
    document.getElementById('error-state').style.display = 'none';
}

// 隱藏載入狀態
function hideLoading() {
    document.getElementById('loading-state').style.display = 'none';
}

// 顯示錯誤
function showError(message) {
    document.getElementById('loading-state').style.display = 'none';
    document.getElementById('mapping-table-container').style.display = 'none';
    document.getElementById('error-state').style.display = 'block';
    document.getElementById('error-message').textContent = message;
}

// 顯示通知
function showNotification(message, type = 'info') {
    // 創建通知元素
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <div class="notification-content">
            <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}"></i>
            <span>${message}</span>
        </div>
    `;
    
    // 添加樣式
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${type === 'success' ? '#d4edda' : type === 'error' ? '#f8d7da' : '#d1ecf1'};
        color: ${type === 'success' ? '#155724' : type === 'error' ? '#721c24' : '#0c5460'};
        border: 1px solid ${type === 'success' ? '#c3e6cb' : type === 'error' ? '#f5c6cb' : '#bee5eb'};
        border-radius: 8px;
        padding: 15px 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1000;
        animation: slideInRight 0.3s ease-out;
    `;
    
    // 添加到頁面
    document.body.appendChild(notification);
    
    // 3秒後自動移除
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease-in';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    }, 3000);
}

// 添加動畫樣式
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    @keyframes slideOutRight {
        from {
            transform: translateX(0);
            opacity: 1;
        }
        to {
            transform: translateX(100%);
            opacity: 0;
        }
    }
    
    .notification-content {
        display: flex;
        align-items: center;
        gap: 10px;
    }
`;
document.head.appendChild(style);
