// 映射配置頁面JavaScript
let customers = [];
let mappingData = {};

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
            customers = result.customers;
            mappingData = result.existing_mapping || {};
            renderMappingTable();
            hideLoading();
            
            // 顯示數據來源
            if (result.source === 'mapping_table') {
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

// 渲染映射表格
function renderMappingTable() {
    const tbody = document.getElementById('mapping-tbody');
    tbody.innerHTML = '';

    customers.forEach((customer, index) => {
        const row = document.createElement('tr');
        row.setAttribute('data-row-index', index);

        // 獲取現有數據
        const regionValue = mappingData.regions && mappingData.regions[customer] ? mappingData.regions[customer] : '';
        const scheduleValue = mappingData.schedule_breakpoints && mappingData.schedule_breakpoints[customer] ? mappingData.schedule_breakpoints[customer] : '';
        const etdValue = mappingData.etd && mappingData.etd[customer] ? mappingData.etd[customer] : '';
        const etaValue = mappingData.eta && mappingData.eta[customer] ? mappingData.eta[customer] : '';

        // 解析 ETD 和 ETA 的週別和星期
        const etdParsed = parseWeekDay(etdValue);
        const etaParsed = parseWeekDay(etaValue);

        row.innerHTML = `
            <td><input type="text" data-row="${index}" data-field="customer" placeholder="輸入客戶簡稱" value="${customer}" class="customer-name-input"></td>
            <td><input type="text" data-row="${index}" data-field="region" placeholder="輸入地區代碼" value="${regionValue}"></td>
            <td>
                <select data-row="${index}" data-field="schedule" class="mapping-select">
                    ${generateSelectOptions(SCHEDULE_OPTIONS, scheduleValue)}
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
            <td class="action-cell">
                <button class="btn btn-danger btn-sm delete-btn" onclick="deleteCustomer(${index})" title="刪除此客戶">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });

    // 顯示表格
    document.getElementById('mapping-table-container').style.display = 'block';
}

// 新增客戶
function addNewCustomer() {
    const newCustomerName = `新客戶_${Date.now()}`;
    customers.push(newCustomerName);

    // 初始化新客戶的映射數據
    if (!mappingData.regions) mappingData.regions = {};
    if (!mappingData.schedule_breakpoints) mappingData.schedule_breakpoints = {};
    if (!mappingData.etd) mappingData.etd = {};
    if (!mappingData.eta) mappingData.eta = {};

    mappingData.regions[newCustomerName] = '';
    mappingData.schedule_breakpoints[newCustomerName] = '';
    mappingData.etd[newCustomerName] = '';
    mappingData.eta[newCustomerName] = '';

    renderMappingTable();

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

    showNotification('已新增客戶，請填寫客戶簡稱和相關資訊', 'info');
}

// 待刪除的客戶索引（用於確認刪除）
let pendingDeleteIndex = null;

// 顯示刪除確認彈跳視窗
function showDeleteConfirmModal(index) {
    const customerName = customers[index];
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
                    <button class="btn btn-danger" onclick="confirmDeleteCustomer()">
                        <i class="fas fa-trash"></i> 確認刪除
                    </button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }

    // 更新訊息
    document.getElementById('delete-confirm-message').textContent =
        `確定要刪除客戶「${customerName}」嗎？`;

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

// 確認刪除客戶
function confirmDeleteCustomer() {
    if (pendingDeleteIndex === null) return;

    const index = pendingDeleteIndex;
    const customerName = customers[index];

    // 從陣列中移除
    customers.splice(index, 1);

    // 從映射數據中移除
    if (mappingData.regions && mappingData.regions[customerName]) {
        delete mappingData.regions[customerName];
    }
    if (mappingData.schedule_breakpoints && mappingData.schedule_breakpoints[customerName]) {
        delete mappingData.schedule_breakpoints[customerName];
    }
    if (mappingData.etd && mappingData.etd[customerName]) {
        delete mappingData.etd[customerName];
    }
    if (mappingData.eta && mappingData.eta[customerName]) {
        delete mappingData.eta[customerName];
    }

    // 隱藏 Modal
    hideDeleteConfirmModal();

    // 重新渲染表格
    renderMappingTable();
    showNotification(`已刪除客戶「${customerName}」`, 'success');
}

// 刪除客戶（顯示確認視窗）
function deleteCustomer(index) {
    showDeleteConfirmModal(index);
}

// 保存映射配置
async function saveMapping() {
    try {
        // 收集所有輸入數據 - 使用新的 data-row 結構
        const rows = document.querySelectorAll('#mapping-tbody tr');
        const mapping = {
            regions: {},
            schedule_breakpoints: {},
            etd: {},
            eta: {}
        };

        // 用於檢查重複的客戶名稱
        const customerNames = new Set();
        let hasDuplicate = false;
        let hasEmptyCustomer = false;

        rows.forEach(row => {
            const customerInput = row.querySelector('input[data-field="customer"]');
            const regionInput = row.querySelector('input[data-field="region"]');
            const scheduleSelect = row.querySelector('select[data-field="schedule"]');
            // ETD 和 ETA 現在是兩個下拉選單的組合
            const etdWeekSelect = row.querySelector('select[data-field="etd-week"]');
            const etdDaySelect = row.querySelector('select[data-field="etd-day"]');
            const etaWeekSelect = row.querySelector('select[data-field="eta-week"]');
            const etaDaySelect = row.querySelector('select[data-field="eta-day"]');

            const customerName = customerInput ? customerInput.value.trim() : '';

            // 檢查空的客戶名稱
            if (!customerName) {
                hasEmptyCustomer = true;
                return;
            }

            // 檢查重複
            if (customerNames.has(customerName)) {
                hasDuplicate = true;
                return;
            }
            customerNames.add(customerName);

            // 收集數據
            const regionValue = regionInput ? regionInput.value.trim() : '';
            const scheduleValue = scheduleSelect ? scheduleSelect.value : '';
            // 組合 ETD 和 ETA 的週別和星期
            const etdWeek = etdWeekSelect ? etdWeekSelect.value : '';
            const etdDay = etdDaySelect ? etdDaySelect.value : '';
            const etaWeek = etaWeekSelect ? etaWeekSelect.value : '';
            const etaDay = etaDaySelect ? etaDaySelect.value : '';
            const etdValue = combineWeekDay(etdWeek, etdDay);
            const etaValue = combineWeekDay(etaWeek, etaDay);

            if (regionValue) mapping.regions[customerName] = regionValue;
            if (scheduleValue) mapping.schedule_breakpoints[customerName] = scheduleValue;
            if (etdValue) mapping.etd[customerName] = etdValue;
            if (etaValue) mapping.eta[customerName] = etaValue;
        });

        // 驗證
        if (hasEmptyCustomer) {
            showNotification('客戶簡稱不能為空，請檢查後再保存', 'error');
            return;
        }

        if (hasDuplicate) {
            showNotification('存在重複的客戶簡稱，請檢查後再保存', 'error');
            return;
        }

        // 發送到服務器
        const response = await fetch('/save_mapping', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(mapping)
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
