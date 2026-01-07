// 全局變量
let currentStep = 1;
let uploadedFiles = {
    erp: false,
    forecast: false,
    transit: false
};

// 在途文件是否為必填項目
let transitRequired = true;

// ========================================
// IT 測試模式支援
// ========================================

// 檢查是否為 IT 測試模式
function isITTestMode() {
    return window.IT_TEST_MODE && window.IT_TEST_MODE.enabled && window.IT_TEST_MODE.selectedCustomer;
}

// 獲取測試模式的客戶 ID
function getTestCustomerId() {
    if (isITTestMode()) {
        return window.IT_TEST_MODE.selectedCustomer.id;
    }
    return null;
}

// 獲取測試模式的客戶資訊
function getTestCustomerInfo() {
    if (isITTestMode()) {
        return window.IT_TEST_MODE.selectedCustomer;
    }
    return null;
}

// 為 FormData 添加測試模式參數
function appendTestModeParams(formData) {
    if (isITTestMode()) {
        formData.append('test_mode', 'true');
        formData.append('customer_id', getTestCustomerId());
        console.log('🧪 測試模式：添加 test_mode=true, customer_id=' + getTestCustomerId());
    }
}

// 為 JSON 請求添加測試模式參數
function getTestModeJsonParams() {
    if (isITTestMode()) {
        const params = {
            test_mode: true,
            customer_id: getTestCustomerId()
        };
        console.log('🧪 測試模式：添加 JSON 參數', params);
        return params;
    }
    return {};
}

// 重置所有狀態（供 IT 測試模式使用）
function resetAllStates() {
    // 重置上傳 Session ID
    resetUploadSessionId();

    // 重置上傳狀態
    uploadedFiles = {
        erp: false,
        forecast: false,
        transit: false
    };

    // 重置上傳檔案資訊
    uploadedFileInfo = {
        erp: null,
        forecast: null,
        transit: null
    };

    // 重置在途需求檢查結果
    forecastTransitCheck = null;

    // 重置在途必填狀態
    transitRequired = true;
    const transitCheckbox = document.getElementById('transit-required-checkbox');
    if (transitCheckbox) {
        transitCheckbox.checked = true;
    }

    // 重置UI - 上傳區域
    ['erp', 'forecast', 'transit'].forEach(type => {
        const uploadBox = document.getElementById(`${type}-upload-box`);
        const status = document.getElementById(`${type}-status`);
        const fileInput = document.getElementById(`${type}-file`);

        if (uploadBox) uploadBox.style.display = 'block';
        if (status) status.style.display = 'none';
        if (fileInput) fileInput.value = '';
    });

    // 重置 checklist 狀態
    if (typeof updateUploadChecklist === 'function') {
        updateUploadChecklist();
    }

    // 重置下一步按鈕狀態
    const nextBtn = document.getElementById('upload-next-btn');
    if (nextBtn) {
        nextBtn.disabled = true;
    }

    // 隱藏所有處理結果
    const resultIds = ['cleanup-result', 'mapping-result', 'forecast-result'];
    resultIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.style.display = 'none';
            element.innerHTML = '';
        }
    });

    // 重置步驟
    currentStep = 1;
    if (typeof updateProgress === 'function') {
        updateProgress();
    }
    if (typeof showSection === 'function') {
        showSection('upload');
    }

    console.log('🔄 所有狀態已重置');
}

// 儲存上傳檔案的詳細資訊
let uploadedFileInfo = {
    erp: null,
    forecast: null,
    transit: null
};

// 儲存 Forecast 上傳時的在途需求檢查結果
let forecastTransitCheck = null;

// 上傳 Session ID - 前端產生，確保同一批上傳的檔案都在同一個資料夾
let uploadSessionId = null;

// 產生新的上傳 Session ID（格式：YYYYMMDD_HHMMSS）
function generateUploadSessionId() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    return `${year}${month}${day}_${hours}${minutes}${seconds}`;
}

// 獲取當前上傳 Session ID（如果不存在則產生新的）
function getUploadSessionId() {
    if (!uploadSessionId) {
        uploadSessionId = generateUploadSessionId();
        console.log('📁 產生新的上傳 Session ID:', uploadSessionId);
    }
    return uploadSessionId;
}

// 重置上傳 Session ID（開始新的上傳流程時使用）
function resetUploadSessionId() {
    uploadSessionId = generateUploadSessionId();
    console.log('🔄 重置上傳 Session ID:', uploadSessionId);
    return uploadSessionId;
}

// DOM元素
const progressSteps = document.querySelectorAll('.step');
const sections = {
    upload: document.getElementById('upload-section'),
    cleanup: document.getElementById('cleanup-section'),
    mapping: document.getElementById('mapping-section'),
    forecast: document.getElementById('forecast-section'),
    download: document.getElementById('download-section')
};

// 初始化
document.addEventListener('DOMContentLoaded', async function() {
    // 頁面載入時產生新的上傳 Session ID
    resetUploadSessionId();

    initializeEventListeners();
    updateProgress();
});

// 事件監聽器
function initializeEventListeners() {
    // 文件上傳
    document.getElementById('erp-file').addEventListener('change', handleErpUpload);
    document.getElementById('forecast-file').addEventListener('change', handleForecastUpload);
    document.getElementById('transit-file').addEventListener('change', handleTransitUpload);

    // 在途文件必填選項變更
    const transitCheckbox = document.getElementById('transit-required-checkbox');
    if (transitCheckbox) {
        transitCheckbox.addEventListener('change', handleTransitRequiredChange);
    }

    // 處理按鈕
    document.getElementById('cleanup-btn').addEventListener('click', handleCleanup);
    document.getElementById('mapping-config-btn').addEventListener('click', openMappingConfig);
    document.getElementById('mapping-process-btn').addEventListener('click', handleMappingProcess);
    document.getElementById('forecast-btn').addEventListener('click', handleForecast);

    // 文件上傳區域的映射配置按鈕
    const uploadMappingBtn = document.getElementById('upload-mapping-config-btn');
    if (uploadMappingBtn) {
        uploadMappingBtn.addEventListener('click', openMappingConfig);
    }
}

// ERP文件上傳處理
async function handleErpUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // 檢查文件類型
    const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!allowedTypes.includes(file.type)) {
        showError('erp', '請選擇Excel文件 (.xlsx 或 .xls)');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);
    // 傳遞前端產生的 upload session id，確保檔案存到同一個資料夾
    formData.append('upload_session_id', getUploadSessionId());
    // 添加測試模式參數（如果是 IT 測試模式）
    appendTestModeParams(formData);

    try {
        showLoading('erp');
        console.log('開始上傳ERP文件:', file.name, '大小:', file.size, 'bytes', 'sessionId:', getUploadSessionId(), 'testMode:', isITTestMode());

        const response = await fetch('/upload_erp', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        console.log('ERP上傳結果:', result);

        if (result.success) {
            showUploadSuccess('erp', result);
            uploadedFiles.erp = true;
            // 儲存檔案資訊
            uploadedFileInfo.erp = {
                filename: file.name,
                rows: result.rows,
                columns: result.columns ? result.columns.length : 0,
                fileSize: result.file_size
            };
            checkUploadComplete();
        } else {
            // 檢查是否為格式驗證錯誤
            if (result.validation_error) {
                showValidationError('erp', result.message, result.details);
            } else {
                showError('erp', result.message);
            }
        }
    } catch (error) {
        console.error('ERP上傳錯誤:', error);
        showError('erp', '上傳失敗: ' + error.message);
    }
}

// Forecast文件上傳處理（支援多檔案）
async function handleForecastUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    // 檢查所有文件類型
    const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    for (let i = 0; i < files.length; i++) {
        // 有些瀏覽器可能不會正確識別 xlsx 類型，所以也檢查副檔名
        const fileName = files[i].name.toLowerCase();
        const isValidType = allowedTypes.includes(files[i].type) || fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
        if (!isValidType) {
            showError('forecast', `文件 "${files[i].name}" 不是有效的 Excel 文件 (.xlsx 或 .xls)`);
            return;
        }
    }

    // 取得合併選項
    const mergeCheckbox = document.getElementById('merge-forecast-files');
    const shouldMerge = mergeCheckbox ? mergeCheckbox.checked : true;

    // 顯示多檔案進度
    showForecastMultiUploadProgress(files);

    const formData = new FormData();
    // 添加所有文件到 FormData
    for (let i = 0; i < files.length; i++) {
        formData.append('files', files[i]);
        console.log(`準備上傳Forecast文件 ${i + 1}:`, files[i].name, '大小:', files[i].size, 'bytes');
    }
    // 添加合併選項
    formData.append('merge_files', shouldMerge ? 'true' : 'false');
    // 傳遞前端產生的 upload session id，確保檔案存到同一個資料夾
    formData.append('upload_session_id', getUploadSessionId());
    // 添加測試模式參數（如果是 IT 測試模式）
    appendTestModeParams(formData);

    try {
        console.log(`開始上傳 ${files.length} 個 Forecast 文件，合併模式: ${shouldMerge}, sessionId: ${getUploadSessionId()}, testMode: ${isITTestMode()}`);

        // 模擬每個檔案的進度更新
        for (let i = 0; i < files.length; i++) {
            updateFileProgress(i, 'uploading', 50);
            await sleep(200);
        }

        const response = await fetch('/upload_forecast', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        console.log('Forecast上傳結果:', result);

        if (result.success) {
            // 更新所有進度為成功
            for (let i = 0; i < files.length; i++) {
                updateFileProgress(i, 'success', 100);
            }
            await sleep(500);

            showForecastMultiUploadSuccess(result);
            uploadedFiles.forecast = true;
            // 儲存檔案資訊
            uploadedFileInfo.forecast = {
                files: Array.from(files).map(f => f.name),
                fileCount: files.length,
                totalRows: result.total_rows || result.rows || 0,
                totalSize: result.total_size || result.file_size,
                details: result.files || null
            };

            // 儲存在途需求檢查結果，並根據結果自動調整 transitRequired
            if (result.transit_check) {
                forecastTransitCheck = result.transit_check;
                // 如果所有廠區都不需要在途文件，自動將 transitRequired 設為 false
                if (!result.transit_check.has_transit_requirement &&
                    result.transit_check.transit_not_required_regions &&
                    result.transit_check.transit_not_required_regions.length > 0) {
                    transitRequired = false;
                    // 更新前端 checkbox 狀態
                    const transitCheckbox = document.getElementById('transit-required-checkbox');
                    if (transitCheckbox) {
                        transitCheckbox.checked = false;
                        // 觸發 change 事件以更新 UI
                        transitCheckbox.dispatchEvent(new Event('change'));
                    }
                    console.log('📋 根據 Forecast 廠區檢查結果，自動設定在途文件為非必填');
                }
            }

            checkUploadComplete();
        } else {
            // 標記失敗的檔案
            for (let i = 0; i < files.length; i++) {
                updateFileProgress(i, 'error', 100);
            }
            await sleep(300);

            // 檢查是否為格式驗證錯誤或有詳細說明的錯誤
            if (result.validation_error) {
                showValidationError('forecast', result.message, result.details);
            } else if (result.details) {
                // 有詳細說明的錯誤（如工作表保護）
                showProtectionError('forecast', result.message, result.details);
            } else {
                showError('forecast', result.message);
            }
        }
    } catch (error) {
        console.error('Forecast上傳錯誤:', error);
        showError('forecast', '上傳失敗: ' + error.message);
    }
}

// 顯示多檔案上傳進度
function showForecastMultiUploadProgress(files) {
    const uploadBox = document.getElementById('forecast-upload-box');
    const status = document.getElementById('forecast-status');

    uploadBox.style.display = 'none';
    status.style.display = 'block';

    let progressHtml = `
        <div class="status-content" style="flex-direction: column; width: 100%;">
            <div class="status-text" style="width: 100%; text-align: center; margin-bottom: 15px;">
                <div class="status-title"><i class="fas fa-cloud-upload-alt"></i> 上傳中...</div>
                <div class="status-details">正在處理 ${files.length} 個檔案</div>
            </div>
            <div class="multi-upload-progress">
    `;

    for (let i = 0; i < files.length; i++) {
        const fileName = files[i].name.length > 25 ? files[i].name.substring(0, 22) + '...' : files[i].name;
        progressHtml += `
            <div class="file-progress-item" id="file-progress-${i}">
                <div class="file-progress-header">
                    <i class="fas fa-file-excel"></i>
                    <span class="file-progress-name" title="${files[i].name}">${fileName}</span>
                    <span class="file-progress-status uploading" id="file-status-${i}">等待中</span>
                </div>
                <div class="file-progress-bar">
                    <div class="file-progress-fill" id="file-progress-fill-${i}" style="width: 0%"></div>
                </div>
            </div>
        `;
    }

    progressHtml += '</div></div>';
    status.innerHTML = progressHtml;
}

// 更新單個檔案的進度
function updateFileProgress(index, status, percentage) {
    const fillEl = document.getElementById(`file-progress-fill-${index}`);
    const statusEl = document.getElementById(`file-status-${index}`);

    if (fillEl) {
        fillEl.style.width = percentage + '%';
        fillEl.className = 'file-progress-fill';
        if (status === 'success') fillEl.classList.add('success');
        if (status === 'error') fillEl.classList.add('error');
    }

    if (statusEl) {
        statusEl.className = 'file-progress-status ' + status;
        const statusText = {
            'uploading': '上傳中',
            'validating': '驗證中',
            'success': '完成',
            'error': '失敗'
        };
        statusEl.textContent = statusText[status] || status;
    }
}

// 顯示 Forecast 多檔案上傳成功
function showForecastMultiUploadSuccess(result) {
    const status = document.getElementById('forecast-status');
    const fileCount = result.file_count || 1;
    const totalRows = result.total_rows || result.rows || 0;
    const totalSize = result.total_size ? formatFileSize(result.total_size) : (result.file_size ? formatFileSize(result.file_size) : '');
    const isMergeMode = result.merge_mode !== false; // 預設為合併模式

    // 儲存檔案資訊供單檔刪除使用
    window.forecastFilesInfo = result.files || [];
    window.forecastMergeMode = isMergeMode;

    let detailsHtml = '';
    if (fileCount === 1) {
        // 單檔案顯示
        const fileSize = result.file_size ? ` (${formatFileSize(result.file_size)})` : '';
        detailsHtml = `${result.rows} 行數據，${result.columns ? result.columns.length : 0} 個欄位${fileSize}`;
    } else {
        // 多檔案顯示
        detailsHtml = `<strong>${fileCount} 個檔案</strong>，共 ${totalRows} 行數據`;
        if (totalSize) {
            detailsHtml += ` (${totalSize})`;
        }

        // 顯示每個檔案的詳細資訊
        if (result.files && result.files.length > 0) {
            detailsHtml += '<div class="multi-file-details" id="forecast-file-list">';
            result.files.forEach((file, index) => {
                const shortName = file.name.length > 25 ? file.name.substring(0, 22) + '...' : file.name;
                const savedName = file.saved_name || '';
                // 只有在非合併模式下且有 saved_name 時才顯示單檔刪除按鈕
                const deleteBtn = (!isMergeMode && savedName)
                    ? `<button class="btn-delete-single-file" onclick="deleteSingleForecastFile('${savedName}', ${index})" title="刪除此檔案"><i class="fas fa-times"></i></button>`
                    : '';
                detailsHtml += `
                    <div class="file-item" id="forecast-file-${index}">
                        <i class="fas fa-file-excel"></i>
                        <span class="file-name" title="${file.name}">${shortName}</span>
                        <span class="file-rows">${file.rows} 行</span>
                        ${deleteBtn}
                    </div>`;
            });
            detailsHtml += '</div>';
        }
    }

    // 檢查是否有在途提醒
    let transitAlertHtml = '';
    if (result.transit_check && result.transit_check.message) {
        const check = result.transit_check;
        let alertClass = 'info';
        let alertIcon = 'info-circle';

        if (check.has_transit_requirement) {
            alertClass = 'warning';
            alertIcon = 'exclamation-triangle';
        }

        transitAlertHtml = `
            <div class="transit-alert ${alertClass}">
                <i class="fas fa-${alertIcon}"></i>
                <span>${check.message}</span>
            </div>
        `;
    }

    status.innerHTML = `
        <div class="status-content" style="flex-direction: column; align-items: flex-start; width: 100%;">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 10px; width: 100%;">
                <i class="fas fa-check-circle status-icon success"></i>
                <div class="status-text" style="flex: 1;">
                    <div class="status-title">上傳成功</div>
                </div>
                <button class="btn-delete-upload" onclick="resetUpload('forecast')" title="刪除並重新上傳">
                    <i class="fas fa-trash-alt"></i>
                </button>
            </div>
            <div class="status-details" style="width: 100%;">${detailsHtml}</div>
            ${transitAlertHtml}
        </div>
    `;
}

// 在途文件上傳處理
async function handleTransitUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // 檢查文件類型
    const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!allowedTypes.includes(file.type)) {
        showError('transit', '請選擇Excel文件 (.xlsx 或 .xls)');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);
    // 傳遞前端產生的 upload session id，確保檔案存到同一個資料夾
    formData.append('upload_session_id', getUploadSessionId());
    // 添加測試模式參數（如果是 IT 測試模式）
    appendTestModeParams(formData);

    try {
        showLoading('transit');
        console.log('開始上傳在途文件:', file.name, '大小:', file.size, 'bytes', 'sessionId:', getUploadSessionId(), 'testMode:', isITTestMode());

        const response = await fetch('/upload_transit', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        console.log('在途文件上傳結果:', result);

        if (result.success) {
            showUploadSuccess('transit', result);
            uploadedFiles.transit = true;
            // 儲存檔案資訊
            uploadedFileInfo.transit = {
                filename: file.name,
                rows: result.rows,
                columns: result.columns ? result.columns.length : 0,
                fileSize: result.file_size
            };
            checkUploadComplete();
        } else {
            // 檢查是否為格式驗證錯誤
            if (result.validation_error) {
                showValidationError('transit', result.message, result.details);
            } else {
                showError('transit', result.message);
            }
        }
    } catch (error) {
        console.error('在途文件上傳錯誤:', error);
        showError('transit', '上傳失敗: ' + error.message);
    }
}

// 處理在途文件必填選項變更
function handleTransitRequiredChange(event) {
    transitRequired = event.target.checked;
    const uploadItem = document.querySelector('#transit-upload-box').closest('.upload-item');
    const uploadBox = document.getElementById('transit-upload-box');
    const fileInput = document.getElementById('transit-file');
    const selectBtn = uploadBox.querySelector('.btn');

    if (transitRequired) {
        // 必填模式
        uploadItem.classList.remove('transit-optional');
        // 啟用文件選擇
        if (selectBtn) {
            selectBtn.disabled = false;
        }
        if (fileInput) {
            fileInput.disabled = false;
        }
        console.log('在途文件設為必填');
    } else {
        // 選填模式 - 禁用文件選擇
        uploadItem.classList.add('transit-optional');
        if (selectBtn) {
            selectBtn.disabled = true;
        }
        if (fileInput) {
            fileInput.disabled = true;
        }
        console.log('在途文件設為選填（已禁用上傳）');
    }

    // 重新檢查上傳完成狀態
    checkUploadComplete();
}

// 顯示載入狀態
function showLoading(type) {
    const uploadBox = document.getElementById(`${type}-upload-box`);
    const status = document.getElementById(`${type}-status`);
    
    uploadBox.style.display = 'none';
    status.style.display = 'block';
    status.innerHTML = `
        <div class="status-content">
            <div class="spinner"></div>
            <div class="status-text">
                <div class="status-title">上傳中...</div>
                <div class="status-details">請稍候</div>
            </div>
        </div>
    `;
}

// 顯示上傳成功
function showUploadSuccess(type, result) {
    const status = document.getElementById(`${type}-status`);
    const fileSize = result.file_size ? ` (${formatFileSize(result.file_size)})` : '';
    status.innerHTML = `
        <div class="status-content" style="width: 100%;">
            <i class="fas fa-check-circle status-icon success"></i>
            <div class="status-text" style="flex: 1;">
                <div class="status-title">上傳成功</div>
                <div class="status-details">${result.rows} 行數據，${result.columns.length} 個欄位${fileSize}</div>
            </div>
            <button class="btn-delete-upload" onclick="resetUpload('${type}')" title="刪除並重新上傳">
                <i class="fas fa-trash-alt"></i>
            </button>
        </div>
    `;
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 重置上傳狀態（刪除已上傳的文件，允許重新上傳）
async function resetUpload(type) {
    const uploadBox = document.getElementById(`${type}-upload-box`);
    const status = document.getElementById(`${type}-status`);
    const fileInput = document.getElementById(`${type}-file`);

    // 呼叫後端 API 刪除伺服器上的檔案
    try {
        const response = await fetch('/api/delete_upload', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                type: type,
                upload_session_id: uploadSessionId
            })
        });

        const result = await response.json();
        if (!result.success) {
            console.error('刪除檔案失敗:', result.message);
        } else {
            console.log('伺服器檔案已刪除:', result.deleted_files);
        }
    } catch (error) {
        console.error('刪除檔案請求失敗:', error);
    }

    // 重置 UI
    uploadBox.style.display = 'block';
    status.style.display = 'none';

    // 清空文件輸入
    if (fileInput) {
        fileInput.value = '';
    }

    // 重置上傳狀態
    uploadedFiles[type] = false;
    uploadedFileInfo[type] = null;

    // 更新檢查清單和下一步按鈕
    checkUploadComplete();

    showNotification(`已刪除${type === 'erp' ? 'ERP淨需求' : type === 'forecast' ? 'Forecast' : '在途'}文件，請重新上傳`, 'info');
}

// 刪除單一 Forecast 檔案
async function deleteSingleForecastFile(savedName, index) {
    try {
        const response = await fetch('/api/delete_upload', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                type: 'forecast',
                upload_session_id: uploadSessionId,
                filename: savedName
            })
        });

        const result = await response.json();
        if (result.success) {
            console.log('單檔刪除成功:', savedName);

            // 從 UI 中移除該檔案項目
            const fileItem = document.getElementById(`forecast-file-${index}`);
            if (fileItem) {
                fileItem.remove();
            }

            // 更新 forecastFilesInfo
            if (window.forecastFilesInfo) {
                window.forecastFilesInfo = window.forecastFilesInfo.filter((_, i) => i !== index);
            }

            // 檢查是否還有剩餘檔案
            if (result.remaining_count === 0) {
                // 沒有剩餘檔案，重置整個 forecast 上傳狀態
                resetUpload('forecast');
                showNotification('所有 Forecast 檔案已刪除，請重新上傳', 'info');
            } else {
                // 更新檔案數量顯示
                updateForecastFileCount(result.remaining_count);

                // 更新在途提醒
                updateTransitAlert(result.transit_check);

                showNotification(`已刪除檔案，剩餘 ${result.remaining_count} 個檔案`, 'info');
            }
        } else {
            console.error('單檔刪除失敗:', result.message);
            showNotification('刪除檔案失敗: ' + result.message, 'error');
        }
    } catch (error) {
        console.error('單檔刪除請求失敗:', error);
        showNotification('刪除檔案失敗', 'error');
    }
}

// 更新 Forecast 檔案數量顯示
function updateForecastFileCount(count) {
    const status = document.getElementById('forecast-status');
    if (!status) return;

    // 更新標題中的檔案數量
    const detailsDiv = status.querySelector('.status-details');
    if (detailsDiv) {
        const strongEl = detailsDiv.querySelector('strong');
        if (strongEl) {
            strongEl.textContent = `${count} 個檔案`;
        }
    }
}

// 更新在途提醒
function updateTransitAlert(transitCheck) {
    const status = document.getElementById('forecast-status');
    if (!status) return;

    // 移除現有的在途提醒
    const existingAlert = status.querySelector('.transit-alert');
    if (existingAlert) {
        existingAlert.remove();
    }

    // 如果有新的在途提醒，加入
    if (transitCheck && transitCheck.message) {
        let alertClass = 'info';
        let alertIcon = 'info-circle';

        if (transitCheck.has_transit_requirement) {
            alertClass = 'warning';
            alertIcon = 'exclamation-triangle';
        }

        const alertHtml = `
            <div class="transit-alert ${alertClass}">
                <i class="fas fa-${alertIcon}"></i>
                <span>${transitCheck.message}</span>
            </div>
        `;

        // 找到 status-content 並在最後加入提醒
        const statusContent = status.querySelector('.status-content');
        if (statusContent) {
            statusContent.insertAdjacentHTML('beforeend', alertHtml);
        }
    }
}

// 顯示錯誤
function showError(type, message) {
    const uploadBox = document.getElementById(`${type}-upload-box`);
    const status = document.getElementById(`${type}-status`);

    uploadBox.style.display = 'block';
    status.style.display = 'none';

    // 顯示錯誤提示
    showNotification(message, 'error');
}

// 顯示格式驗證錯誤（帶詳細信息的彈窗）
function showValidationError(type, message, details) {
    const uploadBox = document.getElementById(`${type}-upload-box`);
    const status = document.getElementById(`${type}-status`);

    uploadBox.style.display = 'block';
    status.style.display = 'none';

    // 重置文件輸入
    const fileInput = document.getElementById(`${type}-file`);
    if (fileInput) {
        fileInput.value = '';
    }

    // 創建驗證錯誤彈窗
    showValidationModal(type, message, details);
}

// 顯示工作表保護錯誤（帶詳細說明的彈窗）
function showProtectionError(type, message, details) {
    const uploadBox = document.getElementById(`${type}-upload-box`);
    const status = document.getElementById(`${type}-status`);

    uploadBox.style.display = 'block';
    status.style.display = 'none';

    // 重置文件輸入
    const fileInput = document.getElementById(`${type}-file`);
    if (fileInput) {
        fileInput.value = '';
    }

    // 創建保護錯誤彈窗
    showProtectionModal(type, message, details);
}

// 顯示工作表保護錯誤彈窗
function showProtectionModal(type, message, details) {
    // 移除已存在的彈窗
    const existingModal = document.getElementById('protection-modal');
    if (existingModal) {
        existingModal.remove();
    }

    // 文件類型名稱對應
    const typeNames = {
        'erp': 'ERP 淨需求文件',
        'forecast': 'Forecast 文件',
        'transit': '在途文件'
    };
    const typeName = typeNames[type] || type;

    const modalHtml = `
        <div id="protection-modal" class="modal-overlay" style="display: flex;">
            <div class="modal-content" style="max-width: 500px;">
                <div class="modal-header" style="background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);">
                    <i class="fas fa-lock"></i>
                    <span>工作表保護錯誤</span>
                </div>
                <div class="modal-body" style="padding: 25px;">
                    <div style="text-align: center; margin-bottom: 20px;">
                        <i class="fas fa-exclamation-triangle" style="font-size: 48px; color: #e74c3c;"></i>
                    </div>
                    <div style="background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 15px; margin-bottom: 20px;">
                        <div style="font-weight: 600; color: #856404; margin-bottom: 8px;">
                            <i class="fas fa-info-circle"></i> ${message}
                        </div>
                        <div style="color: #856404; font-size: 0.9rem;">
                            ${details}
                        </div>
                    </div>
                    <div style="background: #e8f4fd; border-radius: 8px; padding: 15px;">
                        <div style="font-weight: 600; color: #0c5460; margin-bottom: 10px;">
                            <i class="fas fa-tools"></i> 操作步驟：
                        </div>
                        <ol style="margin: 0; padding-left: 20px; color: #0c5460; font-size: 0.9rem;">
                            <li>用 Excel 開啟 Forecast 檔案</li>
                            <li>點擊「檢閱」分頁</li>
                            <li>點擊「取消保護工作表」</li>
                            <li>儲存檔案後重新上傳</li>
                        </ol>
                    </div>
                </div>
                <div class="modal-footer" style="justify-content: center;">
                    <button class="btn btn-primary" onclick="closeProtectionModal()">
                        <i class="fas fa-check"></i> 我知道了
                    </button>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
}

// 關閉保護錯誤彈窗
function closeProtectionModal() {
    const modal = document.getElementById('protection-modal');
    if (modal) {
        modal.remove();
    }
}

// 顯示驗證錯誤彈窗
function showValidationModal(type, message, details) {
    // 移除已存在的彈窗
    const existingModal = document.getElementById('validation-modal');
    if (existingModal) {
        existingModal.remove();
    }

    // 文件類型名稱對應
    const typeNames = {
        'erp': 'ERP 淨需求文件',
        'forecast': 'Forecast 文件',
        'transit': '在途文件'
    };

    // 構建詳細錯誤列表
    let detailsHtml = '';
    if (details && details.length > 0) {
        detailsHtml = `
            <div class="validation-details">
                <div class="details-title">錯誤詳情：</div>
                <ul class="details-list">
                    ${details.map(d => `<li>${d}</li>`).join('')}
                </ul>
            </div>
        `;
    }

    // 創建彈窗
    const modal = document.createElement('div');
    modal.id = 'validation-modal';
    modal.className = 'validation-modal';
    modal.innerHTML = `
        <div class="validation-modal-overlay"></div>
        <div class="validation-modal-content">
            <div class="validation-modal-header">
                <i class="fas fa-exclamation-triangle"></i>
                <span>文件格式驗證失敗</span>
            </div>
            <div class="validation-modal-body">
                <div class="validation-file-type">${typeNames[type] || type}</div>
                <div class="validation-message">${message}</div>
                ${detailsHtml}
                <div class="validation-hint">
                    <i class="fas fa-info-circle"></i>
                    請確認上傳的文件格式與範本一致後重新上傳
                </div>
            </div>
            <div class="validation-modal-footer">
                <button class="validation-close-btn" onclick="closeValidationModal()">
                    <i class="fas fa-check"></i> 我知道了
                </button>
            </div>
        </div>
    `;

    // 添加樣式
    const style = document.createElement('style');
    style.id = 'validation-modal-style';
    style.textContent = `
        .validation-modal {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 10000;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .validation-modal-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(3px);
        }

        .validation-modal-content {
            position: relative;
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 500px;
            width: 90%;
            animation: modalSlideIn 0.3s ease-out;
            overflow: hidden;
        }

        @keyframes modalSlideIn {
            from {
                opacity: 0;
                transform: scale(0.9) translateY(-20px);
            }
            to {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
        }

        .validation-modal-header {
            background: linear-gradient(135deg, #e74c3c, #c0392b);
            color: white;
            padding: 20px 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 1.2em;
            font-weight: 600;
        }

        .validation-modal-header i {
            font-size: 1.5em;
        }

        .validation-modal-body {
            padding: 24px;
        }

        .validation-file-type {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 10px 16px;
            font-weight: 600;
            color: #495057;
            margin-bottom: 16px;
            display: inline-block;
        }

        .validation-message {
            color: #c0392b;
            font-size: 1.1em;
            font-weight: 500;
            margin-bottom: 16px;
            line-height: 1.5;
        }

        .validation-details {
            background: #fff5f5;
            border: 1px solid #fed7d7;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
        }

        .details-title {
            font-weight: 600;
            color: #c53030;
            margin-bottom: 8px;
        }

        .details-list {
            margin: 0;
            padding-left: 20px;
            color: #742a2a;
            font-size: 0.95em;
            line-height: 1.6;
        }

        .details-list li {
            margin-bottom: 4px;
        }

        .validation-hint {
            background: #e8f4fd;
            border-radius: 8px;
            padding: 12px 16px;
            color: #1a56db;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.95em;
        }

        .validation-modal-footer {
            padding: 16px 24px 24px;
            text-align: center;
        }

        .validation-close-btn {
            background: linear-gradient(135deg, #c87941, #a66332);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 32px;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .validation-close-btn:hover {
            background: linear-gradient(135deg, #a66332, #8b5228);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(166, 99, 50, 0.3);
        }
    `;

    // 添加到頁面
    if (!document.getElementById('validation-modal-style')) {
        document.head.appendChild(style);
    }
    document.body.appendChild(modal);

    // 點擊遮罩關閉
    modal.querySelector('.validation-modal-overlay').addEventListener('click', closeValidationModal);
}

// 關閉驗證錯誤彈窗
function closeValidationModal() {
    const modal = document.getElementById('validation-modal');
    if (modal) {
        modal.style.animation = 'modalSlideOut 0.2s ease-in forwards';
        setTimeout(() => {
            modal.remove();
        }, 200);
    }
}

// 添加彈窗關閉動畫
const modalCloseStyle = document.createElement('style');
modalCloseStyle.textContent = `
    @keyframes modalSlideOut {
        from {
            opacity: 1;
            transform: scale(1) translateY(0);
        }
        to {
            opacity: 0;
            transform: scale(0.9) translateY(-20px);
        }
    }
`;
document.head.appendChild(modalCloseStyle);

// 檢查上傳完成（更新 checklist 和下一步按鈕狀態）
function checkUploadComplete() {
    // 更新 checklist 狀態
    updateUploadChecklist();

    // 檢查是否全部上傳完成（根據 transitRequired 決定是否需要在途文件）
    const transitCheck = transitRequired ? uploadedFiles.transit : true;
    const allUploaded = uploadedFiles.erp && uploadedFiles.forecast && transitCheck;

    // 更新下一步按鈕狀態
    const nextBtn = document.getElementById('upload-next-btn');
    if (nextBtn) {
        nextBtn.disabled = !allUploaded;
    }

    // 顯示提示訊息
    if (allUploaded) {
        const msg = transitRequired || uploadedFiles.transit
            ? '所有文件上傳完成，請點擊「下一步」繼續'
            : '必填文件上傳完成（在途文件為選填），請點擊「下一步」繼續';
        showNotification(msg, 'success');
    } else {
        // 顯示還需要上傳哪些文件
        const missing = [];
        if (!uploadedFiles.erp) missing.push('ERP淨需求文件');
        if (!uploadedFiles.forecast) missing.push('Forecast文件');
        if (transitRequired && !uploadedFiles.transit) missing.push('在途文件');

        if (missing.length > 0 && (uploadedFiles.erp || uploadedFiles.forecast || uploadedFiles.transit)) {
            showNotification(`還需要上傳：${missing.join('、')}`, 'info');
        }
    }
}

// 更新上傳 checklist 狀態
function updateUploadChecklist() {
    const checkErp = document.getElementById('check-erp');
    const checkForecast = document.getElementById('check-forecast');
    const checkTransit = document.getElementById('check-transit');

    if (checkErp) {
        if (uploadedFiles.erp) {
            checkErp.classList.add('completed');
            checkErp.querySelector('i').className = 'fas fa-check-circle';
        } else {
            checkErp.classList.remove('completed');
            checkErp.querySelector('i').className = 'fas fa-circle';
        }
    }

    if (checkForecast) {
        if (uploadedFiles.forecast) {
            checkForecast.classList.add('completed');
            checkForecast.querySelector('i').className = 'fas fa-check-circle';
        } else {
            checkForecast.classList.remove('completed');
            checkForecast.querySelector('i').className = 'fas fa-circle';
        }
    }

    if (checkTransit) {
        if (uploadedFiles.transit) {
            checkTransit.classList.add('completed');
            checkTransit.classList.remove('optional', 'skipped');
            checkTransit.querySelector('i').className = 'fas fa-check-circle';
            checkTransit.querySelector('span').textContent = '在途文件';
        } else if (!transitRequired) {
            // 選填且未上傳時顯示為已跳過
            checkTransit.classList.add('skipped');
            checkTransit.classList.remove('completed', 'optional');
            checkTransit.querySelector('i').className = 'fas fa-minus-circle';
            checkTransit.querySelector('span').textContent = '在途文件（不須提供）';
        } else {
            checkTransit.classList.remove('completed', 'optional', 'skipped');
            checkTransit.querySelector('i').className = 'fas fa-circle';
            checkTransit.querySelector('span').textContent = '在途文件';
        }
    }
}

// 下一步：進入數據清理
function goToNextStep() {
    // 再次檢查是否全部上傳完成（根據 transitRequired 決定是否需要在途文件）
    const transitCheck = transitRequired ? uploadedFiles.transit : true;
    if (!uploadedFiles.erp || !uploadedFiles.forecast || !transitCheck) {
        const missing = [];
        if (!uploadedFiles.erp) missing.push('ERP淨需求文件');
        if (!uploadedFiles.forecast) missing.push('Forecast文件');
        if (transitRequired && !uploadedFiles.transit) missing.push('在途文件');

        showNotification(`請先完成上傳：${missing.join('、')}`, 'error');
        return;
    }

    // 顯示確認對話框
    showUploadConfirmModal();
}

// 顯示上傳確認對話框
function showUploadConfirmModal() {
    // 建立 Modal HTML - 手風琴設計
    const modalHtml = `
        <div class="modal-overlay" id="upload-confirm-modal">
            <div class="modal-container upload-confirm-modal">
                <div class="modal-header">
                    <h3><i class="fas fa-clipboard-check"></i> 確認上傳檔案</h3>
                    <button class="modal-close" onclick="closeUploadConfirmModal()">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                <div class="modal-body">
                    <p class="confirm-subtitle">點擊項目可展開查看詳細資訊</p>

                    <div class="confirm-file-list">
                        ${buildAccordionItem('erp', 'ERP 淨需求', 'fa-file-excel', uploadedFileInfo.erp)}
                        ${buildAccordionItem('forecast', 'Forecast', 'fa-chart-bar', uploadedFileInfo.forecast)}
                        ${buildAccordionItem('transit', '在途數據', 'fa-truck', uploadedFileInfo.transit)}
                    </div>

                    ${buildTransitRequirementBanner()}
                    ${buildMergeBanner()}
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" onclick="closeUploadConfirmModal()">
                        <i class="fas fa-arrow-left"></i> 返回修改
                    </button>
                    <button class="btn btn-primary" onclick="confirmAndProceed()">
                        <i class="fas fa-check"></i> 確認，進入下一步
                    </button>
                </div>
            </div>
        </div>
    `;

    // 移除舊的 modal（如果存在）
    const existingModal = document.getElementById('upload-confirm-modal');
    if (existingModal) {
        existingModal.remove();
    }

    // 插入 modal 到 body
    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // 顯示 modal
    setTimeout(() => {
        document.getElementById('upload-confirm-modal').classList.add('show');
    }, 10);
}

// 建立手風琴項目
function buildAccordionItem(type, title, icon, info) {
    // 處理在途文件跳過的情況
    if (type === 'transit' && !info && !transitRequired) {
        return `
            <div class="confirm-accordion-item transit-skipped-item">
                <div class="confirm-accordion-header">
                    <div class="accordion-icon ${type} skipped">
                        <i class="fas ${icon}"></i>
                    </div>
                    <div class="accordion-summary">
                        <div class="accordion-title">${title}</div>
                        <div class="accordion-brief skipped-text">已跳過（選填項目）</div>
                    </div>
                    <div class="accordion-status skipped">
                        <i class="fas fa-minus-circle"></i>
                    </div>
                </div>
            </div>
        `;
    }

    if (!info) return '';

    let briefText = '';
    let detailsHtml = '';

    if (type === 'forecast' && info.fileCount > 1) {
        // 多檔案 Forecast
        briefText = `${info.fileCount} 個檔案 · ${info.totalRows.toLocaleString()} 行`;
        detailsHtml = `
            <div class="accordion-details">
                <div class="accordion-details-row">
                    <span class="accordion-details-label">檔案數量</span>
                    <span class="accordion-details-value">${info.fileCount} 個</span>
                </div>
                <div class="accordion-details-row">
                    <span class="accordion-details-label">總資料行數</span>
                    <span class="accordion-details-value">${info.totalRows.toLocaleString()} 行</span>
                </div>
                ${info.totalSize ? `
                <div class="accordion-details-row">
                    <span class="accordion-details-label">總檔案大小</span>
                    <span class="accordion-details-value">${formatFileSize(info.totalSize)}</span>
                </div>
                ` : ''}
                <div class="accordion-file-list">
                    ${info.files.map(f => `<div class="accordion-file-item"><i class="fas fa-file-excel"></i> ${f}</div>`).join('')}
                </div>
            </div>
        `;
    } else if (type === 'forecast') {
        // 單檔案 Forecast
        briefText = `${info.files[0]}`;
        detailsHtml = `
            <div class="accordion-details">
                <div class="accordion-details-row">
                    <span class="accordion-details-label">檔案名稱</span>
                    <span class="accordion-details-value">${info.files[0]}</span>
                </div>
                <div class="accordion-details-row">
                    <span class="accordion-details-label">資料行數</span>
                    <span class="accordion-details-value">${info.totalRows.toLocaleString()} 行</span>
                </div>
                ${info.totalSize ? `
                <div class="accordion-details-row">
                    <span class="accordion-details-label">檔案大小</span>
                    <span class="accordion-details-value">${formatFileSize(info.totalSize)}</span>
                </div>
                ` : ''}
            </div>
        `;
    } else {
        // ERP 或 Transit
        briefText = info.filename;
        detailsHtml = `
            <div class="accordion-details">
                <div class="accordion-details-row">
                    <span class="accordion-details-label">檔案名稱</span>
                    <span class="accordion-details-value">${info.filename}</span>
                </div>
                <div class="accordion-details-row">
                    <span class="accordion-details-label">資料行數</span>
                    <span class="accordion-details-value">${info.rows.toLocaleString()} 行</span>
                </div>
                <div class="accordion-details-row">
                    <span class="accordion-details-label">欄位數量</span>
                    <span class="accordion-details-value">${info.columns} 欄</span>
                </div>
                ${info.fileSize ? `
                <div class="accordion-details-row">
                    <span class="accordion-details-label">檔案大小</span>
                    <span class="accordion-details-value">${formatFileSize(info.fileSize)}</span>
                </div>
                ` : ''}
            </div>
        `;
    }

    // 截斷過長的摘要文字
    const shortBrief = briefText.length > 30 ? briefText.substring(0, 27) + '...' : briefText;

    return `
        <div class="confirm-accordion-item" onclick="toggleAccordion(this)">
            <div class="confirm-accordion-header">
                <div class="accordion-icon ${type}">
                    <i class="fas ${icon}"></i>
                </div>
                <div class="accordion-summary">
                    <div class="accordion-title">${title}</div>
                    <div class="accordion-brief" title="${briefText}">${shortBrief}</div>
                </div>
                <div class="accordion-status">
                    <i class="fas fa-check-circle"></i>
                    <i class="fas fa-chevron-down accordion-toggle"></i>
                </div>
            </div>
            <div class="confirm-accordion-content">
                ${detailsHtml}
            </div>
        </div>
    `;
}

// 建立獨立的合併狀態橫幅（直接讀取合併區塊 checkbox 的當前狀態）
function buildMergeBanner() {
    const info = uploadedFileInfo.forecast;
    if (!info) return '';

    // 直接讀取合併區塊 checkbox 的當前狀態，而不是上傳時儲存的值
    const mergeCheckbox = document.getElementById('merge-forecast-files');
    const shouldMerge = mergeCheckbox ? mergeCheckbox.checked : true;

    // 只有一個檔案時顯示不適用
    if (info.fileCount <= 1) {
        return `
            <div class="confirm-merge-banner merge-na">
                <div class="merge-icon">
                    <i class="fas fa-file-alt"></i>
                </div>
                <span class="merge-text">Forecast 合併選項：不適用（僅單一檔案）</span>
            </div>
        `;
    }

    if (shouldMerge) {
        return `
            <div class="confirm-merge-banner merge-enabled">
                <div class="merge-icon">
                    <i class="fas fa-object-group"></i>
                </div>
                <span class="merge-text">Forecast 下載時將合併為單一檔案</span>
            </div>
        `;
    } else {
        return `
            <div class="confirm-merge-banner merge-disabled">
                <div class="merge-icon">
                    <i class="fas fa-copy"></i>
                </div>
                <span class="merge-text">Forecast 下載時將分別輸出各檔案</span>
            </div>
        `;
    }
}

// 建立在途需求提示區塊
function buildTransitRequirementBanner() {
    // 如果沒有在途檢查結果，不顯示
    if (!forecastTransitCheck || !forecastTransitCheck.message) {
        return '';
    }

    const check = forecastTransitCheck;
    let bannerClass = 'transit-info';
    let icon = 'fa-check-circle';

    if (check.has_transit_requirement) {
        bannerClass = 'transit-warning';
        icon = 'fa-exclamation-triangle';
    }

    // 建立廠區列表
    let regionDetails = '';
    if (check.transit_not_required_regions && check.transit_not_required_regions.length > 0) {
        regionDetails += `<div class="transit-region-list">
            <span class="region-label">不需在途：</span>
            <span class="region-values">${check.transit_not_required_regions.join('、')}</span>
        </div>`;
    }
    if (check.transit_required_regions && check.transit_required_regions.length > 0) {
        regionDetails += `<div class="transit-region-list">
            <span class="region-label">需要在途：</span>
            <span class="region-values warning">${check.transit_required_regions.join('、')}</span>
        </div>`;
    }

    return `
        <div class="confirm-transit-banner ${bannerClass}">
            <div class="transit-banner-header">
                <div class="transit-icon">
                    <i class="fas ${icon}"></i>
                </div>
                <span class="transit-text">${check.message}</span>
            </div>
            ${regionDetails}
        </div>
    `;
}

// 切換手風琴展開/收合
function toggleAccordion(element) {
    element.classList.toggle('expanded');
}

// 關閉確認對話框
function closeUploadConfirmModal() {
    const modal = document.getElementById('upload-confirm-modal');
    if (modal) {
        modal.classList.remove('show');
        setTimeout(() => {
            modal.remove();
        }, 300);
    }
}

// 確認並進入下一步
async function confirmAndProceed() {
    // 檢查是否有勾選合併且上傳了多個 Forecast 檔案
    const mergeCheckbox = document.getElementById('merge-forecast-files');
    const shouldMerge = mergeCheckbox ? mergeCheckbox.checked : true;
    const forecastInfo = uploadedFileInfo.forecast;
    const hasMultipleFiles = forecastInfo && forecastInfo.fileCount > 1;

    if (shouldMerge && hasMultipleFiles) {
        // 需要合併，先呼叫後端 API 進行合併
        try {
            // 更新 modal 顯示合併中
            const confirmBtn = document.querySelector('#upload-confirm-modal .btn-primary');
            if (confirmBtn) {
                confirmBtn.disabled = true;
                confirmBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 合併檔案中...';
            }

            const response = await fetch('/merge_forecast_files', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    upload_session_id: getUploadSessionId(),
                    ...getTestModeJsonParams()
                })
            });

            const result = await response.json();

            if (!result.success) {
                showNotification(`合併失敗: ${result.message}`, 'error');
                if (confirmBtn) {
                    confirmBtn.disabled = false;
                    confirmBtn.innerHTML = '<i class="fas fa-check"></i> 確認，進入下一步';
                }
                return;
            }

            if (result.merged) {
                console.log(`Forecast 檔案已合併: ${result.merged_filename}`);
                // 更新上傳檔案資訊
                uploadedFileInfo.forecast.merged = true;
                uploadedFileInfo.forecast.mergedFilename = result.merged_filename;
            }

        } catch (error) {
            console.error('合併檔案失敗:', error);
            showNotification('合併檔案失敗，請重試', 'error');
            const confirmBtn = document.querySelector('#upload-confirm-modal .btn-primary');
            if (confirmBtn) {
                confirmBtn.disabled = false;
                confirmBtn.innerHTML = '<i class="fas fa-check"></i> 確認，進入下一步';
            }
            return;
        }
    }

    closeUploadConfirmModal();

    // 更新進度
    currentStep = 2;
    updateProgress();

    // 顯示清理區域
    showSection('cleanup');

    showNotification('進入數據清理階段', 'success');
}

// 數據清理處理（支援多檔案）
async function handleCleanup() {
    try {
        showButtonLoading('cleanup-btn');
        showProgress('cleanup-progress', 'cleanup-progress-fill', 'cleanup-progress-text');

        // 模擬進度更新
        updateProgressBar('cleanup-progress-fill', 'cleanup-progress-text', 0, '開始清理數據...');
        await sleep(500);

        updateProgressBar('cleanup-progress-fill', 'cleanup-progress-text', 30, '讀取Forecast文件...');
        await sleep(500);

        updateProgressBar('cleanup-progress-fill', 'cleanup-progress-text', 60, '清理供應數量數據...');
        await sleep(500);

        updateProgressBar('cleanup-progress-fill', 'cleanup-progress-text', 80, '清理庫存數量數據...');
        await sleep(500);

        const response = await fetch('/process_forecast_cleanup', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                upload_session_id: getUploadSessionId(),
                ...getTestModeJsonParams()
            })
        });

        const result = await response.json();

        updateProgressBar('cleanup-progress-fill', 'cleanup-progress-text', 100, '清理完成！');
        await sleep(500);

        hideProgress('cleanup-progress');

        if (result.success) {
            // 根據是否為多檔案模式顯示不同的結果
            if (result.multi_file && result.files && result.files.length > 0) {
                showMultiFileCleanupResult(result);
            } else {
                showProcessResult('cleanup-result', result.message, 'success');
            }
            currentStep = 3;
            updateProgress();
            showSection('mapping');
            showNotification('數據清理完成，可以配置映射表', 'success');
        } else {
            showProcessResult('cleanup-result', result.message, 'error');
        }
    } catch (error) {
        hideProgress('cleanup-progress');
        showProcessResult('cleanup-result', '清理失敗: ' + error.message, 'error');
    } finally {
        hideButtonLoading('cleanup-btn');
    }
}

// 顯示多檔案清理結果
function showMultiFileCleanupResult(result) {
    const resultContainer = document.getElementById('cleanup-result');
    if (!resultContainer) return;

    const fileCount = result.file_count || 0;
    const totalCells = result.total_cleaned_cells || 0;
    const files = result.files || [];

    let filesHtml = '';
    files.forEach((file, index) => {
        const statusIcon = file.status === 'success'
            ? '<i class="fas fa-check-circle" style="color: #27ae60;"></i>'
            : '<i class="fas fa-times-circle" style="color: #e74c3c;"></i>';
        const statusText = file.status === 'success'
            ? `清理 ${file.cleaned_cells} 個單元格`
            : (file.message || '清理失敗');
        const shortName = file.name.length > 30 ? file.name.substring(0, 27) + '...' : file.name;

        filesHtml += `
            <div class="cleanup-file-item">
                ${statusIcon}
                <span class="file-name" title="${file.name}">${shortName}</span>
                <span class="file-status">${statusText}</span>
            </div>
        `;
    });

    resultContainer.innerHTML = `
        <div class="result-box success">
            <div class="result-header">
                <i class="fas fa-check-circle"></i>
                <span>數據清理完成</span>
            </div>
            <div class="result-summary">
                <strong>${fileCount} 個檔案</strong>，共清理 <strong>${totalCells}</strong> 個單元格
            </div>
            <div class="cleanup-files-list">
                ${filesHtml}
            </div>
        </div>
    `;
    resultContainer.style.display = 'block';
}

// 打開映射配置
function openMappingConfig() {
    // 構建 URL，包含 IT 測試模式參數
    let url = '/mapping';
    if (isITTestMode()) {
        url += `?test_mode=true&customer_id=${getTestCustomerId()}&customer_name=${encodeURIComponent(window.IT_TEST_MODE.selectedCustomer.display_name)}`;
    }

    // 在新窗口打開映射配置頁面
    const mappingWindow = window.open(url, '_blank', 'width=1200,height=800,scrollbars=yes,resizable=yes');

    // 監聯窗口關閉事件
    const checkClosed = setInterval(() => {
        if (mappingWindow.closed) {
            clearInterval(checkClosed);
            // 窗口關閉後，檢查是否有新的映射數據
            checkMappingData();
        }
    }, 1000);
}

// 檢查映射數據
async function checkMappingData() {
    try {
        let url = '/get_mapping_data';
        if (isITTestMode()) {
            url += `?test_mode=true&customer_id=${getTestCustomerId()}`;
        }
        const response = await fetch(url);
        const result = await response.json();
        
        if (result.success && result.source === 'mapping_table') {
            showNotification('映射配置已更新', 'success');
        }
    } catch (error) {
        console.log('檢查映射數據失敗:', error);
    }
}

// 映射處理
async function handleMappingProcess() {
    try {
        showButtonLoading('mapping-process-btn');
        showProgress('mapping-progress', 'mapping-progress-fill', 'mapping-progress-text');

        // 模擬進度更新
        updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 0, '開始映射整合...');
        await sleep(500);

        updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 20, '讀取ERP文件...');
        await sleep(500);

        if (transitRequired && uploadedFiles.transit) {
            updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 40, '讀取在途文件...');
            await sleep(500);
        } else if (!transitRequired) {
            updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 40, '跳過在途文件（不須提供）...');
            await sleep(300);
        }

        updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 60, '讀取映射配置...');
        await sleep(500);

        updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 80, '應用映射關係...');
        await sleep(500);

        const response = await fetch('/process_erp_mapping', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                upload_session_id: getUploadSessionId(),
                transit_required: transitRequired,
                ...getTestModeJsonParams()
            })
        });

        const result = await response.json();

        updateProgressBar('mapping-progress-fill', 'mapping-progress-text', 100, '整合完成！');
        await sleep(500);

        hideProgress('mapping-progress');

        if (result.success) {
            let detailMessage = `
                <div style="margin-top: 10px;">
                    <div>✅ ERP 數據整合完成：${result.erp_rows} 行</div>
            `;

            // 根據是否有在途數據顯示不同的訊息
            if (result.transit_skipped) {
                detailMessage += `<div>⏭️ 在途數據：已跳過（選填項目）</div>`;
            } else {
                detailMessage += `<div>✅ 在途數據整合完成：${result.transit_rows} 行</div>`;
            }

            detailMessage += `
                    <div style="margin-top: 8px; font-size: 0.9em; color: #666;">
                        <div>• ${result.erp_file}</div>
            `;

            if (!result.transit_skipped) {
                detailMessage += `<div>• ${result.transit_file}</div>`;
            }

            detailMessage += `
                    </div>
                </div>
            `;
            showProcessResult('mapping-result', result.message + detailMessage, 'success');
            currentStep = 4;
            updateProgress();
            showSection('forecast');
            showNotification('ERP 和在途數據整合完成，可以開始FORECAST處理', 'success');
        } else {
            showProcessResult('mapping-result', result.message, 'error');
        }
    } catch (error) {
        hideProgress('mapping-progress');
        showProcessResult('mapping-result', '整合失敗: ' + error.message, 'error');
    } finally {
        hideButtonLoading('mapping-process-btn');
    }
}

// FORECAST處理
async function handleForecast() {
    try {
        showButtonLoading('forecast-btn');
        showProgress('forecast-progress', 'forecast-progress-fill', 'forecast-progress-text');
        
        // 模擬進度更新
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 0, '開始FORECAST處理...');
        await sleep(500);
        
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 20, '載入數據文件...');
        await sleep(500);
        
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 40, '識別數據塊...');
        await sleep(500);
        
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 60, '匹配客戶數據...');
        await sleep(500);
        
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 80, '計算目標日期...');
        await sleep(500);
        
        const response = await fetch('/run_forecast', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                upload_session_id: getUploadSessionId(),
                transit_required: transitRequired,
                ...getTestModeJsonParams()
            })
        });
        
        const result = await response.json();
        
        updateProgressBar('forecast-progress-fill', 'forecast-progress-text', 100, 'FORECAST處理完成！');
        await sleep(500);
        
        hideProgress('forecast-progress');
        
        if (result.success) {
            let detailMessage = '';

            if (result.multi_file) {
                // 多檔案模式
                detailMessage = `
                    <div style="margin-top: 10px;">
                        <div style="font-weight: 600; margin-bottom: 8px;">📁 多檔案處理結果：</div>
                        <div style="margin-left: 20px;">
                            <div>📊 處理檔案數：${result.success_count}/${result.file_count} 個成功</div>
                        </div>
                        <div style="font-weight: 600; margin-top: 12px; margin-bottom: 8px;">📊 ERP 數據填寫結果（合計）：</div>
                        <div style="margin-left: 20px;">
                            <div>✅ 成功填寫：${result.total_erp_filled} 筆</div>
                            <div>⏭️  跳過記錄：${result.total_erp_skipped} 筆</div>
                        </div>
                `;

                // 如果有 Transit 數據
                if (result.transit_file_skipped) {
                    detailMessage += `
                        <div style="font-weight: 600; margin-top: 12px; margin-bottom: 8px;">🚚 Transit 數據：</div>
                        <div style="margin-left: 20px;">
                            <div style="color: #888;">⏭️ 已跳過（選填項目）</div>
                        </div>
                    `;
                } else if (result.total_transit_filled !== undefined) {
                    detailMessage += `
                        <div style="font-weight: 600; margin-top: 12px; margin-bottom: 8px;">🚚 Transit 數據填寫結果（合計）：</div>
                        <div style="margin-left: 20px;">
                            <div>✅ 成功填寫：${result.total_transit_filled} 筆</div>
                            <div>⏭️  跳過記錄：${result.total_transit_skipped} 筆</div>
                        </div>
                    `;
                }

                detailMessage += '</div>';

                // 更新下載區域為多檔案模式
                updateDownloadSectionMultiFile(result.files);
            } else {
                // 單檔案模式（合併模式）
                detailMessage = `
                    <div style="margin-top: 10px;">
                        <div style="font-weight: 600; margin-bottom: 8px;">📊 ERP 數據填寫結果：</div>
                        <div style="margin-left: 20px;">
                            <div>✅ 成功填寫：${result.erp_filled} 筆</div>
                            <div>⏭️  跳過記錄：${result.erp_skipped} 筆</div>
                        </div>
                `;

                // 如果有 Transit 數據
                if (result.transit_file_skipped) {
                    detailMessage += `
                        <div style="font-weight: 600; margin-top: 12px; margin-bottom: 8px;">🚚 Transit 數據：</div>
                        <div style="margin-left: 20px;">
                            <div style="color: #888;">⏭️ 已跳過（選填項目）</div>
                        </div>
                    `;
                } else if (result.transit_filled !== undefined) {
                    detailMessage += `
                        <div style="font-weight: 600; margin-top: 12px; margin-bottom: 8px;">🚚 Transit 數據填寫結果：</div>
                        <div style="margin-left: 20px;">
                            <div>✅ 成功填寫：${result.transit_filled} 筆</div>
                            <div>⏭️  跳過記錄：${result.transit_skipped} 筆</div>
                        </div>
                    `;
                }

                detailMessage += '</div>';

                // 更新下載區域為單檔案模式，傳入檔名
                updateDownloadSectionSingleFile(result.file || 'forecast_result.xlsx');
            }

            showProcessResult('forecast-result', result.message + detailMessage, 'success');

            // 更新完成資訊（日期和時間）
            updateCompletionInfo();

            showSection('download');
            showNotification('FORECAST處理完成！', 'success');
        } else {
            showProcessResult('forecast-result', result.message, 'error');
        }
    } catch (error) {
        hideProgress('forecast-progress');
        showProcessResult('forecast-result', '處理失敗: ' + error.message, 'error');
    } finally {
        hideButtonLoading('forecast-btn');
    }
}

// 下載文件
function downloadFile(filename) {
    window.open(`/download/${filename}`, '_blank');
}

// 更新下載區域為單檔案模式
function updateDownloadSectionSingleFile(filename) {
    // 使用傳入的檔名，預設為 forecast_result.xlsx
    const outputFilename = filename || 'forecast_result.xlsx';
    const downloadContainer = document.querySelector('.download-single');
    if (downloadContainer) {
        // 移除多檔案模式的 class
        downloadContainer.classList.remove('multi-file-mode');

        downloadContainer.innerHTML = `
            <div class="download-item">
                <i class="fas fa-file-excel"></i>
                <div class="download-info">
                    <div class="download-title">FORECAST處理結果</div>
                    <div class="download-desc">${outputFilename}</div>
                </div>
                <button class="btn btn-primary" onclick="downloadFile('${outputFilename}')">
                    <i class="fas fa-download"></i> 下載
                </button>
            </div>
        `;
    }
}

// 更新下載區域為多檔案模式
function updateDownloadSectionMultiFile(files) {
    const downloadContainer = document.querySelector('.download-single');
    if (downloadContainer && files && files.length > 0) {
        // 添加多檔案模式的 class，啟用垂直排列樣式
        downloadContainer.classList.add('multi-file-mode');

        // 儲存檔案列表供批量下載使用
        window.multiFileDownloadList = files.map(f => f.output);

        // 批量下載按鈕（放在最上方）
        let downloadAllBtnHtml = '';
        if (files.length > 1) {
            downloadAllBtnHtml = `
                <div class="download-all-section">
                    <button class="btn btn-download-all" onclick="downloadAllFiles()">
                        <i class="fas fa-cloud-download-alt"></i>
                        批量下載全部檔案（共 ${files.length} 個）
                    </button>
                </div>
            `;
        }

        // 檔案列表（垂直排列，可捲動）
        let filesListHtml = '<div class="download-files-list">';
        files.forEach((file, index) => {
            filesListHtml += `
                <div class="download-file-row">
                    <div class="download-file-info">
                        <i class="fas fa-file-excel file-icon"></i>
                        <div class="file-details">
                            <div class="file-name">${file.output}</div>
                            <div class="file-stats">ERP: ${file.erp_filled} 筆 | Transit: ${file.transit_filled} 筆</div>
                        </div>
                    </div>
                    <button class="btn btn-outline-primary btn-sm" onclick="downloadFile('${file.output}')">
                        <i class="fas fa-download"></i> 下載
                    </button>
                </div>
            `;
        });
        filesListHtml += '</div>';

        downloadContainer.innerHTML = downloadAllBtnHtml + filesListHtml;
    }
}

// 批量下載所有檔案
function downloadAllFiles() {
    const filenames = window.multiFileDownloadList || [];
    if (filenames.length === 0) {
        showNotification('沒有可下載的檔案', 'error');
        return;
    }

    showNotification(`開始下載 ${filenames.length} 個檔案...`, 'info');

    filenames.forEach((filename, index) => {
        // 延遲下載避免瀏覽器阻擋
        setTimeout(() => {
            // 使用 iframe 方式下載，避免開啟多個分頁
            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.src = `/download/${filename}`;
            document.body.appendChild(iframe);

            // 5秒後移除 iframe
            setTimeout(() => {
                document.body.removeChild(iframe);
            }, 5000);
        }, index * 800);  // 每個檔案間隔 800ms
    });
}

// 更新完成資訊（日期和時間）
function updateCompletionInfo() {
    const now = new Date();

    // 格式化日期：YYYY/MM/DD
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const dateStr = `${year}/${month}/${day}`;

    // 格式化時間：HH:MM:SS
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const timeStr = `${hours}:${minutes}:${seconds}`;

    // 更新顯示
    const dateElement = document.getElementById('completion-date');
    const timeElement = document.getElementById('completion-time');
    const customerElement = document.getElementById('completion-customer');

    if (dateElement) dateElement.textContent = dateStr;
    if (timeElement) timeElement.textContent = timeStr;

    // 如果是 IT 測試模式，顯示測試客戶名稱
    // 非測試模式時保留模板中設定的客戶名稱
    if (customerElement && isITTestMode()) {
        const customer = getTestCustomerInfo();
        customerElement.textContent = `${customer.company} - ${customer.display_name}`;
    }
}

// 更新進度
function updateProgress() {
    progressSteps.forEach((step, index) => {
        const stepNumber = index + 1;
        step.classList.remove('active', 'completed');
        
        if (stepNumber < currentStep) {
            step.classList.add('completed');
        } else if (stepNumber === currentStep) {
            step.classList.add('active');
        }
    });
}

// 顯示區域
function showSection(sectionName) {
    // 隱藏所有區域
    Object.values(sections).forEach(section => {
        if (section) {
            section.classList.add('section-hidden');
            section.classList.remove('section-visible');
        }
    });
    
    // 顯示指定區域
    const targetSection = sections[sectionName];
    if (targetSection) {
        targetSection.classList.remove('section-hidden');
        targetSection.classList.add('section-visible');
    }
}

// 顯示處理結果
function showProcessResult(elementId, message, type) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerHTML = message;
        element.className = `process-result ${type}`;
        element.style.display = 'block';
    }
}

// 顯示按鈕載入狀態
function showButtonLoading(buttonId) {
    const button = document.getElementById(buttonId);
    if (button) {
        button.disabled = true;
        button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 處理中...';
    }
}

// 隱藏按鈕載入狀態
function hideButtonLoading(buttonId) {
    const button = document.getElementById(buttonId);
    if (button) {
        button.disabled = false;
        // 恢復原始文字（需要根據按鈕類型調整）
        if (buttonId === 'cleanup-btn') {
            button.innerHTML = '<i class="fas fa-play"></i> 開始清理';
        } else if (buttonId === 'mapping-process-btn') {
            button.innerHTML = '<i class="fas fa-play"></i> 開始整合';
        } else if (buttonId === 'forecast-btn') {
            button.innerHTML = '<i class="fas fa-play"></i> 開始FORECAST處理';
        }
    }
}

// 進度條相關函數
function showProgress(containerId, fillId, textId) {
    const container = document.getElementById(containerId);
    if (container) {
        container.style.display = 'block';
    }
}

function hideProgress(containerId) {
    const container = document.getElementById(containerId);
    if (container) {
        container.style.display = 'none';
    }
}

function updateProgressBar(fillId, textId, percentage, text) {
    const fill = document.getElementById(fillId);
    const textElement = document.getElementById(textId);
    
    if (fill) {
        fill.style.width = percentage + '%';
    }
    
    if (textElement) {
        textElement.textContent = text;
    }
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
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

// ========================================
// 回上一步功能
// ========================================

// 步驟名稱對應
const stepNames = {
    1: '文件上傳',
    2: '數據清理',
    3: '映射整合',
    4: 'FORECAST處理',
    5: '結果下載'
};

// 確認回到上一步
function confirmGoBack(fromStep) {
    const targetStep = fromStep;
    const targetStepName = stepNames[targetStep];

    // 創建確認彈窗
    showGoBackModal(fromStep, targetStepName);
}

// 顯示回上一步確認彈窗
function showGoBackModal(fromStep, targetStepName) {
    // 移除已存在的彈窗
    const existingModal = document.getElementById('goback-modal');
    if (existingModal) {
        existingModal.remove();
    }

    // 創建彈窗
    const modal = document.createElement('div');
    modal.id = 'goback-modal';
    modal.className = 'goback-modal';
    modal.innerHTML = `
        <div class="goback-modal-overlay"></div>
        <div class="goback-modal-content">
            <div class="goback-modal-header">
                <i class="fas fa-exclamation-circle"></i>
                <span>確認返回上一步</span>
            </div>
            <div class="goback-modal-body">
                <div class="goback-warning">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>注意！此操作無法復原</span>
                </div>
                <div class="goback-message">
                    您即將返回「<strong>${targetStepName}</strong>」步驟。
                </div>
                <div class="goback-info">
                    <p>返回後，以下變更將會發生：</p>
                    <ul>
                        ${getResetInfo(fromStep)}
                    </ul>
                </div>
                <div class="goback-hint">
                    <i class="fas fa-info-circle"></i>
                    返回後需要重新執行該步驟的操作
                </div>
            </div>
            <div class="goback-modal-footer">
                <button class="goback-cancel-btn" onclick="closeGoBackModal()">
                    <i class="fas fa-times"></i> 取消
                </button>
                <button class="goback-confirm-btn" onclick="executeGoBack(${fromStep})">
                    <i class="fas fa-arrow-left"></i> 確定返回
                </button>
            </div>
        </div>
    `;

    // 添加樣式
    addGoBackModalStyles();

    // 添加到頁面
    document.body.appendChild(modal);

    // 點擊遮罩關閉
    modal.querySelector('.goback-modal-overlay').addEventListener('click', closeGoBackModal);
}

// 獲取重置信息
function getResetInfo(fromStep) {
    const resetInfoMap = {
        1: '<li>需要重新上傳所有文件</li>',
        2: '<li>數據清理結果將被清除</li><li>需要重新執行數據清理</li>',
        3: '<li>映射整合結果將被清除</li><li>需要重新執行映射整合</li>',
        4: '<li>FORECAST處理結果將被清除</li><li>需要重新執行FORECAST處理</li>'
    };
    return resetInfoMap[fromStep] || '<li>需要重新執行該步驟</li>';
}

// 添加回上一步彈窗樣式
function addGoBackModalStyles() {
    if (document.getElementById('goback-modal-style')) return;

    const style = document.createElement('style');
    style.id = 'goback-modal-style';
    style.textContent = `
        .goback-modal {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 10000;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .goback-modal-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(3px);
        }

        .goback-modal-content {
            position: relative;
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 480px;
            width: 90%;
            animation: gobackModalSlideIn 0.3s ease-out;
            overflow: hidden;
        }

        @keyframes gobackModalSlideIn {
            from {
                opacity: 0;
                transform: scale(0.9) translateY(-20px);
            }
            to {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
        }

        @keyframes gobackModalSlideOut {
            from {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
            to {
                opacity: 0;
                transform: scale(0.9) translateY(-20px);
            }
        }

        .goback-modal-header {
            background: linear-gradient(135deg, #f39c12, #e67e22);
            color: white;
            padding: 20px 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 1.2em;
            font-weight: 600;
        }

        .goback-modal-header i {
            font-size: 1.5em;
        }

        .goback-modal-body {
            padding: 24px;
        }

        .goback-warning {
            background: #fff3cd;
            border: 1px solid #ffc107;
            border-radius: 8px;
            padding: 12px 16px;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 10px;
            color: #856404;
            font-weight: 600;
        }

        .goback-warning i {
            color: #f39c12;
            font-size: 1.2em;
        }

        .goback-message {
            font-size: 1.1em;
            color: #333;
            margin-bottom: 16px;
            line-height: 1.5;
        }

        .goback-message strong {
            color: #c87941;
        }

        .goback-info {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
        }

        .goback-info p {
            margin: 0 0 10px 0;
            font-weight: 600;
            color: #495057;
        }

        .goback-info ul {
            margin: 0;
            padding-left: 20px;
            color: #6c757d;
            font-size: 0.95em;
            line-height: 1.6;
        }

        .goback-info li {
            margin-bottom: 4px;
        }

        .goback-hint {
            background: #e8f4fd;
            border-radius: 8px;
            padding: 12px 16px;
            color: #1a56db;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.95em;
        }

        .goback-modal-footer {
            padding: 16px 24px 24px;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
        }

        .goback-cancel-btn {
            background: #6c757d;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 24px;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .goback-cancel-btn:hover {
            background: #5a6268;
        }

        .goback-confirm-btn {
            background: linear-gradient(135deg, #f39c12, #e67e22);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 24px;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .goback-confirm-btn:hover {
            background: linear-gradient(135deg, #e67e22, #d35400);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(243, 156, 18, 0.3);
        }
    `;
    document.head.appendChild(style);
}

// 關閉回上一步彈窗
function closeGoBackModal() {
    const modal = document.getElementById('goback-modal');
    if (modal) {
        modal.querySelector('.goback-modal-content').style.animation = 'gobackModalSlideOut 0.2s ease-in forwards';
        setTimeout(() => {
            modal.remove();
        }, 200);
    }
}

// 執行返回上一步
function executeGoBack(fromStep) {
    closeGoBackModal();

    // 根據來源步驟執行不同的重置邏輯
    switch(fromStep) {
        case 1: // 從數據清理返回文件上傳
            goToStep1();
            break;
        case 2: // 從映射整合返回數據清理
            goToStep2();
            break;
        case 3: // 從FORECAST處理返回映射整合
            goToStep3();
            break;
        case 4: // 從結果下載返回FORECAST處理
            goToStep4();
            break;
    }

    showNotification('已返回上一步，請重新執行操作', 'info');
}

// 返回步驟1：文件上傳
async function goToStep1() {
    // 重置前端的上傳 Session ID（產生新的資料夾 ID）
    resetUploadSessionId();
    console.log('上傳 Session ID 已重置，下次上傳將使用新的資料夾');

    // 重置上傳狀態
    uploadedFiles = {
        erp: false,
        forecast: false,
        transit: false
    };

    // 重置UI - 上傳區域
    ['erp', 'forecast', 'transit'].forEach(type => {
        const uploadBox = document.getElementById(`${type}-upload-box`);
        const status = document.getElementById(`${type}-status`);
        const fileInput = document.getElementById(`${type}-file`);

        if (uploadBox) uploadBox.style.display = 'block';
        if (status) status.style.display = 'none';
        if (fileInput) fileInput.value = '';
    });

    // 重置 checklist 狀態
    updateUploadChecklist();

    // 重置下一步按鈕狀態
    const nextBtn = document.getElementById('upload-next-btn');
    if (nextBtn) {
        nextBtn.disabled = true;
    }

    // 隱藏處理結果
    hideAllResults();

    // 更新步驟
    currentStep = 1;
    updateProgress();
    showSection('upload');
}

// 返回步驟2：數據清理
function goToStep2() {
    // 隱藏清理結果
    const cleanupResult = document.getElementById('cleanup-result');
    if (cleanupResult) {
        cleanupResult.style.display = 'none';
        cleanupResult.innerHTML = '';
    }

    // 更新步驟
    currentStep = 2;
    updateProgress();
    showSection('cleanup');
}

// 返回步驟3：映射整合
function goToStep3() {
    // 隱藏映射結果
    const mappingResult = document.getElementById('mapping-result');
    if (mappingResult) {
        mappingResult.style.display = 'none';
        mappingResult.innerHTML = '';
    }

    // 更新步驟
    currentStep = 3;
    updateProgress();
    showSection('mapping');
}

// 返回步驟4：FORECAST處理
function goToStep4() {
    // 隱藏FORECAST結果
    const forecastResult = document.getElementById('forecast-result');
    if (forecastResult) {
        forecastResult.style.display = 'none';
        forecastResult.innerHTML = '';
    }

    // 隱藏下載區域
    const downloadSection = document.getElementById('download-section');
    if (downloadSection) {
        downloadSection.style.display = 'none';
    }

    // 更新步驟
    currentStep = 4;
    updateProgress();
    showSection('forecast');
}

// 隱藏所有處理結果
function hideAllResults() {
    const resultIds = ['cleanup-result', 'mapping-result', 'forecast-result'];
    resultIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.style.display = 'none';
            element.innerHTML = '';
        }
    });
}

// 確認重新開始
function confirmRestart() {
    // 創建確認彈窗
    showRestartModal();
}

// 顯示重新開始確認彈窗
function showRestartModal() {
    // 移除已存在的彈窗
    const existingModal = document.getElementById('goback-modal');
    if (existingModal) {
        existingModal.remove();
    }

    // 創建彈窗
    const modal = document.createElement('div');
    modal.id = 'goback-modal';
    modal.className = 'goback-modal';
    modal.innerHTML = `
        <div class="goback-modal-overlay"></div>
        <div class="goback-modal-content">
            <div class="goback-modal-header" style="background: linear-gradient(135deg, #e74c3c, #c0392b);">
                <i class="fas fa-redo"></i>
                <span>確認重新開始</span>
            </div>
            <div class="goback-modal-body">
                <div class="goback-warning" style="background: #f8d7da; border-color: #f5c6cb; color: #721c24;">
                    <i class="fas fa-exclamation-triangle" style="color: #e74c3c;"></i>
                    <span>注意！所有進度將被清除</span>
                </div>
                <div class="goback-message">
                    您即將重新開始整個流程。
                </div>
                <div class="goback-info">
                    <p>重新開始後，以下內容將被清除：</p>
                    <ul>
                        <li>所有已上傳的文件</li>
                        <li>數據清理結果</li>
                        <li>映射整合結果</li>
                        <li>FORECAST處理結果</li>
                    </ul>
                </div>
                <div class="goback-hint">
                    <i class="fas fa-info-circle"></i>
                    確定要重新開始嗎？
                </div>
            </div>
            <div class="goback-modal-footer">
                <button class="goback-cancel-btn" onclick="closeGoBackModal()">
                    <i class="fas fa-times"></i> 取消
                </button>
                <button class="goback-confirm-btn" style="background: linear-gradient(135deg, #e74c3c, #c0392b);" onclick="executeRestart()">
                    <i class="fas fa-redo"></i> 確定重新開始
                </button>
            </div>
        </div>
    `;

    // 添加樣式
    addGoBackModalStyles();

    // 添加到頁面
    document.body.appendChild(modal);

    // 點擊遮罩關閉
    modal.querySelector('.goback-modal-overlay').addEventListener('click', closeGoBackModal);
}

// 執行重新開始
async function executeRestart() {
    closeGoBackModal();
    // goToStep1 內部會自動重置 session
    await goToStep1();
    showNotification('已重新開始，請上傳文件', 'info');
}
