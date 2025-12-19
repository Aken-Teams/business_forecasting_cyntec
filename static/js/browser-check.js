/**
 * 瀏覽器檢測腳本
 * 檢測是否使用 Chrome 瀏覽器並顯示提示彈窗
 */

// 檢測瀏覽器類型
function detectBrowser() {
    const userAgent = navigator.userAgent;
    const vendor = navigator.vendor;

    // 先檢測Edge（因為Edge的User Agent也包含Chrome）
    const isEdge = /Edg/.test(userAgent);

    // 再檢測Chrome（排除Edge）
    const isChrome = /Chrome/.test(userAgent) && /Google Inc/.test(vendor) && !isEdge;

    // 檢測其他瀏覽器
    const isFirefox = /Firefox/.test(userAgent);
    const isSafari = /Safari/.test(userAgent) && !/Chrome/.test(userAgent);
    const isOpera = /Opera/.test(userAgent) || /OPR/.test(userAgent);

    return {
        isChrome: isChrome,
        isEdge: isEdge,
        isFirefox: isFirefox,
        isSafari: isSafari,
        isOpera: isOpera,
        userAgent: userAgent,
        vendor: vendor
    };
}

// 創建並插入 Chrome 阻擋彈窗 HTML
function createChromeBlockModal() {
    // 如果已經存在則不重複創建
    if (document.getElementById('chrome-block-modal')) {
        return;
    }

    const modalHTML = `
    <!-- Chrome 瀏覽器阻擋彈窗 -->
    <div id="chrome-block-modal" class="modal" style="display: none;">
        <div class="modal-content">
            <div class="modal-header">
                <i class="fas fa-exclamation-triangle"></i>
                <h2>瀏覽器不支援</h2>
            </div>
            <div class="modal-body">
                <p>檢測到您正在使用 Google Chrome 瀏覽器。</p>
                <p>由於快取問題，本系統建議使用以下方式以獲得最佳體驗。</p>
                <div class="browser-info">
                    <div class="browser-item">
                        <i class="fab fa-edge"></i>
                        <div class="browser-details">
                            <h3>Microsoft Edge</h3>
                            <p>推薦瀏覽器，完全相容且無快取問題</p>
                        </div>
                    </div>
                    <div class="browser-item">
                        <i class="fab fa-firefox-browser"></i>
                        <div class="browser-details">
                            <h3>Mozilla Firefox</h3>
                            <p>支援瀏覽器，穩定可靠且相容性佳</p>
                        </div>
                    </div>
                </div>
                <div class="modal-actions">
                    <button class="btn btn-primary" onclick="openInEdge()">
                        <i class="fab fa-edge"></i> 用 Edge 開啟
                    </button>
                    <button class="btn btn-firefox" onclick="openInFirefox()">
                        <i class="fab fa-firefox-browser"></i> 用 Firefox 開啟
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Firefox 教學彈窗 -->
    <div id="firefox-tutorial-modal" class="modal" style="display: none;">
        <div class="modal-content">
            <div class="modal-header" style="background: linear-gradient(135deg, #ff9500, #ff7b00);">
                <i class="fab fa-firefox-browser"></i>
                <h2>Firefox 使用教學</h2>
            </div>
            <div class="modal-body">
                <p>請按照以下步驟在 Firefox 瀏覽器中使用本系統：</p>

                <div class="tutorial-steps">
                    <div class="tutorial-step">
                        <div class="step-number">1</div>
                        <div class="step-content">
                            <h4>複製網址</h4>
                            <p>請複製以下網址到剪貼簿：</p>
                            <div class="url-copy-box">
                                <input type="text" id="tutorial-url" value="" readonly>
                                <button onclick="copyUrlToClipboard()" class="copy-btn">
                                    <i class="fas fa-copy"></i> 複製
                                </button>
                            </div>
                        </div>
                    </div>

                    <div class="tutorial-step">
                        <div class="step-number">2</div>
                        <div class="step-content">
                            <h4>開啟 Firefox</h4>
                            <p>在您的電腦上開啟 Mozilla Firefox 瀏覽器</p>
                        </div>
                    </div>

                    <div class="tutorial-step">
                        <div class="step-number">3</div>
                        <div class="step-content">
                            <h4>貼上網址</h4>
                            <p>在 Firefox 的網址列中貼上網址並按 Enter 鍵</p>
                        </div>
                    </div>

                    <div class="tutorial-step">
                        <div class="step-number">4</div>
                        <div class="step-content">
                            <h4>開始使用</h4>
                            <p>現在您可以在 Firefox 中正常使用本系統了！</p>
                        </div>
                    </div>
                </div>

                <div class="modal-actions">
                    <button class="btn btn-primary" onclick="closeFirefoxTutorial()">
                        <i class="fas fa-check"></i> 我知道了
                    </button>
                </div>
            </div>
        </div>
    </div>`;

    // 插入到 body 末尾
    document.body.insertAdjacentHTML('beforeend', modalHTML);

    // 阻止模態框外部點擊關閉
    const chromeModal = document.getElementById('chrome-block-modal');
    if (chromeModal) {
        chromeModal.addEventListener('click', function(e) {
            e.stopPropagation();
        });
    }
}

// 顯示Chrome阻擋彈窗
function showChromeModal() {
    const modal = document.getElementById('chrome-block-modal');
    if (modal) {
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
    }
}

// 用Firefox瀏覽器開啟當前頁面
function openInFirefox() {
    const currentUrl = window.location.href;

    // 設置教學彈窗中的網址
    const tutorialUrl = document.getElementById('tutorial-url');
    if (tutorialUrl) {
        tutorialUrl.value = currentUrl;
    }

    try {
        // 嘗試使用 Firefox 的 URL scheme
        const firefoxUrl = `firefox://open-url?url=${encodeURIComponent(currentUrl)}`;

        // 創建一個隱藏的連結來觸發 Firefox
        const link = document.createElement('a');
        link.href = firefoxUrl;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        // 延遲檢查是否成功跳轉
        setTimeout(() => {
            if (document.visibilityState === 'visible') {
                showFirefoxTutorial();
            }
        }, 2000);

    } catch (error) {
        showFirefoxTutorial();
    }
}

// 顯示Firefox教學彈窗
function showFirefoxTutorial() {
    const modal = document.getElementById('firefox-tutorial-modal');
    if (modal) {
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
    }
}

// 關閉Firefox教學彈窗
function closeFirefoxTutorial() {
    const modal = document.getElementById('firefox-tutorial-modal');
    if (modal) {
        modal.style.display = 'none';
        document.body.style.overflow = 'auto';
    }
}

// 複製網址到剪貼簿
function copyUrlToClipboard() {
    const urlInput = document.getElementById('tutorial-url');
    if (!urlInput) return;

    urlInput.select();
    urlInput.setSelectionRange(0, 99999);

    try {
        document.execCommand('copy');
        const copyBtn = event.target.closest('.copy-btn');
        if (copyBtn) {
            const originalText = copyBtn.innerHTML;
            copyBtn.innerHTML = '<i class="fas fa-check"></i> 已複製';
            copyBtn.style.background = '#28a745';

            setTimeout(() => {
                copyBtn.innerHTML = originalText;
                copyBtn.style.background = '';
            }, 2000);
        }
    } catch (err) {
        alert('無法複製到剪貼簿，請手動複製網址');
    }
}

// 用Edge瀏覽器開啟當前頁面
function openInEdge() {
    const currentUrl = window.location.href;
    const edgeUrl = `microsoft-edge:${currentUrl}`;

    const redirectTime = Date.now();
    window.edgeRedirectTime = redirectTime;

    try {
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        iframe.src = edgeUrl;
        document.body.appendChild(iframe);

        window.location.href = edgeUrl;

        let blurHandled = false;
        const handleBlur = () => {
            if (!blurHandled) {
                blurHandled = true;
                window.removeEventListener('blur', handleBlur);
                return;
            }
        };
        window.addEventListener('blur', handleBlur);

        setTimeout(() => {
            if (Date.now() - redirectTime < 4000 && document.visibilityState === 'visible' && !blurHandled) {
                alert('無法自動開啟Edge瀏覽器，請手動複製網址到Edge瀏覽器中開啟：\n' + currentUrl);
            }
            if (iframe.parentNode) {
                iframe.parentNode.removeChild(iframe);
            }
        }, 3000);

    } catch (error) {
        alert('無法自動開啟Edge瀏覽器，請手動複製網址到Edge瀏覽器中開啟：\n' + currentUrl);
    }
}

// 初始化瀏覽器檢測
function initBrowserCheck() {
    const browser = detectBrowser();
    console.log('瀏覽器檢測結果:', browser);

    // 只要是Chrome瀏覽器就顯示阻擋彈窗
    if (browser.isChrome) {
        console.log('檢測到Chrome瀏覽器，顯示阻擋視窗');
        createChromeBlockModal();
        showChromeModal();
    }
}

// 頁面載入時自動執行檢測
document.addEventListener('DOMContentLoaded', initBrowserCheck);
