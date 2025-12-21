from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, flash, make_response, session
from functools import wraps
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta
import os
import json
import warnings
import hashlib
import time
from dotenv import load_dotenv
warnings.filterwarnings('ignore')

# 載入環境變數
load_dotenv()

# 導入資料庫模組
from database import (
    get_db_connection, init_database, create_default_users,
    verify_user, log_activity, log_upload, log_process, get_user_by_id,
    get_customer_mappings, save_customer_mappings, has_customer_mappings,
    # IT/Admin 管理介面函數
    get_upload_records, get_process_records, get_activity_logs_filtered,
    get_all_customer_mappings, get_users_with_company, get_all_users,
    # 用戶管理函數
    create_user, update_user, delete_user, update_activity_logs_enum,
    # 管理者客戶映射 CRUD 函數
    admin_create_customer_mapping, admin_update_customer_mapping,
    admin_delete_customer_mapping, admin_get_customer_mapping_by_id,
    # 規則管理 CRUD 函數
    get_all_processing_rules, get_processing_rules_by_category,
    get_processing_rule_by_id, update_processing_rule,
    create_processing_rule, delete_processing_rule, toggle_processing_rule_status,
    get_processing_rules_by_user
)

def normalize_date_for_mapping(date_value):
    """
    為 mapping 階段標準化日期格式
    處理各種日期格式：文字格式、pandas Timestamp、datetime 對象等
    """
    try:
        # 如果是空值或 NaN
        if pd.isna(date_value) or date_value is None:
            return None
        
        # 如果已經是 datetime 對象，轉換為標準字串格式
        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value.strftime("%Y/%m/%d")
        
        # 如果是字串，嘗試解析並標準化
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str or date_str.lower() in ['nan', 'none', '']:
                return None
            
            # 嘗試多種日期格式
            date_formats = [
                "%Y/%m/%d",      # 2025/10/01
                "%Y-%m-%d",      # 2025-10-01
                "%m/%d/%Y",      # 10/01/2025
                "%d/%m/%Y",      # 01/10/2025
                "%Y%m%d",        # 20251001
            ]
            
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    return parsed_date.strftime("%Y/%m/%d")
                except ValueError:
                    continue
            
            # 如果所有格式都失敗，使用 pandas 自動解析
            try:
                parsed_date = pd.to_datetime(date_str)
                return parsed_date.strftime("%Y/%m/%d")
            except:
                print(f"    ⚠️ 無法解析日期格式: {date_str}")
                return None
        
        # 其他類型，嘗試用 pandas 轉換
        try:
            parsed_date = pd.to_datetime(date_value)
            return parsed_date.strftime("%Y/%m/%d")
        except:
            print(f"    ⚠️ 無法處理的日期類型: {type(date_value)} - {date_value}")
            return None
            
    except Exception as e:
        print(f"    ❌ 日期標準化失敗: {e}")
        return None

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'default_secret_key')
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)  # Session 有效期 8 小時

# ========================================
# 登入驗證裝飾器
# ========================================

def login_required(f):
    """登入驗證裝飾器"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            # 如果是 API 請求，返回 JSON
            if request.path.startswith('/api/') or request.path.startswith('/upload') or request.path.startswith('/process') or request.path.startswith('/download'):
                return jsonify({'success': False, 'message': '請先登入', 'redirect': '/login'}), 401
            # 否則重定向到登入頁面
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """管理者權限裝飾器"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '請先登入', 'redirect': '/login'}), 401
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '權限不足，僅限管理者'}), 403
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


def it_or_admin_required(f):
    """IT 或管理者權限裝飾器"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '請先登入', 'redirect': '/login'}), 401
            return redirect(url_for('login'))
        if session.get('role') not in ['admin', 'it']:
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '權限不足，僅限 IT 或管理者'}), 403
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def get_current_user():
    """獲取當前登入用戶"""
    if 'user_id' in session:
        return {
            'id': session.get('user_id'),
            'username': session.get('username'),
            'display_name': session.get('display_name'),
            'role': session.get('role'),
            'company': session.get('company')
        }
    return None

def get_client_ip():
    """獲取客戶端 IP"""
    if request.headers.get('X-Forwarded-For'):
        return request.headers.get('X-Forwarded-For').split(',')[0].strip()
    return request.remote_addr

# ========================================
# 文件格式驗證功能
# ========================================

# 範本文件路徑
COMPARE_FOLDER = 'compare'

def get_template_columns(template_type, username=None):
    """
    獲取範本文件的欄位結構
    template_type: 'erp', 'transit', 'forecast'
    username: 用戶名稱，用於指定客戶專屬模板目錄

    範本路徑優先順序：
    1. compare/{username}/{template_type}.xlsx（客戶專屬模板）
    2. compare/{template_type}.xlsx（通用模板，向後兼容）
    """
    try:
        # 優先使用客戶專屬模板
        if username:
            template_file = os.path.join(COMPARE_FOLDER, username, f'{template_type}.xlsx')
            if os.path.exists(template_file):
                print(f"📋 使用客戶專屬模板: {template_file}")
            else:
                # 客戶專屬模板不存在，回退到通用模板
                template_file = os.path.join(COMPARE_FOLDER, f'{template_type}.xlsx')
                print(f"⚠️ 客戶專屬模板不存在，使用通用模板: {template_file}")
        else:
            template_file = os.path.join(COMPARE_FOLDER, f'{template_type}.xlsx')

        if not os.path.exists(template_file):
            # 提供更清楚的錯誤訊息
            if username:
                expected_path = f'compare\\{username}\\{template_type}.xlsx'
            else:
                expected_path = f'compare\\{template_type}.xlsx'
            return None, f'範本文件不存在: {expected_path}'

        if template_type == 'forecast':
            # Forecast 使用 header=None 讀取原始結構
            df = pd.read_excel(template_file, nrows=20, header=None)
            return df, None
        else:
            # ERP 和 Transit 讀取欄位名稱
            df = pd.read_excel(template_file, nrows=1)
            return list(df.columns), None
    except Exception as e:
        return None, f'讀取範本文件失敗: {str(e)}'

def validate_erp_format(uploaded_file_path, username=None):
    """
    驗證 ERP 文件格式
    必須欄位名稱和順序完全一致
    username: 用戶名稱，用於指定客戶專屬模板目錄
    """
    try:
        # 獲取範本欄位（根據用戶名稱取得對應模板）
        template_columns, error = get_template_columns('erp', username)
        if error:
            return False, error, []

        # 讀取上傳的文件
        uploaded_df = pd.read_excel(uploaded_file_path, nrows=1)
        uploaded_columns = list(uploaded_df.columns)

        # 檢查欄位數量
        if len(uploaded_columns) != len(template_columns):
            return False, f'欄位數量不符：預期 {len(template_columns)} 個欄位，實際 {len(uploaded_columns)} 個欄位', []

        # 檢查欄位名稱和順序
        mismatched = []
        for i, (template_col, uploaded_col) in enumerate(zip(template_columns, uploaded_columns)):
            # 標準化比較（去除空白和換行符）
            template_clean = str(template_col).strip().replace('\n', '')
            uploaded_clean = str(uploaded_col).strip().replace('\n', '')
            if template_clean != uploaded_clean:
                mismatched.append({
                    'index': i,
                    'expected': template_clean,
                    'actual': uploaded_clean
                })

        if mismatched:
            error_details = []
            for m in mismatched[:5]:  # 最多顯示5個錯誤
                error_details.append(f"欄位 {m['index']+1}: 預期「{m['expected']}」，實際「{m['actual']}」")

            if len(mismatched) > 5:
                error_details.append(f"...還有 {len(mismatched)-5} 個欄位不符")

            return False, '欄位名稱或順序不符', error_details

        return True, 'ERP 文件格式驗證通過', []

    except Exception as e:
        return False, f'驗證過程發生錯誤: {str(e)}', []

def validate_transit_format(uploaded_file_path, username=None):
    """
    驗證在途文件格式
    必須欄位名稱和順序完全一致
    username: 用戶名稱，用於指定客戶專屬模板目錄
    """
    try:
        # 獲取範本欄位（根據用戶名稱取得對應模板）
        template_columns, error = get_template_columns('transit', username)
        if error:
            return False, error, []

        # 讀取上傳的文件
        uploaded_df = pd.read_excel(uploaded_file_path, nrows=1)
        uploaded_columns = list(uploaded_df.columns)

        # 檢查欄位數量
        if len(uploaded_columns) != len(template_columns):
            return False, f'欄位數量不符：預期 {len(template_columns)} 個欄位，實際 {len(uploaded_columns)} 個欄位', []

        # 檢查欄位名稱和順序
        mismatched = []
        for i, (template_col, uploaded_col) in enumerate(zip(template_columns, uploaded_columns)):
            # 標準化比較（去除空白和換行符）
            template_clean = str(template_col).strip().replace('\n', '')
            uploaded_clean = str(uploaded_col).strip().replace('\n', '')
            if template_clean != uploaded_clean:
                mismatched.append({
                    'index': i,
                    'expected': template_clean,
                    'actual': uploaded_clean
                })

        if mismatched:
            error_details = []
            for m in mismatched[:5]:  # 最多顯示5個錯誤
                error_details.append(f"欄位 {m['index']+1}: 預期「{m['expected']}」，實際「{m['actual']}」")

            if len(mismatched) > 5:
                error_details.append(f"...還有 {len(mismatched)-5} 個欄位不符")

            return False, '欄位名稱或順序不符', error_details

        return True, '在途文件格式驗證通過', []

    except Exception as e:
        return False, f'驗證過程發生錯誤: {str(e)}', []

def validate_forecast_format(uploaded_file_path, username=None):
    """
    驗證 Forecast 文件格式
    檢查整體結構是否與範本一致（不比對具體數據和日期）
    username: 用戶名稱，用於指定客戶專屬模板目錄
    """
    try:
        # 獲取範本結構（根據用戶名稱取得對應模板）
        template_df, error = get_template_columns('forecast', username)
        if error:
            return False, error, []

        # 讀取上傳的文件（不使用 header）
        uploaded_df = pd.read_excel(uploaded_file_path, nrows=20, header=None)

        # 1. 檢查欄位數量是否一致
        if len(uploaded_df.columns) != len(template_df.columns):
            return False, f'欄位數量不符：預期 {len(template_df.columns)} 個欄位，實際 {len(uploaded_df.columns)} 個欄位', []

        # 2. 檢查固定位置的標識符
        structure_checks = []

        # 檢查 A1 是否為 "筆數"
        a1_template = str(template_df.iloc[0, 0]).strip() if pd.notna(template_df.iloc[0, 0]) else ''
        a1_uploaded = str(uploaded_df.iloc[0, 0]).strip() if pd.notna(uploaded_df.iloc[0, 0]) else ''
        if a1_template != a1_uploaded:
            structure_checks.append(f"A1 欄位: 預期「{a1_template}」，實際「{a1_uploaded}」")

        # 檢查 A2 是否為 "需求週數"
        a2_template = str(template_df.iloc[1, 0]).strip() if pd.notna(template_df.iloc[1, 0]) else ''
        a2_uploaded = str(uploaded_df.iloc[1, 0]).strip() if pd.notna(uploaded_df.iloc[1, 0]) else ''
        if a2_template != a2_uploaded:
            structure_checks.append(f"A2 欄位: 預期「{a2_template}」，實際「{a2_uploaded}」")

        # 檢查 A3 是否為 "客戶名稱"
        a3_template = str(template_df.iloc[2, 0]).strip() if pd.notna(template_df.iloc[2, 0]) else ''
        a3_uploaded = str(uploaded_df.iloc[2, 0]).strip() if pd.notna(uploaded_df.iloc[2, 0]) else ''
        if a3_template != a3_uploaded:
            structure_checks.append(f"A3 欄位: 預期「{a3_template}」，實際「{a3_uploaded}」")

        # 檢查 A4 是否為 "客戶料號"
        a4_template = str(template_df.iloc[3, 0]).strip() if pd.notna(template_df.iloc[3, 0]) else ''
        a4_uploaded = str(uploaded_df.iloc[3, 0]).strip() if pd.notna(uploaded_df.iloc[3, 0]) else ''
        if a4_template != a4_uploaded:
            structure_checks.append(f"A4 欄位: 預期「{a4_template}」，實際「{a4_uploaded}」")

        # 檢查 C1 是否為 "Web"
        c1_template = str(template_df.iloc[0, 2]).strip() if pd.notna(template_df.iloc[0, 2]) else ''
        c1_uploaded = str(uploaded_df.iloc[0, 2]).strip() if pd.notna(uploaded_df.iloc[0, 2]) else ''
        if c1_template != c1_uploaded:
            structure_checks.append(f"C1 欄位: 預期「{c1_template}」，實際「{c1_uploaded}」")

        # 檢查第4行的標題結構（B4, C4, D4）
        b4_template = str(template_df.iloc[3, 1]).strip() if pd.notna(template_df.iloc[3, 1]) else ''
        b4_uploaded = str(uploaded_df.iloc[3, 1]).strip() if pd.notna(uploaded_df.iloc[3, 1]) else ''
        if b4_template != b4_uploaded:
            structure_checks.append(f"B4 欄位: 預期「{b4_template}」，實際「{b4_uploaded}」")

        c4_template = str(template_df.iloc[3, 2]).strip() if pd.notna(template_df.iloc[3, 2]) else ''
        c4_uploaded = str(uploaded_df.iloc[3, 2]).strip() if pd.notna(uploaded_df.iloc[3, 2]) else ''
        if c4_template != c4_uploaded:
            structure_checks.append(f"C4 欄位: 預期「{c4_template}」，實際「{c4_uploaded}」")

        d4_template = str(template_df.iloc[3, 3]).strip() if pd.notna(template_df.iloc[3, 3]) else ''
        d4_uploaded = str(uploaded_df.iloc[3, 3]).strip() if pd.notna(uploaded_df.iloc[3, 3]) else ''
        if d4_template != d4_uploaded:
            structure_checks.append(f"D4 欄位: 預期「{d4_template}」，實際「{d4_uploaded}」")

        if structure_checks:
            return False, 'Forecast 文件結構不符', structure_checks

        return True, 'Forecast 文件格式驗證通過', []

    except Exception as e:
        return False, f'驗證過程發生錯誤: {str(e)}', []

# ========================================

# 生成版本號用於防止快取
def get_version():
    """生成基於時間戳的版本號"""
    return str(int(time.time()))

# 添加版本控制到模板上下文
@app.context_processor
def inject_version():
    return dict(version=get_version())

# 全局變量存儲上傳的文件
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
FILE_RETENTION_DAYS = 30  # 檔案保留天數

# 確保文件夾存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True)
os.makedirs('static/css', exist_ok=True)
os.makedirs('static/js', exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_or_create_session_folder(user_id, folder_type='uploads'):
    """
    獲取或建立用戶的 session 資料夾
    資料夾結構: {folder_type}/{user_id}/{YYYYMMDD_HHMMSS}/

    所有上傳和處理的檔案都使用同一個 session 時間戳，
    確保同一次操作的檔案都在同一個資料夾中。
    """
    # 使用統一的 session 時間戳 key（不區分 uploads/processed）
    session_key = 'current_session_timestamp'
    session_timestamp = session.get(session_key)

    # 如果沒有 session 時間戳，建立新的
    if not session_timestamp:
        session_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        session[session_key] = session_timestamp
        print(f"📁 建立新的 session 時間戳: {session_timestamp}")
    else:
        print(f"📁 使用現有 session 時間戳: {session_timestamp}")

    # 建立資料夾路徑
    if folder_type == 'uploads':
        folder_path = os.path.join(UPLOAD_FOLDER, str(user_id), session_timestamp)
    else:
        folder_path = os.path.join(PROCESSED_FOLDER, str(user_id), session_timestamp)

    # 確保資料夾存在
    os.makedirs(folder_path, exist_ok=True)

    return folder_path, session_timestamp

def reset_session_folder():
    """
    重置 session 資料夾時間戳
    當用戶重新開始流程時調用此函數
    """
    if 'current_session_timestamp' in session:
        old_timestamp = session.pop('current_session_timestamp')
        print(f"🔄 已重置 session 時間戳 (舊: {old_timestamp})")
    # 同時清除檔案路徑
    for key in ['current_erp_file', 'current_forecast_file', 'current_transit_file', 'current_processed_folder']:
        if key in session:
            session.pop(key)

def get_session_folder_path(user_id, folder_type='uploads'):
    """
    獲取當前 session 的資料夾路徑（不建立新的）
    """
    session_key = 'current_session_timestamp'
    session_timestamp = session.get(session_key)

    if not session_timestamp:
        return None, None

    if folder_type == 'uploads':
        folder_path = os.path.join(UPLOAD_FOLDER, str(user_id), session_timestamp)
    else:
        folder_path = os.path.join(PROCESSED_FOLDER, str(user_id), session_timestamp)

    return folder_path, session_timestamp

def cleanup_old_folders():
    """清理超過保留期限的資料夾"""
    import shutil
    try:
        cutoff_date = datetime.now() - timedelta(days=FILE_RETENTION_DAYS)
        cleaned_count = 0

        for base_folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
            if not os.path.exists(base_folder):
                continue

            # 遍歷用戶資料夾
            for user_folder in os.listdir(base_folder):
                user_folder_path = os.path.join(base_folder, user_folder)
                if not os.path.isdir(user_folder_path):
                    # 處理舊格式的檔案（直接在 uploads/processed 下的檔案）
                    if os.path.isfile(user_folder_path):
                        file_mtime = datetime.fromtimestamp(os.path.getmtime(user_folder_path))
                        if file_mtime < cutoff_date:
                            os.remove(user_folder_path)
                            cleaned_count += 1
                            print(f"🗑️ 已清理過期檔案: {user_folder}")
                    continue

                # 遍歷 session 時間戳資料夾
                for session_folder in os.listdir(user_folder_path):
                    session_folder_path = os.path.join(user_folder_path, session_folder)
                    if not os.path.isdir(session_folder_path):
                        continue

                    # 嘗試解析資料夾名稱中的時間戳
                    try:
                        folder_date = datetime.strptime(session_folder, '%Y%m%d_%H%M%S')
                        if folder_date < cutoff_date:
                            shutil.rmtree(session_folder_path)
                            cleaned_count += 1
                            print(f"🗑️ 已清理過期資料夾: {base_folder}/{user_folder}/{session_folder}")
                    except ValueError:
                        # 無法解析時間戳，使用資料夾修改時間
                        folder_mtime = datetime.fromtimestamp(os.path.getmtime(session_folder_path))
                        if folder_mtime < cutoff_date:
                            shutil.rmtree(session_folder_path)
                            cleaned_count += 1
                            print(f"🗑️ 已清理過期資料夾: {base_folder}/{user_folder}/{session_folder}")

                # 如果用戶資料夾為空，也刪除
                if os.path.exists(user_folder_path) and not os.listdir(user_folder_path):
                    os.rmdir(user_folder_path)
                    print(f"🗑️ 已清理空用戶資料夾: {base_folder}/{user_folder}")

        if cleaned_count > 0:
            print(f"✅ 共清理 {cleaned_count} 個過期資料夾/檔案")
        return cleaned_count
    except Exception as e:
        print(f"❌ 清理過期資料夾失敗: {e}")
        return 0

def get_user_file_path(file_type):
    """從 session 獲取當前用戶的檔案路徑"""
    key = f'current_{file_type}_file'
    return session.get(key)

def set_user_file_path(file_type, filepath):
    """設置當前用戶的檔案路徑到 session"""
    key = f'current_{file_type}_file'
    session[key] = filepath

def get_user_processed_folder():
    """獲取當前用戶的 processed 資料夾路徑"""
    user = get_current_user()
    if not user:
        return PROCESSED_FOLDER

    folder_path, _ = get_session_folder_path(user['id'], 'processed')
    if folder_path and os.path.exists(folder_path):
        return folder_path

    # 如果還沒有 session 資料夾，建立一個
    folder_path, _ = get_or_create_session_folder(user['id'], 'processed')
    return folder_path

# ========================================
# 登入/登出路由
# ========================================

@app.route('/login')
def login():
    """登入頁面"""
    # 如果已登入，跳轉到首頁
    if 'user_id' in session:
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/api/login', methods=['POST'])
def api_login():
    """登入 API"""
    try:
        data = request.get_json()
        username = data.get('username', '').strip()
        password = data.get('password', '')

        if not username or not password:
            return jsonify({'success': False, 'message': '請輸入帳號和密碼'})

        # 驗證用戶
        user = verify_user(username, password)

        if user:
            # 設置 Session
            session.permanent = True
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['display_name'] = user['display_name']
            session['role'] = user['role']
            session['company'] = user['company']

            # 記錄登入活動
            log_activity(
                user_id=user['id'],
                username=user['username'],
                action_type='login',
                action_detail=f"用戶 {user['display_name']} 登入成功",
                ip_address=get_client_ip(),
                user_agent=request.headers.get('User-Agent')
            )

            print(f"✅ 用戶登入成功: {user['username']} ({user['display_name']})")

            # 根據角色決定跳轉頁面
            redirect_url = '/'
            if user['role'] == 'admin':
                redirect_url = '/admin'
            elif user['role'] == 'it':
                redirect_url = '/it'

            return jsonify({
                'success': True,
                'message': '登入成功',
                'redirect_url': redirect_url,
                'user': {
                    'username': user['username'],
                    'display_name': user['display_name'],
                    'role': user['role'],
                    'company': user['company']
                }
            })
        else:
            # 記錄登入失敗
            log_activity(
                user_id=None,
                username=username,
                action_type='login_failed',
                action_detail=f"登入失敗：帳號 {username}",
                ip_address=get_client_ip(),
                user_agent=request.headers.get('User-Agent')
            )

            print(f"❌ 登入失敗: {username}")
            return jsonify({'success': False, 'message': '帳號或密碼錯誤'})

    except Exception as e:
        print(f"❌ 登入處理錯誤: {e}")
        return jsonify({'success': False, 'message': '系統錯誤，請稍後再試'})

@app.route('/api/logout', methods=['POST'])
@login_required
def api_logout():
    """登出 API"""
    try:
        user = get_current_user()
        if user:
            # 記錄登出活動
            log_activity(
                user_id=user['id'],
                username=user['username'],
                action_type='logout',
                action_detail=f"用戶 {user['display_name']} 登出",
                ip_address=get_client_ip(),
                user_agent=request.headers.get('User-Agent')
            )
            print(f"✅ 用戶登出: {user['username']}")

        # 清除 Session
        session.clear()
        return jsonify({'success': True, 'message': '已登出'})

    except Exception as e:
        print(f"❌ 登出處理錯誤: {e}")
        session.clear()
        return jsonify({'success': True, 'message': '已登出'})

@app.route('/logout')
def logout():
    """登出頁面"""
    user = get_current_user()
    if user:
        log_activity(
            user_id=user['id'],
            username=user['username'],
            action_type='logout',
            action_detail=f"用戶 {user['display_name']} 登出",
            ip_address=get_client_ip(),
            user_agent=request.headers.get('User-Agent')
        )
    session.clear()
    return redirect(url_for('login'))

@app.route('/api/user')
@login_required
def api_get_user():
    """獲取當前用戶資訊"""
    user = get_current_user()
    if user:
        return jsonify({'success': True, 'user': user})
    return jsonify({'success': False, 'message': '未登入'})

@app.route('/api/reset_session', methods=['POST'])
@login_required
def api_reset_session():
    """重置當前的 session 資料夾時間戳（當用戶重新開始流程時調用）"""
    try:
        reset_session_folder()
        return jsonify({'success': True, 'message': 'Session 已重置'})
    except Exception as e:
        print(f"❌ 重置 session 失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/migrate_mapping', methods=['POST'])
@login_required
def api_migrate_mapping():
    """
    將現有的 mapping_data.json 遷移到資料庫
    僅限管理員或 IT 人員使用
    """
    user = get_current_user()

    # 檢查權限（僅管理員和 IT 可執行）
    if user['role'] not in ['admin', 'it']:
        return jsonify({'success': False, 'message': '權限不足，僅管理員和 IT 人員可執行此操作'})

    try:
        mapping_file = os.path.join('mapping', 'mapping_data.json')

        if not os.path.exists(mapping_file):
            return jsonify({'success': False, 'message': 'mapping_data.json 檔案不存在'})

        # 讀取 JSON 檔案
        with open(mapping_file, 'r', encoding='utf-8') as f:
            mapping_data = json.load(f)

        # 儲存到當前用戶的資料庫
        if save_customer_mappings(user['id'], mapping_data):
            print(f"✅ 已將 mapping_data.json 遷移至資料庫 (user: {user['username']})")

            # 記錄活動
            log_activity(user['id'], user['username'], 'mapping_start',
                       f"遷移 mapping_data.json 至資料庫", get_client_ip(), request.headers.get('User-Agent'))

            return jsonify({
                'success': True,
                'message': f'成功將 mapping 資料遷移至資料庫（用戶：{user["display_name"]}）',
                'customers_count': len(set(
                    list(mapping_data.get('regions', {}).keys()) +
                    list(mapping_data.get('schedule_breakpoints', {}).keys()) +
                    list(mapping_data.get('etd', {}).keys()) +
                    list(mapping_data.get('eta', {}).keys())
                ))
            })
        else:
            return jsonify({'success': False, 'message': '儲存至資料庫失敗'})

    except Exception as e:
        print(f"❌ 遷移 mapping 失敗: {e}")
        return jsonify({'success': False, 'message': f'遷移失敗: {str(e)}'})

# ========================================
# 主要頁面路由
# ========================================

@app.route('/')
@login_required
def index():
    user = get_current_user()
    response = make_response(render_template('index.html', user=user))
    # 設置不緩存HTML頁面
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/upload_erp', methods=['POST'])
@login_required
def upload_erp():
    user = get_current_user()
    original_filename = ''
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        original_filename = file.filename

        if file and allowed_file(file.filename):
            # 使用資料夾管理結構：uploads/{user_id}/{session_timestamp}/erp_data.xlsx
            upload_folder, session_timestamp = get_or_create_session_folder(user['id'], 'uploads')
            filename = 'erp_data.xlsx'
            filepath = os.path.join(upload_folder, filename)

            # 清理超過保留期限的資料夾
            cleanup_old_folders()

            # 保存文件
            file.save(filepath)
            print(f"ERP文件已保存到: {filepath} (session: {session_timestamp})")

            # 檢查文件是否真的存在
            if not os.path.exists(filepath):
                return jsonify({'success': False, 'message': '文件保存失敗'})

            # ========== 格式驗證 ==========
            # 檢查是否為測試模式，如果是則使用被測試客戶的 username
            test_mode = request.form.get('test_mode') == 'true'
            customer_id = request.form.get('customer_id')
            template_username = user['username']

            if test_mode and customer_id and user['role'] in ['admin', 'it']:
                # 測試模式：取得被測試客戶的 username
                test_customer = get_user_by_id(int(customer_id))
                if test_customer:
                    template_username = test_customer['username']
                    print(f"[測試模式] 使用客戶 {template_username} 的模板進行驗證")

            print(f"開始驗證 ERP 文件格式...（用戶: {template_username}）")
            is_valid, message, details = validate_erp_format(filepath, template_username)

            if not is_valid:
                # 驗證失敗，刪除已上傳的文件
                os.remove(filepath)
                print(f"❌ ERP 文件格式驗證失敗: {message}")
                if details:
                    for detail in details:
                        print(f"   - {detail}")

                # 記錄上傳失敗
                log_upload(user['id'], 'erp', original_filename, 0, 0, 0, 'validation_failed', message)
                log_activity(user['id'], user['username'], 'upload_erp_failed',
                           f"ERP 文件上傳失敗：{message}", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': False,
                    'message': f'ERP 文件格式驗證失敗：{message}',
                    'details': details,
                    'validation_error': True
                })

            print(f"✅ ERP 文件格式驗證通過")
            # ========== 格式驗證結束 ==========

            # 讀取文件並返回基本信息
            try:
                df = pd.read_excel(filepath)
                file_size = os.path.getsize(filepath)
                print(f"ERP文件讀取成功: {len(df)} 行, {len(df.columns)} 欄位")

                # 將檔案路徑存入 session，供後續處理使用
                set_user_file_path('erp', filepath)

                # 記錄上傳成功
                log_upload(user['id'], 'erp', original_filename, file_size, len(df), len(df.columns), 'success')
                log_activity(user['id'], user['username'], 'upload_erp',
                           f"ERP 文件上傳成功：{original_filename} -> {filename}", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': True,
                    'message': 'ERP文件上傳成功（格式驗證通過）',
                    'rows': len(df),
                    'columns': list(df.columns),
                    'file_size': file_size,
                    'saved_filename': filename
                })
            except Exception as e:
                print(f"ERP文件讀取失敗: {str(e)}")
                log_upload(user['id'], 'erp', original_filename, 0, 0, 0, 'failed', str(e))
                return jsonify({'success': False, 'message': f'文件讀取失敗: {str(e)}'})

        return jsonify({'success': False, 'message': '不支持的文件格式'})

    except Exception as e:
        print(f"ERP上傳處理錯誤: {str(e)}")
        if user:
            log_upload(user['id'], 'erp', original_filename, 0, 0, 0, 'failed', str(e))
        return jsonify({'success': False, 'message': f'上傳處理失敗: {str(e)}'})

@app.route('/upload_forecast', methods=['POST'])
@login_required
def upload_forecast():
    """
    上傳 Forecast 文件（支援多檔案上傳）
    - 多檔案：使用 'files' 欄位
    - 單檔案：使用 'file' 欄位（向後兼容）
    - 多個 Forecast 檔案會合併成一個 forecast_data.xlsx
    """
    user = get_current_user()
    original_filenames = []
    try:
        # 檢查是否有上傳檔案（支援多檔案 'files' 或單檔案 'file'）
        files_list = []
        if 'files' in request.files:
            files_list = request.files.getlist('files')
        elif 'file' in request.files:
            files_list = [request.files['file']]

        if not files_list or all(f.filename == '' for f in files_list):
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        # 過濾掉空檔名
        files_list = [f for f in files_list if f.filename != '']
        if not files_list:
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        # 檢查所有文件格式
        for file in files_list:
            if not allowed_file(file.filename):
                return jsonify({'success': False, 'message': f'不支持的文件格式: {file.filename}'})
            original_filenames.append(file.filename)

        # 使用資料夾管理結構
        upload_folder, session_timestamp = get_or_create_session_folder(user['id'], 'uploads')

        # 取得測試模式參數
        test_mode = request.form.get('test_mode') == 'true'
        customer_id = request.form.get('customer_id')
        template_username = user['username']

        if test_mode and customer_id and user['role'] in ['admin', 'it']:
            test_customer = get_user_by_id(int(customer_id))
            if test_customer:
                template_username = test_customer['username']
                print(f"[測試模式] 使用客戶 {template_username} 的模板進行驗證")

        # ========== 處理多檔案上傳 ==========
        if len(files_list) == 1:
            # 單檔案上傳：維持原有邏輯
            file = files_list[0]
            original_filename = file.filename
            filename = 'forecast_data.xlsx'
            filepath = os.path.join(upload_folder, filename)

            file.save(filepath)
            print(f"Forecast文件已保存到: {filepath} (session: {session_timestamp})")

            if not os.path.exists(filepath):
                return jsonify({'success': False, 'message': '文件保存失敗'})

            # 格式驗證
            print(f"開始驗證 Forecast 文件格式...（用戶: {template_username}）")
            is_valid, message, details = validate_forecast_format(filepath, template_username)

            if not is_valid:
                os.remove(filepath)
                print(f"❌ Forecast 文件格式驗證失敗: {message}")
                log_upload(user['id'], 'forecast', original_filename, 0, 0, 0, 'validation_failed', message)
                log_activity(user['id'], user['username'], 'upload_forecast_failed',
                           f"Forecast 文件上傳失敗：{message}", get_client_ip(), request.headers.get('User-Agent'))
                return jsonify({
                    'success': False,
                    'message': f'Forecast 文件格式驗證失敗：{message}',
                    'details': details,
                    'validation_error': True
                })

            print(f"✅ Forecast 文件格式驗證通過")

            # 讀取並返回
            df = pd.read_excel(filepath)
            file_size = os.path.getsize(filepath)
            set_user_file_path('forecast', filepath)

            log_upload(user['id'], 'forecast', original_filename, file_size, len(df), len(df.columns), 'success')
            log_activity(user['id'], user['username'], 'upload_forecast',
                       f"Forecast 文件上傳成功：{original_filename}", get_client_ip(), request.headers.get('User-Agent'))

            return jsonify({
                'success': True,
                'message': 'Forecast文件上傳成功（格式驗證通過）',
                'rows': len(df),
                'columns': list(df.columns),
                'file_size': file_size,
                'file_count': 1,
                'saved_filename': filename
            })

        else:
            # 多檔案上傳模式
            # 取得合併選項（預設為合併）
            merge_files = request.form.get('merge_files', 'true') == 'true'
            print(f"=== 多檔案上傳模式：收到 {len(files_list)} 個 Forecast 文件，合併模式: {merge_files} ===")

            all_dataframes = []
            files_info = []
            total_size = 0
            validation_errors = []

            # 先儲存所有檔案到暫存位置
            temp_files = []
            for idx, file in enumerate(files_list):
                temp_filename = f'forecast_temp_{idx}.xlsx'
                temp_filepath = os.path.join(upload_folder, temp_filename)
                file.save(temp_filepath)
                temp_files.append((file.filename, temp_filepath))
                print(f"  暫存文件 {idx + 1}: {file.filename} -> {temp_filepath}")

            # 驗證並讀取每個檔案
            for original_name, temp_path in temp_files:
                print(f"驗證 Forecast 文件: {original_name}（用戶: {template_username}）")
                is_valid, message, details = validate_forecast_format(temp_path, template_username)

                if not is_valid:
                    validation_errors.append({
                        'filename': original_name,
                        'message': message,
                        'details': details
                    })
                    continue

                try:
                    df = pd.read_excel(temp_path)
                    file_size = os.path.getsize(temp_path)
                    total_size += file_size

                    all_dataframes.append(df)
                    files_info.append({
                        'name': original_name,
                        'rows': len(df),
                        'columns': len(df.columns),
                        'size': file_size
                    })
                    print(f"  ✅ {original_name}: {len(df)} 行, {len(df.columns)} 欄位")

                except Exception as e:
                    validation_errors.append({
                        'filename': original_name,
                        'message': f'讀取失敗: {str(e)}',
                        'details': []
                    })

            # 如果有驗證錯誤
            if validation_errors:
                # 清理暫存檔案
                for _, temp_path in temp_files:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

                error_details = []
                for err in validation_errors:
                    error_details.append(f"{err['filename']}: {err['message']}")
                    if err['details']:
                        error_details.extend([f"  - {d}" for d in err['details']])

                log_activity(user['id'], user['username'], 'upload_forecast_failed',
                           f"Forecast 多檔案上傳失敗：{len(validation_errors)} 個文件驗證失敗", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': False,
                    'message': f'{len(validation_errors)} 個文件格式驗證失敗',
                    'details': error_details,
                    'validation_error': True
                })

            # 檢查是否有有效檔案
            if not all_dataframes:
                return jsonify({'success': False, 'message': '沒有有效的 Forecast 文件'})

            # ========== 根據合併選項處理 ==========
            if merge_files:
                # 合併模式：使用 xlwings 操作 Excel 原生複製貼上，保留格式且速度快
                import shutil
                import time

                print(f"開始合併 {len(temp_files)} 個 Forecast 檔案（使用 Excel 原生複製貼上保留格式）...")
                merge_start = time.time()

                # ===== 步驟 1：直接複製第一個檔案作為基礎 =====
                first_temp_path = temp_files[0][1]
                final_filename = 'forecast_data.xlsx'
                final_filepath = os.path.join(upload_folder, final_filename)
                shutil.copy2(first_temp_path, final_filepath)
                print(f"  複製第一個檔案完成，耗時 {time.time() - merge_start:.2f} 秒")

                # ===== 步驟 2：使用 xlwings 進行真正的 Excel 複製貼上 =====
                if len(temp_files) > 1:
                    import xlwings as xw

                    # 啟動 Excel（隱藏模式）
                    app = xw.App(visible=False, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = False

                    try:
                        # 打開目標檔案
                        dest_wb = app.books.open(final_filepath)
                        dest_ws = dest_wb.sheets[0]

                        # 嘗試取消目標工作表保護
                        try:
                            if dest_ws.api.ProtectContents:
                                dest_ws.api.Unprotect()
                                print(f"  已取消目標工作表保護")
                        except Exception as unprotect_err:
                            # 如果有密碼保護，無法取消
                            app.quit()
                            # 清理暫存檔案
                            for _, temp_path in temp_files:
                                if os.path.exists(temp_path):
                                    os.remove(temp_path)
                            return jsonify({
                                'success': False,
                                'message': 'Forecast 檔案的工作表有密碼保護，無法進行合併',
                                'details': '請先手動在 Excel 中取消工作表保護（檢閱 → 取消保護工作表），然後再重新上傳。'
                            })

                        # 取得第一個檔案的欄數（用於複製範圍）
                        first_max_col = dest_ws.used_range.last_cell.column

                        for file_idx in range(1, len(temp_files)):
                            src_path = temp_files[file_idx][1]

                            # 打開來源檔案
                            src_wb = app.books.open(src_path)
                            src_ws = src_wb.sheets[0]

                            # 嘗試取消來源工作表保護
                            try:
                                if src_ws.api.ProtectContents:
                                    src_ws.api.Unprotect()
                                    print(f"  已取消來源檔案 {file_idx + 1} 工作表保護")
                            except Exception as src_unprotect_err:
                                src_wb.close()
                                app.quit()
                                # 清理暫存檔案
                                for _, temp_path in temp_files:
                                    if os.path.exists(temp_path):
                                        os.remove(temp_path)
                                return jsonify({
                                    'success': False,
                                    'message': f'第 {file_idx + 1} 個 Forecast 檔案的工作表有密碼保護，無法進行合併',
                                    'details': '請先手動在 Excel 中取消工作表保護（檢閱 → 取消保護工作表），然後再重新上傳。'
                                })

                            # 取得來源資料範圍（跳過標題，從第2行開始）
                            src_last_row = src_ws.used_range.last_cell.row
                            src_last_col = src_ws.used_range.last_cell.column

                            if src_last_row < 2:
                                # 只有標題，沒有資料
                                src_wb.close()
                                print(f"  檔案 {file_idx + 1} 沒有資料行，跳過")
                                continue

                            # 複製來源資料區域（從第2行到最後一行）
                            src_range = src_ws.range(f'A2:{xw.utils.col_name(src_last_col)}{src_last_row}')

                            # 找到目標的下一個空白行
                            dest_last_row = dest_ws.used_range.last_cell.row
                            dest_start_row = dest_last_row + 1

                            # 複製並貼上（保留格式）
                            src_range.copy()
                            dest_cell = dest_ws.range(f'A{dest_start_row}')
                            dest_cell.paste(paste='all')

                            # 清除剪貼簿
                            app.api.CutCopyMode = False

                            rows_copied = src_last_row - 1  # 減去標題行
                            print(f"  合併檔案 {file_idx + 1} 完成（{rows_copied} 行），耗時 {time.time() - merge_start:.2f} 秒")

                            # 關閉來源檔案（不儲存）
                            src_wb.close()

                        # 儲存並關閉目標檔案
                        dest_wb.save()
                        dest_wb.close()

                    finally:
                        # 確保 Excel 應用程式關閉
                        app.quit()

                # 計算總行數
                total_rows = sum(df.shape[0] for df in all_dataframes)
                print(f"=== 多檔案合併完成：{final_filepath}，總行數：{total_rows}，總耗時 {time.time() - merge_start:.2f} 秒 ===")

                # 清理暫存檔案
                for _, temp_path in temp_files:
                    if os.path.exists(temp_path) and temp_path != final_filepath:
                        os.remove(temp_path)

                # 取得合併後的檔案大小
                merged_size = os.path.getsize(final_filepath)

                # 儲存路徑到 session（單一合併檔案）
                set_user_file_path('forecast', final_filepath)
                set_user_file_path('forecast_merge_mode', True)
                set_user_file_path('forecast_files', None)

                # 記錄日誌
                filenames_str = ', '.join(original_filenames)
                # 計算總欄數（從第一個 dataframe 取得）
                total_columns = len(all_dataframes[0].columns) if all_dataframes else 0
                log_upload(user['id'], 'forecast', filenames_str, merged_size, total_rows, total_columns, 'success')
                log_activity(user['id'], user['username'], 'upload_forecast',
                           f"Forecast 多檔案上傳成功：{len(files_list)} 個文件已合併", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': True,
                    'message': f'{len(files_list)} 個 Forecast 文件上傳並合併成功',
                    'file_count': len(files_list),
                    'total_rows': total_rows,
                    'total_size': merged_size,
                    'files': files_info,
                    'merge_mode': True,
                    'saved_filename': final_filename
                })

            else:
                # 不合併模式：將暫存檔案重新命名為正式檔案
                saved_files = []
                total_rows = 0

                for idx, (original_name, temp_path) in enumerate(temp_files):
                    # 產生正式檔名：forecast_data_1.xlsx, forecast_data_2.xlsx, ...
                    final_filename = f'forecast_data_{idx + 1}.xlsx'
                    final_filepath = os.path.join(upload_folder, final_filename)

                    # 移動暫存檔案到正式位置
                    import shutil
                    shutil.move(temp_path, final_filepath)

                    saved_files.append({
                        'original_name': original_name,
                        'saved_name': final_filename,
                        'path': final_filepath,
                        'rows': files_info[idx]['rows']
                    })
                    total_rows += files_info[idx]['rows']

                    print(f"  儲存文件 {idx + 1}: {original_name} -> {final_filename}")

                print(f"=== 多檔案分開儲存完成：{len(saved_files)} 個文件 ===")

                # 儲存多檔案資訊到 session
                set_user_file_path('forecast', None)  # 不設定單一檔案路徑
                set_user_file_path('forecast_merge_mode', False)
                set_user_file_path('forecast_files', saved_files)

                # 記錄日誌
                filenames_str = ', '.join(original_filenames)
                log_upload(user['id'], 'forecast', filenames_str, total_size, total_rows, files_info[0]['columns'] if files_info else 0, 'success')
                log_activity(user['id'], user['username'], 'upload_forecast',
                           f"Forecast 多檔案上傳成功：{len(files_list)} 個文件（不合併）", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': True,
                    'message': f'{len(files_list)} 個 Forecast 文件上傳成功（不合併）',
                    'file_count': len(files_list),
                    'total_rows': total_rows,
                    'total_size': total_size,
                    'files': files_info,
                    'merge_mode': False,
                    'saved_files': [f['saved_name'] for f in saved_files]
                })

    except Exception as e:
        print(f"Forecast上傳處理錯誤: {str(e)}")
        if user:
            log_upload(user['id'], 'forecast', original_filename, 0, 0, 0, 'failed', str(e))
        return jsonify({'success': False, 'message': f'上傳處理失敗: {str(e)}'})

@app.route('/upload_transit', methods=['POST'])
@login_required
def upload_transit():
    user = get_current_user()
    original_filename = ''
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': '沒有選擇文件'})

        original_filename = file.filename

        if file and allowed_file(file.filename):
            # 使用資料夾管理結構：uploads/{user_id}/{session_timestamp}/transit_data.xlsx
            upload_folder, session_timestamp = get_or_create_session_folder(user['id'], 'uploads')
            filename = 'transit_data.xlsx'
            filepath = os.path.join(upload_folder, filename)

            # 保存文件
            file.save(filepath)
            print(f"在途文件已保存到: {filepath} (session: {session_timestamp})")

            # 檢查文件是否真的存在
            if not os.path.exists(filepath):
                return jsonify({'success': False, 'message': '文件保存失敗'})

            # ========== 格式驗證 ==========
            # 檢查是否為測試模式，如果是則使用被測試客戶的 username
            test_mode = request.form.get('test_mode') == 'true'
            customer_id = request.form.get('customer_id')
            template_username = user['username']

            if test_mode and customer_id and user['role'] in ['admin', 'it']:
                # 測試模式：取得被測試客戶的 username
                test_customer = get_user_by_id(int(customer_id))
                if test_customer:
                    template_username = test_customer['username']
                    print(f"[測試模式] 使用客戶 {template_username} 的模板進行驗證")

            print(f"開始驗證在途文件格式...（用戶: {template_username}）")
            is_valid, message, details = validate_transit_format(filepath, template_username)

            if not is_valid:
                # 驗證失敗，刪除已上傳的文件
                os.remove(filepath)
                print(f"❌ 在途文件格式驗證失敗: {message}")
                if details:
                    for detail in details:
                        print(f"   - {detail}")

                # 記錄上傳失敗
                log_upload(user['id'], 'transit', original_filename, 0, 0, 0, 'validation_failed', message)
                log_activity(user['id'], user['username'], 'upload_transit_failed',
                           f"在途文件上傳失敗：{message}", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': False,
                    'message': f'在途文件格式驗證失敗：{message}',
                    'details': details,
                    'validation_error': True
                })

            print(f"✅ 在途文件格式驗證通過")
            # ========== 格式驗證結束 ==========

            # 讀取文件並返回基本信息
            try:
                df = pd.read_excel(filepath)
                file_size = os.path.getsize(filepath)
                print(f"在途文件讀取成功: {len(df)} 行, {len(df.columns)} 欄位")

                # 顯示欄位資訊
                print(f"✅ 在途文件欄位資訊:")
                for i, col in enumerate(df.columns):
                    print(f"   索引{i}: {col}")

                # 將檔案路徑存入 session，供後續處理使用
                set_user_file_path('transit', filepath)
                print(f"✅ Transit 檔案路徑已設置到 session: {filepath}")
                print(f"   Session keys: {list(session.keys())}")
                print(f"   current_transit_file: {session.get('current_transit_file')}")

                # 記錄上傳成功
                log_upload(user['id'], 'transit', original_filename, file_size, len(df), len(df.columns), 'success')
                log_activity(user['id'], user['username'], 'upload_transit',
                           f"在途文件上傳成功：{original_filename} -> {filename}", get_client_ip(), request.headers.get('User-Agent'))

                return jsonify({
                    'success': True,
                    'message': '在途文件上傳成功（格式驗證通過）',
                    'rows': len(df),
                    'columns': list(df.columns),
                    'file_size': file_size,
                    'saved_filename': filename
                })
            except Exception as e:
                print(f"在途文件讀取失敗: {str(e)}")
                log_upload(user['id'], 'transit', original_filename, 0, 0, 0, 'failed', str(e))
                return jsonify({'success': False, 'message': f'文件讀取失敗: {str(e)}'})

        return jsonify({'success': False, 'message': '不支持的文件格式'})

    except Exception as e:
        print(f"在途文件上傳處理錯誤: {str(e)}")
        if user:
            log_upload(user['id'], 'transit', original_filename, 0, 0, 0, 'failed', str(e))
        return jsonify({'success': False, 'message': f'上傳處理失敗: {str(e)}'})

@app.route('/mapping')
@login_required
def mapping():
    user = get_current_user()
    response = make_response(render_template('mapping.html', user=user))
    # 設置不緩存HTML頁面
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/get_mapping_data')
@login_required
def get_mapping_data():
    user = get_current_user()
    try:
        # 1. 首先嘗試從資料庫讀取用戶的 mapping 資料
        if has_customer_mappings(user['id']):
            print(f"從資料庫讀取用戶 {user['username']} 的 mapping 資料...")
            existing_mapping = get_customer_mappings(user['id'])

            if existing_mapping:
                # 收集所有客戶名稱
                all_customers = set()
                all_customers.update(existing_mapping.get('regions', {}).keys())
                all_customers.update(existing_mapping.get('schedule_breakpoints', {}).keys())
                all_customers.update(existing_mapping.get('etd', {}).keys())
                all_customers.update(existing_mapping.get('eta', {}).keys())

                return jsonify({
                    'success': True,
                    'customers': list(all_customers),
                    'customer_column': '客戶簡稱',
                    'existing_mapping': existing_mapping,
                    'source': 'database'
                })

        # 2. 如果資料庫沒有資料，嘗試從 mapping 表 Excel 讀取（向後相容）
        mapping_file = os.path.join('mapping', 'mapping表.xlsx')
        if os.path.exists(mapping_file):
            print("資料庫無資料，從 mapping表.xlsx 讀取...")
            mapping_df = pd.read_excel(mapping_file)

            # 找到客戶簡稱欄位（通常是A欄位）
            customer_col = mapping_df.columns[0]  # A欄位

            # 找到其他欄位
            region_col = None
            schedule_col = None
            etd_col = None
            eta_col = None

            for col in mapping_df.columns:
                if '地區' in str(col) or 'region' in str(col).lower():
                    region_col = col
                elif '排程' in str(col) or '斷點' in str(col):
                    schedule_col = col
                elif 'ETD' in str(col):
                    etd_col = col
                elif 'ETA' in str(col):
                    eta_col = col

            # 構建現有的映射數據
            existing_mapping = {
                'regions': {},
                'schedule_breakpoints': {},
                'etd': {},
                'eta': {}
            }

            for idx, row in mapping_df.iterrows():
                customer = str(row[customer_col])
                if region_col and pd.notna(row[region_col]):
                    existing_mapping['regions'][customer] = str(row[region_col])
                if schedule_col and pd.notna(row[schedule_col]):
                    existing_mapping['schedule_breakpoints'][customer] = str(row[schedule_col])
                if etd_col and pd.notna(row[etd_col]):
                    existing_mapping['etd'][customer] = str(row[etd_col])
                if eta_col and pd.notna(row[eta_col]):
                    existing_mapping['eta'][customer] = str(row[eta_col])

            return jsonify({
                'success': True,
                'customers': list(existing_mapping['regions'].keys()),
                'customer_column': customer_col,
                'existing_mapping': existing_mapping,
                'source': 'mapping_table'
            })

        # 3. 如果都沒有，則從ERP文件獲取客戶列表
        erp_file = get_user_file_path('erp')
        if not erp_file or not os.path.exists(erp_file):
            return jsonify({'success': False, 'message': '請先上傳ERP文件或配置映射表'})

        print("從ERP文件獲取客戶數據...")
        df = pd.read_excel(erp_file)

        # 找到客戶簡稱欄位
        customer_col = None
        for col in df.columns:
            if '客戶' in str(col) and '簡稱' in str(col):
                customer_col = col
                break

        if customer_col is None:
            return jsonify({'success': False, 'message': '找不到客戶簡稱欄位'})

        # 獲取唯一的客戶簡稱
        unique_customers = df[customer_col].dropna().unique().tolist()

        return jsonify({
            'success': True,
            'customers': unique_customers,
            'customer_column': customer_col,
            'existing_mapping': {},
            'source': 'erp_file'
        })
    except Exception as e:
        print(f"獲取映射數據失敗: {str(e)}")
        return jsonify({'success': False, 'message': f'獲取映射數據失敗: {str(e)}'})

@app.route('/save_mapping', methods=['POST'])
@login_required
def save_mapping():
    user = get_current_user()
    try:
        mapping_data = request.json

        # 統計客戶數量
        all_customers = set()
        all_customers.update(mapping_data.get('regions', {}).keys())
        all_customers.update(mapping_data.get('schedule_breakpoints', {}).keys())
        all_customers.update(mapping_data.get('etd', {}).keys())
        all_customers.update(mapping_data.get('eta', {}).keys())
        customer_count = len(all_customers)

        # 1. 儲存到資料庫（主要儲存方式，按用戶區分）
        if save_customer_mappings(user['id'], mapping_data):
            print(f"✅ 已儲存 mapping 資料到資料庫 (user: {user['username']})")

            # 記錄 LOG：mapping 配置保存成功
            log_activity(
                user_id=user['id'],
                username=user['username'],
                action_type='mapping_config_save',
                action_detail=f"用戶 {user['display_name']} 保存映射配置，共 {customer_count} 個客戶",
                ip_address=get_client_ip(),
                user_agent=request.headers.get('User-Agent')
            )
        else:
            # 記錄 LOG：mapping 配置保存失敗
            log_activity(
                user_id=user['id'],
                username=user['username'],
                action_type='mapping_config_failed',
                action_detail=f"用戶 {user['display_name']} 保存映射配置失敗：儲存至資料庫失敗",
                ip_address=get_client_ip(),
                user_agent=request.headers.get('User-Agent')
            )
            return jsonify({'success': False, 'message': '儲存映射資料到資料庫失敗'})

        # 2. 同時更新 Excel 文件（向後相容，保持舊格式）
        mapping_excel_file = os.path.join('mapping', 'mapping表.xlsx')
        if os.path.exists(mapping_excel_file):
            # 讀取現有Excel文件
            mapping_df = pd.read_excel(mapping_excel_file)

            # 更新數據
            for idx, row in mapping_df.iterrows():
                customer = str(row.iloc[0])  # 客戶簡稱欄位

                # 更新各個欄位
                if customer in mapping_data.get('regions', {}):
                    mapping_df.iloc[idx, 1] = mapping_data['regions'][customer]  # 客戶需求地區
                if customer in mapping_data.get('schedule_breakpoints', {}):
                    mapping_df.iloc[idx, 3] = mapping_data['schedule_breakpoints'][customer]  # 排程出貨日期斷點
                if customer in mapping_data.get('etd', {}):
                    mapping_df.iloc[idx, 4] = mapping_data['etd'][customer]  # ETD
                if customer in mapping_data.get('eta', {}):
                    mapping_df.iloc[idx, 5] = mapping_data['eta'][customer]  # ETA

            # 保存更新後的Excel文件
            mapping_df.to_excel(mapping_excel_file, index=False)
            print(f"✅ 已同步更新 Excel 文件: {mapping_excel_file}")

        return jsonify({'success': True, 'message': '映射表保存成功（已儲存至資料庫）'})
    except Exception as e:
        print(f"❌ 保存映射表失敗: {str(e)}")
        # 記錄 LOG：mapping 配置保存異常
        log_activity(
            user_id=user['id'],
            username=user['username'],
            action_type='mapping_config_failed',
            action_detail=f"用戶 {user['display_name']} 保存映射配置異常：{str(e)}",
            ip_address=get_client_ip(),
            user_agent=request.headers.get('User-Agent')
        )
        return jsonify({'success': False, 'message': f'保存映射表失敗: {str(e)}'})

@app.route('/process_forecast_cleanup', methods=['POST'])
@login_required
def process_forecast_cleanup():
    user = get_current_user()
    start_time = time.time()
    try:
        # 檢查是否為測試模式（支援無 body 的請求）
        data = {}
        try:
            data = request.get_json(silent=True) or {}
        except:
            pass
        test_mode = data.get('test_mode', False)

        # 記錄開始處理
        log_activity(user['id'], user['username'], 'cleanup_start',
                   f"開始 Forecast 數據清理{' (測試模式)' if test_mode else ''}", get_client_ip(), request.headers.get('User-Agent'))

        # 檢查是否為多檔案模式
        merge_mode = get_user_file_path('forecast_merge_mode')
        forecast_files = get_user_file_path('forecast_files')  # 多檔案列表
        forecast_file = get_user_file_path('forecast')  # 單一檔案（合併模式）

        # 使用資料夾管理結構
        processed_folder, session_timestamp = get_or_create_session_folder(user['id'], 'processed')

        # 存儲 processed 資料夾路徑到 session
        session['current_processed_folder'] = processed_folder

        # 多檔案分開模式
        if merge_mode is False and forecast_files and len(forecast_files) > 0:
            print(f"=== 多檔案清理模式：{len(forecast_files)} 個檔案 ===")

            total_cleaned_count = 0
            cleaned_files_info = []

            for idx, file_info in enumerate(forecast_files):
                file_path = file_info.get('path')
                original_name = file_info.get('original_name') or file_info.get('name', f'forecast_{idx + 1}.xlsx')

                if not file_path or not os.path.exists(file_path):
                    print(f"  ⚠️ 檔案不存在: {file_path}")
                    cleaned_files_info.append({
                        'name': original_name,
                        'cleaned_cells': 0,
                        'status': 'error',
                        'message': '檔案不存在'
                    })
                    continue

                print(f"  清理檔案 {idx + 1}/{len(forecast_files)}: {original_name}")

                try:
                    # 使用openpyxl保持格式
                    wb = load_workbook(file_path)
                    ws = wb.active

                    # 清理數據
                    cleaned_count = 0

                    for row_idx in range(1, ws.max_row + 1):
                        # 檢查K欄位（第11列）是否為"供應數量"
                        k_cell = ws.cell(row=row_idx, column=11)
                        if k_cell.value and str(k_cell.value) == "供應數量":
                            # 清空L~AW欄位（第12列到第49列）
                            for col_idx in range(12, min(50, ws.max_column + 1)):
                                cell = ws.cell(row=row_idx, column=col_idx)
                                if cell.value != 0:
                                    cell.value = 0
                                    cleaned_count += 1

                        # 檢查I欄位（第9列）是否包含"庫存數量"
                        i_cell = ws.cell(row=row_idx, column=9)
                        if i_cell.value and "庫存數量" in str(i_cell.value):
                            next_row_i_cell = ws.cell(row=row_idx + 1, column=9)
                            if next_row_i_cell.value != 0:
                                next_row_i_cell.value = 0
                                cleaned_count += 1

                    # 保存清理後的文件
                    cleaned_filename = f'cleaned_forecast_{idx + 1}.xlsx'
                    cleaned_file = os.path.join(processed_folder, cleaned_filename)
                    wb.save(cleaned_file)

                    total_cleaned_count += cleaned_count
                    cleaned_files_info.append({
                        'name': original_name,
                        'cleaned_cells': cleaned_count,
                        'status': 'success',
                        'cleaned_path': cleaned_file
                    })
                    print(f"    ✅ 清理了 {cleaned_count} 個單元格")

                except Exception as file_error:
                    print(f"    ❌ 清理失敗: {str(file_error)}")
                    cleaned_files_info.append({
                        'name': original_name,
                        'cleaned_cells': 0,
                        'status': 'error',
                        'message': str(file_error)
                    })

            # 更新 session 中的清理後檔案路徑
            cleaned_paths = [f for f in cleaned_files_info if f['status'] == 'success']
            set_user_file_path('cleaned_forecast_files', cleaned_paths)

            duration = time.time() - start_time
            print(f"=== 多檔案清理完成：共清理 {total_cleaned_count} 個單元格 ===")

            log_process(user['id'], 'cleanup', 'success', f'清理了 {len(cleaned_paths)} 個檔案，共 {total_cleaned_count} 個單元格', duration)
            log_activity(user['id'], user['username'], 'cleanup_success',
                       f"Forecast 多檔案數據清理成功，{len(cleaned_paths)} 個檔案，共 {total_cleaned_count} 個單元格", get_client_ip(), request.headers.get('User-Agent'))

            return jsonify({
                'success': True,
                'message': f'Forecast數據清理完成，清理了 {len(cleaned_paths)} 個檔案，共 {total_cleaned_count} 個單元格',
                'multi_file': True,
                'file_count': len(cleaned_paths),
                'files': cleaned_files_info,
                'total_cleaned_cells': total_cleaned_count
            })

        # 單檔案模式（合併模式或原本就只上傳一個檔案）
        else:
            if not forecast_file or not os.path.exists(forecast_file):
                log_process(user['id'], 'cleanup', 'failed', '請先上傳Forecast文件')
                return jsonify({'success': False, 'message': '請先上傳Forecast文件'})

            # 使用openpyxl保持格式
            wb = load_workbook(forecast_file)
            ws = wb.active

            print("開始清理Forecast數據，保持原始格式...")

            # 清理數據
            cleaned_count = 0

            for row_idx in range(1, ws.max_row + 1):
                # 檢查K欄位（第11列）是否為"供應數量"
                k_cell = ws.cell(row=row_idx, column=11)
                if k_cell.value and str(k_cell.value) == "供應數量":
                    # 清空L~AW欄位（第12列到第49列）
                    for col_idx in range(12, min(50, ws.max_column + 1)):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if cell.value != 0:
                            cell.value = 0
                            cleaned_count += 1

                # 檢查I欄位（第9列）是否包含"庫存數量"
                i_cell = ws.cell(row=row_idx, column=9)
                if i_cell.value and "庫存數量" in str(i_cell.value):
                    next_row_i_cell = ws.cell(row=row_idx + 1, column=9)
                    if next_row_i_cell.value != 0:
                        next_row_i_cell.value = 0
                        cleaned_count += 1

            # 保存清理後的文件
            cleaned_file = os.path.join(processed_folder, 'cleaned_forecast.xlsx')
            wb.save(cleaned_file)

            duration = time.time() - start_time
            print(f"Forecast數據清理完成，清理了 {cleaned_count} 個單元格")

            # 記錄處理成功
            log_process(user['id'], 'cleanup', 'success', f'清理了 {cleaned_count} 個單元格', duration)
            log_activity(user['id'], user['username'], 'cleanup_success',
                       f"Forecast 數據清理成功，清理了 {cleaned_count} 個單元格", get_client_ip(), request.headers.get('User-Agent'))

            return jsonify({
                'success': True,
                'message': f'Forecast數據清理完成，清理了 {cleaned_count} 個單元格',
                'multi_file': False,
                'file': 'cleaned_forecast.xlsx',
                'cleaned_cells': cleaned_count
            })

    except Exception as e:
        duration = time.time() - start_time
        print(f"Forecast數據清理失敗: {str(e)}")
        log_process(user['id'], 'cleanup', 'failed', str(e), duration)
        log_activity(user['id'], user['username'], 'cleanup_failed',
                   f"Forecast 數據清理失敗：{str(e)}", get_client_ip(), request.headers.get('User-Agent'))
        return jsonify({'success': False, 'message': f'數據清理失敗: {str(e)}'})

@app.route('/process_erp_mapping', methods=['POST'])
@login_required
def process_erp_mapping():
    user = get_current_user()
    start_time = time.time()
    try:
        # 檢查是否為測試模式
        data = request.get_json() or {}
        test_mode = data.get('test_mode', False)
        test_customer_id = data.get('customer_id')

        # 決定使用哪個用戶的 mapping 資料
        mapping_user_id = user['id']
        if test_mode and test_customer_id and (user['role'] in ['admin', 'it']):
            mapping_user_id = test_customer_id
            print(f"[測試模式] 使用客戶 ID {test_customer_id} 的 mapping 資料")

        # 記錄開始處理
        log_activity(user['id'], user['username'], 'mapping_start',
                   f"開始 ERP 和在途數據整合{' (測試模式)' if test_mode else ''}", get_client_ip(), request.headers.get('User-Agent'))

        # 從 session 獲取當前用戶的檔案路徑
        erp_file = get_user_file_path('erp')
        transit_file = get_user_file_path('transit')
        mapping_excel_file = os.path.join('mapping', 'mapping表.xlsx')

        # 除錯日誌
        print(f"[映射整合] 檢查檔案路徑:")
        print(f"  - ERP 檔案路徑: {erp_file}")
        print(f"  - Transit 檔案路徑: {transit_file}")
        print(f"  - ERP 檔案存在: {os.path.exists(erp_file) if erp_file else False}")
        print(f"  - Transit 檔案存在: {os.path.exists(transit_file) if transit_file else False}")
        print(f"  - Session keys: {list(session.keys())}")

        if not erp_file or not os.path.exists(erp_file):
            log_process(user['id'], 'mapping', 'failed', '請先上傳ERP文件')
            return jsonify({'success': False, 'message': '請先上傳ERP文件'})

        if not transit_file or not os.path.exists(transit_file):
            log_process(user['id'], 'mapping', 'failed', '請先上傳在途文件')
            return jsonify({'success': False, 'message': '請先上傳在途文件'})

        # 檢查 mapping 資料來源（優先資料庫）
        mapping_data = None

        # 1. 優先從資料庫讀取（使用 mapping_user_id）
        if has_customer_mappings(mapping_user_id):
            print(f"從資料庫讀取用戶 ID {mapping_user_id} 的 mapping 資料...")
            mapping_data = get_customer_mappings(mapping_user_id)

        # 2. 如果資料庫沒有，嘗試從 JSON 檔案讀取（向後相容）
        if not mapping_data:
            mapping_file = os.path.join('mapping', 'mapping_data.json')
            if os.path.exists(mapping_file):
                print("從 mapping_data.json 讀取映射資料...")
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mapping_data = json.load(f)

        # 3. 檢查是否有 mapping 資料
        if not mapping_data and not os.path.exists(mapping_excel_file):
            log_process(user['id'], 'mapping', 'failed', '請先配置映射表')
            return jsonify({'success': False, 'message': '請先配置映射表'})

        # === 1. 處理 ERP 數據整合 ===
        print("開始整合 ERP 數據...")
        erp_df = pd.read_excel(erp_file)

        # 標準化排程出貨日期欄位（處理文字/日期格式不一致問題）
        if '排程出貨日期' in erp_df.columns:
            print("🔧 標準化排程出貨日期欄位格式...")

            # 顯示處理前的數據類型統計
            date_col = erp_df['排程出貨日期']
            print(f"   處理前數據類型統計:")
            print(f"   - 總記錄數: {len(date_col)}")
            print(f"   - 非空記錄數: {date_col.notna().sum()}")
            print(f"   - 字串類型: {sum(isinstance(x, str) for x in date_col if pd.notna(x))}")
            print(f"   - 日期類型: {sum(isinstance(x, (datetime, pd.Timestamp)) for x in date_col if pd.notna(x))}")

            # 標準化日期格式
            erp_df['排程出貨日期'] = erp_df['排程出貨日期'].apply(normalize_date_for_mapping)

            # 顯示處理後的結果
            processed_col = erp_df['排程出貨日期']
            print(f"   處理後結果:")
            print(f"   - 成功標準化: {processed_col.notna().sum()}")
            print(f"   - 標準化失敗: {processed_col.isna().sum()}")
            if processed_col.notna().sum() > 0:
                sample_dates = processed_col.dropna().head(3).tolist()
                print(f"   - 範例日期: {sample_dates}")

            print("✅ 排程出貨日期欄位已標準化")
        
        # 找到客戶簡稱欄位
        customer_col = None
        for col in erp_df.columns:
            if '客戶' in str(col) and '簡稱' in str(col):
                customer_col = col
                break
        
        if customer_col is None:
            return jsonify({'success': False, 'message': 'ERP文件找不到客戶簡稱欄位'})
        
        # 應用映射到 ERP
        erp_df['客戶需求地區'] = erp_df[customer_col].map(mapping_data.get('regions', {}))
        erp_df['排程出貨日期斷點'] = erp_df[customer_col].map(mapping_data.get('schedule_breakpoints', {}))
        erp_df['ETD'] = erp_df[customer_col].map(mapping_data.get('etd', {}))
        erp_df['ETA'] = erp_df[customer_col].map(mapping_data.get('eta', {}))
        
        # 按排程出貨日期排序
        if '排程出貨日期' in erp_df.columns:
            erp_df = erp_df.sort_values('排程出貨日期')

        # 使用資料夾管理結構：processed/{user_id}/{session_timestamp}/integrated_erp.xlsx
        processed_folder, session_timestamp = get_or_create_session_folder(user['id'], 'processed')

        # 保存整合後的 ERP 文件
        integrated_erp_file = os.path.join(processed_folder, 'integrated_erp.xlsx')
        erp_df.to_excel(integrated_erp_file, index=False)
        print(f"✅ ERP數據整合完成: {len(erp_df)} 行 (session: {session_timestamp})")
        
        # === 2. 處理 Transit 數據整合 ===
        print("開始整合在途數據...")
        transit_df = pd.read_excel(transit_file)

        # 建立 mapping 字典（從 mapping_data 轉換格式）
        mapping_dict = {}
        if mapping_data:
            # 從資料庫格式的 mapping_data 建立字典
            all_customers = set()
            all_customers.update(mapping_data.get('regions', {}).keys())
            all_customers.update(mapping_data.get('schedule_breakpoints', {}).keys())
            all_customers.update(mapping_data.get('etd', {}).keys())
            all_customers.update(mapping_data.get('eta', {}).keys())

            for customer in all_customers:
                mapping_dict[customer] = {
                    'region': mapping_data.get('regions', {}).get(customer, ''),
                    'schedule_breakpoint': mapping_data.get('schedule_breakpoints', {}).get(customer, ''),
                    'etd': mapping_data.get('etd', {}).get(customer, ''),
                    'eta': mapping_data.get('eta', {}).get(customer, '')
                }
            print(f"從資料庫/JSON 建立 mapping 字典，共 {len(mapping_dict)} 個客戶")
        elif os.path.exists(mapping_excel_file):
            # 向後相容：從 Excel 讀取 mapping
            mapping_excel_df = pd.read_excel(mapping_excel_file)

            # mapping 表的欄位結構
            mapping_customer_col = mapping_excel_df.columns[0]  # A欄位
            mapping_region_col = mapping_excel_df.columns[1] if len(mapping_excel_df.columns) > 1 else None
            mapping_schedule_col = mapping_excel_df.columns[3] if len(mapping_excel_df.columns) > 3 else None
            mapping_etd_col = mapping_excel_df.columns[4] if len(mapping_excel_df.columns) > 4 else None
            mapping_eta_col = mapping_excel_df.columns[5] if len(mapping_excel_df.columns) > 5 else None

            for idx, row in mapping_excel_df.iterrows():
                customer = str(row[mapping_customer_col])
                mapping_dict[customer] = {
                    'region': str(row[mapping_region_col]) if mapping_region_col and pd.notna(row[mapping_region_col]) else '',
                    'schedule_breakpoint': str(row[mapping_schedule_col]) if mapping_schedule_col and pd.notna(row[mapping_schedule_col]) else '',
                    'etd': str(row[mapping_etd_col]) if mapping_etd_col and pd.notna(row[mapping_etd_col]) else '',
                    'eta': str(row[mapping_eta_col]) if mapping_eta_col and pd.notna(row[mapping_eta_col]) else ''
                }
            print(f"從 Excel 建立 mapping 字典，共 {len(mapping_dict)} 個客戶")
        
        # transit_data 的 E 欄位（索引4）與 mapping 表 A 欄位比對
        if len(transit_df.columns) < 5:
            return jsonify({'success': False, 'message': '在途文件欄位不足，需要至少 E 欄位'})
        
        transit_customer_col = transit_df.columns[4]  # E欄位
        print(f"在途文件 E 欄位名稱: {transit_customer_col}")
        
        # 新版在途文件結構（12欄位）：
        # 索引0-11: Tw, Ship Number, Invoice Date, Location, 客戶簡稱, Ordered Item, Pj Item, Qty, ETA, Stauts, 集團客戶, 週別
        # 整合後會新增4個欄位到末尾，變成索引12-15：客戶需求地區, 排程出貨日期斷點, ETD, ETA
        
        # 應用映射到 Transit（新增到文件末尾）
        transit_df['客戶需求地區'] = transit_df[transit_customer_col].apply(
            lambda x: mapping_dict.get(str(x), {}).get('region', '') if pd.notna(x) else ''
        )
        transit_df['排程出貨日期斷點'] = transit_df[transit_customer_col].apply(
            lambda x: mapping_dict.get(str(x), {}).get('schedule_breakpoint', '') if pd.notna(x) else ''
        )
        transit_df['ETD'] = transit_df[transit_customer_col].apply(
            lambda x: mapping_dict.get(str(x), {}).get('etd', '') if pd.notna(x) else ''
        )
        transit_df['ETA_mapping'] = transit_df[transit_customer_col].apply(
            lambda x: mapping_dict.get(str(x), {}).get('eta', '') if pd.notna(x) else ''
        )
        
        # 顯示整合後的欄位結構
        print(f"✅ 在途數據整合完成，欄位結構:")
        for i, col in enumerate(transit_df.columns):
            print(f"   索引{i}: {col}")
        print(f"   整合後總欄位數: {len(transit_df.columns)}")
        
        # 注意：整合後的結構（總共16個欄位，索引0-15）
        # 索引8: ETA (原始文件中的ETA)
        # 索引12: 客戶需求地區 (整合後新增)
        # 索引13: 排程出貨日期斷點 (整合後新增)
        # 索引14: ETD (整合後新增)
        # 索引15: ETA_mapping (整合後新增，來自mapping表)

        # 保存整合後的 Transit 文件（使用同一個 session 資料夾）
        integrated_transit_file = os.path.join(processed_folder, 'integrated_transit.xlsx')
        transit_df.to_excel(integrated_transit_file, index=False)
        print(f"✅ 在途數據整合完成: {len(transit_df)} 行 (session: {session_timestamp})")

        # 存儲 processed 資料夾路徑到 session
        session['current_processed_folder'] = processed_folder

        duration = time.time() - start_time
        # 記錄處理成功
        log_process(user['id'], 'mapping', 'success', f'ERP: {len(erp_df)} 行, Transit: {len(transit_df)} 行', duration)
        log_activity(user['id'], user['username'], 'mapping_success',
                   f"ERP 和在途數據整合成功", get_client_ip(), request.headers.get('User-Agent'))

        return jsonify({
            'success': True,
            'message': 'ERP 和在途數據整合完成',
            'erp_file': 'integrated_erp.xlsx',
            'transit_file': 'integrated_transit.xlsx',
            'erp_rows': len(erp_df),
            'transit_rows': len(transit_df)
        })
    except Exception as e:
        duration = time.time() - start_time
        print(f"數據整合失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        log_process(user['id'], 'mapping', 'failed', str(e), duration)
        log_activity(user['id'], user['username'], 'mapping_failed',
                   f"ERP 和在途數據整合失敗：{str(e)}", get_client_ip(), request.headers.get('User-Agent'))
        return jsonify({'success': False, 'message': f'數據整合失敗: {str(e)}'})

@app.route('/run_forecast', methods=['POST'])
@login_required
def run_forecast():
    user = get_current_user()
    start_time = time.time()
    try:
        # 檢查是否為測試模式
        data = request.get_json() or {}
        test_mode = data.get('test_mode', False)

        # 記錄開始處理
        log_activity(user['id'], user['username'], 'forecast_start',
                   f"開始 FORECAST 處理{' (測試模式)' if test_mode else ''}", get_client_ip(), request.headers.get('User-Agent'))

        # 從 session 獲取當前的 processed 資料夾
        processed_folder = session.get('current_processed_folder')
        if not processed_folder:
            # 嘗試使用 session 資料夾路徑
            processed_folder, _ = get_session_folder_path(user['id'], 'processed')

        if not processed_folder or not os.path.exists(processed_folder):
            log_process(user['id'], 'forecast', 'failed', '請先完成數據清理和整合')
            return jsonify({'success': False, 'message': '請先完成數據清理和整合'})

        # 檢查必要文件是否存在
        cleaned_forecast = os.path.join(processed_folder, 'cleaned_forecast.xlsx')
        integrated_erp = os.path.join(processed_folder, 'integrated_erp.xlsx')
        integrated_transit = os.path.join(processed_folder, 'integrated_transit.xlsx')

        if not os.path.exists(cleaned_forecast):
            log_process(user['id'], 'forecast', 'failed', '請先完成Forecast數據清理')
            return jsonify({'success': False, 'message': '請先完成Forecast數據清理'})

        if not os.path.exists(integrated_erp):
            log_process(user['id'], 'forecast', 'failed', '請先完成ERP數據整合')
            return jsonify({'success': False, 'message': '請先完成ERP數據整合'})

        print("開始FORECAST處理...")
        print(f"清理後的Forecast文件: {cleaned_forecast}")
        print(f"整合後的ERP文件: {integrated_erp}")

        # 檢查是否有 Transit 文件
        has_transit = os.path.exists(integrated_transit)
        if has_transit:
            print(f"整合後的Transit文件: {integrated_transit}")
        else:
            print("⚠️ 未找到Transit文件，將跳過Transit數據處理")

        # 執行FORECAST處理
        from ultra_fast_forecast_processor import UltraFastForecastProcessor

        processor = UltraFastForecastProcessor(
            forecast_file=cleaned_forecast,
            erp_file=integrated_erp,
            transit_file=integrated_transit if has_transit else None,
            output_folder=processed_folder  # 傳遞輸出資料夾路徑
        )

        success = processor.process_all_blocks()

        if success:
            # 檢查結果文件是否真的生成了
            result_file = os.path.join(processed_folder, 'forecast_result.xlsx')
            if os.path.exists(result_file):
                file_size = os.path.getsize(result_file)
                duration = time.time() - start_time
                print(f"FORECAST處理完成，結果文件: {result_file} (大小: {file_size} bytes)")

                # 記錄處理成功
                log_process(user['id'], 'forecast', 'success',
                          f'ERP填入: {processor.total_filled}, Transit填入: {processor.total_transit_filled if has_transit else 0}', duration)
                log_activity(user['id'], user['username'], 'forecast_success',
                           f"FORECAST 處理成功", get_client_ip(), request.headers.get('User-Agent'))

                result_data = {
                    'success': True,
                    'message': 'FORECAST處理完成',
                    'file': 'forecast_result.xlsx',
                    'erp_filled': processor.total_filled,
                    'erp_skipped': processor.total_skipped,
                    'file_size': file_size
                }

                # 如果有 Transit 數據，也返回 Transit 統計
                if has_transit:
                    result_data['transit_filled'] = processor.total_transit_filled
                    result_data['transit_skipped'] = processor.total_transit_skipped

                return jsonify(result_data)
            else:
                duration = time.time() - start_time
                print("錯誤：結果文件未生成")
                log_process(user['id'], 'forecast', 'failed', '結果文件未生成', duration)
                return jsonify({'success': False, 'message': 'FORECAST處理完成但結果文件未生成'})
        else:
            duration = time.time() - start_time
            print("FORECAST處理失敗")
            log_process(user['id'], 'forecast', 'failed', '處理失敗', duration)
            log_activity(user['id'], user['username'], 'forecast_failed',
                       "FORECAST 處理失敗", get_client_ip(), request.headers.get('User-Agent'))
            return jsonify({'success': False, 'message': 'FORECAST處理失敗'})

    except Exception as e:
        duration = time.time() - start_time
        print(f"FORECAST處理異常: {str(e)}")
        import traceback
        traceback.print_exc()
        log_process(user['id'], 'forecast', 'failed', str(e), duration)
        log_activity(user['id'], user['username'], 'forecast_failed',
                   f"FORECAST 處理失敗：{str(e)}", get_client_ip(), request.headers.get('User-Agent'))
        return jsonify({'success': False, 'message': f'FORECAST處理失敗: {str(e)}'})

@app.route('/check_files')
@login_required
def check_files():
    try:
        # 從 session 獲取當前用戶的檔案路徑
        erp_file = get_user_file_path('erp')
        forecast_file = get_user_file_path('forecast')
        transit_file = get_user_file_path('transit')

        erp_exists = erp_file and os.path.exists(erp_file)
        forecast_exists = forecast_file and os.path.exists(forecast_file)
        transit_exists = transit_file and os.path.exists(transit_file)

        return jsonify({
            'success': True,
            'erp_exists': erp_exists,
            'forecast_exists': forecast_exists,
            'transit_exists': transit_exists,
            'erp_size': os.path.getsize(erp_file) if erp_exists else 0,
            'forecast_size': os.path.getsize(forecast_file) if forecast_exists else 0,
            'transit_size': os.path.getsize(transit_file) if transit_exists else 0
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/download/<filename>')
@login_required
def download_file(filename):
    user = get_current_user()

    # 從 session 獲取當前的 processed 資料夾
    processed_folder = session.get('current_processed_folder')
    if not processed_folder:
        processed_folder, _ = get_session_folder_path(user['id'], 'processed')

    # 嘗試從用戶的 session 資料夾下載
    if processed_folder:
        file_path = os.path.join(processed_folder, filename)
        if os.path.exists(file_path):
            # 記錄下載活動
            file_size = os.path.getsize(file_path)
            log_activity(user['id'], user['username'], 'download',
                       f"下載文件：{filename} ({file_size} bytes)", get_client_ip(), request.headers.get('User-Agent'))

            # 為下載檔案加上時間戳
            session_timestamp = session.get('current_session_timestamp', '')
            if session_timestamp:
                # 從檔名分離副檔名
                name, ext = os.path.splitext(filename)
                download_filename = f"{name}_{session_timestamp}{ext}"
            else:
                download_filename = filename

            return send_file(file_path, as_attachment=True, download_name=download_filename)

    # 向後相容：嘗試從舊的 processed 資料夾下載
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    if os.path.exists(file_path):
        file_size = os.path.getsize(file_path)
        log_activity(user['id'], user['username'], 'download',
                   f"下載文件：{filename} ({file_size} bytes)", get_client_ip(), request.headers.get('User-Agent'))
        return send_file(file_path, as_attachment=True)

    return jsonify({'success': False, 'message': '文件不存在'})

# 靜態資源路由，添加版本控制
@app.route('/static/<path:filename>')
def static_files(filename):
    """提供靜態文件並設置適當的緩存策略"""
    file_path = os.path.join('static', filename)
    if os.path.exists(file_path):
        response = make_response(send_file(file_path))

        # 根據文件類型設置不同的緩存策略
        if filename.endswith(('.css', '.js')):
            # CSS和JS文件設置較短的緩存時間，但允許緩存
            response.headers['Cache-Control'] = 'public, max-age=3600'  # 1小時
        elif filename.endswith(('.png', '.jpg', '.jpeg', '.gif', '.ico')):
            # 圖片文件可以緩存更長時間
            response.headers['Cache-Control'] = 'public, max-age=86400'  # 24小時
        else:
            # 其他文件不緩存
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'

        return response
    else:
        return jsonify({'error': 'File not found'}), 404


# ========================================
# IT/Admin 管理介面路由
# ========================================

@app.route('/admin')
@admin_required
def admin_dashboard():
    """管理者首頁"""
    user = get_current_user()
    response = make_response(render_template('admin.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/it')
@it_or_admin_required
def it_dashboard():
    """IT 人員首頁"""
    user = get_current_user()
    response = make_response(render_template('it_dashboard.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/logs')
@it_or_admin_required
def logs_page():
    """LOG 查看頁面"""
    user = get_current_user()
    response = make_response(render_template('logs.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/users_manage')
@admin_required
def users_manage_page():
    """用戶管理頁面（僅管理者）"""
    user = get_current_user()
    response = make_response(render_template('users_manage.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/mappings_view')
@admin_required
def mappings_view_page():
    """客戶映射檢視頁面（僅管理者）"""
    user = get_current_user()
    response = make_response(render_template('mappings_view.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/test_function')
@it_or_admin_required
def test_function_page():
    """測試功能頁面"""
    user = get_current_user()
    response = make_response(render_template('test_function.html', user=user))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


# ========================================
# IT/Admin 管理介面 API
# ========================================

@app.route('/api/logs/activity')
@it_or_admin_required
def api_get_activity_logs():
    """取得活動日誌"""
    try:
        user_id = request.args.get('user_id', type=int)
        action_type = request.args.get('action_type')
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)

        offset = (page - 1) * per_page
        records, total = get_activity_logs_filtered(
            user_id=user_id, action_type=action_type,
            start_date=start_date, end_date=end_date,
            limit=per_page, offset=offset
        )

        # 處理日期格式（JSON 序列化）
        for record in records:
            if record.get('created_at'):
                record['created_at'] = record['created_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({
            'success': True,
            'records': records,
            'total': total,
            'page': page,
            'per_page': per_page,
            'total_pages': (total + per_page - 1) // per_page if total > 0 else 0
        })
    except Exception as e:
        print(f"❌ 取得活動日誌失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/logs/upload')
@it_or_admin_required
def api_get_upload_records():
    """取得上傳記錄"""
    try:
        user_id = request.args.get('user_id', type=int)
        file_type = request.args.get('file_type')
        status = request.args.get('status')
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)

        offset = (page - 1) * per_page
        records, total = get_upload_records(
            user_id=user_id, file_type=file_type, status=status,
            start_date=start_date, end_date=end_date,
            limit=per_page, offset=offset
        )

        # 處理日期格式
        for record in records:
            if record.get('created_at'):
                record['created_at'] = record['created_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({
            'success': True,
            'records': records,
            'total': total,
            'page': page,
            'per_page': per_page,
            'total_pages': (total + per_page - 1) // per_page if total > 0 else 0
        })
    except Exception as e:
        print(f"❌ 取得上傳記錄失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/logs/process')
@it_or_admin_required
def api_get_process_records():
    """取得處理記錄"""
    try:
        user_id = request.args.get('user_id', type=int)
        process_type = request.args.get('process_type')
        status = request.args.get('status')
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)

        offset = (page - 1) * per_page
        records, total = get_process_records(
            user_id=user_id, process_type=process_type, status=status,
            start_date=start_date, end_date=end_date,
            limit=per_page, offset=offset
        )

        # 處理日期格式
        for record in records:
            if record.get('created_at'):
                record['created_at'] = record['created_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({
            'success': True,
            'records': records,
            'total': total,
            'page': page,
            'per_page': per_page,
            'total_pages': (total + per_page - 1) // per_page if total > 0 else 0
        })
    except Exception as e:
        print(f"❌ 取得處理記錄失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/users')
@admin_required
def api_get_all_users():
    """取得所有用戶列表（僅管理者）"""
    try:
        users = get_all_users()

        # 處理日期格式
        for user in users:
            if user.get('created_at'):
                user['created_at'] = user['created_at'].strftime('%Y-%m-%d %H:%M:%S')
            if user.get('last_login'):
                user['last_login'] = user['last_login'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({'success': True, 'users': users})
    except Exception as e:
        print(f"❌ 取得用戶列表失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/users', methods=['POST'])
@admin_required
def api_create_user():
    """建立新用戶（僅管理者）"""
    try:
        data = request.get_json()

        # 驗證必填欄位
        required_fields = ['username', 'password', 'display_name']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'message': f'缺少必填欄位: {field}'})

        # 驗證角色
        role = data.get('role', 'user')
        if role not in ['admin', 'it', 'user']:
            return jsonify({'success': False, 'message': '無效的角色'})

        # 建立用戶
        success, message, user_id = create_user(
            username=data['username'],
            password=data['password'],
            display_name=data['display_name'],
            role=role,
            company=data.get('company'),
            is_active=data.get('is_active', True)
        )

        if success:
            # 建立用戶專屬的 compare 資料夾
            user_compare_folder = os.path.join(app.root_path, 'compare', data['username'])
            try:
                os.makedirs(user_compare_folder, exist_ok=True)
                print(f"✅ 已建立用戶範本資料夾: {user_compare_folder}")
            except Exception as folder_error:
                print(f"⚠️ 建立用戶範本資料夾失敗: {folder_error}")

            # 記錄操作日誌
            admin_user = session.get('user', {})
            log_activity(
                user_id=admin_user.get('id'),
                username=admin_user.get('username'),
                action_type='user_create',
                action_detail=f"建立用戶: {data['username']} ({data['display_name']}), 角色: {role}",
                ip_address=request.remote_addr,
                user_agent=request.headers.get('User-Agent')
            )
            return jsonify({'success': True, 'message': message, 'user_id': user_id})
        else:
            return jsonify({'success': False, 'message': message})

    except Exception as e:
        print(f"❌ 建立用戶失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/users/<int:user_id>', methods=['PUT'])
@admin_required
def api_update_user(user_id):
    """更新用戶資料（僅管理者）"""
    try:
        data = request.get_json()

        # 取得原用戶資料（用於日誌記錄）
        original_user = get_user_by_id(user_id)
        if not original_user:
            return jsonify({'success': False, 'message': '用戶不存在'})

        # 驗證角色（如果有提供）
        if 'role' in data and data['role'] not in ['admin', 'it', 'user']:
            return jsonify({'success': False, 'message': '無效的角色'})

        # 建立更新參數
        update_params = {}
        if 'username' in data:
            update_params['username'] = data['username']
        if 'display_name' in data:
            update_params['display_name'] = data['display_name']
        if 'password' in data and data['password']:
            update_params['password'] = data['password']
        if 'role' in data:
            update_params['role'] = data['role']
        if 'company' in data:
            update_params['company'] = data['company']
        if 'is_active' in data:
            update_params['is_active'] = data['is_active']

        if not update_params:
            return jsonify({'success': False, 'message': '沒有要更新的資料'})

        # 更新用戶
        success, message = update_user(user_id, **update_params)

        if success:
            # 記錄操作日誌
            admin_user = session.get('user', {})

            # 判斷是切換狀態還是一般更新
            if len(update_params) == 1 and 'is_active' in update_params:
                action_type = 'user_toggle_status'
                status_text = '啟用' if update_params['is_active'] else '停用'
                action_detail = f"切換用戶狀態: {original_user['username']} -> {status_text}"
            else:
                action_type = 'user_update'
                changes = []
                if 'username' in update_params:
                    changes.append(f"用戶名: {original_user['username']} -> {update_params['username']}")
                if 'display_name' in update_params:
                    changes.append(f"顯示名稱: {update_params['display_name']}")
                if 'role' in update_params:
                    changes.append(f"角色: {update_params['role']}")
                if 'company' in update_params:
                    changes.append(f"公司: {update_params['company']}")
                if 'password' in update_params:
                    changes.append("密碼已更新")
                if 'is_active' in update_params:
                    changes.append(f"狀態: {'啟用' if update_params['is_active'] else '停用'}")
                action_detail = f"更新用戶 {original_user['username']}: " + ", ".join(changes)

            log_activity(
                user_id=admin_user.get('id'),
                username=admin_user.get('username'),
                action_type=action_type,
                action_detail=action_detail,
                ip_address=request.remote_addr,
                user_agent=request.headers.get('User-Agent')
            )
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'message': message})

    except Exception as e:
        print(f"❌ 更新用戶失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/users/<int:user_id>', methods=['DELETE'])
@admin_required
def api_delete_user(user_id):
    """刪除用戶（僅管理者）"""
    try:
        # 取得原用戶資料（用於日誌記錄）
        original_user = get_user_by_id(user_id)
        if not original_user:
            return jsonify({'success': False, 'message': '用戶不存在'})

        # 防止刪除自己
        current_user = session.get('user', {})
        if current_user.get('id') == user_id:
            return jsonify({'success': False, 'message': '無法刪除自己的帳號'})

        # 刪除用戶
        success, message, deleted_username = delete_user(user_id)

        if success:
            # 記錄操作日誌
            admin_user = session.get('user', {})
            log_activity(
                user_id=admin_user.get('id'),
                username=admin_user.get('username'),
                action_type='user_delete',
                action_detail=f"刪除用戶: {deleted_username} ({original_user['display_name']})",
                ip_address=request.remote_addr,
                user_agent=request.headers.get('User-Agent')
            )
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'message': message})

    except Exception as e:
        print(f"❌ 刪除用戶失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/mappings')
@admin_required
def api_get_all_mappings():
    """取得所有用戶的客戶映射（僅管理者）"""
    try:
        mappings = get_all_customer_mappings()

        # 處理日期格式
        for mapping in mappings:
            if mapping.get('created_at'):
                mapping['created_at'] = mapping['created_at'].strftime('%Y-%m-%d %H:%M:%S')
            if mapping.get('updated_at'):
                mapping['updated_at'] = mapping['updated_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({'success': True, 'mappings': mappings})
    except Exception as e:
        print(f"❌ 取得客戶映射失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/mappings', methods=['POST'])
@admin_required
def api_create_mapping():
    """新增客戶映射（僅管理者）"""
    try:
        data = request.get_json()
        user_id = data.get('user_id')
        customer_name = data.get('customer_name')
        delivery_location = data.get('delivery_location')
        region = data.get('region')
        schedule_breakpoint = data.get('schedule_breakpoint')
        etd = data.get('etd')
        eta = data.get('eta')

        if not user_id or not customer_name:
            return jsonify({'success': False, 'message': '用戶ID和客戶簡稱為必填'})

        if not region:
            return jsonify({'success': False, 'message': '客戶廠區為必填'})

        success, message, mapping_id = admin_create_customer_mapping(
            user_id, customer_name, delivery_location, region, schedule_breakpoint, etd, eta
        )

        if success:
            # 記錄活動
            admin_user = get_current_user()
            log_activity(admin_user['id'], 'mapping_create',
                        f"新增客戶映射: {customer_name} (user_id={user_id})",
                        request.remote_addr)

        return jsonify({'success': success, 'message': message, 'mapping_id': mapping_id})
    except Exception as e:
        print(f"❌ 新增客戶映射失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/mappings/<int:mapping_id>', methods=['PUT'])
@admin_required
def api_update_mapping(mapping_id):
    """更新客戶映射（僅管理者）"""
    try:
        data = request.get_json()

        # 只傳遞有值的欄位
        update_data = {}
        for field in ['customer_name', 'delivery_location', 'region', 'schedule_breakpoint', 'etd', 'eta']:
            if field in data:
                update_data[field] = data[field]

        if not update_data:
            return jsonify({'success': False, 'message': '沒有提供要更新的欄位'})

        success, message = admin_update_customer_mapping(mapping_id, **update_data)

        if success:
            # 記錄活動
            admin_user = get_current_user()
            log_activity(admin_user['id'], 'mapping_update',
                        f"更新客戶映射 ID={mapping_id}",
                        request.remote_addr)

        return jsonify({'success': success, 'message': message})
    except Exception as e:
        print(f"❌ 更新客戶映射失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/mappings/<int:mapping_id>', methods=['DELETE'])
@admin_required
def api_delete_mapping(mapping_id):
    """刪除客戶映射（僅管理者）"""
    try:
        # 先取得映射資訊用於記錄
        mapping = admin_get_customer_mapping_by_id(mapping_id)

        success, message = admin_delete_customer_mapping(mapping_id)

        if success and mapping:
            # 記錄活動
            admin_user = get_current_user()
            log_activity(admin_user['id'], 'mapping_delete',
                        f"刪除客戶映射: {mapping['customer_name']} (company={mapping.get('company', 'N/A')})",
                        request.remote_addr)

        return jsonify({'success': success, 'message': message})
    except Exception as e:
        print(f"❌ 刪除客戶映射失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/mappings/<int:mapping_id>')
@admin_required
def api_get_mapping(mapping_id):
    """取得單一客戶映射（僅管理者）"""
    try:
        mapping = admin_get_customer_mapping_by_id(mapping_id)
        if not mapping:
            return jsonify({'success': False, 'message': '映射不存在'})

        # 處理日期格式
        if mapping.get('created_at'):
            mapping['created_at'] = mapping['created_at'].strftime('%Y-%m-%d %H:%M:%S')
        if mapping.get('updated_at'):
            mapping['updated_at'] = mapping['updated_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({'success': True, 'mapping': mapping})
    except Exception as e:
        print(f"❌ 取得客戶映射失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


# ==================== 規則管理 API ====================

@app.route('/rules')
@admin_required
def rules_view():
    """規則管理頁面"""
    user = get_current_user()
    return render_template('rules.html', user=user)


@app.route('/api/admin/rules')
@admin_required
def api_get_all_rules():
    """取得所有處理規則"""
    try:
        category = request.args.get('category')
        user_id = request.args.get('user_id', type=int)

        if user_id:
            if category:
                rules = get_processing_rules_by_category(category, user_id)
            else:
                rules = get_processing_rules_by_user(user_id)
        else:
            if category:
                rules = get_processing_rules_by_category(category)
            else:
                rules = get_all_processing_rules()

        # 處理日期格式
        for rule in rules:
            if rule.get('created_at'):
                rule['created_at'] = rule['created_at'].strftime('%Y-%m-%d %H:%M:%S')
            if rule.get('updated_at'):
                rule['updated_at'] = rule['updated_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({'success': True, 'rules': rules})
    except Exception as e:
        print(f"❌ 取得處理規則失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/rules/<int:rule_id>')
@admin_required
def api_get_rule(rule_id):
    """取得單一處理規則"""
    try:
        rule = get_processing_rule_by_id(rule_id)
        if not rule:
            return jsonify({'success': False, 'message': '規則不存在'})

        if rule.get('created_at'):
            rule['created_at'] = rule['created_at'].strftime('%Y-%m-%d %H:%M:%S')
        if rule.get('updated_at'):
            rule['updated_at'] = rule['updated_at'].strftime('%Y-%m-%d %H:%M:%S')

        return jsonify({'success': True, 'rule': rule})
    except Exception as e:
        print(f"❌ 取得處理規則失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/rules', methods=['POST'])
@admin_required
def api_create_rule():
    """新增處理規則"""
    try:
        data = request.get_json()
        user_id = data.get('user_id')
        rule_name = data.get('rule_name', '').strip()
        rule_category = data.get('rule_category', '').strip()
        rule_description = data.get('rule_description', '').strip()
        rule_config = data.get('rule_config')
        display_order = data.get('display_order', 0)

        if not user_id:
            return jsonify({'success': False, 'message': '用戶 ID 為必填'})

        if not rule_name:
            return jsonify({'success': False, 'message': '規則名稱為必填'})

        if rule_category not in ['erp', 'transit', 'forecast', 'mapping', 'cleanup']:
            return jsonify({'success': False, 'message': '無效的規則類別'})

        success, message, rule_id = create_processing_rule(
            user_id, rule_name, rule_category, rule_description, rule_config, display_order
        )

        if success:
            return jsonify({'success': True, 'message': message, 'rule_id': rule_id})
        return jsonify({'success': False, 'message': message})
    except Exception as e:
        print(f"❌ 新增處理規則失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/rules/<int:rule_id>', methods=['PUT'])
@admin_required
def api_update_rule(rule_id):
    """更新處理規則"""
    try:
        data = request.get_json()
        update_data = {}

        if 'rule_name' in data:
            update_data['rule_name'] = data['rule_name'].strip()
        if 'rule_description' in data:
            update_data['rule_description'] = data['rule_description'].strip()
        if 'rule_config' in data:
            update_data['rule_config'] = data['rule_config']
        if 'is_active' in data:
            update_data['is_active'] = data['is_active']
        if 'display_order' in data:
            update_data['display_order'] = data['display_order']

        success, message = update_processing_rule(rule_id, **update_data)

        return jsonify({'success': success, 'message': message})
    except Exception as e:
        print(f"❌ 更新處理規則失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/rules/<int:rule_id>', methods=['DELETE'])
@admin_required
def api_delete_rule(rule_id):
    """刪除處理規則"""
    try:
        success, message = delete_processing_rule(rule_id)
        return jsonify({'success': success, 'message': message})
    except Exception as e:
        print(f"❌ 刪除處理規則失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/admin/rules/<int:rule_id>/toggle', methods=['POST'])
@admin_required
def api_toggle_rule_status(rule_id):
    """切換規則啟用狀態"""
    try:
        success, message, new_status = toggle_processing_rule_status(rule_id)
        return jsonify({'success': success, 'message': message, 'is_active': new_status})
    except Exception as e:
        print(f"❌ 切換規則狀態失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/test/customers')
@it_or_admin_required
def api_get_test_customers():
    """取得可測試的客戶列表"""
    try:
        customers = get_users_with_company()
        return jsonify({'success': True, 'customers': customers})
    except Exception as e:
        print(f"❌ 取得客戶列表失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/test/run', methods=['POST'])
@it_or_admin_required
def api_run_test():
    """執行測試流程"""
    import shutil

    user = get_current_user()
    data = request.get_json()
    customer_id = data.get('customer_id')

    if not customer_id:
        return jsonify({'success': False, 'message': '請選擇客戶'})

    try:
        # 取得客戶資訊
        customer = get_user_by_id(customer_id)
        if not customer:
            return jsonify({'success': False, 'message': '客戶不存在'})

        company = customer['company']

        # 檢查測試檔案是否存在
        test_folder = os.path.join('test_data', company)
        if not os.path.exists(test_folder):
            return jsonify({'success': False, 'message': f'測試資料夾不存在: test_data/{company}'})

        required_files = ['erp_data.xlsx', 'forecast_data.xlsx', 'transit_data.xlsx']
        missing_files = []
        for f in required_files:
            if not os.path.exists(os.path.join(test_folder, f)):
                missing_files.append(f)

        if missing_files:
            return jsonify({
                'success': False,
                'message': f'缺少測試檔案: {", ".join(missing_files)}'
            })

        # 建立測試 session
        test_session_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S') + '_test'
        upload_folder = os.path.join(UPLOAD_FOLDER, str(customer_id), test_session_timestamp)
        processed_folder = os.path.join(PROCESSED_FOLDER, str(customer_id), test_session_timestamp)

        os.makedirs(upload_folder, exist_ok=True)
        os.makedirs(processed_folder, exist_ok=True)

        # 複製測試檔案
        for f in required_files:
            shutil.copy(
                os.path.join(test_folder, f),
                os.path.join(upload_folder, f)
            )

        # 取得客戶的 username 用於模板驗證
        customer_username = customer['username']

        # ========== 格式驗證（使用客戶專屬模板）==========
        validation_errors = []

        # 驗證 ERP 格式
        erp_test_file = os.path.join(upload_folder, 'erp_data.xlsx')
        is_valid, message, details = validate_erp_format(erp_test_file, customer_username)
        if not is_valid:
            validation_errors.append({
                'file': 'ERP',
                'message': message,
                'details': details
            })

        # 驗證 Forecast 格式
        forecast_test_file = os.path.join(upload_folder, 'forecast_data.xlsx')
        is_valid, message, details = validate_forecast_format(forecast_test_file, customer_username)
        if not is_valid:
            validation_errors.append({
                'file': 'Forecast',
                'message': message,
                'details': details
            })

        # 驗證在途格式
        transit_test_file = os.path.join(upload_folder, 'transit_data.xlsx')
        is_valid, message, details = validate_transit_format(transit_test_file, customer_username)
        if not is_valid:
            validation_errors.append({
                'file': '在途',
                'message': message,
                'details': details
            })

        # 如果有驗證錯誤，返回錯誤訊息
        if validation_errors:
            # 清理已複製的測試檔案
            shutil.rmtree(upload_folder, ignore_errors=True)
            shutil.rmtree(processed_folder, ignore_errors=True)

            error_messages = []
            for err in validation_errors:
                error_messages.append(f"{err['file']}: {err['message']}")
                if err['details']:
                    for detail in err['details'][:3]:
                        error_messages.append(f"  - {detail}")

            return jsonify({
                'success': False,
                'message': '測試檔案格式驗證失敗',
                'validation_errors': validation_errors,
                'details': error_messages
            })

        print(f"✅ 測試檔案格式驗證通過（使用模板: compare/{customer_username}/）")
        # ========== 格式驗證結束 ==========

        # 記錄測試開始
        log_activity(
            user_id=user['id'],
            username=user['username'],
            action_type='forecast_start',
            action_detail=f"執行測試：{company} (customer_id: {customer_id})",
            ip_address=get_client_ip(),
            user_agent=request.headers.get('User-Agent')
        )

        # 執行完整流程
        # 1. 數據清理
        from openpyxl import load_workbook as openpyxl_load_workbook
        forecast_file = os.path.join(upload_folder, 'forecast_data.xlsx')
        wb = openpyxl_load_workbook(forecast_file)
        ws = wb.active

        cleaned_count = 0
        for row_idx in range(1, ws.max_row + 1):
            k_cell = ws.cell(row=row_idx, column=11)
            if k_cell.value and str(k_cell.value) == "供應數量":
                for col_idx in range(12, min(50, ws.max_column + 1)):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value != 0:
                        cell.value = 0
                        cleaned_count += 1

            i_cell = ws.cell(row=row_idx, column=9)
            if i_cell.value and "庫存數量" in str(i_cell.value):
                next_row_i_cell = ws.cell(row=row_idx + 1, column=9)
                if next_row_i_cell.value != 0:
                    next_row_i_cell.value = 0
                    cleaned_count += 1

        cleaned_file = os.path.join(processed_folder, 'cleaned_forecast.xlsx')
        wb.save(cleaned_file)

        # 2. ERP 和 Transit 整合
        erp_file = os.path.join(upload_folder, 'erp_data.xlsx')
        transit_file = os.path.join(upload_folder, 'transit_data.xlsx')

        erp_df = pd.read_excel(erp_file)
        transit_df = pd.read_excel(transit_file)

        # 取得客戶的 mapping 資料
        mapping_data = get_customer_mappings(customer_id)
        if not mapping_data:
            mapping_data = {'regions': {}, 'schedule_breakpoints': {}, 'etd': {}, 'eta': {}}

        # 標準化日期格式
        if '排程出貨日期' in erp_df.columns:
            erp_df['排程出貨日期'] = erp_df['排程出貨日期'].apply(normalize_date_for_mapping)

        # 找到客戶簡稱欄位
        customer_col = None
        for col in erp_df.columns:
            if '客戶' in str(col) and '簡稱' in str(col):
                customer_col = col
                break

        if customer_col:
            erp_df['客戶需求地區'] = erp_df[customer_col].map(mapping_data.get('regions', {}))
            erp_df['排程出貨日期斷點'] = erp_df[customer_col].map(mapping_data.get('schedule_breakpoints', {}))
            erp_df['ETD'] = erp_df[customer_col].map(mapping_data.get('etd', {}))
            erp_df['ETA'] = erp_df[customer_col].map(mapping_data.get('eta', {}))

        if '排程出貨日期' in erp_df.columns:
            erp_df = erp_df.sort_values('排程出貨日期')

        integrated_erp_file = os.path.join(processed_folder, 'integrated_erp.xlsx')
        erp_df.to_excel(integrated_erp_file, index=False)

        # Transit 整合
        mapping_dict = {}
        all_customers = set()
        all_customers.update(mapping_data.get('regions', {}).keys())
        for cust in all_customers:
            mapping_dict[cust] = {
                'region': mapping_data.get('regions', {}).get(cust, ''),
                'schedule_breakpoint': mapping_data.get('schedule_breakpoints', {}).get(cust, ''),
                'etd': mapping_data.get('etd', {}).get(cust, ''),
                'eta': mapping_data.get('eta', {}).get(cust, '')
            }

        if len(transit_df.columns) >= 5:
            transit_customer_col = transit_df.columns[4]
            transit_df['客戶需求地區'] = transit_df[transit_customer_col].apply(
                lambda x: mapping_dict.get(str(x), {}).get('region', '') if pd.notna(x) else ''
            )
            transit_df['排程出貨日期斷點'] = transit_df[transit_customer_col].apply(
                lambda x: mapping_dict.get(str(x), {}).get('schedule_breakpoint', '') if pd.notna(x) else ''
            )
            transit_df['ETD'] = transit_df[transit_customer_col].apply(
                lambda x: mapping_dict.get(str(x), {}).get('etd', '') if pd.notna(x) else ''
            )
            transit_df['ETA_mapping'] = transit_df[transit_customer_col].apply(
                lambda x: mapping_dict.get(str(x), {}).get('eta', '') if pd.notna(x) else ''
            )

        integrated_transit_file = os.path.join(processed_folder, 'integrated_transit.xlsx')
        transit_df.to_excel(integrated_transit_file, index=False)

        # 3. 執行 FORECAST 處理
        from ultra_fast_forecast_processor import UltraFastForecastProcessor

        processor = UltraFastForecastProcessor(
            forecast_file=cleaned_file,
            erp_file=integrated_erp_file,
            transit_file=integrated_transit_file,
            output_folder=processed_folder
        )

        success = processor.process_all_blocks()

        if success:
            result_file = os.path.join(processed_folder, 'forecast_result.xlsx')
            if os.path.exists(result_file):
                # 記錄測試成功
                log_activity(
                    user_id=user['id'],
                    username=user['username'],
                    action_type='forecast_success',
                    action_detail=f"測試完成：{company}，ERP填入: {processor.total_filled}，Transit填入: {processor.total_transit_filled}",
                    ip_address=get_client_ip(),
                    user_agent=request.headers.get('User-Agent')
                )

                return jsonify({
                    'success': True,
                    'message': f'測試完成：{company}',
                    'result': {
                        'company': company,
                        'cleaned_cells': cleaned_count,
                        'erp_filled': processor.total_filled,
                        'erp_skipped': processor.total_skipped,
                        'transit_filled': processor.total_transit_filled,
                        'transit_skipped': processor.total_transit_skipped,
                        'result_folder': processed_folder,
                        'download_path': f'/api/files/download/{customer_id}/{test_session_timestamp}/forecast_result.xlsx'
                    }
                })

        return jsonify({'success': False, 'message': '測試執行失敗'})

    except Exception as e:
        print(f"❌ 測試執行失敗: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'測試執行失敗: {str(e)}'})


@app.route('/api/files/browse')
@it_or_admin_required
def api_browse_files():
    """瀏覽 uploads 或 processed 資料夾"""
    try:
        folder_type = request.args.get('type', 'processed')
        user_id = request.args.get('user_id')

        base_folder = UPLOAD_FOLDER if folder_type == 'uploads' else PROCESSED_FOLDER

        # 取得用戶 ID 到名稱的映射
        user_mapping = {}
        all_users = get_all_users()
        for u in all_users:
            user_mapping[str(u['id'])] = f"{u['display_name']} ({u['company']})"

        if user_id:
            base_folder = os.path.join(base_folder, str(user_id))

        if not os.path.exists(base_folder):
            return jsonify({'success': True, 'files': [], 'folders': [], 'user_mapping': user_mapping})

        files = []
        folders = []

        # 列出第一層目錄
        for item in os.listdir(base_folder):
            item_path = os.path.join(base_folder, item)
            if os.path.isdir(item_path):
                # 計算資料夾大小和檔案數
                file_count = 0
                total_size = 0
                for root, dirs, filenames in os.walk(item_path):
                    for f in filenames:
                        file_count += 1
                        total_size += os.path.getsize(os.path.join(root, f))

                folders.append({
                    'name': item,
                    'path': item if not user_id else f"{user_id}/{item}",
                    'file_count': file_count,
                    'total_size': total_size,
                    'modified': datetime.fromtimestamp(os.path.getmtime(item_path)).strftime('%Y-%m-%d %H:%M:%S')
                })
            else:
                files.append({
                    'name': item,
                    'path': item if not user_id else f"{user_id}/{item}",
                    'size': os.path.getsize(item_path),
                    'modified': datetime.fromtimestamp(os.path.getmtime(item_path)).strftime('%Y-%m-%d %H:%M:%S')
                })

        return jsonify({'success': True, 'files': files, 'folders': folders, 'user_mapping': user_mapping})
    except Exception as e:
        print(f"❌ 瀏覽檔案失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/files/browse/<path:folder_path>')
@it_or_admin_required
def api_browse_folder(folder_path):
    """瀏覽指定資料夾的內容"""
    try:
        folder_type = request.args.get('type', 'processed')
        base_folder = UPLOAD_FOLDER if folder_type == 'uploads' else PROCESSED_FOLDER

        # 取得用戶 ID 到名稱的映射
        user_mapping = {}
        all_users = get_all_users()
        for u in all_users:
            user_mapping[str(u['id'])] = f"{u['display_name']} ({u['company']})"

        # 安全性檢查
        if '..' in folder_path:
            return jsonify({'success': False, 'message': '無效的路徑'}), 400

        # 處理 Windows 路徑分隔符
        folder_path = folder_path.replace('\\', '/')
        target_folder = os.path.join(base_folder, folder_path)

        print(f"[檔案瀏覽] 目標路徑: {target_folder}")

        if not os.path.exists(target_folder) or not os.path.isdir(target_folder):
            print(f"[檔案瀏覽] 路徑不存在或不是目錄")
            return jsonify({'success': True, 'files': [], 'folders': [], 'user_mapping': user_mapping})

        files = []
        folders = []

        for item in os.listdir(target_folder):
            item_path = os.path.join(target_folder, item)
            # 使用正斜線統一路徑
            rel_path = f"{folder_path}/{item}"

            if os.path.isdir(item_path):
                # 計算子資料夾檔案數
                file_count = len([f for f in os.listdir(item_path) if os.path.isfile(os.path.join(item_path, f))])
                folders.append({
                    'name': item,
                    'path': rel_path,
                    'file_count': file_count,
                    'modified': datetime.fromtimestamp(os.path.getmtime(item_path)).strftime('%Y-%m-%d %H:%M:%S')
                })
            else:
                files.append({
                    'name': item,
                    'path': rel_path,
                    'size': os.path.getsize(item_path),
                    'modified': datetime.fromtimestamp(os.path.getmtime(item_path)).strftime('%Y-%m-%d %H:%M:%S')
                })

        print(f"[檔案瀏覽] 找到 {len(folders)} 個資料夾, {len(files)} 個檔案")
        return jsonify({'success': True, 'files': files, 'folders': folders, 'current_path': folder_path, 'user_mapping': user_mapping})
    except Exception as e:
        print(f"❌ 瀏覽資料夾失敗: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/files/download/<path:filepath>')
@it_or_admin_required
def api_download_admin_file(filepath):
    """下載指定檔案（IT/Admin 專用）"""
    user = get_current_user()

    # 安全性檢查
    if '..' in filepath or filepath.startswith('/'):
        return jsonify({'success': False, 'message': '無效的檔案路徑'}), 400

    folder_type = request.args.get('type', 'processed')
    base_folder = UPLOAD_FOLDER if folder_type == 'uploads' else PROCESSED_FOLDER

    full_path = os.path.join(base_folder, filepath)

    if os.path.exists(full_path) and os.path.isfile(full_path):
        # 記錄下載活動
        file_size = os.path.getsize(full_path)
        log_activity(
            user_id=user['id'],
            username=user['username'],
            action_type='download',
            action_detail=f"[管理] 下載文件：{filepath} ({file_size} bytes)",
            ip_address=get_client_ip(),
            user_agent=request.headers.get('User-Agent')
        )
        return send_file(full_path, as_attachment=True)

    return jsonify({'success': False, 'message': '檔案不存在'}), 404


@app.route('/api/users/list')
@it_or_admin_required
def api_get_users_list():
    """取得用戶列表（用於篩選下拉選單）"""
    try:
        users = get_all_users()
        user_list = [{'id': u['id'], 'display_name': u['display_name'], 'company': u['company']} for u in users]
        return jsonify({'success': True, 'users': user_list})
    except Exception as e:
        print(f"❌ 取得用戶列表失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/customers/login-options')
def api_get_login_customers():
    """取得登入頁面的客戶選項列表（公開 API，不需登入）"""
    try:
        users = get_users_with_company()
        # 只返回必要的資訊：username（用於登入）、display_name（顯示名稱）、company（公司）
        customer_list = [
            {
                'username': u['username'],
                'display_name': u['display_name'],
                'company': u['company'] or u['display_name']
            }
            for u in users
        ]
        return jsonify({'success': True, 'customers': customer_list})
    except Exception as e:
        print(f"❌ 取得登入客戶列表失敗: {e}")
        return jsonify({'success': False, 'message': str(e)})


if __name__ == '__main__':
    # 初始化資料庫
    print("正在初始化資料庫...")
    if init_database():
        print("✅ 資料庫初始化成功")
        # 更新 activity_logs ENUM（新增用戶管理操作類型）
        print("正在更新 activity_logs ENUM...")
        update_activity_logs_enum()
        # 建立預設帳號
        print("正在建立預設帳號...")
        if create_default_users():
            print("✅ 預設帳號建立成功")
        else:
            print("⚠️ 預設帳號建立失敗或已存在")
    else:
        print("❌ 資料庫初始化失敗")

    # 啟動時清理超過保留期限的資料夾
    print(f"正在清理超過 {FILE_RETENTION_DAYS} 天的資料夾...")
    cleanup_old_folders()

    app.run(debug=True, host='0.0.0.0', port=12026)
