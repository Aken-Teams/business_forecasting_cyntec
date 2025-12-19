# -*- coding: utf-8 -*-
"""
資料庫連線和操作模組
"""

import pymysql
from pymysql.cursors import DictCursor
from datetime import datetime
import hashlib
import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

# 資料庫連線設定（從環境變數讀取）
DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 3306)),
    'database': os.getenv('DB_NAME', 'database'),
    'user': os.getenv('DB_USER', 'root'),
    'password': os.getenv('DB_PASSWORD', ''),
    'charset': 'utf8mb4',
    'cursorclass': DictCursor
}

def get_db_connection():
    """建立資料庫連線"""
    try:
        connection = pymysql.connect(**DB_CONFIG)
        return connection
    except Exception as e:
        print(f"❌ 資料庫連線失敗: {e}")
        return None

def init_database():
    """初始化資料庫表格"""
    connection = get_db_connection()
    if not connection:
        print("❌ 無法初始化資料庫")
        return False

    try:
        with connection.cursor() as cursor:
            # 建立用戶表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(50) UNIQUE NOT NULL,
                    password_hash VARCHAR(128) NOT NULL,
                    display_name VARCHAR(100) NOT NULL,
                    role ENUM('admin', 'it', 'user') NOT NULL DEFAULT 'user',
                    company VARCHAR(100),
                    is_active BOOLEAN DEFAULT TRUE,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    last_login DATETIME
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 建立操作日誌表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS activity_logs (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    user_id INT,
                    username VARCHAR(50),
                    action_type ENUM(
                        'login', 'logout', 'login_failed',
                        'upload_erp', 'upload_forecast', 'upload_transit',
                        'upload_erp_failed', 'upload_forecast_failed', 'upload_transit_failed',
                        'cleanup_start', 'cleanup_success', 'cleanup_failed',
                        'mapping_start', 'mapping_success', 'mapping_failed',
                        'forecast_start', 'forecast_success', 'forecast_failed',
                        'download',
                        'mapping_config_save', 'mapping_config_failed',
                        'user_create', 'user_update', 'user_delete', 'user_toggle_status'
                    ) NOT NULL,
                    action_detail TEXT,
                    ip_address VARCHAR(45),
                    user_agent TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 建立檔案上傳記錄表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS upload_records (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    user_id INT,
                    file_type ENUM('erp', 'forecast', 'transit') NOT NULL,
                    original_filename VARCHAR(255),
                    file_size BIGINT,
                    row_count INT,
                    column_count INT,
                    upload_status ENUM('success', 'failed', 'validation_failed') NOT NULL,
                    error_message TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 建立處理記錄表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS process_records (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    user_id INT,
                    process_type ENUM('cleanup', 'mapping', 'forecast') NOT NULL,
                    process_status ENUM('started', 'success', 'failed') NOT NULL,
                    process_detail TEXT,
                    duration_seconds FLOAT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 建立客戶映射表（每個帳號有獨立的 mapping 資料）
            # 唯一值：user_id + customer_name + region（客戶簡稱 + 客戶廠區）
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS customer_mappings (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    user_id INT NOT NULL,
                    customer_name VARCHAR(100) NOT NULL,
                    delivery_location VARCHAR(100),
                    region VARCHAR(50) NOT NULL,
                    schedule_breakpoint VARCHAR(50),
                    etd VARCHAR(50),
                    eta VARCHAR(50),
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
                    UNIQUE KEY unique_user_customer_region (user_id, customer_name, region)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 確保 delivery_location 欄位存在（用於舊資料庫升級）
            try:
                cursor.execute("""
                    ALTER TABLE customer_mappings
                    ADD COLUMN delivery_location VARCHAR(100) AFTER customer_name
                """)
                print("✅ 已新增 delivery_location 欄位")
            except Exception as alter_error:
                # 欄位已存在則忽略
                if "Duplicate column" not in str(alter_error):
                    print(f"ℹ️ delivery_location 欄位檢查: {alter_error}")

            # 更新唯一索引（從 user_id + customer_name 改為 user_id + customer_name + region）
            try:
                # 先刪除舊的唯一索引
                cursor.execute("ALTER TABLE customer_mappings DROP INDEX unique_user_customer")
                print("✅ 已刪除舊的唯一索引 unique_user_customer")
            except Exception as drop_error:
                pass  # 索引可能不存在

            try:
                # 建立新的唯一索引
                cursor.execute("""
                    ALTER TABLE customer_mappings
                    ADD UNIQUE KEY unique_user_customer_region (user_id, customer_name, region)
                """)
                print("✅ 已建立新的唯一索引 unique_user_customer_region")
            except Exception as add_error:
                if "Duplicate key name" not in str(add_error):
                    print(f"ℹ️ 唯一索引檢查: {add_error}")

            connection.commit()
            print("✅ 資料庫表格初始化完成")
            return True

    except Exception as e:
        print(f"❌ 資料庫初始化失敗: {e}")
        return False
    finally:
        connection.close()


def update_activity_logs_enum():
    """更新 activity_logs 表的 action_type ENUM，新增用戶管理和映射管理相關操作"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                ALTER TABLE activity_logs
                MODIFY COLUMN action_type ENUM(
                    'login', 'logout', 'login_failed',
                    'upload_erp', 'upload_forecast', 'upload_transit',
                    'upload_erp_failed', 'upload_forecast_failed', 'upload_transit_failed',
                    'cleanup_start', 'cleanup_success', 'cleanup_failed',
                    'mapping_start', 'mapping_success', 'mapping_failed',
                    'forecast_start', 'forecast_success', 'forecast_failed',
                    'download',
                    'mapping_config_save', 'mapping_config_failed',
                    'user_create', 'user_update', 'user_delete', 'user_toggle_status',
                    'mapping_create', 'mapping_update', 'mapping_delete'
                ) NOT NULL
            """)
            connection.commit()
            print("✅ activity_logs ENUM 更新完成")
            return True
    except Exception as e:
        print(f"❌ 更新 activity_logs ENUM 失敗: {e}")
        return False
    finally:
        connection.close()


def hash_password(password):
    """密碼雜湊"""
    salt = os.getenv('PASSWORD_SALT', 'default_salt')
    return hashlib.sha256((password + salt).encode()).hexdigest()

def create_default_users():
    """建立預設帳號"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            # 預設帳號列表
            default_users = [
                {
                    'username': 'admin',
                    'password': 'admin123',
                    'display_name': '系統管理員',
                    'role': 'admin',
                    'company': '智合科技'
                },
                {
                    'username': 'it_user',
                    'password': 'it123456',
                    'display_name': 'IT 管理人員',
                    'role': 'it',
                    'company': '智合科技'
                },
                {
                    'username': 'quanta',
                    'password': 'quanta123',
                    'display_name': '廣達用戶',
                    'role': 'user',
                    'company': '廣達電腦'
                }
            ]

            for user in default_users:
                # 檢查用戶是否已存在
                cursor.execute("SELECT id FROM users WHERE username = %s", (user['username'],))
                if cursor.fetchone():
                    print(f"⚠️ 用戶 {user['username']} 已存在，跳過")
                    continue

                # 建立新用戶
                cursor.execute("""
                    INSERT INTO users (username, password_hash, display_name, role, company)
                    VALUES (%s, %s, %s, %s, %s)
                """, (
                    user['username'],
                    hash_password(user['password']),
                    user['display_name'],
                    user['role'],
                    user['company']
                ))
                print(f"✅ 已建立用戶: {user['username']} ({user['display_name']})")

            connection.commit()
            return True

    except Exception as e:
        print(f"❌ 建立預設帳號失敗: {e}")
        return False
    finally:
        connection.close()

def verify_user(username, password):
    """驗證用戶登入"""
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, username, display_name, role, company, is_active
                FROM users
                WHERE username = %s AND password_hash = %s
            """, (username, hash_password(password)))

            user = cursor.fetchone()

            if user and user['is_active']:
                # 更新最後登入時間
                cursor.execute("""
                    UPDATE users SET last_login = %s WHERE id = %s
                """, (datetime.now(), user['id']))
                connection.commit()
                return user

            return None

    except Exception as e:
        print(f"❌ 用戶驗證失敗: {e}")
        return None
    finally:
        connection.close()

def log_activity(user_id, username, action_type, action_detail=None, ip_address=None, user_agent=None):
    """記錄用戶活動"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO activity_logs (user_id, username, action_type, action_detail, ip_address, user_agent)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (user_id, username, action_type, action_detail, ip_address, user_agent))
            connection.commit()
            return True
    except Exception as e:
        print(f"❌ 記錄活動失敗: {e}")
        return False
    finally:
        connection.close()

def log_upload(user_id, file_type, original_filename, file_size, row_count, column_count, status, error_message=None):
    """記錄檔案上傳"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO upload_records (user_id, file_type, original_filename, file_size, row_count, column_count, upload_status, error_message)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (user_id, file_type, original_filename, file_size, row_count, column_count, status, error_message))
            connection.commit()
            return True
    except Exception as e:
        print(f"❌ 記錄上傳失敗: {e}")
        return False
    finally:
        connection.close()

def log_process(user_id, process_type, process_status, process_detail=None, duration_seconds=None):
    """記錄處理流程"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO process_records (user_id, process_type, process_status, process_detail, duration_seconds)
                VALUES (%s, %s, %s, %s, %s)
            """, (user_id, process_type, process_status, process_detail, duration_seconds))
            connection.commit()
            return True
    except Exception as e:
        print(f"❌ 記錄處理失敗: {e}")
        return False
    finally:
        connection.close()

def get_user_by_id(user_id):
    """根據 ID 取得用戶資料"""
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, username, display_name, role, company, is_active, last_login
                FROM users WHERE id = %s
            """, (user_id,))
            return cursor.fetchone()
    except Exception as e:
        print(f"❌ 取得用戶失敗: {e}")
        return None
    finally:
        connection.close()

def get_all_users():
    """取得所有用戶（管理員功能）"""
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, username, display_name, role, company, is_active, created_at, last_login
                FROM users ORDER BY created_at DESC
            """)
            return cursor.fetchall()
    except Exception as e:
        print(f"❌ 取得用戶列表失敗: {e}")
        return []
    finally:
        connection.close()

def get_activity_logs(user_id=None, limit=100):
    """取得活動日誌"""
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            if user_id:
                cursor.execute("""
                    SELECT * FROM activity_logs
                    WHERE user_id = %s
                    ORDER BY created_at DESC
                    LIMIT %s
                """, (user_id, limit))
            else:
                cursor.execute("""
                    SELECT * FROM activity_logs
                    ORDER BY created_at DESC
                    LIMIT %s
                """, (limit,))
            return cursor.fetchall()
    except Exception as e:
        print(f"❌ 取得活動日誌失敗: {e}")
        return []
    finally:
        connection.close()

# ========================================
# 客戶映射資料操作函數
# ========================================

def get_customer_mappings(user_id):
    """
    取得指定用戶的所有客戶映射資料
    返回格式與原 mapping_data.json 相同
    """
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT customer_name, delivery_location, region, schedule_breakpoint, etd, eta
                FROM customer_mappings
                WHERE user_id = %s
            """, (user_id,))
            rows = cursor.fetchall()

            # 轉換為原有的 JSON 格式
            mapping_data = {
                'delivery_locations': {},
                'regions': {},
                'schedule_breakpoints': {},
                'etd': {},
                'eta': {}
            }

            for row in rows:
                customer = row['customer_name']
                if row['delivery_location']:
                    mapping_data['delivery_locations'][customer] = row['delivery_location']
                if row['region']:
                    mapping_data['regions'][customer] = row['region']
                if row['schedule_breakpoint']:
                    mapping_data['schedule_breakpoints'][customer] = row['schedule_breakpoint']
                if row['etd']:
                    mapping_data['etd'][customer] = row['etd']
                if row['eta']:
                    mapping_data['eta'][customer] = row['eta']

            return mapping_data

    except Exception as e:
        print(f"❌ 取得客戶映射資料失敗: {e}")
        return None
    finally:
        connection.close()

def save_customer_mappings(user_id, mapping_data):
    """
    儲存用戶的客戶映射資料（完整替換模式）
    會先刪除用戶的所有現有 mapping 資料，再插入新資料
    這樣當客戶被刪除或重命名時，舊資料會被清除

    mapping_data 格式：
    {
        'delivery_locations': {'客戶名稱': '送貨地點', ...},
        'regions': {'客戶名稱': '客戶廠區', ...},
        'schedule_breakpoints': {'客戶名稱': '斷點', ...},
        'etd': {'客戶名稱': 'ETD', ...},
        'eta': {'客戶名稱': 'ETA', ...}
    }
    """
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            # 先刪除該用戶的所有現有映射資料
            cursor.execute("""
                DELETE FROM customer_mappings WHERE user_id = %s
            """, (user_id,))
            print(f"🔄 已清除用戶 {user_id} 的舊映射資料")

            # 收集所有客戶名稱
            all_customers = set()
            all_customers.update(mapping_data.get('delivery_locations', {}).keys())
            all_customers.update(mapping_data.get('regions', {}).keys())
            all_customers.update(mapping_data.get('schedule_breakpoints', {}).keys())
            all_customers.update(mapping_data.get('etd', {}).keys())
            all_customers.update(mapping_data.get('eta', {}).keys())

            # 插入新的映射資料
            for customer in all_customers:
                delivery_location = mapping_data.get('delivery_locations', {}).get(customer, '')
                region = mapping_data.get('regions', {}).get(customer, '')
                schedule_breakpoint = mapping_data.get('schedule_breakpoints', {}).get(customer, '')
                etd = mapping_data.get('etd', {}).get(customer, '')
                eta = mapping_data.get('eta', {}).get(customer, '')

                cursor.execute("""
                    INSERT INTO customer_mappings
                    (user_id, customer_name, delivery_location, region, schedule_breakpoint, etd, eta)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (user_id, customer, delivery_location, region, schedule_breakpoint, etd, eta))

            connection.commit()
            print(f"✅ 已儲存 {len(all_customers)} 個客戶的映射資料 (user_id: {user_id})")
            return True

    except Exception as e:
        print(f"❌ 儲存客戶映射資料失敗: {e}")
        connection.rollback()
        return False
    finally:
        connection.close()

def delete_customer_mapping(user_id, customer_name):
    """刪除指定用戶的特定客戶映射"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                DELETE FROM customer_mappings
                WHERE user_id = %s AND customer_name = %s
            """, (user_id, customer_name))
            connection.commit()
            return cursor.rowcount > 0

    except Exception as e:
        print(f"❌ 刪除客戶映射失敗: {e}")
        return False
    finally:
        connection.close()

def delete_all_customer_mappings(user_id):
    """刪除指定用戶的所有客戶映射"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                DELETE FROM customer_mappings
                WHERE user_id = %s
            """, (user_id,))
            connection.commit()
            print(f"✅ 已刪除用戶 {user_id} 的所有映射資料")
            return True

    except Exception as e:
        print(f"❌ 刪除所有客戶映射失敗: {e}")
        return False
    finally:
        connection.close()

def get_customer_mapping_list(user_id):
    """
    取得用戶的客戶映射列表（用於前端顯示）
    返回格式：[{'customer_name': '...', 'delivery_location': '...', 'region': '...', ...}, ...]
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT customer_name, delivery_location, region, schedule_breakpoint, etd, eta, updated_at
                FROM customer_mappings
                WHERE user_id = %s
                ORDER BY customer_name
            """, (user_id,))
            return cursor.fetchall()

    except Exception as e:
        print(f"❌ 取得客戶映射列表失敗: {e}")
        return []
    finally:
        connection.close()

def has_customer_mappings(user_id):
    """檢查用戶是否有任何映射資料"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT COUNT(*) as count FROM customer_mappings WHERE user_id = %s
            """, (user_id,))
            result = cursor.fetchone()
            return result['count'] > 0

    except Exception as e:
        print(f"❌ 檢查客戶映射失敗: {e}")
        return False
    finally:
        connection.close()

# ========================================
# IT/Admin 管理介面查詢函數
# ========================================

def get_upload_records(user_id=None, file_type=None, status=None,
                       start_date=None, end_date=None, limit=100, offset=0):
    """
    取得上傳記錄，支援多種篩選條件

    參數:
        user_id: 用戶ID（可選）
        file_type: 檔案類型 'erp'/'forecast'/'transit'（可選）
        status: 上傳狀態 'success'/'failed'/'validation_failed'（可選）
        start_date: 開始日期（可選）
        end_date: 結束日期（可選）
        limit: 返回記錄數（預設100）
        offset: 跳過的記錄數（分頁用）

    返回: (records, total) - 記錄列表和總數
    """
    connection = get_db_connection()
    if not connection:
        return [], 0

    try:
        with connection.cursor() as cursor:
            # 構建查詢條件
            conditions = []
            params = []

            if user_id:
                conditions.append("ur.user_id = %s")
                params.append(user_id)
            if file_type:
                conditions.append("ur.file_type = %s")
                params.append(file_type)
            if status:
                conditions.append("ur.upload_status = %s")
                params.append(status)
            if start_date:
                conditions.append("ur.created_at >= %s")
                params.append(start_date)
            if end_date:
                conditions.append("ur.created_at <= %s")
                params.append(end_date + " 23:59:59")

            where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""

            # 獲取總數
            count_sql = f"""
                SELECT COUNT(*) as total FROM upload_records ur
                LEFT JOIN users u ON ur.user_id = u.id
                {where_clause}
            """
            cursor.execute(count_sql, params)
            total = cursor.fetchone()['total']

            # 獲取數據
            sql = f"""
                SELECT ur.*, u.username, u.display_name, u.company
                FROM upload_records ur
                LEFT JOIN users u ON ur.user_id = u.id
                {where_clause}
                ORDER BY ur.created_at DESC
                LIMIT %s OFFSET %s
            """
            cursor.execute(sql, params + [limit, offset])
            records = cursor.fetchall()

            return records, total
    except Exception as e:
        print(f"❌ 取得上傳記錄失敗: {e}")
        return [], 0
    finally:
        connection.close()


def get_process_records(user_id=None, process_type=None, status=None,
                        start_date=None, end_date=None, limit=100, offset=0):
    """
    取得處理記錄，支援多種篩選條件

    參數:
        user_id: 用戶ID（可選）
        process_type: 處理類型 'cleanup'/'mapping'/'forecast'（可選）
        status: 處理狀態 'started'/'success'/'failed'（可選）
        start_date: 開始日期（可選）
        end_date: 結束日期（可選）
        limit: 返回記錄數（預設100）
        offset: 跳過的記錄數（分頁用）

    返回: (records, total) - 記錄列表和總數
    """
    connection = get_db_connection()
    if not connection:
        return [], 0

    try:
        with connection.cursor() as cursor:
            conditions = []
            params = []

            if user_id:
                conditions.append("pr.user_id = %s")
                params.append(user_id)
            if process_type:
                conditions.append("pr.process_type = %s")
                params.append(process_type)
            if status:
                conditions.append("pr.process_status = %s")
                params.append(status)
            if start_date:
                conditions.append("pr.created_at >= %s")
                params.append(start_date)
            if end_date:
                conditions.append("pr.created_at <= %s")
                params.append(end_date + " 23:59:59")

            where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""

            # 獲取總數
            count_sql = f"""
                SELECT COUNT(*) as total FROM process_records pr
                LEFT JOIN users u ON pr.user_id = u.id
                {where_clause}
            """
            cursor.execute(count_sql, params)
            total = cursor.fetchone()['total']

            # 獲取數據
            sql = f"""
                SELECT pr.*, u.username, u.display_name, u.company
                FROM process_records pr
                LEFT JOIN users u ON pr.user_id = u.id
                {where_clause}
                ORDER BY pr.created_at DESC
                LIMIT %s OFFSET %s
            """
            cursor.execute(sql, params + [limit, offset])
            records = cursor.fetchall()

            return records, total
    except Exception as e:
        print(f"❌ 取得處理記錄失敗: {e}")
        return [], 0
    finally:
        connection.close()


def get_activity_logs_filtered(user_id=None, action_type=None,
                               start_date=None, end_date=None,
                               limit=100, offset=0):
    """
    取得活動日誌，支援完整篩選

    參數:
        user_id: 用戶ID（可選）
        action_type: 操作類型（可選）
        start_date: 開始日期（可選）
        end_date: 結束日期（可選）
        limit: 返回記錄數（預設100）
        offset: 跳過的記錄數（分頁用）

    返回: (records, total) - 記錄列表和總數
    """
    connection = get_db_connection()
    if not connection:
        return [], 0

    try:
        with connection.cursor() as cursor:
            conditions = []
            params = []

            if user_id:
                conditions.append("al.user_id = %s")
                params.append(user_id)
            if action_type:
                conditions.append("al.action_type = %s")
                params.append(action_type)
            if start_date:
                conditions.append("al.created_at >= %s")
                params.append(start_date)
            if end_date:
                conditions.append("al.created_at <= %s")
                params.append(end_date + " 23:59:59")

            where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""

            # 獲取總數
            count_sql = f"""
                SELECT COUNT(*) as total FROM activity_logs al
                LEFT JOIN users u ON al.user_id = u.id
                {where_clause}
            """
            cursor.execute(count_sql, params)
            total = cursor.fetchone()['total']

            # 獲取數據
            sql = f"""
                SELECT al.*, u.display_name, u.company
                FROM activity_logs al
                LEFT JOIN users u ON al.user_id = u.id
                {where_clause}
                ORDER BY al.created_at DESC
                LIMIT %s OFFSET %s
            """
            cursor.execute(sql, params + [limit, offset])
            records = cursor.fetchall()

            return records, total
    except Exception as e:
        print(f"❌ 取得活動日誌失敗: {e}")
        return [], 0
    finally:
        connection.close()


def get_all_customer_mappings():
    """
    取得所有用戶的客戶映射資料（管理者功能）

    返回: 映射記錄列表，包含用戶資訊
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT cm.*, u.username, u.display_name, u.company
                FROM customer_mappings cm
                LEFT JOIN users u ON cm.user_id = u.id
                ORDER BY u.company, cm.customer_name
            """)
            return cursor.fetchall()
    except Exception as e:
        print(f"❌ 取得所有客戶映射失敗: {e}")
        return []
    finally:
        connection.close()


def get_users_with_company():
    """
    取得所有用戶及其公司資訊（用於測試功能的客戶選擇）
    僅返回一般用戶（role='user'）且啟用狀態的帳號

    返回: 用戶列表
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, username, display_name, company, role, is_active
                FROM users
                WHERE role = 'user' AND is_active = TRUE
                ORDER BY company, display_name
            """)
            return cursor.fetchall()
    except Exception as e:
        print(f"❌ 取得用戶列表失敗: {e}")
        return []
    finally:
        connection.close()


# ==================== 管理者客戶映射 CRUD ====================

def admin_create_customer_mapping(user_id, customer_name, delivery_location=None, region=None, schedule_breakpoint=None, etd=None, eta=None):
    """
    管理者新增客戶映射

    參數:
        user_id: 用戶ID
        customer_name: 客戶簡稱（必填）
        delivery_location: 送貨地點
        region: 客戶廠區（必填）
        schedule_breakpoint: 排程出貨日期斷點
        etd: ETD
        eta: ETA

    返回: (success, message, mapping_id)

    唯一值：user_id + customer_name + region
    """
    # 驗證必填欄位
    if not region:
        return False, "客戶廠區為必填欄位", None

    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗", None

    try:
        with connection.cursor() as cursor:
            # 檢查是否已存在相同的 user_id + customer_name + region
            cursor.execute("""
                SELECT id FROM customer_mappings
                WHERE user_id = %s AND customer_name = %s AND region = %s
            """, (user_id, customer_name, region))

            if cursor.fetchone():
                return False, f"該用戶已存在相同客戶簡稱+客戶廠區的映射（{customer_name} + {region}）", None

            # 新增映射
            cursor.execute("""
                INSERT INTO customer_mappings
                (user_id, customer_name, delivery_location, region, schedule_breakpoint, etd, eta)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (user_id, customer_name, delivery_location, region, schedule_breakpoint, etd, eta))

            connection.commit()
            mapping_id = cursor.lastrowid
            return True, "映射新增成功", mapping_id

    except Exception as e:
        print(f"❌ 新增客戶映射失敗: {e}")
        return False, str(e), None
    finally:
        connection.close()


def admin_update_customer_mapping(mapping_id, **kwargs):
    """
    管理者更新客戶映射

    參數:
        mapping_id: 映射ID
        **kwargs: 可更新的欄位 (customer_name, delivery_location, region, schedule_breakpoint, etd, eta)

    返回: (success, message)

    唯一值：user_id + customer_name + region
    """
    # 驗證 region 不能為空（如果有提供的話）
    if 'region' in kwargs and not kwargs['region']:
        return False, "客戶廠區為必填欄位，不能設為空值"

    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    allowed_fields = ['customer_name', 'delivery_location', 'region', 'schedule_breakpoint', 'etd', 'eta']
    update_fields = []
    values = []

    for field in allowed_fields:
        if field in kwargs:
            update_fields.append(f"{field} = %s")
            values.append(kwargs[field])

    if not update_fields:
        return False, "沒有提供要更新的欄位"

    values.append(mapping_id)

    try:
        with connection.cursor() as cursor:
            # 檢查映射是否存在，並取得完整資料
            cursor.execute("SELECT id, user_id, customer_name, region FROM customer_mappings WHERE id = %s", (mapping_id,))
            existing = cursor.fetchone()
            if not existing:
                return False, "映射不存在"

            # 如果要更新 customer_name 或 region，檢查是否會與其他記錄重複
            new_customer_name = kwargs.get('customer_name', existing['customer_name'])
            new_region = kwargs.get('region', existing['region'])

            if 'customer_name' in kwargs or 'region' in kwargs:
                cursor.execute("""
                    SELECT id FROM customer_mappings
                    WHERE user_id = %s AND customer_name = %s AND region = %s AND id != %s
                """, (existing['user_id'], new_customer_name, new_region, mapping_id))
                if cursor.fetchone():
                    return False, f"該用戶已存在相同客戶簡稱+客戶廠區的映射（{new_customer_name} + {new_region}）"

            # 執行更新
            sql = f"UPDATE customer_mappings SET {', '.join(update_fields)} WHERE id = %s"
            cursor.execute(sql, values)
            connection.commit()

            return True, "映射更新成功"

    except Exception as e:
        print(f"❌ 更新客戶映射失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def admin_delete_customer_mapping(mapping_id):
    """
    管理者刪除客戶映射

    參數:
        mapping_id: 映射ID

    返回: (success, message)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    try:
        with connection.cursor() as cursor:
            # 檢查映射是否存在
            cursor.execute("SELECT id, customer_name FROM customer_mappings WHERE id = %s", (mapping_id,))
            existing = cursor.fetchone()
            if not existing:
                return False, "映射不存在"

            # 刪除映射
            cursor.execute("DELETE FROM customer_mappings WHERE id = %s", (mapping_id,))
            connection.commit()

            return True, f"映射「{existing['customer_name']}」已刪除"

    except Exception as e:
        print(f"❌ 刪除客戶映射失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def admin_get_customer_mapping_by_id(mapping_id):
    """
    管理者根據 ID 取得單一客戶映射

    參數:
        mapping_id: 映射ID

    返回: 映射資料或 None
    """
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT cm.*, u.username, u.display_name, u.company
                FROM customer_mappings cm
                LEFT JOIN users u ON cm.user_id = u.id
                WHERE cm.id = %s
            """, (mapping_id,))
            return cursor.fetchone()
    except Exception as e:
        print(f"❌ 取得客戶映射失敗: {e}")
        return None
    finally:
        connection.close()


def create_user(username, password, display_name, role='user', company=None, is_active=True):
    """
    建立新用戶

    參數:
        username: 用戶名
        password: 密碼（明文，會自動加密）
        display_name: 顯示名稱
        role: 角色 (admin/it/user)
        company: 公司名稱
        is_active: 是否啟用

    返回: (success, message, user_id)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗", None

    try:
        with connection.cursor() as cursor:
            # 檢查用戶名是否已存在
            cursor.execute("SELECT id FROM users WHERE username = %s", (username,))
            if cursor.fetchone():
                return False, "用戶名已存在", None

            # 建立用戶
            cursor.execute("""
                INSERT INTO users (username, password_hash, display_name, role, company, is_active)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (username, hash_password(password), display_name, role, company, is_active))
            connection.commit()

            user_id = cursor.lastrowid
            return True, "用戶建立成功", user_id
    except Exception as e:
        print(f"❌ 建立用戶失敗: {e}")
        return False, str(e), None
    finally:
        connection.close()


def update_user(user_id, **kwargs):
    """
    更新用戶資料

    參數:
        user_id: 用戶 ID
        **kwargs: 要更新的欄位 (username, password, display_name, role, company, is_active)

    返回: (success, message)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    try:
        with connection.cursor() as cursor:
            # 檢查用戶是否存在
            cursor.execute("SELECT id, username FROM users WHERE id = %s", (user_id,))
            user = cursor.fetchone()
            if not user:
                return False, "用戶不存在"

            # 如果要更新用戶名，檢查是否重複
            if 'username' in kwargs and kwargs['username'] != user['username']:
                cursor.execute("SELECT id FROM users WHERE username = %s AND id != %s",
                             (kwargs['username'], user_id))
                if cursor.fetchone():
                    return False, "用戶名已被使用"

            # 建立更新語句
            update_fields = []
            update_values = []

            allowed_fields = ['username', 'display_name', 'role', 'company', 'is_active']
            for field in allowed_fields:
                if field in kwargs:
                    update_fields.append(f"{field} = %s")
                    update_values.append(kwargs[field])

            # 處理密碼
            if 'password' in kwargs and kwargs['password']:
                update_fields.append("password_hash = %s")
                update_values.append(hash_password(kwargs['password']))

            if not update_fields:
                return False, "沒有要更新的欄位"

            update_values.append(user_id)
            sql = f"UPDATE users SET {', '.join(update_fields)} WHERE id = %s"
            cursor.execute(sql, tuple(update_values))
            connection.commit()

            return True, "用戶更新成功"
    except Exception as e:
        print(f"❌ 更新用戶失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def delete_user(user_id):
    """
    刪除用戶

    參數:
        user_id: 用戶 ID

    返回: (success, message, deleted_username)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗", None

    try:
        with connection.cursor() as cursor:
            # 取得用戶資訊
            cursor.execute("SELECT username, display_name FROM users WHERE id = %s", (user_id,))
            user = cursor.fetchone()
            if not user:
                return False, "用戶不存在", None

            # 刪除用戶
            cursor.execute("DELETE FROM users WHERE id = %s", (user_id,))
            connection.commit()

            return True, "用戶刪除成功", user['username']
    except Exception as e:
        print(f"❌ 刪除用戶失敗: {e}")
        return False, str(e), None
    finally:
        connection.close()


# 測試連線和初始化
if __name__ == "__main__":
    print("測試資料庫連線...")
    conn = get_db_connection()
    if conn:
        print("✅ 資料庫連線成功")
        conn.close()

        print("\n初始化資料庫表格...")
        init_database()

        print("\n更新 activity_logs ENUM...")
        update_activity_logs_enum()

        print("\n建立預設帳號...")
        create_default_users()
    else:
        print("❌ 資料庫連線失敗")
