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
                    original_filename VARCHAR(2000),
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

            # 建立處理規則表（儲存預測處理的計算規則，依客戶區分）
            # 類別順序：upload -> cleanup -> mapping -> erp -> transit -> forecast -> output
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS processing_rules (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    user_id INT NOT NULL,
                    rule_name VARCHAR(100) NOT NULL,
                    rule_category ENUM('upload', 'cleanup', 'mapping', 'erp', 'transit', 'forecast', 'output') NOT NULL,
                    rule_description TEXT,
                    rule_config JSON,
                    is_active BOOLEAN DEFAULT TRUE,
                    display_order INT DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)

            # 確保 user_id 欄位存在（用於舊資料庫升級）
            try:
                # 檢查 user_id 欄位是否存在
                cursor.execute("""
                    SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = DATABASE()
                    AND TABLE_NAME = 'processing_rules'
                    AND COLUMN_NAME = 'user_id'
                """)
                if not cursor.fetchone():
                    # 欄位不存在，需要升級
                    # 先取得廣達用戶的 ID 作為預設值
                    cursor.execute("SELECT id FROM users WHERE username = 'quanta'")
                    quanta_user = cursor.fetchone()
                    default_user_id = quanta_user['id'] if quanta_user else 1

                    # 加入 user_id 欄位（允許 NULL）
                    cursor.execute("""
                        ALTER TABLE processing_rules
                        ADD COLUMN user_id INT NULL AFTER id
                    """)
                    print("✅ 已新增 processing_rules.user_id 欄位")

                    # 將舊資料設定為預設用戶
                    cursor.execute("""
                        UPDATE processing_rules SET user_id = %s WHERE user_id IS NULL
                    """, (default_user_id,))
                    print(f"✅ 已將舊規則資料關聯到用戶 ID: {default_user_id}")

                    # 設定為 NOT NULL
                    cursor.execute("""
                        ALTER TABLE processing_rules
                        MODIFY COLUMN user_id INT NOT NULL
                    """)

                    # 加入外鍵約束
                    try:
                        cursor.execute("""
                            ALTER TABLE processing_rules
                            ADD CONSTRAINT fk_processing_rules_user
                            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
                        """)
                        print("✅ 已新增外鍵約束")
                    except Exception as fk_error:
                        if "Duplicate" not in str(fk_error):
                            print(f"⚠️ 外鍵約束警告: {fk_error}")
            except Exception as alter_error:
                print(f"⚠️ 資料表升級檢查: {alter_error}")

            # 檢查是否需要初始化預設規則（為廣達用戶初始化）
            cursor.execute("SELECT COUNT(*) as count FROM processing_rules")
            if cursor.fetchone()['count'] == 0:
                # 取得廣達用戶的 ID
                cursor.execute("SELECT id FROM users WHERE username = 'quanta'")
                quanta_user = cursor.fetchone()
                if quanta_user:
                    init_default_processing_rules(cursor, quanta_user['id'])

            # 升級 upload_records.original_filename 欄位長度（支援多檔案上傳）
            try:
                # 先檢查目前欄位長度
                cursor.execute("""
                    SELECT CHARACTER_MAXIMUM_LENGTH
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = DATABASE()
                    AND TABLE_NAME = 'upload_records'
                    AND COLUMN_NAME = 'original_filename'
                """)
                result = cursor.fetchone()
                current_length = result['CHARACTER_MAXIMUM_LENGTH'] if result else 0

                if current_length < 2000:
                    cursor.execute("""
                        ALTER TABLE upload_records
                        MODIFY COLUMN original_filename VARCHAR(2000)
                    """)
                    print(f"✅ 已升級 upload_records.original_filename 欄位長度 ({current_length} -> 2000)")
                else:
                    print(f"ℹ️ upload_records.original_filename 欄位長度已足夠 ({current_length})")
            except Exception as alter_error:
                print(f"⚠️ 升級 original_filename 欄位失敗: {alter_error}")

            connection.commit()
            print("✅ 資料庫表格初始化完成")
            return True

    except Exception as e:
        print(f"❌ 資料庫初始化失敗: {e}")
        return False
    finally:
        connection.close()


def init_default_processing_rules(cursor, user_id):
    """
    初始化預設的處理規則（給指定客戶）

    規則按照實際處理流程排序：
    1. 上傳階段 (upload) - 上傳 3 個必要檔案
    2. 清理階段 (cleanup) - 清理 Forecast 舊資料
    3. 客戶映射 (mapping) - 客戶資料映射規則
    4. ERP 處理 (erp) - ERP 資料匹配與計算
    5. Transit 處理 (transit) - 在途資料處理
    6. Forecast 處理 (forecast) - 定位與填入數值
    7. 輸出階段 (output) - 產生結果並下載

    參數:
        cursor: 資料庫游標
        user_id: 客戶的用戶 ID
    """
    import json

    default_rules = [
        # ==================== 階段一：上傳文件 ====================
        {
            'rule_name': '上傳文件流程',
            'rule_category': 'upload',
            'rule_description': '上傳系統所需的三個 Excel 檔案，系統會自動驗證檔案格式',
            'rule_config': json.dumps({
                'upload_steps': [
                    {
                        'step': 1,
                        'name': '上傳 Forecast 檔案',
                        'description': '選擇並上傳客戶提供的 Forecast Excel 檔案',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 2,
                        'name': '上傳 ERP 淨需求檔案',
                        'description': '選擇並上傳 ERP 匯出的淨需求資料',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 3,
                        'name': '上傳在途資料檔案',
                        'description': '選擇並上傳在途貨物清單',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 4,
                        'name': '檔案格式驗證',
                        'description': '系統自動檢查每個檔案的欄位格式是否符合規範，驗證通過後才能進行下一步處理',
                        'auto': True,
                        'validation': True
                    }
                ]
            }, ensure_ascii=False),
            'display_order': 0
        },
        # ==================== 階段二：清理舊資料 ====================
        {
            'rule_name': 'Forecast 清理規則',
            'rule_category': 'cleanup',
            'rule_description': '處理前清理 Forecast 的舊資料',
            'rule_config': json.dumps({
                'cleanup_rules': [
                    {
                        'name': '供應數量清理',
                        'condition': 'K欄位 = "供應數量"',
                        'action': '清空 L~AW 欄位（第12~49欄）的數值',
                        'purpose': '移除舊的預測資料'
                    },
                    {
                        'name': '庫存數量清理',
                        'condition': 'I欄位 包含 "庫存數量"',
                        'action': '清空下一行的 I 欄位',
                        'purpose': '移除庫存標記'
                    }
                ]
            }, ensure_ascii=False),
            'display_order': 1
        },
        # ==================== 階段三：客戶映射 ====================
        {
            'rule_name': '客戶映射規則',
            'rule_category': 'mapping',
            'rule_description': '客戶資料的映射對應設定',
            'rule_config': json.dumps({
                'mapping_fields': {
                    '客戶簡稱': '匹配原始資料中的客戶名稱',
                    '客戶廠區': '對應到 Forecast 的 D 欄（客戶需求地區）',
                    '送貨地點': '送貨目的地資訊',
                    '排程出貨日期斷點': '決定週期計算的基準日（預設禮拜四）',
                    'ETD': '預計出發日期',
                    'ETA': '預計到達日期（用於計算目標週期）'
                },
                'unique_key': '用戶ID + 客戶簡稱 + 客戶廠區'
            }, ensure_ascii=False),
            'display_order': 2
        },
        # ==================== 階段四：ERP 處理 ====================
        {
            'rule_name': 'ERP 資料匹配規則',
            'rule_category': 'erp',
            'rule_description': 'ERP 淨需求文件與 Forecast 的匹配方式',
            'rule_config': json.dumps({
                'match_keys': ['客戶料號', '客戶需求地區'],
                'source_columns': {
                    '客戶簡稱': 'A欄位',
                    '客戶料號': 'B欄位',
                    '排程出貨日期': 'C欄位',
                    '淨需求': 'D欄位'
                },
                'description': '使用 客戶料號 + 客戶需求地區 作為匹配鍵，在 Forecast 中找到對應的資料區塊'
            }, ensure_ascii=False),
            'display_order': 3
        },
        {
            'rule_name': 'ERP 日期計算規則',
            'rule_category': 'erp',
            'rule_description': '從排程出貨日期計算目標週期',
            'rule_config': json.dumps({
                'date_field': '排程出貨日期',
                'breakpoint_field': '排程出貨日期斷點',
                'default_breakpoint': '禮拜四',
                'eta_field': 'ETA',
                'default_eta': '下下週二',
                'week_calculation': {
                    'step1': '取得排程出貨日期',
                    'step2': '根據排程出貨日期斷點（預設禮拜四）計算該週的起迄日',
                    'step3': '根據 ETA（如：下下週二）計算目標日期',
                    'step4': '在 Forecast 的 L~AW 欄位中找到符合日期範圍的欄位'
                },
                'eta_formats': {
                    '本週X': '當前計算週 + 星期X',
                    '下週X': '下一週 + 星期X',
                    '下下週X': '兩週後 + 星期X'
                }
            }, ensure_ascii=False),
            'display_order': 4
        },
        {
            'rule_name': 'ERP 數值轉換規則',
            'rule_category': 'erp',
            'rule_description': '淨需求數值的轉換計算',
            'rule_config': json.dumps({
                'source_field': '淨需求',
                'multiplier': 1000,
                'description': '淨需求值 × 1000 = 實際填入值',
                'example': '淨需求 = 5，填入 Forecast = 5000'
            }, ensure_ascii=False),
            'display_order': 5
        },
        # ==================== 階段五：Transit 處理 ====================
        {
            'rule_name': 'Transit 資料匹配規則',
            'rule_category': 'transit',
            'rule_description': '在途文件與 Forecast 的匹配方式',
            'rule_config': json.dumps({
                'match_keys': ['客戶需求地區(M欄)', 'Ordered Item(F欄)'],
                'source_columns': {
                    '客戶簡稱': 'E欄位',
                    'Ordered Item': 'F欄位（對應客戶料號）',
                    'Qty': 'H欄位（數量）',
                    'ETA': 'I欄位（預計到達日期）'
                },
                'description': '使用 客戶需求地區 + Ordered Item 作為匹配鍵'
            }, ensure_ascii=False),
            'display_order': 6
        },
        {
            'rule_name': 'Transit 日期處理規則',
            'rule_category': 'transit',
            'rule_description': '在途文件的 ETA 日期處理',
            'rule_config': json.dumps({
                'date_field': 'ETA (I欄位)',
                'date_formats': ['YYYY/MM/DD', 'YYYY-MM-DD', 'YYYYMMDD'],
                'description': '直接使用 I 欄位的 ETA 日期作為目標日期，找到 Forecast 對應的週期欄位'
            }, ensure_ascii=False),
            'display_order': 7
        },
        {
            'rule_name': 'Transit 數值轉換規則',
            'rule_category': 'transit',
            'rule_description': '在途數量的轉換計算',
            'rule_config': json.dumps({
                'source_field': 'Qty (H欄位)',
                'multiplier': 1000,
                'description': 'Qty 值 × 1000 = 實際填入值',
                'example': 'Qty = 3，填入 Forecast = 3000'
            }, ensure_ascii=False),
            'display_order': 8
        },
        # ==================== 階段六：Forecast 處理 ====================
        {
            'rule_name': 'Forecast 資料區塊識別規則',
            'rule_category': 'forecast',
            'rule_description': '如何在 Forecast 中識別資料區塊',
            'rule_config': json.dumps({
                'block_identification': {
                    'primary_key': 'A欄位（客戶料號）',
                    'secondary_key': 'D欄位（客戶需求地區）',
                    'method': '掃描 A 欄和 D 欄，相同的 (料號, 地區) 組合視為同一資料區塊'
                },
                'data_range': {
                    'date_columns': 'L~AW 欄位（第12~49欄）',
                    'start_date_row': '資料區塊起始行 + 1',
                    'end_date_row': '資料區塊起始行 + 2',
                    'target_row_marker': '供應數量'
                }
            }, ensure_ascii=False),
            'display_order': 9
        },
        {
            'rule_name': 'Forecast 目標欄位定位規則',
            'rule_category': 'forecast',
            'rule_description': '如何找到要填入數值的位置',
            'rule_config': json.dumps({
                'column_finding': {
                    'scan_range': 'L~AW 欄位',
                    'date_range_check': '起始日期 ≤ 目標日期 ≤ 結束日期',
                    'description': '在日期範圍內找到符合目標日期的欄位'
                },
                'row_finding': {
                    'marker': 'K欄位 = "供應數量"',
                    'scan_range': '資料區塊起始行 ~ 起始行+18',
                    'description': '找到標記為「供應數量」的行作為填入位置'
                }
            }, ensure_ascii=False),
            'display_order': 10
        },
        {
            'rule_name': 'Forecast 數值累加規則',
            'rule_category': 'forecast',
            'rule_description': 'ERP 和 Transit 數值的累加邏輯',
            'rule_config': json.dumps({
                'accumulation_logic': {
                    'rule': '相同位置的數值會累加，不會覆蓋',
                    'example': {
                        'ERP填入': 5000,
                        'Transit填入': 3000,
                        '最終結果': 8000
                    },
                    'description': '如果 ERP 和 Transit 都有數值要填入同一個格子，會將兩者相加'
                }
            }, ensure_ascii=False),
            'display_order': 11
        },
        # ==================== 階段七：輸出結果 ====================
        {
            'rule_name': '輸出與下載',
            'rule_category': 'output',
            'rule_description': '處理完成後的輸出流程',
            'rule_config': json.dumps({
                'output_steps': [
                    {
                        'step': 1,
                        'name': '驗證處理結果',
                        'description': '系統自動檢查所有資料是否正確填入',
                        'auto': True
                    },
                    {
                        'step': 2,
                        'name': '產生處理報告',
                        'description': '統計成功匹配筆數、失敗筆數等資訊',
                        'auto': True
                    },
                    {
                        'step': 3,
                        'name': '下載處理結果',
                        'description': '下載已填入數據的 Forecast Excel 檔案',
                        'file_type': 'Excel (.xlsx)',
                        'output': True
                    }
                ]
            }, ensure_ascii=False),
            'display_order': 12
        }
    ]

    for rule in default_rules:
        cursor.execute("""
            INSERT INTO processing_rules (user_id, rule_name, rule_category, rule_description, rule_config, display_order)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (user_id, rule['rule_name'], rule['rule_category'], rule['rule_description'],
              rule['rule_config'], rule['display_order']))

    print(f"✅ 已初始化預設處理規則 (user_id: {user_id})")


def update_processing_rules_enum():
    """
    更新 processing_rules 表的 rule_category ENUM
    新增 output 類別，並按照實際處理流程排序
    順序：upload -> cleanup -> mapping -> erp -> transit -> forecast -> output
    """
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            # 更新 ENUM 類型，按照實際處理流程排序
            cursor.execute("""
                ALTER TABLE processing_rules
                MODIFY COLUMN rule_category ENUM('upload', 'cleanup', 'mapping', 'erp', 'transit', 'forecast', 'output') NOT NULL
            """)
            connection.commit()
            print("✅ processing_rules ENUM 更新完成（按流程排序，新增 output）")
            return True
    except Exception as e:
        print(f"❌ 更新 processing_rules ENUM 失敗: {e}")
        return False
    finally:
        connection.close()


def add_upload_rule_to_user(user_id):
    """為指定用戶新增上傳文件流程規則（包含上傳步驟與格式驗證）"""
    import json
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    try:
        with connection.cursor() as cursor:
            # 檢查是否已有此規則
            cursor.execute("""
                SELECT id FROM processing_rules
                WHERE user_id = %s AND rule_category = 'upload'
            """, (user_id,))
            if cursor.fetchone():
                return True, "上傳流程規則已存在"

            # 新增上傳流程規則（包含上傳步驟與格式驗證）
            rule_config = json.dumps({
                'upload_steps': [
                    {
                        'step': 1,
                        'name': '上傳 Forecast 檔案',
                        'description': '選擇並上傳客戶提供的 Forecast Excel 檔案',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 2,
                        'name': '上傳 ERP 淨需求檔案',
                        'description': '選擇並上傳 ERP 匯出的淨需求資料',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 3,
                        'name': '上傳在途資料檔案',
                        'description': '選擇並上傳在途貨物清單',
                        'file_type': 'Excel (.xlsx)',
                        'required': True,
                        'validation': True
                    },
                    {
                        'step': 4,
                        'name': '檔案格式驗證',
                        'description': '系統自動檢查每個檔案的欄位格式是否符合規範，驗證通過後才能進行下一步處理',
                        'auto': True,
                        'validation': True
                    }
                ]
            }, ensure_ascii=False)

            cursor.execute("""
                INSERT INTO processing_rules (user_id, rule_name, rule_category, rule_description, rule_config, display_order)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (user_id, '上傳文件流程', 'upload', '上傳系統所需的三個 Excel 檔案，系統會自動驗證檔案格式', rule_config, 0))
            connection.commit()

            return True, "上傳流程規則新增成功"
    except Exception as e:
        print(f"❌ 新增上傳流程規則失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def add_upload_rule_to_all_users():
    """為所有一般用戶新增上傳文件流程規則"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            # 取得所有一般用戶
            cursor.execute("SELECT id, username FROM users WHERE role = 'user'")
            users = cursor.fetchall()

        for user in users:
            success, msg = add_upload_rule_to_user(user['id'])
            print(f"  - {user['username']}: {msg}")

        return True
    except Exception as e:
        print(f"❌ 批次新增上傳流程規則失敗: {e}")
        return False
    finally:
        connection.close()


def add_output_rule_to_user(user_id):
    """為指定用戶新增輸出與下載規則"""
    import json
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    try:
        with connection.cursor() as cursor:
            # 檢查是否已有此規則
            cursor.execute("""
                SELECT id FROM processing_rules
                WHERE user_id = %s AND rule_category = 'output'
            """, (user_id,))
            if cursor.fetchone():
                return True, "輸出規則已存在"

            # 新增輸出與下載規則
            rule_config = json.dumps({
                'output_steps': [
                    {
                        'step': 1,
                        'name': '驗證處理結果',
                        'description': '系統自動檢查所有資料是否正確填入',
                        'auto': True
                    },
                    {
                        'step': 2,
                        'name': '產生處理報告',
                        'description': '統計成功匹配筆數、失敗筆數等資訊',
                        'auto': True
                    },
                    {
                        'step': 3,
                        'name': '下載處理結果',
                        'description': '下載已填入數據的 Forecast Excel 檔案',
                        'file_type': 'Excel (.xlsx)',
                        'output': True
                    }
                ]
            }, ensure_ascii=False)

            cursor.execute("""
                INSERT INTO processing_rules (user_id, rule_name, rule_category, rule_description, rule_config, display_order)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (user_id, '輸出與下載', 'output', '處理完成後的輸出流程', rule_config, 12))
            connection.commit()

            return True, "輸出規則新增成功"
    except Exception as e:
        print(f"❌ 新增輸出規則失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def add_output_rule_to_all_users():
    """為所有一般用戶新增輸出與下載規則"""
    connection = get_db_connection()
    if not connection:
        return False

    try:
        with connection.cursor() as cursor:
            # 取得所有一般用戶
            cursor.execute("SELECT id, username FROM users WHERE role = 'user'")
            users = cursor.fetchall()

        for user in users:
            success, msg = add_output_rule_to_user(user['id'])
            print(f"  - {user['username']}: {msg}")

        return True
    except Exception as e:
        print(f"❌ 批次新增輸出規則失敗: {e}")
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

def get_customer_mappings_raw(user_id):
    """
    取得指定用戶的所有客戶映射資料（原始格式）
    返回完整的記錄列表，每筆記錄包含 customer_name, delivery_location, region 等
    用於 Pegatron 等需要多欄位匹配的客戶
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT customer_name, delivery_location, region, schedule_breakpoint, etd, eta
                FROM customer_mappings
                WHERE user_id = %s
            """, (user_id,))
            return cursor.fetchall()

    except Exception as e:
        print(f"❌ 取得客戶映射資料（原始格式）失敗: {e}")
        return []
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


# ==================== 規則管理 CRUD ====================

def get_all_processing_rules(user_id=None):
    """
    取得所有處理規則

    參數:
        user_id: 用戶 ID（可選，如果指定則只返回該用戶的規則）

    返回: 規則列表
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            if user_id:
                cursor.execute("""
                    SELECT pr.id, pr.user_id, pr.rule_name, pr.rule_category, pr.rule_description,
                           pr.rule_config, pr.is_active, pr.display_order, pr.created_at, pr.updated_at,
                           u.username, u.display_name, u.company
                    FROM processing_rules pr
                    LEFT JOIN users u ON pr.user_id = u.id
                    WHERE pr.user_id = %s
                    ORDER BY pr.rule_category, pr.display_order
                """, (user_id,))
            else:
                cursor.execute("""
                    SELECT pr.id, pr.user_id, pr.rule_name, pr.rule_category, pr.rule_description,
                           pr.rule_config, pr.is_active, pr.display_order, pr.created_at, pr.updated_at,
                           u.username, u.display_name, u.company
                    FROM processing_rules pr
                    LEFT JOIN users u ON pr.user_id = u.id
                    ORDER BY u.company, pr.rule_category, pr.display_order
                """)
            rules = cursor.fetchall()

            # 處理 JSON 欄位
            import json
            for rule in rules:
                if rule['rule_config']:
                    try:
                        rule['rule_config'] = json.loads(rule['rule_config'])
                    except:
                        pass

            return rules
    except Exception as e:
        print(f"❌ 取得處理規則失敗: {e}")
        return []
    finally:
        connection.close()


def get_processing_rules_by_user(user_id):
    """
    取得指定用戶的所有處理規則

    參數:
        user_id: 用戶 ID

    返回: 規則列表
    """
    return get_all_processing_rules(user_id)


def get_processing_rules_by_category(category, user_id=None):
    """
    根據類別取得處理規則

    參數:
        category: 規則類別 ('erp', 'transit', 'forecast', 'mapping', 'cleanup')
        user_id: 用戶 ID（可選）

    返回: 規則列表
    """
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            if user_id:
                cursor.execute("""
                    SELECT pr.id, pr.user_id, pr.rule_name, pr.rule_category, pr.rule_description,
                           pr.rule_config, pr.is_active, pr.display_order, pr.created_at, pr.updated_at,
                           u.username, u.display_name, u.company
                    FROM processing_rules pr
                    LEFT JOIN users u ON pr.user_id = u.id
                    WHERE pr.rule_category = %s AND pr.user_id = %s
                    ORDER BY pr.display_order
                """, (category, user_id))
            else:
                cursor.execute("""
                    SELECT pr.id, pr.user_id, pr.rule_name, pr.rule_category, pr.rule_description,
                           pr.rule_config, pr.is_active, pr.display_order, pr.created_at, pr.updated_at,
                           u.username, u.display_name, u.company
                    FROM processing_rules pr
                    LEFT JOIN users u ON pr.user_id = u.id
                    WHERE pr.rule_category = %s
                    ORDER BY u.company, pr.display_order
                """, (category,))
            rules = cursor.fetchall()

            import json
            for rule in rules:
                if rule['rule_config']:
                    try:
                        rule['rule_config'] = json.loads(rule['rule_config'])
                    except:
                        pass

            return rules
    except Exception as e:
        print(f"❌ 取得處理規則失敗: {e}")
        return []
    finally:
        connection.close()


def get_processing_rule_by_id(rule_id):
    """
    根據 ID 取得單一規則

    參數:
        rule_id: 規則 ID

    返回: 規則資料或 None
    """
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT pr.id, pr.user_id, pr.rule_name, pr.rule_category, pr.rule_description,
                       pr.rule_config, pr.is_active, pr.display_order, pr.created_at, pr.updated_at,
                       u.username, u.display_name, u.company
                FROM processing_rules pr
                LEFT JOIN users u ON pr.user_id = u.id
                WHERE pr.id = %s
            """, (rule_id,))
            rule = cursor.fetchone()

            if rule and rule['rule_config']:
                import json
                try:
                    rule['rule_config'] = json.loads(rule['rule_config'])
                except:
                    pass

            return rule
    except Exception as e:
        print(f"❌ 取得處理規則失敗: {e}")
        return None
    finally:
        connection.close()


def update_processing_rule(rule_id, **kwargs):
    """
    更新處理規則

    參數:
        rule_id: 規則 ID
        **kwargs: 可更新的欄位 (rule_name, rule_description, rule_config, is_active, display_order)

    返回: (success, message)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    allowed_fields = ['rule_name', 'rule_description', 'rule_config', 'is_active', 'display_order']
    update_fields = []
    values = []

    import json
    for field in allowed_fields:
        if field in kwargs:
            update_fields.append(f"{field} = %s")
            value = kwargs[field]
            # 如果是 rule_config 且為 dict，轉為 JSON
            if field == 'rule_config' and isinstance(value, dict):
                value = json.dumps(value, ensure_ascii=False)
            values.append(value)

    if not update_fields:
        return False, "沒有提供要更新的欄位"

    values.append(rule_id)

    try:
        with connection.cursor() as cursor:
            # 檢查規則是否存在
            cursor.execute("SELECT id FROM processing_rules WHERE id = %s", (rule_id,))
            if not cursor.fetchone():
                return False, "規則不存在"

            sql = f"UPDATE processing_rules SET {', '.join(update_fields)} WHERE id = %s"
            cursor.execute(sql, values)
            connection.commit()

            return True, "規則更新成功"
    except Exception as e:
        print(f"❌ 更新處理規則失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def create_processing_rule(user_id, rule_name, rule_category, rule_description=None, rule_config=None, display_order=0):
    """
    建立新的處理規則

    參數:
        user_id: 用戶 ID（必填）
        rule_name: 規則名稱
        rule_category: 規則類別 ('erp', 'transit', 'forecast', 'mapping', 'cleanup')
        rule_description: 規則描述
        rule_config: 規則配置 (dict)
        display_order: 顯示順序

    返回: (success, message, rule_id)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗", None

    try:
        import json
        config_json = json.dumps(rule_config, ensure_ascii=False) if rule_config else None

        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO processing_rules (user_id, rule_name, rule_category, rule_description, rule_config, display_order)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (user_id, rule_name, rule_category, rule_description, config_json, display_order))
            connection.commit()

            return True, "規則建立成功", cursor.lastrowid
    except Exception as e:
        print(f"❌ 建立處理規則失敗: {e}")
        return False, str(e), None
    finally:
        connection.close()


def delete_processing_rule(rule_id):
    """
    刪除處理規則

    參數:
        rule_id: 規則 ID

    返回: (success, message)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗"

    try:
        with connection.cursor() as cursor:
            # 檢查規則是否存在
            cursor.execute("SELECT rule_name FROM processing_rules WHERE id = %s", (rule_id,))
            rule = cursor.fetchone()
            if not rule:
                return False, "規則不存在"

            cursor.execute("DELETE FROM processing_rules WHERE id = %s", (rule_id,))
            connection.commit()

            return True, f"規則「{rule['rule_name']}」已刪除"
    except Exception as e:
        print(f"❌ 刪除處理規則失敗: {e}")
        return False, str(e)
    finally:
        connection.close()


def toggle_processing_rule_status(rule_id):
    """
    切換規則啟用狀態

    參數:
        rule_id: 規則 ID

    返回: (success, message, new_status)
    """
    connection = get_db_connection()
    if not connection:
        return False, "資料庫連線失敗", None

    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT is_active FROM processing_rules WHERE id = %s", (rule_id,))
            rule = cursor.fetchone()
            if not rule:
                return False, "規則不存在", None

            new_status = not rule['is_active']
            cursor.execute("UPDATE processing_rules SET is_active = %s WHERE id = %s", (new_status, rule_id))
            connection.commit()

            status_text = "啟用" if new_status else "停用"
            return True, f"規則已{status_text}", new_status
    except Exception as e:
        print(f"❌ 切換規則狀態失敗: {e}")
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

        print("\n更新 processing_rules ENUM（按流程排序，新增 output）...")
        update_processing_rules_enum()

        print("\n為現有用戶新增上傳流程規則...")
        add_upload_rule_to_all_users()

        print("\n為現有用戶新增輸出與下載規則...")
        add_output_rule_to_all_users()

        print("\n建立預設帳號...")
        create_default_users()
    else:
        print("❌ 資料庫連線失敗")
