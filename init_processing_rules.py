#!/usr/bin/env python3
"""
處理規則初始化腳本

用於初始化或更新客戶的處理規則配置
- 廣達 (Quanta): 多檔案合併模式，在途選填
- 和碩 (Pegatron): 多檔案獨立模式，在途按廠區檢查

使用方式:
    python init_processing_rules.py                    # 初始化所有用戶
    python init_processing_rules.py --user quanta      # 只初始化廣達
    python init_processing_rules.py --user pegatron    # 只初始化和碩
    python init_processing_rules.py --update-quanta    # 更新廣達現有規則
    python init_processing_rules.py --list             # 列出所有用戶規則數量
"""

import argparse
import sys
from database import (
    get_db_connection,
    init_default_processing_rules,
    init_pegatron_processing_rules,
    init_processing_rules_for_user,
    update_quanta_processing_rules,
    get_all_processing_rules
)


def get_user_by_username(username):
    """根據用戶名取得用戶資訊"""
    connection = get_db_connection()
    if not connection:
        return None

    try:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT id, username, display_name, company FROM users WHERE username = %s",
                (username,)
            )
            return cursor.fetchone()
    finally:
        connection.close()


def get_all_regular_users():
    """取得所有一般用戶"""
    connection = get_db_connection()
    if not connection:
        return []

    try:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT id, username, display_name, company FROM users WHERE role = 'user'"
            )
            return cursor.fetchall()
    finally:
        connection.close()


def list_user_rules():
    """列出所有用戶的規則數量"""
    connection = get_db_connection()
    if not connection:
        print("[ERROR] 資料庫連線失敗")
        return

    try:
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT u.id, u.username, u.display_name, u.company, COUNT(pr.id) as rule_count
                FROM users u
                LEFT JOIN processing_rules pr ON u.id = pr.user_id
                WHERE u.role = 'user'
                GROUP BY u.id, u.username, u.display_name, u.company
                ORDER BY u.id
            """)
            users = cursor.fetchall()

            print("\n" + "=" * 70)
            print("用戶處理規則統計")
            print("=" * 70)
            print(f"{'ID':<5} {'用戶名':<15} {'顯示名稱':<15} {'公司':<15} {'規則數':<10}")
            print("-" * 70)

            for user in users:
                print(f"{user['id']:<5} {user['username']:<15} {user['display_name']:<15} "
                      f"{user['company'] or '未設定':<15} {user['rule_count']:<10}")

            print("=" * 70)
    finally:
        connection.close()


def init_user_rules(username):
    """為指定用戶初始化規則"""
    user = get_user_by_username(username)
    if not user:
        print(f"[ERROR] 找不到用戶: {username}")
        return False

    # 根據用戶名或公司判斷類型
    company_lower = (user['company'] or '').lower()
    username_lower = username.lower()

    if 'pegatron' in company_lower or 'pegatron' in username_lower or '和碩' in (user['company'] or ''):
        company_type = 'pegatron'
    else:
        company_type = 'quanta'

    print(f"\n初始化 {user['display_name']} ({company_type}) 的處理規則...")

    success, message = init_processing_rules_for_user(user['id'], company_type)
    if success:
        print(f"[OK] {message}")
    else:
        print(f"[ERROR] {message}")

    return success


def init_all_users():
    """為所有一般用戶初始化規則"""
    users = get_all_regular_users()
    if not users:
        print("[ERROR] 沒有找到一般用戶")
        return

    print(f"\n找到 {len(users)} 個一般用戶，開始初始化...")
    print("-" * 50)

    success_count = 0
    for user in users:
        company_lower = (user['company'] or '').lower()
        username_lower = user['username'].lower()

        if 'pegatron' in company_lower or 'pegatron' in username_lower or '和碩' in (user['company'] or ''):
            company_type = 'pegatron'
        else:
            company_type = 'quanta'

        print(f"\n處理 {user['display_name']} (ID: {user['id']}, 類型: {company_type})...")

        success, message = init_processing_rules_for_user(user['id'], company_type)
        if success:
            print(f"  [OK] {message}")
            success_count += 1
        else:
            print(f"  [ERROR] {message}")

    print("-" * 50)
    print(f"完成！成功: {success_count}/{len(users)}")


def update_quanta_rules():
    """更新廣達現有規則"""
    user = get_user_by_username('quanta')
    if not user:
        print("[ERROR] 找不到廣達用戶")
        return False

    print(f"\n更新廣達 ({user['display_name']}) 的處理規則...")

    success, message = update_quanta_processing_rules(user['id'])
    if success:
        print(f"[OK] {message}")
    else:
        print(f"[ERROR] {message}")

    return success


def show_user_rules(username):
    """顯示指定用戶的規則詳情"""
    user = get_user_by_username(username)
    if not user:
        print(f"[ERROR] 找不到用戶: {username}")
        return

    rules = get_all_processing_rules(user['id'])

    print("\n" + "=" * 80)
    print(f"{user['display_name']} ({user['company']}) 的處理規則")
    print("=" * 80)

    categories = {}
    for rule in rules:
        cat = rule['rule_category']
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(rule)

    category_names = {
        'upload': '上傳文件',
        'cleanup': '資料清理',
        'mapping': '客戶映射',
        'erp': 'ERP 處理',
        'transit': '在途處理',
        'forecast': 'Forecast 處理',
        'output': '輸出結果'
    }

    for cat in ['upload', 'cleanup', 'mapping', 'erp', 'transit', 'forecast', 'output']:
        if cat in categories:
            print(f"\n【{category_names.get(cat, cat)}】")
            print("-" * 40)
            for rule in categories[cat]:
                status = "啟用" if rule['is_active'] else "停用"
                print(f"  • {rule['rule_name']} [{status}]")
                if rule['rule_description']:
                    print(f"    {rule['rule_description']}")

    print("\n" + "=" * 80)
    print(f"共 {len(rules)} 條規則")


def main():
    parser = argparse.ArgumentParser(description='處理規則初始化工具')
    parser.add_argument('--user', '-u', help='指定要初始化的用戶名')
    parser.add_argument('--update-quanta', action='store_true', help='更新廣達現有規則')
    parser.add_argument('--list', '-l', action='store_true', help='列出所有用戶規則數量')
    parser.add_argument('--show', '-s', help='顯示指定用戶的規則詳情')
    parser.add_argument('--all', '-a', action='store_true', help='初始化所有用戶')

    args = parser.parse_args()

    if args.list:
        list_user_rules()
    elif args.show:
        show_user_rules(args.show)
    elif args.update_quanta:
        update_quanta_rules()
    elif args.user:
        init_user_rules(args.user)
    elif args.all:
        init_all_users()
    else:
        # 預設顯示幫助
        parser.print_help()
        print("\n範例:")
        print("  python init_processing_rules.py --list              # 列出所有用戶")
        print("  python init_processing_rules.py --show quanta       # 顯示廣達規則")
        print("  python init_processing_rules.py --user pegatron     # 初始化和碩")
        print("  python init_processing_rules.py --all               # 初始化所有用戶")


if __name__ == "__main__":
    main()
