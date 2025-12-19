import pandas as pd
import sys
import os

def process_erp_and_mapping():
    """
    處理ERP文件和mapping表，根據客戶簡稱匹配並整合地區資料
    """
    try:
        # 讀取ERP文件
        erp_file_path = "erp/20250924 廣達淨需求.xlsx"
        mapping_file_path = "mapping/mapping表.xlsx"
        
        print("正在讀取ERP文件...")
        erp_df = pd.read_excel(erp_file_path)
        
        print("正在讀取mapping表...")
        mapping_df = pd.read_excel(mapping_file_path)
        
        print("ERP文件欄位:")
        print(erp_df.columns.tolist())
        print("\nERP文件前5行數據:")
        print(erp_df.head())
        
        print("\nmapping表欄位:")
        print(mapping_df.columns.tolist())
        print("\nmapping表前5行數據:")
        print(mapping_df.head())
        
        # 檢查D欄位是否存在於ERP文件中
        if len(erp_df.columns) >= 4:
            d_column_name = erp_df.columns[3]  # D欄位 (索引3)
            print(f"\nERP文件D欄位名稱: {d_column_name}")
            print(f"D欄位前10個值: {erp_df[d_column_name].head(10).tolist()}")
        else:
            print("錯誤: ERP文件沒有足夠的欄位")
            return
        
        # 檢查A欄位是否存在於mapping表中
        if len(mapping_df.columns) >= 1:
            a_column_name = mapping_df.columns[0]  # A欄位 (索引0)
            print(f"\nmapping表A欄位名稱: {a_column_name}")
            print(f"A欄位前10個值: {mapping_df[a_column_name].head(10).tolist()}")
        else:
            print("錯誤: mapping表沒有足夠的欄位")
            return
            
        # 尋找所有需要的mapping欄位
        region_columns = [col for col in mapping_df.columns if '地區' in str(col) or 'region' in str(col).lower()]
        schedule_columns = [col for col in mapping_df.columns if '排程' in str(col) or '斷點' in str(col)]
        etd_columns = [col for col in mapping_df.columns if 'ETD' in str(col)]
        eta_columns = [col for col in mapping_df.columns if 'ETA' in str(col)]
        
        # 確定各個欄位
        if region_columns:
            region_column = region_columns[0]
            print(f"\n找到地區欄位: {region_column}")
        else:
            print("\n未找到明確的地區欄位")
            region_column = None
            
        if schedule_columns:
            schedule_column = schedule_columns[0]
            print(f"找到排程出貨日期斷點欄位: {schedule_column}")
        else:
            print("未找到排程出貨日期斷點欄位")
            schedule_column = None
            
        if etd_columns:
            etd_column = etd_columns[0]
            print(f"找到ETD欄位: {etd_column}")
        else:
            print("未找到ETD欄位")
            etd_column = None
            
        if eta_columns:
            eta_column = eta_columns[0]
            print(f"找到ETA欄位: {eta_column}")
        else:
            print("未找到ETA欄位")
            eta_column = None
            
        return erp_df, mapping_df, d_column_name, a_column_name, region_column, schedule_column, etd_column, eta_column
        
    except Exception as e:
        print(f"讀取文件時發生錯誤: {e}")
        return None, None, None, None, None, None, None, None

def merge_data(erp_df, mapping_df, erp_customer_col, mapping_customer_col, region_col, schedule_col, etd_col, eta_col):
    """
    根據客戶簡稱匹配並整合數據
    """
    try:
        print(f"\n開始整合數據...")
        print(f"使用ERP文件欄位 '{erp_customer_col}' 與mapping表欄位 '{mapping_customer_col}' 進行匹配")
        
        # 創建各個mapping字典
        mappings = {}
        
        if region_col:
            mappings['客戶需求地區'] = dict(zip(mapping_df[mapping_customer_col], mapping_df[region_col]))
            print(f"客戶需求地區mapping樣本: {dict(list(mappings['客戶需求地區'].items())[:5])}")
        
        if schedule_col:
            mappings['排程出貨日期斷點'] = dict(zip(mapping_df[mapping_customer_col], mapping_df[schedule_col]))
            print(f"排程出貨日期斷點mapping樣本: {dict(list(mappings['排程出貨日期斷點'].items())[:5])}")
        
        if etd_col:
            mappings['ETD'] = dict(zip(mapping_df[mapping_customer_col], mapping_df[etd_col]))
            print(f"ETD mapping樣本: {dict(list(mappings['ETD'].items())[:5])}")
        
        if eta_col:
            mappings['ETA'] = dict(zip(mapping_df[mapping_customer_col], mapping_df[eta_col]))
            print(f"ETA mapping樣本: {dict(list(mappings['ETA'].items())[:5])}")
        
        # 在ERP數據中添加各個mapping欄位
        for field_name, mapping_dict in mappings.items():
            erp_df[field_name] = erp_df[erp_customer_col].map(mapping_dict)
            matched_count = erp_df[field_name].notna().sum()
            print(f"{field_name} 匹配成功: {matched_count}/{len(erp_df)} 筆")
        
        # 統計總體匹配結果
        total_count = len(erp_df)
        if region_col:
            matched_count = erp_df['客戶需求地區'].notna().sum()
            print(f"\n總體匹配統計:")
            print(f"總記錄數: {total_count}")
            print(f"成功匹配: {matched_count}")
            print(f"未匹配: {total_count - matched_count}")
            
            # 顯示未匹配的客戶簡稱
            unmatched = erp_df[erp_df['客戶需求地區'].isna()][erp_customer_col].unique()
            if len(unmatched) > 0:
                print(f"\n未匹配的客戶簡稱樣本: {list(unmatched[:10])}")
        
        return erp_df
        
    except Exception as e:
        print(f"整合數據時發生錯誤: {e}")
        return None

def save_result(df, output_filename="整合後的廣達淨需求.xlsx"):
    """
    保存整合後的結果
    """
    try:
        df.to_excel(output_filename, index=False)
        print(f"\n結果已保存到: {output_filename}")
        
        # 顯示結果摘要
        print(f"\n結果摘要:")
        print(f"總記錄數: {len(df)}")
        print(f"欄位數: {len(df.columns)}")
        print(f"欄位名稱: {df.columns.tolist()}")
        
        return True
        
    except Exception as e:
        print(f"保存文件時發生錯誤: {e}")
        return False

if __name__ == "__main__":
    print("開始處理ERP和mapping表數據...")
    
    # 讀取和檢查數據
    erp_df, mapping_df, erp_customer_col, mapping_customer_col, region_col, schedule_col, etd_col, eta_col = process_erp_and_mapping()
    
    if erp_df is not None and mapping_df is not None:
        # 整合數據
        merged_df = merge_data(erp_df, mapping_df, erp_customer_col, mapping_customer_col, region_col, schedule_col, etd_col, eta_col)
        
        if merged_df is not None:
            # 保存結果
            save_result(merged_df)
            print("\n處理完成!")
        else:
            print("數據整合失敗")
    else:
        print("文件讀取失敗")
