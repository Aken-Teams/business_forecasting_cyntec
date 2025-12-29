# -*- coding: utf-8 -*-
"""
Pegatron Forecast Processor
專門處理 Pegatron 的 Transit + ERP -> Forecast 邏輯
使用 LibreOffice 跨平台方案保留 .xls/.xlsx 格式和公式
支援 Windows 和 Linux 環境
"""
import pandas as pd
import os
import shutil
from datetime import datetime, timedelta


class PegatronForecastProcessor:
    """
    Pegatron 專用 Forecast 處理器
    - Transit 和 ERP 都填入 ETA QTY 行
    - 使用 LibreOffice 跨平台方案保留格式和公式
    - 支援累加邏輯
    - 追蹤已分配狀態，避免重複分配
    """

    def __init__(self, forecast_file, erp_file, transit_file=None, output_folder=None, output_filename=None):
        self.forecast_file = forecast_file
        self.erp_file = erp_file
        self.transit_file = transit_file
        self.output_folder = output_folder
        # 確保輸出為 xlsx 格式（避免轉換失敗）
        if output_filename:
            self.output_filename = os.path.splitext(output_filename)[0] + '.xlsx'
        else:
            self.output_filename = "forecast_result.xlsx"

        # 統計資料
        self.total_filled = 0
        self.total_skipped = 0
        self.total_transit_filled = 0
        self.total_transit_skipped = 0

        # 追蹤已分配的索引
        self.allocated_erp_indices = []
        self.allocated_transit_indices = []

        # 保存 DataFrame 以便更新已分配狀態
        self.erp_df = None
        self.transit_df = None

    def get_week_end_by_breakpoint(self, schedule_date, breakpoint_text):
        """
        根據排程出貨日期和斷點計算該斷點週的結束日
        斷點就是週的結束日
        """
        breakpoint_map = {
            '週一': 0, '週二': 1, '週三': 2, '週四': 3,
            '週五': 4, '週六': 5, '週日': 6,
            '禮拜一': 0, '禮拜二': 1, '禮拜三': 2, '禮拜四': 3,
            '禮拜五': 4, '禮拜六': 5, '禮拜日': 6,
        }

        breakpoint_weekday = breakpoint_map.get(breakpoint_text, 2)  # 預設週三
        current_weekday = schedule_date.weekday()

        if current_weekday <= breakpoint_weekday:
            days_to_breakpoint = breakpoint_weekday - current_weekday
        else:
            days_to_breakpoint = 7 - (current_weekday - breakpoint_weekday)

        week_end = schedule_date + timedelta(days=days_to_breakpoint)
        return week_end

    def calculate_erp_eta_target_date(self, week_end, eta_text):
        """
        根據 ETA 文字和斷點週結束日計算目標日期
        ETA 格式: 本週X, 下週X, 下下週X
        """
        eta_weekday_map = {
            '一': 0, '二': 1, '三': 2, '四': 3,
            '五': 4, '六': 5, '日': 6, '天': 6,
        }

        if not eta_text or pd.isna(eta_text):
            return None

        eta_text = str(eta_text).strip()

        if '下下週' in eta_text or '下下周' in eta_text:
            weeks_offset = 2
            weekday_char = eta_text.replace('下下週', '').replace('下下周', '').strip()
        elif '下週' in eta_text or '下周' in eta_text:
            weeks_offset = 1
            weekday_char = eta_text.replace('下週', '').replace('下周', '').strip()
        elif '本週' in eta_text or '本周' in eta_text:
            weeks_offset = 0
            weekday_char = eta_text.replace('本週', '').replace('本周', '').strip()
        else:
            return None

        target_weekday = eta_weekday_map.get(weekday_char, 1)  # 預設週二
        week_end_weekday = week_end.weekday()

        days_diff = target_weekday - week_end_weekday
        target_date = week_end + timedelta(days=7 * weeks_offset + days_diff)

        return target_date

    def find_week_column(self, target_date, date_columns):
        """
        找到目標日期所在週的欄位
        Forecast 日期已經是週一
        """
        if isinstance(target_date, datetime):
            target_date = target_date.date()

        days_since_monday = target_date.weekday()
        week_monday = target_date - timedelta(days=days_since_monday)

        if week_monday in date_columns:
            return date_columns[week_monday]

        for week_start, col_idx in date_columns.items():
            week_end = week_start + timedelta(days=6)
            if week_start <= target_date <= week_end:
                return col_idx

        return None

    def process_all_blocks(self):
        """主處理函數"""
        try:
            print("=" * 70)
            print("Pegatron Forecast Processor - Transit + ERP -> ETA QTY")
            print("=" * 70)

            # 1. 讀取資料
            # 根據檔案格式選擇引擎
            forecast_ext = os.path.splitext(self.forecast_file)[1].lower()
            if forecast_ext == '.xls':
                forecast_df = pd.read_excel(self.forecast_file, header=None, engine='xlrd')
            else:
                forecast_df = pd.read_excel(self.forecast_file, header=None)
            print(f"Forecast 行數: {len(forecast_df)}, 欄數: {len(forecast_df.columns)}")

            self.erp_df = pd.read_excel(self.erp_file)
            print(f"ERP 行數: {len(self.erp_df)}")

            # 確保 ERP 有「已分配」欄位
            if '已分配' not in self.erp_df.columns:
                self.erp_df['已分配'] = ''

            self.transit_df = None
            if self.transit_file and os.path.exists(self.transit_file):
                self.transit_df = pd.read_excel(self.transit_file)
                print(f"Transit 行數: {len(self.transit_df)}")
                # 確保 Transit 有「已分配」欄位
                if '已分配' not in self.transit_df.columns:
                    self.transit_df['已分配'] = ''

            # 2. 建立 Forecast 區塊結構
            # F+G 欄位 = Line 客戶採購單號, I 欄位 (row+1) = Ordered Item
            forecast_blocks = []
            row_idx = 2
            while row_idx < len(forecast_df):
                m_val = forecast_df.iloc[row_idx, 12] if pd.notna(forecast_df.iloc[row_idx, 12]) else ''
                if m_val == 'WEEK#':
                    f_val = str(forecast_df.iloc[row_idx, 5]).strip() if pd.notna(forecast_df.iloc[row_idx, 5]) else ''
                    g_val = str(forecast_df.iloc[row_idx, 6]).strip() if pd.notna(forecast_df.iloc[row_idx, 6]) else ''
                    line_po = f"{f_val}-{g_val}" if f_val and g_val else ''

                    ordered_item = ''
                    if row_idx + 1 < len(forecast_df):
                        ordered_item = str(forecast_df.iloc[row_idx + 1, 8]).strip() if pd.notna(forecast_df.iloc[row_idx + 1, 8]) else ''

                    eta_qty_row = row_idx + 4  # ETA QTY 行

                    forecast_blocks.append({
                        'start_row': row_idx,
                        'line_po': line_po,
                        'ordered_item': ordered_item,
                        'eta_qty_row': eta_qty_row + 1,  # Excel 1-based
                    })
                    row_idx += 8
                else:
                    row_idx += 1

            print(f"\n找到 {len(forecast_blocks)} 個 Forecast 區塊")

            # 3. 取得日期欄位對應
            date_columns = {}
            for col_idx in range(14, len(forecast_df.columns)):
                date_val = forecast_df.iloc[1, col_idx]
                if pd.notna(date_val):
                    if isinstance(date_val, str):
                        try:
                            date_obj = pd.to_datetime(date_val)
                            date_columns[date_obj.date()] = col_idx + 1
                        except:
                            pass
                    elif isinstance(date_val, (datetime, pd.Timestamp)):
                        date_columns[date_val.date()] = col_idx + 1

            print(f"日期欄位對應: {len(date_columns)} 個日期")

            all_updates = []

            # 4. 處理 Transit 資料
            if self.transit_df is not None:
                print("\n=== 處理 Transit 資料 ===")
                transit_updates = self._process_transit(self.transit_df, forecast_blocks, date_columns)
                all_updates.extend(transit_updates)
                self.total_transit_filled = len(transit_updates)
                print(f"Transit 更新筆數: {self.total_transit_filled}")

            # 5. 處理 ERP 資料
            print("\n=== 處理 ERP 資料 ===")
            erp_updates = self._process_erp(self.erp_df, forecast_blocks, date_columns)
            all_updates.extend(erp_updates)
            self.total_filled = len(erp_updates)
            print(f"ERP 更新筆數: {self.total_filled}")

            if not all_updates:
                print("\n沒有需要更新的資料，複製原始檔案作為輸出")
                # 即使沒有更新，也要建立輸出檔案讓用戶可以下載
                if self.output_folder:
                    output_path = os.path.join(self.output_folder, self.output_filename)
                else:
                    output_path = self.output_filename
                shutil.copy2(self.forecast_file, output_path)
                print(f"已輸出到: {output_path}")
                return True

            # 6. 使用 COM 寫入 Excel
            print(f"\n=== 使用 COM 更新 {len(all_updates)} 個儲存格 ===")
            success = self._write_to_excel(all_updates)

            if success:
                # 7. 更新並保存已分配狀態
                self._save_allocation_status()

                print("\n" + "=" * 50)
                print("處理完成")
                print("=" * 50)
                print(f"Transit 更新: {self.total_transit_filled} 筆")
                print(f"ERP 更新: {self.total_filled} 筆")

            return success

        except Exception as e:
            print(f"處理失敗: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _process_transit(self, transit_df, forecast_blocks, date_columns):
        """處理 Transit 資料"""
        updates = []

        # 檢查必要欄位
        if 'Line 客戶採購單號' not in transit_df.columns:
            print("Transit 缺少 'Line 客戶採購單號' 欄位")
            return updates

        transit_with_line = transit_df[transit_df['Line 客戶採購單號'].notna()]
        print(f"有 Line 客戶採購單號 的在途記錄: {len(transit_with_line)}")

        for idx, transit_row in transit_with_line.iterrows():
            # 跳過已分配的記錄
            if transit_df.at[idx, '已分配'] == '✓':
                continue

            transit_line_po = str(transit_row['Line 客戶採購單號']).strip()
            transit_ordered_item = str(transit_row['Ordered Item']).strip() if 'Ordered Item' in transit_row and pd.notna(transit_row['Ordered Item']) else ''
            transit_qty = transit_row['Qty'] if 'Qty' in transit_row and pd.notna(transit_row['Qty']) else 0
            transit_eta = transit_row['ETA'] if 'ETA' in transit_row else None

            if not transit_line_po or not transit_ordered_item or transit_qty == 0:
                continue

            # 找到匹配的 Forecast 區塊
            matched_block = None
            for block in forecast_blocks:
                if block['line_po'] == transit_line_po and block['ordered_item'] == transit_ordered_item:
                    matched_block = block
                    break

            if not matched_block:
                print(f"  ⚠️ Transit 跳過: {transit_line_po}, {transit_ordered_item} - 找不到匹配的 Forecast 區塊")
                self.total_transit_skipped += 1
                continue

            if not pd.notna(transit_eta):
                print(f"  ⚠️ Transit 跳過: {transit_line_po}, {transit_ordered_item} - ETA 為空")
                self.total_transit_skipped += 1
                continue

            eta_date = pd.to_datetime(transit_eta).date()
            days_since_monday = eta_date.weekday()
            week_start = eta_date - timedelta(days=days_since_monday)

            if week_start not in date_columns:
                print(f"  ⚠️ Transit 跳過: {transit_line_po}, {transit_ordered_item} - 找不到 ETA 日期 {eta_date} 對應的欄位")
                self.total_transit_skipped += 1
                continue

            excel_col = date_columns[week_start]
            eta_qty_value = transit_qty * 1000

            print(f"  Transit: {transit_line_po}, {transit_ordered_item} -> Row {matched_block['eta_qty_row']}, Col {excel_col}, 值={eta_qty_value}")
            updates.append((matched_block['eta_qty_row'], excel_col, eta_qty_value))

            # 記錄已分配的索引
            self.allocated_transit_indices.append(idx)

        return updates

    def _process_erp(self, erp_df, forecast_blocks, date_columns):
        """處理 ERP 資料"""
        updates = []

        # 檢查必要欄位
        required_cols = ['客戶需求地區', 'Line 客戶採購單號', '客戶料號', '淨需求', '排程出貨日期', '排程出貨日期斷點', 'ETA']
        missing_cols = [c for c in required_cols if c not in erp_df.columns]
        if missing_cols:
            print(f"ERP 缺少欄位: {missing_cols}")
            return updates

        erp_with_mapping = erp_df[erp_df['客戶需求地區'].notna() & (erp_df['客戶需求地區'] != '')]
        print(f"有 mapping 的 ERP 記錄: {len(erp_with_mapping)}")

        for idx, erp_row in erp_with_mapping.iterrows():
            # 跳過已分配的記錄
            if erp_df.at[idx, '已分配'] == '✓':
                continue

            erp_line_po = str(erp_row['Line 客戶採購單號']).strip() if pd.notna(erp_row['Line 客戶採購單號']) else ''
            erp_pn = str(erp_row['客戶料號']).strip() if pd.notna(erp_row['客戶料號']) else ''
            erp_qty = erp_row['淨需求'] if pd.notna(erp_row['淨需求']) else 0
            erp_schedule_date = erp_row['排程出貨日期']
            erp_breakpoint = str(erp_row['排程出貨日期斷點']).strip() if pd.notna(erp_row['排程出貨日期斷點']) else ''
            erp_eta = str(erp_row['ETA']).strip() if pd.notna(erp_row['ETA']) else ''

            if not erp_line_po or not erp_pn or erp_qty == 0:
                continue

            # 找到匹配的 Forecast 區塊
            matched_block = None
            for block in forecast_blocks:
                if block['line_po'] == erp_line_po and block['ordered_item'] == erp_pn:
                    matched_block = block
                    break

            if not matched_block:
                print(f"  ⚠️ ERP 跳過: {erp_line_po}, {erp_pn} - 找不到匹配的 Forecast 區塊")
                self.total_skipped += 1
                continue

            if pd.isna(erp_schedule_date) or not erp_breakpoint or not erp_eta:
                print(f"  ⚠️ ERP 跳過: {erp_line_po}, {erp_pn} - 缺少排程日期/斷點/ETA")
                self.total_skipped += 1
                continue

            schedule_date = pd.to_datetime(erp_schedule_date)
            week_end = self.get_week_end_by_breakpoint(schedule_date, erp_breakpoint)
            target_date = self.calculate_erp_eta_target_date(week_end, erp_eta)

            if target_date is None:
                print(f"  ⚠️ ERP 跳過: {erp_line_po}, {erp_pn} - 無法計算目標日期 (ETA={erp_eta})")
                self.total_skipped += 1
                continue

            excel_col = self.find_week_column(target_date, date_columns)

            if excel_col is None:
                print(f"  ⚠️ ERP 跳過: {erp_line_po}, {erp_pn} - 找不到目標日期 {target_date.date()} 對應的欄位")
                self.total_skipped += 1
                continue

            forecast_value = erp_qty * 1000

            print(f"  ERP: {erp_line_po}, {erp_pn} -> Row {matched_block['eta_qty_row']}, Col {excel_col}, 值={forecast_value}")
            updates.append((matched_block['eta_qty_row'], excel_col, forecast_value))

            # 記錄已分配的索引
            self.allocated_erp_indices.append(idx)

        return updates

    def _write_to_excel(self, updates):
        """使用 LibreOffice 跨平台方案寫入 Excel，保留格式和公式"""
        from libreoffice_utils import write_to_excel_libreoffice

        # 決定輸出路徑
        if self.output_folder:
            output_path = os.path.join(self.output_folder, self.output_filename)
        else:
            output_path = self.output_filename

        try:
            success = write_to_excel_libreoffice(self.forecast_file, updates, output_path)
            if success:
                print(f"\n已輸出到: {output_path}")
            return success
        except Exception as e:
            print(f"寫入 Excel 失敗: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _save_allocation_status(self):
        """保存已分配狀態到 ERP 和 Transit 檔案"""
        try:
            # 更新 ERP 已分配狀態
            if self.erp_df is not None and self.allocated_erp_indices:
                for idx in self.allocated_erp_indices:
                    self.erp_df.at[idx, '已分配'] = '✓'
                # 保存 ERP 檔案
                self.erp_df.to_excel(self.erp_file, index=False)
                print(f"已更新 ERP 已分配狀態: {len(self.allocated_erp_indices)} 筆")

            # 更新 Transit 已分配狀態
            if self.transit_df is not None and self.allocated_transit_indices:
                for idx in self.allocated_transit_indices:
                    self.transit_df.at[idx, '已分配'] = '✓'
                # 保存 Transit 檔案
                self.transit_df.to_excel(self.transit_file, index=False)
                print(f"已更新 Transit 已分配狀態: {len(self.allocated_transit_indices)} 筆")

        except Exception as e:
            print(f"保存已分配狀態失敗: {e}")
            import traceback
            traceback.print_exc()
