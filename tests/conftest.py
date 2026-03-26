# -*- coding: utf-8 -*-
"""共用 fixtures"""
import sys
import os
import pytest

# 確保專案根目錄在 import path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from liteon_forecast_processor import LiteonForecastProcessor


@pytest.fixture
def processor():
    """建立最小 LiteonForecastProcessor（不載入任何檔案）"""
    p = LiteonForecastProcessor.__new__(LiteonForecastProcessor)
    p.merged_mode = False
    p.col_offset = 0
    p.pending_changes = []
    p.date_map = {}
    p.daily_end_col = None
    p.weekly_start_col = None
    p.first_weekly_date = None
    p.monthly_map = {}
    p.plant_daily_end_dates = {}
    p.material_commit_rows = {}
    return p


def _build_date_map(start_col, daily_start, daily_days, weekly_start_date, weekly_cols, monthly_map):
    """建構 date_map dict"""
    from datetime import timedelta
    date_map = {}
    for i in range(daily_days):
        d = daily_start + timedelta(days=i)
        date_map[d] = start_col + i
    daily_end_col = start_col + daily_days - 1
    weekly_start_col = daily_end_col + 1
    first_weekly_date = weekly_start_date
    for i in range(weekly_cols):
        wd = weekly_start_date + timedelta(weeks=i)
        date_map[wd] = weekly_start_col + i
    for month_date, col in monthly_map.items():
        date_map[month_date] = col
    return date_map, daily_end_col, weekly_start_col, first_weekly_date


@pytest.fixture
def processor_with_dates_0309(processor):
    """模擬 03/09-start 的完整日期結構"""
    from datetime import date, timedelta
    dm, dec, wsc, fwd = _build_date_map(
        start_col=11,
        daily_start=date(2026, 3, 9),
        daily_days=31,
        weekly_start_date=date(2026, 4, 13),
        weekly_cols=22,
        monthly_map={
            date(2026, 9, 1): 64,
            date(2026, 10, 1): 65,
            date(2026, 11, 1): 66,
        }
    )
    processor.date_map = dm
    processor.daily_end_col = dec
    processor.weekly_start_col = wsc
    processor.first_weekly_date = fwd
    processor.monthly_map = {
        date(2026, 9, 1): 64,
        date(2026, 10, 1): 65,
        date(2026, 11, 1): 66,
    }
    return processor


@pytest.fixture
def processor_with_dates_0307(processor):
    """模擬 03/07-start (六天 GAP)"""
    from datetime import date
    dm, dec, wsc, fwd = _build_date_map(
        start_col=11,
        daily_start=date(2026, 3, 7),
        daily_days=31,
        weekly_start_date=date(2026, 4, 13),
        weekly_cols=22,
        monthly_map={
            date(2026, 9, 1): 64,
            date(2026, 10, 1): 65,
            date(2026, 11, 1): 66,
        }
    )
    processor.date_map = dm
    processor.daily_end_col = dec
    processor.weekly_start_col = wsc
    processor.first_weekly_date = fwd
    processor.monthly_map = {
        date(2026, 9, 1): 64,
        date(2026, 10, 1): 65,
        date(2026, 11, 1): 66,
    }
    return processor
