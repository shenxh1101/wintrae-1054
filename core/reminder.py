import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List


class ReminderGenerator:
    CATEGORY_INTERVALS = {
        '慢病': 90,
        '孕产妇': 30,
        '儿童': 90,
        '老年人': 180,
    }

    def __init__(self, category: str):
        if category not in self.CATEGORY_INTERVALS:
            raise ValueError(f"不支持的分类: {category}")
        self.category = category
        self.interval_days = self.CATEGORY_INTERVALS[category]

    def get_latest_followup(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df

        df = df.copy()
        df['随访日期'] = pd.to_datetime(df['随访日期'], errors='coerce')
        df = df.dropna(subset=['随访日期'])

        df_sorted = df.sort_values('随访日期', ascending=False)
        latest = df_sorted.drop_duplicates(subset=['身份证号'], keep='first')

        return latest

    def find_overdue_persons(self, df: pd.DataFrame, reference_date: datetime = None) -> Dict:
        if reference_date is None:
            reference_date = datetime.now()

        latest_df = self.get_latest_followup(df)

        if latest_df.empty:
            return {
                '超期人员': pd.DataFrame(),
                '超期人数': 0,
                '总人数': 0,
                '超期率': '0%',
                '参考日期': reference_date.strftime('%Y-%m-%d'),
                '随访间隔天数': self.interval_days,
            }

        latest_df = latest_df.copy()
        latest_df['距今天数'] = (reference_date - latest_df['随访日期']).dt.days
        latest_df['应随访日期'] = latest_df['随访日期'] + timedelta(days=self.interval_days)
        latest_df['超期天数'] = (reference_date - latest_df['应随访日期']).dt.days

        overdue = latest_df[latest_df['超期天数'] > 0].copy()

        overdue['随访日期'] = overdue['随访日期'].dt.strftime('%Y-%m-%d')
        overdue['应随访日期'] = overdue['应随访日期'].dt.strftime('%Y-%m-%d')

        overdue = overdue.sort_values('超期天数', ascending=False)

        total_persons = len(latest_df)
        overdue_count = len(overdue)
        overdue_rate = f"{(overdue_count / total_persons * 100):.2f}%" if total_persons > 0 else "0%"

        return {
            '超期人员': overdue,
            '超期人数': overdue_count,
            '总人数': total_persons,
            '超期率': overdue_rate,
            '参考日期': reference_date.strftime('%Y-%m-%d'),
            '随访间隔天数': self.interval_days,
        }

    def generate_next_month_reminder(self, df: pd.DataFrame, reference_date: datetime = None) -> Dict:
        if reference_date is None:
            reference_date = datetime.now()

        latest_df = self.get_latest_followup(df)

        if latest_df.empty:
            return {
                '下月待随访人员': pd.DataFrame(),
                '待随访人数': 0,
                '总人数': 0,
                '参考日期': reference_date.strftime('%Y-%m-%d'),
                '随访间隔天数': self.interval_days,
            }

        latest_df = latest_df.copy()
        latest_df['下次随访日期'] = latest_df['随访日期'] + timedelta(days=self.interval_days)

        next_month_start = (reference_date.replace(day=1) + timedelta(days=32)).replace(day=1)
        next_month_end = (next_month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

        mask = (latest_df['下次随访日期'] >= next_month_start) & (latest_df['下次随访日期'] <= next_month_end)
        due_next_month = latest_df[mask].copy()

        due_next_month['随访日期'] = due_next_month['随访日期'].dt.strftime('%Y-%m-%d')
        due_next_month['下次随访日期'] = due_next_month['下次随访日期'].dt.strftime('%Y-%m-%d')

        due_next_month = due_next_month.sort_values('下次随访日期')

        return {
            '下月待随访人员': due_next_month,
            '待随访人数': len(due_next_month),
            '总人数': len(latest_df),
            '参考日期': reference_date.strftime('%Y-%m-%d'),
            '随访间隔天数': self.interval_days,
        }

    def generate_village_reminder(self, df: pd.DataFrame, reference_date: datetime = None) -> Dict:
        overdue_result = self.find_overdue_persons(df, reference_date)
        overdue_df = overdue_result['超期人员']

        if overdue_df.empty:
            return {
                '村居待随访清单': {},
                '村居数量': 0,
                '总超期人数': 0,
            }

        village_groups = {}
        for village, group in overdue_df.groupby('村居'):
            village_groups[village] = {
                '超期人数': len(group),
                '人员清单': group,
            }

        sorted_villages = dict(sorted(village_groups.items(), key=lambda x: x[1]['超期人数'], reverse=True))

        return {
            '村居待随访清单': sorted_villages,
            '村居数量': len(sorted_villages),
            '总超期人数': overdue_result['超期人数'],
        }

    def get_overdue_level(self, days_overdue: int) -> str:
        if days_overdue <= 7:
            return '轻度超期'
        elif days_overdue <= 30:
            return '中度超期'
        else:
            return '重度超期'

    def add_overdue_level(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        if '超期天数' in df.columns:
            df['超期等级'] = df['超期天数'].apply(self.get_overdue_level)
        return df
