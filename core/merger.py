import pandas as pd
from datetime import datetime
from typing import List, Dict


class DataMerger:
    def __init__(self, category: str):
        self.category = category

    def merge_multiple_files(self, file_paths: List[str]) -> Dict:
        all_data = []
        file_info = []

        for file_path in file_paths:
            try:
                df = self._read_file(file_path)
                df['来源文件'] = file_path.split('\\')[-1]
                all_data.append(df)
                file_info.append({
                    '文件名': file_path.split('\\')[-1],
                    '记录数': len(df),
                })
            except Exception as e:
                file_info.append({
                    '文件名': file_path.split('\\')[-1],
                    '记录数': 0,
                    '错误': str(e),
                })

        if not all_data:
            return {
                '合并数据': pd.DataFrame(),
                '文件信息': file_info,
                '总记录数': 0,
            }

        merged_df = pd.concat(all_data, ignore_index=True)

        return {
            '合并数据': merged_df,
            '文件信息': file_info,
            '总记录数': len(merged_df),
        }

    def _read_file(self, file_path: str) -> pd.DataFrame:
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            return pd.read_excel(file_path, dtype={'身份证号': str, '联系电话': str})
        elif file_path.endswith('.csv'):
            return pd.read_csv(file_path, dtype={'身份证号': str, '联系电话': str})
        else:
            raise ValueError(f"不支持的文件格式: {file_path}")

    def remove_duplicates(self, df: pd.DataFrame) -> Dict:
        if df.empty:
            return {
                '去重后数据': df,
                '重复记录数': 0,
                '去重后记录数': 0,
                '重复明细': [],
            }

        df = df.copy()
        df['_原始行号'] = df.index + 2

        id_card_col = '身份证号'
        date_col = '随访日期'

        dup_mask = df.duplicated(subset=[id_card_col, date_col], keep=False)
        duplicate_rows = df[dup_mask].copy()

        if duplicate_rows.empty:
            return {
                '去重后数据': df.drop(columns=['_原始行号']),
                '重复记录数': 0,
                '去重后记录数': len(df),
                '重复明细': [],
            }

        duplicate_details = []
        grouped = duplicate_rows.groupby([id_card_col, date_col])

        for (id_card, date), group in grouped:
            if len(group) > 1:
                duplicate_details.append({
                    '身份证号': id_card,
                    '随访日期': date,
                    '姓名': group['姓名'].iloc[0] if '姓名' in group.columns else '',
                    '重复次数': len(group),
                    '涉及行号': ', '.join(map(str, group['_原始行号'].tolist())),
                    '来源文件': ', '.join(group['来源文件'].unique().tolist()) if '来源文件' in group.columns else '',
                })

        df_sorted = df.sort_values(by=date_col, ascending=False)
        df_deduped = df_sorted.drop_duplicates(subset=[id_card_col, date_col], keep='first')
        df_deduped = df_deduped.sort_index()
        df_deduped = df_deduped.drop(columns=['_原始行号'])

        return {
            '去重后数据': df_deduped,
            '重复记录数': len(df) - len(df_deduped),
            '去重后记录数': len(df_deduped),
            '重复明细': duplicate_details,
        }

    def merge_by_person(self, df: pd.DataFrame) -> Dict:
        if df.empty:
            return {
                '按人合并数据': df,
                '人数': 0,
                '总随访次数': 0,
            }

        id_card_col = '身份证号'
        date_col = '随访日期'

        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        person_summary = df.groupby(id_card_col).agg(
            姓名=('姓名', 'first'),
            联系电话=('联系电话', 'first'),
            村居=('村居', 'first'),
            随访次数=(date_col, 'count'),
            最近随访日期=(date_col, 'max'),
            最早随访日期=(date_col, 'min'),
        ).reset_index()

        person_summary['最近随访日期'] = person_summary['最近随访日期'].dt.strftime('%Y-%m-%d')
        person_summary['最早随访日期'] = person_summary['最早随访日期'].dt.strftime('%Y-%m-%d')

        total_visits = person_summary['随访次数'].sum()

        return {
            '按人合并数据': person_summary,
            '人数': len(person_summary),
            '总随访次数': total_visits,
        }
