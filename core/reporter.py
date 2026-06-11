import pandas as pd
import os
from datetime import datetime
from typing import Dict


class ReportGenerator:
    def __init__(self, category: str):
        self.category = category

    def generate_completion_report(self, df: pd.DataFrame, reference_date: datetime = None) -> Dict:
        if reference_date is None:
            reference_date = datetime.now()

        if df.empty:
            return {
                '总人数': 0,
                '已随访人数': 0,
                '未随访人数': 0,
                '完成率': '0%',
                '村居统计': pd.DataFrame(),
                '生成时间': reference_date.strftime('%Y-%m-%d %H:%M:%S'),
            }

        df = df.copy()
        df['随访日期'] = pd.to_datetime(df['随访日期'], errors='coerce')

        df_sorted = df.sort_values('随访日期', ascending=False)
        latest_visits = df_sorted.drop_duplicates(subset=['身份证号'], keep='first')

        total_persons = len(latest_visits)
        completed_persons = len(latest_visits[latest_visits['随访结果'].notna() & (latest_visits['随访结果'] != '')])

        completion_rate = f"{(completed_persons / total_persons * 100):.2f}%" if total_persons > 0 else "0%"

        village_stats = self._get_village_completion_stats(latest_visits)

        return {
            '总人数': total_persons,
            '已随访人数': completed_persons,
            '未随访人数': total_persons - completed_persons,
            '完成率': completion_rate,
            '村居统计': village_stats,
            '生成时间': reference_date.strftime('%Y-%m-%d %H:%M:%S'),
        }

    def _get_village_completion_stats(self, latest_visits: pd.DataFrame) -> pd.DataFrame:
        if latest_visits.empty:
            return pd.DataFrame()

        village_stats = latest_visits.groupby('村居').agg(
            总人数=('身份证号', 'count'),
            已随访人数=('随访结果', lambda x: x.notna().sum()),
        ).reset_index()

        village_stats['完成率'] = village_stats.apply(
            lambda row: f"{(row['已随访人数'] / row['总人数'] * 100):.2f}%" if row['总人数'] > 0 else '0%',
            axis=1
        )
        village_stats = village_stats.sort_values('总人数', ascending=False)

        return village_stats

    def generate_summary_report(self, df: pd.DataFrame, validation_result: Dict,
                                overdue_result: Dict, merge_result: Dict,
                                output_path: str, file_info_list: list = None,
                                next_month_df: pd.DataFrame = None) -> Dict:
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                summary_df = self._create_summary_sheet(df, validation_result, overdue_result, merge_result)
                summary_df.to_excel(writer, index=False, sheet_name='汇总概览')

                if file_info_list:
                    source_df = self._create_source_sheet(file_info_list)
                    source_df.to_excel(writer, index=False, sheet_name='导入来源')

                village_review_df = self._create_village_review_sheet(
                    df, validation_result, overdue_result, next_month_df
                )
                village_review_df.to_excel(writer, index=False, sheet_name='村居复核')

                if validation_result.get('错误明细'):
                    error_df = pd.DataFrame(validation_result['错误明细'])
                    error_df.to_excel(writer, index=False, sheet_name='异常明细')

                if validation_result.get('错误汇总'):
                    error_summary_df = pd.DataFrame(
                        list(validation_result['错误汇总'].items()),
                        columns=['错误类型', '数量']
                    ).sort_values('数量', ascending=False)
                    error_summary_df.to_excel(writer, index=False, sheet_name='错误统计')

                if not overdue_result.get('超期人员', pd.DataFrame()).empty:
                    overdue_df = overdue_result['超期人员']
                    if '超期等级' not in overdue_df.columns:
                        from .reminder import ReminderGenerator
                        rg = ReminderGenerator(self.category)
                        overdue_df = rg.add_overdue_level(overdue_df)
                    overdue_df.to_excel(writer, index=False, sheet_name='超期未访')

                village_stats = self._get_village_completion_stats(
                    self._get_latest_visits(df)
                )
                if not village_stats.empty:
                    village_stats.to_excel(writer, index=False, sheet_name='村居统计')

            return {
                '成功': True,
                '消息': f'汇总报告已生成: {output_path}',
                '路径': output_path,
            }
        except Exception as e:
            return {
                '成功': False,
                '消息': f'生成报告失败: {str(e)}',
                '路径': output_path,
            }

    def _create_source_sheet(self, file_info_list: list) -> pd.DataFrame:
        data = []
        for idx, info in enumerate(file_info_list, 1):
            if '错误' in info:
                status = '失败'
                record_count = 0
                error_msg = info['错误']
            else:
                status = '成功'
                record_count = info['记录数']
                error_msg = ''

            data.append({
                '序号': idx,
                '文件名': info['文件名'],
                '记录数': record_count,
                '状态': status,
                '错误信息': error_msg,
            })

        total_success = sum(1 for info in file_info_list if '错误' not in info)
        total_fail = sum(1 for info in file_info_list if '错误' in info)
        total_records = sum(info.get('记录数', 0) for info in file_info_list)

        data.append({
            '序号': '',
            '文件名': '合计',
            '记录数': total_records,
            '状态': f'成功 {total_success} 个，失败 {total_fail} 个',
            '错误信息': '',
        })

        return pd.DataFrame(data)

    def _create_village_review_sheet(self, df: pd.DataFrame, validation_result: Dict,
                                     overdue_result: Dict, next_month_df: pd.DataFrame = None) -> pd.DataFrame:
        latest_visits = self._get_latest_visits(df)

        if latest_visits.empty:
            return pd.DataFrame(columns=['村居', '总人数', '有效记录数', '超期人数', '下月待随访人数', '异常记录数'])

        village_persons = latest_visits.groupby('村居')['身份证号'].nunique().reset_index()
        village_persons.columns = ['村居', '总人数']

        error_detail = validation_result.get('错误明细', [])
        if error_detail:
            error_df = pd.DataFrame(error_detail)
            if '村居' in error_df.columns:
                village_errors = error_df.groupby('村居').size().reset_index(name='异常记录数')
            else:
                village_errors = pd.DataFrame(columns=['村居', '异常记录数'])
        else:
            village_errors = pd.DataFrame(columns=['村居', '异常记录数'])

        valid_records = validation_result.get('有效记录数', 0)
        if valid_records > 0 and not latest_visits.empty:
            village_valid = latest_visits.groupby('村居').size().reset_index(name='有效记录数')
            village_valid['有效记录数'] = village_valid['有效记录数'].astype(int)
        else:
            village_valid = pd.DataFrame(columns=['村居', '有效记录数'])

        overdue_df = overdue_result.get('超期人员', pd.DataFrame())
        if not overdue_df.empty and '村居' in overdue_df.columns:
            village_overdue = overdue_df.groupby('村居')['身份证号'].nunique().reset_index()
            village_overdue.columns = ['村居', '超期人数']
        else:
            village_overdue = pd.DataFrame(columns=['村居', '超期人数'])

        if next_month_df is not None and not next_month_df.empty and '村居' in next_month_df.columns:
            village_next = next_month_df.groupby('村居')['身份证号'].nunique().reset_index()
            village_next.columns = ['村居', '下月待随访人数']
        else:
            village_next = pd.DataFrame(columns=['村居', '下月待随访人数'])

        review_df = village_persons
        for other_df in [village_valid, village_overdue, village_next, village_errors]:
            if not other_df.empty:
                review_df = review_df.merge(other_df, on='村居', how='left')

        for col in ['有效记录数', '超期人数', '下月待随访人数', '异常记录数']:
            if col not in review_df.columns:
                review_df[col] = 0
            else:
                review_df[col] = review_df[col].fillna(0).astype(int)

        review_df = review_df.sort_values('总人数', ascending=False).reset_index(drop=True)

        total_row = pd.DataFrame([{
            '村居': '合计',
            '总人数': review_df['总人数'].sum(),
            '有效记录数': review_df['有效记录数'].sum(),
            '超期人数': review_df['超期人数'].sum(),
            '下月待随访人数': review_df['下月待随访人数'].sum(),
            '异常记录数': review_df['异常记录数'].sum(),
        }])

        review_df = pd.concat([review_df, total_row], ignore_index=True)

        return review_df

    def _get_latest_visits(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        df = df.copy()
        df['随访日期'] = pd.to_datetime(df['随访日期'], errors='coerce')
        df_sorted = df.sort_values('随访日期', ascending=False)
        return df_sorted.drop_duplicates(subset=['身份证号'], keep='first')

    def _create_summary_sheet(self, df: pd.DataFrame, validation_result: Dict,
                              overdue_result: Dict, merge_result: Dict) -> pd.DataFrame:
        latest_visits = self._get_latest_visits(df)
        total_persons = len(latest_visits)

        completed = len(latest_visits[latest_visits['随访结果'].notna() & (latest_visits['随访结果'] != '')])
        completion_rate = f"{(completed / total_persons * 100):.2f}%" if total_persons > 0 else "0%"

        summary_data = [
            ['分类', self.category],
            ['生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            ['', ''],
            ['【数据质量】', ''],
            ['总记录数', validation_result.get('总记录数', 0)],
            ['有效记录数', validation_result.get('有效记录数', 0)],
            ['无效记录数', validation_result.get('无效记录数', 0)],
            ['数据有效率', validation_result.get('有效率', '0%')],
            ['', ''],
            ['【合并去重】', ''],
            ['合并前记录数', merge_result.get('总记录数', 0) if merge_result else 0],
            ['重复记录数', merge_result.get('重复记录数', 0) if merge_result else 0],
            ['去重后记录数', merge_result.get('去重后记录数', 0) if merge_result else 0],
            ['', ''],
            ['【随访情况】', ''],
            ['管理总人数', total_persons],
            ['已随访人数', completed],
            ['完成率', completion_rate],
            ['', ''],
            ['【超期情况】', ''],
            ['总人数', overdue_result.get('总人数', 0)],
            ['超期未访人数', overdue_result.get('超期人数', 0)],
            ['超期率', overdue_result.get('超期率', '0%')],
            ['随访间隔(天)', overdue_result.get('随访间隔天数', '-')],
            ['参考日期', overdue_result.get('参考日期', '-')],
        ]

        return pd.DataFrame(summary_data, columns=['项目', '数值'])

    def generate_next_month_reminder_file(self, next_month_df: pd.DataFrame,
                                          output_path: str) -> Dict:
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                default_columns = ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '下次随访日期']
                if next_month_df.empty:
                    next_month_df = pd.DataFrame(columns=default_columns)
                next_month_df.to_excel(writer, index=False, sheet_name='下月待随访')

                if next_month_df.empty or '村居' not in next_month_df.columns:
                    village_stats = pd.DataFrame(columns=['村居', '待随访人数'])
                else:
                    village_stats = next_month_df.groupby('村居').agg(
                        待随访人数=('身份证号', 'count'),
                    ).reset_index().sort_values('待随访人数', ascending=False)
                    if village_stats.empty:
                        village_stats = pd.DataFrame(columns=['村居', '待随访人数'])
                village_stats.to_excel(writer, index=False, sheet_name='村居统计')

                total_persons = len(next_month_df) if not next_month_df.empty else 0
                summary_data = {
                    '项目': ['分类', '总人数', '下月待随访人数', '生成时间', '备注'],
                    '数值': [
                        self.category,
                        total_persons,
                        total_persons,
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        '本月无待随访人员' if total_persons == 0 else '',
                    ]
                }
                pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='汇总')

            return {
                '成功': True,
                '消息': f'下月提醒文件已生成: {output_path}',
                '待随访人数': len(next_month_df) if not next_month_df.empty else 0,
                '路径': output_path,
            }
        except Exception as e:
            return {
                '成功': False,
                '消息': f'生成下月提醒失败: {str(e)}',
                '路径': output_path,
            }
