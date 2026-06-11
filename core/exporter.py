import pandas as pd
import os
from typing import Dict, List


class DataExporter:
    def __init__(self, category: str):
        self.category = category

    def export_by_village(self, df: pd.DataFrame, output_dir: str) -> Dict:
        if df.empty:
            return {
                '成功': False,
                '消息': '数据为空，无法导出',
                '文件列表': [],
            }

        os.makedirs(output_dir, exist_ok=True)
        village_col = '村居'

        if village_col not in df.columns:
            return {
                '成功': False,
                '消息': '数据中缺少"村居"字段',
                '文件列表': [],
            }

        exported_files = []
        villages = df[village_col].fillna('未知村居').unique()

        for village in villages:
            village_df = df[df[village_col].fillna('未知村居') == village].copy()
            safe_village_name = str(village).replace('/', '_').replace('\\', '_')
            filename = f"{self.category}_{safe_village_name}.xlsx"
            filepath = os.path.join(output_dir, filename)

            try:
                village_df.to_excel(filepath, index=False, sheet_name='随访记录')
                exported_files.append({
                    '村居': village,
                    '文件名': filename,
                    '记录数': len(village_df),
                    '路径': filepath,
                })
            except Exception as e:
                exported_files.append({
                    '村居': village,
                    '文件名': filename,
                    '记录数': len(village_df),
                    '错误': str(e),
                })

        return {
            '成功': True,
            '消息': f'成功导出 {len(exported_files)} 个村居文件',
            '文件列表': exported_files,
            '输出目录': output_dir,
        }

    def export_overdue_by_village(self, overdue_df: pd.DataFrame, output_dir: str) -> Dict:
        if overdue_df.empty:
            return {
                '成功': False,
                '消息': '无超期人员数据',
                '文件列表': [],
            }

        os.makedirs(output_dir, exist_ok=True)
        village_col = '村居'

        if village_col not in overdue_df.columns:
            return {
                '成功': False,
                '消息': '数据中缺少"村居"字段',
                '文件列表': [],
            }

        exported_files = []
        villages = overdue_df[village_col].fillna('未知村居').unique()

        for village in villages:
            village_df = overdue_df[overdue_df[village_col].fillna('未知村居') == village].copy()
            safe_village_name = str(village).replace('/', '_').replace('\\', '_')
            filename = f"{self.category}_{safe_village_name}_超期待随访.xlsx"
            filepath = os.path.join(output_dir, filename)

            try:
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                    village_df.to_excel(writer, index=False, sheet_name='超期人员清单')

                    summary_data = {
                        '项目': ['村居', '超期人数', '轻度超期(≤7天)', '中度超期(8-30天)', '重度超期(>30天)'],
                        '数值': [
                            village,
                            len(village_df),
                            len(village_df[village_df['超期天数'] <= 7]) if '超期天数' in village_df.columns else '-',
                            len(village_df[(village_df['超期天数'] > 7) & (village_df['超期天数'] <= 30)]) if '超期天数' in village_df.columns else '-',
                            len(village_df[village_df['超期天数'] > 30]) if '超期天数' in village_df.columns else '-',
                        ]
                    }
                    pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='统计汇总')

                exported_files.append({
                    '村居': village,
                    '文件名': filename,
                    '超期人数': len(village_df),
                    '路径': filepath,
                })
            except Exception as e:
                exported_files.append({
                    '村居': village,
                    '文件名': filename,
                    '超期人数': len(village_df),
                    '错误': str(e),
                })

        return {
            '成功': True,
            '消息': f'成功导出 {len(exported_files)} 个村居超期随访清单',
            '文件列表': exported_files,
            '输出目录': output_dir,
        }

    def export_single_file(self, df: pd.DataFrame, output_path: str, sheet_name: str = '数据') -> Dict:
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            df.to_excel(output_path, index=False, sheet_name=sheet_name)
            return {
                '成功': True,
                '消息': f'成功导出文件: {output_path}',
                '记录数': len(df),
                '路径': output_path,
            }
        except Exception as e:
            return {
                '成功': False,
                '消息': f'导出失败: {str(e)}',
                '记录数': len(df),
                '路径': output_path,
            }
