import re
import pandas as pd
from datetime import datetime
from typing import Dict, List, Tuple


class DataValidator:
    CATEGORY_CONFIGS = {
        '慢病': {
            'required_fields': ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果'],
            'followup_interval_days': 90,
        },
        '孕产妇': {
            'required_fields': ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '孕周'],
            'followup_interval_days': 30,
        },
        '儿童': {
            'required_fields': ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '年龄'],
            'followup_interval_days': 90,
        },
        '老年人': {
            'required_fields': ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '年龄'],
            'followup_interval_days': 180,
        },
    }

    def __init__(self, category: str):
        if category not in self.CATEGORY_CONFIGS:
            raise ValueError(f"不支持的分类: {category}")
        self.category = category
        self.config = self.CATEGORY_CONFIGS[category]

    def validate_id_card(self, id_card: str) -> Tuple[bool, str]:
        if pd.isna(id_card) or not str(id_card).strip():
            return False, "身份证号缺失"
        id_str = str(id_card).strip()
        if len(id_str) not in (15, 18):
            return False, "身份证号长度不正确"
        if len(id_str) == 18:
            pattern = r'^[1-9]\d{5}(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{3}[\dXx]$'
            if not re.match(pattern, id_str):
                return False, "身份证号格式不正确"
            if not self._verify_id_checksum(id_str):
                return False, "身份证号校验码错误"
        return True, ""

    def _verify_id_checksum(self, id18: str) -> bool:
        weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
        check_codes = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
        total = sum(int(id18[i]) * weights[i] for i in range(17))
        expected = check_codes[total % 11]
        return id18[17].upper() == expected

    def validate_phone(self, phone: str) -> Tuple[bool, str]:
        if pd.isna(phone) or not str(phone).strip():
            return False, "联系电话缺失"
        phone_str = str(phone).strip()
        pattern = r'^1[3-9]\d{9}$|^0\d{2,3}-?\d{7,8}$'
        if not re.match(pattern, phone_str):
            return False, "电话号码格式不正确"
        return True, ""

    def validate_date(self, date_val) -> Tuple[bool, str]:
        if pd.isna(date_val) or (isinstance(date_val, str) and not date_val.strip()):
            return False, "随访日期缺失"
        try:
            if isinstance(date_val, datetime):
                return True, ""
            date_str = str(date_val).strip()
            for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y年%m月%d日']:
                try:
                    datetime.strptime(date_str, fmt)
                    return True, ""
                except ValueError:
                    continue
            return False, "随访日期格式不正确"
        except Exception:
            return False, "随访日期格式不正确"

    def validate_followup_result(self, result: str) -> Tuple[bool, str]:
        if pd.isna(result) or not str(result).strip():
            return False, "随访结果缺失"
        result_str = str(result).strip()
        valid_results = ['满意', '基本满意', '不满意', '完成', '未完成', '正常', '异常']
        if result_str not in valid_results:
            return True, ""
        return True, ""

    def validate_row(self, row: pd.Series) -> Dict:
        errors = []
        valid = True

        id_valid, id_msg = self.validate_id_card(row.get('身份证号', ''))
        if not id_valid:
            valid = False
            errors.append(id_msg)

        phone_valid, phone_msg = self.validate_phone(row.get('联系电话', ''))
        if not phone_valid:
            valid = False
            errors.append(phone_msg)

        date_valid, date_msg = self.validate_date(row.get('随访日期', ''))
        if not date_valid:
            valid = False
            errors.append(date_msg)

        result_valid, result_msg = self.validate_followup_result(row.get('随访结果', ''))
        if not result_valid:
            valid = False
            errors.append(result_msg)

        for field in self.config['required_fields']:
            if field not in row.index or pd.isna(row.get(field, '')) or (
                    isinstance(row.get(field, ''), str) and not str(row.get(field, '')).strip()):
                if field not in ['身份证号', '联系电话', '随访日期', '随访结果']:
                    valid = False
                    errors.append(f"{field}缺失")

        return {
            'valid': valid,
            'errors': errors,
            'error_count': len(errors),
        }

    def validate_dataframe(self, df: pd.DataFrame) -> Dict:
        total_rows = len(df)
        valid_rows = 0
        invalid_rows = 0
        error_details = []
        error_summary = {}

        for idx, row in df.iterrows():
            result = self.validate_row(row)
            if result['valid']:
                valid_rows += 1
            else:
                invalid_rows += 1
                error_details.append({
                    '行号': idx + 2,
                    '姓名': row.get('姓名', ''),
                    '身份证号': row.get('身份证号', ''),
                    '错误数量': result['error_count'],
                    '错误详情': '; '.join(result['errors']),
                })
                for err in result['errors']:
                    error_summary[err] = error_summary.get(err, 0) + 1

        return {
            '总记录数': total_rows,
            '有效记录数': valid_rows,
            '无效记录数': invalid_rows,
            '有效率': f"{(valid_rows / total_rows * 100):.2f}%" if total_rows > 0 else "0%",
            '错误明细': error_details,
            '错误汇总': error_summary,
        }

    def get_missing_fields(self, df: pd.DataFrame) -> List[str]:
        missing = []
        for field in self.config['required_fields']:
            if field not in df.columns:
                missing.append(field)
        return missing
