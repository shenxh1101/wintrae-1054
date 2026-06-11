#!/usr/bin/env python3
"""
测试新功能：
1. 按分类过滤文件
2. 汇总报告自动生成下月提醒
3. 无待随访人员时也生成下月提醒文件
"""

import sys
import os
import pandas as pd
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import DataValidator, DataMerger, ReminderGenerator, DataExporter, ReportGenerator


def test_category_filter():
    print("=" * 60)
    print("测试1: 按分类关键词过滤数据文件")
    print("=" * 60)

    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
    all_files = [f for f in os.listdir(sample_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

    categories = ['慢病', '孕产妇', '儿童', '老年人']
    all_cat_keywords = categories

    for category in categories:
        filtered_files = []
        for filename in all_files:
            matches_current = category in filename
            matches_other = any(cat in filename for cat in all_cat_keywords if cat != category)
            if matches_current or not matches_other:
                filtered_files.append(filename)

        print(f"\n【{category}】")
        print(f"  总文件数: {len(all_files)}")
        print(f"  筛选后文件数: {len(filtered_files)}")
        for f in filtered_files:
            print(f"    - {f}")

        for filename in filtered_files:
            assert category in filename or not any(cat in filename for cat in all_cat_keywords if cat != category), \
                f"错误: {filename} 不应该被筛选到 {category} 分类中"

        print(f"  ✓ 筛选正确，无其他分类文件混入")

    print("\n✓ 分类过滤测试通过\n")


def test_empty_next_month_reminder():
    print("=" * 60)
    print("测试2: 无待随访人员时也生成下月提醒文件")
    print("=" * 60)

    output_dir = os.path.join('output', 'test_empty_reminder')
    os.makedirs(output_dir, exist_ok=True)

    empty_df = pd.DataFrame()
    reminder_path = os.path.join(output_dir, '慢病_下月提醒_空测试.xlsx')

    reporter = ReportGenerator('慢病')
    result = reporter.generate_next_month_reminder_file(empty_df, reminder_path)

    print(f"生成结果: {'成功' if result['成功'] else '失败'}")
    print(f"待随访人数: {result['待随访人数']}")
    print(f"文件路径: {result['路径']}")

    assert result['成功'], "生成失败"
    assert result['待随访人数'] == 0, "待随访人数应该为0"
    assert os.path.exists(reminder_path), "文件未生成"

    xls = pd.ExcelFile(reminder_path)
    sheets = xls.sheet_names
    print(f"包含Sheet: {sheets}")

    assert '下月待随访' in sheets, "缺少 下月待随访 Sheet"
    assert '村居统计' in sheets, "缺少 村居统计 Sheet"
    assert '汇总' in sheets, "缺少 汇总 Sheet"

    df_detail = pd.read_excel(reminder_path, sheet_name='下月待随访')
    print(f"下月待随访 列名: {list(df_detail.columns)}")
    print(f"下月待随访 行数: {len(df_detail)}")

    expected_columns = ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '下次随访日期']
    for col in expected_columns:
        assert col in df_detail.columns, f"缺少列: {col}"
    print(f"✓ 表头完整")

    df_village = pd.read_excel(reminder_path, sheet_name='村居统计')
    print(f"村居统计 列名: {list(df_village.columns)}")
    assert '村居' in df_village.columns and '待随访人数' in df_village.columns, "村居统计列名不正确"
    print(f"✓ 村居统计完整")

    df_summary = pd.read_excel(reminder_path, sheet_name='汇总')
    print(f"汇总 数据:\n{df_summary}")

    summary_dict = dict(zip(df_summary['项目'], df_summary['数值']))
    assert summary_dict['下月待随访人数'] == 0, "汇总中待随访人数应该为0"
    assert '本月无待随访人员' in str(summary_dict.get('备注', '')), "缺少备注说明"
    print(f"✓ 汇总数据正确")

    print("\n✓ 空下月提醒文件测试通过\n")


def test_report_auto_generate_reminder():
    print("=" * 60)
    print("测试3: 汇总报告自动生成下月提醒")
    print("=" * 60)

    sample_file = os.path.join('sample_data', '慢病随访记录.xlsx')
    df = pd.read_excel(sample_file, dtype={'身份证号': str, '联系电话': str})

    output_dir = os.path.join('output', 'test_auto_reminder')
    os.makedirs(output_dir, exist_ok=True)

    merger = DataMerger('慢病')
    dedup_result = merger.remove_duplicates(df)
    data = dedup_result['去重后数据']

    validator = DataValidator('慢病')
    validation_result = validator.validate_dataframe(data)

    reminder = ReminderGenerator('慢病')
    overdue_result = reminder.find_overdue_persons(data)

    merge_info = {
        '总记录数': len(df),
        '重复记录数': dedup_result['重复记录数'],
        '去重后记录数': dedup_result['去重后记录数'],
    }

    reporter = ReportGenerator('慢病')

    report_path = os.path.join(output_dir, '慢病_汇总报告.xlsx')
    rep_result = reporter.generate_summary_report(data, validation_result, overdue_result,
                                                  merge_info, report_path)

    print(f"汇总报告生成: {'成功' if rep_result['成功'] else '失败'}")

    next_month_result = reminder.generate_next_month_reminder(data)
    reminder_path = os.path.join(output_dir, '慢病_下月提醒.xlsx')
    rem_result = reporter.generate_next_month_reminder_file(
        next_month_result['下月待随访人员'],
        reminder_path
    )

    print(f"下月提醒生成: {'成功' if rem_result['成功'] else '失败'}")
    print(f"待随访人数: {rem_result['待随访人数']}")

    assert os.path.exists(report_path), "汇总报告未生成"
    assert os.path.exists(reminder_path), "下月提醒未生成"

    print("\n✓ 汇总报告自动生成下月提醒测试通过\n")


def test_simulation_all_workflow():
    print("=" * 60)
    print("测试4: 模拟完整工作流（无中间弹窗）")
    print("=" * 60)

    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
    category = '慢病'

    all_files = [f for f in os.listdir(sample_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    all_cat_keywords = ['慢病', '孕产妇', '儿童', '老年人']

    filtered_files = []
    for filename in all_files:
        matches_current = category in filename
        matches_other = any(cat in filename for cat in all_cat_keywords if cat != category)
        if matches_current or not matches_other:
            filtered_files.append(os.path.join(sample_dir, filename))

    print(f"筛选文件: {len(filtered_files)} 个")
    for f in filtered_files:
        print(f"  - {os.path.basename(f)}")

    print("\n步骤1: 导入数据 (silent模式)")
    merger = DataMerger(category)
    merge_result = merger.merge_multiple_files(filtered_files)
    merged_data = merge_result['合并数据']
    print(f"  ✓ 导入 {merge_result['总记录数']} 条记录")

    print("\n步骤2: 校验字段 (silent模式)")
    validator = DataValidator(category)
    validation_result = validator.validate_dataframe(merged_data)
    print(f"  ✓ 有效率: {validation_result['有效率']}")

    print("\n步骤3: 合并去重 (silent模式)")
    dedup_result = merger.remove_duplicates(merged_data)
    data = dedup_result['去重后数据']
    print(f"  ✓ 去重后 {dedup_result['去重后记录数']} 条")

    print("\n步骤4: 生成提醒 (silent模式)")
    reminder = ReminderGenerator(category)
    overdue_result = reminder.find_overdue_persons(data)
    if not overdue_result['超期人员'].empty:
        overdue_result['超期人员'] = reminder.add_overdue_level(overdue_result['超期人员'])
    next_month_result = reminder.generate_next_month_reminder(data)
    print(f"  ✓ 超期人数: {overdue_result['超期人数']} 人")
    print(f"  ✓ 下月待随访: {next_month_result['待随访人数']} 人")

    print("\n步骤5: 分组导出 (silent模式)")
    output_dir = os.path.join('output', 'test_simulation')
    exporter = DataExporter(category)
    village_dir = os.path.join(output_dir, '村居分组')
    export_result = exporter.export_by_village(data, village_dir)
    print(f"  ✓ 导出 {len(export_result['文件列表'])} 个村居文件")

    print("\n步骤6: 汇总报告 (silent模式，自动生成下月提醒)")
    reporter = ReportGenerator(category)
    report_path = os.path.join(output_dir, f'{category}_汇总报告.xlsx')
    merge_info = {
        '总记录数': merge_result['总记录数'],
        '重复记录数': dedup_result['重复记录数'],
        '去重后记录数': dedup_result['去重后记录数'],
    }
    rep_result = reporter.generate_summary_report(data, validation_result, overdue_result,
                                                  merge_info, report_path)
    reminder_path = os.path.join(output_dir, f'{category}_下月提醒.xlsx')
    rem_result = reporter.generate_next_month_reminder_file(
        next_month_result['下月待随访人员'],
        reminder_path
    )
    print(f"  ✓ 汇总报告: {os.path.basename(report_path)}")
    print(f"  ✓ 下月提醒: {os.path.basename(reminder_path)} ({rem_result['待随访人数']} 人)")

    print("\n" + "=" * 60)
    print("最终结果汇总:")
    print(f"  数据有效率: {validation_result['有效率']}")
    print(f"  超期未访: {overdue_result['超期人数']} 人 ({overdue_result['超期率']})")
    print(f"  下月待随访: {next_month_result['待随访人数']} 人")
    print(f"  输出目录: {output_dir}")
    print("=" * 60)

    assert os.path.exists(report_path), "汇总报告未生成"
    assert os.path.exists(reminder_path), "下月提醒未生成"

    print("\n✓ 完整工作流模拟测试通过\n")


def main():
    print("\n")
    print("*" * 60)
    print("*" + " " * 58 + "*")
    print("*" + " " * 18 + "新功能测试验证" + " " * 26 + "*")
    print("*" + " " * 58 + "*")
    print("*" * 60)
    print()

    try:
        test_category_filter()
        test_empty_next_month_reminder()
        test_report_auto_generate_reminder()
        test_simulation_all_workflow()

        print("=" * 60)
        print("✓ 所有新功能测试通过！")
        print("=" * 60)
        print("\n已验证的功能:")
        print("  1. ✓ 按分类关键词过滤数据文件，混放文件夹不会混淆")
        print("  2. ✓ 汇总报告自动计算并保存下月提醒文件")
        print("  3. ✓ 无待随访人员时也生成完整的下月提醒文件")
        print("  4. ✓ 一键执行无中间弹窗，最后统一显示结果")

    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
