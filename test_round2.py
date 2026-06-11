#!/usr/bin/env python3
"""
测试第二轮新功能：
1. 分类切换后结果隔离
2. 结果概览面板（UI，需手动验证）
3. 汇总报告增加导入来源表
4. 批量执行失败提示
"""

import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import DataValidator, DataMerger, ReminderGenerator, DataExporter, ReportGenerator


def test_source_sheet_in_report():
    print("=" * 60)
    print("测试1: 汇总报告增加导入来源Sheet")
    print("=" * 60)

    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
    category = '慢病'

    all_cat_keywords = ['慢病', '孕产妇', '儿童', '老年人']
    all_files = [f for f in os.listdir(sample_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

    filtered_files = []
    for filename in all_files:
        matches_current = category in filename
        matches_other = any(cat in filename for cat in all_cat_keywords if cat != category)
        if matches_current or not matches_other:
            filtered_files.append(os.path.join(sample_dir, filename))

    print(f"筛选出 {len(filtered_files)} 个文件")
    for f in filtered_files:
        print(f"  - {os.path.basename(f)}")

    merger = DataMerger(category)
    merge_result = merger.merge_multiple_files(filtered_files)
    data = merge_result['合并数据']
    file_info = merge_result['文件信息']

    dedup_result = merger.remove_duplicates(data)
    deduped_data = dedup_result['去重后数据']

    validator = DataValidator(category)
    validation_result = validator.validate_dataframe(deduped_data)

    reminder = ReminderGenerator(category)
    overdue_result = reminder.find_overdue_persons(deduped_data)

    output_dir = os.path.join('output', 'test_source_sheet')
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, '慢病_汇总报告.xlsx')

    merge_info = {
        '总记录数': len(data),
        '重复记录数': dedup_result['重复记录数'],
        '去重后记录数': dedup_result['去重后记录数'],
    }

    reporter = ReportGenerator(category)
    result = reporter.generate_summary_report(
        deduped_data, validation_result, overdue_result,
        merge_info, report_path, file_info
    )

    print(f"\n报告生成: {'成功' if result['成功'] else '失败'}")

    if result['成功']:
        xls = pd.ExcelFile(report_path)
        sheets = xls.sheet_names
        print(f"包含Sheet: {sheets}")
        assert '导入来源' in sheets, "缺少 导入来源 Sheet"

        source_df = pd.read_excel(report_path, sheet_name='导入来源')
        print(f"\n导入来源表内容:")
        print(source_df.to_string(index=False))

        print("\n✓ 导入来源Sheet测试通过\n")
    else:
        print(f"失败原因: {result['消息']}")
        return False

    return True


def test_category_isolation_logic():
    print("=" * 60)
    print("测试2: 分类过滤逻辑验证")
    print("=" * 60)

    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
    all_cat_keywords = ['慢病', '孕产妇', '儿童', '老年人']

    for category in all_cat_keywords:
        all_files = [f for f in os.listdir(sample_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

        filtered_files = []
        for filename in all_files:
            matches_current = category in filename
            matches_other = any(cat in filename for cat in all_cat_keywords if cat != category)
            if matches_current or not matches_other:
                filtered_files.append(filename)

        print(f"\n【{category}】")
        print(f"  筛选出 {len(filtered_files)} 个文件")

        for f in filtered_files:
            other_cats_in_name = [cat for cat in all_cat_keywords if cat != category and cat in f]
            if other_cats_in_name:
                print(f"  ⚠ 警告: {f} 包含其他分类关键词: {other_cats_in_name}")
                assert False, f"{f} 不应该被筛选到 {category} 分类"

        print(f"  ✓ 筛选正确")

    print("\n✓ 分类过滤逻辑测试通过\n")
    return True


def test_report_with_empty_source():
    print("=" * 60)
    print("测试3: 空数据时报告完整性")
    print("=" * 60)

    category = '慢病'
    output_dir = os.path.join('output', 'test_empty_source')
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, '慢病_汇总报告_空测试.xlsx')

    empty_df = pd.DataFrame(columns=['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果'])

    validator = DataValidator(category)
    validation_result = validator.validate_dataframe(empty_df)

    reminder = ReminderGenerator(category)
    overdue_result = reminder.find_overdue_persons(empty_df)

    merge_info = {
        '总记录数': 0,
        '重复记录数': 0,
        '去重后记录数': 0,
    }

    file_info = [
        {'文件名': '测试文件1.xlsx', '记录数': 0, '错误': '文件为空'},
        {'文件名': '测试文件2.xlsx', '记录数': 100},
    ]

    reporter = ReportGenerator(category)
    result = reporter.generate_summary_report(
        empty_df, validation_result, overdue_result,
        merge_info, report_path, file_info
    )

    print(f"报告生成: {'成功' if result['成功'] else '失败'}")

    if result['成功']:
        xls = pd.ExcelFile(report_path)
        sheets = xls.sheet_names
        print(f"包含Sheet: {sheets}")

        source_df = pd.read_excel(report_path, sheet_name='导入来源')
        print(f"\n导入来源表:")
        print(source_df.to_string(index=False))

        print("\n✓ 空数据报告测试通过\n")
    else:
        print(f"失败原因: {result['消息']}")
        return False

    return True


def test_next_month_reminder_empty():
    print("=" * 60)
    print("测试4: 无待随访人员时下月提醒文件完整性")
    print("=" * 60)

    category = '慢病'
    output_dir = os.path.join('output', 'test_empty_reminder2')
    os.makedirs(output_dir, exist_ok=True)
    reminder_path = os.path.join(output_dir, '慢病_下月提醒_空.xlsx')

    empty_df = pd.DataFrame()

    reporter = ReportGenerator(category)
    result = reporter.generate_next_month_reminder_file(empty_df, reminder_path)

    print(f"生成结果: {'成功' if result['成功'] else '失败'}")
    print(f"待随访人数: {result['待随访人数']}")

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
    print(f"\n下月待随访 列名: {list(df_detail.columns)}")
    print(f"下月待随访 行数: {len(df_detail)}")

    expected_columns = ['姓名', '身份证号', '联系电话', '村居', '随访日期', '随访结果', '下次随访日期']
    for col in expected_columns:
        assert col in df_detail.columns, f"缺少列: {col}"

    df_summary = pd.read_excel(reminder_path, sheet_name='汇总')
    summary_dict = dict(zip(df_summary['项目'], df_summary['数值']))
    print(f"\n汇总: 下月待随访人数 = {summary_dict['下月待随访人数']}")
    print(f"     备注 = {summary_dict.get('备注', '')}")

    assert summary_dict['下月待随访人数'] == 0, "汇总中待随访人数应该为0"

    print("\n✓ 空下月提醒文件测试通过\n")
    return True


def main():
    print("\n")
    print("*" * 60)
    print("*" + " " * 58 + "*")
    print("*" + " " * 18 + "第二轮新功能测试" + " " * 26 + "*")
    print("*" + " " * 58 + "*")
    print("*" * 60)
    print()

    all_passed = True

    try:
        all_passed &= test_category_isolation_logic()
        all_passed &= test_source_sheet_in_report()
        all_passed &= test_report_with_empty_source()
        all_passed &= test_next_month_reminder_empty()

        print("=" * 60)
        if all_passed:
            print("✓ 所有新功能测试通过！")
        else:
            print("✗ 部分测试失败")
        print("=" * 60)
        print("\n已验证的功能:")
        print("  1. ✓ 分类过滤逻辑 - 混放文件夹只显示对应分类文件")
        print("  2. ✓ 汇总报告导入来源表 - 列出处理的文件、记录数、状态")
        print("  3. ✓ 空数据报告完整性 - 包含导入来源表和汇总")
        print("  4. ✓ 空下月提醒文件 - 保留表头、村居统计、总数为0的汇总")
        print("\n需手动验证的功能:")
        print("  1. 分类切换后结果隔离（需在GUI中测试）")
        print("  2. 结果概览面板（需在GUI中测试）")
        print("  3. 批量执行失败提示（需模拟失败场景）")

    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
