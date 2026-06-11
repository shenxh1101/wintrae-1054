#!/usr/bin/env python3
"""
测试核心模块功能
"""

import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import DataValidator, DataMerger, ReminderGenerator, DataExporter, ReportGenerator


def test_validator():
    print("=" * 60)
    print("测试数据校验模块...")
    print("=" * 60)

    sample_file = os.path.join('sample_data', '慢病随访记录.xlsx')
    df = pd.read_excel(sample_file, dtype={'身份证号': str, '联系电话': str})

    validator = DataValidator('慢病')
    result = validator.validate_dataframe(df)

    print(f"总记录数: {result['总记录数']}")
    print(f"有效记录数: {result['有效记录数']}")
    print(f"无效记录数: {result['无效记录数']}")
    print(f"有效率: {result['有效率']}")
    print(f"错误类型数量: {len(result['错误汇总'])}")
    for err_type, count in list(result['错误汇总'].items())[:5]:
        print(f"  - {err_type}: {count}")

    missing = validator.get_missing_fields(df)
    print(f"缺失字段: {missing if missing else '无'}")

    print("✓ 校验模块测试通过\n")
    return result


def test_merger():
    print("=" * 60)
    print("测试数据合并模块...")
    print("=" * 60)

    files = [
        os.path.join('sample_data', '慢病随访记录.xlsx'),
        os.path.join('sample_data', '慢病随访记录_第二批.xlsx'),
    ]

    merger = DataMerger('慢病')
    merge_result = merger.merge_multiple_files(files)

    print(f"合并后记录数: {merge_result['总记录数']}")
    for info in merge_result['文件信息']:
        print(f"  - {info['文件名']}: {info['记录数']} 条")

    dedup_result = merger.remove_duplicates(merge_result['合并数据'])
    print(f"去重后记录数: {dedup_result['去重后记录数']}")
    print(f"重复记录数: {dedup_result['重复记录数']}")
    print(f"发现重复组数: {len(dedup_result['重复明细'])}")

    person_result = merger.merge_by_person(dedup_result['去重后数据'])
    print(f"管理人数: {person_result['人数']}")
    print(f"总随访次数: {person_result['总随访次数']}")

    print("✓ 合并模块测试通过\n")
    return dedup_result['去重后数据'], merge_result, dedup_result


def test_reminder(df):
    print("=" * 60)
    print("测试提醒生成模块...")
    print("=" * 60)

    reminder = ReminderGenerator('慢病')

    overdue_result = reminder.find_overdue_persons(df)
    print(f"管理总人数: {overdue_result['总人数']}")
    print(f"超期人数: {overdue_result['超期人数']}")
    print(f"超期率: {overdue_result['超期率']}")
    print(f"参考日期: {overdue_result['参考日期']}")
    print(f"随访间隔: {overdue_result['随访间隔天数']} 天")

    overdue_with_level = reminder.add_overdue_level(overdue_result['超期人员'])
    if not overdue_with_level.empty:
        level_counts = overdue_with_level['超期等级'].value_counts()
        print("超期等级分布:")
        for level, count in level_counts.items():
            print(f"  - {level}: {count} 人")

    next_month_result = reminder.generate_next_month_reminder(df)
    print(f"下月待随访人数: {next_month_result['待随访人数']}")

    village_result = reminder.generate_village_reminder(df)
    print(f"涉及村居数: {village_result['村居数量']}")

    print("✓ 提醒模块测试通过\n")
    return overdue_result, next_month_result


def test_exporter(df, overdue_df):
    print("=" * 60)
    print("测试分组导出模块...")
    print("=" * 60)

    output_dir = os.path.join('output', 'test_export')
    exporter = DataExporter('慢病')

    result = exporter.export_by_village(df, output_dir)
    print(f"按村居导出: {'成功' if result['成功'] else '失败'}")
    if result['成功']:
        print(f"导出文件数: {len(result['文件列表'])}")
        for f in result['文件列表'][:5]:
            print(f"  - {f['文件名']}: {f['记录数']} 条")

    if not overdue_df.empty:
        overdue_dir = os.path.join(output_dir, 'overdue')
        overdue_result = exporter.export_overdue_by_village(overdue_df, overdue_dir)
        print(f"超期清单导出: {'成功' if overdue_result['成功'] else '失败'}")
        if overdue_result['成功']:
            print(f"超期文件数: {len(overdue_result['文件列表'])}")

    print("✓ 导出模块测试通过\n")


def test_reporter(df, validation_result, merge_result, dedup_result, overdue_result):
    print("=" * 60)
    print("测试报告生成模块...")
    print("=" * 60)

    reporter = ReportGenerator('慢病')

    output_dir = os.path.join('output', 'test_report')
    os.makedirs(output_dir, exist_ok=True)

    report_path = os.path.join(output_dir, '慢病_汇总报告.xlsx')
    merge_info = {
        '总记录数': merge_result['总记录数'],
        '重复记录数': dedup_result['重复记录数'],
        '去重后记录数': dedup_result['去重后记录数'],
    }
    result = reporter.generate_summary_report(df, validation_result, overdue_result,
                                              merge_info, report_path)
    print(f"汇总报告生成: {'成功' if result['成功'] else '失败'}")
    print(f"报告路径: {report_path}")

    completion = reporter.generate_completion_report(df)
    print(f"\n完成率统计:")
    print(f"  总人数: {completion['总人数']}")
    print(f"  已随访: {completion['已随访人数']}")
    print(f"  完成率: {completion['完成率']}")
    print(f"  涉及村居: {len(completion['村居统计'])} 个")

    print("✓ 报告模块测试通过\n")


def main():
    print("\n")
    print("*" * 60)
    print("*" + " " * 58 + "*")
    print("*" + " " * 15 + "公共卫生管理自动化工具 - 核心模块测试" + " " * 9 + "*")
    print("*" + " " * 58 + "*")
    print("*" * 60)
    print()

    try:
        validation_result = test_validator()
        deduped_df, merge_result, dedup_result = test_merger()
        overdue_result, next_month_result = test_reminder(deduped_df)

        overdue_df = overdue_result['超期人员']
        if not overdue_df.empty:
            from core.reminder import ReminderGenerator
            rg = ReminderGenerator('慢病')
            overdue_df = rg.add_overdue_level(overdue_df)

        test_exporter(deduped_df, overdue_df)
        test_reporter(deduped_df, validation_result, merge_result, dedup_result, overdue_result)

        print("=" * 60)
        print("✓ 所有核心模块测试通过！")
        print("=" * 60)
        print("\n输出文件位于: output/ 目录")
        print("测试数据位于: sample_data/ 目录")

    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
