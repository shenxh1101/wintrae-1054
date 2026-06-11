#!/usr/bin/env python3
"""
生成测试数据
"""

import pandas as pd
import os
import random
from datetime import datetime, timedelta


def generate_id_card():
    prefix = "33010219"
    birth_year = random.randint(50, 95)
    birth_month = random.randint(1, 12)
    birth_day = random.randint(1, 28)
    suffix = f"{random.randint(100, 999)}{random.randint(0, 9)}"
    return f"{prefix}{birth_year:02d}{birth_month:02d}{birth_day:02d}{suffix}"


def generate_phone():
    prefixes = ['138', '139', '150', '151', '152', '188', '189', '135', '136', '137']
    return random.choice(prefixes) + ''.join(str(random.randint(0, 9)) for _ in range(8))


def generate_sample_data(category, count=50, output_dir='sample_data'):
    os.makedirs(output_dir, exist_ok=True)

    villages = ['东村村', '西村村', '南村村', '北村村', '中心村', '和平村', '建设村', '新华村']
    surnames = ['张', '王', '李', '赵', '陈', '刘', '杨', '黄', '周', '吴']
    names = ['伟', '芳', '娜', '敏', '静', '丽', '强', '磊', '军', '洋', '勇', '艳', '杰', '娟', '涛']

    data = []
    id_cards = [generate_id_card() for _ in range(count)]

    for i in range(count):
        name = random.choice(surnames) + random.choice(names)
        id_card = id_cards[i]
        phone = generate_phone()
        village = random.choice(villages)

        base_date = datetime.now() - timedelta(days=random.randint(0, 200))
        followup_date = base_date.strftime('%Y-%m-%d')

        results = ['满意', '基本满意', '满意', '满意', '基本满意', '正常', '完成']
        result = random.choice(results)

        row = {
            '姓名': name,
            '身份证号': id_card,
            '联系电话': phone,
            '村居': village,
            '随访日期': followup_date,
            '随访结果': result,
        }

        if category == '孕产妇':
            row['孕周'] = random.randint(12, 38)
        elif category in ['儿童', '老年人']:
            if category == '儿童':
                row['年龄'] = random.randint(1, 6)
            else:
                row['年龄'] = random.randint(65, 85)

        if random.random() < 0.08:
            row['身份证号'] = ''
        if random.random() < 0.05:
            row['联系电话'] = ''
        if random.random() < 0.06:
            row['随访日期'] = ''
        if random.random() < 0.04:
            row['随访结果'] = ''

        data.append(row)

    df = pd.DataFrame(data)

    output_path = os.path.join(output_dir, f'{category}随访记录.xlsx')
    df.to_excel(output_path, index=False)
    print(f"已生成: {output_path} ({len(df)} 条记录)")

    return output_path


def generate_duplicate_data(category, output_dir='sample_data'):
    src_path = os.path.join(output_dir, f'{category}随访记录.xlsx')
    if not os.path.exists(src_path):
        return

    df = pd.read_excel(src_path, dtype={'身份证号': str, '联系电话': str})
    valid_df = df[df['随访日期'].notna() & (df['随访日期'] != '')]
    if valid_df.empty:
        return

    dup_count = min(10, len(valid_df) // 5)
    dup_df = valid_df.sample(dup_count).copy()

    for idx, row in dup_df.iterrows():
        date_val = str(row['随访日期'])
        try:
            new_date = (datetime.strptime(date_val, '%Y-%m-%d') -
                        timedelta(days=random.randint(30, 90))).strftime('%Y-%m-%d')
            dup_df.at[idx, '随访日期'] = new_date
        except (ValueError, TypeError):
            pass

    result_df = pd.concat([df, dup_df], ignore_index=True)
    result_df = result_df.sample(frac=1).reset_index(drop=True)

    output_path = os.path.join(output_dir, f'{category}随访记录_第二批.xlsx')
    result_df.to_excel(output_path, index=False)
    print(f"已生成: {output_path} ({len(result_df)} 条记录, 含重复)")


def main():
    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')

    categories = ['慢病', '孕产妇', '儿童', '老年人']

    print("正在生成测试数据...")
    print("=" * 50)

    for cat in categories:
        generate_sample_data(cat, count=60, output_dir=sample_dir)
        generate_duplicate_data(cat, output_dir=sample_dir)
        print()

    print("=" * 50)
    print(f"测试数据已生成到: {sample_dir}")
    print("共 4 个分类，每个分类 2 个文件")


if __name__ == '__main__':
    main()
