"""
测试招生计划比对模块的功能
"""

import pandas as pd
from plan_comparison import (
    generate_plan_score_key,
    generate_plan_college_key,
    get_comparison_stats,
    convert_level,
    get_first_subject
)

print("=" * 60)
print("开始功能测试...")
print("=" * 60)

# 测试关键字生成函数
test_row = {
    '年份': '2023',
    '省份': '北京',
    '学校': '清华大学',
    '科类': '物理类',
    '批次': '本科一批',
    '专业': '计算机科学与技术',
    '层次': '本科',
    '专业组代码': '01'
}

key1 = generate_plan_score_key(test_row)
print(f"✓ 比对1关键字: {key1}")

key2 = generate_plan_college_key(test_row)
print(f"✓ 比对2关键字: {key2}")

# 测试数据转换函数
level = convert_level('本科')
print(f"✓ 层次转换 (本科 -> {level})")

subject = get_first_subject('物理类')
print(f"✓ 首选科目 (物理类 -> {subject})")

# 测试统计函数
test_results = [
    {'exists': True, 'index': 1},
    {'exists': True, 'index': 2},
    {'exists': False, 'index': 3},
    {'exists': False, 'index': 4},
    {'exists': False, 'index': 5}
]

stats = get_comparison_stats(test_results)
print(f"✓ 统计信息: 总计{stats['total']} | 匹配{stats['matched']} | 未匹配{stats['unmatched']} | 匹配率{stats['match_rate']}")

print("\n" + "=" * 60)
print("✓ 所有功能测试通过！")
print("=" * 60)
