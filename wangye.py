import streamlit as st
import pandas as pd
import os
import logging
import re
from difflib import SequenceMatcher
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import numbers
import base64
import sys
from io import BytesIO


# ============================
# 初始化设置
# ============================
# 设置页面配置
st.set_page_config(
    page_title="数据处理工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.info("启动数据处理工具。")


# ============================
# 学业桥数据处理相关工具函数
# ============================

# ======== 路径兼容函数 =========
def resource_path(relative_path):
    """兼容 PyCharm 开发环境 和 PyInstaller 打包后的路径"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# ======== 加载学校数据 =========
try:
    school_data_path = resource_path("school_data.xlsx")
    school_df = pd.read_excel(school_data_path)
    VALID_SCHOOL_NAMES = set(school_df['学校名称'].dropna().str.strip())
    logging.info(f"成功加载 {len(VALID_SCHOOL_NAMES)} 个有效学校名称")
except Exception as e:
    logging.error(f"读取 school_data.xlsx 出错：{e}")
    VALID_SCHOOL_NAMES = set()
    st.warning("学校数据加载失败，学校名称检查功能将不可用")

# ======== 加载招生专业数据 =========
try:
    major_data_path = resource_path("招生专业.xlsx")
    major_df = pd.read_excel(major_data_path)
    VALID_MAJOR_COMBOS = set(major_df['招生专业'].dropna().astype(str).str.strip())
    logging.info(f"成功加载 {len(VALID_MAJOR_COMBOS)} 个有效专业组合")
except Exception as e:
    logging.error(f"读取 招生专业.xlsx 出错：{e}")
    VALID_MAJOR_COMBOS = set()
    st.warning("专业数据加载失败，专业匹配功能将不可用")


def check_school_name(name):
    if pd.isna(name) or not str(name).strip():
        return '学校名称为空'
    return '匹配' if name.strip() in VALID_SCHOOL_NAMES else '不匹配'


def check_major_combo(major, level):
    if pd.isna(major) or pd.isna(level):
        return "数据缺失"
    combo = f"{str(major).strip()}{str(level).strip()}"
    return "匹配" if combo in VALID_MAJOR_COMBOS else "不匹配"


CUSTOM_WHITELIST = {
    "宏福校区", "沙河校区", "中外合作办学", "珠海校区", "江北校区", "津南校区", "开封校区",
    "联合办学", "校企合作", "合作办学", "威海校区", "深圳校区", "苏州校区", "平果校区",
    "江南校区", "合川校区", "长安校区", "崇安校区", "南校区", "东校区", "都市园艺", "甘肃兰州"
}

TYPO_DICT = {
    "教助": "救助",
    "指辉": "指挥",
    "料学": "科学",
    "话言": "语言",
    "5十3": "5+3",
    "5十3一体化": "5+3一体化",
    "“5十3”一体化": "“5+3”一体化",
    "5+31体化": "5+3一体化",
    "5+3体化": "5+3一体化",
    "色言": "色盲",
    "NIT": "NIIT",
    "色育": "色盲",
    "人围": "入围",
    "项月": "项目",
    "币范类": "师范类",
    "投课": "授课",
    "就薄": "就读",
    "电请": "申请",
    "中国面": "中国画",
    "火数民族": "少数民族",
    "色自": "色盲",
    "色盲色弱申报": "色盲色弱慎报",
    "数学与应用数笑": "数学与应用数学",
    "法学十": "法学+",
    "中溴": "中澳"
}

REGEX_PATTERNS = {
    'excess_punct': re.compile(r'[，、。！？；,;.!? ]+'),
    'outer_punct': re.compile(r'^[，、。！？；,;.!? ]+|[，、。！？；,;.!? ]+$'),
    'consecutive_right': re.compile(r'）{2,}')
}
NESTED_PAREN_PATTERN = re.compile(r'（（(.*?)））')
CONSECUTIVE_REPEAT_PATTERN = re.compile(r'（(.+?)）\s*（\1）')


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def normalize_brackets(text):
    """统一各种括号为中文括号并处理不完整括号"""
    if pd.isna(text) or not str(text).strip():
        return text
    text = str(text).strip()

    # 替换所有括号变体为中文括号
    text = re.sub(r'[\{\[\【]', '（', text)  # 左括号
    text = re.sub(r'[\}\]\】]', '）', text)  # 右括号
    text = re.sub(r'[<《]', '（', text)  # 左书名号替换为左括号
    text = re.sub(r'[>》]', '）', text)  # 右书名号替换为右括号

    # 补全普通括号
    if '（' in text and '）' not in text:
        text += '）'
    if '）' in text and '（' not in text:
        text = '（' + text

    # 处理连续右括号
    text = REGEX_PATTERNS['consecutive_right'].sub('）', text)

    return text


def clean_outer_punctuation(text):
    """清理最外层括号外的标点符号"""
    if pd.isna(text) or not str(text).strip():
        return text
    text = str(text).strip()
    text = REGEX_PATTERNS['outer_punct'].sub('', text)
    parts = re.split(r'(（.*?）)', text)
    cleaned_parts = []
    for part in parts:
        if part.startswith('（') and part.endswith('）'):
            cleaned_parts.append(part)
        else:
            cleaned_parts.append(REGEX_PATTERNS['outer_punct'].sub('', part))
    return ''.join(cleaned_parts)


def check_score_consistency(row):
    """检查分数一致性：最高分 >= 平均分 >= 最低分"""
    issues = []
    try:
        max_score = float(row['最高分']) if pd.notna(row['最高分']) else None
        avg_score = float(row['平均分']) if pd.notna(row['平均分']) else None
        min_score = float(row['最低分']) if pd.notna(row['最低分']) else None

        if max_score is not None and avg_score is not None and max_score < avg_score:
            issues.append(f"最高分({max_score}) < 平均分({avg_score})")

        if max_score is not None and min_score is not None and max_score < min_score:
            issues.append(f"最高分({max_score}) < 最低分({min_score})")

        if avg_score is not None and min_score is not None and avg_score < min_score:
            issues.append(f"平均分({avg_score}) < 最低分({min_score})")

    except (ValueError, TypeError) as e:
        issues.append(f"分数格式错误: {str(e)}")

    return '；'.join(issues) if issues else '无问题'


def analyze_and_fix(text):
    if pd.isna(text) or not str(text).strip():
        return text, []

    # 1. 括号规范化
    text = normalize_brackets(text)
    text = clean_outer_punctuation(text)
    original = text
    issues = []

    # 白名单跳过
    if text in CUSTOM_WHITELIST:
        return text, []

    # 括号匹配补全
    left, right = text.count('（'), text.count('）')
    if left != right:
        if left > right:
            text += '）' * (left - right)
            issues.append(f"补充缺失右括号 {left - right} 个")
        else:
            text = '（' * (right - left) + text
            issues.append(f"补充缺失左括号 {right - left} 个")

    # 处理嵌套括号
    text2 = NESTED_PAREN_PATTERN.sub(r'（\1）', text)
    if text2 != text:
        issues.append("存在嵌套括号")
    text = text2

    # 处理重复括号内容
    text, n = CONSECUTIVE_REPEAT_PATTERN.subn(r'（\1）', text)
    if n > 0:
        issues.append("存在重复括号内容")

    # 括号内容清洗
    def fix_paren(m):
        c = m.group(1)
        f = c.strip('，、,;；')
        if f != c:
            if c[0] in '，、,;；':
                issues.append(f"括号内容开头多标点：'{c}'")
            if c[-1] in '，、,;；':
                issues.append(f"括号内容结尾多标点：'{c}'")
        return f'（{f}）'

    text = re.sub(r'（(.*?)）', fix_paren, text)

    # 括号内去重
    seen = set()

    def dedup(m):
        c = m.group(1)
        if c in seen:
            issues.append(f"重复内容：{c}")
            return ''
        seen.add(c)
        return f'（{c}）'

    text = re.sub(r'（(.*?)）', dedup, text)

    # 简化多余标点
    text = REGEX_PATTERNS['excess_punct'].sub(lambda m: m.group(0)[0], text)

    # 相似重复检测
    contents = list(dict.fromkeys(re.findall(r'（(.*?)）', original)))
    for i in range(len(contents)):
        for j in range(i + 1, len(contents)):
            if similar(contents[i], contents[j]) >= 0.8:
                issues.append(f"相似重复：'{contents[i]}' 与 '{contents[j]}'")

    # 规则字典校正
    for typo, corr in TYPO_DICT.items():
        if typo in text:
            text = text.replace(typo, corr)
            issues.append(f"错别字：'{typo}'→'{corr}'")

    return text, issues


def process_chunk(chunk):
    """处理数据块"""
    # 学校名称检查
    if '学校名称' in chunk.columns:
        chunk['学校匹配结果'] = chunk['学校名称'].apply(check_school_name)

    # 专业匹配检查
    if '招生专业' in chunk.columns and '一级层次' in chunk.columns:
        chunk['招生专业匹配结果'] = chunk.apply(
            lambda r: check_major_combo(r['招生专业'], r['一级层次']), axis=1)

    # 备注处理 - 修改这部分
    if '专业备注' in chunk.columns:
        def process_remark(remark):
            if pd.isna(remark) or not str(remark).strip():
                return '无问题', ''
            fixed_text, issues = analyze_and_fix(remark)
            return '；'.join(issues) if issues else '无问题', fixed_text

        chunk[['备注检查结果', '修改后备注']] = chunk['专业备注'].apply(
            lambda x: pd.Series(process_remark(x)))

    # 分数检查
    score_columns = ['最高分', '平均分', '最低分']
    if all(col in chunk.columns for col in score_columns):
        chunk['分数检查结果'] = chunk.apply(check_score_consistency, axis=1)

    # 选科要求处理
    if '选科要求' in chunk.columns:
        def proc_req(req):
            if pd.isna(req) or not str(req).strip():
                return ["", ""]
            s = str(req).strip()
            if "不限" in s:
                return ["不限科目专业组", ""]
            if len(s) == 1:
                return ["单科、多科均需选考", s]
            if "且" in s:
                return ["单科、多科均需选考", s.replace("且", "")]
            if "或" in s:
                return ["多门选考", s.replace("或", "")]
            return ["", ""]

        chunk[['选科要求说明', '次选']] = chunk['选科要求'].apply(
            lambda x: pd.Series(proc_req(x)))

    # 招生科类处理
    if '招生科类' in chunk.columns:
        chunk['招生科类'] = chunk['招生科类'].replace({'物理': '物理类', '历史': '历史类'})
        chunk['首选科目'] = chunk['招生科类'].apply(
            lambda x: str(x)[0] if x in ['物理类', '历史类'] else "")

    return chunk




# ============================
# 院校分提取相关函数
# ============================
expected_columns = [
    '学校名称', '省份', '招生专业', '专业方向（选填）', '专业备注（选填）', '一级层次', '招生科类', '招生批次',
    '招生类型（选填）', '最高分', '最低分', '平均分', '最低分位次（选填）', '招生人数（选填）', '数据来源',
    '专业组代码', '首选科目', '选科要求', '次选科目', '专业代码', '招生代码', '录取人数（选填）'
]
columns_to_convert = [
    '专业组代码', '专业代码', '招生代码', '最高分', '最低分', '平均分', '最低分位次（选填）',
    '招生人数（选填）'
]

def process_score_file(file_path):
    try:
        df = pd.read_excel(file_path, header=2, dtype={
            '专业组代码': str,
            '专业代码': str,
            '招生代码': str,
            '最高分': str,
            '最低分': str,
            '平均分': str,
            '最低分位次（选填）': str,
            '招生人数（选填）': str,
            '录取人数（选填）': str
        }, keep_default_na=False, engine='openpyxl')
    except Exception as e:
        raise Exception(f"读取文件错误：{e}")

    missing_columns = [col for col in expected_columns if col not in df.columns]
    if missing_columns:
        raise Exception(f"文件缺少以下列：{missing_columns}")

    df['最低分'] = pd.to_numeric(df['最低分'], errors='coerce')
    df['最高分'] = pd.to_numeric(df['最高分'], errors='coerce')
    df['招生人数（选填）'] = pd.to_numeric(df['招生人数（选填）'], errors='coerce')
    df['录取人数（选填）'] = pd.to_numeric(df['录取人数（选填）'], errors='coerce')
    df = df.dropna(subset=['最低分'])

    if df.empty:
        raise Exception("数据处理后为空。")

    df['招生类型（选填）'] = df['招生类型（选填）'].replace([None], '')

    # 首选科目转换逻辑
    if '首选科目' in df.columns:
        df['首选科目'] = df['首选科目'].str.strip()  # 去除前后空格
        df['首选科目'] = df['首选科目'].replace({
            '历': '历史',
            '物': '物理',
            '历史': '历史',  # 确保已经是"历史"的不变
            '物理': '物理'  # 确保已经是"物理"的不变
        })


    try:
        # 分组字段（含专业组代码）
        group_with_code = ['学校名称', '省份', '一级层次', '招生科类', '招生批次', '专业组代码', '招生类型（选填）']

        # 取每组最低分所在行索引
        min_indices = df.groupby(group_with_code)['最低分'].idxmin()

        # 取每组最高分数（单纯数值，不取对应行）
        max_scores = df.groupby(group_with_code)['最高分'].max()

        # 取最低分行数据
        result = df.loc[min_indices].copy()

        # 招生人数为分组总和
        enroll_groups = df.groupby(group_with_code)['招生人数（选填）'].sum()

        # 录取人数为分组总和
        code_groups = df.groupby(group_with_code)['录取人数（选填）'].sum()

        def get_group_total(row, column_name):
            key = tuple(row[col] for col in group_with_code)
            if column_name == '招生人数（选填）':
                return code_groups.get(key, '')
            elif column_name == '录取人数（选填）':
                return enroll_groups.get(key, '')
            return ''

        result['招生人数（选填）'] = result.apply(lambda row: get_group_total(row, '招生人数（选填）'), axis=1)
        result['录取人数（选填）'] = result.apply(lambda row: get_group_total(row, '录取人数（选填）'), axis=1)


    except Exception as e:
        raise Exception(f"分组字段错误：{e}")

    if result.empty:
        raise Exception("筛选结果为空。")

    # 保留期望列，但排除招生专业和专业方向、专业备注、选科要求、次选科目
    selected_columns = [col for col in expected_columns if col in result.columns and col not in ['招生专业', '专业方向（选填）', '专业备注（选填）', '选科要求', '次选科目']]
    result = result[selected_columns]

    output_path = file_path.replace('.xlsx', '_院校分.xlsx')

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for col in ['专业组代码', '专业代码', '招生代码']:
                if col in result.columns:
                    col_idx = result.columns.get_loc(col) + 1
                    for row in range(2, len(result) + 2):
                        worksheet.cell(row=row, column=col_idx).number_format = numbers.FORMAT_TEXT

            for col in columns_to_convert:
                if col in result.columns and col not in ['专业组代码', '专业代码', '招生代码']:
                    col_idx = result.columns.get_loc(col) + 1
                    for cell in list(worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, values_only=False))[0]:
                        cell.number_format = numbers.FORMAT_TEXT

        return output_path
    except Exception as e:
        raise Exception(f"文件保存失败：{e}")

# ============================
# 保持文本格式
# ============================
def process_remarks_file(file_path, progress_callback=None):
    try:
        # 读取文件时，确保这些字段始终以字符串格式读取
        df = pd.read_excel(file_path, header=2, dtype={
            '专业组代码': str,
            '专业代码': str,
            '招生代码': str,
        }, engine='openpyxl')
    except Exception as e:
        raise Exception(f"读取文件错误：{e}")
    for col in ['专业组代码', '专业代码', '招生代码']:
        if col in df.columns:
            df[col] = df[col].astype(str)
    target_col = None
    for col in df.columns:
        if "专业备注" in str(col):
            target_col = col
            break
    if not target_col:
        raise Exception("未找到'专业备注'相关列")
    if target_col != '专业备注':
        df = df.rename(columns={target_col: '专业备注'})
    chunks = []
    for i in range(0, len(df), 1000):
        chunks.append(df.iloc[i:i + 1000].copy())
    results = {}
    total_chunks = len(chunks)
    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
        future_to_index = {executor.submit(process_chunk, chunk): idx for idx, chunk in enumerate(chunks)}
        for count, future in enumerate(as_completed(future_to_index)):
            idx = future_to_index[future]
            results[idx] = future.result()
            if progress_callback:
                progress_callback(count + 1, total_chunks)
    ordered_results = [results[i] for i in sorted(results.keys())]
    final_result = pd.concat(ordered_results)
    output_path = file_path.replace('.xlsx', '_检查结果.xlsx')
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            final_result.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            # 保持指定列从第三行开始文本格式
            for col in ['专业组代码', '专业代码', '招生代码']:
                if col in final_result.columns:
                    col_idx = final_result.columns.get_loc(col) + 1  # 转换为Excel列号（A=1）
                    # 从第三行开始设置格式（Excel行号为3，对应Python的索引为2）
                    for row in range(3, len(final_result) + 2):  # 工作表行号从3开始（索引2）
                        cell = worksheet.cell(row=row, column=col_idx)
                        cell.value = final_result.iloc[row - 3][col]  # 数据从第三行开始填充
                        cell.number_format = numbers.FORMAT_TEXT
    except Exception as e:
        raise Exception(f"保存文件错误：{e}")
    return output_path

# ============================
# 一分一段数据处理
# ============================

def process_segmentation_file(file_path):
    output_path = os.path.splitext(file_path)[0] + "_校验结果.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    ws['E7'] = '累计人数校验结果'
    ws['F7'] = '分数校验结果'
    ws['F2'] = '年份校验'

    # 校验 B2 是否为 2025
    if ws['B2'].value != 2025:
        ws['G2'] = f"× 应为2025，当前为：{ws['B2'].value}"
    else:
        ws['G2'] = "√"

    region = ws['B3'].value
    suffix = "-750"
    if region == "上海":
        suffix = "-660"
    elif region == "海南":
        suffix = "-900"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # ---------- 第8行特殊处理 ----------
    row = 8
    curr_score = ws[f"A{row}"].value
    curr_num = ws[f"B{row}"].value
    curr_total = ws[f"C{row}"].value

    try:
        score_int = int(float(str(curr_score).split('-')[0]))
    except:
        score_int = None

    inserted = False
    if curr_total is not None:
        if curr_num is None or curr_num == "":
            # 没有人数 → 自动计算
            if row == 8:
                ws[f"B{row}"] = curr_total
            else:
                prev_total = ws[f"C{row - 1}"].value
                if prev_total is not None:
                    ws[f"B{row}"] = curr_total - prev_total
        else:
            # 有人数和累计人数不一致时插入补断点行
            if curr_num != curr_total:
                try:
                    insert_score = score_int + 1
                    insert_num = curr_total - curr_num
                    ws.insert_rows(row)
                    ws[f"A{row}"] = f"{insert_score}{suffix}"  # ✅ 仅加后缀在新增行
                    ws[f"B{row}"] = insert_num
                    ws[f"C{row}"] = insert_num
                    for col in ['A', 'B', 'C', 'E', 'F']:
                        ws[f"{col}{row}"].fill = yellow_fill
                    ws[f"E{row}"] = "补断点"
                    ws[f"F{row}"] = "补断点"
                    inserted = True
                except:
                    pass

    # 仅当没有插入行时，第8行加后缀
    if not inserted and score_int is not None:
        ws[f"A{row}"] = f"{score_int}{suffix}"

    # ---------- 补断点逻辑 ----------
    while row < ws.max_row:
        curr = ws[f"A{row}"].value
        next = ws[f"A{row + 1}"].value
        try:
            curr_score_int = int(str(curr).split('-')[0])
            next_score_int = int(str(next).split('-')[0])
        except:
            row += 1
            continue

        if curr_score_int - next_score_int > 1:
            missing_score = curr_score_int - 1
            ws.insert_rows(row + 1)
            ws[f"A{row + 1}"] = missing_score
            ws[f"B{row + 1}"] = 0
            ws[f"C{row + 1}"] = ws[f"C{row}"].value
            for col in ['A', 'B', 'C', 'E', 'F']:
                ws[f"{col}{row + 1}"].fill = yellow_fill
            ws[f"E{row + 1}"] = "补断点"
            ws[f"F{row + 1}"] = "补断点"
        else:
            row += 1

    # ---------- 校验与自动补人数 ----------
    for row in range(8, ws.max_row + 1):
        curr_score = ws[f"A{row}"].value
        curr_num = ws[f"B{row}"].value
        curr_total = ws[f"C{row}"].value
        prev_total = ws[f"C{row - 1}"].value if row > 8 else None
        prev_score = ws[f"A{row - 1}"].value if row > 8 else None

        # 自动补人数
        if (curr_num is None or curr_num == "") and curr_total is not None:
            if row == 8:
                ws[f"B{row}"] = curr_total
                curr_num = curr_total
            elif prev_total is not None:
                try:
                    calc = curr_total - prev_total
                    ws[f"B{row}"] = calc
                    curr_num = calc
                except:
                    pass

        # 校验累计人数
        if prev_total is not None and curr_num is not None and curr_total is not None:
            expected_total = prev_total + curr_num
            if expected_total == curr_total:
                if ws[f"E{row}"].value != "补断点":
                    ws[f"E{row}"] = "√"
            else:
                if ws[f"E{row}"].value != "补断点":
                    ws[f"E{row}"] = f"× 应为{expected_total}"

        # 校验分数差
        try:
            curr_score_num = float(str(curr_score).split('-')[0])
            prev_score_num = float(str(prev_score).split('-')[0])
        except:
            curr_score_num = prev_score_num = None

        if curr_score_num is not None and prev_score_num is not None:
            diff = prev_score_num - curr_score_num
            if diff == 1:
                if ws[f"F{row}"].value != "补断点":
                    ws[f"F{row}"] = "√"
            else:
                if ws[f"F{row}"].value != "补断点":
                    ws[f"F{row}"] = f"× 差值{diff}"
        else:
            if ws[f"F{row}"].value != "补断点":
                ws[f"F{row}"] = "× 分数非数字，无法校验"

    wb.save(output_path)
    return output_path


# ============================
# 专业组代码匹配
# ============================
# 匹配参数
SIMILARITY_THRESHOLD = 0.5

tableA_fields = [
    "学校名称", "省份", "招生专业", "专业备注（选填）",
    "一级层次", "招生科类", "招生批次", "招生类型（选填）"
]

rename_mapping_B = {
    "学校": "学校名称",
    "省份": "省份",
    "层次": "一级层次",
    "科类": "招生科类",
    "批次": "招生批次",
    "招生类型": "招生类型（选填）",
    "专业": "招生专业",
    "备注": "专业备注（选填）"
}


def clean_remark(text):
    """更彻底的备注清洗函数"""
    if pd.isna(text):
        return ""

    # 统一转换为小写并去除首尾空格
    cleaned = str(text).strip().lower()

    # 移除括号及其内容（保留括号内的文本）
    cleaned = re.sub(r'[\(\)（）]', '', cleaned)

    # 移除常见分隔符（保留分隔符之间的文本）
    cleaned = re.sub(r'[;；、,:：]', ' ', cleaned)

    # 合并多余空格
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    return cleaned


def fuzzy_match(row, b_dict):
    key = row["组合键"]
    candidates = b_dict.get(key, [])
    if not candidates:
        return None

    remark_a = row["专业备注（选填）_清洗"]
    best_match = None
    max_similarity = 0

    # 特殊处理空备注情况
    if not remark_a:
        # 如果A备注为空，优先选择B备注也为空的记录
        empty_remarks = [c for c in candidates if not c["专业备注（选填）_清洗"]]
        if empty_remarks:
            return empty_remarks[0]["专业组代码"]
        # 如果没有完全空白的备注，则选择第一个（或根据其他逻辑）
        return candidates[0]["专业组代码"]

    for candidate in candidates:
        remark_b = candidate["专业备注（选填）_清洗"]

        # 1. 优先判断核心关键词匹配
        # 提取A备注中的关键词（按空格分割）
        keywords_a = set(remark_a.split())

        # 检查A的所有关键词是否都出现在B备注中
        if keywords_a and all(kw in remark_b for kw in keywords_a):
            return candidate["专业组代码"]

        # 2. 部分包含关系（子串匹配）
        if remark_a in remark_b:
            return candidate["专业组代码"]

        # 3. 相似度匹配（使用集合相似度，更高效）
        set_a = set(remark_a.split())
        set_b = set(remark_b.split())

        if not set_a or not set_b:
            similarity = 0
        else:
            # 使用Jaccard相似度
            intersection = len(set_a & set_b)
            union = len(set_a | set_b)
            similarity = intersection / union if union > 0 else 0

        if similarity > max_similarity and similarity >= SIMILARITY_THRESHOLD:
            max_similarity = similarity
            best_match = candidate

    return best_match["专业组代码"] if best_match else None


def process_data(dfA, dfB):
    # 确保导入所需库

    dfB.rename(columns=rename_mapping_B, inplace=True)

    # 清洗备注字段（使用优化后的清洗函数）
    dfA["专业备注（选填）_清洗"] = dfA["专业备注（选填）"].apply(clean_remark)
    dfB["专业备注（选填）_清洗"] = dfB["专业备注（选填）"].apply(clean_remark)

    # 构建组合键（不含备注）
    key_fields = [f for f in tableA_fields if f != "专业备注（选填）"]
    dfA["组合键"] = dfA[key_fields].fillna("").astype(str).apply(
        lambda x: "|".join([str(i).strip() for i in x]), axis=1)
    dfB["组合键"] = dfB[key_fields].fillna("").astype(str).apply(
        lambda x: "|".join([str(i).strip() for i in x]), axis=1)

    # 构建B表字典：组合键 → 记录列表
    b_dict = dfB.groupby("组合键").apply(lambda x: x.to_dict("records")).to_dict()

    def get_code(row):
        key = row["组合键"]
        candidates = b_dict.get(key, [])

        # 情况1：无候选记录
        if not candidates:
            return None

        # 情况2：唯一候选记录
        if len(candidates) == 1:
            return candidates[0]["专业组代码"]

        # 情况3：多个候选记录，使用模糊匹配
        return fuzzy_match(row, b_dict)

    dfA["专业组代码"] = dfA.apply(get_code, axis=1)

    return dfA



# ============================
# Streamlit页面布局
# ============================
# 页面标题
st.title("📊 数据处理工具")
st.markdown("---")

# 功能说明
with st.expander("📌 功能说明", expanded=True):
    st.markdown("""
    1. 上传的文件使用库中专业分、院校分、招生计划、一分一段的模板，直接上传即可，无需删减
    2. 备注检查中，检查出来括号有问题的内容还需要自己再过一遍；整个文件的备注需要大概看看有没有错别字
    3. 校验一分一段时，内容不能为文本格式
    """)

# 更新日志对话框
with st.expander("📢 版本更新（2025.6.12更新）", expanded=False):
    st.markdown("""
    ### 2025.6.12更新
    院校分提取逻辑更新  
      - 提取最高分改为取同一个“学校-省份-层次-科类-批次-类型（-专业组代码）”下的最高分

    ### 历史更新

    #### 2025.4.14更新
    • 招生代码和专业代码保持文本格式  
    • 增加功能说明  
    • 优化工具界面  

    #### 2025.4.16更新
    • 优化了院校分提取处理逻辑  

    #### 2025.5.22更新
    • 更新了院校分提取中录取人数的处理逻辑（建议进行抽查）  
    • 学业桥数据处理中增加了最高分、平均分、最低分的校验，会在最后加一列校验结果  

    #### 2025.5.23更新
    • 学业桥数据处理中增加了学校名称和招生专业的匹配  

    #### 2025.5.27更新
    • 学业桥数据处理中，增加了"招生科类"、"首选科目"、"选科要求"，"次选科目"的处理  
      - 学业桥提供的"3+1+2"省份的招生科类为"物理"、"历史"，可以直接转换为标准的"物理类"、"历史类"  
      - "3+1+2"省份的首选科目可以直接根据招生科类提取  
      - 新增了选科要求、次选科目的处理，可直接转换为标准格式，无需手动处理（处理后的数据在文档最后几列）  

    #### 2025.5.30更新
    新增"一分一段数据处理"  
      - 可直接校验分数、累计人数  
      - 自动补断点  
      - 自动增加"最高分——满分"的区间（上海满分660，海南满分900）  
      
    ### 2025.6.6更新
    "一分一段数据处理"优化  
      - 自动补充"最高分——满分"的区间（上海满分660，海南满分900）  
      - 只有累计人数没有人数时，可计算人数，无需手动操作  
      - 补断点的分数标注颜色，并在分数和人数校验中标注"补断点"
      
    """)

# 创建选项卡
tab1, tab2, tab3, tab4 = st.tabs(["院校分提取", "学业桥数据处理", "一分一段校验", "专业组代码匹配（可以用，需要检查！）"])

# ====================== 院校分提取 ======================
with tab1:
    st.header("院校分提取")

    # 文件上传
    uploaded_file = st.file_uploader("选择Excel文件", type=["xlsx"], key="score_file")

    if uploaded_file is not None:
        st.success(f"已选择文件: {uploaded_file.name}")

        # 显示处理进度
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("准备处理...")

        # 处理按钮
        if st.button("开始数据处理", key="process_score"):
            try:
                # 保存上传的文件到临时位置
                temp_file = "temp_score.xlsx"
                with open(temp_file, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # 处理文件
                for percent_complete in range(0, 101, 10):
                    progress_bar.progress(percent_complete)
                    status_text.text(f"处理中... {percent_complete}%")

                    # 模拟处理过程，实际使用时替换为您的process_score_file函数
                    if percent_complete == 100:
                        output_path = process_score_file(temp_file)

                # 处理完成
                status_text.text("处理完成！")
                st.balloons()

                # 提供下载链接
                with open(output_path, "rb") as f:
                    bytes_data = f.read()
                b64 = base64.b64encode(bytes_data).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="院校分.xlsx">点击下载处理结果</a>'
                st.markdown(href, unsafe_allow_html=True)

                # 清理临时文件
                os.remove(temp_file)
                os.remove(output_path)

            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")

# ====================== 学业桥数据处理 ======================
with tab2:
    st.header("学业桥数据处理")

    # 文件上传
    uploaded_file = st.file_uploader("选择Excel文件", type=["xlsx"], key="remarks_file")

    if uploaded_file is not None:
        st.success(f"已选择文件: {uploaded_file.name}")

        # 显示处理进度
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("准备处理...")

        # 处理按钮
        if st.button("开始数据处理", key="process_remarks"):
            try:
                # 保存上传的文件到临时位置
                temp_file = "temp_remarks.xlsx"
                with open(temp_file, "wb") as f:
                    f.write(uploaded_file.getbuffer())


                # 进度回调函数
                def update_progress(current, total):
                    percent = int((current / total) * 100)
                    progress_bar.progress(percent)
                    status_text.text(f"处理中... {percent}%")


                # 处理文件
                output_path = process_remarks_file(temp_file, progress_callback=update_progress)

                # 处理完成
                progress_bar.progress(100)
                status_text.text("处理完成！")
                st.balloons()

                # 提供下载链接
                with open(output_path, "rb") as f:
                    bytes_data = f.read()
                b64 = base64.b64encode(bytes_data).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="学业桥数据.xlsx">点击下载处理结果</a>'
                st.markdown(href, unsafe_allow_html=True)

                # 清理临时文件
                os.remove(temp_file)
                os.remove(output_path)

            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")

# ====================== 一分一段校验 ======================
with tab3:
    st.header("一分一段校验")

    # 文件上传
    uploaded_file = st.file_uploader("选择Excel文件", type=["xlsx"], key="segmentation_file")

    if uploaded_file is not None:
        st.success(f"已选择文件: {uploaded_file.name}")

        # 显示处理进度
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("准备处理...")

        # 处理按钮
        if st.button("开始数据处理", key="process_segmentation"):
            try:
                # 保存上传的文件到临时位置
                temp_file = "temp_segmentation.xlsx"
                with open(temp_file, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # 处理文件
                for percent_complete in range(0, 101, 10):
                    progress_bar.progress(percent_complete)
                    status_text.text(f"处理中... {percent_complete}%")

                    # 模拟处理过程，实际使用时替换为您的process_segmentation_file函数
                    if percent_complete == 100:
                        output_path = process_segmentation_file(temp_file)

                # 处理完成
                status_text.text("处理完成！")
                st.balloons()

                # 提供下载链接
                with open(output_path, "rb") as f:
                    bytes_data = f.read()
                b64 = base64.b64encode(bytes_data).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="一分一段.xlsx">点击下载处理结果</a>'
                st.markdown(href, unsafe_allow_html=True)

                # 清理临时文件
                os.remove(temp_file)
                os.remove(output_path)

            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")


# ====================== 专业组代码匹配 ======================
with tab4:
    st.header("专业组代码匹配（需要检查！）")

    uploaded_fileA = st.file_uploader("上传专业分导入模板", type=["xls", "xlsx"], key="fileA")
    uploaded_fileB = st.file_uploader("上传招生计划数据导出文件", type=["xls", "xlsx"], key="fileB")

    if uploaded_fileA and uploaded_fileB:
        st.success(f"已选择文件：{uploaded_fileA.name} 和 {uploaded_fileB.name}")

        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("等待开始处理...")

        if st.button("开始数据处理", key="start_match"):
            try:
                # 保存临时文件
                temp_fileA = "tempA.xlsx"
                temp_fileB = "tempB.xlsx"
                with open(temp_fileA, "wb") as f:
                    f.write(uploaded_fileA.getbuffer())
                with open(temp_fileB, "wb") as f:
                    f.write(uploaded_fileB.getbuffer())

                status_text.text("读取文件...")
                progress_bar.progress(10)

                dfA = pd.read_excel(temp_fileA, header=2)
                dfB = pd.read_excel(temp_fileB)

                status_text.text("开始处理数据...")
                for percent_complete in range(20, 101, 20):
                    progress_bar.progress(percent_complete)
                    # 模拟处理时间，如果不需要可以去掉
                    # time.sleep(0.2)

                result_df = process_data(dfA, dfB)

                status_text.text("处理完成！准备导出...")
                progress_bar.progress(100)

                # 导出结果到内存
                output = BytesIO()
                result_df.to_excel(output, index=False)
                output.seek(0)

                b64 = base64.b64encode(output.read()).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="专业组代码匹配结果.xlsx">点击下载匹配结果</a>'
                st.markdown(href, unsafe_allow_html=True)

                # 清理临时文件
                os.remove(temp_fileA)
                os.remove(temp_fileB)

                status_text.text("已完成，结果可下载。")
                st.balloons()

            except Exception as e:
                st.error(f"处理错误：{e}")
    else:
        st.info("请先上传两个Excel文件")


# 页脚
st.markdown("---")
st.markdown("© 数据处理", unsafe_allow_html=True)