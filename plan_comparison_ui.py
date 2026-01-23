<<<<<<< HEAD
"""
招生计划数据比对与转换工具 - Streamlit UI模块
提供Streamlit界面组件和用户交互逻辑
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import logging
import base64
from plan_comparison import (
    load_excel_from_bytes,
    compare_plan_vs_score,
    compare_plan_vs_college,
    get_comparison_stats,
    get_unique_provinces,
    get_unique_batches,
    convert_data_to_score_format,
    export_results_to_excel,
    export_converted_data_to_excel
)

logger = logging.getLogger(__name__)


# ==================== 初始化Session State ====================

def init_session_state():
    """初始化Streamlit Session State"""
    if 'plan_df' not in st.session_state:
        st.session_state.plan_df = None
    if 'score_df' not in st.session_state:
        st.session_state.score_df = None
    if 'college_df' not in st.session_state:
        st.session_state.college_df = None
    
    if 'plan_score_results' not in st.session_state:
        st.session_state.plan_score_results = None
    if 'plan_college_results' not in st.session_state:
        st.session_state.plan_college_results = None
    
    if 'converted_data' not in st.session_state:
        st.session_state.converted_data = None
    if 'conversion_source' not in st.session_state:
        st.session_state.conversion_source = None


# ==================== 文件加载 ====================

def load_files_section():
    """文件上传部分"""
    st.subheader("📁 文件上传")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.write("**招生计划文件**")
        plan_file = st.file_uploader("选择招生计划Excel文件", type=["xlsx", "xls"], key="plan_file")
        if plan_file:
            try:
                st.session_state.plan_df = load_excel_from_bytes(plan_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.plan_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")
    
    with col2:
        st.write("**专业分文件**")
        score_file = st.file_uploader("选择专业分Excel文件", type=["xlsx", "xls"], key="score_file")
        if score_file:
            try:
                st.session_state.score_df = load_excel_from_bytes(score_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.score_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")
    
    with col3:
        st.write("**院校分文件**")
        college_file = st.file_uploader("选择院校分Excel文件", type=["xlsx", "xls"], key="college_file")
        if college_file:
            try:
                st.session_state.college_df = load_excel_from_bytes(college_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.college_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")


# ==================== 比对操作 ====================

def comparison_operations():
    """比对操作部分"""
    st.subheader("🔍 数据比对")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("比对1：招生计划 vs 专业分", key="compare_plan_score"):
            if st.session_state.plan_df is None:
                st.error("请先上传招生计划文件")
            elif st.session_state.score_df is None:
                st.error("请先上传专业分文件")
            else:
                with st.spinner("正在进行比对1..."):
                    try:
                        st.session_state.plan_score_results = compare_plan_vs_score(
                            st.session_state.plan_df,
                            st.session_state.score_df
                        )
                        st.success("✓ 比对1完成")
                        st.session_state.conversion_source = 'planScore'
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col2:
        if st.button("比对2：招生计划 vs 院校分", key="compare_plan_college"):
            if st.session_state.plan_df is None:
                st.error("请先上传招生计划文件")
            elif st.session_state.college_df is None:
                st.error("请先上传院校分文件")
            else:
                with st.spinner("正在进行比对2..."):
                    try:
                        st.session_state.plan_college_results = compare_plan_vs_college(
                            st.session_state.plan_df,
                            st.session_state.college_df
                        )
                        st.success("✓ 比对2完成")
                        st.session_state.conversion_source = 'planCollege'
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col3:
        if st.button("全部比对", key="compare_all"):
            has_plan = st.session_state.plan_df is not None
            has_score = st.session_state.score_df is not None
            has_college = st.session_state.college_df is not None
            
            if not has_plan:
                st.error("请先上传招生计划文件")
            elif not (has_score or has_college):
                st.error("请至少上传专业分或院校分文件")
            else:
                with st.spinner("正在执行全部比对..."):
                    try:
                        if has_score:
                            st.session_state.plan_score_results = compare_plan_vs_score(
                                st.session_state.plan_df,
                                st.session_state.score_df
                            )
                        if has_college:
                            st.session_state.plan_college_results = compare_plan_vs_college(
                                st.session_state.plan_df,
                                st.session_state.college_df
                            )
                        st.success("✓ 全部比对完成")
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col4:
        if st.button("重置所有数据", key="reset_all"):
            st.session_state.plan_df = None
            st.session_state.score_df = None
            st.session_state.college_df = None
            st.session_state.plan_score_results = None
            st.session_state.plan_college_results = None
            st.session_state.converted_data = None
            st.session_state.conversion_source = None
            st.success("✓ 已重置所有数据")


# ==================== 结果显示 ====================

def display_comparison_results():
    """显示比对结果"""
    
    # 比对1结果
    if st.session_state.plan_score_results:
        st.subheader("📊 比对1：招生计划 vs 专业分")
        
        results = st.session_state.plan_score_results
        stats = get_comparison_stats(results)
        
        # 统计信息
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总记录数", stats['total'])
        col2.metric("匹配记录数", stats['matched'], delta="✓")
        col3.metric("未匹配记录数", stats['unmatched'], delta="✗")
        col4.metric("匹配率", stats['match_rate'])
        
        # 筛选选项
        col1, col2, col3 = st.columns(3)
        
        with col1:
            provinces = ['全部'] + get_unique_provinces(results)
            selected_province = st.selectbox(
                "按省份筛选",
                provinces,
                key="plan_score_province"
            )
        
        with col2:
            batches = ['全部'] + get_unique_batches(results)
            selected_batch = st.selectbox(
                "按批次筛选",
                batches,
                key="plan_score_batch"
            )
        
        with col3:
            match_status = st.selectbox(
                "匹配状态",
                ['全部', '匹配', '未匹配'],
                key="plan_score_status"
            )
        
        # 过滤数据
        filtered_results = results
        
        if selected_province != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['省份']).strip() == selected_province]
        
        if selected_batch != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['批次']).strip() == selected_batch]
        
        if match_status == '匹配':
            filtered_results = [r for r in filtered_results if r['exists']]
        elif match_status == '未匹配':
            filtered_results = [r for r in filtered_results if not r['exists']]
        
        # 显示表格
        st.write(f"**显示 {len(filtered_results)} 条记录**")
        
        display_data = []
        for result in filtered_results[:500]:  # 限制显示500条
            row = {
                '序号': result['index'],
                '状态': '✓ 匹配' if result['exists'] else '✗ 未匹配',
                **result['key_fields']
            }
            display_data.append(row)
        
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
        
        # 导出按钮
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 导出比对结果", key="export_plan_score_results"):
                try:
                    file_bytes = export_results_to_excel(results, "plan_score_results.xlsx")
                    st.download_button(
                        label="下载 比对1 结果",
                        data=file_bytes,
                        file_name="招生计划vs专业分_比对结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"导出失败: {str(e)}")
        
        with col2:
            if st.button("🔄 转换未匹配数据为专业分格式", key="convert_plan_score"):
                unmatched = [r for r in results if not r['exists']]
                if not unmatched:
                    st.warning("没有未匹配的数据")
                else:
                    try:
                        converted = convert_data_to_score_format(unmatched, st.session_state.plan_df)
                        st.session_state.converted_data = converted
                        st.session_state.conversion_source = 'planScore'
                        st.success(f"✓ 已转换 {len(converted)} 条未匹配数据")
                    except Exception as e:
                        st.error(f"转换失败: {str(e)}")
    
    # 比对2结果
    if st.session_state.plan_college_results:
        st.subheader("📊 比对2：招生计划 vs 院校分")
        
        results = st.session_state.plan_college_results
        stats = get_comparison_stats(results)
        
        # 统计信息
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总记录数", stats['total'])
        col2.metric("匹配记录数", stats['matched'], delta="✓")
        col3.metric("未匹配记录数", stats['unmatched'], delta="✗")
        col4.metric("匹配率", stats['match_rate'])
        
        # 筛选选项
        col1, col2, col3 = st.columns(3)
        
        with col1:
            provinces = ['全部'] + get_unique_provinces(results)
            selected_province = st.selectbox(
                "按省份筛选",
                provinces,
                key="plan_college_province"
            )
        
        with col2:
            batches = ['全部'] + get_unique_batches(results)
            selected_batch = st.selectbox(
                "按批次筛选",
                batches,
                key="plan_college_batch"
            )
        
        with col3:
            match_status = st.selectbox(
                "匹配状态",
                ['全部', '匹配', '未匹配'],
                key="plan_college_status"
            )
        
        # 过滤数据
        filtered_results = results
        
        if selected_province != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['省份']).strip() == selected_province]
        
        if selected_batch != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['批次']).strip() == selected_batch]
        
        if match_status == '匹配':
            filtered_results = [r for r in filtered_results if r['exists']]
        elif match_status == '未匹配':
            filtered_results = [r for r in filtered_results if not r['exists']]
        
        # 显示表格
        st.write(f"**显示 {len(filtered_results)} 条记录**")
        
        display_data = []
        for result in filtered_results[:500]:  # 限制显示500条
            row = {
                '序号': result['index'],
                '状态': '✓ 匹配' if result['exists'] else '✗ 未匹配',
                **result['key_fields']
            }
            display_data.append(row)
        
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
        
        # 导出按钮
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 导出比对结果", key="export_plan_college_results"):
                try:
                    file_bytes = export_results_to_excel(results, "plan_college_results.xlsx")
                    st.download_button(
                        label="下载 比对2 结果",
                        data=file_bytes,
                        file_name="招生计划vs院校分_比对结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"导出失败: {str(e)}")
        
        with col2:
            if st.button("🔄 转换未匹配数据为专业分格式", key="convert_plan_college"):
                unmatched = [r for r in results if not r['exists']]
                if not unmatched:
                    st.warning("没有未匹配的数据")
                else:
                    try:
                        converted = convert_data_to_score_format(unmatched, st.session_state.plan_df)
                        st.session_state.converted_data = converted
                        st.session_state.conversion_source = 'planCollege'
                        st.success(f"✓ 已转换 {len(converted)} 条未匹配数据")
                    except Exception as e:
                        st.error(f"转换失败: {str(e)}")


# ==================== 转换和导出 ====================

def conversion_export_section():
    """转换和导出部分"""
    
    if st.session_state.converted_data:
        st.subheader("🎯 未匹配数据转换")
        
        converted_data = st.session_state.converted_data
        source = st.session_state.conversion_source
        
        # 统计信息
        col1, col2, col3 = st.columns(3)
        col1.metric("待转换记录数", len(converted_data))
        col2.metric("转换来源", '比对1' if source == 'planScore' else '比对2')
        
        # 预览
        st.write("**预览前10条转换结果：**")
        preview_df = pd.DataFrame(converted_data[:10])
        st.dataframe(preview_df, use_container_width=True)
        
        # 导出按钮
        if st.button("💾 导出为专业分导入模板格式", key="export_converted"):
            try:
                # 获取招生年份
                admission_year = ''
                if st.session_state.plan_df is not None and '年份' in st.session_state.plan_df.columns:
                    admission_year = str(st.session_state.plan_df['年份'].iloc[0])
                
                file_bytes = export_converted_data_to_excel(converted_data, admission_year)
                st.download_button(
                    label="下载 未匹配数据（专业分格式）",
                    data=file_bytes,
                    file_name="未匹配数据_专业分格式.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("✓ 已生成导出文件")
            except Exception as e:
                st.error(f"导出失败: {str(e)}")


# ==================== 主UI函数 ====================

def render_ui():
    """渲染完整的招生计划比对UI"""
    # 注意：页面配置已在主文件中设置，这里不再设置
    
    # 初始化状态
    init_session_state()
    
    # 标题和说明
    st.header("🎓 招生计划数据比对与转换工具")
    st.markdown("""
    上传招生计划、专业分和院校分文件进行比对，快速定位未匹配数据，
    并可自动转换为专业分导入模板格式。
    """)
    
    # 说明
    with st.expander("📝 使用说明", expanded=False):
        st.markdown("""
        **工作流程：**
        1. **上传文件** - 上传招生计划、专业分和院校分文件
        2. **数据比对** - 执行比对1、比对2或全部比对
        3. **结果检查** - 查看匹配情况，过滤和导出结果
        4. **数据转换** - 将未匹配数据转换为专业分格式
        
        **比对字段说明：**
        - **比对1** (招生计划 vs 专业分)：年份、省份、学校、科类、批次、专业、层次、专业组代码
        - **比对2** (招生计划 vs 院校分)：年份、省份、学校、科类、批次、专业组代码
        """)
    
    # 文件上传
    load_files_section()
    
    st.divider()
    
    # 比对操作
    comparison_operations()
    
    st.divider()
    
    # 结果显示
    display_comparison_results()
    
    st.divider()
    
    # 转换导出
    conversion_export_section()
    
    st.divider()
    
    # 页脚
    st.markdown("---")
    st.markdown("© 招生计划数据比对工具 | Python + Pandas + Streamlit")
=======
"""
招生计划数据比对与转换工具 - Streamlit UI模块
提供Streamlit界面组件和用户交互逻辑
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import logging
import base64
from plan_comparison import (
    load_excel_from_bytes,
    compare_plan_vs_score,
    compare_plan_vs_college,
    get_comparison_stats,
    get_unique_provinces,
    get_unique_batches,
    convert_data_to_score_format,
    export_results_to_excel,
    export_converted_data_to_excel
)

logger = logging.getLogger(__name__)


# ==================== 初始化Session State ====================

def init_session_state():
    """初始化Streamlit Session State"""
    if 'plan_df' not in st.session_state:
        st.session_state.plan_df = None
    if 'score_df' not in st.session_state:
        st.session_state.score_df = None
    if 'college_df' not in st.session_state:
        st.session_state.college_df = None
    
    if 'plan_score_results' not in st.session_state:
        st.session_state.plan_score_results = None
    if 'plan_college_results' not in st.session_state:
        st.session_state.plan_college_results = None
    
    if 'converted_data' not in st.session_state:
        st.session_state.converted_data = None
    if 'conversion_source' not in st.session_state:
        st.session_state.conversion_source = None


# ==================== 文件加载 ====================

def load_files_section():
    """文件上传部分"""
    st.subheader("📁 文件上传")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.write("**招生计划文件**")
        plan_file = st.file_uploader("选择招生计划Excel文件", type=["xlsx", "xls"], key="plan_file")
        if plan_file:
            try:
                st.session_state.plan_df = load_excel_from_bytes(plan_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.plan_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")
    
    with col2:
        st.write("**专业分文件**")
        score_file = st.file_uploader("选择专业分Excel文件", type=["xlsx", "xls"], key="score_file")
        if score_file:
            try:
                st.session_state.score_df = load_excel_from_bytes(score_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.score_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")
    
    with col3:
        st.write("**院校分文件**")
        college_file = st.file_uploader("选择院校分Excel文件", type=["xlsx", "xls"], key="college_file")
        if college_file:
            try:
                st.session_state.college_df = load_excel_from_bytes(college_file.getvalue())
                st.success(f"✓ 已加载 {len(st.session_state.college_df)} 条记录")
            except Exception as e:
                st.error(f"加载失败: {str(e)}")


# ==================== 比对操作 ====================

def comparison_operations():
    """比对操作部分"""
    st.subheader("🔍 数据比对")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("比对1：招生计划 vs 专业分", key="compare_plan_score"):
            if st.session_state.plan_df is None:
                st.error("请先上传招生计划文件")
            elif st.session_state.score_df is None:
                st.error("请先上传专业分文件")
            else:
                with st.spinner("正在进行比对1..."):
                    try:
                        st.session_state.plan_score_results = compare_plan_vs_score(
                            st.session_state.plan_df,
                            st.session_state.score_df
                        )
                        st.success("✓ 比对1完成")
                        st.session_state.conversion_source = 'planScore'
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col2:
        if st.button("比对2：招生计划 vs 院校分", key="compare_plan_college"):
            if st.session_state.plan_df is None:
                st.error("请先上传招生计划文件")
            elif st.session_state.college_df is None:
                st.error("请先上传院校分文件")
            else:
                with st.spinner("正在进行比对2..."):
                    try:
                        st.session_state.plan_college_results = compare_plan_vs_college(
                            st.session_state.plan_df,
                            st.session_state.college_df
                        )
                        st.success("✓ 比对2完成")
                        st.session_state.conversion_source = 'planCollege'
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col3:
        if st.button("全部比对", key="compare_all"):
            has_plan = st.session_state.plan_df is not None
            has_score = st.session_state.score_df is not None
            has_college = st.session_state.college_df is not None
            
            if not has_plan:
                st.error("请先上传招生计划文件")
            elif not (has_score or has_college):
                st.error("请至少上传专业分或院校分文件")
            else:
                with st.spinner("正在执行全部比对..."):
                    try:
                        if has_score:
                            st.session_state.plan_score_results = compare_plan_vs_score(
                                st.session_state.plan_df,
                                st.session_state.score_df
                            )
                        if has_college:
                            st.session_state.plan_college_results = compare_plan_vs_college(
                                st.session_state.plan_df,
                                st.session_state.college_df
                            )
                        st.success("✓ 全部比对完成")
                    except Exception as e:
                        st.error(f"比对失败: {str(e)}")
    
    with col4:
        if st.button("重置所有数据", key="reset_all"):
            st.session_state.plan_df = None
            st.session_state.score_df = None
            st.session_state.college_df = None
            st.session_state.plan_score_results = None
            st.session_state.plan_college_results = None
            st.session_state.converted_data = None
            st.session_state.conversion_source = None
            st.success("✓ 已重置所有数据")


# ==================== 结果显示 ====================

def display_comparison_results():
    """显示比对结果"""
    
    # 比对1结果
    if st.session_state.plan_score_results:
        st.subheader("📊 比对1：招生计划 vs 专业分")
        
        results = st.session_state.plan_score_results
        stats = get_comparison_stats(results)
        
        # 统计信息
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总记录数", stats['total'])
        col2.metric("匹配记录数", stats['matched'], delta="✓")
        col3.metric("未匹配记录数", stats['unmatched'], delta="✗")
        col4.metric("匹配率", stats['match_rate'])
        
        # 筛选选项
        col1, col2, col3 = st.columns(3)
        
        with col1:
            provinces = ['全部'] + get_unique_provinces(results)
            selected_province = st.selectbox(
                "按省份筛选",
                provinces,
                key="plan_score_province"
            )
        
        with col2:
            batches = ['全部'] + get_unique_batches(results)
            selected_batch = st.selectbox(
                "按批次筛选",
                batches,
                key="plan_score_batch"
            )
        
        with col3:
            match_status = st.selectbox(
                "匹配状态",
                ['全部', '匹配', '未匹配'],
                key="plan_score_status"
            )
        
        # 过滤数据
        filtered_results = results
        
        if selected_province != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['省份']).strip() == selected_province]
        
        if selected_batch != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['批次']).strip() == selected_batch]
        
        if match_status == '匹配':
            filtered_results = [r for r in filtered_results if r['exists']]
        elif match_status == '未匹配':
            filtered_results = [r for r in filtered_results if not r['exists']]
        
        # 显示表格
        st.write(f"**显示 {len(filtered_results)} 条记录**")
        
        display_data = []
        for result in filtered_results[:500]:  # 限制显示500条
            row = {
                '序号': result['index'],
                '状态': '✓ 匹配' if result['exists'] else '✗ 未匹配',
                **result['key_fields']
            }
            display_data.append(row)
        
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
        
        # 导出按钮
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 导出比对结果", key="export_plan_score_results"):
                try:
                    file_bytes = export_results_to_excel(results, "plan_score_results.xlsx")
                    st.download_button(
                        label="下载 比对1 结果",
                        data=file_bytes,
                        file_name="招生计划vs专业分_比对结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"导出失败: {str(e)}")
        
        with col2:
            if st.button("🔄 转换未匹配数据为专业分格式", key="convert_plan_score"):
                unmatched = [r for r in results if not r['exists']]
                if not unmatched:
                    st.warning("没有未匹配的数据")
                else:
                    try:
                        converted = convert_data_to_score_format(unmatched, st.session_state.plan_df)
                        st.session_state.converted_data = converted
                        st.session_state.conversion_source = 'planScore'
                        st.success(f"✓ 已转换 {len(converted)} 条未匹配数据")
                    except Exception as e:
                        st.error(f"转换失败: {str(e)}")
    
    # 比对2结果
    if st.session_state.plan_college_results:
        st.subheader("📊 比对2：招生计划 vs 院校分")
        
        results = st.session_state.plan_college_results
        stats = get_comparison_stats(results)
        
        # 统计信息
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总记录数", stats['total'])
        col2.metric("匹配记录数", stats['matched'], delta="✓")
        col3.metric("未匹配记录数", stats['unmatched'], delta="✗")
        col4.metric("匹配率", stats['match_rate'])
        
        # 筛选选项
        col1, col2, col3 = st.columns(3)
        
        with col1:
            provinces = ['全部'] + get_unique_provinces(results)
            selected_province = st.selectbox(
                "按省份筛选",
                provinces,
                key="plan_college_province"
            )
        
        with col2:
            batches = ['全部'] + get_unique_batches(results)
            selected_batch = st.selectbox(
                "按批次筛选",
                batches,
                key="plan_college_batch"
            )
        
        with col3:
            match_status = st.selectbox(
                "匹配状态",
                ['全部', '匹配', '未匹配'],
                key="plan_college_status"
            )
        
        # 过滤数据
        filtered_results = results
        
        if selected_province != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['省份']).strip() == selected_province]
        
        if selected_batch != '全部':
            filtered_results = [r for r in filtered_results 
                               if str(r['key_fields']['批次']).strip() == selected_batch]
        
        if match_status == '匹配':
            filtered_results = [r for r in filtered_results if r['exists']]
        elif match_status == '未匹配':
            filtered_results = [r for r in filtered_results if not r['exists']]
        
        # 显示表格
        st.write(f"**显示 {len(filtered_results)} 条记录**")
        
        display_data = []
        for result in filtered_results[:500]:  # 限制显示500条
            row = {
                '序号': result['index'],
                '状态': '✓ 匹配' if result['exists'] else '✗ 未匹配',
                **result['key_fields']
            }
            display_data.append(row)
        
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
        
        # 导出按钮
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 导出比对结果", key="export_plan_college_results"):
                try:
                    file_bytes = export_results_to_excel(results, "plan_college_results.xlsx")
                    st.download_button(
                        label="下载 比对2 结果",
                        data=file_bytes,
                        file_name="招生计划vs院校分_比对结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"导出失败: {str(e)}")
        
        with col2:
            if st.button("🔄 转换未匹配数据为专业分格式", key="convert_plan_college"):
                unmatched = [r for r in results if not r['exists']]
                if not unmatched:
                    st.warning("没有未匹配的数据")
                else:
                    try:
                        converted = convert_data_to_score_format(unmatched, st.session_state.plan_df)
                        st.session_state.converted_data = converted
                        st.session_state.conversion_source = 'planCollege'
                        st.success(f"✓ 已转换 {len(converted)} 条未匹配数据")
                    except Exception as e:
                        st.error(f"转换失败: {str(e)}")


# ==================== 转换和导出 ====================

def conversion_export_section():
    """转换和导出部分"""
    
    if st.session_state.converted_data:
        st.subheader("🎯 未匹配数据转换")
        
        converted_data = st.session_state.converted_data
        source = st.session_state.conversion_source
        
        # 统计信息
        col1, col2, col3 = st.columns(3)
        col1.metric("待转换记录数", len(converted_data))
        col2.metric("转换来源", '比对1' if source == 'planScore' else '比对2')
        
        # 预览
        st.write("**预览前10条转换结果：**")
        preview_df = pd.DataFrame(converted_data[:10])
        st.dataframe(preview_df, use_container_width=True)
        
        # 导出按钮
        if st.button("💾 导出为专业分导入模板格式", key="export_converted"):
            try:
                # 获取招生年份
                admission_year = ''
                if st.session_state.plan_df is not None and '年份' in st.session_state.plan_df.columns:
                    admission_year = str(st.session_state.plan_df['年份'].iloc[0])
                
                file_bytes = export_converted_data_to_excel(converted_data, admission_year)
                st.download_button(
                    label="下载 未匹配数据（专业分格式）",
                    data=file_bytes,
                    file_name="未匹配数据_专业分格式.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("✓ 已生成导出文件")
            except Exception as e:
                st.error(f"导出失败: {str(e)}")


# ==================== 主UI函数 ====================

def render_ui():
    """渲染完整的招生计划比对UI"""
    # 注意：页面配置已在主文件中设置，这里不再设置
    
    # 初始化状态
    init_session_state()
    
    # 标题和说明
    st.header("🎓 招生计划数据比对与转换工具")
    st.markdown("""
    上传招生计划、专业分和院校分文件进行比对，快速定位未匹配数据，
    并可自动转换为专业分导入模板格式。
    """)
    
    # 说明
    with st.expander("📝 使用说明", expanded=False):
        st.markdown("""
        **工作流程：**
        1. **上传文件** - 上传招生计划、专业分和院校分文件
        2. **数据比对** - 执行比对1、比对2或全部比对
        3. **结果检查** - 查看匹配情况，过滤和导出结果
        4. **数据转换** - 将未匹配数据转换为专业分格式
        
        **比对字段说明：**
        - **比对1** (招生计划 vs 专业分)：年份、省份、学校、科类、批次、专业、层次、专业组代码
        - **比对2** (招生计划 vs 院校分)：年份、省份、学校、科类、批次、专业组代码
        """)
    
    # 文件上传
    load_files_section()
    
    st.divider()
    
    # 比对操作
    comparison_operations()
    
    st.divider()
    
    # 结果显示
    display_comparison_results()
    
    st.divider()
    
    # 转换导出
    conversion_export_section()
    
    st.divider()
    
    # 页脚
    st.markdown("---")
    st.markdown("© 招生计划数据比对工具 | Python + Pandas + Streamlit")
>>>>>>> a2d3e7d (auto-commit before pull)
