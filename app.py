import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="数据清洗工具",
    page_icon="📊",
    layout="wide"
)

# 自定义样式
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; color: #1E3A8A; text-align: center; margin-bottom: 2rem; }
    .sub-header { font-size: 1.8rem; color: #3B82F6; margin-top: 2rem; margin-bottom: 1rem; }
    .success-box { background-color: #D1FAE5; padding: 1rem; border-radius: 0.5rem; border-left: 4px solid #10B981; }
    .info-box { background-color: #DBEAFE; padding: 1rem; border-radius: 0.5rem; border-left: 4px solid #3B82F6; }
    .stButton > button { background-color: #3B82F6; color: white; font-weight: bold; border: none; padding: 0.5rem 2rem; border-radius: 0.5rem; }
    .stButton > button:hover { background-color: #2563EB; }
</style>
""", unsafe_allow_html=True)

# 应用标题
st.markdown('<h1 class="main-header">📊 Excel数据清洗工具</h1>', unsafe_allow_html=True)

# 创建标签页
tab1, tab2, tab3 = st.tabs(["📥 数据合并", "🔄 数据匹配", "🧹 数据清洗"])

# ============================================
# 1. 数据合并模块
# ============================================
with tab1:
    st.markdown('<h2 class="sub-header">数据合并</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 上传第一个文件")
        file1 = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'], key="file1")
        if file1:
            try:
                df1 = pd.read_excel(file1)
                st.success(f"✅ 成功读取: {file1.name}")
                st.info(f"📊 维度: {df1.shape[0]} 行 × {df1.shape[1]} 列")
                with st.expander("📋 预览"):
                    st.dataframe(df1.head())
            except Exception as e:
                st.error(f"❌ 错误: {str(e)}")
    
    with col2:
        st.markdown("### 上传第二个文件")
        file2 = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'], key="file2")
        if file2:
            try:
                df2 = pd.read_excel(file2)
                st.success(f"✅ 成功读取: {file2.name}")
                st.info(f"📊 维度: {df2.shape[0]} 行 × {df2.shape[1]} 列")
                with st.expander("📋 预览"):
                    st.dataframe(df2.head())
            except Exception as e:
                st.error(f"❌ 错误: {str(e)}")
    
    if 'file1' in locals() and 'file2' in locals():
        st.markdown("### 合并设置")
        
        merge_type = st.selectbox(
            "选择合并方式",
            ["垂直合并（上下堆叠）", "水平合并（左右连接）", "根据共同列合并"]
        )
        
        if merge_type == "根据共同列合并":
            common_cols = list(set(df1.columns) & set(df2.columns))
            if common_cols:
                merge_on = st.selectbox("选择合并列", common_cols)
                merge_how = st.selectbox("合并方式", ["inner", "left", "right", "outer"])
            else:
                st.warning("⚠️ 没有共同列")
        
        if st.button("🔄 合并数据", type="primary"):
            try:
                if merge_type == "垂直合并（上下堆叠）":
                    merged_df = pd.concat([df1, df2], ignore_index=True)
                elif merge_type == "水平合并（左右连接）":
                    merged_df = pd.concat([df1, df2], axis=1)
                elif merge_type == "根据共同列合并" and 'merge_on' in locals():
                    merged_df = pd.merge(df1, df2, on=merge_on, how=merge_how)
                
                st.success(f"✅ 合并成功！{merged_df.shape[0]} 行 × {merged_df.shape[1]} 列")
                
                with st.expander("🔍 查看结果"):
                    st.dataframe(merged_df.head())
                
                # 下载
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    merged_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 下载结果",
                    data=output.getvalue(),
                    file_name="合并数据.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ 合并错误: {str(e)}")

# ============================================
# 2. 数据匹配模块 (VLOOKUP功能)
# ============================================
with tab2:
    st.markdown('<h2 class="sub-header">数据匹配 (类似VLOOKUP)</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 上传主表 (X1)")
        main_file = st.file_uploader("选择主表文件", type=['xlsx', 'xls'], key="main_file")
        if main_file:
            try:
                main_df = pd.read_excel(main_file)
                st.success(f"✅ 成功读取主表")
                st.info(f"📊 维度: {main_df.shape[0]} 行 × {main_df.shape[1]} 列")
                with st.expander("📋 预览"):
                    st.dataframe(main_df.head())
            except Exception as e:
                st.error(f"❌ 错误: {str(e)}")
    
    with col2:
        st.markdown("### 上传匹配表")
        lookup_file = st.file_uploader("选择匹配表文件", type=['xlsx', 'xls'], key="lookup_file")
        if lookup_file:
            try:
                lookup_df = pd.read_excel(lookup_file)
                st.success(f"✅ 成功读取匹配表")
                st.info(f"📊 维度: {lookup_df.shape[0]} 行 × {lookup_df.shape[1]} 列")
                with st.expander("📋 预览"):
                    st.dataframe(lookup_df.head())
            except Exception as e:
                st.error(f"❌ 错误: {str(e)}")
    
    if 'main_df' in locals() and 'lookup_df' in locals():
        st.markdown("### 匹配设置")
        
        col1, col2 = st.columns(2)
        with col1:
            main_columns = list(main_df.columns)
            lookup_key = st.selectbox("🔍 选择主表查找字段", main_columns)
        
        with col2:
            lookup_columns = list(lookup_df.columns)
            match_key = st.selectbox("🎯 选择匹配表对应字段", lookup_columns)
        
        # 选择匹配字段
        available_fields = [col for col in lookup_columns if col != match_key]
        if available_fields:
            selected_fields = st.multiselect(
                "📋 选择要匹配的字段（可多选）",
                available_fields
            )
            
            if selected_fields:
                # 智能汇总方式选择
                field_settings = {}
                for field in selected_fields:
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.markdown(f"**{field}**")
                    with col2:
                        if lookup_df[field].dtype in ['int64', 'float64']:
                            agg_method = st.selectbox(
                                "汇总方式",
                                ["求和", "平均值", "最大值", "最小值", "第一个值"],
                                key=f"agg_{field}"
                            )
                        else:
                            agg_method = st.selectbox(
                                "汇总方式",
                                ["第一个值", "连接（逗号分隔）"],
                                key=f"agg_{field}"
                            )
                        field_settings[field] = agg_method
                
                if st.button("🔗 执行匹配", type="primary"):
                    try:
                        result_df = main_df.copy()
                        
                        for field, agg_method in field_settings.items():
                            temp_lookup = lookup_df[[match_key, field]].copy()
                            
                            if agg_method == "求和":
                                temp_lookup = temp_lookup.groupby(match_key)[field].sum().reset_index()
                            elif agg_method == "平均值":
                                temp_lookup = temp_lookup.groupby(match_key)[field].mean().reset_index()
                            elif agg_method == "最大值":
                                temp_lookup = temp_lookup.groupby(match_key)[field].max().reset_index()
                            elif agg_method == "最小值":
                                temp_lookup = temp_lookup.groupby(match_key)[field].min().reset_index()
                            elif agg_method == "连接（逗号分隔）":
                                temp_lookup = temp_lookup.groupby(match_key)[field].apply(lambda x: ', '.join(map(str, x))).reset_index()
                            else:  # 第一个值
                                temp_lookup = lookup_df[[match_key, field]].drop_duplicates(subset=match_key, keep='first')
                            
                            result_df = pd.merge(result_df, temp_lookup, left_on=lookup_key, right_on=match_key, how='left')
                            result_df = result_df.rename(columns={field: f"匹配_{field}"})
                        
                        # 移除多余的匹配键列
                        if match_key in result_df.columns and match_key != lookup_key:
                            result_df = result_df.drop(columns=[match_key])
                        
                        st.success(f"✅ 匹配成功！{result_df.shape[0]} 行 × {result_df.shape[1]} 列")
                        
                        # 统计
                        matched_count = result_df[result_df[[f"匹配_{f}" for f in selected_fields]].notna().any(axis=1)].shape[0]
                        total_count = result_df.shape[0]
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("匹配成功行数", matched_count)
                        with col2:
                            st.metric("匹配成功率", f"{matched_count/total_count*100:.1f}%")
                        
                        with st.expander("🔍 查看匹配结果"):
                            st.dataframe(result_df.head())
                        
                        # 下载
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False)
                        
                        st.download_button(
                            label="📥 下载匹配结果",
                            data=output.getvalue(),
                            file_name="匹配结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ 匹配错误: {str(e)}")
        else:
            st.warning("⚠️ 匹配表中没有其他可匹配的字段")

# ============================================
# 3. 数据清洗模块
# ============================================
with tab3:
    st.markdown('<h2 class="sub-header">数据清洗</h2>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("📤 上传需要清洗的Excel文件", type=['xlsx', 'xls'], key="clean_file")
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 成功读取: {uploaded_file.name}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("行数", df.shape[0])
            with col2:
                st.metric("列数", df.shape[1])
            
            with st.expander("📋 数据预览"):
                st.dataframe(df.head())
            
            st.markdown("### 🔍 数据概览")
            
            # 字段信息
            columns = list(df.columns)
            st.write("**字段列表:**")
            for i, col in enumerate(columns):
                dtype = str(df[col].dtype)
                missing = df[col].isnull().sum()
                st.write(f"{i+1}. **{col}** - 类型: `{dtype}` - 缺失值: {missing}")
            
            st.markdown("### 🛠️ 清洗操作")
            
            operation = st.selectbox(
                "选择清洗操作",
                ["删除列", "重命名列", "处理缺失值", "删除重复行", "数据类型转换"]
            )
            
            if operation == "删除列":
                cols_to_drop = st.multiselect("选择要删除的列", columns)
                if cols_to_drop and st.button("🗑️ 删除列", type="secondary"):
                    df = df.drop(columns=cols_to_drop)
                    st.success(f"✅ 已删除 {len(cols_to_drop)} 列")
                    st.rerun()
            
            elif operation == "重命名列":
                col_to_rename = st.selectbox("选择要重命名的列", columns)
                new_name = st.text_input("新列名", value=col_to_rename)
                if new_name and st.button("✏️ 重命名", type="secondary"):
                    df = df.rename(columns={col_to_rename: new_name})
                    st.success(f"✅ 已重命名为 '{new_name}'")
                    st.rerun()
            
            elif operation == "处理缺失值":
                cols_with_missing = [col for col in columns if df[col].isnull().any()]
                if cols_with_missing:
                    col_to_fill = st.selectbox("选择要处理的列", cols_with_missing)
                    
                    fill_method = st.selectbox(
                        "填充方法",
                        ["删除行", "向前填充", "向后填充", "填充固定值", "填充均值", "填充中位数"]
                    )
                    
                    if fill_method == "填充固定值":
                        fill_value = st.text_input("填充值")
                    
                    if st.button("🔧 处理缺失值", type="secondary"):
                        na_count = df[col_to_fill].isnull().sum()
                        
                        if fill_method == "删除行":
                            df = df.dropna(subset=[col_to_fill])
                            st.success(f"✅ 已删除 {na_count} 行")
                        elif fill_method == "向前填充":
                            df[col_to_fill] = df[col_to_fill].ffill()
                            st.success(f"✅ 已向前填充 {na_count} 个缺失值")
                        elif fill_method == "向后填充":
                            df[col_to_fill] = df[col_to_fill].bfill()
                            st.success(f"✅ 已向后填充 {na_count} 个缺失值")
                        elif fill_method == "填充固定值" and 'fill_value' in locals():
                            df[col_to_fill] = df[col_to_fill].fillna(fill_value)
                            st.success(f"✅ 已填充 {na_count} 个缺失值")
                        elif fill_method == "填充均值":
                            if pd.api.types.is_numeric_dtype(df[col_to_fill]):
                                df[col_to_fill] = df[col_to_fill].fillna(df[col_to_fill].mean())
                                st.success(f"✅ 已用均值填充 {na_count} 个缺失值")
                            else:
                                st.error("❌ 该列不是数值类型")
                        elif fill_method == "填充中位数":
                            if pd.api.types.is_numeric_dtype(df[col_to_fill]):
                                df[col_to_fill] = df[col_to_fill].fillna(df[col_to_fill].median())
                                st.success(f"✅ 已用中位数填充 {na_count} 个缺失值")
                            else:
                                st.error("❌ 该列不是数值类型")
                        
                        st.rerun()
                else:
                    st.info("✅ 没有发现缺失值")
            
            elif operation == "删除重复行":
                subset = st.multiselect("基于哪些列检查重复（留空则检查所有列）", columns)
                keep = st.selectbox("保留哪一行", ["第一行", "最后一行"])
                
                if st.button("🌀 删除重复", type="secondary"):
                    original_len = len(df)
                    df = df.drop_duplicates(
                        subset=subset if subset else None,
                        keep='first' if keep == "第一行" else 'last'
                    )
                    removed = original_len - len(df)
                    st.success(f"✅ 已删除 {removed} 行重复数据")
                    st.rerun()
            
            elif operation == "数据类型转换":
                col_to_convert = st.selectbox("选择要转换的列", columns)
                target_type = st.selectbox(
                    "目标数据类型",
                    ["字符串", "整数", "浮点数", "日期时间"]
                )
                
                if st.button("🔄 转换类型", type="secondary"):
                    try:
                        original_dtype = str(df[col_to_convert].dtype)
                        
                        if target_type == "字符串":
                            df[col_to_convert] = df[col_to_convert].astype(str)
                        elif target_type == "整数":
                            df[col_to_convert] = pd.to_numeric(df[col_to_convert], errors='coerce').astype('Int64')
                        elif target_type == "浮点数":
                            df[col_to_convert] = pd.to_numeric(df[col_to_convert], errors='coerce').astype(float)
                        elif target_type == "日期时间":
                            df[col_to_convert] = pd.to_datetime(df[col_to_convert], errors='coerce')
                        
                        st.success(f"✅ 已将 '{col_to_convert}' 从 {original_dtype} 转换为 {target_type}")
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ 转换失败: {str(e)}")
            
            # 显示清洗后数据
            st.markdown("### 📊 清洗后数据")
            st.dataframe(df.head())
            
            # 下载
            output = BytesIO()
            df.to_excel(output, index=False)
            
            st.download_button(
                label="📥 下载清洗后数据",
                data=output.getvalue(),
                file_name="清洗后数据.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ 读取错误: {str(e)}")

# ============================================
# 侧边栏说明
# ============================================
with st.sidebar:
    st.markdown("## 📖 使用说明")
    
    with st.expander("🔧 数据合并模块"):
        st.markdown("""
        **功能：** 合并两个Excel文件
        **步骤：**
        1. 上传两个Excel文件
        2. 选择合并方式
        3. 点击"合并数据"按钮
        """)
    
    with st.expander("🔗 数据匹配模块"):
        st.markdown("""
        **功能：** 类似Excel的VLOOKUP
        **步骤：**
        1. 上传主表和匹配表
        2. 自动识别字段并选择
        3. 选择要匹配的字段
        4. 设置汇总方式
        5. 点击"执行匹配"按钮
        """)
    
    with st.expander("🧹 数据清洗模块"):
        st.markdown("""
        **功能：** 数据清洗和预处理
        **可用操作：**
        - 删除不需要的列
        - 重命名列
        - 处理缺失值
        - 删除重复行
        - 转换数据类型
        """)
    
    st.markdown("---")
    st.markdown("### 📊 支持的文件格式")
    st.markdown("- Excel文件 (.xlsx, .xls)")
    
    st.markdown("---")
    st.markdown("### ⚠️ 注意事项")
    st.markdown("""
    1. 最大文件大小：200MB
    2. 处理大文件时可能需要较长时间
    3. 所有操作在浏览器中完成，数据不会上传到服务器
    4. 建议在处理前备份原始数据
    """)
