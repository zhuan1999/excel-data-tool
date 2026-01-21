import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from functools import lru_cache
import warnings
warnings.filterwarnings('ignore')

# ============================================
# 全局配置和缓存
# ============================================
st.set_page_config(
    page_title="Excel数据处理工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS优化加载速度
st.markdown("""
<style>
    /* 减少页面重绘 */
    .stApp {
        contain: content;
        will-change: transform;
    }
    
    /* 优化表格渲染 */
    .stDataFrame {
        will-change: transform;
    }
    
    /* 加载动画 */
    .loading-spinner {
        display: none;
    }
    
    /* 按钮样式优化 */
    .stButton > button {
        transition: all 0.2s ease;
    }
    
    /* 缓存状态提示 */
    .cache-status {
        font-size: 0.8rem;
        color: #666;
        font-style: italic;
    }
</style>
""", unsafe_allow_html=True)

# 初始化session state用于模块间数据传递
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'data_pipeline' not in st.session_state:
    st.session_state.data_pipeline = {}

# 缓存装饰器 - 显著提升数据处理速度
@lru_cache(maxsize=5)
def cached_read_excel(file_bytes, file_name):
    """缓存读取Excel文件，避免重复读取"""
    try:
        if file_name.endswith('.csv'):
            df = pd.read_csv(BytesIO(file_bytes))
        else:
            df = pd.read_excel(BytesIO(file_bytes))
        return df
    except Exception as e:
        st.error(f"读取文件失败: {str(e)}")
        return None

def convert_column_types(df):
    """智能转换列类型，解决日期和数字类型问题"""
    for col in df.columns:
        # 尝试转换为日期
        try:
            if df[col].dtype == 'object':
                # 尝试多种日期格式
                date_col = pd.to_datetime(df[col], errors='coerce', infer_datetime_format=True)
                if date_col.notna().sum() > 0:  # 如果有有效的日期
                    df[col] = date_col
        except:
            pass
        
        # 尝试转换为数字
        try:
            if df[col].dtype == 'object':
                numeric_col = pd.to_numeric(df[col], errors='coerce')
                if numeric_col.notna().sum() > len(df) * 0.5:  # 如果超过50%是数字
                    df[col] = numeric_col
        except:
            pass
    
    return df

# ============================================
# 侧边栏 - 数据管道管理
# ============================================
with st.sidebar:
    st.title("📁 数据管道管理")
    
    # 显示当前数据管道状态
    if st.session_state.data_pipeline:
        st.markdown("### 当前管道数据")
        for name, data_info in st.session_state.data_pipeline.items():
            with st.expander(f"📄 {name}", expanded=False):
                st.write(f"形状: {data_info['shape']}")
                st.write(f"内存: {data_info['memory_mb']:.2f} MB")
                st.write(f"更新时间: {data_info['timestamp'].strftime('%H:%M:%S')}")
                
                if st.button(f"加载到当前模块", key=f"load_{name}"):
                    st.session_state.processed_data = data_info['data'].copy()
                    st.success(f"已加载 {name} 到当前模块")
                    st.rerun()
    else:
        st.info("暂无管道数据")
    
    st.markdown("---")
    st.markdown("### 🚀 性能优化")
    
    # 性能设置
    use_cache = st.checkbox("启用数据缓存", value=True)
    optimize_memory = st.checkbox("启用内存优化", value=True)
    
    if st.button("🔄 清理缓存"):
        cached_read_excel.cache_clear()
        st.success("缓存已清理")
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📖 使用说明")
    st.info("""
    **工作流程：**
    1. 每个模块处理的数据会自动保存到数据管道
    2. 可以从侧边栏加载之前处理的数据
    3. 支持模块间的数据传递
    4. 自动处理数据类型匹配问题
    """)

# ============================================
# 主应用标题
# ============================================
st.title("🚀 Excel数据处理工具 (优化版)")

# ============================================
# 数据合并模块 - 修复日期字段问题
# ============================================
st.header("📥 数据合并模块")

# 快速操作栏
col1, col2, col3 = st.columns(3)
with col1:
    load_from_pipeline = st.checkbox("从管道加载数据", key="merge_load_pipeline")
with col2:
    clear_cache = st.button("🔄 刷新视图")
with col3:
    if st.button("💾 保存到管道", key="merge_save"):
        if 'merged_df' in locals():
            st.session_state.data_pipeline['合并结果'] = {
                'data': merged_df,
                'shape': merged_df.shape,
                'memory_mb': merged_df.memory_usage(deep=True).sum() / 1024 / 1024,
                'timestamp': pd.Timestamp.now()
            }
            st.success("已保存到数据管道")

if load_from_pipeline and st.session_state.data_pipeline:
    selected_data = st.selectbox("选择管道数据", list(st.session_state.data_pipeline.keys()))
    if selected_data:
        merged_df = st.session_state.data_pipeline[selected_data]['data'].copy()
        st.success(f"已加载 {selected_data}")
else:
    col1, col2 = st.columns(2)
    
    with col1:
        file1 = st.file_uploader("选择第一个文件", type=['xlsx', 'xls', 'csv'], key="merge_file1")
        if file1:
            with st.spinner("快速读取中..."):
                # 使用缓存读取
                df1 = cached_read_excel(file1.getvalue(), file1.name)
                if df1 is not None:
                    df1 = convert_column_types(df1)  # 转换数据类型
                    st.success(f"✅ {file1.name} - {df1.shape[0]}行×{df1.shape[1]}列")
                    
                    # 显示数据类型信息
                    date_cols = [col for col in df1.columns if pd.api.types.is_datetime64_any_dtype(df1[col])]
                    if date_cols:
                        st.info(f"📅 检测到日期字段: {', '.join(date_cols)}")
    
    with col2:
        file2 = st.file_uploader("选择第二个文件", type=['xlsx', 'xls', 'csv'], key="merge_file2")
        if file2:
            with st.spinner("快速读取中..."):
                df2 = cached_read_excel(file2.getvalue(), file2.name)
                if df2 is not None:
                    df2 = convert_column_types(df2)  # 转换数据类型
                    st.success(f"✅ {file2.name} - {df2.shape[0]}行×{df2.shape[1]}列")

if 'df1' in locals() and 'df2' in locals():
    st.markdown("### 合并设置")
    
    merge_method = st.radio(
        "选择合并方式",
        ["垂直合并（追加行）", "水平合并（连接列）", "主键合并（类似SQL JOIN）"],
        horizontal=True
    )
    
    if merge_method == "主键合并（类似SQL JOIN）":
        # 智能识别共同列
        common_cols = list(set(df1.columns) & set(df2.columns))
        if common_cols:
            col1, col2 = st.columns(2)
            with col1:
                merge_on = st.selectbox("选择主键列", common_cols)
            with col2:
                merge_how = st.selectbox("合并方式", ["inner", "left", "right", "outer"])
        else:
            st.warning("⚠️ 未找到共同列")
    
    if st.button("🚀 快速合并", type="primary", use_container_width=True):
        with st.spinner("正在合并数据..."):
            try:
                if merge_method == "垂直合并（追加行）":
                    # 对齐列名，确保所有列都存在
                    all_columns = list(set(df1.columns) | set(df2.columns))
                    df1_aligned = df1.reindex(columns=all_columns)
                    df2_aligned = df2.reindex(columns=all_columns)
                    merged_df = pd.concat([df1_aligned, df2_aligned], ignore_index=True)
                    
                elif merge_method == "水平合并（连接列）":
                    merged_df = pd.concat([df1, df2], axis=1)
                    
                elif merge_method == "主键合并（类似SQL JOIN）":
                    # 确保主键列类型一致
                    if merge_on in df1.columns and merge_on in df2.columns:
                        # 转换为字符串类型避免匹配错误
                        df1[merge_on] = df1[merge_on].astype(str).str.strip()
                        df2[merge_on] = df2[merge_on].astype(str).str.strip()
                        merged_df = pd.merge(df1, df2, on=merge_on, how=merge_how, suffixes=('_表1', '_表2'))
                
                # 处理日期字段 - 确保日期格式正确
                for col in merged_df.columns:
                    if merged_df[col].dtype == 'object':
                        # 尝试转换为日期
                        try:
                            date_series = pd.to_datetime(merged_df[col], errors='ignore')
                            if date_series.dtype != 'object':  # 如果转换成功
                                merged_df[col] = date_series
                        except:
                            pass
                
                st.success(f"✅ 合并完成！共 {merged_df.shape[0]} 行 × {merged_df.shape[1]} 列")
                
                # 保存到session state
                st.session_state.processed_data = merged_df.copy()
                
                # 显示结果
                with st.expander("📊 查看合并结果", expanded=True):
                    st.dataframe(merged_df.head(50), use_container_width=True)
                    
                    # 数据统计
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总行数", merged_df.shape[0])
                    with col2:
                        st.metric("总列数", merged_df.shape[1])
                    with col3:
                        missing_total = merged_df.isnull().sum().sum()
                        st.metric("缺失值", missing_total)
                
                # 下载按钮
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    merged_df.to_excel(writer, index=False, sheet_name='合并数据')
                
                st.download_button(
                    label="📥 下载合并结果",
                    data=output.getvalue(),
                    file_name="合并数据.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # 自动保存到管道
                st.session_state.data_pipeline['合并结果'] = {
                    'data': merged_df,
                    'shape': merged_df.shape,
                    'memory_mb': merged_df.memory_usage(deep=True).sum() / 1024 / 1024,
                    'timestamp': pd.Timestamp.now()
                }
                
            except Exception as e:
                st.error(f"❌ 合并错误: {str(e)}")
                st.code(f"错误详情: {e.__class__.__name__}")

# ============================================
# 数据匹配模块 - 修复类型匹配问题
# ============================================
st.header("🔄 数据匹配模块")

# 快速选择数据源
data_source_option = st.radio(
    "选择数据源",
    ["从管道加载数据", "上传新文件"],
    horizontal=True,
    key="match_source"
)

if data_source_option == "从管道加载数据" and st.session_state.data_pipeline:
    # 选择管道数据
    pipeline_options = list(st.session_state.data_pipeline.keys())
    selected_main = st.selectbox("选择主表数据", pipeline_options, key="main_from_pipe")
    selected_lookup = st.selectbox("选择匹配表数据", pipeline_options, key="lookup_from_pipe")
    
    if selected_main and selected_lookup:
        main_df = st.session_state.data_pipeline[selected_main]['data'].copy()
        lookup_df = st.session_state.data_pipeline[selected_lookup]['data'].copy()
        st.success(f"✅ 已加载: 主表({selected_main}), 匹配表({selected_lookup})")
else:
    # 传统文件上传
    col1, col2 = st.columns(2)
    
    with col1:
        main_file = st.file_uploader("上传主表", type=['xlsx', 'xls', 'csv'], key="match_main")
        if main_file:
            with st.spinner("加载主表..."):
                main_df = cached_read_excel(main_file.getvalue(), main_file.name)
                if main_df is not None:
                    main_df = convert_column_types(main_df)
                    st.success(f"✅ {main_file.name} - {main_df.shape[0]}行")
    
    with col2:
        lookup_file = st.file_uploader("上传匹配表", type=['xlsx', 'xls', 'csv'], key="match_lookup")
        if lookup_file:
            with st.spinner("加载匹配表..."):
                lookup_df = cached_read_excel(lookup_file.getvalue(), lookup_file.name)
                if lookup_df is not None:
                    lookup_df = convert_column_types(lookup_df)
                    st.success(f"✅ {lookup_file.name} - {lookup_df.shape[0]}行")

if 'main_df' in locals() and 'lookup_df' in locals():
    st.markdown("### 🔍 匹配设置")
    
    # 快速字段选择
    col1, col2 = st.columns(2)
    
    with col1:
        main_columns = list(main_df.columns)
        main_key = st.selectbox(
            "选择主表匹配字段",
            main_columns,
            help="用于匹配的关键字段",
            key="main_key_select"
        )
        
        # 显示字段信息
        if main_key:
            dtype_main = str(main_df[main_key].dtype)
            unique_count = main_df[main_key].nunique()
            st.caption(f"类型: {dtype_main} | 唯一值: {unique_count}")
    
    with col2:
        lookup_columns = list(lookup_df.columns)
        lookup_key = st.selectbox(
            "选择匹配表对应字段",
            lookup_columns,
            help="匹配表中对应的关键字段",
            key="lookup_key_select"
        )
        
        if lookup_key:
            dtype_lookup = str(lookup_df[lookup_key].dtype)
            unique_count = lookup_df[lookup_key].nunique()
            st.caption(f"类型: {dtype_lookup} | 唯一值: {unique_count}")
    
    # 类型兼容性检查
    if 'main_key' in locals() and 'lookup_key' in locals():
        # 自动检测类型是否需要转换
        if main_df[main_key].dtype != lookup_df[lookup_key].dtype:
            st.warning(f"⚠️ 字段类型不匹配: 主表({main_df[main_key].dtype}) vs 匹配表({lookup_df[lookup_key].dtype})")
            
            # 自动转换选项
            auto_fix = st.checkbox("🔄 自动转换类型为字符串", value=True)
            if auto_fix:
                with st.spinner("正在标准化字段类型..."):
                    main_df[main_key] = main_df[main_key].astype(str).str.strip()
                    lookup_df[lookup_key] = lookup_df[lookup_key].astype(str).str.strip()
                st.success("✅ 字段类型已统一为字符串")
    
    # 选择要匹配的字段
    available_fields = [col for col in lookup_df.columns if col != lookup_key]
    if available_fields:
        selected_fields = st.multiselect(
            "选择要匹配的字段",
            available_fields,
            help="选择需要从匹配表添加到主表的字段"
        )
        
        if selected_fields:
            # 快速汇总方式选择
            st.markdown("### ⚙️ 汇总设置")
            
            # 批量设置相同汇总方式
            col1, col2 = st.columns([2, 1])
            with col1:
                batch_agg = st.selectbox(
                    "批量设置汇总方式",
                    ["第一个值", "求和", "平均值", "最大值", "最小值", "计数"],
                    help="为所有选中字段设置相同的汇总方式"
                )
            
            with col2:
                if st.button("应用批量设置", use_container_width=True):
                    st.success(f"已为所有字段应用'{batch_agg}'")
            
            # 显示字段设置
            field_settings = {}
            for field in selected_fields:
                with st.expander(f"字段: {field}", expanded=False):
                    col1, col2 = st.columns(2)
                    with col1:
                        if lookup_df[field].dtype in ['int64', 'float64', 'Int64', 'Float64']:
                            agg_options = ["第一个值", "求和", "平均值", "最大值", "最小值", "计数"]
                            default_idx = agg_options.index(batch_agg) if batch_agg in agg_options else 0
                        else:
                            agg_options = ["第一个值", "计数", "连接(逗号分隔)"]
                            default_idx = agg_options.index(batch_agg) if batch_agg in agg_options else 0
                        
                        agg_method = st.selectbox(
                            "汇总方式",
                            agg_options,
                            index=default_idx,
                            key=f"agg_{field}"
                        )
                    
                    with col2:
                        new_name = st.text_input(
                            "新列名",
                            value=f"匹配_{field}",
                            key=f"name_{field}"
                        )
                    
                    field_settings[field] = {
                        'agg': agg_method,
                        'new_name': new_name
                    }
            
            # 执行匹配
            if st.button("🚀 执行智能匹配", type="primary", use_container_width=True):
                with st.spinner("正在执行匹配..."):
                    try:
                        # 创建结果数据框
                        result_df = main_df.copy()
                        
                        # 处理每个字段
                        for field, settings in field_settings.items():
                            agg_method = settings['agg']
                            new_col_name = settings['new_name']
                            
                            # 准备匹配数据
                            temp_lookup = lookup_df[[lookup_key, field]].copy()
                            
                            # 处理重复键
                            if agg_method == "第一个值":
                                temp_lookup = temp_lookup.drop_duplicates(subset=[lookup_key], keep='first')
                            elif agg_method == "求和":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].sum().reset_index()
                            elif agg_method == "平均值":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].mean().reset_index()
                            elif agg_method == "最大值":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].max().resetindex()
                            elif agg_method == "最小值":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].min().reset_index()
                            elif agg_method == "计数":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].count().reset_index()
                            elif agg_method == "连接(逗号分隔)":
                                temp_lookup = temp_lookup.groupby(lookup_key)[field].apply(
                                    lambda x: ', '.join([str(i) for i in x if pd.notna(i)])
                                ).reset_index()
                            
                            # 确保字段类型一致
                            result_df[main_key] = result_df[main_key].astype(str).str.strip()
                            temp_lookup[lookup_key] = temp_lookup[lookup_key].astype(str).str.strip()
                            
                            # 执行匹配
                            result_df = result_df.merge(
                                temp_lookup,
                                how='left',
                                left_on=main_key,
                                right_on=lookup_key,
                                suffixes=('', '_match')
                            )
                            
                            # 重命名匹配的列
                            if field in result_df.columns:
                                result_df = result_df.rename(columns={field: new_col_name})
                        
                        # 清理多余的列
                        if lookup_key in result_df.columns and lookup_key != main_key:
                            result_df = result_df.drop(columns=[lookup_key])
                        
                        # 去重
                        result_df = result_df.loc[:, ~result_df.columns.duplicated()]
                        
                        st.success(f"✅ 匹配完成！共 {result_df.shape[0]} 行 × {result_df.shape[1]} 列")
                        
                        # 显示结果
                        with st.expander("📊 查看匹配结果", expanded=True):
                            st.dataframe(result_df.head(50), use_container_width=True)
                            
                            # 匹配统计
                            matched_cols = [settings['new_name'] for settings in field_settings.values()]
                            if matched_cols:
                                matched_count = result_df[matched_cols].notna().any(axis=1).sum()
                                match_rate = (matched_count / len(result_df)) * 100
                                
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.metric("匹配成功行数", matched_count)
                                with col2:
                                    st.metric("匹配成功率", f"{match_rate:.1f}%")
                        
                        # 下载结果
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='匹配结果')
                        
                        st.download_button(
                            label="📥 下载匹配结果",
                            data=output.getvalue(),
                            file_name="匹配结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        # 保存到管道
                        st.session_state.processed_data = result_df.copy()
                        st.session_state.data_pipeline['匹配结果'] = {
                            'data': result_df,
                            'shape': result_df.shape,
                            'memory_mb': result_df.memory_usage(deep=True).sum() / 1024 / 1024,
                            'timestamp': pd.Timestamp.now()
                        }
                        
                    except Exception as e:
                        st.error(f"❌ 匹配错误: {str(e)}")
                        st.code(f"错误详情:\n{e.__class__.__name__}: {str(e)}")
    else:
        st.warning("匹配表中没有其他可匹配的字段")

# ============================================
# 数据清洗模块
# ============================================
st.header("🧹 数据清洗模块")

# 数据源选择
clean_source = st.radio(
    "选择清洗数据源",
    ["从管道加载数据", "上传新文件", "使用上一个模块结果"],
    horizontal=True,
    key="clean_source"
)

if clean_source == "从管道加载数据" and st.session_state.data_pipeline:
    pipeline_options = list(st.session_state.data_pipeline.keys())
    selected_clean = st.selectbox("选择要清洗的数据", pipeline_options, key="clean_from_pipe")
    if selected_clean:
        clean_df = st.session_state.data_pipeline[selected_clean]['data'].copy()
        st.success(f"✅ 已加载: {selected_clean}")
elif clean_source == "使用上一个模块结果" and st.session_state.processed_data is not None:
    clean_df = st.session_state.processed_data.copy()
    st.success("✅ 已加载上一个模块的结果")
else:
    clean_file = st.file_uploader("上传要清洗的文件", type=['xlsx', 'xls', 'csv'], key="clean_file")
    if clean_file:
        with st.spinner("快速加载中..."):
            clean_df = cached_read_excel(clean_file.getvalue(), clean_file.name)
            if clean_df is not None:
                clean_df = convert_column_types(clean_df)
                st.success(f"✅ {clean_file.name} - {clean_df.shape[0]}行×{clean_df.shape[1]}列")

if 'clean_df' in locals():
    st.markdown("### 📊 数据概览")
    
    # 快速统计
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("总行数", clean_df.shape[0])
    with col2:
        st.metric("总列数", clean_df.shape[1])
    with col3:
        missing_total = clean_df.isnull().sum().sum()
        st.metric("缺失值", missing_total)
    with col4:
        dup_rows = clean_df.duplicated().sum()
        st.metric("重复行", dup_rows)
    
    # 快速清洗操作
    st.markdown("### ⚡ 快速清洗操作")
    
    quick_ops = st.columns(5)
    
    with quick_ops[0]:
        if st.button("删除重复行", use_container_width=True):
            original_len = len(clean_df)
            clean_df = clean_df.drop_duplicates()
            removed = original_len - len(clean_df)
            st.success(f"✅ 已删除 {removed} 行重复数据")
            st.rerun()
    
    with quick_ops[1]:
        if st.button("删除空值行", use_container_width=True):
            original_len = len(clean_df)
            clean_df = clean_df.dropna()
            removed = original_len - len(clean_df)
            st.success(f"✅ 已删除 {removed} 行空值数据")
            st.rerun()
    
    with quick_ops[2]:
        if st.button("重置索引", use_container_width=True):
            clean_df = clean_df.reset_index(drop=True)
            st.success("✅ 索引已重置")
            st.rerun()
    
    with quick_ops[3]:
        if st.button("标准化列名", use_container_width=True):
            clean_df.columns = [str(col).strip().replace(' ', '_') for col in clean_df.columns]
            st.success("✅ 列名已标准化")
            st.rerun()
    
    with quick_ops[4]:
        if st.button("预览数据", use_container_width=True):
            with st.expander("数据预览", expanded=True):
                st.dataframe(clean_df.head(20), use_container_width=True)
    
    # 高级清洗选项
    st.markdown("### 🔧 高级清洗选项")
    
    tab1, tab2, tab3, tab4 = st.tabs(["列操作", "行操作", "数据类型", "批量处理"])
    
    with tab1:
        col_operation = st.selectbox(
            "选择列操作",
            ["重命名列", "删除列", "移动列", "提取列"]
        )
        
        if col_operation == "重命名列":
            col_to_rename = st.selectbox("选择要重命名的列", clean_df.columns)
            new_name = st.text_input("新列名", value=col_to_rename)
            if st.button("执行重命名", key="rename_col"):
                clean_df = clean_df.rename(columns={col_to_rename: new_name})
                st.success(f"✅ 已重命名为: {new_name}")
                st.rerun()
        
        elif col_operation == "删除列":
            cols_to_drop = st.multiselect("选择要删除的列", clean_df.columns)
            if cols_to_drop and st.button("执行删除", key="drop_cols"):
                clean_df = clean_df.drop(columns=cols_to_drop)
                st.success(f"✅ 已删除 {len(cols_to_drop)} 列")
                st.rerun()
    
    with tab2:
        row_operation = st.selectbox(
            "选择行操作",
            ["删除空行", "删除重复", "筛选行", "排序"]
        )
        
        if row_operation == "排序":
            sort_col = st.selectbox("按哪列排序", clean_df.columns)
            sort_asc = st.checkbox("升序排序", value=True)
            if st.button("执行排序", key="sort_rows"):
                clean_df = clean_df.sort_values(by=sort_col, ascending=sort_asc)
                st.success("✅ 排序完成")
                st.rerun()
    
    with tab3:
        dtype_operation = st.selectbox(
            "数据类型转换",
            ["自动检测类型", "转换为字符串", "转换为数值", "转换为日期"]
        )
        
        col_to_convert = st.selectbox("选择要转换的列", clean_df.columns)
        
        if st.button("执行转换", key="convert_dtype"):
            if dtype_operation == "自动检测类型":
                clean_df[col_to_convert] = pd.to_numeric(
                    clean_df[col_to_convert], errors='ignore'
                )
            elif dtype_operation == "转换为字符串":
                clean_df[col_to_convert] = clean_df[col_to_convert].astype(str)
            elif dtype_operation == "转换为数值":
                clean_df[col_to_convert] = pd.to_numeric(
                    clean_df[col_to_convert], errors='coerce'
                )
            elif dtype_operation == "转换为日期":
                clean_df[col_to_convert] = pd.to_datetime(
                    clean_df[col_to_convert], errors='coerce'
                )
            
            st.success(f"✅ 类型转换完成")
            st.rerun()
    
    # 最终结果和下载
    st.markdown("### 💾 最终结果")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        with st.expander("清洗后数据预览", expanded=True):
            st.dataframe(clean_df.head(30), use_container_width=True)
    
    with col2:
        # 保存到管道
        if st.button("💾 保存到管道", use_container_width=True):
            st.session_state.data_pipeline['清洗结果'] = {
                'data': clean_df,
                'shape': clean_df.shape,
                'memory_mb': clean_df.memory_usage(deep=True).sum() / 1024 / 1024,
                'timestamp': pd.Timestamp.now()
            }
            st.success("✅ 已保存到数据管道")
        
        # 下载按钮
        output = BytesIO()
        clean_df.to_excel(output, index=False)
        
        st.download_button(
            label="📥 下载清洗结果",
            data=output.getvalue(),
            file_name="清洗结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ============================================
# 页脚
# ============================================
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9rem;'>
    🚀 优化版数据处理工具 | 支持数据管道传递 | 自动类型转换 | 快速处理
</div>
""", unsafe_allow_html=True)
