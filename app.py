import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="空巢青年智能厨具 KANO分析系统",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义样式（更美观）
st.markdown("""
<style>
    .main { background-color: #f5f7fa; }
    .stButton>button { 
        background-color: #4CAF50; 
        color: white; 
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
    }
    .stDownloadButton>button {
        background-color: #2196F3;
        color: white;
    }
    .kano-m { color: #d62728; font-weight: bold; }  /* 必备-红 */
    .kano-o { color: #ff7f0e; font-weight: bold; }  /* 期望-橙 */
    .kano-a { color: #2ca02c; font-weight: bold; }  /* 魅力-绿 */
    .kano-i { color: #1f77b4; font-weight: bold; }  /* 无差异-蓝 */
</style>
""", unsafe_allow_html=True)

# ==================== 配置信息（与原脚本一致）====================
FUNC_MAP = {
    "省心一键菜谱": [
        "1、正向题 Q1：当该智能烹饪器具具备省心一键菜谱功能时，您的感受是？",
        "2、反向题 Q2：当该智能烹饪器具不具备省心一键菜谱功能时，您的感受是？"
    ],
    "安全监护与异常预警": [
        "3、正向题 Q3：当该智能烹饪器具具备安全监护与异常预警功能时，您的感受是？",
        "4、反向题 Q4：当该智能烹饪器具不具备安全监护与异常预警功能时，您的感受是？"
    ],
    "定时预约自动完成": [
        "5、正向题 Q5：当该智能烹饪器具具备定时预约自动完成功能时，您的感受是？",
        "6、反向题 Q6：当该智能烹饪器具不具备定时预约自动完成功能时，您的感受是？"
    ],
    "健康关怀与提醒": [
        "7、正向题 Q7：当该智能烹饪器具具备健康关怀与提醒功能时，您的感受是？",
        "8、反向题 Q8：当该智能烹饪器具不具备健康关怀与提醒功能时，您的感受是？"
    ],
    "轻量化使用": [
        "9、正向题 Q9：当该智能烹饪器具具备轻量化使用功能时，您的感受是？",
        "10、反向题 Q10：当该智能烹饪器具不具备轻量化使用功能时，您的感受是？"
    ],
    "防油烟": [
        "11、正向题 Q15：当该智能烹饪器具具备防油烟功能时，您的感受是？",
        "12、反向题 Q16：当该智能烹饪器具不具备防油烟功能时，您的感受是？"
    ],
    "自动保温": [
        "13、正向题 Q17：当该智能烹饪器具具备自动保温功能时，您的感受是？",
        "14、反向题 Q18：当该智能烹饪器具不具备自动保温功能时，您的感受是？"
    ],
    "一人食小容量": [
        "15、正向题 Q19：当该智能烹饪器具具备一人食小容量功能时，您的感受是？",
        "16、反向题 Q20：当该智能烹饪器具不具备一人食小容量功能时，您的感受是？"
    ],
    "极致安全": [
        "17、正向题 Q21：当该智能烹饪器具具备极致安全功能时，您的感受是？",
        "18、反向题 Q22：当该智能烹饪器具不具备极致安全功能时，您的感受是？"
    ],
    "易清洁少维护": [
        "19、正向题 Q23：当该智能烹饪器具具备易清洁少维护功能时，您的感受是？",
        "20、反向题 Q24：当该智能烹饪器具不具备易清洁少维护功能时，您的感受是？"
    ],
    "低噪音低干扰": [
        "21、正向题 Q27：若烹饪机具备低噪音低干扰功能时，您的感受是？",
        "22、反向题 Q28：若烹饪机不具备低噪音低干扰功能时，您的感受是？"
    ],
    "温暖治愈视觉外观": [
        "23、正向题 Q31：若烹饪机具备温暖治愈视觉外观时，您的感受是？",
        "24、反向题 Q32：若烹饪机不具备温暖治愈视觉外观时，您的感受是？"
    ],
    "操作区友好设计": [
        "25、正向题 Q33：若厨具具备操作区友好设计时，您的感受是？",
        "26、反向题 Q34：若厨具不具备操作区友好设计时，您的感受是？"
    ],
    "小巧温润不占地": [
        "27、正向题 Q35：若厨具具备小巧温润不占地的外观时，您的感受是？",
        "28、反向题 Q36：若厨具不具备小巧温润不占地的外观时，您的感受是？"
    ],
    "温暖低饱和色": [
        "29、正向题 Q37：若厨具采用温暖低饱和色的色彩方案时，您的感受是？",
        "30、反向题 Q38：若厨具不采用温暖低饱和色的色彩方案时，您的感受是？"
    ],
    "触感温润安全材料": [
        "31、正向题 Q39：若厨具采用触感温润安全材料的材质时，您的感受是？",
        "32、反向题 Q40：若厨具不采用触感温润安全材料的材质时，您的感受是？"
    ],
    "哑光细磨砂亲肤": [
        "33、正向题 Q41：若厨具采用哑光细磨砂柔和亲肤的表面处理时，您的感受是？",
        "34、反向题 Q42：若厨具不采用哑光细磨砂柔和亲肤的表面处理时，您的感受是？"
    ],
    "零学习成本交互": [
        "35、正向题 Q43：若厨具具备零学习成本的交互设计时，您的感受是？",
        "36、反向题 Q44：若厨具不具备零学习成本的交互设计时，您的感受是？"
    ],
    "安全状态强反馈": [
        "37、正向题 Q45：若厨具具备安全与状态强反馈的交互设计时，您的感受是？",
        "38、反向题 Q46：若厨具不具备安全与状态强反馈的交互设计时，您的感受是？"
    ],
    "情感化关怀交互": [
        "39、正向题 Q49：若厨具具备情感化关怀交互的设计时，您的感受是？",
        "40、反向题 Q50：若厨具不具备情感化关怀交互的设计时，您的感受是？"
    ]
}

SCORE_MAP = {
    "非常不满意": 1, "不满意": 2, "无所谓": 3, 
    "满意": 4, "非常满意": 5
}

KANO_RULES = {
    (4, 1): "必备属性M", (4, 2): "必备属性M",
    (5, 1): "必备属性M", (5, 2): "必备属性M",
    (1, 4): "魅力属性A", (1, 5): "魅力属性A",
    (2, 4): "魅力属性A", (2, 5): "魅力属性A",
    (3, 3): "无差异属性I",
    (4, 4): "无差异属性I", (5, 5): "无差异属性I",
    (1, 1): "反向属性R", (2, 2): "反向属性R",
    (4, 3): "期望属性O", (5, 3): "期望属性O",
    (3, 4): "期望属性O", (3, 5): "期望属性O"
}

# ==================== KANO分析函数（与原脚本一致）====================
def analyze_single_func(df, func_name, pos_col, neg_col):
    """单个功能的KANO计算"""
    data = df[[pos_col, neg_col]].copy()
    data.columns = ["positive", "negative"]
    
    # 映射为数字
    data["pos_score"] = data["positive"].map(SCORE_MAP)
    data["neg_score"] = data["negative"].map(SCORE_MAP)
    
    # 过滤无效值
    data = data.dropna(subset=["pos_score", "neg_score"])
    
    # 分类
    data["kano_type"] = data.apply(
        lambda row: KANO_RULES.get((int(row["pos_score"]), int(row["neg_score"])), "可疑结果Q"),
        axis=1
    )
    
    count = data["kano_type"].value_counts(normalize=True).to_dict()
    best_type = max(count, key=count.get) if count else "无"
    
    return {
        "功能": func_name,
        "样本数": len(data),
        "必备属性M": count.get("必备属性M", 0),
        "期望属性O": count.get("期望属性O", 0),
        "魅力属性A": count.get("魅力属性A", 0),
        "无差异属性I": count.get("无差异属性I", 0),
        "反向属性R": count.get("反向属性R", 0),
        "可疑结果Q": count.get("可疑结果Q", 0),
        "最终分类": best_type
    }

def calc_better_worse(row):
    M, O, A, I = row["必备属性M"], row["期望属性O"], row["魅力属性A"], row["无差异属性I"]
    total = M + O + A + I
    if total == 0:
        return 0, 0
    better = (A + O) / total
    worse = -(M + O) / total
    return better, worse

# ==================== UI界面 ====================
st.title("🍳 空巢青年智能烹饪厨具 KANO分析系统")
st.markdown("上传Excel问卷数据，自动完成KANO模型分类、Better-Worse系数计算及可视化")

# 侧边栏信息
with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
    1. **准备数据**：Excel需包含Q1-Q50列（与原脚本一致）
    2. **上传文件**：支持 .xlsx 格式
    3. **自动分析**：点击"开始分析"按钮
    4. **查看结果**：表格 + 图表 + 可下载的Excel
    """)
    st.divider()
    st.markdown("**预期列名示例：**")
    st.code("Q1: 当具备...时，您的感受是？\nQ2: 当不具备...时，您的感受是？", language=None)

# 文件上传
uploaded_file = st.file_uploader("📤 上传Excel数据文件", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        st.success(f"✅ 成功读取数据！共 {len(df)} 行，{len(df.columns)} 列")
        
        # 显示数据预览（可折叠）
        with st.expander("🔍 预览原始数据"):
            st.dataframe(df.head(10), use_container_width=True)
        
        # 开始分析按钮
        if st.button("🚀 开始KANO分析", type="primary", use_container_width=True):
            with st.spinner("正在进行KANO分类计算..."):
                
                # 批量分析
                result_list = []
                progress_bar = st.progress(0)
                
                for idx, (func, (pos_col, neg_col)) in enumerate(FUNC_MAP.items()):
                    try:
                        res = analyze_single_func(df, func, pos_col, neg_col)
                        result_list.append(res)
                    except Exception as e:
                        st.warning(f"⚠️ {func} 分析失败：{str(e)}")
                    progress_bar.progress((idx + 1) / len(FUNC_MAP))
                
                result_df = pd.DataFrame(result_list)
                
                # 计算Better-Worse系数
                result_df[["Better系数", "Worse系数"]] = result_df.apply(
                    lambda row: pd.Series(calc_better_worse(row)), axis=1
                )
                
                # 格式化百分比
                for col in ["必备属性M", "期望属性O", "魅力属性A", "无差异属性I", "反向属性R", "可疑结果Q"]:
                    result_df[col] = result_df[col].apply(lambda x: f"{x:.1%}")
                
                result_df["Better系数"] = result_df["Better系数"].apply(lambda x: f"{x:.3f}")
                result_df["Worse系数"] = result_df["Worse系数"].apply(lambda x: f"{x:.3f}")
                
                st.success("✅ 分析完成！")
                
                # ==================== 结果展示 ====================
                st.divider()
                st.header("📊 分析结果")
                
                # 1. 分类结果表格
                st.subheader("1. KANO需求分类表")
                
                # 高亮显示分类
                def highlight_kano(val):
                    colors = {
                        "必备属性M": "background-color: #ffcccc",
                        "期望属性O": "background-color: #ffe4cc", 
                        "魅力属性A": "background-color: #ccffcc",
                        "无差异属性I": "background-color: #ccccff",
                        "反向属性R": "background-color: #ffccff"
                    }
                    return colors.get(val, "")
                
                styled_df = result_df.style.applymap(highlight_kano, subset=["最终分类"])
                st.dataframe(styled_df, use_container_width=True, height=600)
                
                # 2. 统计图表
                st.subheader("2. 需求属性分布")
                type_count = result_df["最终分类"].value_counts()
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # 柱状图
                    fig1, ax1 = plt.subplots(figsize=(8, 5))
                    colors_map = {"必备属性M": "#d62728", "期望属性O": "#ff7f0e", 
                                  "魅力属性A": "#2ca02c", "无差异属性I": "#1f77b4", 
                                  "反向属性R": "#9467bd"}
                    type_count.plot(kind="bar", ax=ax1, color=[colors_map.get(x, "gray") for x in type_count.index])
                    ax1.set_title("智能厨具KANO需求属性分布", fontsize=14, pad=20)
                    ax1.set_xlabel("需求类型", fontsize=12)
                    ax1.set_ylabel("功能数量", fontsize=12)
                    ax1.tick_params(axis='x', rotation=0)
                    plt.tight_layout()
                    st.pyplot(fig1)
                
                with col2:
                    # 饼图
                    fig2, ax2 = plt.subplots(figsize=(8, 5))
                    ax2.pie(type_count.values, labels=type_count.index, autopct='%1.1f%%',
                           colors=[colors_map.get(x, "gray") for x in type_count.index])
                    ax2.set_title("需求属性占比", fontsize=14, pad=20)
                    st.pyplot(fig2)
                
                # 3. Better-Worse四象限图
                st.subheader("3. Better-Worse四象限分析")
                
                fig3, ax3 = plt.subplots(figsize=(12, 8))
                
                # 转换回数值用于绘图
                result_df["Better_num"] = result_df["Better系数"].astype(float)
                result_df["Worse_num"] = result_df["Worse系数"].astype(float)
                
                colors = {"必备属性M": "red", "期望属性O": "orange", 
                         "魅力属性A": "green", "无差异属性I": "blue", 
                         "反向属性R": "purple", "可疑结果Q": "gray"}
                
                for _, row in result_df.iterrows():
                    ax3.scatter(row["Worse_num"], row["Better_num"],
                               c=colors.get(row["最终分类"], "gray"), s=120, alpha=0.7, edgecolors='black')
                    ax3.annotate(row["功能"][:6], 
                                (row["Worse_num"], row["Better_num"]),
                                fontsize=9, ha='center', va='bottom')
                
                ax3.axhline(0, color="gray", linestyle="--", alpha=0.5)
                ax3.axvline(0, color="gray", linestyle="--", alpha=0.5)
                ax3.set_xlabel("Worse系数（影响不满）", fontsize=12)
                ax3.set_ylabel("Better系数（提升满意）", fontsize=12)
                ax3.set_title("KANO Better-Worse四象限图（右上为期望属性，左上为魅力属性，右下为必备属性）", fontsize=14)
                ax3.grid(alpha=0.3)
                
                # 添加象限标签
                ax3.text(0.5, 0.9, "期望属性\n高Better | 高Worse绝对值", transform=ax3.transAxes, 
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
                ax3.text(0.2, 0.9, "魅力属性\n高Better | 低Worse绝对值", transform=ax3.transAxes,
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
                ax3.text(0.5, 0.1, "必备属性\n低Better | 高Worse绝对值", transform=ax3.transAxes,
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))
                
                plt.tight_layout()
                st.pyplot(fig3)
                
                # ==================== 导出结果 ====================
                st.divider()
                st.header("💾 导出结果")
                
                # 准备导出数据（恢复数值格式）
                export_df = result_df.copy()
                for col in ["必备属性M", "期望属性O", "魅力属性A", "无差异属性I", "反向属性R", "可疑结果Q"]:
                    export_df[col] = export_df[col].str.replace('%', '').astype(float) / 100
                export_df["Better系数"] = export_df["Better系数"].astype(float)
                export_df["Worse系数"] = export_df["Worse系数"].astype(float)
                
                # 生成Excel文件
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    export_df.to_excel(writer, sheet_name="KANO分类结果", index=False)
                    type_count.to_excel(writer, sheet_name="属性统计")
                
                output.seek(0)
                
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        label="📥 下载完整分析结果 (Excel)",
                        data=output,
                        file_name="KANO分析结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col_dl2:
                    # 导出图表
                    img_buffer = BytesIO()
                    fig3.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
                    img_buffer.seek(0)
                    st.download_button(
                        label="📥 下载四象限图 (PNG)",
                        data=img_buffer,
                        file_name="KANO四象限图.png",
                        mime="image/png"
                    )
                
                # 关键洞察自动分析
                st.divider()
                st.header("💡 关键洞察")
                
                m_count = len(result_df[result_df["最终分类"] == "必备属性M"])
                o_count = len(result_df[result_df["最终分类"] == "期望属性O"])
                a_count = len(result_df[result_df["最终分类"] == "魅力属性A"])
                
                col_m, col_o, col_a = st.columns(3)
                col_m.metric("必备属性 (M) 数量", m_count, f"{m_count/len(result_df):.0%}")
                col_o.metric("期望属性 (O) 数量", o_count, f"{o_count/len(result_df):.0%}")
                col_a.metric("魅力属性 (A) 数量", a_count, f"{a_count/len(result_df):.0%}")
                
                st.info("""
                **优先级建议：**
                - **必备属性 (M)**：必须优先实现，缺失会导致用户极度不满
                - **期望属性 (O)**：投入产出比最高，越好用户越满意
                - **魅力属性 (A)**：差异化竞争优势，提供惊喜感
                """)
                
    except Exception as e:
        st.error(f"❌ 处理文件时出错：{str(e)}")
        st.info("请检查Excel格式是否正确，确保包含Sheet1且列名与问卷一致")

else:
    # 空状态提示
    st.info("👆 请上传Excel文件开始分析")
    
    # 提供模板下载（可选）
    st.divider()
    st.subheader("📄 数据格式参考")
    sample_data = {
        "列名示例": list(FUNC_MAP.values())[0] + ["...其他Q"],
        "说明": ["正向题", "反向题", "继续..."]
    }
    st.table(pd.DataFrame(sample_data))
