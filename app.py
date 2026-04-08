import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="空巢青年智能厨具 KANO分析系统",
    page_icon="🍳",
    layout="wide"
)

# 修复中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 1. 配置信息（适配旧版Q1-Q50格式）====================
SCORE_MAP = {
    "非常不满意": 1,
    "不满意": 2,
    "无所谓": 3,
    "满意": 4,
    "非常满意": 5
}

# 功能与列名映射（与第一版代码一致，匹配Q1-Q50）
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

# KANO分类规则（标准5级量表）
KANO_RULES = {
    (5, 1): "必备属性M", (5, 2): "必备属性M",
    (4, 1): "必备属性M", (4, 2): "必备属性M",
    (5, 3): "期望属性O", (4, 3): "期望属性O",
    (3, 4): "期望属性O", (3, 5): "期望属性O",
    (4, 4): "期望属性O", (5, 4): "期望属性O",
    (5, 5): "期望属性O",
    (1, 5): "魅力属性A", (1, 4): "魅力属性A",
    (2, 5): "魅力属性A", (2, 4): "魅力属性A",
    (1, 3): "魅力属性A", (2, 3): "魅力属性A",
    (3, 3): "无差异属性I",
    (3, 1): "无差异属性I", (3, 2): "无差异属性I",
    (1, 1): "反向属性R", (2, 2): "反向属性R",
    (1, 2): "反向属性R", (2, 1): "反向属性R"
}

# ==================== 2. 分析函数 ====================
def analyze_single_func(df, func_name, pos_col, neg_col):
    """单个功能的KANO计算"""
    if pos_col not in df.columns or neg_col not in df.columns:
        return None
    
    data = df[[pos_col, neg_col]].copy()
    data.columns = ["positive", "negative"]
    
    # 映射为数字
    data["pos_score"] = data["positive"].map(SCORE_MAP)
    data["neg_score"] = data["negative"].map(SCORE_MAP)
    
    # 过滤无效值
    data = data.dropna(subset=["pos_score", "neg_score"])
    
    if len(data) == 0:
        return None
    
    # 分类
    data["kano_type"] = data.apply(
        lambda row: KANO_RULES.get((int(row["pos_score"]), int(row["neg_score"])), "可疑结果Q"),
        axis=1
    )
    
    # 统计占比
    total = len(data)
    type_counts = data["kano_type"].value_counts()
    
    counts = {
        "必备属性M": type_counts.get("必备属性M", 0) / total,
        "期望属性O": type_counts.get("期望属性O", 0) / total,
        "魅力属性A": type_counts.get("魅力属性A", 0) / total,
        "无差异属性I": type_counts.get("无差异属性I", 0) / total,
        "反向属性R": type_counts.get("反向属性R", 0) / total,
        "可疑结果Q": type_counts.get("可疑结果Q", 0) / total
    }
    
    # 取占比最高的
    final_type = max(counts, key=counts.get)
    
    # Better-Worse系数（百分比形式，与问卷星一致）
    better = (counts["魅力属性A"] + counts["期望属性O"]) * 100
    worse = -(counts["必备属性M"] + counts["期望属性O"]) * 100
    
    return {
        "功能": func_name,
        "样本数": total,
        "最终分类": final_type,
        "Better系数": round(better, 2),
        "Worse系数": round(worse, 2),
        **{k: round(v*100, 2) for k, v in counts.items()}
    }

# ==================== 3. Streamlit界面 ====================
st.title("🍳 空巢青年智能厨具 KANO分析系统")
st.markdown("支持旧版问卷格式（Q1-Q50，量表：非常不满意→非常满意）")

with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
    **问卷格式要求：**
    - 包含列：1、正向题 Q1 ... / 2、反向题 Q2 ...
    - 评分选项：非常不满意、不满意、无所谓、满意、非常满意
    - 共20个功能，40个题项（Q1-Q50，部分编号跳过）
    
    **评分映射：**
    - 非常不满意 = 1
    - 不满意 = 2
    - 无所谓 = 3
    - 满意 = 4
    - 非常满意 = 5
    """)

# 文件上传（关键：使用内存读取，非本地路径）
uploaded_file = st.file_uploader("📤 上传Excel问卷数据", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 读取Excel（自动识别Sheet1）
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ 成功读取数据！共 {len(df)} 行")
        
        # 检查列名
        missing_cols = []
        for func, (pos, neg) in FUNC_MAP.items():
            if pos not in df.columns:
                missing_cols.append(pos)
            if neg not in df.columns:
                missing_cols.append(neg)
        
        if missing_cols:
            st.error(f"❌ 缺少 {len(missing_cols)} 个必要列，请检查Excel格式")
            with st.expander("查看缺失的列"):
                for col in missing_cols[:5]:
                    st.code(col)
            st.stop()
        
        # 数据预览
        with st.expander("🔍 预览数据（前5行）"):
            st.dataframe(df.head(), use_container_width=True)
        
        # 分析按钮
        if st.button("🚀 开始KANO分析", type="primary", use_container_width=True):
            with st.spinner("分析中..."):
                result_list = []
                progress_bar = st.progress(0)
                
                for idx, (func_name, (pos_col, neg_col)) in enumerate(FUNC_MAP.items()):
                    res = analyze_single_func(df, func_name, pos_col, neg_col)
                    if res:
                        result_list.append(res)
                    progress_bar.progress((idx + 1) / len(FUNC_MAP))
                
                if not result_list:
                    st.error("分析失败，请检查数据")
                    st.stop()
                
                result_df = pd.DataFrame(result_list)
                
                # 按分类排序
                type_order = {"必备属性M": 0, "期望属性O": 1, "魅力属性A": 2, 
                             "无差异属性I": 3, "反向属性R": 4, "可疑结果Q": 5}
                result_df["排序"] = result_df["最终分类"].map(type_order)
                result_df = result_df.sort_values("排序").drop("排序", axis=1).reset_index(drop=True)
                
                # ==================== 结果展示 ====================
                st.divider()
                st.header("📊 分析结果")
                
                # 1. 主表格
                st.subheader("KANO需求分类表")
                
                def highlight_kano(val):
                    colors = {
                        "必备属性M": "background-color: #ffcccc",
                        "期望属性O": "background-color: #ffe4cc",
                        "魅力属性A": "background-color: #ccffcc",
                        "无差异属性I": "background-color: #ccccff",
                        "反向属性R": "background-color: #ffccff"
                    }
                    return colors.get(val, "")
                
                display_df = result_df[["功能", "最终分类", "Better系数", "Worse系数", 
                                       "必备属性M", "期望属性O", "魅力属性A", "无差异属性I"]].copy()
                
                # 添加百分号显示
                for col in ["Better系数", "Worse系数", "必备属性M", "期望属性O", "魅力属性A", "无差异属性I"]:
                    display_df[col] = display_df[col].astype(str) + "%"
                
                try:
                    styled_df = display_df.style.map(highlight_kano, subset=["最终分类"])
                except:
                    styled_df = display_df
                
                st.dataframe(styled_df, use_container_width=True, height=700)
                
                # 2. 可视化
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("属性分布")
                    type_count = result_df["最终分类"].value_counts()
                    
                    fig1, ax1 = plt.subplots(figsize=(8, 5))
                    colors_map = {"必备属性M": "#d62728", "期望属性O": "#ff7f0e", 
                                  "魅力属性A": "#2ca02c", "无差异属性I": "#1f77b4", 
                                  "反向属性R": "#9467bd"}
                    
                    bars = ax1.bar(type_count.index, type_count.values, 
                                   color=[colors_map.get(x, "gray") for x in type_count.index])
                    ax1.set_ylabel("功能数量")
                    
                    for bar in bars:
                        height = bar.get_height()
                        ax1.text(bar.get_x() + bar.get_width()/2., height,
                                f'{int(height)}', ha='center', va='bottom')
                    
                    st.pyplot(fig1)
                
                with col2:
                    st.subheader("占比饼图")
                    fig2, ax2 = plt.subplots(figsize=(8, 5))
                    ax2.pie(type_count.values, labels=type_count.index, autopct='%1.1f%%',
                           colors=[colors_map.get(x, "gray") for x in type_count.index])
                    st.pyplot(fig2)
                
                # 3. 四象限图
                st.subheader("Better-Worse四象限分析")
                
                fig3, ax3 = plt.subplots(figsize=(12, 8))
                
                for _, row in result_df.iterrows():
                    color = colors_map.get(row["最终分类"], "gray")
                    ax3.scatter(row["Worse系数"], row["Better系数"],
                               c=color, s=150, alpha=0.7, edgecolors='black')
                    ax3.annotate(row["功能"][:6], 
                                (row["Worse系数"], row["Better系数"]),
                                xytext=(5, 5), textcoords='offset points', fontsize=9)
                
                # 参考线
                ax3.axhline(y=50, color='gray', linestyle='--', alpha=0.3)
                ax3.axvline(x=-50, color='gray', linestyle='--', alpha=0.3)
                ax3.axhline(y=0, color='black', linestyle='-', alpha=0.2)
                ax3.axvline(x=0, color='black', linestyle='-', alpha=0.2)
                
                ax3.set_xlabel("Worse系数")
                ax3.set_ylabel("Better系数")
                ax3.set_title("KANO四象限图")
                ax3.grid(alpha=0.3)
                
                # 象限标签
                ax3.text(0.75, 0.95, "期望属性", transform=ax3.transAxes, 
                        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
                ax3.text(0.05, 0.95, "魅力属性", transform=ax3.transAxes,
                        bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
                ax3.text(0.75, 0.05, "必备属性", transform=ax3.transAxes,
                        bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))
                
                st.pyplot(fig3)
                
                # ==================== 导出 ====================
                st.divider()
                st.header("💾 导出结果")
                
                # 准备Excel
                export_df = result_df.copy()
                for col in ["Better系数", "Worse系数", "必备属性M", "期望属性O", 
                           "魅力属性A", "无差异属性I", "反向属性R", "可疑结果Q"]:
                    export_df[col] = export_df[col].astype(str) + "%"
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    export_df.to_excel(writer, sheet_name="KANO结果", index=False)
                
                st.download_button(
                    label="📥 下载Excel报告",
                    data=output.getvalue(),
                    file_name="KANO分析结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 统计卡片
                st.divider()
                cols = st.columns(4)
                for i, (t, c) in enumerate(result_df["最终分类"].value_counts().items()):
                    if i < 4:
                        emoji = {"必备属性M": "🔴", "期望属性O": "🟠", 
                                "魅力属性A": "🟢", "无差异属性I": "🔵"}.get(t, "⚪")
                        cols[i].metric(f"{emoji} {t}", f"{c}个")

    except Exception as e:
        st.error(f"❌ 错误：{str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👆 请上传Excel文件（包含Q1-Q50列的旧版问卷）")
