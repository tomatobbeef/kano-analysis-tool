import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="空巢青年智能厨具 KANO分析系统",
    page_icon="🍳",
    layout="wide"
)

# 修复中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 1. 配置信息（适配新问卷）====================
# 评分映射：新问卷使用"很喜欢/理所当然/无所谓/勉强接受/很不喜欢"
SCORE_MAP = {
    "很喜欢": 5,
    "理所当然": 4,
    "无所谓": 3,
    "勉强接受": 2,
    "很不喜欢": 1
}

# 功能名称与列名映射（20个功能，每个包含正向和反向题）
FUNC_MAP = {
    "一键菜谱同步": [
        "1、智能厨具具备一键菜谱同步、自动匹配烹饪步骤功能—如果有这个功能您的评价是：",
        "1、如果没有这个功能您的评价是："
    ],
    "APP安全监护预警": [
        "2、智能厨具具备APP 安全监护预警（干烧、漏电、溢锅提醒）功能—如果有这个功能您的评价是：",
        "2、如果没有这个功能您的评价是："
    ],
    "定时预约自动完成": [
        "3、智能厨具具备定时预约、烹饪自动完成功能—如果有这个功能您的评价是：",
        "3、如果没有这个功能您的评价是："
    ],
    "健康饮食关怀提醒": [
        "4、智能厨具具备健康饮食关怀与提醒（营养摄入、饮食规律提醒）功能—如果有这个功能您的评价是：",
        "4、如果没有这个功能您的评价是："
    ],
    "APP轻量化操作": [
        "5、智能厨具具备APP 轻量化操作、界面简洁易上手功能—如果有这个功能您的评价是：",
        "5、如果没有这个功能您的评价是："
    ],
    "强效防油烟": [
        "6、智能厨具具备强效防油烟、减少厨房油污功能—如果有这个功能您的评价是：",
        "6、如果没有这个功能您的评价是："
    ],
    "烹饪后自动保温": [
        "7、智能厨具具备烹饪后自动保温、随时吃上热饭功能—如果有这个功能您的评价是：",
        "7、如果没有这个功能您的评价是："
    ],
    "一人食小容量": [
        "8、智能厨具具备一人食专属小容量、不浪费食材设计—如果有这个功能您的评价是：",
        "8、如果没有这个功能您的评价是："
    ],
    "多重安全防护": [
        "9、智能厨具具备多重安全防护、杜绝使用隐患功能—如果有这个功能您的评价是：",
        "9、如果没有这个功能您的评价是："
    ],
    "易拆卸清洗": [
        "10、智能厨具具备机身 / 配件易拆卸、清洗无死角功能—如果有这个功能您的评价是：",
        "10、如果没有这个功能您的评价是："
    ],
    "低噪音运行": [
        "11、智能厨具具备低噪音运行、不产生噪音干扰功能—如果有这个功能您的评价是：",
        "11、如果没有这个功能您的评价是："
    ],
    "温暖治愈外观": [
        "12、智能厨具采用温暖治愈、贴合独居氛围的外观设计—如果有这个功能您的评价是：",
        "12、如果没有这个功能您的评价是："
    ],
    "操作区友好设计": [
        "13、智能厨具采用按键 / 触屏操作区清晰、好操作设计—如果有这个功能您的评价是：",
        "13、如果没有这个功能您的评价是："
    ],
    "小巧不占地": [
        "14、智能厨具采用小巧机身、不占厨房空间设计—如果有这个功能您的评价是：",
        "14、如果没有这个功能您的评价是："
    ],
    "温暖低饱和色": [
        "15、智能厨具采用温暖低饱和色系、视觉舒适设计—如果有这个功能您的评价是：",
        "15、如果没有这个功能您的评价是："
    ],
    "温润安全材质": [
        "16、智能厨具采用温润安全、防烫防滑材质—如果有这个功能您的评价是：",
        "16、如果没有这个功能您的评价是："
    ],
    "触感温润表面处理": [
        "17、智能厨具采用触感温润安全表面处理—如果有这个功能您的评价是：",
        "17、如果没有这个功能您的评价是："
    ],
    "零学习成本交互": [
        "18、智能厨具具备零学习成本、拿到就会用的交互设计—如果有这个功能您的评价是：",
        "18、如果没有这个功能您的评价是："
    ],
    "安全状态强反馈": [
        "19、智能厨具具备安全状态实时强反馈（灯光 / 声音提醒）功能—如果有这个功能您的评价是：",
        "19、如果没有这个功能您的评价是："
    ],
    "情感化关怀陪伴": [
        "20、智能厨具具备情感化关怀，增强陪伴感—如果有这个功能您的评价是：",
        "20、如果没有这个功能您的评价是："
    ]
}

# ==================== 2. KANO分类函数（与问卷星逻辑一致）====================
def classify_kano(pos_score, neg_score):
    """
    KANO分类逻辑（基于标准KANO模型）：
    pos_score: 正向题得分（有功能时的评价）1-5
    neg_score: 反向题得分（无功能时的评价）1-5
    
    注意：反向题中，1=很不喜欢（没有功能我很不喜欢=我需要这个功能），5=很喜欢（没有功能我很喜欢=我不需要）
    """
    p = int(pos_score)
    n = int(neg_score)
    
    # 必备属性 M：有则满意(4/5)，无则不满(1/2)
    if (p in [4, 5]) and (n in [1, 2]):
        return "必备属性M"
    
    # 期望属性 O：有则满意，无则勉强接受/也满意（线性关系，不够惊喜）
    if (p in [4, 5]) and (n in [4, 5]):
        return "期望属性O"
    if (p == 5) and (n == 4):
        return "期望属性O"
    
    # 魅力属性 A：有则很喜欢(5)，无则无所谓(3) - 惊喜型
    if (p == 5) and (n == 3):
        return "魅力属性A"
    if (p == 4) and (n == 3):
        return "魅力属性A"
    
    # 无差异属性 I：有无都无所谓，或态度一致
    if (p == 3) and (n in [2, 3, 4]):
        return "无差异属性I"
    if (p == 4) and (n == 4):  # 有则理所当然，无则理所当然 → 其实是期望？但通常归为无差异或期望
        return "无差异属性I"
    
    # 反向属性 R：有则不满(1/2)，无则喜(4/5)；或有无都喜（不需要这个功能）
    if (p in [1, 2]) and (n in [4, 5]):
        return "反向属性R"
    if (p in [1, 2]) and (n in [1, 2, 3]):  # 有则不满，无则无所谓/不满
        return "反向属性R"
    if (p in [3, 4, 5]) and (n == 5):  # 无则很喜欢（没有更好）
        return "反向属性R"
    
    # 可疑结果 Q：矛盾回答（如很喜欢有且很喜欢无）
    if (p == 5) and (n == 5):
        return "可疑结果Q"
    if (p == 1) and (n == 1):
        return "可疑结果Q"
    
    # 默认归类
    return "无差异属性I"

def analyze_single_func(df, func_name, pos_col, neg_col):
    """单个功能的KANO分析"""
    # 提取数据
    data = df[[pos_col, neg_col]].copy()
    data.columns = ["pos", "neg"]
    
    # 映射为数字
    data["pos_num"] = data["pos"].map(SCORE_MAP)
    data["neg_num"] = data["neg"].map(SCORE_MAP)
    
    # 删除无效值
    data = data.dropna(subset=["pos_num", "neg_num"])
    
    if len(data) == 0:
        return None
    
    # 逐样本分类
    data["kano_type"] = data.apply(lambda row: classify_kano(row["pos_num"], row["neg_num"]), axis=1)
    
    # 统计占比
    type_counts = data["kano_type"].value_counts(normalize=True)
    counts_dict = {
        "必备属性M": type_counts.get("必备属性M", 0),
        "期望属性O": type_counts.get("期望属性O", 0),
        "魅力属性A": type_counts.get("魅力属性A", 0),
        "无差异属性I": type_counts.get("无差异属性I", 0),
        "反向属性R": type_counts.get("反向属性R", 0),
        "可疑结果Q": type_counts.get("可疑结果Q", 0)
    }
    
    # 取占比最高的为最终分类
    final_type = max(counts_dict, key=counts_dict.get)
    
    # Better-Worse系数计算（与问卷星一致的百分比形式）
    total = sum(counts_dict.values())
    if total > 0:
        better = (counts_dict["魅力属性A"] + counts_dict["期望属性O"]) / total * 100
        worse = -(counts_dict["必备属性M"] + counts_dict["期望属性O"]) / total * 100
    else:
        better = 0
        worse = 0
    
    return {
        "功能": func_name,
        "样本数": len(data),
        "最终分类": final_type,
        "Better系数": round(better, 2),
        "Worse系数": round(worse, 2),
        **{k: round(v*100, 2) for k, v in counts_dict.items()}  # 转为百分比
    }

# ==================== 3. Streamlit界面 ====================
st.title("🍳 空巢青年智能烹饪厨具 KANO分析系统")
st.markdown("上传Excel问卷数据，自动完成KANO模型分类与Better-Worse系数计算")

with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
    1. **上传文件**：选择你的Excel问卷数据（.xlsx格式）
    2. **自动分析**：系统自动识别20个功能的正反项
    3. **查看结果**：表格展示分类、Better-Worse系数
    4. **下载报告**：支持导出Excel和图表
    
    **评分标准映射：**
    - 很喜欢 = 5分
    - 理所当然 = 4分  
    - 无所谓 = 3分
    - 勉强接受 = 2分
    - 很不喜欢 = 1分
    """)
    
    st.divider()
    st.markdown("**技术说明**")
    st.info("若分析结果与问卷星有差异，请检查是否有空值或异常回答")

# 文件上传（关键修复：使用uploaded_file而非本地路径）
uploaded_file = st.file_uploader("📤 上传Excel数据文件", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 读取上传的文件（从内存中读取，不是本地路径）
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        st.success(f"✅ 成功读取数据！共 {len(df)} 行，{len(df.columns)} 列")
        
        # 检查必要的列是否存在
        missing_cols = []
        for func, (pos, neg) in FUNC_MAP.items():
            if pos not in df.columns:
                missing_cols.append(pos)
            if neg not in df.columns:
                missing_cols.append(neg)
        
        if missing_cols:
            st.error(f"❌ 缺少以下 {len(missing_cols)} 个必要列：")
            for col in missing_cols[:3]:
                st.code(col, language=None)
            if len(missing_cols) > 3:
                st.info(f"...还有 {len(missing_cols)-3} 个列未显示")
            st.stop()
        
        # 数据预览
        with st.expander("🔍 预览原始数据（前5行）"):
            st.dataframe(df.head(), use_container_width=True)
        
        # 开始分析按钮
        if st.button("🚀 开始KANO分析", type="primary", use_container_width=True):
            with st.spinner("正在进行KANO分类计算..."):
                # 批量分析所有功能
                result_list = []
                progress_bar = st.progress(0)
                
                for idx, (func_name, (pos_col, neg_col)) in enumerate(FUNC_MAP.items()):
                    result = analyze_single_func(df, func_name, pos_col, neg_col)
                    if result:
                        result_list.append(result)
                    progress_bar.progress((idx + 1) / len(FUNC_MAP))
                
                if not result_list:
                    st.error("没有成功分析任何功能，请检查数据格式")
                    st.stop()
                
                # 转为DataFrame
                result_df = pd.DataFrame(result_list)
                
                # 按分类排序（必备>期望>魅力>无差异>反向）
                type_order = {"必备属性M": 0, "期望属性O": 1, "魅力属性A": 2, 
                             "无差异属性I": 3, "反向属性R": 4, "可疑结果Q": 5}
                result_df["排序"] = result_df["最终分类"].map(type_order)
                result_df = result_df.sort_values("排序").drop("排序", axis=1).reset_index(drop=True)
                
                st.success("✅ 分析完成！")
                
                # ==================== 结果展示 ====================
                st.divider()
                st.header("📊 分析结果")
                
                # 1. 主要结果表格
                st.subheader("1. KANO需求分类表")
                
                # 高亮显示
                def highlight_kano(val):
                    colors = {
                        "必备属性M": "background-color: #ffcccc; color: #000",
                        "期望属性O": "background-color: #ffe4cc; color: #000",
                        "魅力属性A": "background-color: #ccffcc; color: #000",
                        "无差异属性I": "background-color: #ccccff; color: #000",
                        "反向属性R": "background-color: #ffccff; color: #000"
                    }
                    return colors.get(val, "")
                
                # 显示主要列
                display_cols = ["功能", "最终分类", "Better系数", "Worse系数", "必备属性M", "期望属性O", "魅力属性A", "无差异属性I"]
                display_df = result_df[display_cols].copy()
                
                # 添加百分号
                for col in ["Better系数", "Worse系数"]:
                    display_df[col] = display_df[col].astype(str) + "%"
                
                try:
                    styled_df = display_df.style.map(highlight_kano, subset=["最终分类"])
                except:
                    styled_df = display_df
                
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
                                  "反向属性R": "#9467bd", "可疑结果Q": "#7f7f7f"}
                    bars = ax1.bar(type_count.index, type_count.values, 
                                   color=[colors_map.get(x, "gray") for x in type_count.index])
                    ax1.set_title("KANO需求属性分布", fontsize=14)
                    ax1.set_ylabel("功能数量")
                    ax1.tick_params(axis='x', rotation=45)
                    
                    for bar in bars:
                        height = bar.get_height()
                        ax1.text(bar.get_x() + bar.get_width()/2., height,
                                f'{int(height)}', ha='center', va='bottom')
                    
                    plt.tight_layout()
                    st.pyplot(fig1)
                
                with col2:
                    # 饼图
                    fig2, ax2 = plt.subplots(figsize=(8, 5))
                    ax2.pie(type_count.values, labels=type_count.index, autopct='%1.1f%%',
                           colors=[colors_map.get(x, "gray") for x in type_count.index])
                    ax2.set_title("需求属性占比")
                    st.pyplot(fig2)
                
                # 3. Better-Worse四象限图
                st.subheader("3. Better-Worse四象限分析")
                
                fig3, ax3 = plt.subplots(figsize=(12, 8))
                
                for _, row in result_df.iterrows():
                    color = colors_map.get(row["最终分类"], "gray")
                    ax3.scatter(row["Worse系数"], row["Better系数"],
                               c=color, s=150, alpha=0.7, edgecolors='black', linewidth=1)
                    
                    # 标注功能名称（缩短）
                    label = row["功能"][:8]
                    ax3.annotate(label, (row["Worse系数"], row["Better系数"]),
                                xytext=(5, 5), textcoords='offset points', fontsize=9)
                
                # 参考线
                ax3.axhline(y=50, color='gray', linestyle='--', alpha=0.3)
                ax3.axvline(x=-50, color='gray', linestyle='--', alpha=0.3)
                ax3.axhline(y=0, color='black', linestyle='-', alpha=0.2)
                ax3.axvline(x=0, color='black', linestyle='-', alpha=0.2)
                
                ax3.set_xlabel("Worse系数（影响不满程度）", fontsize=12)
                ax3.set_ylabel("Better系数（提升满意程度）", fontsize=12)
                ax3.set_title("Better-Worse四象限图（与问卷星一致）", fontsize=14)
                ax3.grid(alpha=0.3)
                
                # 象限标签
                ax3.text(0.75, 0.95, "期望属性\n(高Better/高Worse)", transform=ax3.transAxes, 
                        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
                ax3.text(0.05, 0.95, "魅力属性\n(高Better/低Worse)", transform=ax3.transAxes,
                        bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
                ax3.text(0.75, 0.05, "必备属性\n(低Better/高Worse)", transform=ax3.transAxes,
                        bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))
                ax3.text(0.05, 0.05, "无差异\n(低Better/低Worse)", transform=ax3.transAxes,
                        bbox=dict(boxstyle='round', facecolor='lightgray', alpha=0.5))
                
                plt.tight_layout()
                st.pyplot(fig3)
                
                # ==================== 导出结果 ====================
                st.divider()
                st.header("💾 导出结果")
                
                # 准备Excel数据
                export_df = result_df.copy()
                for col in ["Better系数", "Worse系数", "必备属性M", "期望属性O", "魅力属性A", "无差异属性I", "反向属性R", "可疑结果Q"]:
                    if col in export_df.columns:
                        export_df[col] = export_df[col].astype(str) + "%"
                
                # 生成Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    export_df.to_excel(writer, sheet_name="KANO分析结果", index=False)
                    type_count.to_excel(writer, sheet_name="属性统计")
                output.seek(0)
                
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        label="📥 下载分析结果(Excel)",
                        data=output,
                        file_name="KANO分析结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col_dl2:
                    # 下载四象限图
                    img_buffer = BytesIO()
                    fig3.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
                    img_buffer.seek(0)
                    st.download_button(
                        label="📥 下载四象限图(PNG)",
                        data=img_buffer,
                        file_name="KANO四象限图.png",
                        mime="image/png"
                    )
                
                # 关键洞察
                st.divider()
                st.header("💡 关键洞察")
                
                m_count = len(result_df[result_df["最终分类"] == "必备属性M"])
                o_count = len(result_df[result_df["最终分类"] == "期望属性O"])
                a_count = len(result_df[result_df["最终分类"] == "魅力属性A"])
                i_count = len(result_df[result_df["最终分类"] == "无差异属性I"])
                
                cols = st.columns(4)
                cols[0].metric("🔴 必备属性(M)", m_count)
                cols[1].metric("🟠 期望属性(O)", o_count)
                cols[2].metric("🟢 魅力属性(A)", a_count)
                cols[3].metric("🔵 无差异(I)", i_count)
                
                # 显示各类别的具体功能
                col_list1, col_list2 = st.columns(2)
                
                with col_list1:
                    if m_count > 0:
                        st.subheader("🔴 必备属性（必须实现）")
                        for func in result_df[result_df["最终分类"] == "必备属性M"]["功能"].tolist():
                            st.write(f"- {func}")
                    
                    if o_count > 0:
                        st.subheader("🟠 期望属性（投入产出比高）")
                        for func in result_df[result_df["最终分类"] == "期望属性O"]["功能"].tolist():
                            st.write(f"- {func}")
                
                with col_list2:
                    if a_count > 0:
                        st.subheader("🟢 魅力属性（差异化卖点）")
                        for func in result_df[result_df["最终分类"] == "魅力属性A"]["功能"].tolist():
                            st.write(f"- {func}")
                    
                    if i_count > 0:
                        st.subheader("🔵 无差异属性（可暂缓）")
                        for func in result_df[result_df["最终分类"] == "无差异属性I"]["功能"].tolist():
                            st.write(f"- {func}")

    except Exception as e:
        st.error(f"❌ 处理文件时出错：{str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    # 空状态
    st.info("👆 请上传Excel文件开始分析")
    
    # 显示格式示例
    st.divider()
    st.subheader("📄  expected columns (前3个功能示例)")
    example_data = []
    for i, (func, (pos, neg)) in enumerate(list(FUNC_MAP.items())[:3], 1):
        example_data.append({
            "功能序号": i,
            "功能名称": func,
            "正向题列名片段": pos[:30] + "...",
            "反向题列名片段": neg[:30] + "..."
        })
    st.table(pd.DataFrame(example_data))
