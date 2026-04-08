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

# 修复Matplotlib中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 1. 配置信息（已适配新问卷）====================
# 新问卷的评分映射
SCORE_MAP = {
    "很喜欢": 5,
    "理所当然": 4, 
    "无所谓": 3,
    "勉强接受": 2,
    "很不喜欢": 1
}

# 功能名称与 正向/反向 列名映射（严格对应新Excel的列名）
FUNC_MAP = {
    "省心一键菜谱": [
        "1、智能厨具具备一键菜谱同步、自动匹配烹饪步骤功能—如果有这个功能您的评价是：",
        "1、如果没有这个功能您的评价是："
    ],
    "安全监护与异常预警": [
        "2、智能厨具具备APP 安全监护预警（干烧、漏电、溢锅提醒）功能—如果有这个功能您的评价是：",
        "2、如果没有这个功能您的评价是："
    ],
    "定时预约自动完成": [
        "3、智能厨具具备定时预约、烹饪自动完成功能—如果有这个功能您的评价是：",
        "3、如果没有这个功能您的评价是："
    ],
    "健康关怀与提醒": [
        "4、智能厨具具备健康饮食关怀与提醒（营养摄入、饮食规律提醒）功能—如果有这个功能您的评价是：",
        "4、如果没有这个功能您的评价是："
    ],
    "轻量化使用": [
        "5、智能厨具具备APP 轻量化操作、界面简洁易上手功能—如果有这个功能您的评价是：",
        "5、如果没有这个功能您的评价是："
    ],
    "防油烟": [
        "6、智能厨具具备强效防油烟、减少厨房油污功能—如果有这个功能您的评价是：",
        "6、如果没有这个功能您的评价是："
    ],
    "自动保温": [
        "7、智能厨具具备烹饪后自动保温、随时吃上热饭功能—如果有这个功能您的评价是：",
        "7、如果没有这个功能您的评价是："
    ],
    "一人食小容量": [
        "8、智能厨具具备一人食专属小容量、不浪费食材设计—如果有这个功能您的评价是：",
        "8、如果没有这个功能您的评价是："
    ],
    "极致安全": [
        "9、智能厨具具备多重安全防护、杜绝使用隐患功能—如果有这个功能您的评价是：",
        "9、如果没有这个功能您的评价是："
    ],
    "易清洁少维护": [
        "10、智能厨具具备机身 / 配件易拆卸、清洗无死角功能—如果有这个功能您的评价是：",
        "10、如果没有这个功能您的评价是："
    ],
    "低噪音低干扰": [
        "11、智能厨具具备低噪音运行、不产生噪音干扰功能—如果有这个功能您的评价是：",
        "11、如果没有这个功能您的评价是："
    ],
    "温暖治愈视觉外观": [
        "12、智能厨具采用温暖治愈、贴合独居氛围的外观设计—如果有这个功能您的评价是：",
        "12、如果没有这个功能您的评价是："
    ],
    "操作区友好设计": [
        "13、智能厨具采用按键 / 触屏操作区清晰、好操作设计—如果有这个功能您的评价是：",
        "13、如果没有这个功能您的评价是："
    ],
    "小巧温润不占地": [
        "14、智能厨具采用小巧机身、不占厨房空间设计—如果有这个功能您的评价是：",
        "14、如果没有这个功能您的评价是："
    ],
    "温暖低饱和色": [
        "15、智能厨具采用温暖低饱和色系、视觉舒适设计—如果有这个功能您的评价是：",
        "15、如果没有这个功能您的评价是："
    ],
    "触感温润安全材料": [
        "16、智能厨具采用温润安全、防烫防滑材质—如果有这个功能您的评价是：",
        "16、如果没有这个功能您的评价是："
    ],
    "哑光细磨砂亲肤": [
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
    "情感化关怀交互": [
        "20、智能厨具具备情感化关怀，增强陪伴感—如果有这个功能您的评价是：",
        "20、如果没有这个功能您的评价是："
    ]
}

# KANO分类规则（保持不变）
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

# ==================== KANO分析函数 ====================
def analyze_single_func(df, func_name, pos_col, neg_col):
    """单个功能的KANO计算"""
    data = df[[pos_col, neg_col]].copy()
    data.columns = ["positive", "negative"]
    
    # 映射为数字
    data["pos_score"] = data["positive"].map(SCORE_MAP)
    data["neg_score"] = data["negative"].map(SCORE_MAP)
    
    # 过滤无效值
    data = data.dropna(subset=["pos_score", "neg_score"])
    
    if len(data) == 0:
        return {
            "功能": func_name,
            "样本数": 0,
            "必备属性M": 0,
            "期望属性O": 0,
            "魅力属性A": 0,
            "无差异属性I": 0,
            "反向属性R": 0,
            "可疑结果Q": 0,
            "最终分类": "无数据"
        }
    
    # 分类
    data["kano_type"] = data.apply(
        lambda row: KANO_RULES.get((int(row["pos_score"]), int(row["neg_score"])), "可疑结果Q"),
        axis=1
    )
    
    # 统计占比
    counts = data["kano_type"].value_counts()
    total = len(data)
    
    result = {
        "功能": func_name,
        "样本数": total,
        "必备属性M": counts.get("必备属性M", 0) / total,
        "期望属性O": counts.get("期望属性O", 0) / total,
        "魅力属性A": counts.get("魅力属性A", 0) / total,
        "无差异属性I": counts.get("无差异属性I", 0) / total,
        "反向属性R": counts.get("反向属性R", 0) / total,
        "可疑结果Q": counts.get("可疑结果Q", 0) / total,
    }
    
    # 确定最终分类（占比最高的）
    best_type = max(result, key=lambda k: result[k] if k != "样本数" and k != "功能" else -1)
    result["最终分类"] = best_type
    return result

def calc_better_worse(row):
    M, O, A, I = row["必备属性M"], row["期望属性O"], row["魅力属性A"], row["无差异属性I"]
    total = M + O + A + I
    if total == 0:
        return 0, 0
    better = (A + O) / total  # 提升满意度
    worse = -(M + O) / total  # 降低满意度
    return better, worse

# ==================== UI界面 ====================
st.title("🍳 空巢青年智能烹饪厨具 KANO分析系统")
st.markdown("上传Excel问卷数据，自动完成KANO模型分类、Better-Worse系数计算及可视化")

with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
    1. **准备数据**：上传含20个功能正向/反向题的Excel
    2. **评分标准**：很喜欢(5) → 理所当然(4) → 无所谓(3) → 勉强接受(2) → 很不喜欢(1)
    3. **自动分析**：点击"开始分析"按钮
    4. **查看结果**：表格 + 图表 + 可下载的Excel
    """)

# 文件上传
uploaded_file = st.file_uploader("📤 上传Excel数据文件", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # 读取数据（新问卷默认在Sheet1）
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
            st.error(f"❌ 缺少以下列（共{len(missing_cols)}个）：")
            for col in missing_cols[:5]:  # 只显示前5个避免太长
                st.code(col)
            if len(missing_cols) > 5:
                st.info(f"...还有 {len(missing_cols)-5} 个列未显示")
            st.stop()
        
        # 显示数据预览
        with st.expander("🔍 预览原始数据（前5行）"):
            st.dataframe(df.head(), use_container_width=True)
        
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
                        st.error(f"❌ {func} 分析失败：{str(e)}")
                    progress_bar.progress((idx + 1) / len(FUNC_MAP))
                
                if not result_list:
                    st.error("没有成功分析任何功能，请检查数据格式")
                    st.stop()
                
                result_df = pd.DataFrame(result_list)
                
                # 计算Better-Worse系数
                result_df[["Better系数", "Worse系数"]] = result_df.apply(
                    lambda row: pd.Series(calc_better_worse(row)), axis=1
                )
                
                # 格式化百分比显示
                display_df = result_df.copy()
                for col in ["必备属性M", "期望属性O", "魅力属性A", "无差异属性I", "反向属性R", "可疑结果Q"]:
                    display_df[col] = display_df[col].apply(lambda x: f"{x:.1%}")
                
                display_df["Better系数"] = display_df["Better系数"].apply(lambda x: f"{x:.3f}")
                display_df["Worse系数"] = display_df["Worse系数"].apply(lambda x: f"{x:.3f}")
                
                st.success("✅ 分析完成！")
                
                # ==================== 结果展示 ====================
                st.divider()
                st.header("📊 分析结果")
                
                # 1. 分类结果表格
                st.subheader("1. KANO需求分类表")
                
                # 高亮显示分类
                def highlight_kano(val):
                    colors = {
                        "必备属性M": "background-color: #ffcccc; color: #000",
                        "期望属性O": "background-color: #ffe4cc; color: #000", 
                        "魅力属性A": "background-color: #ccffcc; color: #000",
                        "无差异属性I": "background-color: #ccccff; color: #000",
                        "反向属性R": "background-color: #ffccff; color: #000"
                    }
                    return colors.get(val, "")
                
                # 兼容新旧版本Pandas
                try:
                    styled_df = display_df.style.map(highlight_kano, subset=["最终分类"])
                except AttributeError:
                    styled_df = display_df.style.applymap(highlight_kano, subset=["最终分类"])
                
                st.dataframe(styled_df, use_container_width=True, height=700)
                
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
                    ax1.set_title("智能厨具KANO需求属性分布", fontsize=14, pad=20)
                    ax1.set_xlabel("需求类型", fontsize=12)
                    ax1.set_ylabel("功能数量", fontsize=12)
                    ax1.tick_params(axis='x', rotation=45)
                    
                    # 在柱子上显示数值
                    for bar in bars:
                        height = bar.get_height()
                        ax1.text(bar.get_x() + bar.get_width()/2., height,
                                f'{int(height)}', ha='center', va='bottom')
                    
                    plt.tight_layout()
                    st.pyplot(fig1)
                
                with col2:
                    # 饼图
                    fig2, ax2 = plt.subplots(figsize=(8, 5))
                    wedges, texts, autotexts = ax2.pie(
                        type_count.values, 
                        labels=type_count.index, 
                        autopct='%1.1f%%',
                        colors=[colors_map.get(x, "gray") for x in type_count.index],
                        startangle=90
                    )
                    ax2.set_title("需求属性占比", fontsize=14, pad=20)
                    st.pyplot(fig2)
                
                # 3. Better-Worse四象限图
                st.subheader("3. Better-Worse四象限分析")
                
                fig3, ax3 = plt.subplots(figsize=(14, 10))
                
                colors = {"必备属性M": "red", "期望属性O": "orange", 
                         "魅力属性A": "green", "无差异属性I": "blue", 
                         "反向属性R": "purple", "可疑结果Q": "gray"}
                
                for _, row in result_df.iterrows():
                    ax3.scatter(row["Worse系数"], row["Better系数"],
                               c=colors.get(row["最终分类"], "gray"), 
                               s=150, alpha=0.7, edgecolors='black', linewidth=1)
                    
                    # 调整标签位置避免重叠
                    ax3.annotate(row["功能"][:6], 
                                (row["Worse系数"], row["Better系数"]),
                                fontsize=9, ha='center', va='bottom',
                                xytext=(0, 5), textcoords='offset points')
                
                # 添加象限分割线
                ax3.axhline(0.5, color="gray", linestyle="--", alpha=0.3)
                ax3.axvline(-0.5, color="gray", linestyle="--", alpha=0.3)
                ax3.axhline(0, color="black", linestyle="-", alpha=0.3)
                ax3.axvline(0, color="black", linestyle="-", alpha=0.3)
                
                ax3.set_xlabel("Worse系数（-1到0，越负越重要）", fontsize=12)
                ax3.set_ylabel("Better系数（0到1，越高越能提升满意度）", fontsize=12)
                ax3.set_title("KANO Better-Worse四象限图\n（右上：期望属性 | 左上：魅力属性 | 右下：必备属性）", fontsize=14)
                ax3.grid(alpha=0.3)
                ax3.set_xlim(-1.1, 0.1)
                ax3.set_ylim(-0.1, 1.1)
                
                # 添加象限标签
                ax3.text(-0.25, 0.75, "魅力属性\n高Better | 低Worse", transform=ax3.transAxes, 
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.3), fontsize=10)
                ax3.text(-0.25, 0.25, "无差异属性", transform=ax3.transAxes,
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='lightgray', alpha=0.3), fontsize=10)
                ax3.text(0.75, 0.75, "期望属性\n高Better | 高Worse", transform=ax3.transAxes,
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3), fontsize=10)
                ax3.text(0.75, 0.25, "必备属性\n低Better | 高Worse", transform=ax3.transAxes,
                        ha='center', va='center', bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.3), fontsize=10)
                
                plt.tight_layout()
                st.pyplot(fig3)
                
                # ==================== 导出结果 ====================
                st.divider()
                st.header("💾 导出结果")
                
                # 准备导出数据（数值格式）
                export_df = result_df.copy()
                
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
                    # 导出四象限图
                    img_buffer = BytesIO()
                    fig3.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
                    img_buffer.seek(0)
                    st.download_button(
                        label="📥 下载四象限图 (PNG)",
                        data=img_buffer,
                        file_name="KANO四象限图.png",
                        mime="image/png"
                    )
                
                # 关键洞察
                st.divider()
                st.header("💡 关键洞察与优先级建议")
                
                m_count = len(result_df[result_df["最终分类"] == "必备属性M"])
                o_count = len(result_df[result_df["最终分类"] == "期望属性O"])
                a_count = len(result_df[result_df["最终分类"] == "魅力属性A"])
                i_count = len(result_df[result_df["最终分类"] == "无差异属性I"])
                
                cols = st.columns(4)
                cols[0].metric("🔴 必备属性(M)", m_count, f"{m_count/len(result_df):.0%}")
                cols[1].metric("🟠 期望属性(O)", o_count, f"{o_count/len(result_df):.0%}")
                cols[2].metric("🟢 魅力属性(A)", a_count, f"{a_count/len(result_df):.0%}")
                cols[3].metric("🔵 无差异属性(I)", i_count, f"{i_count/len(result_df):.0%}")
                
                # 具体建议
                st.info("""
                **产品开发优先级建议：**
                
                1. **第一优先级（必备属性 M）**：必须实现，缺失会导致用户极度不满
                2. **第二优先级（期望属性 O）**：投入产出比最高，做得越好用户越满意  
                3. **第三优先级（魅力属性 A）**：差异化卖点，能产生惊喜感和传播点
                4. **最后考虑（无差异属性 I）**：资源有限时可暂缓或简化
                """)
                
                # 显示各分类的具体功能
                col_show1, col_show2 = st.columns(2)
                
                with col_show1:
                    if m_count > 0:
                        st.subheader("🔴 必备属性功能")
                        m_funcs = result_df[result_df["最终分类"] == "必备属性M"]["功能"].tolist()
                        for func in m_funcs:
                            st.write(f"- {func}")
                    
                    if o_count > 0:
                        st.subheader("🟠 期望属性功能")
                        o_funcs = result_df[result_df["最终分类"] == "期望属性O"]["功能"].tolist()
                        for func in o_funcs:
                            st.write(f"- {func}")
                
                with col_show2:
                    if a_count > 0:
                        st.subheader("🟢 魅力属性功能")
                        a_funcs = result_df[result_df["最终分类"] == "魅力属性A"]["功能"].tolist()
                        for func in a_funcs:
                            st.write(f"- {func}")
                    
                    if i_count > 0:
                        st.subheader("🔵 无差异属性功能")
                        i_funcs = result_df[result_df["最终分类"] == "无差异属性I"]["功能"].tolist()
                        for func in i_funcs:
                            st.write(f"- {func}")

    except Exception as e:
        st.error(f"❌ 处理文件时出错：{str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    # 空状态提示
    st.info("👆 请上传Excel文件开始分析")
    
    # 显示预期格式
    st.divider()
    st.subheader("📄 数据格式检查清单")
    check_cols = pd.DataFrame({
        "功能序号": [f"功能{i}" for i in range(1, 21)],
        "正向题示例": [f"{i}、智能厨具具备...—如果有这个功能您的评价是：" for i in range(1, 21)],
        "反向题示例": [f"{i}、如果没有这个功能您的评价是：" for i in range(1, 21)]
    })
    st.dataframe(check_cols, use_container_width=True, height=300)
