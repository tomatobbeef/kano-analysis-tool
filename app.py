import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

# -------------------------- 1. 基础设置 --------------------------
file_path = "你的问卷.xlsx"  # 替换为实际路径

# 需求项名称（与问卷1-20项对应）
demand_names = [
    "一键菜谱同步、自动匹配烹饪步骤",
    "APP安全监护预警（干烧、漏电、溢锅）",
    "定时预约、烹饪自动完成",
    "健康饮食关怀与提醒",
    "APP轻量化操作、界面简洁易上手",
    "强效防油烟、减少厨房油污",
    "烹饪后自动保温",
    "一人食专属小容量、不浪费食材",
    "多重安全防护、杜绝使用隐患",
    "机身/配件易拆卸、清洗无死角",
    "低噪音运行",
    "温暖治愈、贴合独居氛围外观",
    "按键/触屏操作区清晰、好操作",
    "小巧机身、不占厨房空间",
    "温暖低饱和色系、视觉舒适",
    "温润安全、防烫防滑材质",
    "触感温润安全表面处理",
    "零学习成本、拿到就会用交互",
    "安全状态实时强反馈（灯光/声音）",
    "情感化关怀、增强陪伴感"
]

# KANO评价映射（问卷文本→分数）
score_map = {
    "很喜欢": 5,
    "理所当然": 4,
    "无所谓": 3,
    "勉强接受": 2,
    "很不喜欢": 1
}

# -------------------------- 2. 数据读取 --------------------------
df = pd.read_excel(file_path)

# 提取有效列（正向=有功能，反向=无功能）
func_cols = [col for col in df.columns if "—如果有这个功能" in col]
dysfunc_cols = [col for col in df.columns if "如果没有这个功能" in col]

print(f"找到 {len(func_cols)} 个正向题，{len(dysfunc_cols)} 个反向题")

# 文本转分数
df_func = df[func_cols].copy()
df_dysfunc = df[dysfunc_cols].copy()

for col in df_func.columns:
    df_func[col] = df_func[col].map(score_map)
for col in df_dysfunc.columns:
    df_dysfunc[col] = df_dysfunc[col].map(score_map)

# -------------------------- 3. KANO分类（核心修正）--------------------------
def classify_kano(pos_score, neg_score):
    """
    标准KANO分类逻辑：
    - M(必备): 有则满意(4/5)，无则不满(1/2) → (4,1),(4,2),(5,1),(5,2)
    - O(期望): 有则满意，无则无所谓 → (4,3),(5,3),(3,4),(3,5) 等线性关系
    - A(魅力): 有则喜，无则无所谓 → (5,3)其实是期望，(1/2,3)才是魅力？
    
    正确逻辑（基于KANO理论）：
    正向\反向 | 1很喜欢(无) | 2理所当然 | 3无所谓 | 4勉强接受 | 5很不喜欢
    ----------|-------------|-----------|---------|-----------|-----------
    5很喜欢   |      Q      |     A     |    A    |     A     |     O    
    4理所当然 |      M      |     I     |    I    |     I     |     M    
    3无所谓   |      R      |     R     |    I    |     R     |     R    
    2勉强接受 |      R      |     R     |    R    |     R     |     R    
    1很不喜欢 |      R      |     R     |    R    |     R     |     R    
    """
    # 标准化输入
    p = int(pos_score)
    n = int(neg_score)
    
    # 反向属性R：有功能时不喜欢/勉强接受（1/2），或没有时很喜欢（1）——意味着用户不想要这个功能
    if p in [1, 2] or n == 1:
        return "反向属性R"
    
    # 必备属性M：有则理所当然(4)或很喜欢(5)，无则很不喜欢(1)或勉强接受(2)
    if p in [4, 5] and n in [1, 2]:
        return "必备属性M"
    
    # 魅力属性A：有则很喜欢(5)，无则无所谓(3)——惊喜型功能
    if p == 5 and n == 3:
        return "魅力属性A"
    
    # 期望属性O：有则很喜欢/理所当然，无则勉强接受/很不喜欢（线性关系）
    if (p in [4, 5] and n in [4, 5]) or (p == 5 and n == 4):
        return "期望属性O"
    
    # 无差异属性I：有或没有都无所谓/理所当然
    if p == 3 and n in [2, 3, 4]:
        return "无差异属性I"
    if p == 4 and n in [2, 3, 4]:
        return "无差异属性I"
    
    # 其他未覆盖情况标记为可疑
    return "可疑结果Q"

# -------------------------- 4. 逐需求分析 --------------------------
result_list = []

for i in range(20):
    func_name = demand_names[i]
    func_col = df_func.iloc[:, i]
    dys_col = df_dysfunc.iloc[:, i]
    
    valid_count = 0
    type_counts = {"必备属性M": 0, "期望属性O": 0, "魅力属性A": 0, 
                   "无差异属性I": 0, "反向属性R": 0, "可疑结果Q": 0}
    
    # 逐样本分类统计
    for f_score, d_score in zip(func_col, dys_col):
        if pd.isna(f_score) or pd.isna(d_score):
            continue
        valid_count += 1
        kano_type = classify_kano(f_score, d_score)
        type_counts[kano_type] += 1
    
    if valid_count == 0:
        continue
    
    # 计算占比
    for k in type_counts:
        type_counts[k] = type_counts[k] / valid_count
    
    # 取占比最高的为最终分类
    final_type = max(type_counts, key=type_counts.get)
    
    # Better-Worse系数（与问卷星一致，使用百分比）
    # Better = (魅力 + 期望) / (魅力+期望+必备+无差异)
    # Worse = -(必备 + 期望) / (魅力+期望+必备+无差异)
    total_valid = type_counts["魅力属性A"] + type_counts["期望属性O"] + \
                  type_counts["必备属性M"] + type_counts["无差异属性I"]
    
    if total_valid > 0:
        better = (type_counts["魅力属性A"] + type_counts["期望属性O"]) / total_valid * 100
        worse = -(type_counts["必备属性M"] + type_counts["期望属性O"]) / total_valid * 100
    else:
        better = 0
        worse = 0
    
    result_list.append({
        "序号": i+1,
        "功能": func_name,
        "KANO属性": final_type,
        "Better系数": f"{better:.2f}%",
        "Worse系数": f"{worse:.2f}%",
        "样本数": valid_count,
        **{k: f"{v:.2%}" for k, v in type_counts.items()}  # 添加各类型占比详情
    })

# -------------------------- 5. 结果输出 --------------------------
df_result = pd.DataFrame(result_list)

# 按KANO属性排序（必备>期望>魅力>无差异>反向）
type_order = {"必备属性M": 0, "期望属性O": 1, "魅力属性A": 2, "无差异属性I": 3, "反向属性R": 4, "可疑结果Q": 5}
df_result["排序"] = df_result["KANO属性"].map(type_order)
df_result = df_result.sort_values("排序").drop("排序", axis=1).reset_index(drop=True)

print("\n" + "="*100)
print("KANO模型需求分析结果（与问卷星逻辑一致）")
print("="*100)
print(df_result[["序号", "功能", "KANO属性", "Better系数", "Worse系数"]].to_string(index=False))

# 保存详细结果
df_result.to_excel("KANO分析结果_修正版.xlsx", index=False)
print("\n✅ 详细结果已保存至：KANO分析结果_修正版.xlsx")

# -------------------------- 6. 可视化 --------------------------
plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# 四象限图（Better-Worse图）
fig, ax = plt.subplots(figsize=(12, 8))

colors = {"必备属性M": "red", "期望属性O": "orange", "魅力属性A": "green", 
          "无差异属性I": "gray", "反向属性R": "purple", "可疑结果Q": "blue"}

for idx, row in df_result.iterrows():
    better_val = float(row["Better系数"].replace("%", ""))
    worse_val = float(row["Worse系数"].replace("%", ""))
    
    ax.scatter(worse_val, better_val, 
               c=colors.get(row["KANO属性"], "black"), 
               s=120, alpha=0.7, edgecolors='black')
    
    # 标注序号
    ax.annotate(str(row["序号"]), (worse_val, better_val), 
                xytext=(5, 5), textcoords='offset points', fontsize=9)

# 添加参考线
ax.axhline(y=50, color='gray', linestyle='--', alpha=0.3)
ax.axvline(x=-50, color='gray', linestyle='--', alpha=0.3)
ax.axhline(y=0, color='black', linestyle='-', alpha=0.2)
ax.axvline(x=0, color='black', linestyle='-', alpha=0.2)

ax.set_xlabel("Worse系数（影响不满程度）", fontsize=12)
ax.set_ylabel("Better系数（提升满意程度）", fontsize=12)
ax.set_title("KANO Better-Worse四象限分析", fontsize=14)
ax.grid(alpha=0.3)

# 添加象限说明
ax.text(0.05, 0.95, "魅力属性", transform=ax.transAxes, fontsize=10, 
        bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
ax.text(0.05, 0.05, "无差异", transform=ax.transAxes, fontsize=10,
        bbox=dict(boxstyle='round', facecolor='lightgray', alpha=0.5))
ax.text(0.75, 0.95, "期望属性", transform=ax.transAxes, fontsize=10,
        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
ax.text(0.75, 0.05, "必备属性", transform=ax.transAxes, fontsize=10,
        bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))

plt.tight_layout()
plt.savefig("KANO四象限图_修正版.png", dpi=300)
print("✅ 四象限图已保存")
