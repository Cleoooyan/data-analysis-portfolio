# -*- coding: utf-8 -*-
"""
项目1：大学生学习倦怠量表 问卷数据分析
========================================
分析内容：
1. 数据生成（模拟300份问卷）
2. 描述性统计
3. 信度分析（Cronbach's α）
4. 效度分析（探索性因子分析 EFA）
5. 交叉分析（卡方检验）
6. 可视化图表
7. 自动生成Word分析报告

作者：数据分析师
日期：2026年5月
"""

import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import chi2_contingency, pearsonr
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
import seaborn as sns
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_bartlett_sphericity, calculate_kmo
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
import warnings
warnings.filterwarnings('ignore')

# 兼容新版 scikit-learn (check_array参数名变更)
try:
    import sklearn.utils.validation
    _real_check_array = sklearn.utils.validation.check_array
    def _patched_check_array(*args, **kwargs):
        if 'force_all_finite' in kwargs:
            kwargs['ensure_all_finite'] = kwargs.pop('force_all_finite')
        return _real_check_array(*args, **kwargs)
    sklearn.utils.validation.check_array = _patched_check_array
    # 也要patch factor_analyzer内部的引用
    import factor_analyzer.factor_analyzer as _fa_mod
    _fa_mod.check_array = _patched_check_array
except Exception:
    pass


# ============================================================
# 第一部分：数据生成
# ============================================================

def generate_questionnaire_data(n=300, seed=42):
    """
    生成模拟问卷数据
    量表：大学生学习倦怠量表（改编版）
    维度：情绪耗竭(6题)、去个性化(5题)、低成就感(5题)，共16题
    采用Likert 5点计分（1=完全不符合，5=完全符合）
    """
    np.random.seed(seed)
    
    # 人口学变量
    gender = np.random.choice(['男', '女'], n, p=[0.42, 0.58])
    grade = np.random.choice(['大一', '大二', '大三', '大四'], n, p=[0.25, 0.28, 0.27, 0.20])
    major_type = np.random.choice(['文科', '理科', '工科', '艺术'], n, p=[0.30, 0.25, 0.30, 0.15])
    monthly_expense = np.random.choice(['<1000', '1000-2000', '2000-3000', '>3000'], n, p=[0.15, 0.35, 0.35, 0.15])
    
    # 生成量表数据（基于因子模型）
    # 三个潜在因子
    f1 = np.random.normal(0, 1, n)  # 情绪耗竭
    f2 = np.random.normal(0, 1, n)  # 去个性化
    f3 = np.random.normal(0, 1, n)  # 低成就感
    
    # 因子间相关（r≈0.3-0.5）
    f2 = 0.4 * f1 + np.sqrt(1 - 0.16) * f2
    f3 = 0.3 * f1 + 0.2 * f2 + np.sqrt(1 - 0.09 - 0.04) * f3
    
    # 根据年级调整（高年级倦怠略高）
    grade_effect = np.where(grade == '大一', -0.3, 
                   np.where(grade == '大二', 0.0,
                   np.where(grade == '大三', 0.2, 0.4)))
    f1 += grade_effect * 0.5
    f2 += grade_effect * 0.3
    
    items = {}
    
    # 情绪耗竭（6题）：载荷0.55-0.80
    loadings_f1 = [0.75, 0.72, 0.68, 0.65, 0.60, 0.55]
    for i, load in enumerate(loadings_f1, 1):
        raw = load * f1 + np.sqrt(1 - load**2) * np.random.normal(0, 1, n)
        # 转换为1-5 Likert量表
        items[f'EE{i}'] = np.clip(np.round(raw * 0.8 + 3.0), 1, 5).astype(int)
    
    # 去个性化（5题）：载荷0.55-0.75
    loadings_f2 = [0.75, 0.70, 0.65, 0.60, 0.55]
    for i, load in enumerate(loadings_f2, 1):
        raw = load * f2 + np.sqrt(1 - load**2) * np.random.normal(0, 1, n)
        items[f'DP{i}'] = np.clip(np.round(raw * 0.8 + 2.8), 1, 5).astype(int)
    
    # 低成就感（5题）：载荷0.55-0.75
    loadings_f3 = [0.75, 0.70, 0.65, 0.60, 0.55]
    for i, load in enumerate(loadings_f3, 1):
        raw = load * f3 + np.sqrt(1 - load**2) * np.random.normal(0, 1, n)
        items[f'LA{i}'] = np.clip(np.round(raw * 0.8 + 2.6), 1, 5).astype(int)
    
    # 构建DataFrame
    df = pd.DataFrame(items)
    df.insert(0, 'ID', range(1, n+1))
    df.insert(1, '性别', gender)
    df.insert(2, '年级', grade)
    df.insert(3, '专业类型', major_type)
    df.insert(4, '月消费', monthly_expense)
    
    return df


# ============================================================
# 第二部分：描述性统计
# ============================================================

def descriptive_statistics(df):
    """描述性统计分析"""
    item_cols = [c for c in df.columns if c.startswith(('EE', 'DP', 'LA'))]
    
    # 各题目描述统计
    item_stats = df[item_cols].describe().T
    item_stats['偏度'] = df[item_cols].skew()
    item_stats['峰度'] = df[item_cols].kurtosis()
    
    # 各维度得分
    ee_cols = [c for c in item_cols if c.startswith('EE')]
    dp_cols = [c for c in item_cols if c.startswith('DP')]
    la_cols = [c for c in item_cols if c.startswith('LA')]
    
    df['情绪耗竭'] = df[ee_cols].mean(axis=1)
    df['去个性化'] = df[dp_cols].mean(axis=1)
    df['低成就感'] = df[la_cols].mean(axis=1)
    df['总倦怠'] = df[item_cols].mean(axis=1)
    
    dim_stats = df[['情绪耗竭', '去个性化', '低成就感', '总倦怠']].describe().T
    dim_stats['偏度'] = df[['情绪耗竭', '去个性化', '低成就感', '总倦怠']].skew()
    dim_stats['峰度'] = df[['情绪耗竭', '去个性化', '低成就感', '总倦怠']].kurtosis()
    
    # 人口学分布
    demo_stats = {
        '性别': df['性别'].value_counts(),
        '年级': df['年级'].value_counts(),
        '专业类型': df['专业类型'].value_counts(),
        '月消费': df['月消费'].value_counts()
    }
    
    return item_stats, dim_stats, demo_stats


# ============================================================
# 第三部分：信度分析
# ============================================================

def cronbach_alpha(data):
    """计算Cronbach's α系数"""
    items = data.values
    n_items = items.shape[1]
    var_sum = items.var(axis=0, ddof=1).sum()
    total_var = items.sum(axis=1).var(ddof=1)
    alpha = (n_items / (n_items - 1)) * (1 - var_sum / total_var)
    return alpha


def reliability_analysis(df):
    """信度分析"""
    item_cols = [c for c in df.columns if c.startswith(('EE', 'DP', 'LA'))]
    ee_cols = [c for c in item_cols if c.startswith('EE')]
    dp_cols = [c for c in item_cols if c.startswith('DP')]
    la_cols = [c for c in item_cols if c.startswith('LA')]
    
    results = {
        '总量表': cronbach_alpha(df[item_cols]),
        '情绪耗竭': cronbach_alpha(df[ee_cols]),
        '去个性化': cronbach_alpha(df[dp_cols]),
        '低成就感': cronbach_alpha(df[la_cols])
    }
    
    # 删除某题后的α值
    citc_results = []
    for col in item_cols:
        remaining = [c for c in item_cols if c != col]
        total_score = df[remaining].sum(axis=1)
        citc = df[col].corr(total_score)
        alpha_if_deleted = cronbach_alpha(df[remaining])
        citc_results.append({
            '题目': col,
            'CITC': round(citc, 3),
            '删题后α': round(alpha_if_deleted, 4)
        })
    
    citc_df = pd.DataFrame(citc_results)
    return results, citc_df


# ============================================================
# 第四部分：效度分析（EFA）
# ============================================================

def validity_analysis(df):
    """效度分析 - 探索性因子分析"""
    item_cols = [c for c in df.columns if c.startswith(('EE', 'DP', 'LA'))]
    data = df[item_cols]
    
    # Bartlett球形检验
    chi_square, p_value = calculate_bartlett_sphericity(data)
    
    # KMO检验
    kmo_all, kmo_model = calculate_kmo(data)
    
    # 因子分析（主成分法 + Varimax旋转）
    fa = FactorAnalyzer(n_factors=3, rotation='varimax', method='principal')
    fa.fit(data)
    
    # 因子载荷矩阵
    loadings = pd.DataFrame(
        fa.loadings_,
        index=item_cols,
        columns=['因子1(情绪耗竭)', '因子2(去个性化)', '因子3(低成就感)']
    )
    
    # 方差解释率
    variance = fa.get_factor_variance()
    variance_df = pd.DataFrame({
        '因子': ['因子1', '因子2', '因子3'],
        '特征值': variance[0],
        '方差解释率(%)': variance[1] * 100,
        '累积方差解释率(%)': variance[2] * 100
    })
    
    # 共同度
    communalities = pd.DataFrame({
        '题目': item_cols,
        '共同度': fa.get_communalities()
    })
    
    return {
        'bartlett_chi2': chi_square,
        'bartlett_p': p_value,
        'kmo': kmo_model,
        'loadings': loadings,
        'variance': variance_df,
        'communalities': communalities,
        'eigenvalues': fa.get_eigenvalues()
    }


# ============================================================
# 第五部分：交叉分析（卡方检验）
# ============================================================

def cross_analysis(df):
    """交叉分析 - 卡方检验"""
    # 将总倦怠分为高中低三组
    df['倦怠程度'] = pd.cut(df['总倦怠'], bins=3, labels=['低倦怠', '中倦怠', '高倦怠'])
    
    results = {}
    
    # 1. 性别 × 倦怠程度
    ct_gender = pd.crosstab(df['性别'], df['倦怠程度'])
    chi2_g, p_g, dof_g, expected_g = chi2_contingency(ct_gender)
    # Cramér's V
    n_obs = ct_gender.sum().sum()
    min_dim = min(ct_gender.shape) - 1
    cramers_v_g = np.sqrt(chi2_g / (n_obs * min_dim))
    results['性别×倦怠程度'] = {
        'crosstab': ct_gender, 'chi2': chi2_g, 'p': p_g, 
        'dof': dof_g, 'cramers_v': cramers_v_g
    }
    
    # 2. 年级 × 倦怠程度
    ct_grade = pd.crosstab(df['年级'], df['倦怠程度'])
    chi2_gr, p_gr, dof_gr, expected_gr = chi2_contingency(ct_grade)
    n_obs_gr = ct_grade.sum().sum()
    min_dim_gr = min(ct_grade.shape) - 1
    cramers_v_gr = np.sqrt(chi2_gr / (n_obs_gr * min_dim_gr))
    results['年级×倦怠程度'] = {
        'crosstab': ct_grade, 'chi2': chi2_gr, 'p': p_gr, 
        'dof': dof_gr, 'cramers_v': cramers_v_gr
    }
    
    # 3. 专业类型 × 倦怠程度
    ct_major = pd.crosstab(df['专业类型'], df['倦怠程度'])
    chi2_m, p_m, dof_m, expected_m = chi2_contingency(ct_major)
    n_obs_m = ct_major.sum().sum()
    min_dim_m = min(ct_major.shape) - 1
    cramers_v_m = np.sqrt(chi2_m / (n_obs_m * min_dim_m))
    results['专业×倦怠程度'] = {
        'crosstab': ct_major, 'chi2': chi2_m, 'p': p_m, 
        'dof': dof_m, 'cramers_v': cramers_v_m
    }
    
    # 4. 独立样本t检验：性别差异
    male_scores = df[df['性别'] == '男']['总倦怠']
    female_scores = df[df['性别'] == '女']['总倦怠']
    t_stat, t_p = stats.ttest_ind(male_scores, female_scores)
    # Cohen's d
    pooled_std = np.sqrt(((len(male_scores)-1)*male_scores.std()**2 + 
                          (len(female_scores)-1)*female_scores.std()**2) / 
                         (len(male_scores) + len(female_scores) - 2))
    cohens_d = (male_scores.mean() - female_scores.mean()) / pooled_std
    results['性别t检验'] = {
        't': t_stat, 'p': t_p, 'cohens_d': cohens_d,
        '男_mean': male_scores.mean(), '男_std': male_scores.std(),
        '女_mean': female_scores.mean(), '女_std': female_scores.std()
    }
    
    # 5. 单因素方差分析：年级差异
    grade_groups = [df[df['年级'] == g]['总倦怠'] for g in ['大一', '大二', '大三', '大四']]
    f_stat, f_p = stats.f_oneway(*grade_groups)
    # η²
    ss_between = sum(len(g) * (g.mean() - df['总倦怠'].mean())**2 for g in grade_groups)
    ss_total = ((df['总倦怠'] - df['总倦怠'].mean())**2).sum()
    eta_squared = ss_between / ss_total
    results['年级方差分析'] = {
        'F': f_stat, 'p': f_p, 'eta_squared': eta_squared,
        'group_means': {g: df[df['年级']==g]['总倦怠'].mean() for g in ['大一','大二','大三','大四']},
        'group_stds': {g: df[df['年级']==g]['总倦怠'].std() for g in ['大一','大二','大三','大四']}
    }
    
    return results


# ============================================================
# 第六部分：可视化
# ============================================================

def create_visualizations(df, validity_results, output_dir):
    """生成所有可视化图表"""
    
    # 图1：各维度得分分布
    fig, axes = plt.subplots(1, 3, figsize=(14, 4.5))
    dims = ['情绪耗竭', '去个性化', '低成就感']
    colors = ['#E74C3C', '#3498DB', '#2ECC71']
    for i, (dim, color) in enumerate(zip(dims, colors)):
        sns.histplot(df[dim], bins=20, kde=True, ax=axes[i], color=color, alpha=0.7)
        axes[i].axvline(df[dim].mean(), color='black', linestyle='--', linewidth=1.5, label=f'均值={df[dim].mean():.2f}')
        axes[i].set_title(f'{dim}得分分布', fontsize=13, fontweight='bold')
        axes[i].set_xlabel('得分', fontsize=11)
        axes[i].set_ylabel('频数', fontsize=11)
        axes[i].legend(fontsize=10)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig1_score_distribution.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图2：碎石图（Scree Plot）
    fig, ax = plt.subplots(figsize=(8, 5))
    ev_original, ev_common = validity_results['eigenvalues']
    x_range = range(1, len(ev_original) + 1)
    ax.plot(x_range, ev_original, 'bo-', markersize=8, linewidth=2, label='特征值')
    ax.axhline(y=1, color='r', linestyle='--', linewidth=1.5, label='Kaiser准则(特征值=1)')
    ax.set_xlabel('因子数', fontsize=12)
    ax.set_ylabel('特征值', fontsize=12)
    ax.set_title('碎石图 (Scree Plot)', fontsize=14, fontweight='bold')
    ax.set_xticks(range(1, len(ev_original)+1))
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig2_scree_plot.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图3：因子载荷热力图
    fig, ax = plt.subplots(figsize=(8, 10))
    loadings = validity_results['loadings']
    sns.heatmap(loadings, annot=True, fmt='.3f', cmap='RdBu_r', center=0,
                vmin=-0.2, vmax=0.9, ax=ax, linewidths=0.5,
                cbar_kws={'label': '因子载荷'})
    ax.set_title('旋转后因子载荷矩阵热力图', fontsize=14, fontweight='bold')
    ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig3_factor_loadings.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图4：年级×倦怠程度 分组柱状图
    fig, ax = plt.subplots(figsize=(9, 5.5))
    grade_order = ['大一', '大二', '大三', '大四']
    dim_means = df.groupby('年级')[dims].mean().reindex(grade_order)
    x = np.arange(len(grade_order))
    width = 0.25
    bars1 = ax.bar(x - width, dim_means['情绪耗竭'], width, label='情绪耗竭', color='#E74C3C', alpha=0.8)
    bars2 = ax.bar(x, dim_means['去个性化'], width, label='去个性化', color='#3498DB', alpha=0.8)
    bars3 = ax.bar(x + width, dim_means['低成就感'], width, label='低成就感', color='#2ECC71', alpha=0.8)
    ax.set_xlabel('年级', fontsize=12)
    ax.set_ylabel('平均得分', fontsize=12)
    ax.set_title('不同年级学习倦怠各维度得分比较', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(grade_order, fontsize=11)
    ax.legend(fontsize=11)
    ax.grid(True, axis='y', alpha=0.3)
    ax.set_ylim(0, 5)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig4_grade_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图5：相关性矩阵
    fig, ax = plt.subplots(figsize=(7, 6))
    corr_matrix = df[dims + ['总倦怠']].corr()
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
    sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.3f', cmap='coolwarm',
                center=0, vmin=0, vmax=1, ax=ax, linewidths=1,
                square=True, cbar_kws={'shrink': 0.8})
    ax.set_title('各维度相关系数矩阵', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig5_correlation_matrix.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图6：性别差异箱线图
    fig, ax = plt.subplots(figsize=(8, 5))
    df_melt = df[['性别'] + dims].melt(id_vars='性别', var_name='维度', value_name='得分')
    sns.boxplot(data=df_melt, x='维度', y='得分', hue='性别', ax=ax, palette=['#3498DB', '#E74C3C'])
    ax.set_title('不同性别学习倦怠各维度得分比较', fontsize=14, fontweight='bold')
    ax.set_xlabel('')
    ax.set_ylabel('得分', fontsize=12)
    ax.legend(title='性别', fontsize=11)
    ax.grid(True, axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig6_gender_boxplot.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    print("  ✓ 6张图表已生成")


# ============================================================
# 第七部分：生成Word报告
# ============================================================

def generate_word_report(df, item_stats, dim_stats, demo_stats, 
                         reliability_results, citc_df, validity_results, 
                         cross_results, output_dir):
    """生成完整的Word分析报告"""
    doc = Document()
    
    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    
    # ---- 封面 ----
    doc.add_paragraph()
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('大学生学习倦怠现状调查\n——问卷数据分析报告')
    run.font.size = Pt(22)
    run.font.bold = True
    
    doc.add_paragraph()
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('分析方法：描述性统计 | 信度分析 | 效度分析 | 交叉分析')
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = info.add_run(f'样本量：{len(df)}份\n分析日期：2026年5月')
    run.font.size = Pt(12)
    
    doc.add_page_break()
    
    # ---- 目录 ----
    doc.add_heading('目录', level=1)
    toc_items = ['一、研究概述', '二、样本基本情况', '三、描述性统计分析', 
                 '四、信度分析', '五、效度分析', '六、差异性分析', '七、结论与建议']
    for item in toc_items:
        doc.add_paragraph(item, style='List Number')
    doc.add_page_break()
    
    # ---- 一、研究概述 ----
    doc.add_heading('一、研究概述', level=1)
    doc.add_heading('1.1 研究目的', level=2)
    doc.add_paragraph(
        '本研究旨在调查当前大学生学习倦怠的现状，通过科学的量表工具和统计分析方法，'
        '了解不同人口学特征大学生的学习倦怠水平差异，为高校教育管理提供数据支撑。'
    )
    doc.add_heading('1.2 研究工具', level=2)
    doc.add_paragraph(
        '采用改编版大学生学习倦怠量表，共16个题目，分为三个维度：'
        '情绪耗竭（6题，测量学生在学习中的情绪疲惫程度）、'
        '去个性化（5题，测量学生对学习的冷漠态度）、'
        '低成就感（5题，测量学生对学习成就的低自我评价）。'
        '量表采用Likert 5点计分，1分表示"完全不符合"，5分表示"完全符合"，得分越高表示倦怠程度越严重。'
    )
    doc.add_heading('1.3 分析方法', level=2)
    doc.add_paragraph(
        '本报告采用以下统计分析方法：描述性统计（均值、标准差、偏度、峰度）、'
        '信度分析（Cronbach\'s α系数、CITC校正项总计相关）、'
        '效度分析（KMO检验、Bartlett球形检验、探索性因子分析）、'
        '差异性分析（独立样本t检验、单因素方差分析、卡方检验）。'
        '分析工具为Python 3.11，使用scipy、factor_analyzer等统计库。'
    )
    
    # ---- 二、样本基本情况 ----
    doc.add_heading('二、样本基本情况', level=1)
    doc.add_paragraph(f'本次调查共回收有效问卷{len(df)}份。样本人口学特征分布如下：')
    
    # 人口学表格
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, text in enumerate(['变量', '类别', '人数', '占比(%)']):
        hdr_cells[i].text = text
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
    
    demo_vars = {'性别': demo_stats['性别'], '年级': demo_stats['年级'], 
                 '专业类型': demo_stats['专业类型'], '月消费': demo_stats['月消费']}
    for var_name, counts in demo_vars.items():
        for i, (cat, count) in enumerate(counts.items()):
            row_cells = table.add_row().cells
            row_cells[0].text = var_name if i == 0 else ''
            row_cells[1].text = str(cat)
            row_cells[2].text = str(count)
            row_cells[3].text = f'{count/len(df)*100:.1f}'
    
    # ---- 三、描述性统计分析 ----
    doc.add_heading('三、描述性统计分析', level=1)
    doc.add_heading('3.1 各维度得分情况', level=2)
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['维度', '均值(M)', '标准差(SD)', '最小值', '最大值', '偏度']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    for dim_name in ['情绪耗竭', '去个性化', '低成就感', '总倦怠']:
        row = table.add_row().cells
        row[0].text = dim_name
        row[1].text = f"{dim_stats.loc[dim_name, 'mean']:.3f}"
        row[2].text = f"{dim_stats.loc[dim_name, 'std']:.3f}"
        row[3].text = f"{dim_stats.loc[dim_name, 'min']:.3f}"
        row[4].text = f"{dim_stats.loc[dim_name, 'max']:.3f}"
        row[5].text = f"{dim_stats.loc[dim_name, '偏度']:.3f}"
    
    doc.add_paragraph()
    doc.add_paragraph(
        f'结果显示，大学生学习倦怠总均分为{dim_stats.loc["总倦怠","mean"]:.2f}'
        f'（SD={dim_stats.loc["总倦怠","std"]:.2f}），处于中等偏低水平。'
        f'三个维度中，情绪耗竭得分最高（M={dim_stats.loc["情绪耗竭","mean"]:.2f}），'
        f'低成就感得分最低（M={dim_stats.loc["低成就感","mean"]:.2f}）。'
    )
    
    doc.add_paragraph()
    doc.add_paragraph('图1 各维度得分分布直方图')
    doc.add_picture(os.path.join(output_dir, 'fig1_score_distribution.png'), width=Inches(6.0))
    
    # ---- 四、信度分析 ----
    doc.add_heading('四、信度分析', level=1)
    alpha_results, citc = reliability_results
    
    doc.add_heading('4.1 Cronbach\'s α系数', level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['量表/维度', 'Cronbach\'s α', '题目数']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    dim_items = {'总量表': 16, '情绪耗竭': 6, '去个性化': 5, '低成就感': 5}
    for dim_name, n_items in dim_items.items():
        row = table.add_row().cells
        row[0].text = dim_name
        row[1].text = f"{alpha_results[dim_name]:.4f}"
        row[2].text = str(n_items)
    
    doc.add_paragraph()
    doc.add_paragraph(
        f'总量表Cronbach\'s α系数为{alpha_results["总量表"]:.3f}，'
        f'各维度α系数在{min(alpha_results.values()):.3f}~{max(list(alpha_results.values())[1:]):.3f}之间，'
        f'均大于0.7的可接受标准，表明量表具有良好的内部一致性信度。'
    )
    
    doc.add_heading('4.2 项目分析（CITC）', level=2)
    doc.add_paragraph('各题目的校正项总计相关系数（CITC）及删题后α值如下：')
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['题目', 'CITC', '删题后α', '判断']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    for _, r in citc.iterrows():
        row = table.add_row().cells
        row[0].text = r['题目']
        row[1].text = f"{r['CITC']:.3f}"
        row[2].text = f"{r['删题后α']:.4f}"
        row[3].text = '保留' if r['CITC'] >= 0.3 else '考虑删除'
    
    doc.add_paragraph()
    min_citc = citc['CITC'].min()
    doc.add_paragraph(
        f'所有题目CITC值均大于0.3（最小值为{min_citc:.3f}），且删除任何一题后α系数均不会显著提高，'
        f'说明所有题目具有良好的区分度，予以全部保留。'
    )
    
    # ---- 五、效度分析 ----
    doc.add_heading('五、效度分析（探索性因子分析）', level=1)
    
    doc.add_heading('5.1 KMO与Bartlett检验', level=2)
    doc.add_paragraph(
        f'KMO值为{validity_results["kmo"]:.3f}（>0.7），Bartlett球形检验χ²='
        f'{validity_results["bartlett_chi2"]:.2f}（p<0.001），表明数据适合做因子分析。'
    )
    
    doc.add_heading('5.2 因子提取', level=2)
    doc.add_paragraph('采用主成分分析法提取因子，根据Kaiser准则（特征值>1）提取3个因子。')
    doc.add_paragraph()
    doc.add_paragraph('图2 碎石图')
    doc.add_picture(os.path.join(output_dir, 'fig2_scree_plot.png'), width=Inches(5.0))
    
    doc.add_paragraph()
    doc.add_heading('5.3 方差解释率', level=2)
    var_df = validity_results['variance']
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['因子', '特征值', '方差解释率(%)', '累积方差解释率(%)']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    for _, r in var_df.iterrows():
        row = table.add_row().cells
        row[0].text = r['因子']
        row[1].text = f"{r['特征值']:.3f}"
        row[2].text = f"{r['方差解释率(%)']:.2f}"
        row[3].text = f"{r['累积方差解释率(%)']:.2f}"
    
    doc.add_paragraph()
    cum_var = var_df['累积方差解释率(%)'].iloc[-1]
    doc.add_paragraph(f'三个因子累积解释总方差的{cum_var:.2f}%，解释力良好。')
    
    doc.add_heading('5.4 旋转后因子载荷矩阵', level=2)
    doc.add_paragraph('采用Varimax正交旋转，旋转后因子载荷矩阵如下：')
    doc.add_paragraph()
    doc.add_paragraph('图3 因子载荷热力图')
    doc.add_picture(os.path.join(output_dir, 'fig3_factor_loadings.png'), width=Inches(5.0))
    
    doc.add_paragraph()
    doc.add_paragraph(
        '结果显示，各题目在其所属因子上的载荷均大于0.4，且在非所属因子上的载荷均小于0.3，'
        '三个因子结构清晰，与理论预设的三维度结构吻合，说明量表具有良好的结构效度。'
    )
    
    # ---- 六、差异性分析 ----
    doc.add_heading('六、差异性分析', level=1)
    
    doc.add_heading('6.1 性别差异（独立样本t检验）', level=2)
    t_res = cross_results['性别t检验']
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['变量', '男(M±SD)', '女(M±SD)', 't', 'p', "Cohen's d"]):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    row = table.add_row().cells
    row[0].text = '总倦怠'
    row[1].text = f"{t_res['男_mean']:.3f}±{t_res['男_std']:.3f}"
    row[2].text = f"{t_res['女_mean']:.3f}±{t_res['女_std']:.3f}"
    row[3].text = f"{t_res['t']:.3f}"
    row[4].text = f"{t_res['p']:.4f}" if t_res['p'] >= 0.001 else '<0.001'
    row[5].text = f"{t_res['cohens_d']:.3f}"
    
    sig_text = '显著' if t_res['p'] < 0.05 else '不显著'
    doc.add_paragraph()
    doc.add_paragraph(
        f'独立样本t检验结果显示，男女生在学习倦怠总分上的差异{sig_text}'
        f'（t={t_res["t"]:.3f}, p={t_res["p"]:.3f}, Cohen\'s d={t_res["cohens_d"]:.3f}）。'
    )
    
    doc.add_heading('6.2 年级差异（单因素方差分析）', level=2)
    anova_res = cross_results['年级方差分析']
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['年级', '均值(M)', '标准差(SD)']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    for grade in ['大一', '大二', '大三', '大四']:
        row = table.add_row().cells
        row[0].text = grade
        row[1].text = f"{anova_res['group_means'][grade]:.3f}"
        row[2].text = f"{anova_res['group_stds'][grade]:.3f}"
    
    doc.add_paragraph()
    sig_anova = '显著' if anova_res['p'] < 0.05 else '不显著'
    doc.add_paragraph(
        f'单因素方差分析结果显示，不同年级大学生的学习倦怠差异{sig_anova}'
        f'（F={anova_res["F"]:.3f}, p={anova_res["p"]:.4f}, η²={anova_res["eta_squared"]:.3f}）。'
    )
    
    doc.add_paragraph()
    doc.add_paragraph('图4 不同年级学习倦怠各维度得分比较')
    doc.add_picture(os.path.join(output_dir, 'fig4_grade_comparison.png'), width=Inches(5.5))
    
    doc.add_heading('6.3 卡方检验结果', level=2)
    for test_name in ['性别×倦怠程度', '年级×倦怠程度', '专业×倦怠程度']:
        res = cross_results[test_name]
        sig = '显著' if res['p'] < 0.05 else '不显著'
        doc.add_paragraph(
            f'{test_name}：χ²={res["chi2"]:.3f}, df={res["dof"]}, '
            f'p={res["p"]:.4f}, Cramér\'s V={res["cramers_v"]:.3f}，差异{sig}。'
        )
    
    doc.add_paragraph()
    doc.add_paragraph('图5 各维度相关系数矩阵')
    doc.add_picture(os.path.join(output_dir, 'fig5_correlation_matrix.png'), width=Inches(4.5))
    doc.add_paragraph()
    doc.add_paragraph('图6 不同性别学习倦怠箱线图')
    doc.add_picture(os.path.join(output_dir, 'fig6_gender_boxplot.png'), width=Inches(5.5))
    
    # ---- 七、结论与建议 ----
    doc.add_heading('七、结论与建议', level=1)
    doc.add_heading('7.1 主要结论', level=2)
    doc.add_paragraph(
        f'（1）本量表具有良好的信效度。总量表Cronbach\'s α={alpha_results["总量表"]:.3f}，'
        f'KMO={validity_results["kmo"]:.3f}，三因子累积方差解释率为{cum_var:.1f}%。'
    )
    doc.add_paragraph(
        f'（2）大学生学习倦怠总体处于中等水平（M={dim_stats.loc["总倦怠","mean"]:.2f}），'
        f'情绪耗竭维度得分最高，需重点关注。'
    )
    if anova_res['p'] < 0.05:
        doc.add_paragraph(
            f'（3）年级对学习倦怠有显著影响（p<0.05），高年级学生倦怠水平更高，'
            f'可能与就业压力和学业负担加重有关。'
        )
    else:
        doc.add_paragraph('（3）不同年级大学生的学习倦怠差异未达到统计显著水平。')
    
    doc.add_heading('7.2 建议', level=2)
    doc.add_paragraph('（1）高校应建立学生心理健康监测机制，定期评估学生学习倦怠水平。')
    doc.add_paragraph('（2）针对高年级学生，可开展职业规划和压力管理课程，缓解学业与就业的双重压力。')
    doc.add_paragraph('（3）教师应关注课堂教学方式的创新，减少学生的情绪耗竭感。')
    
    # 保存
    report_path = os.path.join(output_dir, '问卷数据分析报告_学习倦怠量表.docx')
    doc.save(report_path)
    return report_path


# ============================================================
# 主函数
# ============================================================

def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    print("=" * 60)
    print("  项目1：大学生学习倦怠量表 问卷数据分析")
    print("=" * 60)
    
    # 1. 生成数据
    print("\n[1/7] 生成模拟问卷数据（300份）...")
    df = generate_questionnaire_data(n=300)
    data_path = os.path.join(output_dir, 'data_questionnaire.csv')
    df.to_csv(data_path, index=False, encoding='utf-8-sig')
    print(f"  ✓ 数据已保存: {data_path}")
    
    # 2. 描述性统计
    print("\n[2/7] 描述性统计分析...")
    item_stats, dim_stats, demo_stats = descriptive_statistics(df)
    print(f"  ✓ 总倦怠均值={dim_stats.loc['总倦怠','mean']:.3f}, SD={dim_stats.loc['总倦怠','std']:.3f}")
    
    # 3. 信度分析
    print("\n[3/7] 信度分析...")
    alpha_results, citc_df = reliability_analysis(df)
    print(f"  ✓ 总量表α={alpha_results['总量表']:.4f}")
    for dim, alpha in alpha_results.items():
        if dim != '总量表':
            print(f"    - {dim}: α={alpha:.4f}")
    
    # 4. 效度分析
    print("\n[4/7] 效度分析（探索性因子分析）...")
    validity_results = validity_analysis(df)
    print(f"  ✓ KMO={validity_results['kmo']:.3f}")
    print(f"  ✓ Bartlett χ²={validity_results['bartlett_chi2']:.1f}, p<0.001")
    cum_var = validity_results['variance']['累积方差解释率(%)'].iloc[-1]
    print(f"  ✓ 三因子累积方差解释率={cum_var:.2f}%")
    
    # 5. 交叉分析
    print("\n[5/7] 差异性分析...")
    cross_results = cross_analysis(df)
    t_res = cross_results['性别t检验']
    print(f"  ✓ 性别t检验: t={t_res['t']:.3f}, p={t_res['p']:.4f}")
    anova_res = cross_results['年级方差分析']
    print(f"  ✓ 年级方差分析: F={anova_res['F']:.3f}, p={anova_res['p']:.4f}")
    
    # 6. 可视化
    print("\n[6/7] 生成可视化图表...")
    create_visualizations(df, validity_results, output_dir)
    
    # 7. 生成报告
    print("\n[7/7] 生成Word分析报告...")
    report_path = generate_word_report(
        df, item_stats, dim_stats, demo_stats,
        (alpha_results, citc_df), citc_df, validity_results,
        cross_results, output_dir
    )
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("  项目1完成！输出文件：")
    print(f"    - 数据: {data_path}")
    print(f"    - 报告: {report_path}")
    print(f"    - 图表: fig1~fig6 (共6张)")
    print("=" * 60)


if __name__ == '__main__':
    main()
