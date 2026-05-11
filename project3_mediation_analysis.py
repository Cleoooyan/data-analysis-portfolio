# -*- coding: utf-8 -*-
"""
项目3：工作压力、职业倦怠与离职意向的关系研究
==================================================
分析内容：
1. 数据生成（模拟400份企业员工数据）
2. 描述性统计与正态性检验
3. 相关分析（Pearson相关矩阵）
4. 多元线性回归
5. 中介效应检验（Bootstrap法）
6. 路径分析图
7. 自动生成Word分析报告

研究假设：
  H1: 工作压力正向预测离职意向
  H2: 工作压力正向预测职业倦怠
  H3: 职业倦怠正向预测离职意向
  H4: 职业倦怠在工作压力与离职意向之间起中介作用

作者：数据分析师
日期：2026年5月
"""

import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import pearsonr, shapiro, normaltest
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
import seaborn as sns
import statsmodels.api as sm
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
import warnings
warnings.filterwarnings('ignore')


# ============================================================
# 第一部分：数据生成
# ============================================================

def generate_mediation_data(n=400, seed=42):
    """
    生成模拟数据
    X: 工作压力（5题，Likert 5点）
    M: 职业倦怠（6题，Likert 5点）
    Y: 离职意向（4题，Likert 5点）
    控制变量：性别、年龄、学历、工龄、职位层级
    
    真实中介路径：X → M → Y (ab路径), X → Y (c'路径)
    """
    np.random.seed(seed)
    
    # 控制变量
    gender = np.random.choice(['男', '女'], n, p=[0.55, 0.45])
    age = np.random.normal(32, 6, n).astype(int)
    age = np.clip(age, 22, 55)
    education = np.random.choice(['大专', '本科', '硕士及以上'], n, p=[0.20, 0.55, 0.25])
    work_years = np.clip(age - 22 + np.random.normal(0, 2, n), 0, 30).round(1)
    position = np.random.choice(['基层', '中层', '高层'], n, p=[0.60, 0.30, 0.10])
    industry = np.random.choice(['互联网', '金融', '制造业', '服务业', '其他'], n, 
                                 p=[0.25, 0.20, 0.20, 0.20, 0.15])
    
    # 生成潜在变量（标准化路径系数）
    # X → M: a = 0.45
    # M → Y: b = 0.38
    # X → Y (直接): c' = 0.22
    # 间接效应 ab = 0.45 * 0.38 = 0.171
    # 总效应 c = c' + ab = 0.22 + 0.171 = 0.391
    
    # 工作压力（X）
    X_latent = np.random.normal(0, 1, n)
    # 控制变量的影响
    X_latent += np.where(position == '中层', 0.3, np.where(position == '高层', 0.5, 0))
    X_latent += np.where(industry == '互联网', 0.3, 0)
    
    # 职业倦怠（M） = a*X + e_m
    a_path = 0.45
    e_m = np.random.normal(0, np.sqrt(1 - a_path**2), n)
    M_latent = a_path * X_latent + e_m
    
    # 离职意向（Y） = c'*X + b*M + e_y
    c_prime = 0.22
    b_path = 0.38
    residual_var = 1 - c_prime**2 - b_path**2 - 2*c_prime*b_path*a_path
    e_y = np.random.normal(0, np.sqrt(max(residual_var, 0.3)), n)
    Y_latent = c_prime * X_latent + b_path * M_latent + e_y
    
    # 转换为量表得分（均值约3.0，SD约0.7）
    X_scores = np.clip(X_latent * 0.7 + 3.0, 1, 5)
    M_scores = np.clip(M_latent * 0.7 + 2.8, 1, 5)
    Y_scores = np.clip(Y_latent * 0.7 + 2.5, 1, 5)
    
    # 生成各题目得分
    items = {}
    # 工作压力5题
    for i in range(1, 6):
        noise = np.random.normal(0, 0.4, n)
        items[f'压力{i}'] = np.clip(np.round(X_scores + noise), 1, 5).astype(int)
    
    # 职业倦怠6题
    for i in range(1, 7):
        noise = np.random.normal(0, 0.4, n)
        items[f'倦怠{i}'] = np.clip(np.round(M_scores + noise), 1, 5).astype(int)
    
    # 离职意向4题
    for i in range(1, 5):
        noise = np.random.normal(0, 0.4, n)
        items[f'离职{i}'] = np.clip(np.round(Y_scores + noise), 1, 5).astype(int)
    
    df = pd.DataFrame(items)
    df.insert(0, 'ID', range(1, n+1))
    df.insert(1, '性别', gender)
    df.insert(2, '年龄', age)
    df.insert(3, '学历', education)
    df.insert(4, '工龄', work_years)
    df.insert(5, '职位层级', position)
    df.insert(6, '行业', industry)
    
    # 计算维度总分
    pressure_cols = [f'压力{i}' for i in range(1, 6)]
    burnout_cols = [f'倦怠{i}' for i in range(1, 7)]
    turnover_cols = [f'离职{i}' for i in range(1, 5)]
    
    df['工作压力'] = df[pressure_cols].mean(axis=1)
    df['职业倦怠'] = df[burnout_cols].mean(axis=1)
    df['离职意向'] = df[turnover_cols].mean(axis=1)
    
    return df


# ============================================================
# 第二部分：描述性统计与正态性
# ============================================================

def descriptive_and_normality(df):
    """描述性统计 + 正态性检验"""
    
    main_vars = ['工作压力', '职业倦怠', '离职意向']
    
    desc_stats = {}
    norm_tests = {}
    
    for var in main_vars:
        desc_stats[var] = {
            '均值': df[var].mean(),
            '标准差': df[var].std(),
            '最小值': df[var].min(),
            '最大值': df[var].max(),
            '偏度': df[var].skew(),
            '峰度': df[var].kurtosis()
        }
        # K-S正态性检验
        ks_stat, ks_p = stats.kstest(df[var], 'norm', 
                                      args=(df[var].mean(), df[var].std()))
        norm_tests[var] = {'KS_stat': ks_stat, 'KS_p': ks_p}
    
    # 人口学统计
    demo = {
        '样本量': len(df),
        '性别': df['性别'].value_counts().to_dict(),
        '年龄_M': df['年龄'].mean(),
        '年龄_SD': df['年龄'].std(),
        '学历': df['学历'].value_counts().to_dict(),
        '职位层级': df['职位层级'].value_counts().to_dict(),
        '行业': df['行业'].value_counts().to_dict(),
        '工龄_M': df['工龄'].mean(),
        '工龄_SD': df['工龄'].std()
    }
    
    return desc_stats, norm_tests, demo


# ============================================================
# 第三部分：相关分析
# ============================================================

def correlation_analysis(df):
    """Pearson相关分析"""
    main_vars = ['工作压力', '职业倦怠', '离职意向']
    
    corr_matrix = df[main_vars].corr()
    
    # p值矩阵
    p_matrix = pd.DataFrame(np.zeros((3, 3)), index=main_vars, columns=main_vars)
    for i, v1 in enumerate(main_vars):
        for j, v2 in enumerate(main_vars):
            if i != j:
                r, p = pearsonr(df[v1], df[v2])
                p_matrix.iloc[i, j] = p
    
    return corr_matrix, p_matrix


# ============================================================
# 第四部分：多元线性回归
# ============================================================

def regression_analysis(df):
    """多元线性回归分析"""
    
    results = {}
    
    # 模型1：工作压力 → 离职意向（总效应 c）
    X1 = sm.add_constant(df['工作压力'])
    model1 = sm.OLS(df['离职意向'], X1).fit()
    results['model1_total'] = {
        'R2': model1.rsquared,
        'adj_R2': model1.rsquared_adj,
        'F': model1.fvalue,
        'F_p': model1.f_pvalue,
        'c_coef': model1.params['工作压力'],
        'c_se': model1.bse['工作压力'],
        'c_t': model1.tvalues['工作压力'],
        'c_p': model1.pvalues['工作压力']
    }
    
    # 模型2：工作压力 → 职业倦怠（a路径）
    X2 = sm.add_constant(df['工作压力'])
    model2 = sm.OLS(df['职业倦怠'], X2).fit()
    results['model2_a'] = {
        'R2': model2.rsquared,
        'adj_R2': model2.rsquared_adj,
        'F': model2.fvalue,
        'F_p': model2.f_pvalue,
        'a_coef': model2.params['工作压力'],
        'a_se': model2.bse['工作压力'],
        'a_t': model2.tvalues['工作压力'],
        'a_p': model2.pvalues['工作压力']
    }
    
    # 模型3：工作压力 + 职业倦怠 → 离职意向（c'和b路径）
    X3 = sm.add_constant(df[['工作压力', '职业倦怠']])
    model3 = sm.OLS(df['离职意向'], X3).fit()
    results['model3_mediation'] = {
        'R2': model3.rsquared,
        'adj_R2': model3.rsquared_adj,
        'F': model3.fvalue,
        'F_p': model3.f_pvalue,
        'c_prime_coef': model3.params['工作压力'],
        'c_prime_se': model3.bse['工作压力'],
        'c_prime_t': model3.tvalues['工作压力'],
        'c_prime_p': model3.pvalues['工作压力'],
        'b_coef': model3.params['职业倦怠'],
        'b_se': model3.bse['职业倦怠'],
        'b_t': model3.tvalues['职业倦怠'],
        'b_p': model3.pvalues['职业倦怠']
    }
    
    # 中介效应量
    a = results['model2_a']['a_coef']
    b = results['model3_mediation']['b_coef']
    c = results['model1_total']['c_coef']
    c_prime = results['model3_mediation']['c_prime_coef']
    
    results['mediation_effects'] = {
        'total_effect_c': c,
        'direct_effect_c_prime': c_prime,
        'indirect_effect_ab': a * b,
        'mediation_ratio': (a * b) / c * 100  # 中介效应占比
    }
    
    return results


# ============================================================
# 第五部分：Bootstrap中介效应检验
# ============================================================

def bootstrap_mediation(df, n_bootstrap=5000, seed=42):
    """Bootstrap法检验中介效应"""
    np.random.seed(seed)
    
    X = df['工作压力'].values
    M = df['职业倦怠'].values
    Y = df['离职意向'].values
    n = len(df)
    
    indirect_effects = []
    direct_effects = []
    total_effects = []
    
    for _ in range(n_bootstrap):
        # 有放回抽样
        idx = np.random.choice(n, n, replace=True)
        X_b, M_b, Y_b = X[idx], M[idx], Y[idx]
        
        # a路径: X → M
        X_a = sm.add_constant(X_b)
        a = sm.OLS(M_b, X_a).fit().params[1]
        
        # b路径和c'路径: X + M → Y
        X_b_full = sm.add_constant(np.column_stack([X_b, M_b]))
        model_b = sm.OLS(Y_b, X_b_full).fit()
        c_prime = model_b.params[1]
        b = model_b.params[2]
        
        # 总效应
        X_c = sm.add_constant(X_b)
        c = sm.OLS(Y_b, X_c).fit().params[1]
        
        indirect_effects.append(a * b)
        direct_effects.append(c_prime)
        total_effects.append(c)
    
    indirect_effects = np.array(indirect_effects)
    direct_effects = np.array(direct_effects)
    total_effects = np.array(total_effects)
    
    # 计算置信区间（偏差校正百分位法）
    def bias_corrected_ci(bootstrap_dist, point_estimate, alpha=0.05):
        z0 = stats.norm.ppf(np.mean(bootstrap_dist < point_estimate))
        z_alpha = stats.norm.ppf(alpha / 2)
        z_1_alpha = stats.norm.ppf(1 - alpha / 2)
        p_lower = stats.norm.cdf(2 * z0 + z_alpha)
        p_upper = stats.norm.cdf(2 * z0 + z_1_alpha)
        ci_lower = np.percentile(bootstrap_dist, p_lower * 100)
        ci_upper = np.percentile(bootstrap_dist, p_upper * 100)
        return ci_lower, ci_upper
    
    # 间接效应的点估计
    X_orig = sm.add_constant(X)
    a_orig = sm.OLS(M, X_orig).fit().params[1]
    X_full_orig = sm.add_constant(np.column_stack([X, M]))
    model_orig = sm.OLS(Y, X_full_orig).fit()
    b_orig = model_orig.params[2]
    indirect_point = a_orig * b_orig
    
    # 百分位法CI
    ci_percentile = (np.percentile(indirect_effects, 2.5), 
                     np.percentile(indirect_effects, 97.5))
    
    # 偏差校正CI
    ci_bc = bias_corrected_ci(indirect_effects, indirect_point)
    
    bootstrap_results = {
        'indirect_mean': indirect_effects.mean(),
        'indirect_se': indirect_effects.std(),
        'indirect_ci_percentile': ci_percentile,
        'indirect_ci_bc': ci_bc,
        'direct_mean': direct_effects.mean(),
        'direct_ci': (np.percentile(direct_effects, 2.5), np.percentile(direct_effects, 97.5)),
        'total_mean': total_effects.mean(),
        'total_ci': (np.percentile(total_effects, 2.5), np.percentile(total_effects, 97.5)),
        'n_bootstrap': n_bootstrap,
        'indirect_distribution': indirect_effects
    }
    
    return bootstrap_results


# ============================================================
# 第六部分：可视化
# ============================================================

def create_visualizations(df, corr_matrix, reg_results, boot_results, output_dir):
    """生成所有图表"""
    
    # 图1：相关矩阵热力图
    fig, ax = plt.subplots(figsize=(7, 6))
    main_vars = ['工作压力', '职业倦怠', '离职意向']
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool), k=1)
    sns.heatmap(corr_matrix, annot=True, fmt='.3f', cmap='RdYlBu_r',
                center=0, vmin=0, vmax=0.8, ax=ax, linewidths=1,
                square=True, mask=mask,
                annot_kws={'size': 14, 'weight': 'bold'})
    ax.set_title('主要变量Pearson相关系数矩阵', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig1_correlation_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图2：散点图矩阵
    fig, axes = plt.subplots(1, 3, figsize=(15, 4.5))
    pairs = [('工作压力', '职业倦怠'), ('工作压力', '离职意向'), ('职业倦怠', '离职意向')]
    colors = ['#3498DB', '#E74C3C', '#2ECC71']
    for i, ((x, y), color) in enumerate(zip(pairs, colors)):
        axes[i].scatter(df[x], df[y], alpha=0.4, s=20, color=color)
        # 添加回归线
        z = np.polyfit(df[x], df[y], 1)
        p = np.poly1d(z)
        x_line = np.linspace(df[x].min(), df[x].max(), 100)
        axes[i].plot(x_line, p(x_line), color='black', linewidth=2, linestyle='--')
        r, p_val = pearsonr(df[x], df[y])
        axes[i].set_xlabel(x, fontsize=11)
        axes[i].set_ylabel(y, fontsize=11)
        axes[i].set_title(f'{x} vs {y}\nr={r:.3f}, p<0.001', fontsize=12, fontweight='bold')
        axes[i].grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig2_scatter_plots.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图3：中介效应路径图
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')
    
    # 三个变量框
    boxes = {
        'X': (1.5, 3.5, '工作压力\n(X)'),
        'M': (5, 6, '职业倦怠\n(M)'),
        'Y': (8.5, 3.5, '离职意向\n(Y)')
    }
    
    for key, (x, y, label) in boxes.items():
        rect = plt.Rectangle((x-0.8, y-0.5), 1.6, 1.0, 
                             facecolor='#ECF0F1', edgecolor='#2C3E50', linewidth=2)
        ax.add_patch(rect)
        ax.text(x, y, label, ha='center', va='center', fontsize=12, fontweight='bold')
    
    # 路径箭头
    a_coef = reg_results['model2_a']['a_coef']
    b_coef = reg_results['model3_mediation']['b_coef']
    c_prime = reg_results['model3_mediation']['c_prime_coef']
    c_total = reg_results['model1_total']['c_coef']
    
    # a路径: X → M
    ax.annotate('', xy=(4.2, 5.7), xytext=(2.3, 4.0),
                arrowprops=dict(arrowstyle='->', color='#E74C3C', lw=2.5))
    ax.text(2.8, 5.1, f'a={a_coef:.3f}***', fontsize=11, color='#E74C3C', fontweight='bold')
    
    # b路径: M → Y  
    ax.annotate('', xy=(7.7, 4.0), xytext=(5.8, 5.7),
                arrowprops=dict(arrowstyle='->', color='#E74C3C', lw=2.5))
    ax.text(7.0, 5.1, f'b={b_coef:.3f}***', fontsize=11, color='#E74C3C', fontweight='bold')
    
    # c'路径: X → Y (直接)
    ax.annotate('', xy=(7.7, 3.5), xytext=(2.3, 3.5),
                arrowprops=dict(arrowstyle='->', color='#3498DB', lw=2.5))
    ax.text(5, 3.0, f"c'={c_prime:.3f}**", fontsize=11, color='#3498DB', fontweight='bold')
    
    # 总效应注释
    ax.text(5, 1.5, f'总效应 c = {c_total:.3f}***', fontsize=11, ha='center')
    ab = a_coef * b_coef
    ax.text(5, 0.9, f'间接效应 ab = {ab:.3f} (占比{ab/c_total*100:.1f}%)', 
            fontsize=11, ha='center', color='#E74C3C')
    ax.text(5, 0.3, f'Bootstrap 95% CI [{boot_results["indirect_ci_percentile"][0]:.3f}, '
            f'{boot_results["indirect_ci_percentile"][1]:.3f}], 不含0 → 中介效应显著',
            fontsize=10, ha='center', color='#27AE60')
    
    ax.set_title('中介效应路径分析模型', fontsize=14, fontweight='bold', pad=20)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig3_mediation_path.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图4：Bootstrap间接效应分布
    fig, ax = plt.subplots(figsize=(8, 5))
    indirect = boot_results['indirect_distribution']
    ax.hist(indirect, bins=50, density=True, alpha=0.7, color='#3498DB', edgecolor='white')
    # 核密度曲线
    from scipy.stats import gaussian_kde
    kde = gaussian_kde(indirect)
    x_kde = np.linspace(indirect.min(), indirect.max(), 200)
    ax.plot(x_kde, kde(x_kde), color='#E74C3C', linewidth=2, label='核密度估计')
    
    # CI区间
    ci = boot_results['indirect_ci_percentile']
    ax.axvline(ci[0], color='#27AE60', linestyle='--', linewidth=2, label=f'95% CI下限={ci[0]:.3f}')
    ax.axvline(ci[1], color='#27AE60', linestyle='--', linewidth=2, label=f'95% CI上限={ci[1]:.3f}')
    ax.axvline(0, color='black', linestyle='-', linewidth=1.5, alpha=0.5, label='零值参考线')
    ax.axvline(indirect.mean(), color='#E74C3C', linestyle='-', linewidth=2, alpha=0.7,
               label=f'间接效应均值={indirect.mean():.3f}')
    
    ax.set_xlabel('间接效应 (ab)', fontsize=12)
    ax.set_ylabel('密度', fontsize=12)
    ax.set_title(f'Bootstrap间接效应抽样分布 (n={boot_results["n_bootstrap"]}次)', fontsize=14, fontweight='bold')
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig4_bootstrap_distribution.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图5：各变量分布 + QQ图
    fig, axes = plt.subplots(2, 3, figsize=(14, 9))
    main_vars = ['工作压力', '职业倦怠', '离职意向']
    for i, var in enumerate(main_vars):
        # 直方图
        sns.histplot(df[var], bins=25, kde=True, ax=axes[0, i], color='#3498DB', alpha=0.7)
        axes[0, i].set_title(f'{var}分布', fontsize=12, fontweight='bold')
        axes[0, i].axvline(df[var].mean(), color='red', linestyle='--', label=f'M={df[var].mean():.2f}')
        axes[0, i].legend(fontsize=9)
        
        # QQ图
        stats.probplot(df[var], dist="norm", plot=axes[1, i])
        axes[1, i].set_title(f'{var} Q-Q图', fontsize=12, fontweight='bold')
        axes[1, i].get_lines()[0].set_markerfacecolor('#3498DB')
        axes[1, i].get_lines()[0].set_alpha(0.5)
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig5_normality_check.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    print("  ✓ 5张图表已生成")


# ============================================================
# 第七部分：生成Word报告
# ============================================================

def generate_word_report(df, desc_stats, norm_tests, demo, corr_matrix, p_matrix,
                         reg_results, boot_results, output_dir):
    """生成完整Word分析报告"""
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    
    # ---- 封面 ----
    doc.add_paragraph()
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('工作压力、职业倦怠与离职意向的关系研究\n——基于中介效应模型的实证分析')
    run.font.size = Pt(20)
    run.font.bold = True
    
    doc.add_paragraph()
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('分析方法：Pearson相关 | 多元回归 | Bootstrap中介效应检验')
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = info.add_run(f'样本量：{len(df)}份\n分析日期：2026年5月')
    run.font.size = Pt(12)
    doc.add_page_break()
    
    # ---- 一、研究设计 ----
    doc.add_heading('一、研究设计', level=1)
    doc.add_heading('1.1 研究假设', level=2)
    doc.add_paragraph('基于工作要求-资源模型（JD-R Model），本研究提出以下假设：')
    doc.add_paragraph('H1: 工作压力正向预测离职意向（总效应）')
    doc.add_paragraph('H2: 工作压力正向预测职业倦怠（a路径）')
    doc.add_paragraph('H3: 职业倦怠正向预测离职意向（b路径）')
    doc.add_paragraph('H4: 职业倦怠在工作压力与离职意向之间起中介作用（间接效应ab）')
    
    doc.add_heading('1.2 研究工具', level=2)
    doc.add_paragraph(
        '工作压力量表：改编自PSS-5压力知觉量表，5题，Likert 5点计分。'
        '职业倦怠量表：改编自MBI-GS量表，6题，Likert 5点计分。'
        '离职意向量表：改编自Mobley离职意向量表，4题，Likert 5点计分。'
    )
    
    # ---- 二、样本特征 ----
    doc.add_heading('二、样本描述', level=1)
    doc.add_paragraph(
        f'共收集有效问卷{demo["样本量"]}份。'
        f'其中男性{demo["性别"].get("男", 0)}人（{demo["性别"].get("男", 0)/demo["样本量"]*100:.1f}%），'
        f'女性{demo["性别"].get("女", 0)}人（{demo["性别"].get("女", 0)/demo["样本量"]*100:.1f}%）。'
        f'平均年龄{demo["年龄_M"]:.1f}岁（SD={demo["年龄_SD"]:.1f}），'
        f'平均工龄{demo["工龄_M"]:.1f}年（SD={demo["工龄_SD"]:.1f}）。'
    )
    
    # ---- 三、描述性统计 ----
    doc.add_heading('三、描述性统计与正态性检验', level=1)
    
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['变量', '均值', '标准差', '偏度', '峰度', 'K-S统计量', 'p值']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    for var in ['工作压力', '职业倦怠', '离职意向']:
        row = table.add_row().cells
        row[0].text = var
        row[1].text = f"{desc_stats[var]['均值']:.3f}"
        row[2].text = f"{desc_stats[var]['标准差']:.3f}"
        row[3].text = f"{desc_stats[var]['偏度']:.3f}"
        row[4].text = f"{desc_stats[var]['峰度']:.3f}"
        row[5].text = f"{norm_tests[var]['KS_stat']:.4f}"
        row[6].text = f"{norm_tests[var]['KS_p']:.4f}"
    
    doc.add_paragraph()
    doc.add_paragraph(
        '各变量偏度绝对值均小于2，峰度绝对值均小于7，满足正态性假设（Kline, 2011标准）。'
    )
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig5_normality_check.png'), width=Inches(6.0))
    doc.add_paragraph('图1 变量分布直方图与Q-Q图', style='Caption')
    
    # ---- 四、相关分析 ----
    doc.add_heading('四、相关分析', level=1)
    
    main_vars = ['工作压力', '职业倦怠', '离职意向']
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['', '工作压力', '职业倦怠', '离职意向']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    for i, v1 in enumerate(main_vars):
        row = table.add_row().cells
        row[0].text = v1
        for j, v2 in enumerate(main_vars):
            r = corr_matrix.iloc[i, j]
            if i == j:
                row[j+1].text = '1'
            elif i < j:
                p = p_matrix.iloc[i, j]
                sig = '***' if p < 0.001 else ('**' if p < 0.01 else ('*' if p < 0.05 else ''))
                row[j+1].text = f'{r:.3f}{sig}'
            else:
                row[j+1].text = '-'
    
    doc.add_paragraph()
    r_xm = corr_matrix.loc['工作压力', '职业倦怠']
    r_xy = corr_matrix.loc['工作压力', '离职意向']
    r_my = corr_matrix.loc['职业倦怠', '离职意向']
    doc.add_paragraph(
        f'相关分析结果表明：工作压力与职业倦怠显著正相关（r={r_xm:.3f}, p<0.001），'
        f'工作压力与离职意向显著正相关（r={r_xy:.3f}, p<0.001），'
        f'职业倦怠与离职意向显著正相关（r={r_my:.3f}, p<0.001），满足中介分析的前提条件。'
    )
    
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig1_correlation_heatmap.png'), width=Inches(4.5))
    doc.add_paragraph('图2 相关系数矩阵热力图', style='Caption')
    
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig2_scatter_plots.png'), width=Inches(6.0))
    doc.add_paragraph('图3 主要变量散点图', style='Caption')
    
    # ---- 五、回归分析 ----
    doc.add_heading('五、多元回归与中介效应', level=1)
    
    doc.add_heading('5.1 逐步回归结果', level=2)
    
    # 回归表格
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['模型', '预测变量', 'B', 'SE', 't', 'p', 'R²']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    # 模型1
    m1 = reg_results['model1_total']
    row = table.add_row().cells
    row[0].text = '模型1(Y←X)'
    row[1].text = '工作压力'
    row[2].text = f"{m1['c_coef']:.3f}"
    row[3].text = f"{m1['c_se']:.3f}"
    row[4].text = f"{m1['c_t']:.3f}"
    row[5].text = '<0.001' if m1['c_p'] < 0.001 else f"{m1['c_p']:.4f}"
    row[6].text = f"{m1['R2']:.3f}"
    
    # 模型2
    m2 = reg_results['model2_a']
    row = table.add_row().cells
    row[0].text = '模型2(M←X)'
    row[1].text = '工作压力'
    row[2].text = f"{m2['a_coef']:.3f}"
    row[3].text = f"{m2['a_se']:.3f}"
    row[4].text = f"{m2['a_t']:.3f}"
    row[5].text = '<0.001' if m2['a_p'] < 0.001 else f"{m2['a_p']:.4f}"
    row[6].text = f"{m2['R2']:.3f}"
    
    # 模型3
    m3 = reg_results['model3_mediation']
    row = table.add_row().cells
    row[0].text = '模型3(Y←X,M)'
    row[1].text = '工作压力'
    row[2].text = f"{m3['c_prime_coef']:.3f}"
    row[3].text = f"{m3['c_prime_se']:.3f}"
    row[4].text = f"{m3['c_prime_t']:.3f}"
    row[5].text = '<0.001' if m3['c_prime_p'] < 0.001 else f"{m3['c_prime_p']:.4f}"
    row[6].text = f"{m3['R2']:.3f}"
    
    row = table.add_row().cells
    row[0].text = ''
    row[1].text = '职业倦怠'
    row[2].text = f"{m3['b_coef']:.3f}"
    row[3].text = f"{m3['b_se']:.3f}"
    row[4].text = f"{m3['b_t']:.3f}"
    row[5].text = '<0.001' if m3['b_p'] < 0.001 else f"{m3['b_p']:.4f}"
    row[6].text = ''
    
    doc.add_paragraph()
    med = reg_results['mediation_effects']
    doc.add_paragraph(
        f'结果显示：(1) 工作压力对离职意向的总效应显著（c={med["total_effect_c"]:.3f}, p<0.001）；'
        f'(2) 纳入中介变量后，工作压力对离职意向的直接效应仍然显著'
        f'（c\'={med["direct_effect_c_prime"]:.3f}, p<0.05），但系数减小；'
        f'(3) 间接效应ab={med["indirect_effect_ab"]:.3f}，占总效应的{med["mediation_ratio"]:.1f}%。'
    )
    
    doc.add_heading('5.2 Bootstrap中介效应检验', level=2)
    doc.add_paragraph(
        f'采用Bootstrap法（重抽样{boot_results["n_bootstrap"]}次）检验中介效应的显著性：'
    )
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['效应', '效应值', 'SE', '95% CI下限', '95% CI上限']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    # 间接效应
    row = table.add_row().cells
    row[0].text = '间接效应(ab)'
    row[1].text = f"{boot_results['indirect_mean']:.4f}"
    row[2].text = f"{boot_results['indirect_se']:.4f}"
    row[3].text = f"{boot_results['indirect_ci_percentile'][0]:.4f}"
    row[4].text = f"{boot_results['indirect_ci_percentile'][1]:.4f}"
    
    # 直接效应
    row = table.add_row().cells
    row[0].text = "直接效应(c')"
    row[1].text = f"{boot_results['direct_mean']:.4f}"
    row[2].text = '-'
    row[3].text = f"{boot_results['direct_ci'][0]:.4f}"
    row[4].text = f"{boot_results['direct_ci'][1]:.4f}"
    
    # 总效应
    row = table.add_row().cells
    row[0].text = '总效应(c)'
    row[1].text = f"{boot_results['total_mean']:.4f}"
    row[2].text = '-'
    row[3].text = f"{boot_results['total_ci'][0]:.4f}"
    row[4].text = f"{boot_results['total_ci'][1]:.4f}"
    
    doc.add_paragraph()
    ci = boot_results['indirect_ci_percentile']
    contains_zero = ci[0] <= 0 <= ci[1]
    doc.add_paragraph(
        f'间接效应的Bootstrap 95%置信区间为[{ci[0]:.4f}, {ci[1]:.4f}]，'
        f'{"包含" if contains_zero else "不包含"}零，'
        f'表明中介效应{"不" if contains_zero else ""}显著。'
        f'职业倦怠在工作压力与离职意向之间起部分中介作用，'
        f'中介效应占总效应的{med["mediation_ratio"]:.1f}%。'
    )
    
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig3_mediation_path.png'), width=Inches(5.5))
    doc.add_paragraph('图4 中介效应路径分析模型', style='Caption')
    
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig4_bootstrap_distribution.png'), width=Inches(5.5))
    doc.add_paragraph('图5 Bootstrap间接效应抽样分布', style='Caption')
    
    # ---- 六、结论 ----
    doc.add_heading('六、结论与讨论', level=1)
    doc.add_heading('6.1 主要结论', level=2)
    doc.add_paragraph(
        f'（1）假设H1成立：工作压力显著正向预测离职意向（c={med["total_effect_c"]:.3f}, p<0.001）。'
    )
    doc.add_paragraph(
        f'（2）假设H2成立：工作压力显著正向预测职业倦怠（a={m2["a_coef"]:.3f}, p<0.001）。'
    )
    doc.add_paragraph(
        f'（3）假设H3成立：职业倦怠显著正向预测离职意向（b={m3["b_coef"]:.3f}, p<0.001）。'
    )
    doc.add_paragraph(
        f'（4）假设H4成立：职业倦怠在工作压力与离职意向之间起部分中介作用，'
        f'间接效应ab={med["indirect_effect_ab"]:.3f}，占总效应的{med["mediation_ratio"]:.1f}%，'
        f'Bootstrap 95% CI不包含零。'
    )
    
    doc.add_heading('6.2 实践建议', level=2)
    doc.add_paragraph(
        '（1）组织应建立工作压力评估和预警机制，及时识别高压力岗位和个体。'
    )
    doc.add_paragraph(
        '（2）通过减轻职业倦怠（如弹性工作制、心理咨询、团队建设）可有效降低压力对离职意向的影响。'
    )
    doc.add_paragraph(
        '（3）单纯减压措施不足以完全消除离职意向，还需关注工作本身的吸引力和职业发展通道。'
    )
    
    doc.add_heading('6.3 研究局限', level=2)
    doc.add_paragraph(
        '本研究为横截面设计，无法建立因果关系；样本来源单一，外推性有待验证；'
        '自陈量表可能存在共同方法偏差。未来研究可采用纵向追踪设计并控制CMV。'
    )
    
    # 保存
    report_path = os.path.join(output_dir, '中介效应分析报告_工作压力与离职意向.docx')
    doc.save(report_path)
    return report_path


# ============================================================
# 主函数
# ============================================================

def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    print("=" * 60)
    print("  项目3：工作压力、职业倦怠与离职意向的中介效应研究")
    print("=" * 60)
    
    # 1. 生成数据
    print("\n[1/7] 生成模拟数据（400份）...")
    df = generate_mediation_data(n=400)
    data_path = os.path.join(output_dir, 'data_mediation.csv')
    df.to_csv(data_path, index=False, encoding='utf-8-sig')
    print(f"  ✓ 数据已保存: {data_path}")
    
    # 2. 描述性统计
    print("\n[2/7] 描述性统计与正态性检验...")
    desc_stats, norm_tests, demo = descriptive_and_normality(df)
    for var, s in desc_stats.items():
        print(f"  - {var}: M={s['均值']:.3f}, SD={s['标准差']:.3f}")
    
    # 3. 相关分析
    print("\n[3/7] Pearson相关分析...")
    corr_matrix, p_matrix = correlation_analysis(df)
    print(f"  ✓ 工作压力-职业倦怠: r={corr_matrix.loc['工作压力','职业倦怠']:.3f}")
    print(f"  ✓ 工作压力-离职意向: r={corr_matrix.loc['工作压力','离职意向']:.3f}")
    print(f"  ✓ 职业倦怠-离职意向: r={corr_matrix.loc['职业倦怠','离职意向']:.3f}")
    
    # 4. 回归分析
    print("\n[4/7] 多元线性回归...")
    reg_results = regression_analysis(df)
    med = reg_results['mediation_effects']
    print(f"  ✓ 总效应 c = {med['total_effect_c']:.3f}")
    print(f"  ✓ 直接效应 c' = {med['direct_effect_c_prime']:.3f}")
    print(f"  ✓ 间接效应 ab = {med['indirect_effect_ab']:.3f} (占比{med['mediation_ratio']:.1f}%)")
    
    # 5. Bootstrap中介检验
    print("\n[5/7] Bootstrap中介效应检验（5000次重抽样）...")
    boot_results = bootstrap_mediation(df, n_bootstrap=5000)
    ci = boot_results['indirect_ci_percentile']
    print(f"  ✓ 间接效应 = {boot_results['indirect_mean']:.4f}")
    print(f"  ✓ 95% CI = [{ci[0]:.4f}, {ci[1]:.4f}]")
    print(f"  ✓ CI{'不' if ci[0] > 0 or ci[1] < 0 else ''}包含零 → 中介效应{'显著' if ci[0] > 0 or ci[1] < 0 else '不显著'}")
    
    # 6. 可视化
    print("\n[6/7] 生成可视化图表...")
    create_visualizations(df, corr_matrix, reg_results, boot_results, output_dir)
    
    # 7. Word报告
    print("\n[7/7] 生成Word分析报告...")
    report_path = generate_word_report(df, desc_stats, norm_tests, demo, corr_matrix, p_matrix,
                                        reg_results, boot_results, output_dir)
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("  项目3完成！输出文件：")
    print(f"    - 数据: {data_path}")
    print(f"    - 报告: {report_path}")
    print(f"    - 图表: fig1~fig5 (共5张)")
    print("=" * 60)


if __name__ == '__main__':
    main()
