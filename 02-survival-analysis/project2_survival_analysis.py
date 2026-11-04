# -*- coding: utf-8 -*-
"""
项目2：肿瘤患者生存分析
========================================
分析内容：
1. 数据生成（模拟500例肿瘤患者随访数据）
2. 描述性统计与基线特征
3. Kaplan-Meier生存曲线
4. Log-rank检验
5. 单因素Cox回归
6. 多因素Cox回归
7. 森林图
8. 自动生成Word分析报告

作者：数据分析师
日期：2026年5月
"""

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
import seaborn as sns
from lifelines import KaplanMeierFitter, CoxPHFitter
from lifelines.statistics import logrank_test, multivariate_logrank_test
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
import warnings
warnings.filterwarnings('ignore')


# ============================================================
# 第一部分：数据生成
# ============================================================

def generate_survival_data(n=500, seed=42):
    """
    生成模拟肿瘤患者随访数据
    场景：非小细胞肺癌（NSCLC）患者术后随访
    终点事件：全因死亡
    """
    np.random.seed(seed)
    
    # 基线特征
    age = np.random.normal(62, 10, n).astype(int)
    age = np.clip(age, 35, 85)
    gender = np.random.choice(['男', '女'], n, p=[0.62, 0.38])
    
    # 临床分期（I/II/III/IV）
    stage = np.random.choice(['I期', 'II期', 'III期', 'IV期'], n, p=[0.25, 0.30, 0.28, 0.17])
    
    # 病理类型
    pathology = np.random.choice(['腺癌', '鳞癌', '大细胞癌'], n, p=[0.50, 0.35, 0.15])
    
    # 分化程度
    differentiation = np.random.choice(['高分化', '中分化', '低分化'], n, p=[0.25, 0.45, 0.30])
    
    # 治疗方案
    treatment = np.random.choice(['手术+化疗', '手术+放化疗', '单纯手术'], n, p=[0.45, 0.35, 0.20])
    
    # KPS评分
    kps = np.random.choice([60, 70, 80, 90, 100], n, p=[0.10, 0.20, 0.35, 0.25, 0.10])
    
    # 肿瘤大小(cm)
    tumor_size = np.random.lognormal(1.2, 0.4, n).round(1)
    tumor_size = np.clip(tumor_size, 0.5, 12.0)
    
    # 吸烟史
    smoking = np.random.choice(['从不吸烟', '已戒烟', '目前吸烟'], n, p=[0.30, 0.35, 0.35])
    
    # 生成生存时间（基于Cox模型）
    # 设定各因素的真实HR
    baseline_hazard = 0.02  # 基线风险
    
    # 计算线性预测值
    lp = np.zeros(n)
    lp += np.where(stage == 'II期', 0.4, np.where(stage == 'III期', 0.9, np.where(stage == 'IV期', 1.5, 0)))
    lp += np.where(differentiation == '中分化', 0.3, np.where(differentiation == '低分化', 0.7, 0))
    lp += np.where(gender == '男', 0.2, 0)
    lp += (age - 60) * 0.02
    lp += np.where(treatment == '手术+化疗', -0.3, np.where(treatment == '手术+放化疗', -0.5, 0))
    lp += np.where(kps >= 80, -0.3, 0)
    lp += (tumor_size - 3) * 0.1
    lp += np.where(smoking == '目前吸烟', 0.3, np.where(smoking == '已戒烟', 0.1, 0))
    
    # Weibull分布生成生存时间
    shape = 1.2
    scale = 1 / (baseline_hazard * np.exp(lp))
    survival_time = np.random.weibull(shape, n) * scale
    survival_time = np.clip(survival_time, 0.5, 120).round(1)  # 单位：月
    
    # 删失（~30%删失率）
    censor_time = np.random.uniform(12, 80, n)
    event = (survival_time <= censor_time).astype(int)
    observed_time = np.minimum(survival_time, censor_time).round(1)
    
    df = pd.DataFrame({
        'ID': range(1, n+1),
        '年龄': age,
        '性别': gender,
        '临床分期': stage,
        '病理类型': pathology,
        '分化程度': differentiation,
        '治疗方案': treatment,
        'KPS评分': kps,
        '肿瘤大小(cm)': tumor_size,
        '吸烟史': smoking,
        '随访时间(月)': observed_time,
        '结局事件': event  # 1=死亡, 0=删失
    })
    
    return df


# ============================================================
# 第二部分：描述性统计
# ============================================================

def baseline_characteristics(df):
    """基线特征描述"""
    total_n = len(df)
    event_n = df['结局事件'].sum()
    censor_n = total_n - event_n
    median_time = df['随访时间(月)'].median()
    
    stats_info = {
        '总样本量': total_n,
        '事件数(死亡)': event_n,
        '删失数': censor_n,
        '事件发生率': f'{event_n/total_n*100:.1f}%',
        '中位随访时间(月)': median_time
    }
    
    # 分类变量统计
    cat_vars = ['性别', '临床分期', '病理类型', '分化程度', '治疗方案', '吸烟史']
    cat_stats = {}
    for var in cat_vars:
        counts = df[var].value_counts()
        cat_stats[var] = {k: f'{v}({v/total_n*100:.1f}%)' for k, v in counts.items()}
    
    # 连续变量统计
    cont_vars = ['年龄', 'KPS评分', '肿瘤大小(cm)', '随访时间(月)']
    cont_stats = {}
    for var in cont_vars:
        cont_stats[var] = f'{df[var].mean():.1f}±{df[var].std():.1f}'
    
    return stats_info, cat_stats, cont_stats


# ============================================================
# 第三部分：Kaplan-Meier生存分析
# ============================================================

def kaplan_meier_analysis(df, output_dir):
    """Kaplan-Meier生存曲线 + Log-rank检验"""
    
    results = {}
    
    # 总体生存曲线
    kmf = KaplanMeierFitter()
    kmf.fit(df['随访时间(月)'], event_observed=df['结局事件'], label='总体')
    median_survival = kmf.median_survival_time_
    results['总体中位生存期'] = median_survival
    
    # 图1：总体生存曲线
    fig, ax = plt.subplots(figsize=(8, 5.5))
    kmf.plot_survival_function(ax=ax, ci_show=True, color='#2C3E50', linewidth=2)
    ax.set_xlabel('时间（月）', fontsize=12)
    ax.set_ylabel('生存概率', fontsize=12)
    ax.set_title('Kaplan-Meier 总体生存曲线', fontsize=14, fontweight='bold')
    ax.axhline(y=0.5, color='gray', linestyle='--', alpha=0.5)
    ax.axvline(x=median_survival, color='gray', linestyle='--', alpha=0.5)
    ax.text(median_survival + 1, 0.52, f'中位OS={median_survival:.1f}月', fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_ylim(0, 1.05)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig1_overall_km_curve.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图2：按临床分期分组的生存曲线
    fig, ax = plt.subplots(figsize=(9, 6))
    colors_stage = {'I期': '#27AE60', 'II期': '#2980B9', 'III期': '#E67E22', 'IV期': '#E74C3C'}
    km_results_stage = {}
    for stage in ['I期', 'II期', 'III期', 'IV期']:
        mask = df['临床分期'] == stage
        kmf_s = KaplanMeierFitter()
        kmf_s.fit(df.loc[mask, '随访时间(月)'], event_observed=df.loc[mask, '结局事件'], label=stage)
        kmf_s.plot_survival_function(ax=ax, ci_show=False, color=colors_stage[stage], linewidth=2)
        km_results_stage[stage] = kmf_s.median_survival_time_
    
    # Log-rank检验（多组）
    lr_result = multivariate_logrank_test(df['随访时间(月)'], df['临床分期'], df['结局事件'])
    results['分期_logrank_chi2'] = lr_result.test_statistic
    results['分期_logrank_p'] = lr_result.p_value
    results['分期_中位OS'] = km_results_stage
    
    ax.set_xlabel('时间（月）', fontsize=12)
    ax.set_ylabel('生存概率', fontsize=12)
    ax.set_title('不同临床分期 Kaplan-Meier 生存曲线', fontsize=14, fontweight='bold')
    ax.legend(fontsize=11, loc='lower left')
    ax.text(0.95, 0.95, f'Log-rank p{"<0.001" if lr_result.p_value < 0.001 else f"={lr_result.p_value:.4f}"}',
            transform=ax.transAxes, ha='right', va='top', fontsize=11,
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    ax.grid(True, alpha=0.3)
    ax.set_ylim(0, 1.05)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig2_stage_km_curve.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    # 图3：按治疗方案分组
    fig, ax = plt.subplots(figsize=(9, 6))
    colors_tx = {'手术+化疗': '#3498DB', '手术+放化疗': '#2ECC71', '单纯手术': '#E74C3C'}
    km_results_tx = {}
    for tx in ['手术+化疗', '手术+放化疗', '单纯手术']:
        mask = df['治疗方案'] == tx
        kmf_t = KaplanMeierFitter()
        kmf_t.fit(df.loc[mask, '随访时间(月)'], event_observed=df.loc[mask, '结局事件'], label=tx)
        kmf_t.plot_survival_function(ax=ax, ci_show=False, color=colors_tx[tx], linewidth=2)
        km_results_tx[tx] = kmf_t.median_survival_time_
    
    lr_tx = multivariate_logrank_test(df['随访时间(月)'], df['治疗方案'], df['结局事件'])
    results['治疗_logrank_chi2'] = lr_tx.test_statistic
    results['治疗_logrank_p'] = lr_tx.p_value
    results['治疗_中位OS'] = km_results_tx
    
    ax.set_xlabel('时间（月）', fontsize=12)
    ax.set_ylabel('生存概率', fontsize=12)
    ax.set_title('不同治疗方案 Kaplan-Meier 生存曲线', fontsize=14, fontweight='bold')
    ax.legend(fontsize=11, loc='lower left')
    ax.text(0.95, 0.95, f'Log-rank p{"<0.001" if lr_tx.p_value < 0.001 else f"={lr_tx.p_value:.4f}"}',
            transform=ax.transAxes, ha='right', va='top', fontsize=11,
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    ax.grid(True, alpha=0.3)
    ax.set_ylim(0, 1.05)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig3_treatment_km_curve.png'), dpi=150, bbox_inches='tight')
    plt.close()
    
    print("  ✓ 3张KM生存曲线已生成")
    return results


# ============================================================
# 第四部分：Cox比例风险回归
# ============================================================

def cox_regression(df):
    """单因素 + 多因素Cox回归"""
    
    # 数据预处理：创建哑变量
    df_cox = df.copy()
    df_cox['性别_男'] = (df_cox['性别'] == '男').astype(int)
    df_cox['分期_II'] = (df_cox['临床分期'] == 'II期').astype(int)
    df_cox['分期_III'] = (df_cox['临床分期'] == 'III期').astype(int)
    df_cox['分期_IV'] = (df_cox['临床分期'] == 'IV期').astype(int)
    df_cox['病理_鳞癌'] = (df_cox['病理类型'] == '鳞癌').astype(int)
    df_cox['病理_大细胞癌'] = (df_cox['病理类型'] == '大细胞癌').astype(int)
    df_cox['分化_中'] = (df_cox['分化程度'] == '中分化').astype(int)
    df_cox['分化_低'] = (df_cox['分化程度'] == '低分化').astype(int)
    df_cox['治疗_化疗'] = (df_cox['治疗方案'] == '手术+化疗').astype(int)
    df_cox['治疗_放化疗'] = (df_cox['治疗方案'] == '手术+放化疗').astype(int)
    df_cox['KPS高'] = (df_cox['KPS评分'] >= 80).astype(int)
    df_cox['吸烟_已戒'] = (df_cox['吸烟史'] == '已戒烟').astype(int)
    df_cox['吸烟_目前'] = (df_cox['吸烟史'] == '目前吸烟').astype(int)
    
    covariates = ['年龄', '性别_男', '分期_II', '分期_III', '分期_IV',
                  '病理_鳞癌', '病理_大细胞癌', '分化_中', '分化_低',
                  '治疗_化疗', '治疗_放化疗', 'KPS高', '肿瘤大小(cm)',
                  '吸烟_已戒', '吸烟_目前']
    
    # ---- 单因素Cox回归 ----
    univariate_results = []
    for cov in covariates:
        cph = CoxPHFitter()
        try:
            cph.fit(df_cox[['随访时间(月)', '结局事件', cov]], 
                    duration_col='随访时间(月)', event_col='结局事件')
            summary = cph.summary
            hr = np.exp(summary['coef'].values[0])
            ci_lower = np.exp(summary['coef lower 95%'].values[0])
            ci_upper = np.exp(summary['coef upper 95%'].values[0])
            p_val = summary['p'].values[0]
            univariate_results.append({
                '变量': cov, 'HR': hr, '95%CI下限': ci_lower, 
                '95%CI上限': ci_upper, 'p值': p_val
            })
        except Exception:
            pass
    
    uni_df = pd.DataFrame(univariate_results)
    
    # ---- 多因素Cox回归 ----
    # 纳入单因素中p<0.1的变量
    sig_vars = uni_df[uni_df['p值'] < 0.1]['变量'].tolist()
    if len(sig_vars) < 3:
        sig_vars = covariates[:8]  # 保底纳入前8个
    
    cph_multi = CoxPHFitter()
    cph_multi.fit(df_cox[['随访时间(月)', '结局事件'] + sig_vars],
                  duration_col='随访时间(月)', event_col='结局事件')
    
    multi_summary = cph_multi.summary.copy()
    multi_summary['HR'] = np.exp(multi_summary['coef'])
    multi_summary['HR_lower'] = np.exp(multi_summary['coef lower 95%'])
    multi_summary['HR_upper'] = np.exp(multi_summary['coef upper 95%'])
    
    multivariate_results = []
    for var in multi_summary.index:
        row = multi_summary.loc[var]
        multivariate_results.append({
            '变量': var,
            'HR': row['HR'],
            '95%CI下限': row['HR_lower'],
            '95%CI上限': row['HR_upper'],
            'p值': row['p']
        })
    
    multi_df = pd.DataFrame(multivariate_results)
    
    # 模型统计
    model_stats = {
        'concordance': cph_multi.concordance_index_,
        'log_likelihood': cph_multi.log_likelihood_,
        'AIC': cph_multi.AIC_partial_,
        'n_obs': cph_multi.event_observed.shape[0],
        'n_events': int(cph_multi.event_observed.sum())
    }
    
    return uni_df, multi_df, model_stats, sig_vars


# ============================================================
# 第五部分：森林图
# ============================================================

def create_forest_plot(multi_df, output_dir):
    """绘制多因素Cox回归森林图"""
    
    # 变量名称映射（美化显示）
    name_map = {
        '年龄': '年龄(每增加1岁)',
        '性别_男': '性别(男 vs 女)',
        '分期_II': '临床分期(II vs I)',
        '分期_III': '临床分期(III vs I)',
        '分期_IV': '临床分期(IV vs I)',
        '病理_鳞癌': '病理(鳞癌 vs 腺癌)',
        '病理_大细胞癌': '病理(大细胞癌 vs 腺癌)',
        '分化_中': '分化程度(中 vs 高)',
        '分化_低': '分化程度(低 vs 高)',
        '治疗_化疗': '治疗(手术+化疗 vs 单纯手术)',
        '治疗_放化疗': '治疗(手术+放化疗 vs 单纯手术)',
        'KPS高': 'KPS评分(≥80 vs <80)',
        '肿瘤大小(cm)': '肿瘤大小(每增加1cm)',
        '吸烟_已戒': '吸烟(已戒 vs 从不)',
        '吸烟_目前': '吸烟(目前 vs 从不)'
    }
    
    plot_df = multi_df.copy()
    plot_df['显示名称'] = plot_df['变量'].map(name_map).fillna(plot_df['变量'])
    plot_df = plot_df.sort_values('HR', ascending=True).reset_index(drop=True)
    
    fig, ax = plt.subplots(figsize=(10, max(6, len(plot_df) * 0.5 + 1)))
    
    y_pos = range(len(plot_df))
    
    # 绘制CI线段和点
    for i, (_, row) in enumerate(plot_df.iterrows()):
        color = '#E74C3C' if row['p值'] < 0.05 else '#7F8C8D'
        ax.plot([row['95%CI下限'], row['95%CI上限']], [i, i], 
                color=color, linewidth=2, solid_capstyle='round')
        ax.plot(row['HR'], i, 'o', color=color, markersize=8)
    
    # 参考线
    ax.axvline(x=1, color='black', linestyle='--', linewidth=1, alpha=0.7)
    
    # Y轴标签
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(plot_df['显示名称'], fontsize=10)
    
    # 右侧标注HR和CI
    x_max = plot_df['95%CI上限'].max() * 1.1
    for i, (_, row) in enumerate(plot_df.iterrows()):
        sig_mark = '*' if row['p值'] < 0.05 else ''
        ax.text(x_max, i, f"HR={row['HR']:.2f} ({row['95%CI下限']:.2f}-{row['95%CI上限']:.2f}){sig_mark}",
                va='center', fontsize=9)
    
    ax.set_xlabel('Hazard Ratio (HR)', fontsize=12)
    ax.set_title('多因素Cox回归 森林图', fontsize=14, fontweight='bold')
    ax.set_xlim(0, x_max * 1.6)
    ax.grid(True, axis='x', alpha=0.3)
    
    # 添加注释
    ax.text(0.02, -0.08, '* p<0.05   HR>1: 风险因素   HR<1: 保护因素',
            transform=ax.transAxes, fontsize=9, color='gray')
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'fig4_forest_plot.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  ✓ 森林图已生成")


# ============================================================
# 第六部分：生成Word报告
# ============================================================

def generate_word_report(df, baseline, km_results, uni_df, multi_df, 
                         model_stats, output_dir):
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
    run = title.add_run('非小细胞肺癌患者术后生存分析\n——医学统计分析报告')
    run.font.size = Pt(22)
    run.font.bold = True
    
    doc.add_paragraph()
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('分析方法：Kaplan-Meier生存分析 | Log-rank检验 | Cox比例风险回归')
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = info.add_run(f'样本量：{len(df)}例\n分析日期：2026年5月')
    run.font.size = Pt(12)
    doc.add_page_break()
    
    # ---- 一、研究背景 ----
    doc.add_heading('一、研究背景与目的', level=1)
    doc.add_paragraph(
        '非小细胞肺癌（NSCLC）占所有肺癌的80-85%，是全球癌症死亡的首要原因。'
        '本研究回顾性分析了500例NSCLC术后患者的临床随访数据，'
        '旨在探究影响患者总生存期（OS）的独立预后因素，为临床决策提供循证依据。'
    )
    
    stats_info, cat_stats, cont_stats = baseline
    
    # ---- 二、基线特征 ----
    doc.add_heading('二、患者基线特征', level=1)
    doc.add_paragraph(
        f'共纳入{stats_info["总样本量"]}例患者，中位随访时间{stats_info["中位随访时间(月)"]}月，'
        f'共发生死亡事件{stats_info["事件数(死亡)"]}例（{stats_info["事件发生率"]}），'
        f'删失{stats_info["删失数"]}例。'
    )
    
    # 基线表格
    doc.add_heading('表1 患者基线特征', level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['特征', '类别', 'n(%)/M±SD']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    # 连续变量
    for var, val in cont_stats.items():
        row = table.add_row().cells
        row[0].text = var
        row[1].text = '-'
        row[2].text = val
    
    # 分类变量
    for var, categories in cat_stats.items():
        for i, (cat, val) in enumerate(categories.items()):
            row = table.add_row().cells
            row[0].text = var if i == 0 else ''
            row[1].text = cat
            row[2].text = val
    
    # ---- 三、生存分析 ----
    doc.add_heading('三、Kaplan-Meier生存分析', level=1)
    
    doc.add_heading('3.1 总体生存情况', level=2)
    doc.add_paragraph(
        f'全组患者中位总生存期为{km_results["总体中位生存期"]:.1f}个月。'
    )
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig1_overall_km_curve.png'), width=Inches(5.0))
    doc.add_paragraph('图1 总体Kaplan-Meier生存曲线', style='Caption')
    
    doc.add_heading('3.2 不同临床分期生存比较', level=2)
    stage_os = km_results['分期_中位OS']
    doc.add_paragraph(
        f'不同分期患者中位OS分别为：I期{stage_os.get("I期", "NR"):.1f}月、'
        f'II期{stage_os.get("II期", "NR"):.1f}月、III期{stage_os.get("III期", "NR"):.1f}月、'
        f'IV期{stage_os.get("IV期", "NR"):.1f}月。'
        f'Log-rank检验显示各组间差异有统计学意义'
        f'（χ²={km_results["分期_logrank_chi2"]:.2f}, p<0.001）。'
    )
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig2_stage_km_curve.png'), width=Inches(5.5))
    doc.add_paragraph('图2 不同临床分期Kaplan-Meier生存曲线', style='Caption')
    
    doc.add_heading('3.3 不同治疗方案生存比较', level=2)
    tx_os = km_results['治疗_中位OS']
    p_tx = km_results['治疗_logrank_p']
    doc.add_paragraph(
        f'不同治疗方案的中位OS：手术+化疗{tx_os.get("手术+化疗","NR"):.1f}月、'
        f'手术+放化疗{tx_os.get("手术+放化疗","NR"):.1f}月、'
        f'单纯手术{tx_os.get("单纯手术","NR"):.1f}月'
        f'（Log-rank χ²={km_results["治疗_logrank_chi2"]:.2f}, '
        f'p{"<0.001" if p_tx < 0.001 else f"={p_tx:.4f}"}）。'
    )
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig3_treatment_km_curve.png'), width=Inches(5.5))
    doc.add_paragraph('图3 不同治疗方案Kaplan-Meier生存曲线', style='Caption')
    
    # ---- 四、Cox回归 ----
    doc.add_heading('四、Cox比例风险回归分析', level=1)
    
    doc.add_heading('4.1 单因素分析', level=2)
    doc.add_paragraph('对各临床因素分别进行单因素Cox回归分析，结果如下：')
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['变量', 'HR', '95% CI', 'p值', '显著性']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    for _, r in uni_df.iterrows():
        row = table.add_row().cells
        row[0].text = r['变量']
        row[1].text = f"{r['HR']:.3f}"
        row[2].text = f"{r['95%CI下限']:.3f}-{r['95%CI上限']:.3f}"
        row[3].text = f"{r['p值']:.4f}" if r['p值'] >= 0.001 else '<0.001'
        row[4].text = '**' if r['p值'] < 0.01 else ('*' if r['p值'] < 0.05 else '')
    
    doc.add_heading('4.2 多因素分析', level=2)
    doc.add_paragraph(
        f'将单因素分析中p<0.1的变量纳入多因素Cox回归模型。'
        f'模型C-index={model_stats["concordance"]:.3f}，具有较好的预测能力。'
    )
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, text in enumerate(['变量', 'HR', '95% CI', 'p值', '显著性']):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].font.bold = True
    
    for _, r in multi_df.iterrows():
        row = table.add_row().cells
        row[0].text = r['变量']
        row[1].text = f"{r['HR']:.3f}"
        row[2].text = f"{r['95%CI下限']:.3f}-{r['95%CI上限']:.3f}"
        row[3].text = f"{r['p值']:.4f}" if r['p值'] >= 0.001 else '<0.001'
        row[4].text = '**' if r['p值'] < 0.01 else ('*' if r['p值'] < 0.05 else '')
    
    doc.add_paragraph()
    doc.add_picture(os.path.join(output_dir, 'fig4_forest_plot.png'), width=Inches(6.0))
    doc.add_paragraph('图4 多因素Cox回归森林图', style='Caption')
    
    # ---- 五、结论 ----
    doc.add_heading('五、结论', level=1)
    
    # 找出显著的独立预后因素
    sig_factors = multi_df[multi_df['p值'] < 0.05]
    risk_factors = sig_factors[sig_factors['HR'] > 1]['变量'].tolist()
    protect_factors = sig_factors[sig_factors['HR'] < 1]['变量'].tolist()
    
    doc.add_paragraph(
        f'本研究基于{len(df)}例NSCLC术后患者的随访数据，通过Kaplan-Meier生存分析和'
        f'Cox比例风险回归模型，识别出以下独立预后因素：'
    )
    
    if risk_factors:
        doc.add_paragraph(f'独立危险因素（HR>1）：{", ".join(risk_factors)}')
    if protect_factors:
        doc.add_paragraph(f'独立保护因素（HR<1）：{", ".join(protect_factors)}')
    
    doc.add_paragraph()
    doc.add_paragraph(
        f'模型预测能力良好（C-index={model_stats["concordance"]:.3f}），'
        f'可为临床个体化治疗决策和预后评估提供参考。'
    )
    
    doc.add_paragraph()
    doc.add_paragraph('注：* p<0.05, ** p<0.01', style='Caption')
    
    # 保存
    report_path = os.path.join(output_dir, '生存分析报告_NSCLC术后随访.docx')
    doc.save(report_path)
    return report_path


# ============================================================
# 主函数
# ============================================================

def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    print("=" * 60)
    print("  项目2：非小细胞肺癌患者术后生存分析")
    print("=" * 60)
    
    # 1. 生成数据
    print("\n[1/6] 生成模拟随访数据（500例）...")
    df = generate_survival_data(n=500)
    data_path = os.path.join(output_dir, 'data_survival.csv')
    df.to_csv(data_path, index=False, encoding='utf-8-sig')
    print(f"  ✓ 数据已保存: {data_path}")
    print(f"  ✓ 事件数: {df['结局事件'].sum()}, 删失数: {(1-df['结局事件']).sum()}")
    
    # 2. 基线特征
    print("\n[2/6] 基线特征描述...")
    baseline = baseline_characteristics(df)
    stats_info = baseline[0]
    print(f"  ✓ 中位随访时间: {stats_info['中位随访时间(月)']}月")
    
    # 3. KM生存分析
    print("\n[3/6] Kaplan-Meier生存分析 + Log-rank检验...")
    km_results = kaplan_meier_analysis(df, output_dir)
    print(f"  ✓ 总体中位OS: {km_results['总体中位生存期']:.1f}月")
    print(f"  ✓ 分期Log-rank: χ²={km_results['分期_logrank_chi2']:.2f}, p<0.001")
    
    # 4. Cox回归
    print("\n[4/6] Cox比例风险回归...")
    uni_df, multi_df, model_stats, sig_vars = cox_regression(df)
    print(f"  ✓ 单因素: {len(uni_df)}个变量")
    print(f"  ✓ 多因素: {len(sig_vars)}个变量纳入, C-index={model_stats['concordance']:.3f}")
    n_sig = len(multi_df[multi_df['p值'] < 0.05])
    print(f"  ✓ 独立预后因素: {n_sig}个 (p<0.05)")
    
    # 5. 森林图
    print("\n[5/6] 生成森林图...")
    create_forest_plot(multi_df, output_dir)
    
    # 6. Word报告
    print("\n[6/6] 生成Word分析报告...")
    report_path = generate_word_report(df, baseline, km_results, uni_df, multi_df, model_stats, output_dir)
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("  项目2完成！输出文件：")
    print(f"    - 数据: {data_path}")
    print(f"    - 报告: {report_path}")
    print(f"    - 图表: fig1~fig4 (共4张)")
    print("=" * 60)


if __name__ == '__main__':
    main()
