#!/usr/bin/env python3
"""
风电项目财务测算主脚本（v2固化标准）。

输入: params dict（见 references/wind-power-model.md）
输出:
  - JSON 数据文件（含完整20年测算结果 + 敏感性分析）
  - 7张图表PNG文件

用法:
    python gen_wind_model.py <project_name> <params.json> <output_dir>
    python gen_wind_model.py "项目名称" '{"epc_price":5.4,"capacity":60,...}' /output/path
"""
import sys, json, os, math
import numpy as np
import numpy_financial as npf
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# ========== 图表中文渲染字体（macOS黑体）==========
HEI_FONT = '/System/Library/AssetsV2/com_apple_MobileAsset_Font8/5feac9245cca79adaf638ded7a4994b1ddb33ca0.asset/AssetData/Hei.ttf'
_chinese_font = fm.FontProperties(fname=HEI_FONT)
plt.rcParams['axes.unicode_minus'] = False


def calculate_model(params):
    """执行完整财务测算，返回(yearly_data, metrics, params_out)。"""
    # 提取参数
    capacity   = params['capacity']          # MW
    epc_price  = params['epc_price']         # 元/W
    util_hours = params['util_hours']         # h
    op_years   = params.get('op_years', 20)           # 年
    cfd_price  = params['cfd_price']         # 元/kWh
    post_cfd   = params['post_cfd_price']     # 元/kWh
    cfd_years  = params['cfd_years']          # 年
    curtail    = params.get('curtailment', 0.05)  # 限电率
    loan_rate  = params['loan_rate']           # 含税
    loan_years = params['loan_years']         # 年
    capital_r  = params.get('capital_ratio', 0.20)  # 资本金比例
    om_rates   = params.get('om_rates', {'1-5': 0.0, '6-10': 0.0, '11-20': 0.0})
    ins_rate   = params.get('insurance_rate', 0.005)
    spare_rate = params.get('spare_rate', 0.001)
    payment    = params.get('payment_method', '年付')

    # 派生计算
    total_inv = epc_price * capacity * 100            # 万元 (元/W × MW × 100 = 万元)
    capital   = total_inv * capital_r                 # 万元
    loan      = total_inv * (1 - capital_r)             # 万元
    annual_power = capacity * util_hours * (1 - curtail) * 1000         # kWh → ÷10000转为万元=kWh

    # 折旧
    dep_base  = total_inv * 0.97                       # 折旧基数
    annual_dep = dep_base / op_years                    # 万元/年
    residual   = total_inv * 0.03                       # 残值

    # 年利率（月利率）
    annual_rate  = loan_rate
    monthly_rate = annual_rate / 12

    # 等额本金还款：每月本金相同，每月利息递减
    monthly_principal = loan * 10000 / loan_years / 12   # 元/月 → 计算中用万元
    yearly_principal_list = []
    yearly_interest_list = []
    remaining = loan * 10000  # 元
    for yr in range(1, loan_years + 1):
        yr_interest = 0.0
        for m in range(12):
            yr_interest += remaining * loan_rate              # 万元 × 年利率 = 万元/年
        yearly_principal_list.append(monthly_principal * 12)     # 万元
        yearly_interest_list.append(yr_interest)
        yr_princ = monthly_principal * 12; remaining -= yr_princ; yearly_principal_list[-1] = yr_princ

    # 每年数据
    data = []
    cumsum_full   = 0.0
    cumsum_equity = 0.0

    for yr in range(1, op_years + 1):
        price = cfd_price if yr <= cfd_years else post_cfd

        # 发电收入（含税）= 年发电量 × 电价 × 1.13
        revenue_vat = annual_power * price / 10000          # kWh × 元/kWh = 元 → ÷10000=万元
        revenue_net = revenue_vat / 1.13                   # 万元 / 1.13 = 万元

        # 进项税抵扣（仅在融资期内，本金+利息 的13%）
        deduct_vat = 0.0
        if yr <= loan_years:
            deduct_vat = (yearly_principal_list[yr-1] + yearly_interest_list[yr-1]) * 0.13 / 1.13

        # O&M 费用（元/W → 万元）
        if yr <= 5:
            om_rate = om_rates.get('1-5', 0.0)
        elif yr <= 10:
            om_rate = om_rates.get('6-10', 0.0)
        else:
            om_rate = om_rates.get('11-20', 0.0)
        om_cost = capacity * 10000 * om_rate / 10000  # 万元

        insurance = total_inv * ins_rate              # 万元
        spare     = total_inv * spare_rate            # 万元
        op_cost   = om_cost + insurance + spare       # 万元

        # 销项税
        vat_output = annual_power * price * 0.13           # 万kWh × 元/kWh × 13% = 万元

        # 城建税及附加（7%增值税）
        city_edu_tax = vat_output * 0.07  # 万元

        # 利息（当期）
        interest = yearly_interest_list[yr-1] if yr <= loan_years else 0.0

        # 本金（当期）
        principal = yearly_principal_list[yr-1] if yr <= loan_years else 0.0

        # 全投资CF = 收入(含税)+进项税抵扣-运营成本-销项税-城建税-调整所得税
        # 调整所得税 = max(0, PBT) × 25%，PBT不含利息（因为全投资CF不加回利息）
        pbt_full = revenue_net - op_cost - annual_dep
        adj_tax  = max(0, pbt_full) * 0.25
        cf_full  = revenue_vat + deduct_vat - op_cost - vat_output - city_edu_tax - adj_tax

        # 资本金CF = 收入(含税)+进项税抵扣-运营成本-销项税-城建税-所得税-本金-利息
        pbt_equity = revenue_net - op_cost - annual_dep - interest
        income_tax = max(0, pbt_equity) * 0.25
        cf_equity  = revenue_vat + deduct_vat - op_cost - vat_output - city_edu_tax - income_tax - principal - interest

        # 净利润（用于DSCR）
        net_profit = max(0, pbt_equity) * (1 - 0.25) if pbt_equity > 0 else 0.0

        # FCFE = 净利润 + 折旧 + 利息抵税(×75%) - 本金
        fcfe = net_profit + annual_dep + interest * (1 - 0.25) - principal

        cumsum_full   += cf_full
        cumsum_equity += cf_equity

        d = {
            'year': yr, 'price': price,
            'revenue_vat': revenue_vat, 'revenue_net': revenue_net,
            'deduct_vat': deduct_vat, 'op_cost': op_cost,
            'om_cost': om_cost, 'insurance': insurance, 'spare': spare,
            'vat_output': vat_output, 'city_edu_tax': city_edu_tax,
            'depreciation': annual_dep, 'interest': interest, 'principal': principal,
            'pbt': pbt_equity, 'net_profit': net_profit,
            'adj_tax': adj_tax, 'income_tax': income_tax,
            'cf_full': cf_full, 'cf_equity': cf_equity,
            'cumsum_full': cumsum_full, 'cumsum_equity': cumsum_equity,
            'fcfe': fcfe,
        }

        # DSCR = (净利润+折旧+利息)/(本金+利息)
        debt_service = principal + interest
        d['dscr'] = (net_profit + annual_dep + interest) / debt_service if debt_service > 1e-6 else None

        data.append(d)

    # 指标计算
    # 全投资IRR
    full_cf = [-total_inv] + [d['cf_full'] for d in data]
    irr_full = npf.irr(full_cf) * 100 if full_cf else 0

    # 资本金IRR
    equity_cf = [-capital] + [d['cf_equity'] for d in data]
    irr_equity = npf.irr(equity_cf) * 100 if equity_cf else 0

    # NPV（WACC=5.5%，资本金成本=8%）
    npv_full   = npf.npv(0.055, full_cf)
    npv_equity = npf.npv(0.08,  equity_cf)

    # 年均净利润/资本金（ROE）
    avg_net_profit = np.mean([d['net_profit'] for d in data])
    roe = avg_net_profit / capital * 100 if capital > 0 else 0

    # 年均EBIT/总资产（ROA）：EBIT = 净利润+利息+所得税
    avg_ebit = np.mean([d['net_profit'] + d['interest'] + d['income_tax'] for d in data])
    roa = avg_ebit / total_inv * 100 if total_inv > 0 else 0

    # ROIC = NOPAT/(债务+权益)，NOPAT=EBIT×(1-25%)
    nopat = avg_ebit * (1 - 0.25)
    roic = nopat / total_inv * 100 if total_inv > 0 else 0

    # ROI = 累计净收入/总投资
    total_net_income = sum(d['revenue_net'] - d['op_cost'] - d['city_edu_tax'] - d['income_tax'] for d in data)
    roi = total_net_income / total_inv * 100 if total_inv > 0 else 0

    metrics = {
        'irr_full':   round(irr_full, 4),
        'irr_equity': round(irr_equity, 4),
        'npv_full':   round(npv_full, 2),
        'npv_equity': round(npv_equity, 2),
        'roe':  round(roe, 2),
        'roa':  round(roa, 2),
        'roic': round(roic, 2),
        'roi':  round(roi, 2),
    }

    params_out = {
        'capacity': capacity, 'epc_price': epc_price, 'util_hours': util_hours,
        'op_years': op_years, 'cfd_price': cfd_price, 'post_cfd_price': post_cfd,
        'cfd_years': cfd_years, 'curtailment': curtail, 'loan_rate': loan_rate,
        'loan_years': loan_years, 'capital_ratio': capital_r,
        'total_inv': total_inv, 'capital': capital, 'loan': loan,
        'annual_power': annual_power, 'residual_value': residual,
        'annual_depreciation': annual_dep,
        'om_rates': om_rates, 'insurance_rate': ins_rate, 'spare_rate': spare_rate,
    }

    return data, metrics, params_out


def sensitivity_analysis(data, metrics, params, params_out):
    """计算5项敏感性分析，返回sens dict。"""
    base_irr = metrics['irr_equity']
    sens = {}

    def irr_params(p):
        """用给定参数重新计算IRR"""
        new_params = dict(params, **p)
        _, m, _ = calculate_model(new_params)
        return m['irr_equity']

    # 1. 电价
    base_cfd = params['cfd_price']
    sens['cfd_price'] = [((d - base_cfd) / base_cfd * 100, irr_params({'cfd_price': d}))
                          for d in [base_cfd * r for r in [0.80, 0.90, 0.95, 1.00, 1.05, 1.10, 1.20]]]

    # 2. 利用小时
    base_uh = params['util_hours']
    sens['util_hours'] = [(((uh - base_uh) / base_uh * 100), irr_params({'util_hours': uh}))
                           for uh in [base_uh * r for r in [0.80, 0.90, 0.95, 1.00, 1.05, 1.10, 1.20]]]

    # 3. 总投资
    base_epc = params['epc_price']
    sens['total_inv'] = [(((epc - base_epc) / base_epc * 100), irr_params({'epc_price': epc}))
                          for epc in [base_epc * r for r in [0.90, 0.95, 1.00, 1.05, 1.10]]]

    # 4. 贷款利率
    base_lr = params['loan_rate']
    sens['loan_rate'] = [((lr - base_lr) * 100, irr_params({'loan_rate': lr}))
                          for lr in [base_lr * r for r in [0.80, 0.90, 0.95, 1.00, 1.05, 1.10, 1.20]]]

    # 5. 运维成本（元/W）
    om_base = params.get('om_rates', {}).get('1-5', 0.0)
    sens['om_cost'] = [(((om - om_base) / max(om_base, 0.001) * 100 if om_base > 0 else 0), irr_params({'om_rates': {'1-5': om, '6-10': om * 1.2, '11-20': om * 1.5}}))
                        for om in [om_base * r for r in [0.50, 0.75, 1.00, 1.25, 1.50, 2.00]]]

    return sens


def generate_charts(data, metrics, sens, params, project_name, output_dir):
    """生成7张分析图表。"""
    os.makedirs(output_dir, exist_ok=True)
    years = [d['year'] for d in data]
    base_irr = metrics['irr_equity']

    def fmt(**kw): return f"{output_dir}/{kw['f']}"

    charts = []

    # 图1：全投资净现金流
    fig, ax = plt.subplots(figsize=(10, 4.5))
    colors = ['#2E86AB' if d['cf_full'] >= 0 else '#E94F37' for d in data]
    bars = ax.bar(years, [d['cf_full'] for d in data], color=colors, edgecolor='white', linewidth=0.5)
    ax.axhline(0, color='#333', linewidth=0.8)
    ax.set_xlabel('年份', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('金额（万元）', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n全投资净现金流趋势（万元）', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.set_xticks(years); ax.set_xticklabels(years, fontsize=7)
    for bar, d in zip(bars, data):
        if abs(d['cf_full']) > 100:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20, f'{d["cf_full"]:.0f}',
                    ha='center', va='bottom', fontsize=6.5, color='#333')
    fig.tight_layout()
    fig.savefig(fmt(f='chart1.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart1.png')

    # 图2：资本金净现金流
    fig, ax = plt.subplots(figsize=(10, 4.5))
    colors = ['#27AE60' if d['cf_equity'] >= 0 else '#E94F37' for d in data]
    bars = ax.bar(years, [d['cf_equity'] for d in data], color=colors, edgecolor='white', linewidth=0.5)
    ax.axhline(0, color='#333', linewidth=0.8)
    ax.set_xlabel('年份', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('金额（万元）', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n资本金净现金流趋势（万元）', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.set_xticks(years); ax.set_xticklabels(years, fontsize=7)
    for bar, d in zip(bars, data):
        if abs(d['cf_equity']) > 50:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15, f'{d["cf_equity"]:.0f}',
                    ha='center', va='bottom', fontsize=6.5, color='#333')
    fig.tight_layout()
    fig.savefig(fmt(f='chart2.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart2.png')

    # 图3：累计现金流
    fig, ax = plt.subplots(figsize=(10, 4.5))
    ax.plot(years, [d['cumsum_full'] for d in data], 'o-', color='#2E86AB', label='全投资', linewidth=1.5, markersize=4)
    ax.plot(years, [d['cumsum_equity'] for d in data], 's-', color='#27AE60', label='资本金', linewidth=1.5, markersize=4)
    ax.axhline(0, color='#333', linewidth=0.8)
    ax.fill_between(years, [d['cumsum_full'] for d in data], alpha=0.1, color='#2E86AB')
    ax.fill_between(years, [d['cumsum_equity'] for d in data], alpha=0.1, color='#27AE60')
    ax.set_xlabel('年份', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('金额（万元）', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n累计现金流对比（万元）', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.legend(prop=_chinese_font, loc='upper left', fontsize=9)
    ax.set_xticks(years); ax.set_xticklabels(years, fontsize=7)
    for i, d in enumerate(data):
        if i % 2 == 0:
            ax.annotate(f'{d["cumsum_full"]:.0f}', (d['year'], d['cumsum_full']),
                         textcoords="offset points", xytext=(0, 8), ha='center', fontsize=6, color='#2E86AB')
            ax.annotate(f'{d["cumsum_equity"]:.0f}', (d['year'], d['cumsum_equity']),
                         textcoords="offset points", xytext=(0, -12), ha='center', fontsize=6, color='#27AE60')
    fig.tight_layout()
    fig.savefig(fmt(f='chart3.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart3.png')

    # 图4：收入成本结构
    x = np.array(years)
    rev = [d['revenue_vat'] for d in data]
    cost = [d['op_cost'] for d in data]
    net = [d['revenue_vat'] - d['op_cost'] for d in data]
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.bar(x - 0.175, rev, width=0.35, label='发电收入（含税）', color='#2E86AB', alpha=0.85)
    ax.bar(x + 0.175, cost, width=0.35, label='运营成本', color='#E94F37', alpha=0.85)
    ax.plot(x, net, 'o-', color='#27AE60', linewidth=1.5, markersize=3, label='净收入')
    ax.set_xlabel('年份', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('金额（万元）', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n收入成本结构（全生命周期，万元）', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.legend(prop=_chinese_font, loc='upper left', fontsize=9)
    ax.set_xticks(x); ax.set_xticklabels(years, fontsize=7)
    fig.tight_layout()
    fig.savefig(fmt(f='chart4.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart4.png')

    # 图5：DSCR
    valid_yr = [d['year'] for d in data if d['dscr'] is not None]
    valid_d  = [d['dscr'] for d in data if d['dscr'] is not None]
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(valid_yr, valid_d, 'o-', color='#8E44AD', linewidth=1.5, markersize=5)
    ax.axhline(1.2, color='#27AE60', linewidth=1, linestyle='--', label='健康线(1.2)')
    ax.axhline(1.0, color='#F39C12', linewidth=1, linestyle='--', label='预警线(1.0)')
    ax.fill_between(valid_yr, valid_d, alpha=0.2, color='#8E44AD')
    ax.set_xlabel('年份', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('DSCR', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n偿债覆盖率（DSCR）', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.legend(prop=_chinese_font, loc='upper right', fontsize=9)
    ax.set_xticks(valid_yr); ax.set_xticklabels(valid_yr, fontsize=7)
    for x, y in zip(valid_yr, valid_d):
        if y > 1.5 or y < 1.3:
            ax.annotate(f'{y:.2f}', (x, y), textcoords="offset points", xytext=(3, 5), fontsize=7, color='#8E44AD')
    fig.tight_layout()
    fig.savefig(fmt(f='chart5.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart5.png')

    # 图6：电价敏感性
    deltas = [d for d, _ in sens['cfd_price']]
    irrs   = [irr for _, irr in sens['cfd_price']]
    colors = ['#27AE60' if irr >= base_irr else '#E94F37' for irr in irrs]
    x_labels = ['%+.0f%%' % d for d in deltas]
    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(x_labels, irrs, color=colors, edgecolor='white')
    ax.axhline(base_irr, color='#2E86AB', linewidth=1.5, linestyle='--', label=f'基准IRR={base_irr:.2f}%')
    ax.set_xlabel('电价变动', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('资本金IRR (%)', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n电价对资本金IRR的影响', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.legend(prop=_chinese_font, fontsize=9)
    for bar, irr in zip(bars, irrs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{irr:.1f}%',
                ha='center', va='bottom', fontsize=8, color='#333', fontproperties=_chinese_font)
    fig.tight_layout()
    fig.savefig(fmt(f='chart6.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart6.png')

    # 图7：多变量敏感性
    # 取共同长度的数据（都用电价敏感性的x轴）
    common_len = min(len(sens['cfd_price']), len(sens['util_hours']),
                     len(sens['total_inv']), len(sens['om_cost']))
    x_labels = ['%+.0f%%' % sens['cfd_price'][j][0] for j in range(common_len)]
    x_pos = np.arange(common_len)
    width = 0.2
    sens_vars = [
        ('电价',     [sens['cfd_price'][j][1]     for j in range(common_len)]),
        ('利用小时',  [sens['util_hours'][j][1]    for j in range(common_len)]),
        ('总投资',   [sens['total_inv'][j][1]     for j in range(common_len)]),
        ('运维成本', [sens['om_cost'][j][1]       for j in range(min(common_len, len(sens['om_cost'])))]),
    ]
    # 确保运维成本列表长度一致
    sens_vars = [(l, irrs[:common_len]) for l, irrs in sens_vars]
    fig, ax = plt.subplots(figsize=(10, 5))
    for i, (label, irrs) in enumerate(sens_vars):
        ax.bar(x_pos + i * width, irrs, width, label=label)
    ax.axhline(base_irr, color='#333', linewidth=1.5, linestyle='--', label=f'基准={base_irr:.2f}%')
    ax.set_xlabel('变动幅度', fontproperties=_chinese_font, fontsize=10)
    ax.set_ylabel('资本金IRR (%)', fontproperties=_chinese_font, fontsize=10)
    ax.set_title(f'{project_name}\n多变量敏感性分析', fontproperties=_chinese_font, fontsize=11, fontweight='bold')
    ax.set_xticks(x_pos + width * 1.5)
    ax.set_xticklabels(x_labels)
    ax.legend(prop=_chinese_font, loc='upper left', fontsize=8)
    fig.tight_layout()
    fig.savefig(fmt(f='chart7.png'), dpi=150, bbox_inches='tight'); plt.close()
    charts.append('chart7.png')

    return charts


def save_json(project_name, params, metrics, data, sens, output_path):
    """保存JSON结果。"""
    output = {
        'project_name': project_name,
        'params': params,
        'metrics': metrics,
        'yearly': data,
        'sens': sens,
    }
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)


# ========== 主入口 ==========
if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("用法: python gen_wind_model.py <project_name> <params_json> <output_dir>")
        sys.exit(1)

    project_name = sys.argv[1]
    params_raw   = json.loads(sys.argv[2])
    # 支持从 extract_params_from_excel.py 输出的 {"project_name":..., "params":...} 格式
    if 'params' in params_raw:
        project_name = params_raw.get('project_name', project_name)
        params = params_raw['params']
    else:
        params = params_raw
    output_dir   = sys.argv[3]

    os.makedirs(output_dir, exist_ok=True)

    print(f"项目: {project_name}")
    data, metrics, params_out = calculate_model(params)
    print(f"全投资IRR: {metrics['irr_full']:.2f}%  资本金IRR: {metrics['irr_equity']:.2f}%")
    print(f"全投资NPV: {metrics['npv_full']:.0f}万元  资本金NPV: {metrics['npv_equity']:.0f}万元")

    print("敏感性分析...")
    sens = sensitivity_analysis(data, metrics, params, params_out)

    print("生成图表...")
    charts = generate_charts(data, metrics, sens, params, project_name, output_dir)
    print(f"图表完成: {len(charts)} 张")

    json_path = os.path.join(output_dir, 'model_data.json')
    save_json(project_name, params_out, metrics, data, sens, json_path)
    print(f"JSON已保存: {json_path}")
