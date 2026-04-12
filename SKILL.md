---
name: wind-power-analysis
description: 风电项目财务测算与报告生成。适用场景：（1）接收风电项目Excel参数文件，完成财务建模（DCF/IRR/NPV/DSCR）；（2）生成包含7张图表的Word和PDF分析报告；（3）对已建模项目进行敏感性分析；（4）更新风电测算固化标准（v2标准：残值=总投资×3%、折旧基数=总投资×97%、DSCR=(净利润+折旧+利息)/(本金+利息)、图表使用Hei.ttf中文化）。当用户发送风电项目Excel或要求"测算"、"财务分析"、"生成报告"时触发。
---

# Wind Power Financial Analysis Skill

## 快速使用

收到 Excel 参数文件后，依次执行：

```bash
# 1. 提取参数
python wind-power-analysis/scripts/extract_params_from_excel.py <excel_file> --json-only > params.json

# 2. 生成财务模型 + 图表
python wind-power-analysis/scripts/gen_wind_model.py "<项目名称>" "$(cat params.json)" <output_dir>

# 3. 生成Word报告
python wind-power-analysis/scripts/gen_wind_report.py \
    <output_dir>/model_data.json \
    <output_dir> \
    wind-power-analysis/assets/jianeng_logo_header.png \
    wind-power-analysis/assets/wechat_qr_100h.png \
    <output_dir>/<项目名>财务分析报告.docx
```

## 工作流程

```
Excel参数文件
    ↓
extract_params_from_excel.py → params.json
    ↓
gen_wind_model.py → model_data.json + 7张图表
    ↓
gen_wind_report.py → Word报告 + FPDF报告
    ↓
发送报告给张总（通过Telegram Bot）
```

## 核心参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| capacity | 60 MW | 装机容量 |
| epc_price | 5.2 元/W | EPC单价 |
| util_hours | 1700 h | 利用小时数 |
| op_years | 20 年 | 运营年限 |
| cfd_price | 0.35 元/kWh | CFD结算电价（前10年） |
| post_cfd_price | 0.36 元/kWh | 市场化电价（10年后） |
| cfd_years | 10 年 | CFD执行年限 |
| curtailment | 5% | 限电率 |
| loan_rate | 3.5% | 含税贷款利率 |
| loan_years | 18 年 | 融资期限 |
| capital_ratio | 20% | 资本金比例 |

详细参数格式见 `references/wind-power-model.md`

## 关键固化标准（v2）

- **残值** = 总投资 × 3%（固定，不随年度递减）
- **年折旧** = 总投资 × 97% ÷ 运营年限
- **DSCR** = (净利润 + 折旧 + 利息) / (本金 + 利息)
- **进项税抵扣** = (年本金 + 年利息) × 13%（仅融资期限内）
- **图表中文** = 使用 `Hei.ttf`（macOS黑体，路径固定）

## 敏感性分析

自动计算5项变量对资本金IRR的影响：
1. 电价（CFD价格）
2. 利用小时数
3. 总投资（EPC单价）
4. 贷款利率
5. 运维成本（O&M费率）

## 报告结构（9章节）

1. 执行摘要（含完整指标说明）
2. 全投资现金流量表（全生命周期20年）
3. 资本金现金流量表（全生命周期20年）
4. 资本金利润表（全生命周期20年）
5. FCFE（资本金自由现金流量）
6. DSCR偿债覆盖率（全生命周期）
7. 敏感性分析（5项因素）
8. 关键计算公式
9. 7张分析图表

## 发送文件给张总

```bash
curl -s -X POST "https://api.telegram.org/bot<BOT_TOKEN>/sendDocument" \
  -F "chat_id=402908483" \
  -F "document=@<file_path>" \
  -F "caption=<说明文字>"
```

## 资源路径

```
wind-power-analysis/
├── scripts/
│   ├── extract_params_from_excel.py  # Excel参数提取
│   ├── gen_wind_model.py               # 财务模型+图表生成
│   └── gen_wind_report.py              # Word报告生成
├── references/
│   └── wind-power-model.md             # 详细财务模型参考
└── assets/
    ├── jianeng_logo_header.png         # 江能Logo（封面用）
    └── wechat_qr_100h.png              # 微信公众号二维码
```
