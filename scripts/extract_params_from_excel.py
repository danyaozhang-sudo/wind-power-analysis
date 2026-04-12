#!/usr/bin/env python3
"""
从风电项目Excel文件中提取测算参数。
支持标准格式（见 references/wind-power-model.md 参数章节）。

用法:
    python extract_params_from_excel.py <excel_file_path>
    python extract_params_from_excel.py <excel_file_path> --json-only

输出:
    params dict, 或带 --json-only 时仅输出 JSON
"""
import sys, json, zipfile, xml.etree.ElementTree as ET
from pathlib import Path

def read_excel_params(xlsx_path):
    """读取风电项目Excel参数文件，返回参数字典。"""
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(xlsx_path) as z:
        # 读取共享字符串
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            for si in tree.findall('.//s:si', ns):
                texts = si.findall('.//s:t', ns)
                strings.append(''.join(t.text or '' for t in texts))

        # 读取工作表
        sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
        rows = sheet.findall('.//s:row', ns)

    params = {}
    project_name = None

    for row in rows:
        cells = row.findall('s:c', ns)
        row_data = []
        for c in cells:
            t = c.get('t', '')
            v_el = c.find('s:v', ns)
            v = v_el.text if v_el is not None else ''
            if t == 's' and v:
                v = strings[int(v)]
            row_data.append(v)
        # row_data: [None, 序号, 内容, 输入, 备注]
        if len(row_data) >= 4:
            # 跳过标题行和空行
            seq = row_data[1] if len(row_data) > 1 else None
            name = row_data[2] if len(row_data) > 2 else None
            value = row_data[3] if len(row_data) > 3 else None

            if name is None or name.strip() == '':
                continue

            name = str(name).strip()

            # 项目名称（第1行）
            if project_name is None and value and 'MW' in str(value):
                project_name = str(value).strip()
                continue

            # 解析各参数
            value_str = str(value).strip() if value else ''

            # 跳过标题行和空值
            if name in ('测算边界', '', ' '):
                continue
            if not value_str:
                continue

            # EPC价格（元/W）
            if 'EPC' in name and '元/W' in name:
                params['epc_price'] = float(value_str)
            # 容量（MW）
            elif '容量' in name and 'MW' in name:
                params['capacity'] = int(float(value_str))
            # 利用小时数
            elif '利用小时' in name:
                params['util_hours'] = int(float(value_str))
            # 运营年数
            elif '运营年数' in name:
                params['op_years'] = int(float(value_str))
            # 支付方式
            elif '支付方式' in name:
                params['payment_method'] = value_str
            # CFD结算电价（元/kWh）
            elif 'CFD结算电价' in name and 'POST' not in name:
                params['cfd_price'] = float(value_str)
            # POST CFD结算电价（元/kWh）
            elif 'POST CFD结算电价' in name or ('POST' in name and '结算电价' in name):
                params['post_cfd_price'] = float(value_str)
            # CFD执行年限
            elif 'CFD执行年限' in name:
                params['cfd_years'] = int(float(value_str))
            # 自用比例
            elif '自用比例' in name:
                try:
                    params['self_use_ratio'] = float(value_str)
                except:
                    params['self_use_ratio'] = 0.0
            # emc协议电价
            elif 'emc协议电价' in name or 'EMC协议电价' in name:
                try:
                    params['emc_price'] = float(value_str)
                except:
                    params['emc_price'] = 0.0
            # 建设期利率
            elif '建设期利率' in name:
                params['construction_rate'] = float(value_str)
            # 运营期利率
            elif '运营期利率' in name:
                params['loan_rate'] = float(value_str)
            # 融资期限（年）
            elif '融资期限' in name:
                params['loan_years'] = int(float(value_str))
            # 运维费用（元/W）
            elif '运维费用' in name:
                # 格式: "1——5：0.xx；6——10：0.xx；11——20：0.xx" 或 "1-5：0.xx；6-10：0.xx；11-20：0.xx"
                import re
                om_str = value_str
                om_rates = {}
                # Match patterns like "1——5：0.12" or "1-5：0.12" etc.
                for seg, sep in [('1——5', '1-5'), ('6——10', '6-10'), ('11——20', '11-20'),
                                  ('1-5', '1-5'), ('6-10', '6-10'), ('11-20', '11-20')]:
                    if seg in om_str:
                        # Find the number after the colon
                        pattern = seg + r'[：:]([\d.]+)'
                        m = re.search(pattern, om_str)
                        if m:
                            om_rates[sep] = float(m.group(1))
                # If we found any, fill in defaults for missing segments
                if om_rates:
                    for s in ['1-5', '6-10', '11-20']:
                        if s not in om_rates:
                            om_rates[s] = 0.0
                    params['om_rates'] = om_rates
            # 保险费（%）
            elif '保险费' in name and '%' in name:
                params['insurance_rate'] = float(value_str.replace('%','')) / 100
            # 备品备件（%）
            elif '备品备件' in name and '%' in name:
                params['spare_rate'] = float(value_str.replace('%','')) / 100
            # 限电率
            elif '限电率' in name:
                params['curtailment'] = float(value_str.replace('%','')) / 100 if '%' in value_str else float(value_str)

    # 设置默认值
    defaults = {
        'capacity': 60, 'epc_price': 5.2, 'util_hours': 1700, 'op_years': 20,
        'cfd_price': 0.35, 'post_cfd_price': 0.36, 'cfd_years': 10,
        'curtailment': 0.05, 'loan_rate': 0.035, 'loan_years': 18,
        'capital_ratio': 0.20, 'self_use_ratio': 0.0, 'emc_price': 0.0,
        'insurance_rate': 0.005, 'spare_rate': 0.001,
        'payment_method': '年付',
        'om_rates': {'1-5': 0.0, '6-10': 0.0, '11-20': 0.0},
    }
    for k, v in defaults.items():
        if k not in params:
            params[k] = v

    # 项目名称：从文件名校准或使用Excel第1行
    if not project_name:
        project_name = Path(xlsx_path).stem
    import re
    # Remove UUID hash from filename
    project_name = re.sub(r'-*[a-f0-9]{8,}.*$', '', project_name, flags=re.IGNORECASE)
    project_name = project_name.strip().replace(' ', '')

    return project_name, params


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python extract_params_from_excel.py <excel_file> [--json-only]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    json_only = '--json-only' in sys.argv

    try:
        project_name, params = read_excel_params(xlsx_path)
        if json_only:
            print(json.dumps({'project_name': project_name, 'params': params}, ensure_ascii=False, indent=2))
        else:
            print(f"项目: {project_name}")
            print(f"参数: {json.dumps(params, ensure_ascii=False, indent=2)}")
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr)
        sys.exit(1)
