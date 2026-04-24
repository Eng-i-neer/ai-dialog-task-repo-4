# -*- coding: utf-8 -*-
"""Test the parser against multiple bill formats."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')

os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
from app.models import Order, OrderFee, ImportLog, FeeCategory

app = create_app()

TEST_FILES = [
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.07.28\（李志）鑫腾跃-中文-对账单20250728.xlsx', '布局1: 标准17列'),
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.08.05\（李志）鑫腾跃-中文-对账单20250805.xlsx', '布局3: 20列(含寄件人等)'),
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.08.11\（李志）鑫腾跃-中文-对账单20250811.xlsx', '布局4: 缺计算公式列'),
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.08.18\（李志）鑫腾跃-中文-对账单20250818.xlsx', '布局混合: COD列序不同'),
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.09.01\（李志）鑫腾跃-中文-对账单20250901.xlsx', '布局1: 含偏远费Sheet'),
    (r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.12.22\李志-鑫腾跃-中文-对账单20251222.xlsx', '布局2: 含实重列18列'),
]

with app.app_context():
    from app.services.excel_parser import _find_header_row, _build_col_map, _match_sheet_type
    from app.services.excel_utils import load_excel

    print("=" * 80)
    print("测试解析器: _find_header_row + _build_col_map")
    print("=" * 80)

    for fpath, desc in TEST_FILES:
        fname = os.path.basename(fpath)
        print(f"\n{'─'*80}")
        print(f"📄 {fname}")
        print(f"   {desc}")

        wb = load_excel(fpath, data_only=True)
        total_sheets = 0
        parsed_sheets = 0
        total_rows = 0
        waybill_samples = []
        issues = []

        for sname in wb.sheetnames:
            ws = wb[sname]
            stype = _match_sheet_type(sname)
            if not stype:
                continue

            total_sheets += 1
            hr = _find_header_row(ws)
            if not hr:
                issues.append(f"  ⚠ [{sname}] 找不到表头行")
                continue

            col_map = _build_col_map(ws, hr)
            if 'waybill' not in col_map:
                issues.append(f"  ⚠ [{sname}] hr={hr} 找不到运单号列, col_map={col_map}")
                continue

            parsed_sheets += 1
            sheet_rows = 0
            for row_idx in range(hr + 1, (ws.max_row or 0) + 1):
                first_cell = ws.cell(row_idx, 1).value
                if first_cell and str(first_cell).strip().startswith('合计'):
                    break
                wv = ws.cell(row_idx, col_map['waybill']).value
                if not wv:
                    continue
                wv = str(wv).strip()
                if not wv or wv == 'None' or wv.startswith('合计'):
                    break
                sheet_rows += 1
                if len(waybill_samples) < 3:
                    country_col = col_map.get('country', 0)
                    country = ws.cell(row_idx, country_col).value if country_col else '?'
                    waybill_samples.append(f"{wv} ({country})")

            total_rows += sheet_rows
            mapped_fields = ', '.join(sorted(col_map.keys()))

        print(f"   Sheets: {parsed_sheets}/{total_sheets} 成功解析")
        print(f"   数据行: {total_rows} 条")
        print(f"   运单样本: {waybill_samples}")
        if issues:
            for iss in issues:
                print(iss)

        wb.close()

    print(f"\n{'='*80}")
    print("测试完成")
