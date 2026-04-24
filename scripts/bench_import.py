# -*- coding: utf-8 -*-
"""Benchmark import speed after optimizations."""
import sys, os, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
app = create_app()

TEST_FILES = [
    r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.07.28\（李志）鑫腾跃-中文-对账单20250728.xlsx',
    r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.08.05\（李志）鑫腾跃-中文-对账单20250805.xlsx',
]

with app.app_context():
    from app.services.excel_parser import parse_agent_bill
    from app.models import ImportLog, Order

    for f in TEST_FILES:
        fname = os.path.basename(f)
        existing = Order.query.filter_by(bill_period=fname.split('对账单')[-1][:8] if '对账单' in fname else None).count()

        log = ImportLog(filename=f'bench_{fname}', file_type='agent_bill', status='uploaded')
        db.session.add(log)
        db.session.commit()

        t0 = time.time()
        result = parse_agent_bill(f, log.id, customer_id=None)
        elapsed = time.time() - t0

        print(f"{fname}")
        print(f"  耗时: {elapsed:.1f}s")
        print(f"  订单: {result.get('orders_count', 0)} (DB已有: {existing})")
        print(f"  Sheet行: {result.get('sheets_processed', 0)}")
        print()

        db.session.rollback()

    print("基准测试完成")
