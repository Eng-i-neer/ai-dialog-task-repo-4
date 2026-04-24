# -*- coding: utf-8 -*-
import sys, os, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
app = create_app()

with app.app_context():
    from app.services.excel_parser import parse_agent_bill
    from app.models import ImportLog, Order

    f = r'c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\中邮原账单\2025.07.28\（李志）鑫腾跃-中文-对账单20250728.xlsx'

    existing = Order.query.filter_by(bill_period='2025-07-28').count()
    print(f"DB中已有 2025-07-28 期订单: {existing} 条")

    log = ImportLog(filename='test_reimport', file_type='agent_bill', status='uploaded')
    db.session.add(log)
    db.session.commit()

    t0 = time.time()
    result = parse_agent_bill(f, log.id, customer_id=None)
    t1 = time.time()

    print(f"解析耗时: {t1 - t0:.1f}s")
    print(f"订单数: {result.get('orders_count', 0)}")
    print(f"Sheet处理数: {result.get('sheets_processed', 0)}")
    print(f"跨期事件: {len(result.get('cross_period_events', []))}")

    db.session.rollback()
    print("测试完成 (已回滚)")
