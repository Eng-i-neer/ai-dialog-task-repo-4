# -*- coding: utf-8 -*-
"""Clean up ghost fee records from wrong previous calculations on cross-period orders."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
app = create_app()

with app.app_context():
    from app.models import Order, OrderFee

    orders_0811 = Order.query.filter(
        Order.bill_period == '2025-08-11'
    ).all()

    deleted = 0
    for o in orders_0811:
        applicable = set(o.applicable_fees)
        fees = OrderFee.query.filter_by(order_id=o.id).all()
        for f in fees:
            if f.category and f.category.code not in applicable:
                if f.category.code in ('HEAD_FREIGHT', 'TAIL_FREIGHT', 'F_SURCHARGE') and f.input_amount is None:
                    print(f"  删除 order={o.id} ({o.waybill_no}) 的多余费用: {f.category.code}")
                    db.session.delete(f)
                    deleted += 1

    db.session.commit()
    print(f"\n共清理 {deleted} 条多余费用记录")
