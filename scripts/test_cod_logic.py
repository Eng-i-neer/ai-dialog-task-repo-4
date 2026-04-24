# -*- coding: utf-8 -*-
"""Test the new COD fee calculation logic."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app, db
app = create_app()

WAYBILLS = ['IT12507101310001', 'IT12507071310019']

with app.app_context():
    from app.models import Order, OrderFee
    from app.services.pricing_engine import calculate_order_fees

    for wb_no in WAYBILLS:
        print(f"\n{'='*60}")
        print(f"运单号: {wb_no}")
        print(f"{'='*60}")

        orders = Order.query.filter_by(waybill_no=wb_no).order_by(Order.bill_period).all()
        for o in orders:
            period = o.bill_period.strftime('%Y%m%d') if o.bill_period else '?'
            print(f"\n  --- {period}期 (order.id={o.id}) ---")
            print(f"    has_head={o.has_head_freight}, has_tail={o.has_tail_freight}")
            print(f"    cod_amount={o.cod_amount}, region={o.region.name if o.region else '?'}")
            print(f"    import_sheets={o.import_sheets}")

            results = calculate_order_fees(o.id)
            print(f"    计算结果:")
            for r in results:
                print(f"      {r['category']}: {r['amount']} ({r['source']})")

            fees = OrderFee.query.filter_by(order_id=o.id).all()
            for f in fees:
                if f.category and f.category.code == 'COD_FEE':
                    print(f"    COD_FEE 详情: input={f.input_amount}, calc={f.calculated_amount}")
