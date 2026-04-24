# -*- coding: utf-8 -*-
"""Test cross-period fee calculation for Italian orders."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
app = create_app()

WAYBILLS = ['IT12507101310001', 'IT12507071310019']

with app.app_context():
    from app.models import Order, OrderFee
    from app.services.pricing_engine import calculate_order_fees, _get_cross_period_merged, _merged_applicable_fees

    for wb_no in WAYBILLS:
        print(f"\n{'='*60}")
        print(f"运单号: {wb_no}")
        print(f"{'='*60}")

        orders = Order.query.filter_by(waybill_no=wb_no).order_by(Order.bill_period).all()
        for o in orders:
            period = o.bill_period.strftime('%Y%m%d') if o.bill_period else '?'
            print(f"\n  --- {period}期 (order.id={o.id}) ---")

            merged = _get_cross_period_merged(o)
            merged_fees = _merged_applicable_fees(o, merged)
            print(f"    merged flags: head={merged['has_head_freight']}, tail={merged['has_tail_freight']}, "
                  f"cod={merged['cod_amount']}, region={merged['region_code']}")
            print(f"    merged applicable_fees: {merged_fees}")
            print(f"    original applicable_fees: {o.applicable_fees}")

            print(f"    重新计算费用...")
            results = calculate_order_fees(o.id)
            print(f"    计算结果: {results}")

            fees = OrderFee.query.filter_by(order_id=o.id).all()
            print(f"    当前费用记录 ({len(fees)} 条):")
            for f in fees:
                cat_name = f.category.name if f.category else '?'
                cat_code = f.category.code if f.category else '?'
                print(f"      {cat_code} ({cat_name}): input={f.input_amount}, calc={f.calculated_amount}, "
                      f"override={f.override_amount}")
