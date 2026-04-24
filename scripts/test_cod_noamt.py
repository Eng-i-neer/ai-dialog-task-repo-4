# -*- coding: utf-8 -*-
"""Find an Italian order from 0728 that does NOT have COD in any period, and test."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app, db
app = create_app()

with app.app_context():
    from app.models import Order, OrderFee, Region
    from app.services.pricing_engine import calculate_order_fees, _get_cross_period_merged, _merged_applicable_fees

    it_region = Region.query.filter_by(code='IT').first()
    if not it_region:
        print("No IT region!")
        sys.exit()

    it_orders = Order.query.filter_by(
        region_id=it_region.id,
        has_head_freight=True
    ).filter(
        Order.bill_period == '2025-07-28'
    ).limit(5).all()

    for o in it_orders:
        all_periods = Order.query.filter_by(waybill_no=o.waybill_no).all()
        any_cod = any(sib.cod_amount for sib in all_periods)
        has_cod_sheet = any('COD' in (sib.import_sheets or '') for sib in all_periods)

        if not has_cod_sheet:
            print(f"\n运单 {o.waybill_no} (id={o.id}): 无COD表, cod_amount={o.cod_amount}")
            print(f"  periods: {[(sib.bill_period.strftime('%Y%m%d'), sib.import_sheets) for sib in all_periods]}")

            merged = _get_cross_period_merged(o)
            applicable = _merged_applicable_fees(o, merged)
            print(f"  merged applicable: {applicable}")

            results = calculate_order_fees(o.id)
            for r in results:
                print(f"    {r['category']}: {r['amount']} ({r['source']})")

            cod_fee = OrderFee.query.filter_by(order_id=o.id).join(OrderFee.category).filter_by(code='COD_FEE').first()
            if cod_fee:
                print(f"  COD_FEE record: input={cod_fee.input_amount}, calc={cod_fee.calculated_amount}")
            break
