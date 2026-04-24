# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app, db
app = create_app()

with app.app_context():
    from app.models import Order, OrderFee
    from sqlalchemy import func

    total = Order.query.count()
    dupes = db.session.query(
        Order.waybill_no, func.count(Order.id)
    ).group_by(Order.waybill_no).having(func.count(Order.id) > 1).count()
    print(f"订单总数: {total}, 重复运单: {dupes}")

    for wb in ['IT12507071310019', 'IT12507101310001']:
        o = Order.query.filter_by(waybill_no=wb).first()
        if o:
            print(f"\n{wb} (id={o.id}):")
            print(f"  periods: {o.import_period_list}")
            print(f"  sheets: {o.import_sheet_list}")
            print(f"  head={o.has_head_freight}, tail={o.has_tail_freight}")
            print(f"  head_w={o.charge_weight_head}, tail_w={o.charge_weight_tail}")
            print(f"  cod={o.cod_amount} {o.cod_currency}")
            fees = OrderFee.query.filter_by(order_id=o.id).all()
            print(f"  费用 ({len(fees)}):")
            for f in fees:
                print(f"    {f.category.code}: calc={f.calculated_amount}, input={f.input_amount}, period={f.import_period}")
