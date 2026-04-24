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
    unique_waybills = db.session.query(func.count(func.distinct(Order.waybill_no))).scalar()
    dupes = db.session.query(
        Order.waybill_no, func.count(Order.id).label('cnt')
    ).group_by(Order.waybill_no).having(func.count(Order.id) > 1).all()

    print(f"总订单数: {total}")
    print(f"唯一运单号: {unique_waybills}")
    print(f"有重复的运单号: {len(dupes)}")
    if dupes:
        for wb, cnt in dupes[:10]:
            orders = Order.query.filter_by(waybill_no=wb).order_by(Order.bill_period).all()
            periods = [o.bill_period.strftime('%Y%m%d') for o in orders]
            print(f"  {wb}: {cnt} 条, periods={periods}")

    from app import db
    total_fees = OrderFee.query.count()
    print(f"\n总费用记录: {total_fees}")
