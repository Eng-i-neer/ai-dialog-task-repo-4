# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app
app = create_app()

with app.app_context():
    from app.models import Order
    for wb in ['IT12507071310019', 'IT12507101310001']:
        orders = Order.query.filter_by(waybill_no=wb).order_by(Order.bill_period).all()
        for o in orders:
            p = o.bill_period.strftime('%Y%m%d') if o.bill_period else '?'
            print(f"{wb} [{p}] head={o.charge_weight_head}, tail={o.charge_weight_tail}, "
                  f"actual={o.actual_weight}, postcode={o.postcode}, address={o.address}")
