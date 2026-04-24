# -*- coding: utf-8 -*-
"""Remove duplicate OrderFee records (same order + same category, keep the one with import_period)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app, db
app = create_app()

with app.app_context():
    from app.models import Order, OrderFee
    from sqlalchemy import func

    dupes = db.session.query(
        OrderFee.order_id, OrderFee.category_id, func.count(OrderFee.id)
    ).group_by(OrderFee.order_id, OrderFee.category_id
    ).having(func.count(OrderFee.id) > 1).all()

    print(f"有重复费用的订单+科目组合: {len(dupes)}")
    deleted = 0

    for order_id, cat_id, cnt in dupes:
        fees = OrderFee.query.filter_by(order_id=order_id, category_id=cat_id
        ).order_by(OrderFee.import_period.desc().nullslast(), OrderFee.id).all()

        keep = fees[0]
        for dup in fees[1:]:
            if dup.input_amount and not keep.input_amount:
                keep.input_amount = dup.input_amount
            if dup.calculated_amount and not keep.calculated_amount:
                keep.calculated_amount = dup.calculated_amount
            if dup.import_period and not keep.import_period:
                keep.import_period = dup.import_period
            if dup.source_sheet and not keep.source_sheet:
                keep.source_sheet = dup.source_sheet
            db.session.delete(dup)
            deleted += 1

    db.session.commit()
    print(f"删除重复费用: {deleted}")
    print(f"剩余费用总数: {OrderFee.query.count()}")
