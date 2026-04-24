# -*- coding: utf-8 -*-
"""
Migrate: merge multiple Order records with the same waybill_no into one.

Steps:
1. For each waybill_no that has >1 Order, pick the earliest (by bill_period)
   as the "primary" order.
2. Merge all data from secondary orders into the primary:
   - Combine import_sheets, import_periods
   - Fill missing fields (charge_weight, cod_amount, etc.)
   - Transfer OrderFee records
   - Merge boolean flags (OR them together)
3. Delete secondary orders (fees already transferred).
4. Drop the old unique constraint and add new one (waybill_no only).
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app, db
app = create_app()

with app.app_context():
    from app.models import Order, OrderFee
    from sqlalchemy import func, text

    dupes = db.session.query(
        Order.waybill_no, func.count(Order.id).label('cnt')
    ).group_by(Order.waybill_no).having(func.count(Order.id) > 1).all()

    print(f"需合并的运单号: {len(dupes)}")
    merged_count = 0
    deleted_count = 0
    fees_transferred = 0

    for waybill_no, cnt in dupes:
        orders = Order.query.filter_by(waybill_no=waybill_no).order_by(Order.bill_period).all()
        primary = orders[0]
        secondaries = orders[1:]

        all_periods = set()
        all_sheets = set()

        for o in orders:
            if o.bill_period:
                all_periods.add(o.bill_period.strftime('%Y%m%d'))
            if o.import_sheets:
                for s in o.import_sheets.split(','):
                    all_sheets.add(s.strip())

        for sec in secondaries:
            if not primary.charge_weight_head and sec.charge_weight_head:
                primary.charge_weight_head = sec.charge_weight_head
            if not primary.charge_weight_tail and sec.charge_weight_tail:
                primary.charge_weight_tail = sec.charge_weight_tail
            if not primary.actual_weight and sec.actual_weight:
                primary.actual_weight = sec.actual_weight
            if not primary.transfer_no and sec.transfer_no:
                primary.transfer_no = sec.transfer_no
            if not primary.customer_ref and sec.customer_ref:
                primary.customer_ref = sec.customer_ref
            if not primary.postcode and sec.postcode:
                primary.postcode = sec.postcode
            if not primary.address and sec.address:
                primary.address = sec.address
            if not primary.dimensions and sec.dimensions:
                primary.dimensions = sec.dimensions
            if sec.cod_amount:
                primary.cod_amount = sec.cod_amount
                primary.cod_currency = sec.cod_currency or 'EUR'
            if sec.has_head_freight:
                primary.has_head_freight = True
            if sec.has_tail_freight:
                primary.has_tail_freight = True
            if sec.needs_return_fee:
                primary.needs_return_fee = True
            if sec.needs_shelf_fee:
                primary.needs_shelf_fee = True
            if sec.needs_vat:
                primary.needs_vat = True
            if sec.is_remote:
                primary.is_remote = True

            sec_fees = OrderFee.query.filter_by(order_id=sec.id).all()
            for fee in sec_fees:
                existing = OrderFee.query.filter_by(
                    order_id=primary.id,
                    category_id=fee.category_id,
                    source_sheet=fee.source_sheet,
                    import_period=fee.import_period
                ).first()

                if not existing:
                    fee.order_id = primary.id
                    fees_transferred += 1
                else:
                    if fee.input_amount and not existing.input_amount:
                        existing.input_amount = fee.input_amount
                    if fee.calculated_amount and not existing.calculated_amount:
                        existing.calculated_amount = fee.calculated_amount
                    db.session.delete(fee)

            db.session.flush()

            db.session.delete(sec)
            deleted_count += 1

        primary.import_periods = ','.join(sorted(all_periods))
        primary.import_sheets = ','.join(sorted(all_sheets))
        merged_count += 1

    for o in Order.query.filter(Order.import_periods.is_(None)).all():
        if o.bill_period:
            o.import_periods = o.bill_period.strftime('%Y%m%d')

    db.session.commit()

    print(f"合并完成:")
    print(f"  合并运单数: {merged_count}")
    print(f"  删除重复记录: {deleted_count}")
    print(f"  转移费用记录: {fees_transferred}")
    print(f"  剩余订单数: {Order.query.count()}")
    print(f"  剩余费用数: {OrderFee.query.count()}")

    print("\n修改数据库约束...")
    try:
        db.session.execute(text("DROP INDEX IF EXISTS uq_waybill_period"))
        db.session.commit()
        print("  已删除旧的 uq_waybill_period 约束")
    except Exception as e:
        db.session.rollback()
        print(f"  删除旧约束失败(可能已不存在): {e}")

    try:
        db.session.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS ix_orders_waybill_no ON orders (waybill_no)"))
        db.session.commit()
        print("  已创建新的 waybill_no 唯一索引")
    except Exception as e:
        db.session.rollback()
        print(f"  创建索引失败: {e}")

    print("\n验证...")
    dup_check = db.session.query(
        Order.waybill_no, func.count(Order.id)
    ).group_by(Order.waybill_no).having(func.count(Order.id) > 1).all()
    print(f"  重复运单号: {len(dup_check)} (应为0)")
