from flask import render_template
from app.routes import main_bp
from app import db
from app.models import Order, Customer, ImportLog, OrderFee
from sqlalchemy import func


@main_bp.route('/')
def dashboard():
    total_orders = db.session.query(func.count(Order.id)).scalar() or 0
    total_customers = db.session.query(func.count(Customer.id)).scalar() or 0

    pending = db.session.query(func.count(Order.id)).filter(
        Order.logistics_status == '待处理'
    ).scalar() or 0

    exported = db.session.query(func.count(Order.id)).filter(
        Order.logistics_status == '已导出'
    ).scalar() or 0

    customer_periods = _build_customer_period_summary()

    recent_imports = ImportLog.query.order_by(ImportLog.created_at.desc()).limit(5).all()

    return render_template('dashboard.html',
        total_orders=total_orders,
        total_customers=total_customers,
        pending=pending,
        exported=exported,
        customer_periods=customer_periods,
        recent_imports=recent_imports
    )


def _build_customer_period_summary():
    """Build a nested structure: {customer: [{period, count, pending, exported, fee_total}]}"""
    results = db.session.query(
        Customer.id,
        Customer.name,
        Customer.currency,
        Order.bill_period,
        func.count(Order.id).label('order_count'),
        func.sum(
            db.case(
                (Order.logistics_status == '待处理', 1),
                else_=0
            )
        ).label('pending_count'),
        func.sum(
            db.case(
                (Order.logistics_status == '已导出', 1),
                else_=0
            )
        ).label('exported_count'),
    ).join(Customer, Order.customer_id == Customer.id)\
     .group_by(Customer.id, Customer.name, Customer.currency, Order.bill_period)\
     .order_by(Customer.name, Order.bill_period.desc())\
     .all()

    customer_map = {}
    for row in results:
        cid = row[0]
        if cid not in customer_map:
            customer_map[cid] = {
                'id': cid,
                'name': row[1],
                'currency': row[2],
                'periods': [],
                'total_orders': 0,
            }
        period_str = row[3].strftime('%m%d') if row[3] else '未知'
        period_date = row[3].strftime('%Y-%m-%d') if row[3] else ''
        customer_map[cid]['periods'].append({
            'label': period_str,
            'date': period_date,
            'count': row[4],
            'pending': row[5] or 0,
            'exported': row[6] or 0,
        })
        customer_map[cid]['total_orders'] += row[4]

    return sorted(customer_map.values(), key=lambda c: -c['total_orders'])
