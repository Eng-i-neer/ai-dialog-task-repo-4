from flask import render_template, request, jsonify
from app.routes import orders_bp
from app import db
from app.models import Order, Customer, Region, OrderFee, FeeCategory
from sqlalchemy import func


def _order_fee_total(order_id):
    """Sum the effective fee for each category of an order."""
    fees = OrderFee.query.filter_by(order_id=order_id).all()
    total = 0.0
    for f in fees:
        val = f.override_amount if f.override_amount is not None else f.calculated_amount
        if val is not None:
            total += val
    return round(total, 2)


FEE_LABELS = {
    'HEAD_FREIGHT': '头程',
    'TAIL_FREIGHT': '尾程',
    'COD_FEE': 'COD手续费',
    'RETURN_FEE': '退件',
    'SHELF_FEE': '上架',
    'F_SURCHARGE': '附加',
    'VAT': '增值税',
    'REMOTE_FEE': '偏远',
    'SECOND_DELIVERY': '二派费',
}

FEE_FILTER_OPTIONS = [
    ('HEAD_FREIGHT', '有头程运费'),
    ('TAIL_FREIGHT', '有尾程运费'),
    ('HAS_COD', '有COD代收'),
    ('RETURN_FEE', '需退件费'),
    ('SHELF_FEE', '需上架费'),
    ('VAT', '需增值税'),
    ('F_SURCHARGE', '有附加费'),
    ('REMOTE_FEE', '偏远订单'),
    ('SECOND_DELIVERY', '有二派费'),
]


@orders_bp.route('/')
def order_list():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 50, type=int)
    search = request.args.get('q', '').strip()
    customer_id = request.args.get('customer_id', type=int)
    region_id = request.args.get('region_id', type=int)
    bill_period = request.args.get('bill_period', '')
    category_filter = request.args.get('category', '').strip()
    cargo_type_filter = request.args.get('cargo_type', '').strip()
    remote_filter = request.args.get('remote', '').strip()

    query = Order.query

    if search:
        query = query.filter(Order.waybill_no.contains(search))
    if customer_id:
        query = query.filter(Order.customer_id == customer_id)
    if region_id:
        query = query.filter(Order.region_id == region_id)
    if bill_period:
        from datetime import date
        try:
            parts = bill_period.split('-')
            bp = date(int(parts[0]), int(parts[1]), int(parts[2]))
            query = query.filter(Order.bill_period == bp)
        except (ValueError, IndexError):
            pass
    if cargo_type_filter:
        query = query.filter(Order.cargo_type == cargo_type_filter)
    if remote_filter == '1':
        query = query.filter(Order.is_remote == True)
    elif remote_filter == '0':
        query = query.filter(db.or_(Order.is_remote == False, Order.is_remote.is_(None)))

    if category_filter:
        if category_filter == 'HEAD_FREIGHT':
            query = query.filter(Order.has_head_freight == True)
        elif category_filter == 'TAIL_FREIGHT':
            query = query.filter(Order.has_tail_freight == True)
        elif category_filter == 'HAS_COD':
            query = query.filter(Order.import_sheets.contains('COD'))
        elif category_filter == 'RETURN_FEE':
            query = query.filter(Order.needs_return_fee == True)
        elif category_filter == 'SHELF_FEE':
            query = query.filter(Order.needs_shelf_fee == True)
        elif category_filter == 'VAT':
            query = query.filter(Order.needs_vat == True)
        elif category_filter == 'F_SURCHARGE':
            query = query.filter(db.or_(Order.has_head_freight == True, Order.has_tail_freight == True))
        elif category_filter == 'REMOTE_FEE':
            query = query.filter(Order.is_remote == True)
        elif category_filter == 'SECOND_DELIVERY':
            query = query.filter(Order.needs_second_delivery == True)

    query = query.order_by(Order.created_at.desc())
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)

    fee_totals = {}
    for o in pagination.items:
        fee_totals[o.id] = _order_fee_total(o.id)

    customers = Customer.query.order_by(Customer.name).all()
    regions = Region.query.order_by(Region.name).all()

    current_customer = Customer.query.get(customer_id) if customer_id else None

    return render_template('orders.html',
        orders=pagination.items,
        pagination=pagination,
        customers=customers,
        regions=regions,
        fee_labels=FEE_LABELS,
        fee_filter_options=FEE_FILTER_OPTIONS,
        search=search,
        customer_id=customer_id,
        region_id=region_id,
        bill_period=bill_period,
        category_filter=category_filter,
        cargo_type_filter=cargo_type_filter,
        remote_filter=remote_filter,
        fee_totals=fee_totals,
        current_customer=current_customer,
    )


@orders_bp.route('/<int:order_id>')
def order_detail(order_id):
    order = Order.query.get_or_404(order_id)
    fees = OrderFee.query.filter_by(order_id=order_id).all()

    has_calculated = any(f.calculated_amount is not None for f in fees)
    if not has_calculated:
        try:
            from app.services.pricing_engine import calculate_order_fees
            calculate_order_fees(order_id)
            fees = OrderFee.query.filter_by(order_id=order_id).all()
        except Exception:
            pass

    categories = FeeCategory.query.all()

    return render_template('order_detail.html',
        order=order, fees=fees, categories=categories)


@orders_bp.route('/api/list')
def api_order_list():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 50, type=int)
    search = request.args.get('q', '').strip()
    customer_id = request.args.get('customer_id', type=int)
    region_id = request.args.get('region_id', type=int)
    bill_period = request.args.get('bill_period', '')

    query = Order.query
    if search:
        query = query.filter(Order.waybill_no.contains(search))
    if customer_id:
        query = query.filter(Order.customer_id == customer_id)
    if region_id:
        query = query.filter(Order.region_id == region_id)
    if bill_period:
        from datetime import date
        try:
            parts = bill_period.split('-')
            bp = date(int(parts[0]), int(parts[1]), int(parts[2]))
            query = query.filter(Order.bill_period == bp)
        except (ValueError, IndexError):
            pass

    pagination = query.order_by(Order.created_at.desc()).paginate(
        page=page, per_page=per_page, error_out=False)

    return jsonify({
        'orders': [o.to_dict() for o in pagination.items],
        'total': pagination.total,
        'pages': pagination.pages,
        'page': page
    })


@orders_bp.route('/api/<int:order_id>/fees', methods=['POST'])
def update_fee(order_id):
    data = request.get_json() or {}
    fee_id = data.get('fee_id')
    if not fee_id:
        return jsonify({'error': 'fee_id 不能为空'}), 400

    fee = OrderFee.query.filter_by(id=fee_id, order_id=order_id).first_or_404()
    override = data.get('override_amount')
    if override is not None:
        try:
            fee.override_amount = float(override)
        except (ValueError, TypeError):
            return jsonify({'error': '金额格式错误'}), 400
        fee.is_manual = True
    else:
        fee.override_amount = None
        fee.is_manual = False

    db.session.commit()
    return jsonify({'success': True, 'fee': fee.to_dict()})


@orders_bp.route('/api/<int:order_id>', methods=['DELETE'])
def delete_order(order_id):
    order = Order.query.get_or_404(order_id)
    db.session.delete(order)
    db.session.commit()
    return jsonify({'success': True})


@orders_bp.route('/api/batch-calculate', methods=['POST'])
def batch_calculate():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    category_codes = data.get('category_codes')

    if not order_ids:
        return jsonify({'error': '请选择订单'}), 400

    from app.services.pricing_engine import batch_calculate as do_batch
    results = do_batch(order_ids, category_codes)
    return jsonify({'success': True, 'results': {str(k): v for k, v in results.items()}})


@orders_bp.route('/api/batch-override', methods=['POST'])
def batch_override():
    """Batch set override_amount for a specific category on selected orders."""
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    category_code = data.get('category_code')
    amount = data.get('amount')

    if not order_ids or not category_code:
        return jsonify({'error': '参数不完整'}), 400

    cat = FeeCategory.query.filter_by(code=category_code).first()
    if not cat:
        return jsonify({'error': f'未知科目: {category_code}'}), 400

    updated = 0
    for oid in order_ids:
        fee = OrderFee.query.filter_by(order_id=oid, category_id=cat.id).first()
        if fee:
            if amount is not None:
                try:
                    fee.override_amount = float(amount)
                except (ValueError, TypeError):
                    continue
                fee.is_manual = True
            else:
                fee.override_amount = None
                fee.is_manual = False
            updated += 1

    db.session.commit()
    return jsonify({'success': True, 'updated': updated})


@orders_bp.route('/api/batch-delete', methods=['POST'])
def batch_delete():
    """Batch delete selected orders."""
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    if not order_ids:
        return jsonify({'error': '请选择订单'}), 400

    deleted = 0
    for oid in order_ids:
        order = Order.query.get(oid)
        if order:
            db.session.delete(order)
            deleted += 1

    db.session.commit()
    return jsonify({'success': True, 'deleted': deleted})


@orders_bp.route('/api/periods')
def api_periods():
    """Get all existing bill periods (for edition selector)."""
    customer_id = request.args.get('customer_id', type=int)
    
    query = db.session.query(
        Order.bill_period,
        func.count(Order.id).label('order_count')
    ).filter(Order.bill_period.isnot(None))
    
    if customer_id:
        query = query.filter(Order.customer_id == customer_id)
    
    results = query.group_by(Order.bill_period)\
        .order_by(Order.bill_period.desc())\
        .all()
    
    periods = []
    for bp, count in results:
        if bp:
            periods.append({
                'date': bp.strftime('%Y-%m-%d'),
                'label': bp.strftime('%m%d'),
                'full_label': bp.strftime('%m%d期'),
                'order_count': count
            })
    
    return jsonify(periods)
