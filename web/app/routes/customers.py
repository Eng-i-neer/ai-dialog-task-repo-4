from flask import render_template, request, jsonify
from app.routes import customers_bp
from app import db
from app.models import Customer


@customers_bp.route('/')
def customer_list():
    customers = Customer.query.order_by(Customer.name).all()
    return render_template('customers.html', customers=customers)


@customers_bp.route('/api', methods=['GET'])
def api_list():
    customers = Customer.query.order_by(Customer.name).all()
    return jsonify([c.to_dict() for c in customers])


@customers_bp.route('/api', methods=['POST'])
def api_create():
    data = request.get_json() or {}
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': '客户名称不能为空'}), 400

    if Customer.query.filter_by(name=name).first():
        return jsonify({'error': '客户名称已存在'}), 400

    c = Customer(
        name=name,
        code=data.get('code', ''),
        currency=data.get('currency', 'EUR'),
        notes=data.get('notes', '')
    )
    db.session.add(c)
    db.session.commit()
    return jsonify(c.to_dict()), 201


@customers_bp.route('/api/<int:cid>', methods=['PUT'])
def api_update(cid):
    c = Customer.query.get_or_404(cid)
    data = request.get_json() or {}

    new_name = data.get('name', '').strip()
    if new_name and new_name != c.name:
        existing = Customer.query.filter_by(name=new_name).first()
        if existing and existing.id != cid:
            return jsonify({'error': '客户名称已存在'}), 400
        c.name = new_name
    if 'code' in data:
        c.code = data['code']
    if 'currency' in data:
        c.currency = data['currency']
    if 'notes' in data:
        c.notes = data['notes']

    db.session.commit()
    return jsonify(c.to_dict())


@customers_bp.route('/api/<int:cid>', methods=['DELETE'])
def api_delete(cid):
    c = Customer.query.get_or_404(cid)
    if c.orders.count() > 0:
        return jsonify({'error': '该客户下有订单记录，无法删除'}), 400
    db.session.delete(c)
    db.session.commit()
    return jsonify({'success': True})
