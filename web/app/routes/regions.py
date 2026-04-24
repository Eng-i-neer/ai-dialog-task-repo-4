from flask import render_template, request, jsonify
from app.routes import regions_bp
from app import db
from app.models import Region


@regions_bp.route('/')
def region_list():
    regions = Region.query.order_by(Region.name).all()
    return render_template('regions.html', regions=regions)


@regions_bp.route('/api', methods=['GET'])
def api_list():
    regions = Region.query.order_by(Region.name).all()
    return jsonify([r.to_dict() for r in regions])


@regions_bp.route('/api', methods=['POST'])
def api_create():
    data = request.get_json() or {}
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': '地区名称不能为空'}), 400

    if Region.query.filter_by(name=name).first():
        return jsonify({'error': '地区名称已存在'}), 400

    vat = data.get('vat_rate')
    try:
        vat = float(vat) if vat is not None and vat != '' else None
    except (ValueError, TypeError):
        return jsonify({'error': 'VAT税率格式错误'}), 400

    r = Region(
        name=name,
        code=data.get('code', ''),
        currency=data.get('currency', ''),
        vat_rate=vat,
        return_rule=data.get('return_rule', '100%')
    )
    db.session.add(r)
    db.session.commit()
    return jsonify(r.to_dict()), 201


@regions_bp.route('/api/<int:rid>', methods=['PUT'])
def api_update(rid):
    r = Region.query.get_or_404(rid)
    data = request.get_json() or {}

    for field in ['name', 'code', 'currency', 'return_rule']:
        if field in data:
            setattr(r, field, data[field])
    if 'vat_rate' in data:
        vat = data['vat_rate']
        try:
            r.vat_rate = float(vat) if vat is not None and vat != '' else None
        except (ValueError, TypeError):
            return jsonify({'error': 'VAT税率格式错误'}), 400

    db.session.commit()
    return jsonify(r.to_dict())


@regions_bp.route('/api/<int:rid>', methods=['DELETE'])
def api_delete(rid):
    r = Region.query.get_or_404(rid)
    if r.orders.count() > 0:
        return jsonify({'error': '该地区下有订单记录，无法删除'}), 400
    db.session.delete(r)
    db.session.commit()
    return jsonify({'success': True})
