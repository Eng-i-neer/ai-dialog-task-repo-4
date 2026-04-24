from flask import render_template, request, jsonify
from app.routes import pricing_bp
from app import db
from app.models import (PricingVersion, PricingRule, FeeCategory, Region,
                         CustomerPricingOverride, Customer, RemotePostcode)


@pricing_bp.route('/')
def pricing_page():
    versions = PricingVersion.query.order_by(PricingVersion.effective_date.desc()).all()
    categories = FeeCategory.query.order_by(FeeCategory.code).all()
    regions = Region.query.order_by(Region.name).all()
    customers = Customer.query.order_by(Customer.name).all()
    return render_template('pricing.html',
        versions=versions, categories=categories,
        regions=regions, customers=customers)


@pricing_bp.route('/remote-postcodes')
def remote_postcodes_page():
    versions = PricingVersion.query.order_by(PricingVersion.effective_date.desc()).all()
    selected_version_id = request.args.get('version_id', type=int)
    return render_template('remote_postcodes.html',
                           versions=versions,
                           selected_version_id=selected_version_id)


# ---------- Version APIs ----------

@pricing_bp.route('/api/versions', methods=['GET'])
def api_versions():
    versions = PricingVersion.query.order_by(PricingVersion.effective_date.desc()).all()
    return jsonify([v.to_dict() for v in versions])


@pricing_bp.route('/api/versions', methods=['POST'])
def api_create_version():
    data = request.get_json() or {}
    name = data.get('name', '').strip()
    eff_date = data.get('effective_date', '')
    if not name or not eff_date:
        return jsonify({'error': '名称和生效日期不能为空'}), 400

    from datetime import date as dt_date
    try:
        parts = eff_date.split('-')
        ed = dt_date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        return jsonify({'error': '日期格式错误'}), 400

    v = PricingVersion(name=name, effective_date=ed, notes=data.get('notes', ''))
    exp = data.get('expire_date')
    if exp:
        try:
            parts = exp.split('-')
            v.expire_date = dt_date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            pass

    db.session.add(v)
    db.session.commit()
    return jsonify(v.to_dict()), 201


@pricing_bp.route('/api/versions/<int:version_id>/activate', methods=['POST'])
def api_activate_version(version_id):
    """Set this version as the only active one."""
    target = PricingVersion.query.get_or_404(version_id)

    PricingVersion.query.filter(PricingVersion.id != version_id).update({'is_active': False})
    target.is_active = True
    db.session.commit()
    return jsonify({'success': True, 'message': f'已激活 {target.name}'})


@pricing_bp.route('/api/versions/<int:version_id>', methods=['DELETE'])
def api_delete_version(version_id):
    v = PricingVersion.query.get_or_404(version_id)
    db.session.delete(v)
    db.session.commit()
    return jsonify({'success': True})


# ---------- Structured summary API ----------

@pricing_bp.route('/api/versions/<int:version_id>/summary')
def api_version_summary(version_id):
    version = PricingVersion.query.get_or_404(version_id)
    rules = PricingRule.query.filter_by(version_id=version_id).all()

    cat_map = {c.id: c.code for c in FeeCategory.query.all()}
    country_data = {}
    global_rules = []

    for r in rules:
        cat_code = cat_map.get(r.category_id, '?')
        region_name = r.region.name if r.region else None
        region_code = r.region.code if r.region else None
        entry = {
            'id': r.id,
            'category': cat_code,
            'cargo_type': r.cargo_type,
            'rule_type': r.rule_type,
            'params': r.get_params(),
        }
        if region_name:
            key = region_code or region_name
            if key not in country_data:
                country_data[key] = {'name': region_name, 'code': region_code, 'rules': []}
            country_data[key]['rules'].append(entry)
        else:
            global_rules.append(entry)

    countries_list = []
    for code, data in sorted(country_data.items(), key=lambda x: x[1]['name']):
        countries_list.append({
            'name': data['name'],
            'code': data['code'],
            'freight': _build_freight_table(data['rules']),
            'cod': _extract_cod(data['rules']),
            'return_fee': _extract_return(data['rules']),
        })

    postcodes_count = RemotePostcode.query.filter_by(version_id=version_id).count()

    return jsonify({
        'version': version.to_dict(),
        'countries': countries_list,
        'global': {
            'shelf_fee': _extract_global_rule(global_rules, 'SHELF_FEE'),
            'f_surcharge': _extract_f_surcharge(global_rules),
            'remote_fee': _extract_global_rule(global_rules, 'REMOTE_FEE'),
            'postcodes_count': postcodes_count,
        },
    })


def _build_freight_table(rules):
    CARGO_LABELS = {'GS': '普货', 'SC': '特货', 'IC': '敏感货'}
    table = {}
    for cargo_code, label in CARGO_LABELS.items():
        table[cargo_code] = {
            'label': label, 'head': None, 'tail_first': None,
            'tail_extra': None, 'carrier': None,
            'head_rule_id': None, 'tail_rule_id': None,
        }
    for r in rules:
        ct = r.get('cargo_type')
        if ct not in CARGO_LABELS:
            continue
        p = r.get('params', {})
        if r['category'] == 'HEAD_FREIGHT':
            table[ct]['head'] = p.get('rate_per_kg')
            table[ct]['head_rule_id'] = r['id']
            if p.get('carrier'):
                table[ct]['carrier'] = p['carrier']
        elif r['category'] == 'TAIL_FREIGHT':
            table[ct]['tail_first'] = p.get('first_price')
            table[ct]['tail_extra'] = p.get('extra_per_kg')
            table[ct]['tail_rule_id'] = r['id']
            if p.get('carrier'):
                table[ct]['carrier'] = p['carrier']
    return table


def _extract_cod(rules):
    for r in rules:
        if r['category'] == 'COD_FEE':
            p = r.get('params', {})
            return {'rate': p.get('rate'), 'min_amount': p.get('min_amount'), 'rule_id': r['id']}
    return None


def _extract_return(rules):
    for r in rules:
        if r['category'] == 'RETURN_FEE':
            p = r.get('params', {})
            return {
                'ratio': p.get('return_ratio', 1.0),
                'first_price': p.get('first_price'),
                'extra_per_kg': p.get('extra_per_kg'),
                'rule_id': r['id'],
            }
    return None


def _extract_global_rule(rules, cat_code):
    for r in rules:
        if r['category'] == cat_code:
            p = r.get('params', {})
            p['_rule_id'] = r['id']
            return p
    return None


def _extract_f_surcharge(rules):
    result = []
    for r in rules:
        if r['category'] == 'F_SURCHARGE':
            p = r.get('params', {})
            result.append({
                'cargo_type': r.get('cargo_type', ''),
                'amount': p.get('amount'),
                'currency': p.get('currency', 'EUR'),
                'rule_id': r['id'],
            })
    return result


# ---------- Rule CRUD ----------

@pricing_bp.route('/api/rules/<int:version_id>', methods=['GET'])
def api_rules(version_id):
    rules = PricingRule.query.filter_by(version_id=version_id).all()
    return jsonify([r.to_dict() for r in rules])


@pricing_bp.route('/api/rules', methods=['POST'])
def api_create_rule():
    data = request.get_json() or {}
    version_id = data.get('version_id')
    category_id = data.get('category_id')
    if not version_id or not category_id:
        return jsonify({'error': '版本和科目不能为空'}), 400
    rule = PricingRule(
        version_id=version_id, category_id=category_id,
        region_id=data.get('region_id') or None,
        cargo_type=data.get('cargo_type') or None,
        rule_type=data.get('rule_type', 'fixed'),
    )
    rule.set_params(data.get('params', {}))
    db.session.add(rule)
    db.session.commit()
    return jsonify(rule.to_dict()), 201


@pricing_bp.route('/api/rules/<int:rule_id>', methods=['PUT'])
def api_update_rule(rule_id):
    rule = PricingRule.query.get_or_404(rule_id)
    data = request.get_json() or {}
    for field in ['category_id', 'region_id', 'cargo_type', 'rule_type']:
        if field in data:
            setattr(rule, field, data[field])
    if 'params' in data:
        rule.set_params(data['params'])
    db.session.commit()
    return jsonify(rule.to_dict())


@pricing_bp.route('/api/rules/<int:rule_id>', methods=['DELETE'])
def api_delete_rule(rule_id):
    rule = PricingRule.query.get_or_404(rule_id)
    db.session.delete(rule)
    db.session.commit()
    return jsonify({'success': True})


# ---------- Override CRUD ----------

@pricing_bp.route('/api/overrides', methods=['GET'])
def api_overrides():
    customer_id = request.args.get('customer_id', type=int)
    query = CustomerPricingOverride.query
    if customer_id:
        query = query.filter_by(customer_id=customer_id)
    return jsonify([o.to_dict() for o in query.all()])


@pricing_bp.route('/api/overrides', methods=['POST'])
def api_create_override():
    data = request.get_json() or {}
    customer_id = data.get('customer_id')
    category_id = data.get('category_id')
    if not customer_id or not category_id:
        return jsonify({'error': '客户和科目不能为空'}), 400
    o = CustomerPricingOverride(
        customer_id=customer_id, category_id=category_id,
        region_id=data.get('region_id') or None,
        cargo_type=data.get('cargo_type') or None,
        rule_type=data.get('rule_type', 'fixed'),
        notes=data.get('notes', '')
    )
    o.set_params(data.get('params', {}))
    if data.get('effective_date'):
        from datetime import date
        try:
            parts = data['effective_date'].split('-')
            o.effective_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            pass
    db.session.add(o)
    db.session.commit()
    return jsonify(o.to_dict()), 201


# ---------- Remote Postcode APIs ----------

@pricing_bp.route('/api/postcodes', methods=['GET'])
def api_postcodes():
    version_id = request.args.get('version_id', type=int)
    search = request.args.get('q', '').strip()
    country = request.args.get('country', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 50, type=int)

    query = RemotePostcode.query
    if version_id:
        query = query.filter_by(version_id=version_id)
    if search:
        query = query.filter(RemotePostcode.postcode.like(f'{search}%'))
    if country:
        query = query.filter(RemotePostcode.country == country)

    total = query.count()
    items = query.order_by(RemotePostcode.postcode).offset((page - 1) * per_page).limit(per_page).all()

    return jsonify({
        'items': [p.to_dict() for p in items],
        'total': total,
        'page': page,
        'pages': (total + per_page - 1) // per_page,
    })


@pricing_bp.route('/api/postcodes/<int:pc_id>', methods=['PUT'])
def api_update_postcode(pc_id):
    pc = RemotePostcode.query.get_or_404(pc_id)
    data = request.get_json() or {}
    for field in ['postcode', 'country', 'zone', 'surcharge_type']:
        if field in data:
            setattr(pc, field, data[field])
    if 'surcharge_amount' in data:
        try:
            pc.surcharge_amount = float(data['surcharge_amount'])
        except (ValueError, TypeError):
            return jsonify({'error': '金额格式错误'}), 400
    db.session.commit()
    return jsonify(pc.to_dict())


@pricing_bp.route('/api/postcodes/<int:pc_id>', methods=['DELETE'])
def api_delete_postcode(pc_id):
    pc = RemotePostcode.query.get_or_404(pc_id)
    db.session.delete(pc)
    db.session.commit()
    return jsonify({'success': True})


@pricing_bp.route('/api/postcodes', methods=['POST'])
def api_create_postcode():
    data = request.get_json() or {}
    version_id = data.get('version_id')
    postcode = data.get('postcode', '').strip()
    if not version_id or not postcode:
        return jsonify({'error': '版本和邮编不能为空'}), 400
    try:
        amt = float(data.get('surcharge_amount', 0))
    except (ValueError, TypeError):
        return jsonify({'error': '金额格式错误'}), 400
    pc = RemotePostcode(
        version_id=version_id, postcode=postcode,
        country=data.get('country', ''),
        zone=data.get('zone', ''),
        surcharge_type=data.get('surcharge_type', 'per_kg'),
        surcharge_amount=amt,
    )
    db.session.add(pc)
    db.session.commit()
    return jsonify(pc.to_dict()), 201


@pricing_bp.route('/api/postcodes/countries', methods=['GET'])
def api_postcode_countries():
    version_id = request.args.get('version_id', type=int)
    query = db.session.query(RemotePostcode.country, db.func.count(RemotePostcode.id))
    if version_id:
        query = query.filter(RemotePostcode.version_id == version_id)
    results = query.group_by(RemotePostcode.country).all()
    return jsonify([{'country': c, 'count': n} for c, n in results if c])
