from flask import render_template, request, jsonify
from app.routes import exchange_rates_bp
from app import db
from app.models import ExchangeRate


@exchange_rates_bp.route('/')
def rates_page():
    rates = ExchangeRate.query.order_by(ExchangeRate.date.desc()).limit(50).all()
    return render_template('exchange_rates.html', rates=rates)


@exchange_rates_bp.route('/api', methods=['GET'])
def api_list():
    rates = ExchangeRate.query.order_by(ExchangeRate.date.desc()).limit(100).all()
    return jsonify([r.to_dict() for r in rates])


@exchange_rates_bp.route('/api', methods=['POST'])
def api_create():
    data = request.get_json() or {}
    from_c = data.get('from_currency', 'EUR')
    to_c = data.get('to_currency', 'CNY')
    rate_val = data.get('rate')
    date_str = data.get('date')

    if not rate_val or not date_str:
        return jsonify({'error': '汇率和日期不能为空'}), 400

    from datetime import date
    try:
        parts = date_str.split('-')
        d = date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        return jsonify({'error': '日期格式错误'}), 400

    try:
        rate_float = float(rate_val)
    except (ValueError, TypeError):
        return jsonify({'error': '汇率格式错误'}), 400

    r = ExchangeRate(
        from_currency=from_c,
        to_currency=to_c,
        rate=rate_float,
        date=d,
        source=data.get('source', 'manual')
    )
    db.session.add(r)
    db.session.commit()
    return jsonify(r.to_dict()), 201


@exchange_rates_bp.route('/api/<int:rid>', methods=['DELETE'])
def api_delete(rid):
    r = ExchangeRate.query.get_or_404(rid)
    db.session.delete(r)
    db.session.commit()
    return jsonify({'success': True})


@exchange_rates_bp.route('/api/latest', methods=['GET'])
def api_latest():
    from_c = request.args.get('from', 'EUR')
    to_c = request.args.get('to', 'CNY')
    rate = ExchangeRate.query.filter_by(
        from_currency=from_c, to_currency=to_c
    ).order_by(ExchangeRate.date.desc()).first()

    if rate:
        return jsonify(rate.to_dict())
    return jsonify({'error': '未找到汇率数据'}), 404
