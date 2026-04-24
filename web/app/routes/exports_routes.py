import os
from flask import render_template, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename
from app.routes import exports_bp
from app import db
from app.models import Customer, FeeCategory, Order, OrderFee


@exports_bp.route('/')
def export_page():
    customers = Customer.query.order_by(Customer.name).all()
    categories = FeeCategory.query.order_by(FeeCategory.code).all()
    return render_template('exports.html',
        customers=customers, categories=categories)


@exports_bp.route('/api/preview', methods=['POST'])
def api_preview():
    data = request.get_json() or {}
    customer_id = data.get('customer_id')
    category_ids = data.get('category_ids', [])
    bill_period = data.get('bill_period')

    query = Order.query
    if customer_id:
        query = query.filter_by(customer_id=customer_id)
    if bill_period:
        from datetime import date
        try:
            parts = bill_period.split('-')
            bp = date(int(parts[0]), int(parts[1]), int(parts[2]))
            query = query.filter_by(bill_period=bp)
        except (ValueError, IndexError):
            pass

    if category_ids:
        order_ids_with_fees = db.session.query(OrderFee.order_id).filter(
            OrderFee.category_id.in_(category_ids)
        ).distinct()
        query = query.filter(Order.id.in_(order_ids_with_fees))

    orders = query.all()
    return jsonify({
        'orders_count': len(orders),
        'orders': [o.to_dict() for o in orders[:20]]
    })


@exports_bp.route('/api/generate', methods=['POST'])
def api_generate():
    data = request.get_json() or {}
    customer_id = data.get('customer_id')
    bill_period = data.get('bill_period')
    category_ids = data.get('category_ids', [])

    if not customer_id:
        return jsonify({'error': '请选择客户'}), 400

    try:
        from app.services.export_engine import generate_export
        filepath = generate_export(customer_id, bill_period, category_ids)
        basename = os.path.basename(filepath)
        return jsonify({
            'success': True,
            'filename': basename,
            'download_url': f'/exports/api/download/{basename}'
        })
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': '导出失败，请重试'}), 500


@exports_bp.route('/api/download/<filename>')
def api_download(filename):
    filename = secure_filename(filename) or os.path.basename(filename)
    export_dir = os.path.join(current_app.config['UPLOAD_FOLDER'], 'exports')
    filepath = os.path.join(export_dir, filename)
    real_export = os.path.realpath(export_dir)
    real_file = os.path.realpath(filepath)
    if not real_file.startswith(real_export):
        return jsonify({'error': '非法路径'}), 403
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=filename)
