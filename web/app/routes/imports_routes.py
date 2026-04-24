import os
import re
import unicodedata
from flask import render_template, request, jsonify, current_app
from app.routes import imports_bp
from app import db
from app.models import ImportLog


def _safe_filename(filename):
    """secure_filename that preserves CJK characters."""
    filename = unicodedata.normalize('NFC', filename)
    filename = filename.replace(os.sep, '_').replace('/', '_')
    filename = re.sub(r'[<>:"|?*\\]', '_', filename)
    filename = re.sub(r'\.\.+', '.', filename)
    filename = filename.strip('. ')
    return filename or 'upload.xlsx'


@imports_bp.route('/')
def import_page():
    logs = ImportLog.query.order_by(ImportLog.created_at.desc()).limit(20).all()
    return render_template('imports.html', logs=logs)


@imports_bp.route('/api/upload', methods=['POST'])
def api_upload():
    if 'file' not in request.files:
        return jsonify({'error': '未选择文件'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': '文件名为空'}), 400

    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': '仅支持 .xlsx / .xls 文件'}), 400

    upload_dir = current_app.config['UPLOAD_FOLDER']
    os.makedirs(upload_dir, exist_ok=True)

    filename = _safe_filename(file.filename)

    filepath = os.path.join(upload_dir, filename)
    file.save(filepath)

    file_type = request.form.get('file_type', 'agent_bill')
    if file_type not in ('agent_bill', 'pricing_file', 'cargo_status'):
        file_type = 'agent_bill'

    original_name = file.filename or filename

    if file_type == 'pricing_file':
        return _handle_pricing_upload(filepath, filename)
    else:
        return jsonify({
            'success': True,
            'mode': 'upload_ready',
            'filepath': filepath,
            'filename': filename,
            'display_name': original_name,
            'file_type': file_type,
        })


@imports_bp.route('/api/confirm', methods=['POST'])
def api_confirm_import():
    """Step 2: user clicks confirm after upload — validate fields, then process."""
    data = request.get_json() or {}
    filepath = data.get('filepath')
    filename = data.get('filename')
    file_type = data.get('file_type', 'agent_bill')
    bill_period = data.get('bill_period', '')
    customer_id = data.get('customer_id', '')

    if not filepath or not os.path.exists(filepath):
        return jsonify({'error': '文件不存在，请重新上传'}), 400

    if not bill_period:
        if file_type == 'agent_bill':
            return jsonify({'error': '请选择账期日期'}), 400
        elif file_type == 'cargo_status':
            return jsonify({'error': '请选择账期日期'}), 400

    if file_type == 'agent_bill' and not customer_id:
        return jsonify({'error': '请选择客户'}), 400

    if file_type == 'cargo_status':
        return _handle_cargo_status_confirm(filepath, filename or os.path.basename(filepath), bill_period)
    else:
        return _handle_agent_confirm(filepath, filename or os.path.basename(filepath), bill_period, customer_id)


def _handle_agent_confirm(filepath, filename, bill_period, customer_id):
    """Process agent bill after user confirmation."""
    log = ImportLog(filename=filename, file_type='agent_bill', status='uploaded')

    if bill_period:
        from datetime import date
        try:
            parts = bill_period.split('-')
            log.bill_period = date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            pass

    db.session.add(log)
    db.session.commit()

    try:
        cid = None
        if customer_id:
            try:
                cid = int(customer_id)
            except (ValueError, TypeError):
                pass

        from app.services.excel_parser import parse_agent_bill
        result = parse_agent_bill(filepath, log.id, customer_id=cid)
        log.orders_count = result.get('orders_count', 0)
        log.status = 'success'
        db.session.commit()

        new_count = result.get('new_orders', 0)
        updated_count = result.get('updated_orders', 0)

        msg_parts = [f'共处理 {log.orders_count} 条订单']
        if new_count:
            msg_parts.append(f'新增 {new_count} 条')
        if updated_count:
            msg_parts.append(f'更新 {updated_count} 条已有订单')
        msg_parts.append('可在订单列表中批量计算费用')

        return jsonify({
            'success': True,
            'log_id': log.id,
            'orders_count': log.orders_count,
            'new_orders': new_count,
            'updated_orders': updated_count,
            'message': '，'.join(msg_parts)
        })
    except Exception as e:
        log.status = 'error'
        log.error_log = str(e)
        db.session.commit()
        return jsonify({'error': f'解析失败: {str(e)}'}), 500


def _handle_cargo_status_confirm(filepath, filename, bill_period):
    """Process cargo status file after user confirmation."""
    log = ImportLog(filename=filename, file_type='cargo_status', status='uploaded')

    bp = None
    if bill_period:
        from datetime import date
        try:
            parts = bill_period.split('-')
            bp = date(int(parts[0]), int(parts[1]), int(parts[2]))
            log.bill_period = bp
        except (ValueError, IndexError):
            pass

    db.session.add(log)
    db.session.commit()

    try:
        from app.services.cargo_status_parser import parse_cargo_status
        result = parse_cargo_status(filepath, bill_period=bp)
        log.orders_count = result.get('updated_count', 0)
        log.status = 'success'
        db.session.commit()

        msg_parts = [f'更新 {result["updated_count"]} 条订单']
        if result.get('remote_marked'):
            msg_parts.append(f'标记 {result["remote_marked"]} 条偏远')
        if result.get('not_found_count'):
            msg_parts.append(f'{result["not_found_count"]} 条未匹配')

        return jsonify({
            'success': True,
            'log_id': log.id,
            'message': '，'.join(msg_parts),
            'result': result,
        })
    except Exception as e:
        log.status = 'error'
        log.error_log = str(e)
        db.session.commit()
        return jsonify({'error': f'解析失败: {str(e)}'}), 500


def _handle_pricing_upload(filepath, filename):
    """Step 1: preview — parse and return summary without writing to DB."""
    try:
        from app.services.pricing_parser import preview_pricing_file
        preview = preview_pricing_file(filepath)

        total_rules = sum(preview.get('rules_by_category', {}).values())
        return jsonify({
            'success': True,
            'mode': 'pricing_preview',
            'filepath': filepath,
            'filename': filename,
            'preview': {
                'countries': preview.get('countries', []),
                'rules_by_category': preview.get('rules_by_category', {}),
                'total_rules': total_rules,
                'vat_updates': len(preview.get('vat_updates', [])),
                'remote_postcodes': preview.get('remote_postcodes_count', 0),
                'currencies': preview.get('currencies', []),
                'warnings': preview.get('warnings', []),
            }
        })
    except Exception as e:
        return jsonify({'error': f'报价文件解析失败: {str(e)}'}), 500


@imports_bp.route('/api/pricing/confirm', methods=['POST'])
def api_pricing_confirm():
    """Step 2: confirm — actually write pricing data to DB."""
    data = request.get_json() or {}
    filepath = data.get('filepath')
    version_name = data.get('version_name', '').strip()
    effective_date_str = data.get('effective_date', '')
    expire_date_str = data.get('expire_date', '')

    if not filepath or not os.path.exists(filepath):
        return jsonify({'error': '文件不存在，请重新上传'}), 400
    if not version_name:
        return jsonify({'error': '请输入版本名称'}), 400
    if not effective_date_str:
        return jsonify({'error': '请选择生效日期'}), 400

    from datetime import date
    try:
        parts = effective_date_str.split('-')
        eff_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        return jsonify({'error': '日期格式错误'}), 400

    exp_date = None
    if expire_date_str:
        try:
            parts = expire_date_str.split('-')
            exp_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            pass

    log = ImportLog(
        filename=os.path.basename(filepath),
        file_type='pricing_file',
        bill_period=eff_date,
        status='uploaded'
    )
    db.session.add(log)
    db.session.flush()

    try:
        from app.services.pricing_parser import commit_pricing_file
        summary = commit_pricing_file(
            filepath,
            version_name=version_name,
            effective_date=eff_date,
            expire_date=exp_date,
            source_filename=os.path.basename(filepath),
        )
        log.orders_count = summary.get('rules_created', 0)
        log.status = 'success'
        db.session.commit()

        msg_parts = []
        if summary.get('rules_created'):
            msg_parts.append(f'{summary["rules_created"]} 条计价规则')
        if summary.get('vat_updated'):
            msg_parts.append(f'{summary["vat_updated"]} 个国家VAT税率更新')
        if summary.get('postcodes_created'):
            msg_parts.append(f'{summary["postcodes_created"]} 个偏远邮编')
        message = '成功导入: ' + '，'.join(msg_parts) if msg_parts else '导入完成（无新数据）'

        return jsonify({
            'success': True,
            'message': message,
            'summary': summary,
        })
    except Exception as e:
        log.status = 'error'
        log.error_log = str(e)
        db.session.commit()
        return jsonify({'error': f'写入失败: {str(e)}'}), 500


@imports_bp.route('/api/logs', methods=['GET'])
def api_logs():
    logs = ImportLog.query.order_by(ImportLog.created_at.desc()).limit(50).all()
    return jsonify([l.to_dict() for l in logs])


@imports_bp.route('/api/logs/<int:log_id>', methods=['GET'])
def api_log_detail(log_id):
    log = ImportLog.query.get_or_404(log_id)
    detail = log.to_dict()

    if log.file_type == 'agent_bill' and log.filename:
        from app.models import Order
        count = Order.query.filter_by(source_file=log.filename).count()
        detail['linked_orders'] = count
    elif log.file_type == 'pricing_file' and log.filename:
        from app.models import PricingVersion
        versions = PricingVersion.query.filter_by(source_file=log.filename).all()
        detail['linked_versions'] = [{'id': v.id, 'name': v.name} for v in versions]

    return jsonify(detail)


@imports_bp.route('/api/logs/<int:log_id>', methods=['DELETE'])
def api_delete_log(log_id):
    log = ImportLog.query.get_or_404(log_id)
    deleted_info = {}

    if log.file_type == 'agent_bill' and log.filename:
        from app.models import Order, OrderFee
        orders = Order.query.filter_by(source_file=log.filename).all()
        order_count = len(orders)
        for order in orders:
            OrderFee.query.filter_by(order_id=order.id).delete()
            db.session.delete(order)
        deleted_info['orders_deleted'] = order_count

    elif log.file_type == 'pricing_file' and log.filename:
        from app.models import PricingVersion
        versions = PricingVersion.query.filter_by(source_file=log.filename).all()
        v_names = [v.name for v in versions]
        for v in versions:
            db.session.delete(v)
        deleted_info['versions_deleted'] = v_names

    if log.filename:
        upload_dir = current_app.config['UPLOAD_FOLDER']
        fpath = os.path.join(upload_dir, log.filename)
        real_upload = os.path.realpath(upload_dir)
        real_fpath = os.path.realpath(fpath)
        if real_fpath.startswith(real_upload) and os.path.exists(real_fpath):
            try:
                os.remove(real_fpath)
                deleted_info['file_removed'] = True
            except OSError:
                deleted_info['file_removed'] = False

    db.session.delete(log)
    db.session.commit()

    return jsonify({'success': True, **deleted_info})
