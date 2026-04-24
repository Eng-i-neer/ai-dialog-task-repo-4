"""
Excel解析服务 - 解析代理账单Excel并写入数据库
以Sheet为单位处理，列名字典匹配（非位置匹配），兼容不同Excel的列顺序差异

核心逻辑：代理账单中的Sheet标识订单需要收取哪些费用（标记），
代理的价格(原币金额)仅做参考记录，我方收费由报价规则独立计算。

每个Sheet处理时会为该Sheet对应的科目创建OrderFee记录，记录：
- import_log_id: 本次导入的ID
- import_period: 账期字符串（如 "20250728"）
- source_sheet: 来源Sheet名称
"""
import os
import re
from datetime import datetime, date
from app import db
from app.models import Order, OrderFee, FeeCategory, Region, Customer, ImportLog
from app.services.excel_utils import load_excel

COUNTRY_MAP = {
    '德国': '德国', 'DE': '德国', 'Germany': '德国',
    '意大利': '意大利', 'IT': '意大利', 'Italy': '意大利',
    '西班牙': '西班牙', 'ES': '西班牙', 'Spain': '西班牙',
    '葡萄牙': '葡萄牙', 'PT': '葡萄牙', 'Portugal': '葡萄牙',
    '克罗地亚': '克罗地亚', 'HR': '克罗地亚', 'Croatia': '克罗地亚',
    '希腊': '希腊', 'GR': '希腊', 'Greece': '希腊',
    '斯洛文尼亚': '斯洛文尼亚', 'SI': '斯洛文尼亚', 'Slovenia': '斯洛文尼亚',
    '匈牙利': '匈牙利', 'HU': '匈牙利', 'Hungary': '匈牙利',
    '捷克': '捷克', 'CZ': '捷克', 'Czech': '捷克',
    '斯洛伐克': '斯洛伐克', 'SK': '斯洛伐克', 'Slovakia': '斯洛伐克',
    '罗马尼亚': '罗马尼亚', 'RO': '罗马尼亚', 'Romania': '罗马尼亚',
    '保加利亚': '保加利亚', 'BG': '保加利亚', 'Bulgaria': '保加利亚',
    '奥地利': '奥地利', 'AT': '奥地利', 'Austria': '奥地利',
    '波兰': '波兰', 'PL': '波兰', 'Poland': '波兰',
}

SHEET_TYPE_MAP = [
    ('头程运费', 'head_freight'),
    ('头程', 'head_freight'),
    ('尾程运费', 'tail_freight'),
    ('尾程派送', 'tail_freight'),
    ('地派服务费', 'tail_freight'),
    ('尾程退件操作费', 'return_fee'),
    ('拒收返程费', 'return_fee'),
    ('拒收返程', 'return_fee'),
    ('补退', 'return_refund'),
    ('上架费', 'shelf_fee'),
    ('上架', 'shelf_fee'),
    ('目的地增值税', 'vat'),
    ('增值税', 'vat'),
    ('VAT', 'vat'),
    ('代收COD手续费', 'cod_fee_sheet'),
    ('代收COD', 'cod'),
    ('COD', 'cod'),
    ('二派费', 'second_delivery'),
    ('二派', 'second_delivery'),
    ('旺季附加费', 'peak_surcharge'),
    ('偏远费', 'remote_fee'),
    ('偏远', 'remote_fee'),
    ('F附加', 'f_surcharge'),
    ('F-附加', 'f_surcharge'),
    ('转寄操作费', 'other'),
    ('转寄操作', 'other'),
    ('短信费', 'other'),
    ('客诉电话费', 'other'),
    ('服务费', 'other'),
    ('账号管理', 'other'),
    ('海外仓', 'other'),
]

SHEET_TYPE_TO_FEE_CODE = {
    'head_freight': 'HEAD_FREIGHT',
    'tail_freight': 'TAIL_FREIGHT',
    'cod': 'COD_FEE',
    'cod_fee_sheet': 'COD_FEE',
    'return_fee': 'RETURN_FEE',
    'return_refund': 'RETURN_REFUND',
    'shelf_fee': 'SHELF_FEE',
    'vat': 'VAT',
    'remote_fee': 'REMOTE_FEE',
    'f_surcharge': 'F_SURCHARGE',
    'second_delivery': 'SECOND_DELIVERY',
    'peak_surcharge': 'MISC_FEE',
    'other': None,
}

CUSTOMER_TEMPLATE_PATTERN = re.compile(r'^\d{8}期')

COLUMN_ALIASES = {
    'waybill': ['运单号码', '运单号', '快递单号'],
    'customer_ref': ['客户单号'],
    'transfer': ['转单号'],
    'country': ['目的地', '目的国'],
    'weight': ['实重', '实际重量'],
    'charge_head': ['头程计费重'],
    'charge_tail': ['尾程计费重', '计费重量'],
    'charge_weight': ['重量(KG)', '重量（KG）', '重量'],
    'product': ['品名', '品类', '中文品名'],
    'pieces': ['件数'],
    'amount': ['原币金额', '金额', '费用', '运费', '收费'],
    'ship_date': ['寄件日期', '发货日期', '揽收日期', '揽件日期'],
    'notes': ['备注'],
    'dimensions': ['尺寸', '体积'],
    'postcode': ['邮编'],
    'cargo_type_col': ['类型'],
    'route': ['指定路线', '路线'],
    'subject': ['科目'],
    'formula': ['计算公式'],
    'sender': ['寄件人'],
    'sender_company': ['寄件公司'],
    'receiver': ['收件人'],
    'quantity': ['数量'],
}

HEADER_REQUIRED_FIELDS = {'运单号码', '寄件日期', '目的地', '原币金额'}


def _normalize_country(raw):
    if not raw:
        return None
    raw = str(raw).strip()
    return COUNTRY_MAP.get(raw, raw)


_region_cache = {}


def _get_or_create_region(country_name):
    if not country_name:
        return None
    if country_name in _region_cache:
        return _region_cache[country_name]
    region = Region.query.filter_by(name=country_name).first()
    if not region:
        region = Region(name=country_name)
        db.session.add(region)
        db.session.flush()
    _region_cache[country_name] = region
    return region


def _safe_float(val):
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _detect_ship_type(product_name, notes=''):
    text = f"{product_name or ''} {notes or ''}".lower()
    if '转寄' in text:
        return '转寄'
    return '直发'


def _detect_cargo_type(product_name, notes='', route=''):
    text = f"{product_name or ''} {notes or ''}"
    if any(k in text for k in ['敏感', '纯电', 'IC']):
        return 'IC'
    if any(k in text for k in ['特货', 'F货', 'F-', 'F手表', 'F-手表', 'SC']):
        return 'SC'
    route_str = route or ''
    if '东欧专线' in route_str:
        return 'IC'
    return 'GS'


def _parse_date_from_filename(filename):
    m = re.search(r'(\d{8})', filename)
    if m:
        ds = m.group(1)
        try:
            return date(int(ds[:4]), int(ds[4:6]), int(ds[6:8]))
        except ValueError:
            pass
    return None


def _match_sheet_type(sheet_name):
    """Match sheet name to a logical sheet type for fee marking.
    Rejects customer template sheets (e.g. '20260330期运费')."""
    if CUSTOMER_TEMPLATE_PATTERN.match(sheet_name):
        return None
    if sheet_name in ('汇总', '理赔', '超期仓储'):
        return None
    for keyword, stype in SHEET_TYPE_MAP:
        if keyword in sheet_name:
            return stype
    return None


def _build_col_map(ws, header_row):
    """Build column index mapping by matching header cell text against COLUMN_ALIASES.
    Pure name-based matching — never relies on column position."""
    col_map = {}
    for col_idx in range(1, min((ws.max_column or 0) + 1, 30)):
        raw = ws.cell(header_row, col_idx).value
        if raw is None:
            continue
        val = str(raw).strip()
        if not val:
            continue

        if '客户' in val and '单号' in val:
            col_map['customer_ref'] = col_idx
            continue

        if val in ('运单号码', '运单号', '快递单号'):
            col_map['waybill'] = col_idx
            continue

        for field, aliases in COLUMN_ALIASES.items():
            if field in col_map:
                continue
            if field in ('waybill', 'customer_ref'):
                continue
            for alias in aliases:
                if val == alias:
                    col_map[field] = col_idx
                    break
            else:
                continue
            break

    return col_map


def _find_header_row(ws):
    """Find the real header row by requiring >=3 of the known header field names.
    Skips metadata rows like '帐单号码:' which only contain one match."""
    for row_idx in range(1, min(15, (ws.max_row or 0) + 1)):
        row_texts = set()
        for col_idx in range(1, min((ws.max_column or 0) + 1, 25)):
            v = ws.cell(row_idx, col_idx).value
            if v is not None:
                row_texts.add(str(v).strip())

        matches = sum(1 for kw in HEADER_REQUIRED_FIELDS if kw in row_texts)
        if matches >= 3:
            return row_idx
    return None


_fee_category_cache = {}


def _get_fee_category(code):
    """Cached FeeCategory lookup to avoid repeated DB queries."""
    if code not in _fee_category_cache:
        cat = FeeCategory.query.filter_by(code=code).first()
        _fee_category_cache[code] = cat
    return _fee_category_cache.get(code)


def _get_or_create_fee(order, fee_code, import_log_id, period_str, sheet_name):
    """Get existing or create new OrderFee for the given order+category.
    If fee already exists (from a previous import), update the source info."""
    cat = _get_fee_category(fee_code)
    if not cat:
        return None

    if order.id:
        fee = OrderFee.query.filter_by(order_id=order.id, category_id=cat.id).first()
    else:
        fee = None

    if not fee:
        fee = OrderFee(
            order_id=order.id,
            category_id=cat.id,
            import_log_id=import_log_id,
            import_period=period_str,
            source_sheet=sheet_name,
            input_currency='EUR',
        )
        db.session.add(fee)
    else:
        if not fee.import_log_id:
            fee.import_log_id = import_log_id
        if not fee.import_period:
            fee.import_period = period_str
        if not fee.source_sheet:
            fee.source_sheet = sheet_name
    return fee


def parse_agent_bill(filepath, import_log_id=None, customer_id=None):
    """
    Parse agent bill Excel. Each Sheet identifies orders and marks
    which of our fees apply. Creates OrderFee records with source tracking.
    """
    global _fee_category_cache, _region_cache
    _fee_category_cache = {}
    _region_cache = {}

    wb = load_excel(filepath, data_only=True)
    filename = os.path.basename(filepath)
    bill_period = _parse_date_from_filename(filename)
    period_str = bill_period.strftime('%Y%m%d') if bill_period else None

    orders_map = {}
    agent_fees_recorded = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        sheet_type = _match_sheet_type(sheet_name)
        if not sheet_type:
            continue

        header_row = _find_header_row(ws)
        if not header_row:
            continue

        col_map = _build_col_map(ws, header_row)
        if 'waybill' not in col_map:
            continue

        is_head_sheet = sheet_type == 'head_freight'
        is_tail_sheet = sheet_type == 'tail_freight'
        fee_code = SHEET_TYPE_TO_FEE_CODE.get(sheet_type)

        for row_idx in range(header_row + 1, (ws.max_row or 0) + 1):
            first_cell = ws.cell(row_idx, 1).value
            if first_cell and str(first_cell).strip().startswith('合计'):
                break

            waybill = ws.cell(row_idx, col_map['waybill']).value
            if not waybill:
                continue
            waybill = str(waybill).strip()
            if not waybill or waybill == 'None':
                continue
            if waybill.startswith('合计') or waybill.startswith('小计'):
                break

            row_charge_weight = None
            for cw_field in ('charge_head', 'charge_tail', 'charge_weight'):
                if cw_field in col_map:
                    row_charge_weight = _safe_float(ws.cell(row_idx, col_map[cw_field]).value)
                    if row_charge_weight is not None:
                        break

            if waybill not in orders_map:
                existing = Order.query.filter_by(waybill_no=waybill).first()
                if existing:
                    orders_map[waybill] = existing

            if waybill not in orders_map:
                country_raw = ws.cell(row_idx, col_map.get('country', 0)).value if 'country' in col_map else None
                country = _normalize_country(country_raw)
                region = _get_or_create_region(country) if country else None

                product_name = str(ws.cell(row_idx, col_map.get('product', 0)).value or '') if 'product' in col_map else ''
                notes = str(ws.cell(row_idx, col_map.get('notes', 0)).value or '') if 'notes' in col_map else ''

                transfer_no = None
                if 'transfer' in col_map:
                    tv = ws.cell(row_idx, col_map['transfer']).value
                    if tv:
                        transfer_no = str(tv).strip() or None

                customer_ref = None
                if 'customer_ref' in col_map:
                    cr = ws.cell(row_idx, col_map['customer_ref']).value
                    if cr:
                        customer_ref = str(cr).strip() or None

                route = str(ws.cell(row_idx, col_map.get('route', 0)).value or '') if 'route' in col_map else ''
                cargo_type = _detect_cargo_type(product_name, notes, route)
                if 'cargo_type_col' in col_map:
                    ct_val = str(ws.cell(row_idx, col_map['cargo_type_col']).value or '').strip()
                    if ct_val in ('GS', 'SC', 'IC'):
                        cargo_type = ct_val
                    elif ct_val in ('普货',):
                        cargo_type = 'GS'
                    elif ct_val in ('特货', 'F货'):
                        cargo_type = 'SC'
                    elif ct_val in ('敏感货', '纯电'):
                        cargo_type = 'IC'

                order = Order(
                    waybill_no=waybill,
                    transfer_no=transfer_no,
                    customer_id=customer_id,
                    region_id=region.id if region else None,
                    import_log_id=import_log_id,
                    bill_period=bill_period,
                    ship_type=_detect_ship_type(product_name, notes),
                    product_name=product_name or None,
                    cargo_type=cargo_type,
                    pieces=int(ws.cell(row_idx, col_map.get('pieces', 0)).value or 1) if 'pieces' in col_map else 1,
                    actual_weight=_safe_float(ws.cell(row_idx, col_map.get('weight', 0)).value) if 'weight' in col_map else None,
                    charge_weight_head=row_charge_weight if is_head_sheet else None,
                    charge_weight_tail=row_charge_weight if is_tail_sheet else None,
                    dimensions=str(ws.cell(row_idx, col_map.get('dimensions', 0)).value or '') if 'dimensions' in col_map else None,
                    customer_ref=customer_ref,
                    postcode=str(ws.cell(row_idx, col_map.get('postcode', 0)).value or '') if 'postcode' in col_map else None,
                    logistics_status='待处理',
                    source_file=filename,
                )

                if 'ship_date' in col_map:
                    sd = ws.cell(row_idx, col_map['ship_date']).value
                    if isinstance(sd, datetime):
                        order.ship_date = sd.date()
                    elif isinstance(sd, date):
                        order.ship_date = sd

                if period_str:
                    order.add_import_period(period_str)

                db.session.add(order)
                db.session.flush()

                orders_map[waybill] = order
            else:
                order = orders_map[waybill]

                if period_str:
                    order.add_import_period(period_str)

                if row_charge_weight is not None:
                    if is_head_sheet and not order.charge_weight_head:
                        order.charge_weight_head = row_charge_weight
                    elif is_tail_sheet and not order.charge_weight_tail:
                        order.charge_weight_tail = row_charge_weight

                if not order.actual_weight and 'weight' in col_map:
                    w = _safe_float(ws.cell(row_idx, col_map['weight']).value)
                    if w:
                        order.actual_weight = w

                if not order.transfer_no and 'transfer' in col_map:
                    tv = ws.cell(row_idx, col_map['transfer']).value
                    if tv:
                        order.transfer_no = str(tv).strip() or None

                if not order.customer_ref and 'customer_ref' in col_map:
                    cr = ws.cell(row_idx, col_map['customer_ref']).value
                    if cr:
                        order.customer_ref = str(cr).strip() or None

                if not order.postcode and 'postcode' in col_map:
                    pc = ws.cell(row_idx, col_map['postcode']).value
                    if pc:
                        order.postcode = str(pc).strip() or None

                if not order.region_id and 'country' in col_map:
                    country_raw = ws.cell(row_idx, col_map['country']).value
                    country = _normalize_country(country_raw)
                    if country:
                        region = _get_or_create_region(country)
                        if region:
                            order.region_id = region.id

            order.add_import_sheet(sheet_name)

            if is_head_sheet:
                order.has_head_freight = True
            if is_tail_sheet:
                order.has_tail_freight = True

            if sheet_type == 'cod':
                amt = _safe_float(ws.cell(row_idx, col_map.get('amount', 0)).value) if 'amount' in col_map else None
                if amt is not None:
                    order.cod_amount = amt
                    order.cod_currency = 'EUR'

            elif sheet_type == 'return_fee':
                order.needs_return_fee = True

            elif sheet_type == 'shelf_fee':
                order.needs_shelf_fee = True

            elif sheet_type == 'second_delivery':
                order.needs_second_delivery = True

            elif sheet_type == 'vat':
                order.needs_vat = True

            if fee_code:
                fee = _get_or_create_fee(order, fee_code, import_log_id, period_str, sheet_name)
                if fee:
                    agent_amt = _safe_float(ws.cell(row_idx, col_map.get('amount', 0)).value) if 'amount' in col_map else None
                    if agent_amt is not None:
                        fee.input_amount = agent_amt

            agent_fees_recorded += 1

        db.session.flush()

    wb.close()

    db.session.commit()

    new_count = sum(1 for o in orders_map.values() if len(o.import_period_list) <= 1)
    updated_count = len(orders_map) - new_count

    return {
        'orders_count': len(orders_map),
        'new_orders': new_count,
        'updated_orders': updated_count,
        'sheets_processed': agent_fees_recorded,
    }
