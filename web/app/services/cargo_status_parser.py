"""
货况表解析服务 - 从货况表中提取订单邮编信息，
用于匹配偏远邮编表来确定哪些订单需要收取偏远费。

正确的偏远费判定流程：
 1. 导入货况表 → 从"直发"和"转寄"表中读取运单号和邮编
 2. 用运单号匹配数据库中已有的订单
 3. 将邮编写入匹配到的订单
 4. 用订单的邮编去计价规则的偏远邮编表中查询是否命中
 5. 命中且有尾程运费则 is_remote=True

邮编来源：
 - 直发表：收件邮编 列
 - 转寄表：收件邮编 或 收件区县 列（部分转寄订单的邮编存在区县列）

处理策略：
 1. 优先处理"直发"表获取邮编和状态
 2. 继续处理"转寄"表，为仍缺少邮编的订单补充数据
"""
import os
import pandas as pd
from app import db
from app.models import Order, Region, RemotePostcode, PricingVersion

COLUMN_ALIASES = {
    'waybill': ['运单号', '运单号码', '快递单号', '面单单号'],
    'postcode': ['收件邮编', '邮编', '邮政编码', '收件区县'],
    'address': ['收件详细地址', '详细地址', '收件地址', '地址'],
    'cargo_status': ['货态', '物流状态'],
    'cargo_status_category': ['归类', '状态归类'],
}

PRIORITY_SHEETS = ['2026直发', '2025直发', '直发']
SKIP_SHEETS = {'国家', 'Sheet1'}


def _safe_str(val):
    if val is None:
        return None
    if isinstance(val, float):
        import math
        if math.isnan(val):
            return None
    s = str(val).strip()
    return s if s and s.lower() != 'nan' and s != 'None' else None


def _normalize_postcode(pc_str):
    """Normalize postcode: strip spaces, remove trailing '.0' from numeric strings."""
    if not pc_str:
        return None
    pc = pc_str.strip()
    if pc.endswith('.0'):
        pc = pc[:-2]
    return pc if pc else None


def _resolve_columns(df_columns):
    """Match DataFrame columns to known field aliases.
    Aliases are ordered by priority (first alias = highest priority).
    For 'waybill', '运单号' is preferred over '面单单号' because
    运单号 contains the system tracking number (e.g. DE12602061410020)
    while 面单单号 is the courier label number (e.g. JJD149990200077398859).
    Two passes: exact match by alias priority, then substring fallback."""
    col_map = {}
    col_list = [str(c).strip() for c in df_columns]
    col_orig = list(df_columns)
    used_cols = set()

    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            for i, col_str in enumerate(col_list):
                if col_orig[i] in used_cols:
                    continue
                if col_str == alias:
                    col_map[field] = col_orig[i]
                    used_cols.add(col_orig[i])
                    break
            if field in col_map:
                break

    for field, aliases in COLUMN_ALIASES.items():
        if field in col_map:
            continue
        for alias in aliases:
            for i, col_str in enumerate(col_list):
                if col_orig[i] in used_cols:
                    continue
                if alias in col_str or col_str in alias:
                    col_map[field] = col_orig[i]
                    used_cols.add(col_orig[i])
                    break
            if field in col_map:
                break
    return col_map


def _process_df(df, col_map, order_map, remote_postcode_keys,
                region_names, seen_waybills, stats):
    wb_col = col_map.get('waybill')
    pc_col = col_map.get('postcode')
    addr_col = col_map.get('address')
    status_col = col_map.get('cargo_status')
    status_cat_col = col_map.get('cargo_status_category')
    if wb_col is None:
        return

    waybills = df[wb_col].values
    postcodes = df[pc_col].values if pc_col else None
    addresses = df[addr_col].values if addr_col else None
    statuses = df[status_col].values if status_col else None
    status_cats = df[status_cat_col].values if status_cat_col else None
    for i in range(len(waybills)):
        waybill = _safe_str(waybills[i])
        if not waybill:
            continue

        orders = order_map.get(waybill)
        if not orders:
            if waybill not in seen_waybills:
                stats['not_found'] += 1
            seen_waybills.add(waybill)
            continue
        seen_waybills.add(waybill)

        pc_raw = _safe_str(postcodes[i]) if postcodes is not None else None
        pc = _normalize_postcode(pc_raw)
        addr = _safe_str(addresses[i]) if addresses is not None else None
        status = _safe_str(statuses[i]) if statuses is not None else None
        status_cat = _safe_str(status_cats[i]) if status_cats is not None else None

        for order in orders:
            changed = False
            if pc:
                order.postcode = pc
                changed = True
                country = region_names.get(order.region_id, '')
                pc_is_remote = (country, pc) in remote_postcode_keys
                should_mark = pc_is_remote and order.has_tail_freight
                if should_mark and not order.is_remote:
                    order.is_remote = True
                    stats['remote_marked'] += 1
                elif not should_mark and order.is_remote:
                    order.is_remote = False
            if addr and not order.address:
                order.address = addr
                changed = True
            if status:
                order.cargo_status = status
                changed = True
            if status_cat:
                order.cargo_status_category = status_cat
                changed = True
            if changed:
                stats['updated'] += 1


_USECOLS_KEYWORDS = ['运单', '面单', '单号', '邮编', '区县', '地址', '货态', '归类']


def _usecols_filter(col_name):
    s = str(col_name)
    return any(kw in s for kw in _USECOLS_KEYWORDS)


def _load_sheets(filepath):
    """Load all sheets at once, only columns matching known keywords."""
    ext = os.path.splitext(filepath)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    return pd.read_excel(filepath, sheet_name=None, header=0,
                         dtype=str, engine=engine,
                         usecols=_usecols_filter)


def _build_remote_postcode_map(version_id):
    """Build a dict of (country, postcode) -> True for O(1) lookup.
    Different countries can share the same postcode numbers,
    so matching must be scoped by country."""
    if not version_id:
        return set()
    keys = set()
    for rp in RemotePostcode.query.filter_by(version_id=version_id).all():
        pc = _normalize_postcode(rp.postcode)
        if pc and rp.country:
            keys.add((rp.country, pc))
    return keys


def _build_region_name_map():
    """Build region_id -> region_name mapping for country lookup."""
    return {r.id: r.name for r in Region.query.all()}


def _get_version_for_period(bill_period):
    """Find the correct pricing version for a given bill period."""
    if bill_period:
        version = PricingVersion.query.filter(
            PricingVersion.effective_date <= bill_period,
            db.or_(
                PricingVersion.expire_date.is_(None),
                PricingVersion.expire_date > bill_period
            )
        ).order_by(PricingVersion.effective_date.desc()).first()
        if version:
            return version

    active = PricingVersion.query.filter_by(is_active=True).first()
    if active:
        return active
    return PricingVersion.query.order_by(
        PricingVersion.effective_date.desc()
    ).first()


def parse_cargo_status(filepath, bill_period=None):
    all_dfs = _load_sheets(filepath)

    order_query = Order.query
    if bill_period:
        order_query = order_query.filter_by(bill_period=bill_period)
    all_orders = order_query.all()
    order_map = {}
    for o in all_orders:
        order_map.setdefault(o.waybill_no, []).append(o)

    total_order_waybills = set(order_map.keys())

    version = _get_version_for_period(bill_period)
    remote_postcode_keys = _build_remote_postcode_map(
        version.id if version else None
    )
    region_names = _build_region_name_map()

    stats = {'updated': 0, 'remote_marked': 0, 'not_found': 0}
    sheets_processed = 0
    seen_waybills = set()

    sheet_names = list(all_dfs.keys())

    priority = []
    secondary = []
    for ps in PRIORITY_SHEETS:
        for sn in sheet_names:
            if ps in sn and sn not in SKIP_SHEETS and sn not in priority:
                priority.append(sn)
    for sn in sheet_names:
        if sn not in priority and sn not in SKIP_SHEETS:
            secondary.append(sn)

    for sheet_name in priority:
        df = all_dfs[sheet_name]
        if df.empty:
            continue
        col_map = _resolve_columns(df.columns)
        if 'waybill' not in col_map:
            continue
        sheets_processed += 1
        _process_df(df, col_map, order_map, remote_postcode_keys,
                    region_names, seen_waybills, stats)

    orders_missing_pc = {wb for wb, olist in order_map.items()
                         if any(not o.postcode for o in olist)}
    need_more = (total_order_waybills - seen_waybills) | orders_missing_pc
    for sheet_name in secondary:
        if not need_more:
            break
        df = all_dfs[sheet_name]
        if df.empty:
            continue
        col_map = _resolve_columns(df.columns)
        if 'waybill' not in col_map:
            continue
        sheets_processed += 1
        _process_df(df, col_map, order_map, remote_postcode_keys,
                    region_names, seen_waybills, stats)
        orders_missing_pc = {wb for wb, olist in order_map.items()
                             if any(not o.postcode for o in olist)}
        need_more = (total_order_waybills - seen_waybills) | orders_missing_pc

    db.session.commit()

    return {
        'updated_count': stats['updated'],
        'remote_marked': stats['remote_marked'],
        'not_found_count': stats['not_found'],
        'sheets_processed': sheets_processed,
        'version_used': version.name if version else None,
        'remote_postcodes_loaded': len(remote_postcode_keys),
    }
