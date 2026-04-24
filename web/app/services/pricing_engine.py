"""
计价引擎 - 根据科目+国家+客户查找规则并计算费用
所有费用由规则引擎计算，禁止使用代理价格。
"""
import math
from datetime import date
from app import db
from app.models import (PricingRule, PricingVersion, CustomerPricingOverride,
                         FeeCategory, Region, Order, OrderFee, ExchangeRate,
                         RemotePostcode)


def get_active_version(bill_date=None):
    """
    按订单账期日期查找对应的报价版本：
    effective_date <= bill_date，且 expire_date 为空或 > bill_date。
    """
    if bill_date:
        version = PricingVersion.query.filter(
            PricingVersion.effective_date <= bill_date,
            db.or_(
                PricingVersion.expire_date.is_(None),
                PricingVersion.expire_date > bill_date
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


def get_rule(version_id, category_code, region_id=None, cargo_type=None):
    """查找匹配的计价规则，优先级：完全匹配 > 无货物类型 > 无地区 > 全通配"""
    category = FeeCategory.query.filter_by(code=category_code).first()
    if not category:
        return None

    filters = [
        PricingRule.version_id == version_id,
        PricingRule.category_id == category.id,
    ]

    if region_id and cargo_type:
        rule = PricingRule.query.filter(
            *filters,
            PricingRule.region_id == region_id,
            PricingRule.cargo_type == cargo_type
        ).first()
        if rule:
            return rule

    if region_id:
        rule = PricingRule.query.filter(
            *filters,
            PricingRule.region_id == region_id,
            db.or_(PricingRule.cargo_type.is_(None), PricingRule.cargo_type == '')
        ).first()
        if rule:
            return rule

    if cargo_type:
        rule = PricingRule.query.filter(
            *filters,
            db.or_(PricingRule.region_id.is_(None)),
            PricingRule.cargo_type == cargo_type
        ).first()
        if rule:
            return rule

    return PricingRule.query.filter(
        *filters,
        db.or_(PricingRule.region_id.is_(None)),
        db.or_(PricingRule.cargo_type.is_(None), PricingRule.cargo_type == '')
    ).first()


def get_customer_override(customer_id, category_code, region_id=None, bill_date=None):
    """查找客户定制规则"""
    category = FeeCategory.query.filter_by(code=category_code).first()
    if not category:
        return None

    query = CustomerPricingOverride.query.filter(
        CustomerPricingOverride.customer_id == customer_id,
        CustomerPricingOverride.category_id == category.id,
    )

    if bill_date:
        query = query.filter(
            db.or_(
                CustomerPricingOverride.effective_date.is_(None),
                CustomerPricingOverride.effective_date <= bill_date
            ),
            db.or_(
                CustomerPricingOverride.expire_date.is_(None),
                CustomerPricingOverride.expire_date >= bill_date
            )
        )

    if region_id:
        override = query.filter(CustomerPricingOverride.region_id == region_id).first()
        if override:
            return override

    return query.filter(
        db.or_(CustomerPricingOverride.region_id.is_(None))
    ).first()


def get_exchange_rate(from_currency='EUR', to_currency='CNY', target_date=None):
    """获取汇率"""
    query = ExchangeRate.query.filter_by(
        from_currency=from_currency,
        to_currency=to_currency
    )
    if target_date:
        query = query.filter(ExchangeRate.date <= target_date)
    rate = query.order_by(ExchangeRate.date.desc()).first()
    return rate.rate if rate else None


def _normalize_postcode(pc_str):
    """Normalize postcode for comparison."""
    if not pc_str:
        return None
    pc = str(pc_str).strip()
    if pc.endswith('.0'):
        pc = pc[:-2]
    return pc if pc else None


def _get_order_country(order):
    """Get the country name for an order from its region."""
    if not order.region_id:
        return None
    region = Region.query.get(order.region_id)
    return region.name if region else None


def _find_remote_postcode(version_id, postcode, country):
    """Find a remote postcode record matching version + postcode + country."""
    if not postcode or not version_id:
        return None
    pc = _normalize_postcode(postcode)
    if not pc:
        return None
    query = RemotePostcode.query.filter_by(
        version_id=version_id,
        postcode=pc
    )
    if country:
        query = query.filter_by(country=country)
    return query.first()


def check_remote(order, version_id):
    """Check if order's postcode matches remote postcodes and update is_remote flag.
    Matching is scoped by country to avoid false positives from
    different countries sharing the same postcode numbers.
    Only orders with has_tail_freight (appeared in a 地派服务费 sheet) are
    eligible for remote fees — pure COD-only orders are excluded."""
    if not order.postcode or not version_id:
        return False
    if not order.has_tail_freight:
        order.is_remote = False
        return False
    country = _get_order_country(order)
    remote = _find_remote_postcode(version_id, order.postcode, country)
    if remote:
        order.is_remote = True
        return True
    order.is_remote = False
    return False


def calculate_remote_fee(order, version_id, exchange_rate=None):
    """
    Calculate remote area surcharge based on postcode + country → zone mapping.
    Reads surcharge_type and surcharge_amount from RemotePostcode record:
      - per_kg: surcharge_amount * charge_weight (EUR)
      - per_piece: surcharge_amount * pieces (EUR)
    """
    if not order.postcode or not version_id:
        return 0
    country = _get_order_country(order)
    remote = _find_remote_postcode(version_id, order.postcode, country)
    if not remote:
        return 0

    surcharge_type = remote.surcharge_type or 'per_kg'
    amount = remote.surcharge_amount or 0

    if surcharge_type == 'per_kg':
        weight = order.charge_weight_tail or order.actual_weight or 0
        return round(amount * weight, 2)
    elif surcharge_type == 'per_piece':
        pieces = order.pieces or 1
        return round(amount * pieces, 2)
    else:
        return round(amount, 2)


def calculate_fee(rule_or_override, order, exchange_rate=None, category_code=None):
    """
    根据规则计算费用。
    category_code 用于选择正确的计费重量：
      HEAD_FREIGHT -> charge_weight_head
      其他 -> charge_weight_tail
    """
    if hasattr(rule_or_override, 'get_params'):
        params = rule_or_override.get_params()
    else:
        params = rule_or_override

    rule_type = getattr(rule_or_override, 'rule_type', None) or params.get('_rule_type') or 'fixed'

    actual = order.actual_weight or 0
    if category_code == 'HEAD_FREIGHT':
        charge_w = order.charge_weight_head or actual
    else:
        charge_w = order.charge_weight_tail or actual

    if rule_type == 'per_kg':
        rate = params.get('rate_per_kg', 0)
        return round(charge_w * rate, 2)

    elif rule_type == 'first_extra':
        first_w = params.get('first_weight', 2)
        first_p = params.get('first_price', 0)
        extra_p = params.get('extra_per_kg', 0)
        weight = charge_w
        if weight <= first_w:
            return round(first_p, 2)
        extra_kg = math.ceil(weight) - first_w
        return round(first_p + extra_kg * extra_p, 2)

    elif rule_type == 'percentage':
        rate = params.get('rate', 0)
        min_amount = params.get('min_amount', 0)
        base_amount = params.get('base_amount') or 0
        if not base_amount and category_code == 'COD_FEE' and order.cod_amount:
            base_amount = abs(order.cod_amount)
        result = float(base_amount) * float(rate)
        if min_amount and (result < min_amount or base_amount == 0):
            return round(min_amount, 2)
        return round(result, 2)

    elif rule_type == 'fixed':
        amount = params.get('amount', 0)
        if exchange_rate and params.get('convert_to_rmb'):
            amount = round(amount * exchange_rate, 2)
        return round(amount, 2)

    elif rule_type == 'tiered':
        tiers = params.get('tiers', [])
        for tier in sorted(tiers, key=lambda t: t.get('max_weight', float('inf'))):
            if charge_w <= tier.get('max_weight', float('inf')):
                return round(tier.get('price', 0), 2)
        if tiers:
            return round(tiers[-1].get('price', 0), 2)
        return 0

    return 0


def _is_watch(order):
    """Check if the product is a watch based on product name."""
    name = (order.product_name or '').lower()
    return '手表' in name or '表' == name


COD_COUNTRIES = {'PL', 'IT', 'AT', 'DE'}


def _get_applicable_fees(order):
    """
    Determine all applicable fees for this order, including COD pre-charge.

    For countries with COD rules (PL/IT/AT/DE), when the order has
    head/tail freight, COD_FEE is always included — even before COD
    amount is confirmed. The minimum fee is charged upfront.
    """
    fees = list(order.applicable_fees)

    this_has_freight = order.has_head_freight or order.has_tail_freight
    region_code = order.region.code if order.region else None
    is_cod_country = region_code in COD_COUNTRIES

    if 'COD_FEE' not in fees and is_cod_country and this_has_freight:
        fees.append('COD_FEE')

    return fees


def calculate_order_fees(order_id, category_codes=None):
    """
    根据订单的跨期合并标记计算各项费用。

    关键规则：
    - 头尾程运费统一按 IC（特殊货）规则查找费率
    - 附加费：手表 30 RMB，其余 2 EUR（李志特殊 1.5 EUR），在有头尾程时收取
    - 偏远费：在地派表出现时即根据邮编收取，不等代理出偏远费科目
    - VAT 依赖尾程运费结果，在尾程之后计算
    - 意大利订单首次出现头尾程时即收取COD手续费（预收取），后续确认差值
    """
    order = Order.query.get(order_id)
    if not order:
        return []

    bill_date = order.bill_period or date.today()
    version = get_active_version(bill_date)
    if not version:
        return []

    exchange_rate = get_exchange_rate('EUR', 'CNY', bill_date)

    if order.postcode:
        check_remote(order, version.id)

    target_codes = _get_applicable_fees(order)
    if category_codes is not None:
        target_codes = [c for c in target_codes if c in set(category_codes)]

    cat_cache = {}
    for cat in FeeCategory.query.all():
        cat_cache[cat.code] = cat

    all_fees = OrderFee.query.filter_by(order_id=order.id).all()
    existing_fees = {}
    cod_fees = []
    for f in all_fees:
        if f.category:
            if f.category.code == 'COD_FEE':
                cod_fees.append(f)
            else:
                existing_fees[f.category.code] = f

    results = []
    for code in target_codes:
        cat = cat_cache.get(code)
        if not cat:
            continue

        if code == 'SECOND_DELIVERY':
            fee = existing_fees.get(code)
            if not fee:
                fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                db.session.add(fee)
            agent_input = fee.input_amount or 0
            if exchange_rate and agent_input:
                amount = round(agent_input * exchange_rate, 2)
            else:
                amount = agent_input
            fee.calculated_amount = amount
            if exchange_rate:
                fee.exchange_rate = exchange_rate
            db.session.flush()
            existing_fees[code] = fee
            results.append({'category': code, 'amount': amount, 'source': 'second_delivery_agent'})
            continue

        if code == 'REMOTE_FEE':
            amount = calculate_remote_fee(order, version.id, exchange_rate)
            fee = existing_fees.get(code)
            if not fee:
                fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                db.session.add(fee)
            fee.calculated_amount = amount
            if exchange_rate:
                fee.exchange_rate = exchange_rate
            db.session.flush()
            existing_fees[code] = fee
            results.append({'category': code, 'amount': amount, 'source': 'remote_postcode'})
            continue

        if code == 'F_SURCHARGE':
            if _is_watch(order):
                amount = 30.0
                source_note = 'watch_30rmb'
            else:
                customer = order.customer
                customer_name = customer.name if customer else ''
                if '李志' in customer_name:
                    eur_amount = 1.5
                else:
                    eur_amount = 2.0
                if exchange_rate:
                    amount = round(eur_amount * exchange_rate, 2)
                else:
                    amount = eur_amount
                source_note = f'surcharge_{eur_amount}eur'

            fee = existing_fees.get(code)
            if not fee:
                fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                db.session.add(fee)
            fee.calculated_amount = amount
            if exchange_rate:
                fee.exchange_rate = exchange_rate
            db.session.flush()
            existing_fees[code] = fee
            results.append({'category': code, 'amount': amount, 'source': source_note})
            continue

        if code == 'COD_FEE':
            cod_amt = order.cod_amount
            region_code = order.region.code if order.region else None
            is_cod_country = region_code in COD_COUNTRIES

            override_cod = None
            if order.customer_id:
                override_cod = get_customer_override(
                    order.customer_id, code, order.region_id, bill_date)
            rule_cod = get_rule(version.id, code, order.region_id, order.cargo_type)
            calc_source_cod = override_cod or rule_cod

            if calc_source_cod:
                if hasattr(calc_source_cod, 'get_params'):
                    params_cod = calc_source_cod.get_params()
                else:
                    params_cod = calc_source_cod
                min_amount_cod = params_cod.get('min_amount', 0)
            else:
                min_amount_cod = 0

            if is_cod_country:
                if not cod_amt:
                    amount = min_amount_cod
                    source_note = 'cod_min_precharge'
                    
                    precharge_fee = None
                    for f in cod_fees:
                        if f.notes and '预收取' in f.notes:
                            precharge_fee = f
                            break
                    
                    if not precharge_fee:
                        precharge_fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                        db.session.add(precharge_fee)
                        cod_fees.append(precharge_fee)
                    
                    precharge_fee.calculated_amount = amount
                    precharge_fee.notes = '预收取（最低费用）'
                    if exchange_rate:
                        precharge_fee.exchange_rate = exchange_rate
                    db.session.flush()
                    results.append({'category': code, 'amount': amount, 'source': source_note})
                else:
                    pct_amount = round(abs(cod_amt) * 0.03, 2)
                    amount = max(pct_amount, min_amount_cod)
                    
                    precharge_fee = None
                    for f in cod_fees:
                        if f.notes and '预收取' in f.notes:
                            precharge_fee = f
                            break
                    
                    if precharge_fee and precharge_fee.calculated_amount is not None:
                        diff_amount = amount - precharge_fee.calculated_amount
                        if diff_amount > 0:
                            difference_fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                            db.session.add(difference_fee)
                            difference_fee.calculated_amount = diff_amount
                            difference_fee.input_amount = cod_amt
                            difference_fee.notes = f'差额（实际{amount} - 预收取{precharge_fee.calculated_amount}）'
                            if exchange_rate:
                                difference_fee.exchange_rate = exchange_rate
                            db.session.flush()
                            results.append({'category': code, 'amount': diff_amount, 'source': 'cod_difference'})
                        else:
                            results.append({'category': code, 'amount': 0, 'source': 'cod_no_difference'})
                    else:
                        calculated_fee = None
                        for f in cod_fees:
                            if not f.notes or '预收取' not in f.notes:
                                calculated_fee = f
                                break
                        
                        if not calculated_fee:
                            calculated_fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                            db.session.add(calculated_fee)
                        
                        calculated_fee.calculated_amount = amount
                        calculated_fee.input_amount = cod_amt
                        calculated_fee.notes = '实际计算'
                        if exchange_rate:
                            calculated_fee.exchange_rate = exchange_rate
                        db.session.flush()
                        results.append({'category': code, 'amount': amount, 'source': 'cod_calculated'})
            else:
                if not cod_amt:
                    continue
                
                pct_amount = round(abs(cod_amt) * 0.03, 2)
                amount = max(pct_amount, min_amount_cod)
                
                calculated_fee = None
                for f in cod_fees:
                    calculated_fee = f
                    break
                
                if not calculated_fee:
                    calculated_fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                    db.session.add(calculated_fee)
                
                calculated_fee.calculated_amount = amount
                calculated_fee.input_amount = cod_amt
                calculated_fee.notes = '实际计算'
                if exchange_rate:
                    calculated_fee.exchange_rate = exchange_rate
                db.session.flush()
                results.append({'category': code, 'amount': amount, 'source': 'cod_calculated'})
            continue

        override = None
        if order.customer_id:
            override = get_customer_override(
                order.customer_id, code, order.region_id, bill_date)

        lookup_cargo_type = 'IC' if code in ('HEAD_FREIGHT', 'TAIL_FREIGHT') else order.cargo_type
        rule = get_rule(version.id, code, order.region_id, lookup_cargo_type)

        if not override and not rule and code == 'VAT':
            amount = 1.2
            fee = existing_fees.get(code)
            if not fee:
                fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
                db.session.add(fee)
            fee.calculated_amount = amount
            if exchange_rate:
                fee.exchange_rate = exchange_rate
            db.session.flush()
            existing_fees[code] = fee
            results.append({'category': code, 'amount': amount, 'source': 'default_vat'})
            continue

        calc_source = override if override else rule
        if not calc_source:
            continue

        amount = calculate_fee(calc_source, order, exchange_rate, category_code=code)

        fee = existing_fees.get(code)
        if not fee:
            fee = OrderFee(order_id=order.id, category_id=cat.id, input_currency='EUR')
            db.session.add(fee)

        fee.calculated_amount = amount
        if exchange_rate:
            fee.exchange_rate = exchange_rate
        db.session.flush()
        existing_fees[code] = fee

        results.append({
            'category': code,
            'amount': amount,
            'source': 'override' if override else 'rule',
        })

    db.session.commit()
    return results


def batch_calculate(order_ids, category_codes=None):
    """批量计算多个订单的费用"""
    results = {}
    for oid in order_ids:
        results[oid] = calculate_order_fees(oid, category_codes)
    return results
