import json
from datetime import datetime, date
from app import db


class Customer(db.Model):
    __tablename__ = 'customers'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    code = db.Column(db.String(50))
    currency = db.Column(db.String(10), default='EUR')
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    orders = db.relationship('Order', backref='customer', lazy='dynamic')
    pricing_overrides = db.relationship('CustomerPricingOverride', backref='customer', lazy='dynamic')

    def to_dict(self):
        return {
            'id': self.id, 'name': self.name, 'code': self.code,
            'currency': self.currency, 'notes': self.notes,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }


class Region(db.Model):
    __tablename__ = 'regions'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    code = db.Column(db.String(10))
    currency = db.Column(db.String(10))
    vat_rate = db.Column(db.Float)
    return_rule = db.Column(db.String(10))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    orders = db.relationship('Order', backref='region', lazy='dynamic')

    def to_dict(self):
        return {
            'id': self.id, 'name': self.name, 'code': self.code,
            'currency': self.currency, 'vat_rate': self.vat_rate,
            'return_rule': self.return_rule
        }


class Order(db.Model):
    __tablename__ = 'orders'
    id = db.Column(db.Integer, primary_key=True)
    waybill_no = db.Column(db.String(50), unique=True, nullable=False)
    transfer_no = db.Column(db.String(100))
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'))
    region_id = db.Column(db.Integer, db.ForeignKey('regions.id'))
    import_log_id = db.Column(db.Integer, db.ForeignKey('import_logs.id'))
    ship_date = db.Column(db.Date)
    bill_period = db.Column(db.Date)
    import_periods = db.Column(db.Text)
    ship_type = db.Column(db.String(10))
    product_name = db.Column(db.String(200))
    cargo_type = db.Column(db.String(10))
    pieces = db.Column(db.Integer, default=1)
    actual_weight = db.Column(db.Float)
    charge_weight_head = db.Column(db.Float)
    charge_weight_tail = db.Column(db.Float)
    dimensions = db.Column(db.String(50))
    customer_ref = db.Column(db.String(100))
    postcode = db.Column(db.String(20))
    address = db.Column(db.String(500))
    cod_amount = db.Column(db.Float)
    cod_currency = db.Column(db.String(10))
    cargo_status = db.Column(db.String(50))
    cargo_status_category = db.Column(db.String(50))
    is_remote = db.Column(db.Boolean, default=False)
    has_head_freight = db.Column(db.Boolean, default=False)
    has_tail_freight = db.Column(db.Boolean, default=False)
    needs_return_fee = db.Column(db.Boolean, default=False)
    needs_shelf_fee = db.Column(db.Boolean, default=False)
    needs_vat = db.Column(db.Boolean, default=False)
    needs_second_delivery = db.Column(db.Boolean, default=False)
    import_sheets = db.Column(db.Text)
    logistics_status = db.Column(db.String(50))
    notes = db.Column(db.Text)
    source_file = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    fees = db.relationship('OrderFee', backref='order', lazy='dynamic', cascade='all, delete-orphan')
    import_log = db.relationship('ImportLog', backref=db.backref('orders', lazy='dynamic'))

    @property
    def import_sheet_list(self):
        """Return list of agent sheet names this order appeared in."""
        if not self.import_sheets:
            return []
        return [s.strip() for s in self.import_sheets.split(',') if s.strip()]

    def add_import_sheet(self, sheet_name):
        """Append a sheet name to the import_sheets list (dedup)."""
        existing = set(self.import_sheet_list)
        if sheet_name not in existing:
            existing.add(sheet_name)
            self.import_sheets = ','.join(sorted(existing))

    @property
    def import_period_list(self):
        """Return sorted list of bill periods this order appeared in."""
        if not self.import_periods:
            if self.bill_period:
                return [self.bill_period.strftime('%Y%m%d')]
            return []
        return [s.strip() for s in self.import_periods.split(',') if s.strip()]

    def add_import_period(self, period_str):
        """Record that this order appeared in a given bill period."""
        existing = set(self.import_period_list)
        if period_str not in existing:
            existing.add(period_str)
            self.import_periods = ','.join(sorted(existing))

    @property
    def applicable_fees(self):
        """Return list of our fee codes applicable to this order."""
        fees = []
        if self.has_head_freight:
            fees.append('HEAD_FREIGHT')
        if self.has_tail_freight:
            fees.append('TAIL_FREIGHT')
        if 'COD' in (self.import_sheets or ''):
            fees.append('COD_FEE')
        if self.needs_return_fee:
            fees.append('RETURN_FEE')
        if self.needs_shelf_fee:
            fees.append('SHELF_FEE')
        if self.needs_vat:
            fees.append('VAT')
        if self.is_remote and self.has_tail_freight:
            fees.append('REMOTE_FEE')
        if self.has_head_freight or self.has_tail_freight:
            fees.append('F_SURCHARGE')
        if self.needs_second_delivery:
            fees.append('SECOND_DELIVERY')
        return fees

    @property
    def receivable_total(self):
        """Sum of all calculated fee amounts."""
        total = 0.0
        for f in self.fees.all():
            val = f.override_amount if f.override_amount is not None else f.calculated_amount
            if val:
                total += val
        return round(total, 2)

    def to_dict(self):
        return {
            'id': self.id, 'waybill_no': self.waybill_no,
            'transfer_no': self.transfer_no,
            'customer': self.customer.name if self.customer else None,
            'customer_id': self.customer_id,
            'region': self.region.name if self.region else None,
            'region_id': self.region_id,
            'ship_date': self.ship_date.isoformat() if self.ship_date else None,
            'bill_period': self.bill_period.isoformat() if self.bill_period else None,
            'import_periods': self.import_period_list,
            'ship_type': self.ship_type,
            'product_name': self.product_name,
            'cargo_type': self.cargo_type,
            'pieces': self.pieces,
            'actual_weight': self.actual_weight,
            'charge_weight_head': self.charge_weight_head,
            'charge_weight_tail': self.charge_weight_tail,
            'dimensions': self.dimensions,
            'customer_ref': self.customer_ref,
            'postcode': self.postcode,
            'address': self.address,
            'cod_amount': self.cod_amount,
            'cod_currency': self.cod_currency,
            'cargo_status': self.cargo_status,
            'is_remote': self.is_remote,
            'logistics_status': self.logistics_status,
            'notes': self.notes,
            'source_file': self.source_file,
            'has_head_freight': self.has_head_freight,
            'has_tail_freight': self.has_tail_freight,
            'needs_return_fee': self.needs_return_fee,
            'needs_shelf_fee': self.needs_shelf_fee,
            'needs_vat': self.needs_vat,
            'needs_second_delivery': self.needs_second_delivery,
            'import_sheets': self.import_sheet_list,
            'applicable_fees': self.applicable_fees,
            'receivable_total': self.receivable_total,
        }


class FeeCategory(db.Model):
    __tablename__ = 'fee_categories'
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    group = db.Column(db.String(50))
    description = db.Column(db.Text)

    def to_dict(self):
        return {
            'id': self.id, 'code': self.code,
            'name': self.name, 'group': self.group,
            'description': self.description
        }


class OrderFee(db.Model):
    __tablename__ = 'order_fees'
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('fee_categories.id'), nullable=False)
    import_log_id = db.Column(db.Integer, db.ForeignKey('import_logs.id'))
    import_period = db.Column(db.String(20))
    source_sheet = db.Column(db.String(100))
    input_amount = db.Column(db.Float)
    input_currency = db.Column(db.String(10), default='EUR')
    calculated_amount = db.Column(db.Float)
    output_currency = db.Column(db.String(10))
    exchange_rate = db.Column(db.Float)
    override_amount = db.Column(db.Float)
    is_manual = db.Column(db.Boolean, default=False)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    category = db.relationship('FeeCategory')
    import_log = db.relationship('ImportLog', backref=db.backref('order_fees', lazy='dynamic'))

    @property
    def final_amount(self):
        return self.override_amount if self.override_amount is not None else self.calculated_amount

    @property
    def period_label(self):
        """Human-readable period label like '0728期'."""
        if self.import_period:
            return f"{self.import_period[4:]}期"
        return None

    def to_dict(self):
        return {
            'id': self.id, 'order_id': self.order_id,
            'category': self.category.name if self.category else None,
            'category_code': self.category.code if self.category else None,
            'import_log_id': self.import_log_id,
            'import_period': self.import_period,
            'period_label': self.period_label,
            'source_sheet': self.source_sheet,
            'input_amount': self.input_amount,
            'input_currency': self.input_currency,
            'calculated_amount': self.calculated_amount,
            'output_currency': self.output_currency,
            'exchange_rate': self.exchange_rate,
            'override_amount': self.override_amount,
            'is_manual': self.is_manual,
            'final_amount': self.final_amount,
            'notes': self.notes
        }


class PricingVersion(db.Model):
    __tablename__ = 'pricing_versions'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    effective_date = db.Column(db.Date, nullable=False)
    expire_date = db.Column(db.Date)
    is_active = db.Column(db.Boolean, default=False, nullable=False)
    source_file = db.Column(db.String(200))
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    rules = db.relationship('PricingRule', backref='version', lazy='dynamic', cascade='all, delete-orphan')

    def to_dict(self):
        return {
            'id': self.id, 'name': self.name,
            'effective_date': self.effective_date.isoformat() if self.effective_date else None,
            'expire_date': self.expire_date.isoformat() if self.expire_date else None,
            'is_active': self.is_active,
            'source_file': self.source_file, 'notes': self.notes,
            'rules_count': self.rules.count()
        }


class PricingRule(db.Model):
    __tablename__ = 'pricing_rules'
    id = db.Column(db.Integer, primary_key=True)
    version_id = db.Column(db.Integer, db.ForeignKey('pricing_versions.id'), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('fee_categories.id'), nullable=False)
    region_id = db.Column(db.Integer, db.ForeignKey('regions.id'))
    cargo_type = db.Column(db.String(10))
    rule_type = db.Column(db.String(30), nullable=False)
    params = db.Column(db.Text)

    category = db.relationship('FeeCategory')
    region = db.relationship('Region')

    def get_params(self):
        if self.params:
            try:
                return json.loads(self.params)
            except (json.JSONDecodeError, ValueError):
                return {}
        return {}

    def set_params(self, params_dict):
        self.params = json.dumps(params_dict, ensure_ascii=False)

    def to_dict(self):
        return {
            'id': self.id, 'version_id': self.version_id,
            'category': self.category.name if self.category else None,
            'category_code': self.category.code if self.category else None,
            'region': self.region.name if self.region else None,
            'region_id': self.region_id,
            'cargo_type': self.cargo_type,
            'rule_type': self.rule_type,
            'params': self.get_params()
        }


class CustomerPricingOverride(db.Model):
    __tablename__ = 'customer_pricing_overrides'
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('fee_categories.id'), nullable=False)
    region_id = db.Column(db.Integer, db.ForeignKey('regions.id'))
    cargo_type = db.Column(db.String(10))
    rule_type = db.Column(db.String(30), default='fixed')
    params = db.Column(db.Text)
    effective_date = db.Column(db.Date)
    expire_date = db.Column(db.Date)
    notes = db.Column(db.Text)

    category = db.relationship('FeeCategory')
    region = db.relationship('Region')

    def get_params(self):
        if self.params:
            try:
                return json.loads(self.params)
            except (json.JSONDecodeError, ValueError):
                return {}
        return {}

    def set_params(self, params_dict):
        self.params = json.dumps(params_dict, ensure_ascii=False)

    def to_dict(self):
        return {
            'id': self.id, 'customer_id': self.customer_id,
            'category': self.category.name if self.category else None,
            'region': self.region.name if self.region else None,
            'cargo_type': self.cargo_type,
            'rule_type': self.rule_type,
            'params': self.get_params(),
            'effective_date': self.effective_date.isoformat() if self.effective_date else None,
            'expire_date': self.expire_date.isoformat() if self.expire_date else None
        }


class RemotePostcode(db.Model):
    __tablename__ = 'remote_postcodes'
    id = db.Column(db.Integer, primary_key=True)
    version_id = db.Column(db.Integer, db.ForeignKey('pricing_versions.id'), nullable=False)
    postcode = db.Column(db.String(20), nullable=False)
    country = db.Column(db.String(50))
    zone = db.Column(db.String(50))
    surcharge_type = db.Column(db.String(20))
    surcharge_amount = db.Column(db.Float)

    version = db.relationship('PricingVersion', backref=db.backref(
        'remote_postcodes', lazy='dynamic', cascade='all, delete-orphan'))

    __table_args__ = (
        db.Index('ix_remote_postcode_lookup', 'version_id', 'postcode'),
    )

    def to_dict(self):
        return {
            'id': self.id, 'version_id': self.version_id,
            'postcode': self.postcode, 'country': self.country,
            'zone': self.zone, 'surcharge_type': self.surcharge_type,
            'surcharge_amount': self.surcharge_amount
        }


class ExchangeRate(db.Model):
    __tablename__ = 'exchange_rates'
    id = db.Column(db.Integer, primary_key=True)
    from_currency = db.Column(db.String(10), nullable=False)
    to_currency = db.Column(db.String(10), nullable=False)
    rate = db.Column(db.Float, nullable=False)
    date = db.Column(db.Date, nullable=False)
    source = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id, 'from_currency': self.from_currency,
            'to_currency': self.to_currency, 'rate': self.rate,
            'date': self.date.isoformat() if self.date else None,
            'source': self.source
        }


class ExportTemplate(db.Model):
    __tablename__ = 'export_templates'
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'))
    category_id = db.Column(db.Integer, db.ForeignKey('fee_categories.id'))
    template_file = db.Column(db.String(200))
    column_mapping = db.Column(db.Text)
    formula_config = db.Column(db.Text)
    notes = db.Column(db.Text)

    customer = db.relationship('Customer')
    category = db.relationship('FeeCategory')


class ImportLog(db.Model):
    __tablename__ = 'import_logs'
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200))
    file_type = db.Column(db.String(30))
    bill_period = db.Column(db.Date)
    orders_count = db.Column(db.Integer)
    status = db.Column(db.String(20))
    error_log = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id, 'filename': self.filename,
            'file_type': self.file_type,
            'bill_period': self.bill_period.isoformat() if self.bill_period else None,
            'orders_count': self.orders_count,
            'status': self.status,
            'error_log': self.error_log,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }
