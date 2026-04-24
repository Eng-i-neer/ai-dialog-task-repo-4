from flask import Blueprint

main_bp = Blueprint('main', __name__)
orders_bp = Blueprint('orders', __name__)
customers_bp = Blueprint('customers', __name__)
regions_bp = Blueprint('regions', __name__)
pricing_bp = Blueprint('pricing', __name__)
imports_bp = Blueprint('imports', __name__)
exports_bp = Blueprint('exports', __name__)
exchange_rates_bp = Blueprint('exchange_rates', __name__)

from app.routes import main, orders, customers, regions, pricing, imports_routes, exports_routes, exchange_rates
