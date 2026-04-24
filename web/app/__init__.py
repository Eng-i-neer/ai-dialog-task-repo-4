import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate

db = SQLAlchemy()
migrate = Migrate()


def _auto_add_columns(app):
    """Add new columns to existing SQLite tables without losing data."""
    import sqlite3
    db_path = app.config.get('SQLALCHEMY_DATABASE_URI', '').replace('sqlite:///', '')
    if not db_path or not os.path.exists(db_path):
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    migrations = [
        ('order_fees', 'import_log_id', 'INTEGER'),
        ('order_fees', 'import_period', 'VARCHAR(20)'),
        ('order_fees', 'source_sheet', 'VARCHAR(100)'),
        ('orders', 'import_periods', 'TEXT'),
        ('orders', 'needs_second_delivery', 'BOOLEAN DEFAULT 0'),
    ]

    for table, column, col_type in migrations:
        try:
            cursor.execute(f"SELECT {column} FROM {table} LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")

    conn.commit()
    conn.close()


def create_app():
    app = Flask(__name__)
    app.config.from_object('app.config.Config')

    db.init_app(app)
    migrate.init_app(app, db)

    from app.routes import main_bp, orders_bp, customers_bp, regions_bp, pricing_bp, imports_bp, exports_bp, exchange_rates_bp
    app.register_blueprint(main_bp)
    app.register_blueprint(orders_bp, url_prefix='/orders')
    app.register_blueprint(customers_bp, url_prefix='/customers')
    app.register_blueprint(regions_bp, url_prefix='/regions')
    app.register_blueprint(pricing_bp, url_prefix='/pricing')
    app.register_blueprint(imports_bp, url_prefix='/imports')
    app.register_blueprint(exports_bp, url_prefix='/exports')
    app.register_blueprint(exchange_rates_bp, url_prefix='/exchange-rates')

    with app.app_context():
        _auto_add_columns(app)

        from app import models
        db.create_all()

        from app.services.seed_data import init_seed_data
        init_seed_data()

    return app
