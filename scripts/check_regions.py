# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'

from app import create_app
app = create_app()

with app.app_context():
    from app.models import Region
    for r in Region.query.all():
        print(f"  id={r.id}, name={r.name}, code={r.code}")
