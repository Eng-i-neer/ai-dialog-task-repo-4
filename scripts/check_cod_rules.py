# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'web'))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['FLASK_ENV'] = 'testing'
from app import create_app
app = create_app()

with app.app_context():
    from app.models import PricingRule, FeeCategory, Region

    cod_cat = FeeCategory.query.filter_by(code='COD_FEE').first()
    if cod_cat:
        print(f"COD_FEE category: id={cod_cat.id}, name={cod_cat.name}")
        rules = PricingRule.query.filter_by(category_id=cod_cat.id).all()
        print(f"COD rules ({len(rules)}):")
        for r in rules:
            region = Region.query.get(r.region_id) if r.region_id else None
            rname = region.name if region else 'ALL'
            print(f"  version={r.version_id}, region={rname}, cargo={r.cargo_type}, "
                  f"rule_type={r.rule_type}, params={r.params}")
    else:
        print("No COD_FEE category found!")

    print("\nAll fee categories:")
    for c in FeeCategory.query.all():
        print(f"  {c.code}: {c.name}")
