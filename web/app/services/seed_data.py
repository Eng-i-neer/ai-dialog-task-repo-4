"""
种子数据初始化 - 预置地区、科目和客户。
科目通过 code 进行 upsert，确保新版本始终覆盖旧数据。
"""
from app import db
from app.models import Region, FeeCategory, Customer

CORE_CATEGORIES = [
    ('HEAD_FREIGHT', '头程运费', '运费', '中国→目的国干线运输，按头程计费重×单价'),
    ('TAIL_FREIGHT', '尾程运费', '运费', '目的国末端派送，首重+续重（尾程计费重）'),
    ('COD_FEE', 'COD手续费', 'COD', '代收货款手续费，按代收金额×费率，含汇率差调整'),
    ('RETURN_FEE', '退件费', '杂费', '拒收返程/退件入仓费，导入代理"尾程退件操作费"表标记'),
    ('SHELF_FEE', '上架费', '杂费', '导入代理"上架费"表中的订单需收取'),
    ('F_SURCHARGE', '附加费', '杂费', '李志:普通1.5EUR特殊30CNY; 其余:2EUR/30CNY'),
    ('VAT', '增值税', '税费', '目的国增值税，导入代理"目的地增值税"表标记'),
    ('REMOTE_FEE', '偏远费', '杂费', '由货况表邮编匹配报价偏远邮编表确定'),
    ('SECOND_DELIVERY', '二次派送费', '杂费', '代理账单"二派费"表标记，二次派送产生的额外费用'),
]


def init_seed_data():
    """Upsert fee categories, seed regions/customers only if empty."""
    changed = False

    existing = {c.code: c for c in FeeCategory.query.all()}
    valid_codes = {code for code, *_ in CORE_CATEGORIES}
    for code, name, group, desc in CORE_CATEGORIES:
        if code in existing:
            cat = existing[code]
            cat.name, cat.group, cat.description = name, group, desc
        else:
            db.session.add(FeeCategory(code=code, name=name, group=group, description=desc))
            changed = True

    for code, cat in existing.items():
        if code not in valid_codes:
            db.session.delete(cat)
            changed = True

    if Region.query.count() == 0:
        regions = [
            Region(name='德国', code='DE', currency='EUR', vat_rate=0.19, return_rule='100%'),
            Region(name='意大利', code='IT', currency='EUR', vat_rate=0.22, return_rule='100%'),
            Region(name='西班牙', code='ES', currency='EUR', vat_rate=0.21, return_rule='100%'),
            Region(name='葡萄牙', code='PT', currency='EUR', vat_rate=0.23, return_rule='100%'),
            Region(name='克罗地亚', code='HR', currency='EUR', vat_rate=0.25, return_rule='70%'),
            Region(name='希腊', code='GR', currency='EUR', vat_rate=0.24, return_rule='70%'),
            Region(name='斯洛文尼亚', code='SI', currency='EUR', vat_rate=0.22, return_rule='70%'),
            Region(name='匈牙利', code='HU', currency='HUF', vat_rate=0.27, return_rule='70%'),
            Region(name='捷克', code='CZ', currency='CZK', vat_rate=0.21, return_rule='70%'),
            Region(name='斯洛伐克', code='SK', currency='EUR', vat_rate=0.20, return_rule='70%'),
            Region(name='罗马尼亚', code='RO', currency='RON', vat_rate=0.19, return_rule='70%'),
            Region(name='保加利亚', code='BG', currency='BGN', vat_rate=0.20, return_rule='70%'),
            Region(name='奥地利', code='AT', currency='EUR', vat_rate=0.20, return_rule='100%'),
            Region(name='波兰', code='PL', currency='PLN', vat_rate=0.23, return_rule='70%'),
        ]
        db.session.add_all(regions)
        changed = True

    if Customer.query.count() == 0:
        customers = [
            Customer(name='李志', code='中文', currency='CNY', notes='大客户，RMB结算，附加费加入尾程杂费'),
            Customer(name='君悦', code='中文1', currency='EUR'),
            Customer(name='小美', code='中文3', currency='EUR'),
            Customer(name='小魏', code='中文4', currency='EUR'),
            Customer(name='J-欧洲', code='中文5', currency='EUR'),
            Customer(name='涵江', code='中文6', currency='EUR'),
            Customer(name='阿甘', code='中文7', currency='EUR'),
            Customer(name='欧弟', code='中文8', currency='EUR'),
            Customer(name='威总', code='中文9', currency='EUR'),
            Customer(name='振总', code='中文10', currency='EUR'),
            Customer(name='峰总', code='中文12', currency='EUR'),
            Customer(name='刚总', code='中文14', currency='EUR'),
            Customer(name='李总-1', code='中文15', currency='EUR'),
            Customer(name='香隅', code='XY-1', currency='EUR'),
        ]
        db.session.add_all(customers)
        changed = True

    db.session.commit()
    return changed
