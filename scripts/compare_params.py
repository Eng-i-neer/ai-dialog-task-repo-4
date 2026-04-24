"""对比当前代码中的退件费参数 vs 报价表的实际数据"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 报价表中的真实数据
pricing = {
    '波兰':     {'first2kg': 4.0, 'extra1kg': 0.9, 'rule': '70%'},
    '罗马尼亚':  {'first2kg': 4.8, 'extra1kg': 0.6, 'rule': '70%'},
    '匈牙利':   {'first2kg': 4.4, 'extra1kg': 0.7, 'rule': '70%'},
    '捷克':     {'first2kg': 4.1, 'extra1kg': 0.8, 'rule': '70%'},
    '斯洛伐克':  {'first2kg': 4.3, 'extra1kg': 0.7, 'rule': '70%'},
    '保加利亚':  {'first2kg': 4.3, 'extra1kg': 0.8, 'rule': '70%'},
    '克罗地亚':  {'first2kg': 6.4, 'extra1kg': 1.0, 'rule': '70%'},
    '斯洛文尼亚': {'first2kg': 5.9, 'extra1kg': 0.9, 'rule': '70%'},
    '西班牙':   {'first2kg': 4.0, 'extra1kg': 1.0, 'rule': '同派送费'},
    '葡萄牙':   {'first2kg': 4.0, 'extra1kg': 1.0, 'rule': '同派送费'},
    '希腊':     {'first2kg': 5.7, 'extra1kg': 0.8, 'rule': '同派送费'},
    '意大利':   {'first2kg': 7.3, 'extra1kg': 1.1, 'rule': '同派送费'},
    '奥地利':   {'first2kg': 6.5, 'extra1kg': 1.0, 'rule': '同派送费'},
    '德国':     {'first2kg': 8.3, 'extra1kg': 1.5, 'rule': '同派送费'},
}

# 当前代码中 _build_return_fee_formula 使用的参数
code_params = {
    '德国':     (8, 1.5, False),
    '奥地利':   (6.5, 1.0, False),
    '意大利':   (6.7, 1.0, False),
    '西班牙':   (4.0, 0.6, False),
    '葡萄牙':   (4.0, 0.6, False),
    '希腊':     (5.7, 0.8, False),
    '克罗地亚':  (5.6, 0.9, True),
    '斯洛文尼亚': (5.9, 0.8, True),
    '波兰':     (4.0, 0.6, True),
    '罗马尼亚':  (4.8, 0.7, True),
    '匈牙利':   (4.4, 0.6, True),
    '捷克':     (4.1, 0.6, True),
    '斯洛伐克':  (4.3, 0.6, True),
    '保加利亚':  (4.3, 0.6, True),
}

print(f"{'国家':<12} {'报价首2KG':>10} {'代码首2KG':>10} {'MATCH?':>8} {'报价续1KG':>10} {'代码续1KG':>10} {'MATCH?':>8} {'报价规则':>10} {'代码70%':>8}")
print("-"*100)
for country in pricing:
    p = pricing[country]
    c = code_params.get(country)
    if c:
        base, step, is_70 = c
        match_base = 'OK' if abs(p['first2kg'] - base) < 0.01 else 'DIFF'
        match_step = 'OK' if abs(p['extra1kg'] - step) < 0.01 else 'DIFF'
        code_rule = '70%' if is_70 else '同派送费'
        match_rule = 'OK' if p['rule'] == code_rule else 'DIFF'
        
        if match_base != 'OK' or match_step != 'OK':
            flag = ' <<<'
        else:
            flag = ''
        
        print(f"{country:<12} {p['first2kg']:>10} {base:>10} {match_base:>8} {p['extra1kg']:>10} {step:>10} {match_step:>8} {p['rule']:>10} {code_rule:>8}{flag}")
    else:
        print(f"{country:<12} {p['first2kg']:>10} {'N/A':>10} {'MISS':>8} {p['extra1kg']:>10} {'N/A':>10} {'MISS':>8}")
