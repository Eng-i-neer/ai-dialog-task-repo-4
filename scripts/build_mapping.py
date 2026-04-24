"""建立中介文件 -> 原始模板 -> 客户名的映射"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'c:\Users\59571\Desktop\deutsch-app\舅妈网站'

# 对照文档映射
code_to_customer = {
    '中文': '李志',
    '中文1': '君悦',
    '中文3': '小美',
    '中文4': '小魏',
    '中文5': 'J',  # J-欧洲 -> 模板中是 J
    '中文6': '涵江',
    '中文7': '阿甘',
    '中文8': '欧弟',
    '中文9': '威总',
    '中文10': '振总',
    '中文12': '峰总',
    '中文14': '刚总',
    '中文15': '李总-1',
}

# 中介文件列表
input_dir = os.path.join(BASE, '中介提供')
input_files = {}
for f in os.listdir(input_dir):
    if f.endswith('.xlsx'):
        # 提取标识: "鑫腾跃 -中文-对账单20260330.xlsx" -> "中文"
        # "鑫腾跃 -中文1-对账单20260330.xlsx" -> "中文1"
        # "鑫腾跃 -XY-1-对账单20260330.xlsx" -> "XY-1" (特殊)
        parts = f.replace('鑫腾跃 -', '').replace('-对账单20260330.xlsx', '')
        input_files[parts] = os.path.join(input_dir, f)

# 原始模板列表
template_dir = os.path.join(BASE, '反馈客户', '原始模板')
template_files = {}
for f in os.listdir(template_dir):
    if f.endswith('.xlsx'):
        template_files[f] = os.path.join(template_dir, f)

print("=== 中介文件 ===")
for code, path in sorted(input_files.items()):
    customer = code_to_customer.get(code, '???')
    print(f"  {code:>8} -> 客户: {customer:>6} | {os.path.basename(path)}")

print(f"\n=== 原始模板 ({len(template_files)} 个) ===")
for name in sorted(template_files.keys()):
    print(f"  {name}")

# 建立映射
print(f"\n=== 匹配结果 ===")
matched = []
unmatched_inputs = []
for code, input_path in sorted(input_files.items()):
    customer = code_to_customer.get(code)
    if not customer:
        unmatched_inputs.append((code, input_path, '对照文档中无此标识'))
        continue
    
    # 在模板中查找包含客户名的文件
    found_template = None
    for tname, tpath in template_files.items():
        if customer in tname:
            found_template = (tname, tpath)
            break
    
    if found_template:
        matched.append({
            'code': code,
            'customer': customer,
            'input': input_path,
            'template': found_template[1],
            'template_name': found_template[0],
        })
        print(f"  OK  {code:>8} -> {customer:>6} | 输入: {os.path.basename(input_path)}")
        print(f"       {'':>8}    {'':>6} | 模板: {found_template[0]}")
    else:
        unmatched_inputs.append((code, input_path, f'客户"{customer}"无对应模板'))

if unmatched_inputs:
    print(f"\n=== 未匹配 ({len(unmatched_inputs)} 个) ===")
    for code, path, reason in unmatched_inputs:
        customer = code_to_customer.get(code, '???')
        print(f"  X   {code:>8} -> {customer:>6} | {reason} | {os.path.basename(path)}")

# 检查有模板但无对应中介文件的
matched_templates = set(m['template_name'] for m in matched)
unmatched_templates = [n for n in template_files.keys() if n not in matched_templates]
if unmatched_templates:
    print(f"\n=== 有模板但无中介文件 ({len(unmatched_templates)} 个) ===")
    for t in sorted(unmatched_templates):
        print(f"  {t}")

print(f"\n=== 总结 ===")
print(f"  中介文件总数: {len(input_files)}")
print(f"  原始模板总数: {len(template_files)}")
print(f"  成功匹配: {len(matched)} 对")
print(f"  未匹配输入: {len(unmatched_inputs)}")
print(f"  未匹配模板: {len(unmatched_templates)}")
