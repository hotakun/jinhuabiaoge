"""Python 对照版本 —— 同样功能，对比语法和性能"""
import sys, csv

if len(sys.argv) < 2:
    print("用法: python compare.py <csv文件>")
    sys.exit(1)

items = {}  # dict: name -> qty

with open(sys.argv[1], 'r', encoding='utf-8-sig') as f:
    reader = csv.reader(f)
    header = next(reader)
    col_name = next((i for i, h in enumerate(header) if '商品名称' in h), -1)
    col_qty  = next((i for i, h in enumerate(header) if '订货数量' in h), -1)
    if col_name < 0 or col_qty < 0:
        print("找不到必要列")
        sys.exit(1)

    next(reader, None)  # skip sub-header
    for row in reader:
        try:
            name = row[col_name]
            qty = int(float(row[col_qty]))
            items[name] = items.get(name, 0) + qty
        except: pass

for name, qty in sorted(items.items()):
    print(f"{name:<35} {qty:>5}")

total = sum(items.values())
print(f"共 {len(items)} 种商品, 出库总计 {total}")
