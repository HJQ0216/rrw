import re
import zipfile

paths = [
    r"D:\钢结构RRW方向\制表\串联系统_RRW单参数计算需求表.xlsx",
    r"D:\钢结构RRW方向\制表\串联系统_RRW多参数计算需求表.xlsx",
]

for p in paths:
    with zipfile.ZipFile(p) as z:
        x = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    print("FILE=", p)
    print("C_COL=", re.findall(r"<[^>]*col[^>]*min=\"3\"[^>]*/>", x))
    print("PANES=", re.findall(r"<[^>]*pane[^>]*/>", x))
    print("SELECTIONS=", re.findall(r"<[^>]*selection[^>]*/>", x))
    print("HIDDEN_COLS=", re.findall(r"<[^>]*col[^>]*hidden=\"1\"[^>]*/>", x))
