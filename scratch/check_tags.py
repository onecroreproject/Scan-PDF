
import re

with open(r"r:\Balakrishnan\1CP_Project\All_In_One_PDF Scan Pdf\All_In_One_PDF\templates\converter\base.html", "r", encoding="utf-8") as f:
    content = f.read()

divs = len(re.findall(r"<div", content))
closers = len(re.findall(r"</div", content))

print(f"Divs: {divs}, Closers: {closers}")
