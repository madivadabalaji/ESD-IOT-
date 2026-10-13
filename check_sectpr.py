import docx
from xml.etree import ElementTree as ET

doc = docx.Document(r'c:\Users\subha\Downloads\Telegram Desktop\PBL Final Documentation templet (3).docx')

for idx in [23, 24, 25, 27, 28]:
    p = doc.paragraphs[idx]
    sectPr = p._p.xpath('./w:pPr/w:sectPr')
    print(f"P{idx}: text={repr(p.text)}, has_sectPr={bool(sectPr)}")
