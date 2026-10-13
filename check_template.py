import docx

doc = docx.Document(r'c:\Users\subha\Downloads\Telegram Desktop\PBL Final Documentation templet (3).docx')
print("Total sections:", len(doc.sections))

for idx, sec in enumerate(doc.sections):
    print(f"\n--- SECTION {idx} ---")
    print(f"Start type: {sec.start_type}")
    print(f"Cols: {sec._sectPr.xpath('./w:cols/@w:num')}")
    print(f"Margins (L,R,T,B): {sec.left_margin.inches:.2f}, {sec.right_margin.inches:.2f}, {sec.top_margin.inches:.2f}, {sec.bottom_margin.inches:.2f}")

# Map paragraphs to sections by looking at sectPr
p_idx = 0
for i, p in enumerate(doc.paragraphs):
    sectPr = p._p.xpath('./w:pPr/w:sectPr')
    if sectPr:
        print(f"Paragraph P{i:02d} terminates section! Text: {repr(p.text[:40])}")
