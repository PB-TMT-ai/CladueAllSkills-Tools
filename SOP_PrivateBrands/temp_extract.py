from docx import Document

doc = Document(r"D:\SOP_PrivateBrands\Documents\JSWO_Plant Operation_to add pictures (1).docx")

# Table 3 (index 2) has activities 11-11.E
table = doc.tables[2]
print(f"Table 3 has {len(table.rows)} rows")

for r_idx, row in enumerate(table.rows):
    print(f"\n--- Row {r_idx} ---")
    for c_idx, cell in enumerate(row.cells):
        text = cell.text.strip()
        print(f"  Col {c_idx}: '''{text}'''")
