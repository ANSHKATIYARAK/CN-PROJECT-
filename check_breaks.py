import docx

files = [
    'BAITE203_Lab_Assessment_Part1_Problem_Analysis_Design_and_Configuration_Software_Proof.docx',
    'BAITE203_Lab_Assessment_Part2_Implementation_and_Performance_Evaluation_Software_Proof.docx',
    'BAITE203_Lab_Assessment_Part3_Presentation_Innovation_and_Optimization_Software_Proof.docx'
]

for fname in files:
    doc = docx.Document(fname)
    print('='*60)
    print('FILE:', fname)
    # Check paragraphs and tables in order
    body_elements = doc._body._element
    element_idx = 0
    for child in body_elements:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            p = docx.text.paragraph.Paragraph(child, doc)
            brs = child.xpath('.//w:br[@w:type="page"]')
            if brs:
                print(f"Elem {element_idx} (p): [PAGE BREAK] Text: '{p.text[:60]}'")
            elif p.style.name.startswith('Heading') or (len(p.runs) > 0 and p.runs[0].bold and p.paragraph_format.space_before and p.paragraph_format.space_before.pt > 10):
                print(f"Elem {element_idx} (Heading): '{p.text[:60]}'")
        elif tag == 'tbl':
            tbl = docx.table.Table(child, doc)
            cant_splits = child.xpath('.//w:cantSplit')
            print(f"Elem {element_idx} (tbl): {len(tbl.rows)} rows x {len(tbl.columns)} cols, cantSplit={len(cant_splits)}")
        element_idx += 1
