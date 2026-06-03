from docx import Document

doc = Document()
doc.add_paragraph("render smoke test")
doc.save("render_tmp_ascii/simple.docx")
