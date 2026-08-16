from docx import Document

# Load an existing .docx file
doc = Document('output.docx')

# Read and print paragraphs
print("--- Paragraphs ---")
for paragraph in doc.paragraphs:
    if paragraph.text.strip():
        print(paragraph.text)

# Read and print table data
print("\n--- Tables ---")
for table in doc.tables:
    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        print(row_data)
