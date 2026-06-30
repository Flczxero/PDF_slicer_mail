import fitz # PyMuPDF
import os

main_pdf_path = os.path.join("..", "1-s2.0-S2772443321000027-main.pdf")
doc = fitz.open(main_pdf_path)

# Create a new empty PDF
doc2 = fitz.open()

# Insert the first page of doc into doc2
doc2.insert_pdf(doc, from_page=0, to_page=0)

out_name = "test-pymupdf.pdf"
# Save with garbage collection enabled to strip unused resources
doc2.save(out_name, garbage=4, deflate=True)

print("Size:", os.path.getsize(out_name))
