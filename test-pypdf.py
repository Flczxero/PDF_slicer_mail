import os
from pypdf import PdfReader, PdfWriter

main_pdf_path = os.path.join("..", "1-s2.0-S2772443321000027-main.pdf")

reader = PdfReader(main_pdf_path)
writer = PdfWriter()

# Just slice the first page
writer.add_page(reader.pages[0])

with open("test-pypdf.pdf", "wb") as out:
    writer.write(out)

print("Size:", os.path.getsize("test-pypdf.pdf"))
