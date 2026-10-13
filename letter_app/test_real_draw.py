import fitz
import sys
import os

pdf_path = "../public iron infusion.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

patient = {
    "name": "BRIDGETTE GIBNEY",
    "dob": "20-4-91",
    "urn": "12345",
    "medicare_number": "3106729603-6",
    "phone": "0401217139",
    "full_address": "123 Fake Street, Melbourne VIC 3000",
    "private_health": "Bupa 987654"
}

from app import build_iron_patient_sticker_text
text = build_iron_patient_sticker_text(patient, is_small=False)

sticker_rect = fitz.Rect(1050, 100, 1550, 450)
rect = fitz.Rect(sticker_rect.x0 + 10, sticker_rect.y0 + 10, sticker_rect.x1 - 10, sticker_rect.y1 - 10)

print("Text to draw:")
print(text)
print("Rect:", rect)

# Let's call shape.insert_textbox directly and check result
shape = page.new_shape()
result = shape.insert_textbox(
    rect,
    text,
    fontsize=20.0,
    fontname="helv",
    color=(0, 0, 0),
    align=0,
    rotate=0
)
print("Insert textbox result:", result)
shape.commit()

doc.save("test_direct_shape.pdf")
pix = page.get_pixmap(dpi=150)
pix.save("test_direct_shape.png")
doc.close()
print("Done")
