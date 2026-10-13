import fitz

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

rect = fitz.Rect(1060, 110, 1540, 440)

print("Text to draw:")
print(text)
print("-" * 40)

for size in [20, 18, 16, 14, 12, 10, 8]:
    shape = page.new_shape()
    res = shape.insert_textbox(rect, text, fontsize=size, fontname="helv", color=(0,0,0))
    print(f"Font size {size}: result = {res}")
    # We won't commit, just inspecting the result
doc.close()
