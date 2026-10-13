import sys
import os

# Add letter_app to path
sys.path.append(os.path.abspath("."))

from app import generate_public_iron_pdf, build_iron_patient_sticker_text, generate_private_iron_pdf
import fitz

patient = {
    "name": "BRIDGETTE GIBNEY",
    "dob": "20-4-91",
    "urn": "12345",
    "medicare_number": "3106729603-6",
    "phone": "0401217139",
    "full_address": "123 Fake Street, Melbourne VIC 3000",
    "private_health": "Bupa 987654"
}

print("Patient sticker text:")
print(repr(build_iron_patient_sticker_text(patient, is_small=False)))

print("Generating Public Iron PDF...")
output_name, output_path = generate_public_iron_pdf(patient)
print("Saved to", output_path)

# Render to png to inspect
doc = fitz.open(output_path)
page = doc[0]
pix = page.get_pixmap(dpi=150)
pix.save("test_app_public_page.png")
doc.close()
print("Saved test_app_public_page.png")

print("Generating Private Iron PDF...")
output_name2, output_path2 = generate_private_iron_pdf(patient)
doc2 = fitz.open(output_path2)
pix2 = doc2[0].get_pixmap(dpi=150)
pix2.save("test_app_private_page1.png")
pix3 = doc2[1].get_pixmap(dpi=150)
pix3.save("test_app_private_page2.png")
doc2.close()
print("Saved private pages to PNG")
