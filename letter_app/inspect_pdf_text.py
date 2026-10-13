import fitz

doc = fitz.open("generated_iron/BRIDGETTE_GIBNEY_public_iron_2026-06-14.pdf")
page = doc[0]

# Extract text blocks with details
blocks = page.get_text("dict")["blocks"]
for b in blocks:
    if "lines" in b:
        for l in b["lines"]:
            for s in l["spans"]:
                print(f"Text: {repr(s['text'])}, Font: {s['font']}, Size: {s['size']}, Color: {s['color']}")
doc.close()
