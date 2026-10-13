import fitz
doc = fitz.open("../public iron infusion.pdf")
page = doc[0]
pix = page.get_pixmap(dpi=150)
pix.save("original_public.png")
doc.close()
print("Saved original_public.png")
