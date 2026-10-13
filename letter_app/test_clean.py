import fitz

pdf_path = "../public iron infusion.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

# Clean the page contents first
page.clean_contents()

# Let's draw text using shape.insert_textbox, without any white rectangle
shape = page.new_shape()
shape.insert_textbox(
    fitz.Rect(1060, 110, 1540, 200),
    "Test Clean: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape.commit()

doc.save("test_clean.pdf")
pix = page.get_pixmap(dpi=150)
pix.save("test_clean.png")
doc.close()
print("Done")
