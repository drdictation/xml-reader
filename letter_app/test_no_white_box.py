import fitz

pdf_path = "../public iron infusion.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

# Let's draw text using shape.insert_textbox, at different Y positions, without any white rectangle
shape = page.new_shape()
shape.insert_textbox(
    fitz.Rect(1060, 110, 1540, 200),
    "Test Top: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape.commit()

shape2 = page.new_shape()
shape2.insert_textbox(
    fitz.Rect(1060, 250, 1540, 340),
    "Test Middle: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape2.commit()

shape3 = page.new_shape()
shape3.insert_textbox(
    fitz.Rect(1060, 360, 1540, 450),
    "Test Bottom: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape3.commit()

doc.save("test_no_white_box.pdf")
pix = page.get_pixmap(dpi=150)
pix.save("test_no_white_box.png")
doc.close()
print("Done")
