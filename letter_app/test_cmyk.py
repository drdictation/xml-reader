import fitz

pdf_path = "../public iron infusion.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

# Test 1: shape.insert_textbox with color=0 (grayscale)
shape = page.new_shape()
shape.insert_textbox(
    fitz.Rect(1060, 110, 1540, 200),
    "Test Grayscale 0: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=0
)
shape.commit()

# Test 2: shape.insert_textbox with color=(0, 0, 0, 1) (CMYK black)
shape2 = page.new_shape()
shape2.insert_textbox(
    fitz.Rect(1060, 250, 1540, 340),
    "Test CMYK Black: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0, 1)
)
shape2.commit()

# Test 3: shape.insert_textbox with color=(0, 0, 0) (RGB black)
shape3 = page.new_shape()
shape3.insert_textbox(
    fitz.Rect(1060, 360, 1540, 450),
    "Test RGB Black: shape.insert_textbox",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape3.commit()

doc.save("test_cmyk.pdf")
pix = page.get_pixmap(dpi=150)
pix.save("test_cmyk.png")
doc.close()
print("Done")
