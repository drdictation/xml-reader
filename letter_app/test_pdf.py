import fitz
import os

pdf_path = "../public iron infusion.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

# Let's draw a white box first
rect = fitz.Rect(1050, 100, 1550, 450)
shape = page.new_shape()
shape.draw_rect(rect)
shape.finish(color=(0.85, 0.85, 0.85), fill=(1, 1, 1), width=0.8)
shape.commit()

# Test 1: page.insert_textbox with color=(0, 0, 0)
page.insert_textbox(
    fitz.Rect(1060, 110, 1540, 200),
    "Test 1: page.insert_textbox (RGB 0,0,0)",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)

# Test 2: page.insert_textbox with color=0 (grayscale black)
page.insert_textbox(
    fitz.Rect(1060, 210, 1540, 300),
    "Test 2: page.insert_textbox (grayscale 0)",
    fontsize=20,
    fontname="helv",
    color=0
)

# Test 3: page.insert_text with color=(0, 0, 0)
page.insert_text(
    fitz.Point(1060, 330),
    "Test 3: page.insert_text (RGB 0,0,0)",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)

# Test 4: shape.insert_textbox with color=(0, 0, 0)
shape2 = page.new_shape()
shape2.insert_textbox(
    fitz.Rect(1060, 360, 1540, 440),
    "Test 4: shape.insert_textbox (RGB 0,0,0)",
    fontsize=20,
    fontname="helv",
    color=(0, 0, 0)
)
shape2.commit()

# Test 5: Draw a small filled black square just to test draw_rect vs text
shape3 = page.new_shape()
shape3.draw_rect(fitz.Rect(1450, 330, 1480, 360))
shape3.finish(color=(0, 0, 0), fill=(0, 0, 0))
shape3.commit()

output_path = "test_output.pdf"
doc.save(output_path)

# Render to png
pix = page.get_pixmap(dpi=150)
pix.save("test_output_page.png")
doc.close()
print("Saved to test_output.pdf and test_output_page.png")
