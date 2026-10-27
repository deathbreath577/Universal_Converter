import os
import pandas as pd
from PIL import Image, ImageDraw
import win32com.client
import pythoncom

out_dir = os.path.abspath(".")

# 1. Create 2 Sample Images (.jpg and .png) for Multi-Pic -> Single PDF testing
img1 = Image.new("RGB", (600, 400), color=(30, 30, 46))
d1 = ImageDraw.Draw(img1)
d1.text((220, 190), "Test Image 1 (JPG)", fill=(255, 255, 255))
img1.save("test_pic1.jpg")

img2 = Image.new("RGBA", (600, 400), color=(74, 144, 226, 255))
d2 = ImageDraw.Draw(img2)
d2.text((220, 190), "Test Image 2 (PNG)", fill=(255, 255, 255))
img2.save("test_pic2.png")

# 2. Create Sample CSV & Excel (.csv and .xlsx)
df = pd.DataFrame({
    "Subject": ["Signals & Systems", "Data Structures", "Digital Design", "Machine Learning"],
    "Credits": [4, 4, 3, 4],
    "Grade": ["S", "A", "A", "S"]
})
df.to_csv("test_data.csv", index=False)
df.to_excel("test_sheet.xlsx", index=False)

# 3. Create Sample Text File (.txt)
with open("test_notes.txt", "w") as f:
    f.write("Universal Converter Test Notes\n==============================\nThis TXT file should convert cleanly into a PDF.")

# 4. Create Sample Word (.docx) & PowerPoint (.pptx) via MS Office COM
pythoncom.CoInitialize()
try:
    # Generate Word Doc
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    doc.Content.Text = "Sample Word Document (.docx)\n\nTesting Word to PDF conversion via engine.py!"
    doc.SaveAs(os.path.join(out_dir, "test_doc.docx"), FileFormat=16) # 16 = wdFormatDocumentDefault
    doc.Close()
    word.Quit()

    # Generate PowerPoint Deck
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    deck = ppt.Presentations.Add(WithWindow=False)
    slide = deck.Slides.Add(1, 1) # 1 = ppLayoutTitle
    slide.Shapes.Title.TextFrame.TextRange.Text = "Universal Converter Test Slide"
    slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Testing PPTX to PDF conversion!"
    deck.SaveAs(os.path.join(out_dir, "test_slides.pptx"))
    deck.Close()
    ppt.Quit()
except Exception as e:
    print(f"Note on Office generation: {e}")
finally:
    pythoncom.CoUninitialize()

print("All sample files generated in your folder!")