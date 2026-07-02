import fitz
from PIL import Image
import pytesseract

def read_pdf(uploaded_file):
    texto = ""
    pdf_bytes = uploaded_file.read()
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            texto += page.get_text("text") + "\n"
    return texto

def read_image(uploaded_file):
    image = Image.open(uploaded_file)
    return pytesseract.image_to_string(image, lang="spa")

def read_file(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith(".pdf"):
        return read_pdf(uploaded_file)

    if name.endswith((".jpg", ".jpeg", ".png")):
        return read_image(uploaded_file)

    return ""
