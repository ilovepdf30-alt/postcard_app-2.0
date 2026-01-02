import fitz  # PyMuPDF
from PIL import Image, ImageTk

def render_pdf_page_to_photoimage(pdf_path: str, page_index: int, canvas_w: int, canvas_h: int) -> ImageTk.PhotoImage:
    doc = fitz.open(pdf_path)
    try:
        page = doc.load_page(page_index)
        rect = page.rect
        target_w = max(200, int(canvas_w) - 20)
        target_h = max(200, int(canvas_h) - 20)
        scale = min(target_w / max(1.0, rect.width), target_h / max(1.0, rect.height))
        scale = max(0.6, min(4.0, scale))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((target_w, target_h))
        return ImageTk.PhotoImage(img)
    finally:
        doc.close()