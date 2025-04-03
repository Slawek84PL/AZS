import fitz
from PIL import Image, ImageTk


def get_resolution(pages):
    cols = 3
    x = 600
    y = 800
    if pages < 3:
        cols = 2
        x = 900
        y = 1080
    return cols, x, y


def build_image_from_page(page, x, y, matrix_arg):
    pix = page.get_pixmap(matrix=fitz.Matrix(matrix_arg, matrix_arg))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.thumbnail((x, y))
    return ImageTk.PhotoImage(img)
