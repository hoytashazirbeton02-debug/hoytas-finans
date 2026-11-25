"""Load the packaged HOYTAŞ brand image for Tk labels."""
from pathlib import Path
import tkinter as tk


def set_brand_logo(label, max_size=96, *, compound="none"):
    path = Path(__file__).with_name("hoytas_logo.png")
    if not path.is_file():
        return
    from PIL import Image, ImageTk

    scale = float(label.tk.call("tk", "scaling")) / (96 / 72)
    pixels = max(1, round(max_size * scale))
    with Image.open(path) as source:
        logo = source.convert("RGBA")
    logo.thumbnail((pixels, pixels), Image.Resampling.LANCZOS)
    label.brand_image = ImageTk.PhotoImage(logo, master=label)
    label.configure(image=label.brand_image, compound=compound)
