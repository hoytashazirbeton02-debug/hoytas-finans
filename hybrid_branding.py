"""Load the packaged HOYTAŞ brand image for Tk labels."""
from pathlib import Path
import tkinter as tk


def set_brand_logo(label, max_size=96, *, compound="none"):
    path = Path(__file__).with_name("hoytas_logo.png")
    if not path.is_file():
        return
    source = tk.PhotoImage(master=label, file=str(path))
    longest = max(source.width(), source.height())
    factor = (longest + max_size - 1) // max_size
    label.brand_image = source.subsample(factor)
    label.configure(image=label.brand_image, compound=compound)
