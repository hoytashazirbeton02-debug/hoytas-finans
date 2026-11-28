"""Exercise date entry and the customer dropdown in the concrete form."""
import tkinter as tk
from tkinter import ttk

import hoytas as h


def verify():
    root = tk.Tk()
    choices = (h.NO_CUSTOMER_SELECTION, "1 - Öztürk Beton", "2 - Akbaşlar")
    fields = [
        ("tarih", "Tarih (GG.AA.YYYY)", "entry", None),
        ("cari", "Cari ara / seçin", "combo_search", choices),
    ]
    errors = []
    def interact():
        window = next(child for child in root.winfo_children() if isinstance(child, tk.Toplevel))
        body = window.winfo_children()[0]
        try:
            day = next(child for child in body.winfo_children() if isinstance(child, h.DateEntry))
            buyer = next(child for child in body.winfo_children() if isinstance(child, ttk.Combobox))
            for digit in "28112026":
                day.insert("insert", digit)
                day._format_input()
            assert day.get() == "28.11.2026"
            assert buyer.get() == ""
            post = buyer.cget("postcommand")
            buyer.tk.call(post)
            assert tuple(buyer["values"]) == choices
            buyer.insert(0, "ozturk")
            buyer.tk.call(post)
            assert tuple(buyer["values"]) == choices[1:2]
            buyer.set(choices[1])
            buyer.tk.call(post)
            assert tuple(buyer["values"]) == choices
            buttons = body.winfo_children()[-1]
            buttons.winfo_children()[0].invoke()
        except Exception as error:
            errors.append(error)
            window.destroy()
    try:
        root.after(100, interact)
        result = h.HoytasApp.prompt(root, "Yeni Beton Satışı", fields)
        if errors:
            raise errors[0]
        assert result == {"tarih": "2026-11-28", "cari": choices[1]}
    finally:
        root.destroy()


if __name__ == "__main__":
    verify()
