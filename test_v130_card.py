"""Windows Tk smoke test for the revised customer card."""
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from pypdf import PdfReader

import hoytas as h


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def double_click(app, tree, row):
    x, y, width, height = tree.bbox(row, "#1")
    for instant in (100, 200):
        tree.event_generate("<ButtonPress-1>", x=x + width // 2, y=y + height // 2, time=instant)
        tree.event_generate("<ButtonRelease-1>", x=x + width // 2, y=y + height // 2, time=instant + 10)
    app.update()


def verify():
    assert h.ISOLATED_PROFILE_DIR
    with h.connect() as db:
        cid = db.execute("""INSERT INTO cari(tur,unvan,created_at)
            VALUES('Müşteri','Deneme Cari','2026-11-01')""").lastrowid
        for name, quantity in (("Çiftlik", 3), ("Akçabel", 6)):
            site = db.execute("INSERT INTO santiye(ad) VALUES(?)", (name,)).lastrowid
            sale = db.execute("""INSERT INTO beton_satis
                (tarih,cari_id,santiye_id,beton_cinsi,miktar_m3,birim_fiyat,mikser_nakliye_ucreti,toplam,odeme_durumu)
                VALUES ('2026-11-18',?,?, 'C30',?,100,50,?,'Vadeli')""",
                (cid, site, quantity, quantity * 100 + 50)).lastrowid
            h.sync_beton_sale_cari_movement(db, sale)

    with patch.object(h.HoytasApp, "_ensure_admin_and_login",
                      lambda app: setattr(app, "current_user", {"username": "Test", "role": "Yönetici"})), \
            patch.object(h.HoytasApp, "check_for_updates_silent"):
        app = h.HoytasApp()
    errors = []
    app.report_callback_exception = lambda *error: errors.append(error)
    try:
        app.update()
        app.show_cari_hareketleri((cid, "Deneme Cari"), selected_tab=1)
        app.update()
        card = next(widget for widget in app.winfo_children() if isinstance(widget, h.tk.Toplevel)
                    and widget.title().startswith("Cari Kartı"))
        widgets = list(descendants(card))
        notebook = next(widget for widget in widgets if isinstance(widget, h.ttk.Notebook))
        page = card.nametowidget(notebook.select())
        table = next(w for w in descendants(page) if isinstance(w, h.ttk.Treeview))
        picker = next(w for w in descendants(page) if isinstance(w, h.ttk.Combobox))
        assert len(table.get_children()) == 2
        picker.set("Çiftlik")
        picker.event_generate("<<ComboboxSelected>>")
        app.update()
        item, = table.get_children()
        assert table.set(item, "santiye") == "Çiftlik"
        assert any("Toplam 3 m³" in w.cget("text") for w in descendants(page) if isinstance(w, h.ttk.Label))
        output = next(w for w in descendants(page) if isinstance(w, h.ttk.Menubutton) and w.cget("text") == "Çıktı Al")
        assert output.winfo_ismapped()
        menu = output.nametowidget(output.cget("menu"))
        with patch.object(app, "_preview_then_save_pdf") as preview:
            menu.invoke(0)
            ids = preview.call_args.kwargs["sale_ids"]
            assert ids == (int(table.set(item, "id")),)
        pdf = Path(h.ISOLATED_PROFILE_DIR) / "beton.pdf"
        app.generate_pdf("Cari Beton Satış Dökümü", str(pdf), cid, sale_ids=ids)
        text = PdfReader(pdf).pages[0].extract_text()
        assert "3 m³" in text and "Çiftlik" in text and "Akçabel" not in text, text
        excel = Path(h.ISOLATED_PROFILE_DIR) / "beton.xlsx"
        with patch.object(app, "_preview_excel", return_value=False) as preview, \
                patch.object(h.filedialog, "asksaveasfilename") as choose:
            menu.invoke(1)
            preview.assert_called_once()
            choose.assert_not_called()
        with patch.object(app, "_preview_excel", return_value=True), \
                patch.object(h.filedialog, "asksaveasfilename", return_value=str(excel)):
            menu.invoke(1)
        book = load_workbook(excel)
        try:
            assert book.active["B2"].value == "Çiftlik"
            assert book.active["D3"].value == 3
        finally:
            book.close()
        with patch.object(app, "edit_selected") as edit:
            double_click(app, table, item)
            edit.assert_called_once_with(table, "beton_satis")
        refreshed = next(w for w in descendants(card) if isinstance(w, h.ttk.Combobox))
        assert refreshed.get() == "Çiftlik"
        card.destroy()
        app._open_module(app.beton_tab)
        app.refresh_all()
        app.update()
        first = app.beton_tree.get_children()[0]
        with patch.object(app, "edit_selected") as edit:
            double_click(app, app.beton_tree, first)
            edit.assert_called_once_with(app.beton_tree, "beton_satis")
        assert not errors, errors
    finally:
        app.destroy()


if __name__ == "__main__":
    verify()
