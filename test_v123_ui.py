"""Standalone Tk smoke test. Run with an isolated profile and a display."""
from unittest.mock import patch
import hoytas as h


def verify(scale, width):
    with patch.object(h.HoytasApp, "_ensure_admin_and_login",
                      lambda app: setattr(app, "current_user", {"username": "Test", "role": "Yönetici"})), \
            patch.object(h.HoytasApp, "check_for_updates_silent"):
        app = h.HoytasApp()
    errors = []
    app.report_callback_exception = lambda *error: errors.append(error)
    try:
        app.tk.call("tk", "scaling", scale)
        app.geometry(f"{width}x650")
        app.update()
        nav = app.module_navigation
        groups = nav.winfo_children()
        assert all(widget.winfo_x() + widget.winfo_width() <= nav.winfo_width() for widget in groups)
        cari = groups[1]
        assert isinstance(cari, h.ttk.Menubutton) and cari.cget("text") == "Cari"
        assert len(cari.winfo_children()) == 1
        assert int(cari.cget("width")) == 0
        assert str(cari.cget("direction")) == "below"

        def click(widget, x=3):
            widget.event_generate("<Enter>", x=x, y=3)
            widget.event_generate("<ButtonPress-1>", x=x, y=3)
            widget.event_generate("<ButtonRelease-1>", x=x, y=3)
            app.update()

        menu = cari.nametowidget(cari.cget("menu"))
        posted = []
        app.tk.call("rename", "tk_popup", "hoytas_native_popup")
        app.tk.createcommand("tk_popup", lambda *args: posted.append(args))
        try:
            click(cari)
            click(cari, x=cari.winfo_width() - 4)
            assert len(posted) == 2
            assert all(item[0] == str(menu) for item in posted)
        finally:
            app.tk.deletecommand("tk_popup")
            app.tk.call("rename", "hoytas_native_popup", "tk_popup")
        assert menu.entrycget(1, "label") == "Borç Takip"
        menu.invoke(1)
        app.update()
        assert app.tabs.select() == str(app.borc_takip_tab)
        tree = app.borc_tree
        app.refresh_borc_takip()
        assert len(tree.get_children()) == 2
        tree._search_var.set("ornek")
        assert len(tree.get_children()) == 1
        assert "100.00 TL" in app.borc_total_label.cget("text")
        with patch.object(app, "_preview_then_save_pdf") as preview:
            app._export_debt_list()
            saved = preview.call_args.kwargs["rows"]
            app._export_debt_list(print_dialog=True)
            assert saved == preview.call_args.kwargs["rows"]
            assert preview.call_args.kwargs["print_dialog"] is True
            assert saved[1][0] == "Örnek İnşaat"
            assert saved[-1][3:] == ("100.00", "1 cari")
        tree._search_var.set("")
        assert len(tree.get_children()) == 2
        close = app.module_controls[str(app.borc_takip_tab)][2]
        assert close.winfo_ismapped()
        click(close)
        assert app.tabs.tab(app.borc_takip_tab, "state") == "hidden"
        assert not errors, errors
    finally:
        app.destroy()


if __name__ == "__main__":
    assert h.ISOLATED_PROFILE_DIR
    with h.connect() as db:
        for cid, name in enumerate(("Örnek İnşaat", "Deneme Beton"), 1):
            db.execute("INSERT INTO cari(id,tur,unvan,created_at) VALUES(?,'Müşteri',?,'2026-11-25')", (cid, name))
            db.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,tutar) VALUES('2026-11-25',?,'Borç','Beton',?)",
                       (cid, cid * 100))
    verify(96 / 72, 1024)
    verify(120 / 72, 800)
