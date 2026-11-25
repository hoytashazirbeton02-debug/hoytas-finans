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
        name, arrow = groups[1].winfo_children()
        assert name.cget("text") == "Cari" and arrow.cget("text") == "▼"

        def click(widget):
            widget.event_generate("<Enter>", x=3, y=3)
            widget.event_generate("<ButtonPress-1>", x=3, y=3)
            widget.event_generate("<ButtonRelease-1>", x=3, y=3)
            app.update()

        menu = arrow.winfo_children()[0]
        # Windows runs a native, blocking menu loop until the user makes a choice.
        # Intercept the popup at that boundary while exercising the actual Tk clicks.
        with patch.object(menu, "tk_popup") as popup:
            click(name)
            popup.assert_not_called()
            click(arrow)
            popup.assert_called_once_with(arrow.winfo_rootx(), arrow.winfo_rooty() + arrow.winfo_height())
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
