"""Regress successive invoice periods with overlapping open receivables."""
from pathlib import Path
import sys
from types import MethodType, SimpleNamespace
from unittest.mock import patch

import hoytas as h
from hybrid_audit_bridge import pending_batches
from hybrid_inbound import apply_batch
from invoice_matching import customer_invoice_totals, sales_in_period, uninvoiced_sales


def verify():
    assert h.ISOLATED_PROFILE_DIR
    source = Path(h.DB_PATH)
    replica = source.with_name("period_replica.db")
    with h.connect(replica):
        pass
    with h.connect() as db:
        customer = db.execute("INSERT INTO cari(tur,unvan,created_at) VALUES('Müşteri','Dönem','2026-11-01')").lastrowid

    def add_sale(day):
        with h.connect() as db:
            sale = db.insert_shared("""INSERT INTO beton_satis(tarih,cari_id,miktar_m3,birim_fiyat,toplam,odeme_durumu)
                VALUES(?,?,5,1000,5000,'Vadeli')""", (day, customer)).lastrowid
            h.sync_beton_sale_cari_movement(db, sale)
        return sale

    november = add_sale("2026-11-30")
    december = add_sale("2026-12-01")
    with h.connect() as db:
        available = uninvoiced_sales(db, customer)
        assert [row[0] for row in sales_in_period(available, "2026-11-01", "2026-11-30")] == [november]
        assert [row[0] for row in sales_in_period(available, "2026-12-01")] == [december]
        assert customer_invoice_totals(db, customer) == (0, 2, 1000000)
    try:
        sales_in_period(available, "2026-12-01", "2026-11-30")
    except ValueError:
        pass
    else:
        raise AssertionError("The reversed date range was accepted.")

    form = dict(tarih="2026-12-01", belge="DÖNEM-1", aciklama="", kdv="20",
                tur="Normal", kod="", tevkifat="0")
    app = SimpleNamespace(prompt=lambda *_: form, _refresh_selected_module=lambda: None, current_user={})
    for name in ("_prompt_for_cari", "_confirm_invoice_posting"):
        setattr(app, name, MethodType(getattr(h.HoytasApp, name), app))
    context = customer, "Dönem"
    with patch.object(h.messagebox, "askyesno", return_value=True), \
            patch.object(h.messagebox, "showinfo"), patch.object(h.messagebox, "showerror") as errors:
        h.HoytasApp.add_fatura(app, context, sale_ids=(november,), post=True)
        with h.connect() as db:
            assert customer_invoice_totals(db, customer) == (600000, 1, 500000)
            invoice = db.execute("SELECT id FROM fatura WHERE belge_no='DÖNEM-1'").fetchone()[0]
        with h.connect() as db:
            batches = pending_batches(db)
        for batch in batches:
            assert apply_batch(lambda: h.connect(replica), *batch) == "applied"
        with h.connect(replica) as db:
            assert customer_invoice_totals(db, customer) == (600000, 1, 500000)
        app.prompt = lambda *_: dict(tarih="2026-12-01", tutar="2000", yontem="Nakit",
                                      banka="", cek_no="", vade="", aciklama="")
        with patch.object(h, "DB_PATH", str(replica)):
            h.HoytasApp.cari_tahsilat(app, context, invoice_id=invoice)
            with h.connect() as db:
                payment, = pending_batches(db)
                assert customer_invoice_totals(db, customer) == (400000, 1, 500000)
        assert apply_batch(h.connect, *payment) == "applied"
        assert apply_batch(h.connect, *payment) == "duplicate"
        form.update(belge="DÖNEM-2", tarih="2026-12-02")
        app.prompt = lambda *_: form
        h.HoytasApp.add_fatura(app, context, sale_ids=(december,), post=True)
        add_sale("2026-12-03")
        errors.assert_not_called()
    with h.connect() as db:
        subsequent = pending_batches(db)[len(batches):]
        assert db.execute("SELECT tahsil_edilecek-odenen FROM fatura ORDER BY id").fetchall() == [(4000,), (6000,)]
        assert customer_invoice_totals(db, customer) == (1000000, 1, 500000)
        assert [row[1] for row in uninvoiced_sales(db, customer)] == ["2026-12-03"]
        assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (customer,)).fetchone() == (15000,)
    for batch in subsequent:
        assert apply_batch(lambda: h.connect(replica), *batch) == "applied"
    with h.connect(replica) as db:
        assert customer_invoice_totals(db, customer) == (1000000, 1, 500000)
    return context


def verify_desktop(context):
    with patch.object(h.HoytasApp, "_ensure_admin_and_login",
                      lambda app: setattr(app, "current_user", {"username": "Test", "role": "Yönetici"})), \
            patch.object(h.HoytasApp, "check_for_updates_silent"):
        app = h.HoytasApp()
    errors = []
    app.report_callback_exception = lambda *error: errors.append(error)

    def click_range():
        window = next(w for w in app.winfo_children() if isinstance(w, h.tk.Toplevel)
                      and w.title() == "Beton Satışını Faturaya Bağla")
        try:
            body, = window.winfo_children()
            dates = next(w for w in body.winfo_children() if isinstance(w, h.ttk.Frame)
                         and sum(isinstance(x, h.ttk.Entry) for x in w.winfo_children()) == 2)
            start, end = [w for w in dates.winfo_children() if isinstance(w, h.ttk.Entry)]
            start.insert(0, "01.12.2026")
            end.insert(0, "02.12.2026")
            buttons = [w for w in dates.winfo_children() if isinstance(w, h.ttk.Button)]
            buttons[0].invoke()
            # Dec 1 was already invoiced, Dec 3 lies outside the selected dates.
            descendants = list(walk(body))
            grid = next(w for w in descendants if isinstance(w, h.ttk.Treeview))
            assert not grid.get_children()
            end.delete(0, "end")
            end.insert(0, "03.12.2026")
            selector = next(w for w in descendants if isinstance(w, h.ttk.Button)
                            and w.cget("text") == "Listelenenleri Seç")
            selector.invoke()
            item, = grid.selection()
            assert grid.set(item, "tarih") == "03.12.2026"
            proceed = next(w for w in descendants if isinstance(w, h.ttk.Button)
                           and w.cget("text") == "Devam Et")
            proceed.invoke()
        finally:
            if window.winfo_exists():
                window.destroy()

    try:
        with patch.object(app, "add_fatura") as draft:
            app.after(80, click_range)
            app.match_concrete_invoice(context)
            assert not errors, errors
            assert draft.call_args.kwargs["post"]
            with h.connect() as db:
                expected, = uninvoiced_sales(db, context[0])
            assert draft.call_args.kwargs["sale_ids"] == (expected[0],)
    finally:
        app.destroy()


def walk(widget):
    for child in widget.winfo_children():
        yield child
        yield from walk(child)


if __name__ == "__main__":
    with patch.dict(h.os.environ, {"HOYTAS_HYBRID_CAPTURE": "1"}):
        context = verify()
    if "--desktop" in sys.argv:
        verify_desktop(context)
