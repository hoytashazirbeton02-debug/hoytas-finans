"""Regress successive invoice periods with overlapping open receivables."""
from types import MethodType, SimpleNamespace
from unittest.mock import patch

import hoytas as h
from invoice_matching import customer_invoice_totals, sales_in_period, uninvoiced_sales


def verify():
    assert h.ISOLATED_PROFILE_DIR
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
        app.prompt = lambda *_: dict(tarih="2026-12-01", tutar="2000", yontem="Nakit",
                                      banka="", cek_no="", vade="", aciklama="")
        h.HoytasApp.cari_tahsilat(app, context, invoice_id=invoice)
        form.update(belge="DÖNEM-2", tarih="2026-12-02")
        app.prompt = lambda *_: form
        h.HoytasApp.add_fatura(app, context, sale_ids=(december,), post=True)
        add_sale("2026-12-03")
        errors.assert_not_called()
    with h.connect() as db:
        assert db.execute("SELECT tahsil_edilecek-odenen FROM fatura ORDER BY id").fetchall() == [(4000,), (6000,)]
        assert customer_invoice_totals(db, customer) == (1000000, 1, 500000)
        assert [row[1] for row in uninvoiced_sales(db, customer)] == ["2026-12-03"]
        assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (customer,)).fetchone() == (15000,)


if __name__ == "__main__":
    verify()
