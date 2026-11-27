"""Test manual debt reconciliation, receipt and PDF totals against an isolated profile."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

from pypdf import PdfReader
import hoytas as h
from invoice_matching import manual_invoice_debit, manual_invoice_debits, selected_sales, invoice_quote, post_invoice

def verify():
    assert h.ISOLATED_PROFILE_DIR
    with h.connect() as db:
        cid = db.execute("INSERT INTO cari(tur,unvan,created_at) VALUES('Müşteri','Deneme','2026-11-01')").lastrowid
        sale = db.execute("""INSERT INTO beton_satis(tarih,cari_id,miktar_m3,birim_fiyat,toplam,odeme_durumu)
            VALUES('2026-11-02',?,5,1000,5000,'Vadeli')""",(cid,)).lastrowid
        h.sync_beton_sale_cari_movement(db, sale)
        fid = db.execute("""INSERT INTO fatura(tarih,cari_id,belge_no,miktar_m3,tahsil_edilecek)
            VALUES('2026-11-03',?,'DENEME',5,6000)""",(cid,)).lastrowid
        mid = db.execute("""INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar)
            VALUES('2026-11-03',?,'Borç','Fatura','Elle girilen fatura borcu',6000)""",(cid,)).lastrowid
        assert not manual_invoice_debits(db, cid, 5000)  # Sale-origin debit must stay excluded.
        debit, = manual_invoice_debits(db, cid, 6000)
        assert debit[0] == mid
        assert manual_invoice_debit(db, cid, 6000, mid) == debit
        quote = invoice_quote(db, cid, fid)
        sales = selected_sales(db, cid, (sale,), h.index_beton_movements(db)[0])
        post_invoice(db, cid, fid, quote, sales, debit)
        assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (cid,)).fetchone() == (6000,)
        assert not manual_invoice_debits(db, cid, 6000)
        assert db.execute("SELECT tur,tutar FROM hareket WHERE kategori='Fatura Farkı'").fetchone() == ('Alacak',5000)
        assert db.execute("SELECT aciklama FROM hareket WHERE id=?", (mid,)).fetchone()[0].endswith(f"[FATURA:{fid}]")
        new_sale = db.execute("""INSERT INTO beton_satis(tarih,cari_id,miktar_m3,birim_fiyat,toplam,odeme_durumu)
            VALUES('2026-11-04',?,3,1000,3000,'Vadeli')""",(cid,)).lastrowid
        h.sync_beton_sale_cari_movement(db, new_sale)

    app = SimpleNamespace()
    app._collect_sales_invoice = Mock()
    tree = SimpleNamespace(selection=lambda: ("item",), set=lambda item, key: {"tur":"Satış","id":fid}[key])
    h.HoytasApp._collect_selected_invoice(app, tree)
    app._collect_sales_invoice.assert_called_once_with(fid)

    pdf = Path(h.ISOLATED_PROFILE_DIR) / "cari.pdf"
    h.HoytasApp.generate_pdf(app, "Cari Hesap Ekstresi", str(pdf), cid)
    report = "\n".join(page.extract_text() for page in PdfReader(pdf).pages)
    for label in ("Faturalı Kalan Borç", "Faturaya Bağlanmamış Satış Tutarı", "Genel Cari Bakiye", "9,000.00"):
        assert label in report, (label, report)


if __name__ == "__main__":
    verify()
