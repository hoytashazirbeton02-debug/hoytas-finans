"""Check invoice accounting against isolated v132 databases."""
from pathlib import Path
from types import MethodType, SimpleNamespace
from unittest.mock import patch

import hoytas as h
from hybrid_audit_bridge import pending_batches
from hybrid_inbound import apply_batch, ReplicaConflict
from invoice_matching import invoice_quote, post_invoice, selected_sales, stored_details


def verify():
    assert h.ISOLATED_PROFILE_DIR
    source = Path(h.DB_PATH)
    replica = source.with_name("replica.db")
    competing = source.with_name("competing.db")
    with patch.dict(h.os.environ, {"HOYTAS_HYBRID_CAPTURE": "1"}):
        with h.connect() as db:
            first = db.insert_shared("INSERT INTO cari(tur,unvan,created_at) VALUES(?,?,?)",
                                     ("Müşteri", "Birinci", "2026-11-01")).lastrowid
            second = db.insert_shared("INSERT INTO cari(tur,unvan,created_at) VALUES(?,?,?)",
                                      ("Müşteri", "İkinci", "2026-11-01")).lastrowid
            sales = []
            for buyer, quantity, concrete, pump, transport in (
                    (first,5,5000,500,300), (first,5,5000,0,0), (second,10,13000,0,0)):
                sale = db.insert_shared("""INSERT INTO beton_satis(tarih,cari_id,miktar_m3,birim_fiyat,toplam,odeme_durumu,pompa_ucreti,mikser_nakliye_ucreti)
                    VALUES(?,?,?,?,?,?,?,?)""", ("2026-11-02", buyer, quantity, concrete / quantity, concrete+pump+transport, "Vadeli",pump,transport)).lastrowid
                h.sync_beton_sale_cari_movement(db, sale)
                sales.append(sale)
            try:
                selected_sales(db, first, (sales[0], sales[2]), h.index_beton_movements(db)[0])
            except ValueError:
                pass
            else:
                raise AssertionError("Different-customer sales were accepted.")
        with h.connect() as db:
            baseline = pending_batches(db)
        for path in (replica, competing):
            with h.connect(path):
                pass
            for batch in baseline:
                assert apply_batch(lambda: h.connect(path), *batch) == "applied"
        assert len(baseline) == 1

        label = next(label for label in h.WITHHOLDING_LABELS if label.startswith("601 ·"))
        form = dict(tarih="2026-11-02", belge="TEST001", aciklama="Beton", miktar="10", fiyat="1000",
                    kdv="20", tur="Tevkifatlı", kod=label, tevkifat="9/10", cari="ignorable")
        app = SimpleNamespace(prompt=lambda *_: form, _refresh_selected_module=lambda: None, current_user={})
        app._prompt_for_cari = MethodType(h.HoytasApp._prompt_for_cari, app)
        app._confirm_invoice_posting = MethodType(h.HoytasApp._confirm_invoice_posting, app)
        with patch.object(h.messagebox, "askyesno", side_effect=[False, True]) as approval, \
                patch.object(h.messagebox, "showinfo"), patch.object(h.messagebox, "showerror") as errors:
            h.HoytasApp.add_fatura(app, (first, "Birinci"), sale_ids=sales[:2], post=True)
            with h.connect() as db:
                assert not db.execute("SELECT 1 FROM fatura").fetchone()
            h.HoytasApp.add_fatura(app, (first, "Birinci"), sale_ids=sales[:2], post=True)
            errors.assert_not_called()
            assert "+1.296,00 TL" in approval.call_args.args[1]
            assert "Pompa Hizmeti: 500,00 TL" in approval.call_args.args[1]
            assert "Mikser Nakliye: 300,00 TL" in approval.call_args.args[1]
        with h.connect() as db:
            invoice_id, payable, stored = db.execute("SELECT id,tahsil_edilecek,belge_bilgileri FROM fatura").fetchone()
            assert payable == 12096
            assert stored_details(stored)["tevkifat_islem_kodu"] == "601"
            assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (first,)).fetchone() == (12096,)
            assert db.execute("SELECT fatura_id FROM beton_satis WHERE cari_id=? ORDER BY id", (first,)).fetchall() == [(invoice_id,)] * 2
            batch = pending_batches(db)[1]
        from pypdf import PdfReader
        output = source.with_name("fatura.pdf")
        h.HoytasApp.generate_pdf(app, "Fatura Taslağı PDF", str(output), invoice_id)
        report = "\n".join(page.extract_text() for page in PdfReader(output).pages)
        assert all(label in report for label in ("Hazır Beton", "Pompa Hizmeti", "Mikser Nakliye", "12.096,00")), report
        assert apply_batch(lambda: h.connect(replica), *batch) == "applied"
        with h.connect(replica) as db:
            assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (first,)).fetchone() == (12096,)
        with h.connect(competing) as db:
            fid = db.execute("""INSERT INTO fatura(id,tarih,cari_id,belge_no,miktar_m3,tahsil_edilecek)
                VALUES(999,'2026-11-02',?,'OTHER',5,5000)""", (first,)).lastrowid
            quote = invoice_quote(db, first, fid)
            post_invoice(db, first, fid, quote, selected_sales(db, first, sales[:1], h.index_beton_movements(db)[0]))
        try:
            apply_batch(lambda: h.connect(competing), *batch)
        except ReplicaConflict:
            pass
        else:
            raise AssertionError("Double link was accepted from another replica.")
        with h.connect(competing) as db:
            assert db.execute("SELECT count(*) FROM fatura").fetchone() == (1,)

        with h.connect() as db:
            fid = db.insert_shared("""INSERT INTO fatura(tarih,cari_id,belge_no,miktar_m3,tahsil_edilecek)
                VALUES(?,?,?,?,?)""", ("2026-11-03", second, "OLD", 10, 12000)).lastrowid
            quote = invoice_quote(db, second, fid)
            selected = selected_sales(db, second, sales[2:], h.index_beton_movements(db)[0])
            post_invoice(db, second, fid, quote, selected)
            assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (second,)).fetchone() == (12000,)
            assert db.execute("SELECT tur,tutar FROM hareket WHERE cari_id=? AND kategori='Fatura Farkı'", (second,)).fetchone() == ("Alacak", 1000)
            standalone = db.insert_shared("""INSERT INTO fatura(tarih,cari_id,belge_no,miktar_m3,tahsil_edilecek)
                VALUES(?,?,?,?,?)""", ("2026-11-03", second, "NEW", 1, 200)).lastrowid
            post_invoice(db, second, standalone, invoice_quote(db, second, standalone), ())
            assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (second,)).fetchone() == (12200,)
            assert not db.execute("PRAGMA foreign_key_check").fetchone()


if __name__ == "__main__":
    verify()
