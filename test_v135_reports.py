"""Verify report routing and PDF content with mixed invoice states."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from pypdf import PdfReader

import hoytas as h
from test_v133_invoice_period import verify as verify_period


def verify():
    with patch.dict(h.os.environ, {"HOYTAS_HYBRID_CAPTURE": "1"}):
        customer, _ = verify_period()
    with h.connect() as db:
        plain = db.execute("INSERT INTO cari(tur,unvan,created_at) VALUES('Müşteri','Peşin','2026-11-01')").lastrowid
        invoice = db.execute("SELECT id FROM fatura WHERE belge_no='DÖNEM-1'").fetchone()[0]
        db.execute("UPDATE beton_satis SET beton_cinsi='C35' WHERE cari_id=? AND fatura_id IS NULL", (customer,))
        sale = db.execute("SELECT id FROM beton_satis WHERE cari_id=? AND fatura_id IS NULL", (customer,)).fetchone()[0]

    app = SimpleNamespace()
    output = Path(h.ISOLATED_PROFILE_DIR) / "report.pdf"

    def read(report, record=None, **kwargs):
        h.HoytasApp.generate_pdf(app, report, str(output), record, **kwargs)
        pdf = PdfReader(output)
        assert pdf.metadata.title == f"HOYTAŞ - {report}"
        return "\n".join(page.extract_text() for page in pdf.pages)

    general, billed, unbilled = h.CARI_STATEMENT_REPORTS
    full = read(general, customer)
    assert "Genel Cari Bakiye" in full and "Faturalı Kalan Borç" in full and "15 m³" in full
    assert "Faturalı Kalan Borç" not in read(general, plain)
    billed_pdf = read(billed, customer)
    assert all(value in billed_pdf for value in ("DÖNEM-1", "DÖNEM-2", "10.000,00", "2.000,00"))
    assert "C35" not in billed_pdf
    unbilled_pdf = read(unbilled, customer)
    assert all(value in unbilled_pdf for value in ("C35", "5.000,00", "03.12.2026"))
    assert "DÖNEM-1" not in unbilled_pdf and "10.000,00" not in unbilled_pdf
    assert "C35" in read("Cari Beton Satış Dökümü", customer, sale_ids=(sale,))
    assert "Cari Hareket Dökümü" in read("Cari Hareket Dökümü", customer)
    assert "DÖNEM-1" in read("Fatura Taslağı PDF", invoice)
    for report in ("Beton Satışları Dökümü", "Gelir-Gider Raporu", "Açık Borç Listesi",
                   "Çek Listesi", "Alınan Çekler Dökümü", "Verilen Çekler Dökümü",
                   *h.CHEQUE_STATUS_REPORTS, "Gelen Alış Faturaları", "Kasa ve Banka Raporu"):
        assert report in read(report), report

    for module in ("Banka İşlemleri", "Nakit Kasa", "Personel", "Maaş Ödemeleri",
                   "Mastarcı Hesabı", "Stok / Depo", "Stok Hareketleri", "Alış Faturaları"):
        content = read(module + " Listesi", display_rows=[["Kayıt"], ["Görünür kayıt"]])
        assert module in content and "Görünür kayıt" in content
        assert "Beton Cinsi" not in content

    app.prompt = Mock(return_value={"rapor": unbilled})
    app.export_single_cari_pdf = Mock()
    h.HoytasApp.choose_cari_statement_pdf(app, customer)
    app.export_single_cari_pdf.assert_called_once_with(customer, unbilled)
    for report, record in ((billed, plain), (unbilled, plain), ("Geçersiz rapor", None)):
        try:
            read(report, record)
        except ValueError:
            continue
        raise AssertionError(f"An empty or unknown report was accepted: {report}")


if __name__ == "__main__":
    verify()
