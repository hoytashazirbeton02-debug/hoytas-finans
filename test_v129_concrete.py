"""Exercise sale creation, edit and both report paths on an isolated profile."""
from pathlib import Path
from types import SimpleNamespace, MethodType
from unittest.mock import patch

from pypdf import PdfReader
from openpyxl import load_workbook

import hoytas as h


def verify():
    assert h.ISOLATED_PROFILE_DIR
    with h.connect() as db:
        customer = db.execute("""INSERT INTO cari(tur,unvan,created_at)
            VALUES('Müşteri','Rapor Denemesi','2026-11-01')""").lastrowid
        old_sale = db.execute("""INSERT INTO beton_satis(tarih,cari_id,miktar_m3,birim_fiyat,toplam,odeme_durumu)
            VALUES('2026-11-01',?,10,100,1000,'Vadeli')""", (customer,)).lastrowid
        h.sync_beton_sale_cari_movement(db, old_sale)

    def sale_prompt(_, fields):
        keys = [field[0] for field in fields]
        assert keys.index("mikser_nakliye_ucreti") == keys.index("pompa_ucreti") + 1
        return dict(tarih="2026-11-02", cari=f"{customer} - Rapor Denemesi", santiye="",
                    beton_cinsi="C30", m3="30,5", fiyat="100", pompa_ucreti="200",
                    mikser_nakliye_ucreti="350,50", durum="Tahsil Edildi",
                    odeme_yontemi="Nakit", aciklama="Deneme")

    app = SimpleNamespace(prompt=sale_prompt, _resolve_site=lambda _: None,
                          _refresh_selected_module=lambda: None, current_user={})
    for name in ("_prompt_for_cari", "_report_filename", "_safe_export_name"):
        setattr(app, name, MethodType(getattr(h.HoytasApp, name), app))
    h.HoytasApp.add_beton(app, (customer, "Rapor Denemesi"))
    with h.connect() as db:
        sale, fee, total = db.execute("""SELECT id,mikser_nakliye_ucreti,toplam FROM beton_satis
            ORDER BY id DESC LIMIT 1""").fetchone()
        assert (fee, total) == (350.50, 3600.50)
        assert db.execute("SELECT mikser_nakliye_ucreti FROM beton_satis WHERE id=?", (old_sale,)).fetchone() == (0,)

    def edit_prompt(_, fields, initial):
        assert initial["mikser_nakliye_ucreti"] == 350.50
        return dict(initial, mikser_nakliye_ucreti="500")

    app.prompt = edit_prompt
    tree = SimpleNamespace(selection=lambda: ["sale"], item=lambda *_: (sale,))
    h.HoytasApp.edit_selected(app, tree, "beton_satis")
    with h.connect() as db:
        assert db.execute("SELECT mikser_nakliye_ucreti,toplam FROM beton_satis WHERE id=?", (sale,)).fetchone() == (500,3750)
        assert db.execute("""SELECT toplam_borc,toplam_alacak,bakiye FROM cari_bakiye
            WHERE cari_id=?""", (customer,)).fetchone() == (4750,3750,1000)
        assert db.execute("SELECT tutar FROM nakit_kasa").fetchone() == (3750,)

    profile = Path(h.ISOLATED_PROFILE_DIR)
    for title in ("Cari Hareket Dökümü", "Cari Beton Satış Dökümü", "Beton Satışları Dökümü"):
        pdf = profile / f"{title}.pdf"
        h.HoytasApp.generate_pdf(app, title, str(pdf), customer)
        page = PdfReader(pdf).pages[0].extract_text()
        if "Satış" in title:
            assert "40.5 m³" in page and "500.00" in page and "Mikser" in page, page
            assert "Tahsilat" not in page, page
        else:
            assert "Tahsilat" in page and "3,750.00" in page, page

    excel = profile / "Cari.xlsx"
    with patch.object(h.filedialog, "asksaveasfilename", return_value=str(excel)), \
            patch.object(h.messagebox, "showinfo"), patch.object(h.messagebox, "showerror") as errors:
        h.HoytasApp.export_single_cari_excel(app, customer, "Rapor Denemesi")
        errors.assert_not_called()
    workbook = load_workbook(excel)
    try:
        sheet = workbook["Beton Satışları"]
        assert (sheet["F3"].value, sheet["G3"].value) == (500,3750)
    finally:
        workbook.close()


if __name__ == "__main__":
    verify()
