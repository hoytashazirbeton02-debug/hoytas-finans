"""Local invoice accounting and GİB operation references."""
import json
import re
from decimal import Decimal, ROUND_HALF_UP

# UBL-TR 1.42 (March 2026) and KDV Genel Uygulama Tebliği, section I/C.
PARTIAL_WITHHOLDING = {
    "601": ("Yapım işleri ve birlikte verilen proje hizmetleri", 4, "2.1.3.2.1"),
    "602": ("Etüt, danışmanlık ve denetim", 9, "2.1.3.2.2"),
    "603": ("Makine, ekipman ve taşıt bakım ve onarımı", 7, "2.1.3.2.3"),
    "604": ("Yemek servisi", 5, "2.1.3.2.4"),
    "605": ("Organizasyon", 5, "2.1.3.2.4"),
    "606": ("İşgücü temini", 9, "2.1.3.2.5"),
    "607": ("Özel güvenlik", 9, "2.1.3.2.5"),
    "608": ("Yapı denetimi", 9, "2.1.3.2.6"),
    "609": ("Fason tekstil, dikim ve aracılık", 7, "2.1.3.2.7"),
    "610": ("Turistik mağazalara müşteri yönlendirme", 9, "2.1.3.2.8"),
    "611": ("Spor kulüplerinin yayın, reklam ve isim hakları", 9, "2.1.3.2.9"),
    "612": ("Temizlik", 9, "2.1.3.2.10"),
    "613": ("Çevre ve bahçe bakımı", 9, "2.1.3.2.10"),
    "614": ("Servis taşımacılığı", 5, "2.1.3.2.11"),
    "615": ("Baskı ve basım", 7, "2.1.3.2.12"),
    "616": ("Tebliğ kapsamındaki diğer hizmetler", 5, "2.1.3.2.13"),
    "617": ("Hurda metalden üretilen külçe", 7, "2.1.3.3.1"),
    "618": ("Hurda dışı metal külçeleri", 7, "2.1.3.3.1"),
    "619": ("Bakır, çinko ve alüminyum ürünleri", 7, "2.1.3.3.2"),
    "620": ("İstisnadan vazgeçilen hurda ve atık", 7, "2.1.3.3.3"),
    "621": ("Hurda ve atıklardan üretilen hammadde", 9, "2.1.3.3.4"),
    "622": ("Pamuk, tiftik, yün, yapağı ve ham deri", 9, "2.1.3.3.5"),
    "623": ("Ağaç ve orman ürünleri", 5, "2.1.3.3.6"),
    "624": ("Karayolu yük taşımacılığı", 2, "2.1.3.2.11"),
    "625": ("Ticari reklam", 3, "2.1.3.2.15"),
    "626": ("Tebliğ kapsamındaki diğer teslimler", 2, "2.1.3.3.7"),
    "627": ("Demir-çelik ürünleri", 5, "2.1.3.3.8"),
}

WITHHOLDING_LABELS = {
    f"{code} · {name} · {rate}/10": (code, rate, basis)
    for code, (name, rate, basis) in PARTIAL_WITHHOLDING.items()
}
GIB_TEBLIG = ("https://cdn.gib.gov.tr/api/gibportal-file/file/getFileResources?"
              "objectKey=arsiv/fileadmin/user_upload/Tebligler/KDV/kdv_genteb.htm")


def cents(value):
    return int((Decimal(str(value)) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def stored_details(value):
    try:
        parsed = json.loads(value or "{}")
    except (ValueError, TypeError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def selected_sales(db, customer, sale_ids, movement_index):
    """Return a quote only for available sales with individually verified debits."""
    rows = []
    for sale_id in sorted(set(sale_ids)):
        row = db.execute("""SELECT id,tarih,miktar_m3,toplam,birim_fiyat,beton_cinsi,
            pompa_ucreti,mikser_nakliye_ucreti FROM beton_satis
            WHERE id=? AND cari_id=? AND fatura_id IS NULL""", (sale_id, customer)).fetchone()
        if row is None:
            raise ValueError("Satış başka bir cariye veya faturaya bağlı ya da artık mevcut değil.")
        linked = movement_index.get(sale_id, ())
        if len(linked) != 1 or (linked[0][1], linked[0][2], cents(linked[0][3])) != (
                customer, row[1], cents(row[3])):
            raise ValueError(f"Satış #{sale_id} için cari borç bağlantısını önce kontrol edin.")
        rows.append((*row, linked[0][0]))
    return tuple(rows)


def invoice_quote(db, customer, invoice_id):
    row = db.execute("""SELECT tarih,belge_no,tahsil_edilecek,odenen,belge_bilgileri,miktar_m3
        FROM fatura WHERE id=? AND cari_id=? AND durum<>'İptal edildi'""", (invoice_id, customer)).fetchone()
    if row is None or stored_details(row[4]).get("cari_isleme") or db.execute(
            "SELECT 1 FROM beton_satis WHERE fatura_id=? LIMIT 1", (invoice_id,)).fetchone():
        raise ValueError("Fatura bu cariye ait, açık ve henüz cariye işlenmemiş olmalıdır.")
    return row


def post_invoice(db, customer, invoice_id, invoice, sales):
    """Apply only the net difference. Persist all rows in the caller's transaction."""
    day, number, payable, paid, stored, quantity = invoice
    if sales and abs(quantity - sum(row[2] for row in sales)) > 0.000001:
        raise ValueError("Faturanın miktarı seçili beton satışlarının toplamına eşit olmalıdır.")
    sale_debt = sum(cents(row[3]) for row in sales)
    net = cents(payable)
    if net <= 0 or cents(paid) > net:
        raise ValueError("Fatura tahsil edilecek tutarı ve önceki ödemeleri kontrol edin.")
    delta = net - sale_debt
    for sale_id, *_ in sales:
        if db.execute("""UPDATE beton_satis SET fatura_id=?
            WHERE id=? AND cari_id=? AND fatura_id IS NULL""",
            (invoice_id, sale_id, customer)).rowcount != 1:
            raise ValueError("Satışın bağlantısı değişti. İşlemi yeniden açın.")
    move_id = None
    if delta:
        move_id = db.insert_shared("""INSERT INTO hareket
            (tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)""",
            (day, customer, "Borç" if delta > 0 else "Alacak",
             "Fatura Farkı" if sales else "Fatura",
             f"Fatura {number} · {'Satış farkı' if sales else 'Satış faturası'} [FATURA:{invoice_id}]",
             abs(delta) / 100, "Fatura")).lastrowid
    details = stored_details(stored)
    details["cari_isleme"] = {"hareket_id": move_id, "satislar": [row[0] for row in sales],
                               "satis_borcu_kurus": sale_debt, "fark_kurus": delta}
    db.execute("UPDATE fatura SET belge_bilgileri=? WHERE id=?",
               (json.dumps(details, ensure_ascii=False), invoice_id))

def record_is_posted(db, table, record_id):
    if table == "beton_satis":
        return db.execute("SELECT 1 FROM beton_satis WHERE id=? AND fatura_id IS NOT NULL", (record_id,)).fetchone()
    if table == "hareket":
        row = db.execute("SELECT aciklama FROM hareket WHERE id=?", (record_id,)).fetchone()
        return bool(row and re.search(r"\[FATURA:\d+\]$", row[0] or ""))
    if table == "fatura":
        row = db.execute("SELECT belge_bilgileri FROM fatura WHERE id=?", (record_id,)).fetchone()
        return bool(row and stored_details(row[0]).get("cari_isleme"))
    return False



def concrete_invoice_lines(sales):
    """Keep different concrete prices separate and aggregate equal service charges."""
    groups = {}
    for _, _, quantity, _, price, kind, pump, transport, _ in sales:
        concrete = ("Hazır Beton" + (f" · {kind}" if kind else ""), Decimal(str(price)), "m³")
        groups[concrete] = groups.get(concrete, Decimal(0)) + Decimal(str(quantity))
        for name, amount in (("Pompa Hizmeti", pump), ("Mikser Nakliye", transport)):
            if amount:
                service = (name, Decimal(str(amount)), "hizmet")
                groups[service] = groups.get(service, Decimal(0)) + 1
    return [
        {"ad": label, "miktar": str(quantity), "birim": unit,
         "fiyat": str(price), "tutar_kurus": cents(quantity * price)}
        for (label, price, unit), quantity in sorted(groups.items(), key=lambda item: item[0][2] != "m³")
    ]


def uninvoiced_sales(db, customer):
    return db.execute("""SELECT b.id,b.tarih,COALESCE(s.ad,'—'),b.miktar_m3,b.toplam
        FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id
        WHERE b.cari_id=? AND b.fatura_id IS NULL ORDER BY b.tarih,b.id""",
        (customer,)).fetchall()


def sales_in_period(sales, start=None, end=None):
    if start and end and start > end:
        raise ValueError("Başlangıç tarihi bitiş tarihinden sonra olamaz.")
    return [row for row in sales if (not start or row[1] >= start) and (not end or row[1] <= end)]


def customer_invoice_totals(db, customer):
    """Keep unbilled sales distinct from open balances tracked on posted invoices."""
    invoices = db.execute("""SELECT tahsil_edilecek,odenen,belge_bilgileri FROM fatura
        WHERE cari_id=? AND durum<>'İptal edildi' AND tahsil_edilecek>odenen""", (customer,))
    open_invoice = sum(cents(payable) - cents(paid) for payable, paid, info in invoices
                       if stored_details(info).get("cari_isleme"))
    unbilled = uninvoiced_sales(db, customer)
    return open_invoice, len(unbilled), sum(cents(row[4]) for row in unbilled)
