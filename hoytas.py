import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date, timedelta
import csv
import os
import sys
import shutil
import subprocess
import tempfile
import zipfile
import re
import hashlib
import json
import urllib.request
import urllib.parse
import urllib.error
import hmac
import secrets
import traceback
from xml.sax.saxutils import escape

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_CURRENT_USER = "Sistem"
TARGET_SCHEMA_VERSION = 50
APP_VERSION = 81
# Veritabanini uygulamanin acildigi konumdan bagimsiz, kalici kullanici klasorunde tut.
USER_DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "HOYTASFinans")
os.makedirs(USER_DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(USER_DATA_DIR, "hoytas.db")
LEGACY_DB_PATH = os.path.join(APP_DIR, "hoytas.db")
# Onceki surumdeki veritabanini ilk acilista bir defa aktar; var olan kayitlari ezme.
if not os.path.exists(DB_PATH) and os.path.isfile(LEGACY_DB_PATH):
    shutil.copy2(LEGACY_DB_PATH, DB_PATH)
BACKUP_DIR = os.path.join(USER_DATA_DIR, "yedekler")
UPDATE_CONFIG_PATH = os.path.join(USER_DATA_DIR, "guncelleme_ayarlari.json")
# Canlı yayın adresi doğrulanmış, anonim HTTPS erişimli bir manifest olmalıdır.
# Önceki OneDrive bağlantısı 401 verdiği için hatalı adres varsayılan olmaktan çıkarıldı.
DEFAULT_UPDATE_MANIFEST_URL = ""
DEFAULT_GITHUB_REPO = "hoytashazirbeton02-debug/hoytas-finans"
# pythonw.exe kullanıldığında konsol olmadığı için beklenmeyen hataları kalıcı loga yaz.
if sys.stderr is None:
    try:
        _startup_log = open(os.path.join(USER_DATA_DIR, "program_hatalari.log"), "a", encoding="utf-8")
        sys.stderr = _startup_log
        sys.stdout = _startup_log
    except OSError:
        pass


def normalize_dropbox_download_url(url):
    """Convert a public Dropbox share URL into its direct-download form."""
    url = str(url or "").strip()
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower()
    if host == "dropbox.com" or host.endswith(".dropbox.com"):
        query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
        query.pop("raw", None)
        query["dl"] = ["1"]
        return urllib.parse.urlunparse((parsed.scheme, parsed.netloc, parsed.path, parsed.params,
                                         urllib.parse.urlencode(query, doseq=True), parsed.fragment))
    return url


def open_pdf_in_viewer(path):
    """PDF'yi Edge'e veya tarayıcıya yönlendirmeden Adobe Acrobat/Reader ile aç."""
    path = os.path.abspath(path)
    if not os.path.isfile(path) or not path.lower().endswith(".pdf"):
        raise RuntimeError("PDF dosyası bulunamadı veya uzantısı .pdf değil.")
    if os.name == "nt":
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pfx86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        local = os.environ.get("LOCALAPPDATA", "")
        candidates = [
            os.path.join(pf, "Adobe", "Acrobat", "Acrobat", "Acrobat.exe"),
            os.path.join(pf, "Adobe", "Acrobat", "Acrobat.exe"),
            os.path.join(pf, "Adobe", "Acrobat DC", "Acrobat", "Acrobat.exe"),
            os.path.join(pf, "Adobe", "Acrobat Reader", "Reader", "AcroRd32.exe"),
            os.path.join(pf, "Adobe", "Acrobat Reader DC", "Reader", "AcroRd32.exe"),
            os.path.join(pfx86, "Adobe", "Acrobat Reader", "Reader", "AcroRd32.exe"),
            os.path.join(pfx86, "Adobe", "Acrobat Reader DC", "Reader", "AcroRd32.exe"),
        ]
        if local:
            candidates.extend([
                os.path.join(local, "Programs", "Adobe", "Acrobat", "Acrobat.exe"),
                os.path.join(local, "Programs", "Adobe", "Acrobat DC", "Acrobat", "Acrobat.exe"),
                os.path.join(local, "Programs", "Adobe", "Acrobat Reader", "Reader", "AcroRd32.exe"),
            ])
        # Adobe'nin kayıt defterindeki App Paths kayıtları bazı kurulumlarda tek güvenilir kaynaktır.
        try:
            import winreg
            keys = (
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
            )
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for key_path in keys:
                    try:
                        with winreg.OpenKey(hive, key_path) as key:
                            exe = winreg.QueryValue(key, None)
                            if exe:
                                candidates.insert(0, exe.strip('"'))
                    except OSError:
                        pass
        except Exception:
            pass
        seen = set()
        for exe in candidates:
            if exe and exe not in seen and os.path.isfile(exe):
                seen.add(exe)
                try:
                    subprocess.Popen([exe, path], close_fds=True)
                    return "Adobe Acrobat / Reader"
                except OSError:
                    continue
        # Bilerek Edge veya varsayılan tarayıcıya düşme: kullanıcıyı yanlış uygulamaya göndermeyelim.
        raise RuntimeError(
            "Adobe Acrobat/Reader uygulaması bulunamadı. PDF dosyası oluşturuldu ancak otomatik açılmadı. "
            "Adobe Acrobat Reader kurulumunu onarın veya kurulu uygulamanın konumunu kontrol edin. "
            "Edge ile açılmayacak."
        )
    raise RuntimeError("PDF otomatik açma bu sistemde desteklenmiyor; oluşturulan PDF dosyasını dosya yöneticisinden açın.")

SCHEMA = """
CREATE TABLE IF NOT EXISTS cari (
 id INTEGER PRIMARY KEY AUTOINCREMENT, tur TEXT NOT NULL, unvan TEXT NOT NULL,
 telefon TEXT, vergi_no TEXT, notlar TEXT, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS hareket (
 id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, cari_id INTEGER,
 tur TEXT NOT NULL, kategori TEXT NOT NULL, aciklama TEXT,
 tutar REAL NOT NULL, odeme_yontemi TEXT,
 FOREIGN KEY(cari_id) REFERENCES cari(id)
);
CREATE TABLE IF NOT EXISTS cek (
 id INTEGER PRIMARY KEY AUTOINCREMENT, yon TEXT NOT NULL, banka TEXT, banka_kodu TEXT, cari_id INTEGER,
 cek_no TEXT, kesideci_lehtar TEXT, vade TEXT NOT NULL, tutar REAL NOT NULL,
 durum TEXT NOT NULL DEFAULT 'Bekliyor', notlar TEXT, FOREIGN KEY(cari_id) REFERENCES cari(id)
);
CREATE TABLE IF NOT EXISTS stok (
 id INTEGER PRIMARY KEY AUTOINCREMENT, kod TEXT, ad TEXT NOT NULL,
 birim TEXT NOT NULL DEFAULT 'Adet', miktar REAL NOT NULL DEFAULT 0,
 birim_fiyat REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS personel (
 id INTEGER PRIMARY KEY AUTOINCREMENT, ad TEXT NOT NULL, gorev TEXT,
 telefon TEXT, maas REAL NOT NULL DEFAULT 0, notlar TEXT
);
CREATE TABLE IF NOT EXISTS beton_satis (
 id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, cari_id INTEGER, santiye_id INTEGER,
 miktar_m3 REAL NOT NULL, birim_fiyat REAL NOT NULL, toplam REAL NOT NULL,
 pompa_ucreti REAL NOT NULL DEFAULT 0,
 beton_cinsi TEXT NOT NULL DEFAULT '', odeme_durumu TEXT NOT NULL, aciklama TEXT,
 FOREIGN KEY(cari_id) REFERENCES cari(id), FOREIGN KEY(santiye_id) REFERENCES santiye(id)
);
CREATE TABLE IF NOT EXISTS santiye (id INTEGER PRIMARY KEY AUTOINCREMENT, ad TEXT NOT NULL, adres TEXT, yetkili TEXT, telefon TEXT, notlar TEXT);
CREATE TABLE IF NOT EXISTS kullanici (id INTEGER PRIMARY KEY AUTOINCREMENT, kullanici_adi TEXT UNIQUE NOT NULL, parola_hash TEXT NOT NULL, salt TEXT NOT NULL, rol TEXT NOT NULL DEFAULT 'Kullanıcı', aktif INTEGER NOT NULL DEFAULT 1, kurtarma_hash TEXT, kurtarma_salt TEXT);
CREATE TABLE IF NOT EXISTS fatura (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, cari_id INTEGER, belge_no TEXT, aciklama TEXT, miktar_m3 REAL NOT NULL DEFAULT 0, birim_fiyat REAL NOT NULL DEFAULT 0, ara_toplam REAL NOT NULL DEFAULT 0, kdv_oran REAL NOT NULL DEFAULT 20, kdv_tutar REAL NOT NULL DEFAULT 0, genel_toplam REAL NOT NULL DEFAULT 0, durum TEXT NOT NULL DEFAULT 'Taslak', FOREIGN KEY(cari_id) REFERENCES cari(id));
CREATE TABLE IF NOT EXISTS gelen_fatura (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, vade TEXT, cari_id INTEGER NOT NULL, belge_no TEXT NOT NULL, aciklama TEXT, ara_toplam REAL NOT NULL DEFAULT 0, kdv_oran REAL NOT NULL DEFAULT 20, kdv_tutar REAL NOT NULL DEFAULT 0, genel_toplam REAL NOT NULL DEFAULT 0, odenen REAL NOT NULL DEFAULT 0, durum TEXT NOT NULL DEFAULT 'Ödenmedi', FOREIGN KEY(cari_id) REFERENCES cari(id));
CREATE TABLE IF NOT EXISTS gelen_fatura_odeme (id INTEGER PRIMARY KEY AUTOINCREMENT, fatura_id INTEGER NOT NULL, tarih TEXT NOT NULL, tutar REAL NOT NULL CHECK(tutar>0), yontem TEXT NOT NULL, aciklama TEXT, FOREIGN KEY(fatura_id) REFERENCES gelen_fatura(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS nakit_kasa (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, yon TEXT NOT NULL, kaynak TEXT NOT NULL, aciklama TEXT, tutar REAL NOT NULL CHECK(tutar > 0), kullanici TEXT);
CREATE TABLE IF NOT EXISTS stok_hareket (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, stok_id INTEGER NOT NULL, yon TEXT NOT NULL, miktar REAL NOT NULL CHECK(miktar>0), aciklama TEXT, FOREIGN KEY(stok_id) REFERENCES stok(id));
CREATE TABLE IF NOT EXISTS maas_odeme (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, personel_id INTEGER NOT NULL, donem TEXT NOT NULL, tutar REAL NOT NULL CHECK(tutar>0), yontem TEXT, aciklama TEXT, FOREIGN KEY(personel_id) REFERENCES personel(id));
CREATE TABLE IF NOT EXISTS banka_hareket (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, banka TEXT NOT NULL, hesap TEXT, yon TEXT NOT NULL, aciklama TEXT, tutar REAL NOT NULL CHECK(tutar>0));
"""

def _verified_sqlite_backup(source_connection, prefix="hoytas_guvenlik"):
    """Create a uniquely named, verified SQLite snapshot (Windows-lock safe)."""
    import time

    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    # tempfile creates an exclusive, process-unique file in the same directory;
    # this prevents two program instances from writing to the same .tmp file.
    fd, temp_target = tempfile.mkstemp(prefix=f"{prefix}_{stamp}_", suffix=".tmp", dir=BACKUP_DIR)
    os.close(fd)
    target = os.path.join(BACKUP_DIR, f"{prefix}_{stamp}_{secrets.token_hex(4)}.db")
    try:
        dest = sqlite3.connect(temp_target, timeout=30)
        try:
            dest.execute("PRAGMA busy_timeout = 30000")
            source_connection.backup(dest, pages=256, sleep=0.1)
            result = dest.execute("PRAGMA integrity_check").fetchone()
            if not result or result[0] != "ok":
                raise RuntimeError(f"Yedek bütünlük kontrolünden geçmedi: {result}")
        finally:
            dest.close()

        if not os.path.isfile(temp_target) or os.path.getsize(temp_target) < 4096:
            raise RuntimeError("Yedek dosyası beklenen boyutta oluşturulamadı.")

        # Antivirus/indexing software may briefly hold a just-written file on
        # Windows. Retry the atomic rename instead of failing immediately.
        last_error = None
        for attempt in range(8):
            try:
                os.replace(temp_target, target)
                return target
            except PermissionError as exc:
                last_error = exc
                time.sleep(0.35 * (attempt + 1))
        raise PermissionError(
            "Windows yedek dosyasını başka bir işlem kullandığı için henüz taşıyamadı. "
            f"Geçici yedek korunuyor: {temp_target}. Program veritabanı değiştirilmedi. "
            f"Ayrıntı: {last_error}"
        )
    except Exception:
        # Do not delete a temp file after a sharing violation: another process
        # may still hold it. A complete verified temp backup is safer retained.
        try:
            if os.path.exists(temp_target) and os.path.getsize(temp_target) < 4096:
                os.remove(temp_target)
        except OSError:
            pass
        raise


def _install_audit_triggers(con):
    """Track business-record changes without ever logging password material."""
    con.execute("""CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        changed_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
        actor TEXT NOT NULL DEFAULT 'Sistem',
        table_name TEXT NOT NULL,
        record_id INTEGER,
        action TEXT NOT NULL CHECK(action IN ('INSERT','UPDATE','DELETE')),
        old_values TEXT,
        new_values TEXT
    )""")
    con.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_time ON audit_log(changed_at DESC, id DESC)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_table ON audit_log(table_name, record_id)")
    tracked = (
        "cari", "hareket", "cek", "stok", "personel", "beton_satis", "santiye",
        "fatura", "gelen_fatura", "gelen_fatura_odeme", "nakit_kasa",
        "stok_hareket", "maas_odeme", "banka_hareket"
    )
    for table in tracked:
        columns = [row[1] for row in con.execute(f"PRAGMA table_info({table})")]
        if not columns or "id" not in columns:
            continue
        old_json = ",".join(f"'{col}', OLD.{col}" for col in columns)
        new_json = ",".join(f"'{col}', NEW.{col}" for col in columns)
        statements = {
            "insert": f"""CREATE TRIGGER IF NOT EXISTS audit_{table}_insert AFTER INSERT ON {table}
                BEGIN INSERT INTO audit_log(actor,table_name,record_id,action,new_values)
                VALUES(hoytas_actor(),'{table}',NEW.id,'INSERT',json_object({new_json})); END""",
            "update": f"""CREATE TRIGGER IF NOT EXISTS audit_{table}_update AFTER UPDATE ON {table}
                BEGIN INSERT INTO audit_log(actor,table_name,record_id,action,old_values,new_values)
                VALUES(hoytas_actor(),'{table}',NEW.id,'UPDATE',json_object({old_json}),json_object({new_json})); END""",
            "delete": f"""CREATE TRIGGER IF NOT EXISTS audit_{table}_delete AFTER DELETE ON {table}
                BEGIN INSERT INTO audit_log(actor,table_name,record_id,action,old_values)
                VALUES(hoytas_actor(),'{table}',OLD.id,'DELETE',json_object({old_json})); END"""
        }
        for sql in statements.values():
            con.execute(sql)


class ClosingSQLiteConnection(sqlite3.Connection):
    """Commit/rollback and always release the database handle on context exit."""
    def __exit__(self, exc_type, exc_value, traceback):
        try:
            return super().__exit__(exc_type, exc_value, traceback)
        finally:
            self.close()


def connect():
    existed = os.path.isfile(DB_PATH) and os.path.getsize(DB_PATH) > 0
    con = sqlite3.connect(DB_PATH, timeout=30, factory=ClosingSQLiteConnection)
    con.row_factory = None
    con.create_function("hoytas_actor", 0, lambda: APP_CURRENT_USER)
    con.execute("PRAGMA busy_timeout = 30000")
    con.execute("PRAGMA foreign_keys = ON")
    con.execute("PRAGMA synchronous = FULL")
    journal_mode = con.execute("PRAGMA journal_mode = WAL").fetchone()[0]
    if str(journal_mode).lower() != "wal":
        con.close()
        raise RuntimeError("Veritabanı güvenli WAL kipine geçirilemedi; mevcut kayıtlar korunarak işlem durduruldu.")
    if con.execute("PRAGMA foreign_keys").fetchone()[0] != 1:
        con.close()
        raise RuntimeError("Veritabanı ilişki kontrolleri etkinleştirilemedi.")
    current_version = con.execute("PRAGMA user_version").fetchone()[0]
    # Her yapısal yükseltmeden önce doğrulanmış ve ayrı isimli yedek alınır.
    if existed and current_version < TARGET_SCHEMA_VERSION:
        _verified_sqlite_backup(con, f"hoytas_v{current_version}_v{TARGET_SCHEMA_VERSION}_oncesi")
    con.executescript(SCHEMA)
    con.execute("""CREATE VIEW IF NOT EXISTS cari_bakiye AS
        SELECT c.id AS cari_id, c.unvan, c.tur,
          COALESCE((SELECT SUM(h.tutar) FROM hareket h WHERE h.cari_id=c.id AND h.tur='Borç'),0) AS toplam_borc,
          COALESCE((SELECT SUM(h.tutar) FROM hareket h WHERE h.cari_id=c.id AND h.tur='Alacak'),0) AS toplam_alacak,
          COALESCE((SELECT SUM(CASE WHEN h.tur='Borç' THEN h.tutar WHEN h.tur='Alacak' THEN -h.tutar ELSE 0 END)
                    FROM hareket h WHERE h.cari_id=c.id),0) AS bakiye
        FROM cari c""")
    # Mevcut veritabanlarını kayıtları silmeden yeni alanlara taşır.
    migrations = {
        "cek": [("banka_kodu", "TEXT"), ("cari_id", "INTEGER")],
        "beton_satis": [("santiye_id", "INTEGER"), ("beton_cinsi", "TEXT NOT NULL DEFAULT ''"), ("pompa_ucreti", "REAL NOT NULL DEFAULT 0")],
        "kullanici": [("kurtarma_hash", "TEXT"), ("kurtarma_salt", "TEXT")],
        "maas_odeme": [("islem_turu", "TEXT NOT NULL DEFAULT 'Maaş'")],
        "fatura": [("fatura_turu", "TEXT NOT NULL DEFAULT 'Normal'"), ("tevkifat_oran", "REAL NOT NULL DEFAULT 0"), ("tevkifat_tutar", "REAL NOT NULL DEFAULT 0"), ("tahsil_edilecek", "REAL NOT NULL DEFAULT 0"), ("odenen", "REAL NOT NULL DEFAULT 0")],
    }
    for table, columns in migrations.items():
        existing = {row[1] for row in con.execute(f"PRAGMA table_info({table})")}
        for col, typ in columns:
            if col not in existing:
                con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")

    con.execute("""CREATE TABLE IF NOT EXISTS fatura_tahsilat (id INTEGER PRIMARY KEY AUTOINCREMENT, fatura_id INTEGER NOT NULL, tarih TEXT NOT NULL, tutar REAL NOT NULL CHECK(tutar > 0), yontem TEXT NOT NULL, aciklama TEXT, FOREIGN KEY(fatura_id) REFERENCES fatura(id) ON DELETE RESTRICT)""")
    con.execute("UPDATE fatura SET tahsil_edilecek=genel_toplam WHERE tahsil_edilecek=0 AND tevkifat_oran=0 AND genel_toplam>0")
    # Eski sürümlerde tarihler GG.AA.YYYY / GG-AA-YYYY olarak kaydedilmiş
    # olabilir. Metin sıralaması bu biçimlerde kronolojik değildir; mevcut
    # kayıtları silmeden ISO YYYY-MM-DD biçimine dönüştür.
    date_columns = {
        "hareket": ("tarih",), "beton_satis": ("tarih",), "cek": ("vade",),
        "fatura": ("tarih",), "fatura_tahsilat": ("tarih",), "gelen_fatura": ("tarih", "vade"),
        "gelen_fatura_odeme": ("tarih",), "nakit_kasa": ("tarih",),
        "stok_hareket": ("tarih",), "maas_odeme": ("tarih",),
        "banka_hareket": ("tarih",),
    }
    for table, cols in date_columns.items():
        existing = {row[1] for row in con.execute(f"PRAGMA table_info({table})")}
        for col in cols:
            if col not in existing:
                continue
            rows = con.execute(f"SELECT id, {col} FROM {table} WHERE {col} IS NOT NULL AND TRIM({col})<>'' AND SUBSTR({col},5,1)<>'-' ").fetchall()
            for row_id, raw_date in rows:
                normalized = None
                for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
                    try:
                        normalized = datetime.strptime(str(raw_date).strip()[:10], fmt).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        pass
                if normalized and normalized != raw_date:
                    con.execute(f"UPDATE {table} SET {col}=? WHERE id=?", (normalized, row_id))
    # İşlem geçmişi ve tetikleyiciler tüm temel kayıt tablolarına kurulur.
    _install_audit_triggers(con)
    # Şema sürümünü ilan etmeden önce veritabanının fiziksel bütünlüğünü doğrula.
    integrity = con.execute("PRAGMA integrity_check").fetchone()
    if not integrity or integrity[0] != "ok":
        con.close()
        raise RuntimeError(f"HOYTAŞ veritabanı bütünlük kontrolünden geçemedi: {integrity}")
    # Sürüm numarası ancak bütün yükseltmeler ve bütünlük kontrolü başarıyla tamamlandıktan sonra yazılır.
    con.execute(f"PRAGMA user_version = {TARGET_SCHEMA_VERSION}")
    return con

def parse_amount(value):
    """Read Turkish/international numeric formats without losing financial precision."""
    import math
    raw = str(value if value is not None else "").replace("₺", "").replace("TL", "").replace("tl", "")
    raw = "".join(raw.split()).replace("\u200b", "")
    if not raw:
        raise ValueError("Tutar boş bırakılamaz.")
    if not re.fullmatch(r"[+-]?[0-9.,]+", raw):
        raise ValueError(f"Geçersiz tutar: {value!r}")
    sign = ""
    if raw[0] in "+-":
        sign, raw = raw[0], raw[1:]
    if not raw or not any(ch.isdigit() for ch in raw):
        raise ValueError("Geçersiz tutar.")
    if "," in raw and "." in raw:
        decimal = "," if raw.rfind(",") > raw.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        whole, fraction = raw.rsplit(decimal, 1)
        if not fraction.isdigit() or not re.fullmatch(r"[0-9]{1,3}(?:" + re.escape(thousands) + r"[0-9]{3})*|[0-9]+", whole):
            raise ValueError(f"Geçersiz tutar: {value!r}")
        normalized = whole.replace(thousands, "") + "." + fraction
    elif raw.count(".") > 1 or raw.count(",") > 1:
        sep = "." if "." in raw else ","
        chunks = raw.split(sep)
        if not chunks[0].isdigit() or any(len(c) != 3 or not c.isdigit() for c in chunks[1:]):
            raise ValueError(f"Geçersiz tutar: {value!r}")
        normalized = "".join(chunks)
    elif "," in raw:
        normalized = raw.replace(",", ".")
    else:
        normalized = raw
    amount = float(sign + normalized)
    if not math.isfinite(amount):
        raise ValueError("Geçersiz tutar.")
    return amount


def safe_float(value, default=0.0):
    """Convert database/UI numeric values including Turkish and spaced formats.

    Examples: 8524130.00, 8 524 130.00, 8.524.130,00 and 8,5.
    Empty optional values become default; malformed values raise a clear ValueError.
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        return default
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    try:
        return parse_amount(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(f"Geçersiz sayısal değer: {value!r}") from exc


def normalize_search_text(value):
    """Normalize Turkish text for forgiving, case-insensitive customer search."""
    import unicodedata
    text = str(value or "").strip().replace("ı", "i").replace("İ", "i").casefold()
    return "".join(ch for ch in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(ch))


def filter_customer_options(options, typed):
    """Return matching suggestions without changing or selecting the user's text."""
    query = normalize_search_text(typed)
    if not query:
        return list(options)
    return [item for item in options if query in normalize_search_text(item)]


def resolve_customer_id(connection, value, allowed_types=("Müşteri", "Diğer")):
    """Resolve a selected 'ID - title' or an exact manually typed customer title."""
    value = str(value or "").strip()
    if not value or value == "(Cari seçilmedi)":
        return None
    match = re.match(r"^(\d+)\s+-\s+", value)
    if match:
        row = connection.execute("SELECT id,tur FROM cari WHERE id=?", (int(match.group(1)),)).fetchone()
        if not row or (allowed_types and row[1] not in allowed_types):
            raise ValueError("Seçilen cari hesap bulunamadı veya bu işlem için uygun değil.")
        return row[0]
    rows = connection.execute("SELECT id,unvan FROM cari WHERE tur IN (%s) ORDER BY id" %
                              ",".join("?" for _ in allowed_types), tuple(allowed_types)).fetchall()
    query = normalize_search_text(value)
    exact = [row[0] for row in rows if normalize_search_text(row[1]) == query]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise ValueError("Aynı ünvanla birden fazla cari var. Lütfen listeden ID'li kaydı seçin.")
    raise ValueError("Yazdığınız cari bulunamadı. Lütfen açılan önerilerden bir cari seçin.")


def linked_beton_sale_id(description):
    """Return the source sale ID for an automatically generated Cari movement."""
    match = re.search(r"\[SATIŞ:(\d+)\]", str(description or ""), flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def concrete_sale_totals(m3_value, price_value="", pump_value="0"):
    """M³ is mandatory; an omitted unit price means a zero-value cari sale."""
    miktar = parse_amount(m3_value)
    fiyat = parse_amount(price_value or "0")
    pompa = parse_amount(pump_value or "0")
    if miktar <= 0 or fiyat < 0 or pompa < 0:
        raise ValueError("Metreküp sıfırdan büyük; fiyat ve pompa ücreti negatif olmamalı.")
    return miktar, fiyat, pompa, round(miktar * fiyat + pompa, 2)


def display_date(value):
    """ISO date from database -> Turkish DD.MM.YYYY for the interface."""
    if not value:
        return value
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except (ValueError, TypeError):
        return value

def format_m3(value):
    """Metreküp miktarını küsuratı kaybetmeden, gereksiz sıfırları göstermeden yaz."""
    try:
        number = safe_float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    return f"{number:,.3f}".rstrip("0").rstrip(".")


def format_beton_aciklama(cinsi, miktar, birim_fiyat, pompa_ucreti=0, toplam=None,
                          santiye="", notlar="", odeme_durumu=""):
    """Beton satışını her ekranda aynı, eksiksiz ve okunaklı dille açıkla."""
    m3 = safe_float(miktar or 0)
    fiyat = safe_float(birim_fiyat or 0)
    pompa = safe_float(pompa_ucreti or 0)
    genel = safe_float(toplam) if toplam is not None else round(m3 * fiyat + pompa, 2)
    parts = [f"{(cinsi or 'Hazır beton').strip()} beton",
             f"{format_m3(m3)} m³",
             f"{fiyat:,.2f} TL/m³"]
    parts.append(f"Pompa: {pompa:,.2f} TL (Var)" if pompa > 0.005 else "Pompa: Yok")
    parts.append(f"Toplam: {genel:,.2f} TL")
    if santiye and str(santiye).strip():
        parts.append(f"Şantiye: {str(santiye).strip()}")
    if odeme_durumu and str(odeme_durumu).strip():
        parts.append(f"Ödeme: {str(odeme_durumu).strip()}")
    if notlar and str(notlar).strip():
        parts.append(f"Not: {str(notlar).strip()}")
    return " • ".join(parts)


def sync_beton_sale_cari_movement(db, sale_id):
    """Keep exactly one sale-originated debit movement in sync with its beton sale.

    Uses the immutable sale ID marker, repairs legacy rows, and never touches
    receipts/payments or unrelated manual cari movements.
    """
    sale = db.execute("""SELECT tarih,cari_id,beton_cinsi,miktar_m3,birim_fiyat,
        COALESCE(pompa_ucreti,0),toplam,COALESCE(aciklama,''),COALESCE(odeme_durumu,'')
        FROM beton_satis WHERE id=?""", (sale_id,)).fetchone()
    if not sale:
        raise ValueError(f"Beton satış kaydı bulunamadı: {sale_id}")
    tarih,cari_id,cinsi,m3,fiyat,pompa,toplam,notlar,durum = sale
    marker = f"%[SATIŞ:{sale_id}]%"
    # Include both stable IDs and pre-v65 legacy descriptions in one repair pass.
    linked = db.execute("""SELECT id FROM hareket WHERE kategori='Satış' AND tur='Borç'
        AND (aciklama LIKE ? OR aciklama LIKE ? OR aciklama LIKE ?) ORDER BY id""",
        (marker, f"Beton satış #{sale_id} - %", f"Beton satış #{sale_id}%")).fetchall()
    if not cari_id:
        for (movement_id,) in linked:
            db.execute("DELETE FROM hareket WHERE id=?", (movement_id,))
        return None
    description = format_beton_aciklama(cinsi,m3,fiyat,pompa,toplam,notlar=notlar,odeme_durumu=durum)
    description = f"{description} [SATIŞ:{sale_id}]"
    if linked:
        keep_id = linked[0][0]
        db.execute("""UPDATE hareket SET tarih=?,cari_id=?,tur='Borç',kategori='Satış',
            aciklama=?,tutar=?,odeme_yontemi='Vadeli' WHERE id=?""",
            (tarih,cari_id,description,toplam,keep_id))
        # Remove only duplicate rows carrying the same explicit legacy/immutable link.
        for (duplicate_id,) in linked[1:]:
            db.execute("DELETE FROM hareket WHERE id=?", (duplicate_id,))
        return keep_id
    cur = db.execute("""INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi)
        VALUES(?,?,?,?,?,?,?)""", (tarih,cari_id,"Borç","Satış",description,toplam,"Vadeli"))
    return cur.lastrowid



def audit_beton_cari_amounts(db, cari_id):
    """Read-only reconciliation: distinguish absent, mismatched and ambiguous links.

    A sale with a linked zero-amount movement is NOT a missing sale.
    Never infer that an unrelated manual debit belongs to a sale.
    """
    sales = db.execute("SELECT id,tarih,toplam FROM beton_satis WHERE cari_id=? ORDER BY tarih,id", (cari_id,)).fetchall()
    movements = db.execute("""SELECT id,cari_id,tarih,tutar,COALESCE(aciklama,'')
        FROM hareket WHERE tur='Borç' AND kategori='Satış' ORDER BY id""").fetchall()
    missing, mismatched, ambiguous = [], [], []
    for sale_id, date, amount in sales:
        import re
        matches = [m for m in movements if
                   f"[SATIŞ:{sale_id}]" in m[4] or
                   re.search(r"Beton satış #" + re.escape(str(sale_id)) + r"(?!\d)", m[4], re.I)]
        if len(matches) == 1:
            m = matches[0]
            if m[1] != cari_id or m[2] != date or abs(safe_float(m[3])-safe_float(amount)) > .009:
                mismatched.append((sale_id, date, safe_float(amount), m[0], safe_float(m[3]), m[1]))
        elif len(matches) > 1:
            ambiguous.append((sale_id, date, safe_float(amount), 'Çoklu bağlantı'))
        else:
            candidates = [m for m in movements if m[1] == cari_id and m[2] == date
                          and '[SATIŞ:' not in m[4] and
                          abs(safe_float(m[3])-safe_float(amount)) < .01]
            if candidates:
                ambiguous.append((sale_id, date, safe_float(amount), 'Kimliği belirsiz eski hareket'))
            else:
                missing.append((sale_id, date, safe_float(amount)))
    return missing, mismatched, ambiguous


def find_missing_beton_movements(db, cari_id):
    """Return unlinked sales; avoid creating a duplicate for plausible legacy entries."""
    sales = db.execute("""SELECT id,tarih,toplam FROM beton_satis
        WHERE cari_id=? ORDER BY tarih,id""", (cari_id,)).fetchall()
    movements = db.execute("""SELECT id,tarih,tutar,COALESCE(aciklama,'') FROM hareket
        WHERE cari_id=? AND tur='Borç' AND kategori='Satış'""", (cari_id,)).fetchall()
    missing, ambiguous = [], []
    for sale_id, date, amount in sales:
        explicit = [m for m in movements if f"[SATIŞ:{sale_id}]" in m[3]
                    or m[3].startswith(f"Beton satış #{sale_id} - ")
                    or m[3] == f"Beton satış #{sale_id}"]
        if explicit:
            continue
        # Existing unmarked sales with same date and amount may already represent
        # this sale. Never guess or duplicate a financial debit.
        possible = [m for m in movements if "[SATIŞ:" not in m[3]
                    and m[1] == date and abs(safe_float(m[2])-safe_float(amount)) < 0.01]
        (ambiguous if possible else missing).append((sale_id, date, amount))
    return missing, ambiguous

def normalize_date(value):
    """Accept DD.MM.YYYY or DD-MM-YYYY; store dates consistently as ISO."""
    value = (value or "").strip()
    if not value:
        return value
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError("Tarih GG.AA.YYYY biçiminde olmalıdır (ör. 26.10.2026).")

class DateTreeview(ttk.Treeview):
    """Format date columns for display without changing database values."""
    def insert(self, parent, index, iid=None, **kw):
        values = kw.get("values")
        if values is not None:
            cols = list(self.cget("columns"))
            cooked = list(values)
            for i, col in enumerate(cols):
                if i < len(cooked) and col.lower() in ("tarih", "vade"):
                    cooked[i] = display_date(cooked[i])
            kw["values"] = cooked
        return super().insert(parent, index, iid=iid, **kw)

def _restore_desktop_shortcut_after_update(program_dir):
    """Restore the stable desktop shortcut after a successful Windows ZIP update."""
    if os.name != "nt":
        return
    script = os.path.join(program_dir, "MASAUSTU_KISAYOLU_OLUSTUR.vbs")
    if not os.path.isfile(script):
        raise RuntimeError("Masaüstü kısayolunu yeniden oluşturacak dosya pakette bulunamadı.")
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    wscript = os.path.join(system_root, "System32", "wscript.exe")
    if not os.path.isfile(wscript):
        wscript = "wscript.exe"
    # /silent argümanı başarı mesajını bastırır; kısayol yine oluşturulur.
    subprocess.Popen([wscript, script, "/silent"], cwd=program_dir, close_fds=True,
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


class HoytasApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HOYTAŞ Hazır Beton | Finans ve İşletme Yönetimi")
        self.geometry("1440x900")
        self.minsize(1180, 720)
        self.configure(bg="#F3F5F4")
        self._ensure_admin_and_login()
        self.style = ttk.Style(self)
        try: self.style.theme_use("clam")
        except tk.TclError: pass
        # Kurumsal HOYTAŞ arayüzü: sakin gri zemin, koyu yeşil ve kontrollü sarı vurgu.
        self.style.configure("TFrame", background="#F1F4F2")
        self.style.configure("TLabel", background="#F1F4F2", foreground="#24352C", font=("Segoe UI", 10))
        self.style.configure("TButton", padding=(12, 8), font=("Segoe UI", 9, "bold"),
                             background="#286044", foreground="#FFFFFF", borderwidth=0, focusthickness=2)
        self.style.map("TButton", background=[("active", "#1D4933"), ("pressed", "#173B2A"), ("disabled", "#D5DCD7")],
                       foreground=[("disabled", "#7C8780")])
        self.style.configure("TEntry", padding=7, fieldbackground="#FFFFFF", foreground="#24352C", bordercolor="#C9D5CD")
        self.style.configure("TCombobox", padding=7, fieldbackground="#FFFFFF", foreground="#24352C", arrowsize=14)
        self.style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], foreground=[("readonly", "#24352C")])
        self.style.configure("TNotebook", background="#E8EEEA", borderwidth=0, tabmargins=(3, 7, 3, 0))
        self.style.configure("TNotebook.Tab", padding=(16, 12), font=("Segoe UI", 10, "bold"),
                             background="#DCE6DF", foreground="#40564A", borderwidth=0)
        self.style.map("TNotebook.Tab", background=[("selected", "#FFFFFF"), ("active", "#C9DCCF")],
                       foreground=[("selected", "#1D5B39"), ("active", "#1F3028")], padding=[("selected", (17, 13))])
        self.style.configure("Treeview", rowheight=34, font=("Segoe UI", 9), background="#FFFFFF",
                             fieldbackground="#FFFFFF", foreground="#24352C", borderwidth=0)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                             background="#DCE8E0", foreground="#214B35", padding=(9, 11), relief="flat")
        self.style.map("Treeview", background=[("selected", "#CDE5D5")], foreground=[("selected", "#173B2A")])
        self.style.configure("TLabelframe", background="#FFFFFF", bordercolor="#D5E0D8", relief="solid", padding=8)
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#286044", font=("Segoe UI", 9, "bold"))
        self.style.configure("TLabelframe", background="#FFFFFF", bordercolor="#D8E0DB", relief="solid")
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#315B46", font=("Segoe UI", 9, "bold"))
        self.style.configure("Metric.TLabelframe", background="#FFFFFF", bordercolor="#D8E0DB", relief="solid")
        self.style.configure("Metric.TLabel", background="#FFFFFF", foreground="#1F3028", font=("Segoe UI", 18, "bold"))
        self._build()
        self.tabs.bind("<<NotebookTabChanged>>", self._refresh_on_tab_change, add="+")
        # Klavye kısayolları: sekmeler, cari arama ve kayıt formları.
        self.bind_all("<Control-f>", self._shortcut_cari_search)
        self.bind_all("<Control-F>", self._shortcut_cari_search)
        self.bind_all("<F5>", self._refresh_from_button)
        self.bind_all("<Control-r>", self._refresh_from_button)
        self.bind_all("<Control-R>", self._refresh_from_button)
        for _i in range(1, 10):
            self.bind_all(f"<Alt-{_i}>", lambda e, i=_i: self._shortcut_tab(i))
        self.refresh_all()
        # Sunucu tanımlıysa açılışta sessizce yeni sürüm kontrolü yap.
        self.after(1800, self.check_for_updates_silent)

    def _refresh_on_tab_change(self, event=None):
        """Sekme değişiminde bütün modülleri aynı veritabanı görüntüsüne getir."""
        if getattr(self, "_refreshing_tab_change", False):
            return
        self._refreshing_tab_change = True
        try:
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Senkronizasyon yapılamadı", f"Sekmeler yenilenirken hata oluştu:\n{exc}", parent=self)
        finally:
            self._refreshing_tab_change = False

    def _refresh_from_button(self, event=None):
        """Tüm modüllerin tablolarını ve özetlerini veritabanından yeniden yükle."""
        try:
            self.refresh_all()
            if hasattr(self, "notice_var"):
                self.notice_var.set("Ekranlar güncellendi.")
            else:
                # Küçük bilgi penceresi açmadan, pencere başlığını kısa süreli değiştirmeden sessiz yenile.
                pass
        except Exception as exc:
            messagebox.showerror("Yenileme yapılamadı", f"Ekranlar yenilenirken hata oluştu:\n{exc}", parent=self)
        return "break"

    def _shortcut_cari_search(self, event=None):
        if hasattr(self, "cari_search"):
            try:
                self.tabs.select(self.cari_tab)
                self.cari_search_entry.focus_set()
                self.cari_search_entry.selection_range(0, "end")
                return "break"
            except tk.TclError:
                pass

    def _shortcut_tab(self, number):
        if hasattr(self, "tabs") and number <= len(self.tabs.tabs()):
            self.tabs.select(number - 1)
            return "break"

    def _hash_password(self, password, salt):
        return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 180000).hex()

    def _ensure_admin_and_login(self):
        with connect() as c:
            count = c.execute("SELECT COUNT(*) FROM kullanici").fetchone()[0]
        if count == 0:
            while True:
                win = tk.Toplevel(self); win.title("HOYTAŞ Hazır Beton - İlk Kurulum"); win.transient(self); win.grab_set(); win.resizable(False, False)
                frm = ttk.Frame(win, padding=22); frm.pack()
                ttk.Label(frm, text="HOYTAŞ Hazır Beton Yönetici Hesabı", font=("Segoe UI", 15, "bold")).grid(row=0,column=0,columnspan=2,pady=(0,12))
                ttk.Label(frm,text="Kullanıcı adı").grid(row=1,column=0,sticky="w",pady=5); user=ttk.Entry(frm,width=28); user.insert(0,"admin"); user.grid(row=1,column=1,pady=5)
                ttk.Label(frm,text="Yeni şifre (en az 8 karakter)").grid(row=2,column=0,sticky="w",pady=5); pw=ttk.Entry(frm,width=28,show="*"); pw.grid(row=2,column=1,pady=5)
                ttk.Label(frm,text="Şifreyi tekrar yazın").grid(row=3,column=0,sticky="w",pady=5); pw2=ttk.Entry(frm,width=28,show="*"); pw2.grid(row=3,column=1,pady=5)
                result={"ok":False}
                def create_user():
                    u=user.get().strip(); a=pw.get(); b=pw2.get()
                    if not u or len(a)<8 or a!=b:
                        messagebox.showerror("Kontrol edin","Kullanıcı adı gerekli; şifre en az 8 karakter olmalı ve iki şifre aynı olmalı.",parent=win); return
                    salt=secrets.token_hex(16)
                    with connect() as c: c.execute("INSERT INTO kullanici(kullanici_adi,parola_hash,salt,rol) VALUES(?,?,?,?)",(u,self._hash_password(a,salt),salt,"Yönetici"))
                    result["ok"]=True; win.destroy()
                ttk.Button(frm,text="Yönetici hesabını oluştur",command=create_user).grid(row=4,column=0,columnspan=2,pady=(14,0))
                self.wait_window(win)
                if not result["ok"]:
                    messagebox.showwarning("Giriş gerekli","Programı kullanmak için yönetici hesabı oluşturulmalı.")
                    self.destroy(); raise SystemExit
                break
        while True:
            win=tk.Toplevel(self); win.title("HOYTAŞ Hazır Beton - Giriş"); win.transient(self); win.grab_set(); win.resizable(False,False)
            frm=ttk.Frame(win,padding=24); frm.pack()
            ttk.Label(frm,text="HOYTAŞ HAZIR BETON",font=("Segoe UI",22,"bold")).grid(row=0,column=0,columnspan=2,pady=(0,3))
            ttk.Label(frm,text="Finans Yönetimi - Güvenli Giriş").grid(row=1,column=0,columnspan=2,pady=(0,14))
            ttk.Label(frm,text="Kullanıcı adı").grid(row=2,column=0,sticky="w",pady=5); user=ttk.Entry(frm,width=28); user.grid(row=2,column=1,pady=5)
            ttk.Label(frm,text="Şifre").grid(row=3,column=0,sticky="w",pady=5); pw=ttk.Entry(frm,width=28,show="*"); pw.grid(row=3,column=1,pady=5)
            result={"ok":False}
            def login():
                global APP_CURRENT_USER
                try:
                    with connect() as c:
                        row=c.execute("SELECT id,kullanici_adi,parola_hash,salt,rol FROM kullanici WHERE kullanici_adi=? AND aktif=1",(user.get().strip(),)).fetchone()
                    if row and hmac.compare_digest(self._hash_password(pw.get(),row[3]),row[2]):
                        self.current_user={"id":row[0],"username":row[1],"role":row[4]}
                        APP_CURRENT_USER = row[1]
                        result["ok"]=True
                        win.destroy()
                    else:
                        messagebox.showerror("Giriş başarısız","Kullanıcı adı veya şifre yanlış.",parent=win)
                except Exception:
                    detail = traceback.format_exc()
                    try:
                        with open(os.path.join(USER_DATA_DIR, "program_hatalari.log"), "a", encoding="utf-8") as log:
                            log.write("\n--- GIRIS ISLEMI HATASI ---\n" + detail + "\n")
                    except OSError:
                        pass
                    messagebox.showerror("Giriş sırasında hata", "Giriş işlemi tamamlanamadı. Verileriniz silinmedi.\nAyrıntı: " + str(detail.splitlines()[-1] if detail.splitlines() else "Bilinmeyen hata"), parent=win)
            ttk.Button(frm,text="Giriş Yap",command=login).grid(row=4,column=0,columnspan=2,pady=(14,0))
            links=ttk.Frame(frm); links.grid(row=5,column=0,columnspan=2,pady=(8,0))
            ttk.Button(links,text="Yeni Üye Ol",command=self.signup_dialog).pack(side="left",padx=4)
            ttk.Button(links,text="Şifremi Unuttum",command=self.forgot_password_dialog).pack(side="left",padx=4)
            pw.bind("<Return>",lambda e:login()); user.focus_set(); self.wait_window(win)
            if result["ok"]: break
            messagebox.showwarning("Giriş gerekli","Giriş yapmadan program açılamaz."); self.destroy(); raise SystemExit

    def signup_dialog(self):
        d=self.prompt("HOYTAŞ Üyelik Oluştur",[("user","Kullanıcı adı","entry",None),("password","Şifre (en az 8 karakter)","entry",None),("recovery","Kurtarma kelimesi (unutmayın)","entry",None)])
        if not d: return
        if not d["user"].strip() or len(d["password"])<8 or len(d["recovery"].strip())<4:
            return messagebox.showerror("Eksik bilgi","Kullanıcı adı, en az 8 karakter şifre ve en az 4 karakter kurtarma kelimesi gerekli.")
        salt=secrets.token_hex(16); rs=secrets.token_hex(16)
        try:
            with connect() as c: c.execute("INSERT INTO kullanici(kullanici_adi,parola_hash,salt,rol,kurtarma_hash,kurtarma_salt) VALUES(?,?,?,?,?,?)",(d["user"].strip(),self._hash_password(d["password"],salt),salt,"Kullanıcı",self._hash_password(d["recovery"].strip().casefold(),rs),rs))
        except sqlite3.IntegrityError: return messagebox.showerror("Kullanıcı mevcut","Bu kullanıcı adı zaten kayıtlı.")
        messagebox.showinfo("Üyelik oluşturuldu","Hesap oluşturuldu. Şimdi kullanıcı adı ve şifrenizle giriş yapabilirsiniz. Kurtarma kelimenizi güvenli yerde saklayın.")

    def forgot_password_dialog(self):
        d=self.prompt("Şifre Kurtarma",[("user","Kullanıcı adı","entry",None),("recovery","Kurtarma kelimesi","entry",None),("password","Yeni şifre (en az 8 karakter)","entry",None)])
        if not d:return
        with connect() as c: row=c.execute("SELECT id,kurtarma_hash,kurtarma_salt FROM kullanici WHERE kullanici_adi=? AND aktif=1",(d["user"].strip(),)).fetchone()
        if not row or not row[1] or not hmac.compare_digest(self._hash_password(d["recovery"].strip().casefold(),row[2]),row[1]):
            return messagebox.showerror("Kurtarma başarısız","Kullanıcı adı veya kurtarma kelimesi yanlış. Kurtarma bilgisi tanımlı değilse yöneticiye başvurun.")
        if len(d["password"])<8:return messagebox.showerror("Geçersiz şifre","Yeni şifre en az 8 karakter olmalı.")
        salt=secrets.token_hex(16)
        with connect() as c:c.execute("UPDATE kullanici SET parola_hash=?,salt=? WHERE id=?",(self._hash_password(d["password"],salt),salt,row[0]))
        messagebox.showinfo("Şifre yenilendi","Yeni şifreniz kaydedildi. Giriş yapabilirsiniz.")

    def change_password_dialog(self):
        d=self.prompt("Şifre Değiştir",[("old","Mevcut şifre","entry",None),("new","Yeni şifre (en az 8 karakter)","entry",None),("recovery","Kurtarma kelimesi (isteğe bağlı)","entry",None)])
        if not d:return
        with connect() as c: row=c.execute("SELECT parola_hash,salt FROM kullanici WHERE id=?",(self.current_user["id"],)).fetchone()
        if not hmac.compare_digest(self._hash_password(d["old"],row[1]),row[0]):return messagebox.showerror("Hata","Mevcut şifre yanlış.")
        if len(d["new"])<8:return messagebox.showerror("Hata","Yeni şifre en az 8 karakter olmalı.")
        salt=secrets.token_hex(16)
        with connect() as c:
            if d["recovery"].strip():
                rs=secrets.token_hex(16); c.execute("UPDATE kullanici SET parola_hash=?,salt=?,kurtarma_hash=?,kurtarma_salt=? WHERE id=?",(self._hash_password(d["new"],salt),salt,self._hash_password(d["recovery"].strip().casefold(),rs),rs,self.current_user["id"]))
            else:c.execute("UPDATE kullanici SET parola_hash=?,salt=? WHERE id=?",(self._hash_password(d["new"],salt),salt,self.current_user["id"]))
        messagebox.showinfo("Başarılı","Şifreniz değiştirildi."+(" Kurtarma kelimesi de güncellendi." if d["recovery"].strip() else ""))

    def _build_santiye(self):
        # Şantiye yönetimi ayrı ana sekme değildir; Beton Satış içinden açılır.
        if getattr(self, "santiye_window", None) is not None and self.santiye_window.winfo_exists():
            self.santiye_window.lift()
            self.refresh_santiye_tree()
            return
        self.santiye_window = tk.Toplevel(self)
        self.santiye_window.title("HOYTAŞ | Şantiye Kayıtları")
        self.santiye_window.geometry("900x440")
        self.santiye_window.transient(self)
        bar = ttk.Frame(self.santiye_window, padding=12); bar.pack(fill="x")
        ttk.Label(bar, text="Şantiye Kayıtları", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="+ Yeni Şantiye Ekle", command=self.add_santiye).pack(side="right")
        self.santiye_tree = self._tree(self.santiye_window, ("id","ad","adres","yetkili","telefon","notlar"),
            ("ID","Şantiye Adı","Adres","Yetkili","Telefon","Notlar"),(50,190,250,140,120,180))
        row=ttk.Frame(self.santiye_window); row.pack(fill="x",padx=12,pady=8)
        ttk.Button(row,text="Seçili Şantiyeyi Düzenle",command=lambda:self.edit_selected(self.santiye_tree,"santiye")).pack(side="left")
        ttk.Button(row, text="Seçili Şantiyeyi Sil", command=lambda:self.delete_selected(self.santiye_tree,"santiye")).pack(side="right")
        self.refresh_santiye_tree()

    def refresh_santiye_tree(self):
        if not hasattr(self, "santiye_tree"):
            return
        try:
            if not self.santiye_tree.winfo_exists(): return
            with connect() as c:
                rows=c.execute("SELECT id,ad,COALESCE(adres,''),COALESCE(yetkili,''),COALESCE(telefon,''),COALESCE(notlar,'') FROM santiye ORDER BY ad").fetchall()
            self.fill_tree(self.santiye_tree, rows)
        except tk.TclError:
            pass

    def _build_gelen_faturalar(self):
        bar=ttk.Frame(self.gelen_fatura_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Gelen Alış Faturaları",font=("Segoe UI",16,"bold")).pack(side="left")
        ttk.Button(bar,text="+ Gelen Fatura Gir",command=self.add_gelen_fatura).pack(side="right")
        ttk.Label(self.gelen_fatura_tab,text="Tedarikçi faturalarını, vadesini ve kısmi/tam ödemelerini takip edin. Bu kayıt resmi e-belge portalına gönderim yapmaz.",foreground="#7a4b00",wraplength=1100).pack(anchor="w",pady=(7,8))
        self.gelen_fatura_tree=self._tree(self.gelen_fatura_tab,("id","tarih","vade","cari","belge","ara","kdv","toplam","odenen","kalan","durum"),("ID","Fatura Tarihi","Vade","Tedarikçi","Fatura No","Ara Toplam","KDV","Genel Toplam","Ödenen","Kalan","Durum"),(42,95,95,180,120,105,85,115,105,105,100))
        row=ttk.Frame(self.gelen_fatura_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Faturaya Ödeme Gir",command=self.odeme_gelen_fatura).pack(side="left")
        ttk.Button(row,text="Seçili Faturayı Düzenle",command=self.edit_gelen_fatura).pack(side="left",padx=8)
        ttk.Button(row,text="Seçili Faturayı Sil",command=lambda:self.delete_selected(self.gelen_fatura_tree,"gelen_fatura")).pack(side="right")

    def add_gelen_fatura(self):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari WHERE tur IN ('Tedarikçi','Diğer') OR tur='Müşteri' ORDER BY unvan").fetchall()
        if not cariler: return messagebox.showinfo("Tedarikçi gerekli","Önce Cari Hesaplar bölümünden tedarikçi ekleyin.",parent=self)
        names=[f"{i} - {n}" for i,n in cariler]
        d=self.prompt("Gelen Alış Faturası",[("tarih","Fatura tarihi (GG.AA.YYYY)","entry",None),("vade","Ödeme vadesi (GG.AA.YYYY)","entry",None),("cari","Tedarikçi","combo",names),("belge","Fatura numarası","entry",None),("aciklama","Açıklama / ürün-hizmet","entry",None),("ara","KDV hariç tutar (TL)","entry",None),("kdv","KDV oranı (%)","combo",["0","1","10","20"])])
        if not d:return
        try:
            tarih=d["tarih"] or date.today().isoformat(); datetime.strptime(tarih,"%Y-%m-%d")
            vade=d["vade"].strip() or None
            if vade: datetime.strptime(vade,"%Y-%m-%d")
            if not d["belge"].strip(): raise ValueError("Fatura numarası zorunludur.")
            ara=parse_amount(d["ara"]); oran=parse_amount(d["kdv"])
            if ara<0 or oran<0 or oran>100: raise ValueError("Tutar veya KDV oranı geçersiz.")
            cid=int(d["cari"].split(" - ",1)[0]); kdv=round(ara*oran/100,2); toplam=round(ara+kdv,2)
            with connect() as c:
                dup=c.execute("SELECT 1 FROM gelen_fatura WHERE cari_id=? AND belge_no=?",(cid,d["belge"].strip())).fetchone()
                if dup: raise ValueError("Bu tedarikçi ve fatura numarası zaten kayıtlı.")
                c.execute("INSERT INTO gelen_fatura(tarih,vade,cari_id,belge_no,aciklama,ara_toplam,kdv_oran,kdv_tutar,genel_toplam,odenen,durum) VALUES(?,?,?,?,?,?,?,?,?,0,'Ödenmedi')",(tarih,vade,cid,d["belge"].strip(),d["aciklama"],ara,oran,kdv,toplam))
            self.refresh_all(); messagebox.showinfo("Kaydedildi",f"Gelen fatura kaydedildi.\\nGenel toplam: {toplam:,.2f} TL\\nÖdeme durumu: Ödenmedi",parent=self)
        except Exception as e: messagebox.showerror("Fatura kaydedilemedi",str(e),parent=self)

    def odeme_gelen_fatura(self):
        sel=self.gelen_fatura_tree.selection()
        if not sel:return messagebox.showinfo("Fatura seçin","Önce ödeme yapılacak faturayı seçin.",parent=self)
        rid=int(self.gelen_fatura_tree.item(sel[0],"values")[0])
        with connect() as c: row=c.execute("SELECT cari_id,belge_no,genel_toplam,odenen FROM gelen_fatura WHERE id=?",(rid,)).fetchone()
        if not row:return
        kalan=round(row[2]-row[3],2)
        if kalan<=0:return messagebox.showinfo("Ödendi","Bu fatura tamamen ödenmiş.",parent=self)
        d=self.prompt("Gelen Fatura Ödemesi",[("tarih","Ödeme tarihi (GG.AA.YYYY)","entry",None),("tutar",f"Ödeme tutarı (kalan {kalan:,.2f} TL)","entry",None),("yontem","Ödeme yöntemi","combo",["Banka","Nakit","Çek","Diğer"]),("aciklama","Açıklama","entry",None)])
        if not d:return
        try:
            dt=d["tarih"] or date.today().isoformat(); datetime.strptime(dt,"%Y-%m-%d"); amount=parse_amount(d["tutar"])
            if amount<=0 or amount>kalan: raise ValueError("Ödeme tutarı sıfırdan büyük ve kalan tutardan fazla olmamalı.")
            newpaid=round(row[3]+amount,2); status="Ödendi" if newpaid>=row[2] else "Kısmen Ödendi"; ref=f"Gelen fatura #{rid} ödeme - {row[1]}"
            with connect() as c:
                c.execute("INSERT INTO gelen_fatura_odeme(fatura_id,tarih,tutar,yontem,aciklama) VALUES(?,?,?,?,?)",(rid,dt,amount,d["yontem"],d["aciklama"]))
                c.execute("UPDATE gelen_fatura SET odenen=?,durum=? WHERE id=?",(newpaid,status,rid))
                c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(dt,row[0],"Gider","Alış Faturası Ödemesi",ref,amount,d["yontem"]))
                if d["yontem"]=="Nakit": c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",(dt,"Para Çıkışı","Gelen Fatura",ref,amount,getattr(self,"current_user",{}).get("username","")))
                elif d["yontem"]=="Banka": c.execute("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",(dt,"Banka","", "Para Çıkışı",ref,amount))
            self.refresh_all(); messagebox.showinfo("Ödeme kaydedildi",f"Ödeme: {amount:,.2f} TL\\nKalan: {round(kalan-amount,2):,.2f} TL\\nDurum: {status}",parent=self)
        except Exception as e:messagebox.showerror("Ödeme kaydedilemedi",str(e),parent=self)

    def edit_gelen_fatura(self):
        sel=self.gelen_fatura_tree.selection()
        if not sel:return messagebox.showinfo("Fatura seçin","Düzenlemek için fatura seçin.",parent=self)
        rid=int(self.gelen_fatura_tree.item(sel[0],"values")[0])
        with connect() as c: old=c.execute("SELECT tarih,vade,cari_id,belge_no,aciklama,ara_toplam,kdv_oran,odenen FROM gelen_fatura WHERE id=?",(rid,)).fetchone(); cariler=c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        if not old:return
        if old[7]>0:return messagebox.showwarning("Düzenleme kısıtlı","Ödemesi bulunan faturanın tutar/tedarikçi bilgisi değiştirilemez. Önce ödeme kayıtlarını kontrol edin.",parent=self)
        names=[f"{i} - {n}" for i,n in cariler]; init={"tarih":old[0],"vade":old[1] or "","cari":next((n for n in names if n.startswith(f"{old[2]} - ")),""),"belge":old[3],"aciklama":old[4] or "","ara":str(old[5]),"kdv":str(old[6])}
        d=self.prompt("Gelen Faturayı Düzenle",[("tarih","Fatura tarihi","entry",None),("vade","Vade (boş olabilir)","entry",None),("cari","Tedarikçi","combo",names),("belge","Fatura numarası","entry",None),("aciklama","Açıklama","entry",None),("ara","KDV hariç tutar","entry",None),("kdv","KDV oranı (%)","combo",["0","1","10","20"])],init)
        if not d:return
        try:
            datetime.strptime(d["tarih"],"%Y-%m-%d"); vade=d["vade"].strip() or None
            if vade:datetime.strptime(vade,"%Y-%m-%d")
            ara=parse_amount(d["ara"]); oran=parse_amount(d["kdv"]); kdv=round(ara*oran/100,2); total=round(ara+kdv,2); cid=int(d["cari"].split(" - ",1)[0])
            with connect() as c:c.execute("UPDATE gelen_fatura SET tarih=?,vade=?,cari_id=?,belge_no=?,aciklama=?,ara_toplam=?,kdv_oran=?,kdv_tutar=?,genel_toplam=?,durum='Ödenmedi' WHERE id=?",(d["tarih"],vade,cid,d["belge"],d["aciklama"],ara,oran,kdv,total,rid))
            self.refresh_all()
        except Exception as e:messagebox.showerror("Düzenleme başarısız",str(e),parent=self)

    def _build_fatura(self):
        bar=ttk.Frame(self.fatura_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Fatura / Belge Taslakları",font=("Segoe UI",16,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Fatura Taslağı",command=self.add_fatura).pack(side="right")
        ttk.Label(self.fatura_tab,text="Bu bölüm tutar ve KDV hesaplamalı yerel fatura taslağı oluşturur. EDM bağlantısı yapılana kadar resmi e-Fatura göndermez.",foreground="#7a4b00",wraplength=1000).pack(anchor="w",pady=(8,0))
        self.fatura_tree=self._tree(self.fatura_tab,("id","tarih","cari","belge","miktar","ara","kdv","toplam","tevkifat","net","odenen","kalan","durum"),("ID","Tarih","Cari","Belge No","m³","Ara Toplam TL","KDV TL","Genel Toplam TL","Tevkifat TL","Net TL","Tahsil TL","Kalan TL","Durum"),(45,100,200,120,70,120,100,120,110,110,110,110,125))
        row=ttk.Frame(self.fatura_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Fatura Tahsilatı Al",command=self.tahsilat_fatura).pack(side="left",padx=(0,8))
        ttk.Button(row,text="Seçili Taslağı Düzenle",command=lambda:self.edit_selected(self.fatura_tree,"fatura")).pack(side="left")
        ttk.Button(row,text="Seçili Taslağı Sil",command=lambda:self.delete_selected(self.fatura_tree,"fatura")).pack(side="right")

    def _build(self):
        header = tk.Frame(self, bg="#1F3028", height=78)
        header.pack(fill="x")
        header.pack_propagate(False)
        brand = tk.Frame(header, bg="#1F3028")
        brand.pack(side="left", fill="y", padx=(22, 10))
        tk.Label(brand, text="HOYTAŞ", bg="#1F3028", fg="#E7BD37",
                 font=("Segoe UI", 23, "bold")).pack(anchor="w", pady=(12, 0))
        tk.Label(brand, text="HAZIR BETON", bg="#1F3028", fg="#FFFFFF",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 9))
        tk.Frame(header, bg="#DDB62E", width=3).pack(side="left", fill="y", pady=16, padx=(0, 14))
        tk.Label(header, text="Finans ve İşletme Yönetimi", bg="#1F3028", fg="#F3F5F4",
                 font=("Segoe UI", 13, "bold")).pack(side="left", pady=20)
        who = getattr(self, "current_user", {})
        user_text = f"KULLANICI  {who.get('username', '—')}   |   {who.get('role', '')}"
        tk.Label(header, text=user_text, bg="#1F3028", fg="#D3DDD6",
                 font=("Segoe UI", 9)).pack(side="right", padx=24)
        tk.Frame(self, bg="#DDB62E", height=3).pack(fill="x")
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=18, pady=(14, 10))
        self.dashboard = ttk.Frame(self.tabs, padding=18)
        self.cari_tab = ttk.Frame(self.tabs, padding=12)
        self.borc_takip_tab = ttk.Frame(self.tabs, padding=12)
        self.alisveris_cari_tab = ttk.Frame(self.tabs, padding=12)
        self.hareket_tab = ttk.Frame(self.tabs, padding=12)
        self.cek_tab = ttk.Frame(self.tabs, padding=12)
        self.stok_tab = ttk.Frame(self.tabs, padding=12)
        self.personel_tab = ttk.Frame(self.tabs, padding=12)
        self.beton_tab = ttk.Frame(self.tabs, padding=12)
        self.fatura_tab = ttk.Frame(self.tabs, padding=12)
        self.gelen_fatura_tab = ttk.Frame(self.tabs, padding=12)
        self.kasa_tab = ttk.Frame(self.tabs, padding=12)
        self.stok_hareket_tab = ttk.Frame(self.tabs, padding=12)
        self.maas_tab = ttk.Frame(self.tabs, padding=12)
        self.banka_tab = ttk.Frame(self.tabs, padding=12)
        # Sekmelerde kısaltma yerine okunaklı iki satırlı tam adlar kullan.
        tab_items = [
            (self.dashboard, "Genel\nBakış"), (self.cari_tab, "Cari\nHesaplar"),
            (self.borc_takip_tab, "Borç\nTakip"), (self.alisveris_cari_tab, "Tedarikçi\nCari"),
            (self.hareket_tab, "Gelir /\nGider"), (self.cek_tab, "Çek\nİşlemleri"),
            (self.stok_tab, "Stok /\nDepo"), (self.personel_tab, "Personel\nİşlemleri"),
            (self.beton_tab, "Beton\nSatış"), (self.fatura_tab, "Fatura /\nBelge"),
            (self.gelen_fatura_tab, "Alış\nFaturaları"), (self.kasa_tab, "Nakit\nKasa"),
            (self.stok_hareket_tab, "Stok\nHareketleri"), (self.maas_tab, "Maaş\nÖdemeleri"),
            (self.banka_tab, "Banka\nİşlemleri"),
        ]
        for frame, title in tab_items:
            self.tabs.add(frame, text=title)
        # Her sekmenin kendine ait işlem rehberi; genel açıklama yerine bağlama özgü bilgi.
        ttk.Label(self.cari_tab, text="Cari bakiyeleri ve faturalar ayrı izlenir. Bir carinin toplam borcunu ve fatura tahsilatını ilgili kayıt üzerinden kontrol edin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.beton_tab, text="Beton cinsi, m³, şantiye, pompa bedeli ve tahsilat durumunu burada takip edin. Satış açıklamasına teslim/şantiye bilgisini yazın.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.fatura_tab, text="Normal ve tevkifatlı fatura taslaklarında m³, KDV, tevkifat ve net tahsilat ayrı hesaplanır. Resmî e-Fatura gönderimi EDM üzerinden yapılır.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.cek_tab, text="Alınan ve verilen çekleri banka, cari, vade ve durum bilgileriyle takip edin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.stok_tab, text="Malzeme, stok miktarı ve birim maliyetlerini takip edin; stok hareketlerini ayrı sekmede kaydedin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.personel_tab, text="Personel bilgileri bu alandadır. Maaş ve nakit avans hareketleri Maaş Ödemeleri sekmesinde izlenir.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.kasa_tab, text="Nakit giriş ve çıkışlarını kaynak ve işlem açıklamasıyla kaydedin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.hareket_tab, text="Gelir, gider, cari borç ve alacak işlemlerini açıklamasıyla kaydedin; ilgili cariyi seçin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.gelen_fatura_tab, text="Tedarikçi faturalarını, vade tarihlerini, yapılan ödemeleri ve kalan tutarları izleyin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.banka_tab, text="Banka giriş/çıkışlarını hesap ve açıklama bilgileriyle takip edin.", foreground="#315B46", wraplength=1200).pack(anchor="w", pady=(0, 8))
        self._build_dashboard()
        self._build_cari()
        self._build_borc_takip()
        self._build_alisveris_cari()
        self._build_hareket()
        self._build_cek()
        self._build_stok()
        self._build_personel()
        self._build_beton()
        self._build_fatura()
        self._build_gelen_faturalar()
        self._build_kasa()
        self._build_extra_tabs()
        bottom = ttk.Frame(self, padding=(18, 4, 18, 12))
        bottom.pack(fill="x")
        bottom_top = ttk.Frame(bottom); bottom_top.pack(fill="x", pady=(0, 5))
        bottom_bottom = ttk.Frame(bottom); bottom_bottom.pack(fill="x")
        for parent, specs in [
            (bottom_top, [("⟳ Yenile (F5)", self._refresh_from_button), ("Excel Raporu", self.export_excel),
                          ("CSV Dışa Aktar", self.export_csv), ("Şimdi Yedekle", self.backup),
                          ("Veritabanı Kontrolü", self.database_health_check), ("PDF Raporları", self.pdf_menu)]),
            (bottom_bottom, [("Şifre Değiştir", self.change_password_dialog), ("Veritabanı Konumu", self.show_db_path),
                             ("Güncelleme URL Ayarları", self.configure_update_url),
                             ("Güncellemeleri Denetle", self.check_for_updates), ("ZIP ile Güncelle", self.install_update_package)])
        ]:
            for label, command in specs:
                ttk.Button(parent, text=label, command=command).pack(side="left", padx=(0, 7))
        if getattr(self,"current_user",{}).get("role")=="Yönetici":
            ttk.Button(bottom_bottom,text="Kullanıcı Yönetimi",command=self.manage_users).pack(side="right")
            ttk.Button(bottom_bottom,text="İşlem Geçmişi",command=self.show_audit_log).pack(side="right", padx=(0,8))

    def configure_update_url(self):
        """Otomatik güncelleme manifest URL'sini program içinden düzenle."""
        try:
            os.makedirs(USER_DATA_DIR, exist_ok=True)
            config = {}
            if os.path.isfile(UPDATE_CONFIG_PATH):
                with open(UPDATE_CONFIG_PATH, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        config = loaded
            current_url = str(config.get("manifest_url", "")).strip()
        except (OSError, ValueError):
            config = {}
            current_url = ""

        win = tk.Toplevel(self)
        win.title("HOYTAŞ - Otomatik Güncelleme URL Ayarı")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Yayın sunucusu (manifest) HTTPS adresi", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Buraya ZIP adresi değil, manifest.json dosyasının doğrudan HTTPS adresi yazılır.", wraplength=540).pack(anchor="w", pady=(5, 10))
        url_var = tk.StringVar(value=current_url)
        entry = ttk.Entry(frame, textvariable=url_var, width=82)
        entry.pack(fill="x", pady=(0, 12))
        entry.focus_set()
        entry.selection_range(0, "end")
        status = ttk.Label(frame, text="Adres kaydedildiğinde program bu adres üzerinden sürüm kontrolü yapar.", foreground="#52645A", wraplength=540)
        status.pack(anchor="w", pady=(0, 12))

        def save_url():
            url = url_var.get().strip()
            parsed = urllib.parse.urlparse(url)
            if url and (parsed.scheme != "https" or not parsed.netloc):
                messagebox.showwarning("Geçersiz adres", "Manifest adresi https:// ile başlamalı ve geçerli bir alan adı içermelidir.", parent=win)
                return
            config["github_repo"] = str(config.get("github_repo", ""))
            config["manifest_url"] = url
            config["channel"] = str(config.get("channel", "stable"))
            try:
                os.makedirs(USER_DATA_DIR, exist_ok=True)
                fd, temp_path = tempfile.mkstemp(prefix="guncelleme_ayarlari_", suffix=".tmp", dir=USER_DATA_DIR)
                try:
                    with os.fdopen(fd, "w", encoding="utf-8") as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
                        f.flush()
                        os.fsync(f.fileno())
                    os.replace(temp_path, UPDATE_CONFIG_PATH)
                finally:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
            except OSError as exc:
                messagebox.showerror("Ayar kaydedilemedi", str(exc), parent=win)
                return
            messagebox.showinfo("HOYTAŞ", "Güncelleme URL adresi kaydedildi.", parent=win)
            win.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Kaydet", command=save_url).pack(side="right")
        ttk.Button(buttons, text="Vazgeç", command=win.destroy).pack(side="right", padx=(0, 8))
        win.bind("<Return>", lambda _event: save_url())
        win.bind("<Escape>", lambda _event: win.destroy())
        win.update_idletasks()
        win.geometry(f"{max(600, win.winfo_reqwidth())}x{win.winfo_reqheight()}")

    def _update_config(self):
        """Eski GitHub ayarını geriye dönük uyumluluk için oku."""
        if not os.path.isfile(UPDATE_CONFIG_PATH):
            os.makedirs(USER_DATA_DIR, exist_ok=True)
            with open(UPDATE_CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({"github_repo": DEFAULT_GITHUB_REPO, "manifest_url": DEFAULT_UPDATE_MANIFEST_URL, "channel": "stable"}, f, ensure_ascii=False, indent=2)
        with open(UPDATE_CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
        repo = str(config.get("github_repo", "")).strip() or DEFAULT_GITHUB_REPO
        if re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo):
            return repo
        return None

    def check_for_updates_silent(self):
        """Açılışta ağ hatalarını kullanıcıyı rahatsız etmeden ele al."""
        try:
            self.check_for_updates(silent=True)
        except Exception as exc:
            try:
                with open(os.path.join(USER_DATA_DIR, "guncelleme.log"), "a", encoding="utf-8") as log:
                    log.write(f"{datetime.now().isoformat(timespec='seconds')} | {exc}\n")
            except OSError:
                pass

    def _fetch_release_with_retry(self, request):
        """Geçici bağlantı kopmalarında sınırlı yeniden deneme."""
        import time
        last_error = None
        for attempt in range(3):
            try:
                with urllib.request.urlopen(request, timeout=20) as response:
                    return json.loads(response.read(2 * 1024 * 1024).decode("utf-8"))
            except (urllib.error.URLError, ConnectionError, TimeoutError, OSError) as exc:
                last_error = exc
                if attempt < 2:
                    time.sleep(attempt + 1)
        raise last_error

    def check_for_updates(self, silent=False):
        """Önce bağımsız HTTPS manifestini, gerekirse eski GitHub kaynağını denetle."""
        try:
            repo = self._update_config()
            try:
                with open(UPDATE_CONFIG_PATH, "r", encoding="utf-8") as config_file:
                    config = json.load(config_file)
            except (OSError, ValueError):
                config = {}
            manifest_url = str(config.get("manifest_url", "")).strip() or DEFAULT_UPDATE_MANIFEST_URL
            if manifest_url:
                self._check_manifest_update(manifest_url, silent=silent)
                return
            if not repo:
                if not silent:
                    messagebox.showinfo(
                        "HOYTAŞ otomatik güncelleme",
                        "Güncelleme istemcisi hazır; ancak yayın adresi henüz tanımlanmadı.\n\n"
                        "GitHub kullanmadan da çalışması için güvenli HTTPS manifest adresi tanımlanabilir.\n"
                        f"Ayar dosyası: {UPDATE_CONFIG_PATH}", parent=self)
                return
            api = f"https://api.github.com/repos/{repo}/releases/latest"
            request = urllib.request.Request(api, headers={"Accept": "application/vnd.github+json", "User-Agent": "HOYTAS-Updater/1.0"})
            release = self._fetch_release_with_retry(request)
            tag = str(release.get("tag_name", ""))
            match = re.search(r"(\d+)", tag)
            if not match:
                raise RuntimeError("Yayın etiketinde sürüm numarası bulunamadı.")
            version = int(match.group(1))
            if version <= APP_VERSION:
                if not silent:
                    messagebox.showinfo("HOYTAŞ güncel", f"Kullandığınız sürüm v{APP_VERSION}. Daha yeni bir sürüm bulunamadı.", parent=self)
                return
            assets = release.get("assets", [])
            package_asset = next((a for a in assets if str(a.get("name", "")).lower().endswith(".zip") and "hoytas" in str(a.get("name", "")).lower()), None)
            if not package_asset:
                raise RuntimeError("Yayında HOYTAŞ ZIP güncelleme paketi bulunamadı.")
            digest = str(package_asset.get("digest") or "")
            if not digest.startswith("sha256:"):
                checksum_asset = next((a for a in assets if str(a.get("name", "")).lower() in ("sha256sums.txt", "hoytas.sha256")), None)
                if checksum_asset:
                    req = urllib.request.Request(checksum_asset["browser_download_url"], headers={"User-Agent": "HOYTAS-Updater/1.0"})
                    with urllib.request.urlopen(req, timeout=15) as response:
                        checksum_text = response.read(65536).decode("utf-8", "replace")
                    for line in checksum_text.splitlines():
                        parts = line.strip().split()
                        if len(parts) >= 2 and parts[-1].lstrip("*") == package_asset.get("name"):
                            digest = "sha256:" + parts[0].lower()
                            break
            if not digest.startswith("sha256:") or not re.fullmatch(r"sha256:[0-9a-fA-F]{64}", digest):
                raise RuntimeError("Yayın paketi SHA-256 doğrulama bilgisi taşımıyor; güvenlik için indirme durduruldu.")
            if not silent and not messagebox.askyesno("HOYTAŞ güncellemesi bulundu", f"Yeni sürüm v{version} bulundu.\n\nŞimdi indirip güvenli şekilde kurayım mı?", parent=self):
                return
            fd, package_path = tempfile.mkstemp(prefix=f"hoytas_v{version}_", suffix=".zip", dir=USER_DATA_DIR)
            os.close(fd)
            try:
                req = urllib.request.Request(package_asset["browser_download_url"], headers={"User-Agent": "HOYTAS-Updater/1.0"})
                hasher = hashlib.sha256()
                total = 0
                with urllib.request.urlopen(req, timeout=60) as response, open(package_path, "wb") as target:
                    while True:
                        chunk = response.read(1024 * 1024)
                        if not chunk:
                            break
                        total += len(chunk)
                        if total > 500 * 1024 * 1024:
                            raise RuntimeError("Güncelleme paketi 500 MB sınırını aşıyor.")
                        hasher.update(chunk)
                        target.write(chunk)
                if hasher.hexdigest().lower() != digest.split(":", 1)[1].lower():
                    raise RuntimeError("İndirilen dosyanın SHA-256 doğrulaması başarısız. Dosya kurulmadı.")
                self.install_update_package(package=package_path, source="internet")
            finally:
                try:
                    os.remove(package_path)
                except OSError:
                    pass
        except urllib.error.URLError as exc:
            if silent:
                self._log_update_error(exc)
            else:
                messagebox.showerror("Güncelleme kontrol edilemedi", f"İnternet veya yayın sunucusu bağlantısı kurulamadı:\n{exc}", parent=self)
        except Exception as exc:
            if silent:
                self._log_update_error(exc)
            else:
                messagebox.showerror("Güncelleme kontrol edilemedi", str(exc), parent=self)

    def _check_manifest_update(self, manifest_url, silent=False):
        """GitHub gerektirmeyen HTTPS JSON manifest kaynağından güvenli güncelleme al."""
        manifest_url = normalize_dropbox_download_url(manifest_url)
        parsed = urllib.parse.urlparse(manifest_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise RuntimeError("Güncelleme manifest adresi HTTPS olmalıdır.")
        request = urllib.request.Request(manifest_url, headers={"User-Agent": "HOYTAS-Updater/1.0", "Cache-Control": "no-cache"})
        with urllib.request.urlopen(request, timeout=20) as response:
            manifest_bytes = response.read(256 * 1024 + 1)
        if len(manifest_bytes) > 256 * 1024:
            raise RuntimeError("Güncelleme manifest dosyası izin verilen boyutu aşıyor.")
        manifest = json.loads(manifest_bytes.decode("utf-8"))
        if not isinstance(manifest, dict):
            raise RuntimeError("Güncelleme manifest biçimi geçersiz.")
        raw_version = manifest.get("version")
        if isinstance(raw_version, bool):
            raise RuntimeError("Manifest sürüm numarası geçersiz.")
        try:
            version = int(raw_version)
        except (TypeError, ValueError):
            raise RuntimeError("Manifest sürüm numarası geçersiz.")
        if version < 1 or version > 9999:
            raise RuntimeError("Manifest sürüm numarası izin verilen aralıkta değil.")
        package_url = normalize_dropbox_download_url(str(manifest.get("package_url", "")).strip())
        digest = str(manifest.get("sha256", "")).strip().lower()
        package_parsed = urllib.parse.urlparse(package_url)
        # Kaynağın güvenliğini sürüm karşılaştırmasından önce doğrula.
        if package_parsed.scheme != "https" or not package_parsed.netloc or package_parsed.username or package_parsed.password:
            raise RuntimeError("Paket adresi kullanıcı bilgisi içermeyen geçerli bir HTTPS adresi olmalıdır.")
        if not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise RuntimeError("Manifest içinde geçerli SHA-256 bilgisi bulunamadı.")
        if version <= APP_VERSION:
            if not silent:
                messagebox.showinfo("HOYTAŞ güncel", f"Kullandığınız sürüm v{APP_VERSION}. Daha yeni sürüm yok.", parent=self)
            return
        # Sessiz kontrol yalnızca "güncel" ve ağ hatası bildirimlerini gizler;
        # yeni sürüm bulunduğunda kullanıcı onayı olmadan kurulum başlatılmaz.
        if not messagebox.askyesno("HOYTAŞ güncellemesi bulundu", f"Yeni sürüm v{version} bulundu.\n\nŞimdi indirip güvenli şekilde kurayım mı?", parent=self):
            return
        os.makedirs(USER_DATA_DIR, exist_ok=True)
        fd, package_path = tempfile.mkstemp(prefix=f"hoytas_v{version}_", suffix=".zip", dir=USER_DATA_DIR)
        os.close(fd)
        try:
            req = urllib.request.Request(package_url, headers={"User-Agent": "HOYTAS-Updater/1.0"})
            hasher = hashlib.sha256()
            total = 0
            with urllib.request.urlopen(req, timeout=60) as response, open(package_path, "wb") as target:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > 500 * 1024 * 1024:
                        raise RuntimeError("Güncelleme paketi 500 MB sınırını aşıyor.")
                    hasher.update(chunk)
                    target.write(chunk)
            if total == 0:
                raise RuntimeError("Yayın sunucusu boş güncelleme paketi gönderdi.")
            if not hmac.compare_digest(hasher.hexdigest(), digest):
                raise RuntimeError("İndirilen dosyanın SHA-256 doğrulaması başarısız; kurulum iptal edildi.")
            self.install_update_package(package=package_path, source="internet")
        finally:
            try:
                os.remove(package_path)
            except OSError:
                pass

    def _log_update_error(self, exc):
        try:
            os.makedirs(USER_DATA_DIR, exist_ok=True)
            with open(os.path.join(USER_DATA_DIR, "guncelleme.log"), "a", encoding="utf-8") as log:
                log.write(f"{datetime.now().isoformat(timespec='seconds')} | {exc}\n")
        except OSError:
            pass

    def _activate_installed_version(self, destination, version):
        """Doğrulanmış kurulu sürümü veri kaybı olmadan etkinleştir."""
        pointer = os.path.join(USER_DATA_DIR, "aktif_program_yolu.txt")
        pointer_tmp = pointer + ".tmp"
        with open(pointer_tmp, "w", encoding="utf-16") as f:
            f.write(destination)
        os.replace(pointer_tmp, pointer)
        try:
            _restore_desktop_shortcut_after_update(destination)
        except Exception as exc:
            self._log_update_error(exc)
        messagebox.showinfo(
            "HOYTAŞ sürümü hazır",
            f"v{version} zaten kurulu. Mevcut dosyalar korunarak aktif sürüm bağlantısı yenilendi.\n"
            "Programı kapatıp masaüstü kısayolundan tekrar açın.",
            parent=self
        )

    def install_update_package(self, package=None, source="manual"):
        """ZIP güncellemesini ayrı sürüm klasörüne güvenli biçimde kur; mevcut sürüme dokunma."""
        if not package:
            package = filedialog.askopenfilename(
                parent=self, title="HOYTAŞ güncelleme ZIP dosyasını seçin",
                filetypes=[("HOYTAŞ güncelleme paketi", "*.zip"), ("Tüm dosyalar", "*.*")]
            )
        if not package:
            return
        try:
            with zipfile.ZipFile(package, "r") as zf:
                bad = zf.testzip()
                if bad:
                    raise RuntimeError(f"ZIP doğrulaması başarısız: {bad}")
                members = zf.infolist()
                if not members or len(members) > 5000:
                    raise RuntimeError("Güncelleme paketi boş veya beklenenden fazla dosya içeriyor.")
                total_uncompressed = 0
                for info in members:
                    name = info.filename.replace("\\", "/")
                    clean_name = name.rstrip("/")
                    parts = clean_name.split("/") if clean_name else []
                    if name.startswith("/") or re.match(r"^[A-Za-z]:", name) or any(part in ("..", "") or ":" in part for part in parts):
                        raise RuntimeError("Güvenlik nedeniyle ZIP içindeki geçersiz dosya yolu reddedildi.")
                    # Unix symlink dosyaları Windows dışında ZIP yol sınırını aşabilir.
                    unix_mode = (info.external_attr >> 16) & 0xFFFF
                    if (unix_mode & 0o170000) == 0o120000:
                        raise RuntimeError("Güvenlik: ZIP içindeki sembolik bağlantı reddedildi.")
                    if info.file_size > 150 * 1024 * 1024:
                        raise RuntimeError("ZIP içinde izin verilen boyutu aşan dosya var.")
                    total_uncompressed += info.file_size
                    if total_uncompressed > 500 * 1024 * 1024:
                        raise RuntimeError("Güncelleme paketi açıldığında 500 MB sınırını aşıyor.")
                py_candidates = [i.filename for i in members if i.filename.replace("\\", "/").endswith("/hoytas.py") or i.filename == "hoytas.py"]
                if len(py_candidates) != 1:
                    raise RuntimeError("Paketin içinde tek bir ana hoytas.py dosyası bulunamadı.")
                main_member = py_candidates[0].replace("\\", "/")
                prefix = main_member.rsplit("/", 1)[0] if "/" in main_member else ""
                change_files = [i.filename for i in members if re.search(r"DEGISIKLIKLER_v(\d+)\.txt$", i.filename, re.I)]
                versions = [int(re.search(r"_v(\d+)\.txt$", n, re.I).group(1)) for n in change_files]
                if not versions:
                    raise RuntimeError("Paket sürüm numarası bulunamadı (DEGISIKLIKLER_vNN.txt eksik).")
                version = max(versions)
                if version <= APP_VERSION:
                    raise RuntimeError(f"Seçilen paket v{version}. Bu program v{APP_VERSION}; daha yeni bir sürüm seçin.")

                root = USER_DATA_DIR
                versions_dir = os.path.join(root, "surumler")
                os.makedirs(versions_dir, exist_ok=True)
                destination = os.path.join(versions_dir, f"v{version}")
                if os.path.exists(destination):
                    # Önceki kurulum tamamlanmışsa aynı paketi yeniden açma.
                    existing_main = os.path.join(destination, "hoytas.py")
                    existing_launcher = os.path.join(destination, "Baslat.vbs")
                    if not (os.path.isfile(existing_main) and os.path.isfile(existing_launcher)):
                        raise RuntimeError(
                            f"v{version} klasörü eksik veya yarım kurulmuş görünüyor. "
                            "Kayıpları önlemek için üzerine yazılmadı; klasörü silmeyin."
                        )
                    with open(existing_main, "r", encoding="utf-8") as installed_file:
                        installed_source = installed_file.read(4096)
                    if not re.search(r"^APP_VERSION\s*=\s*" + str(version) + r"\s*$",
                                     installed_source, re.MULTILINE):
                        raise RuntimeError(f"v{version} klasörünün içeriği sürümle uyuşmuyor; üzerine yazılmadı.")
                    self._activate_installed_version(destination, version)
                    return
                temp_dir = tempfile.mkdtemp(prefix=f"hoytas_v{version}_", dir=versions_dir)
                try:
                    zf.extractall(temp_dir)
                    source_dir = os.path.join(temp_dir, prefix) if prefix else temp_dir
                    if not os.path.isfile(os.path.join(source_dir, "hoytas.py")):
                        raise RuntimeError("ZIP açıldı ancak programın ana dosyası bulunamadı.")
                    if not os.path.isfile(os.path.join(source_dir, "Baslat.vbs")):
                        raise RuntimeError("Paketin doğrudan açılış dosyası (Baslat.vbs) eksik.")
                    # Önce veritabanının SQLite online backup'ını al; canlı dosyayı kopyalama.
                    os.makedirs(BACKUP_DIR, exist_ok=True)
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_path = os.path.join(BACKUP_DIR, f"guncelleme_oncesi_v{version}_{stamp}.db")
                    with sqlite3.connect(DB_PATH, timeout=30) as source_db:
                        with sqlite3.connect(backup_path) as backup_db:
                            source_db.backup(backup_db)
                    with open(backup_path, "rb") as check_file:
                        if not check_file.read(16).startswith(b"SQLite format 3"):
                            raise RuntimeError("Güncelleme öncesi veritabanı yedeği doğrulanamadı.")
                    shutil.copytree(source_dir, destination)
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)

            pointer = os.path.join(USER_DATA_DIR, "aktif_program_yolu.txt")
            # Windows Script Host VBS okuyucusu UTF-16LE bekliyor.
            pointer_tmp = pointer + ".tmp"
            with open(pointer_tmp, "w", encoding="utf-16") as f:
                f.write(destination)
            os.replace(pointer_tmp, pointer)
            shortcut_note = "Masaüstü kısayolu güncel sürüme bağlandı."
            try:
                _restore_desktop_shortcut_after_update(destination)
            except Exception as shortcut_exc:
                shortcut_note = f"Masaüstü kısayolu otomatik oluşturulamadı: {shortcut_exc}"
                self._log_update_error(shortcut_exc)
            messagebox.showinfo(
                "HOYTAŞ güncellemesi hazır",
                f"v{version} ayrı klasöre kuruldu.\n\nVeritabanı yedeği alındı:\n{backup_path}\n\n"
                f"{shortcut_note}\nYeni sürümü açmak için programı kapatıp masaüstündeki HOYTAŞ kısayoluna çift tıklayın.\n"
                "Eski sürüm ve kayıtlar silinmedi.", parent=self
            )
        except Exception as exc:
            messagebox.showerror("Güncelleme yapılamadı", str(exc), parent=self)

    def _build_dashboard(self):
        ttk.Label(self.dashboard, text="Genel Bakış", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(self.dashboard, text="Kayıtlarınız bu bilgisayardaki yerel veritabanında saklanır.",
                  foreground="#596775").pack(anchor="w", pady=(4, 18))
        ttk.Label(self.dashboard,
                  text="Cari: müşteri ve bakiye  •  Beton Satış: m³ ve şantiye  •  Borç Takip: kalan alacak  •  F5: tüm listeleri yenile",
                  foreground="#315B46", font=("Segoe UI", 10, "bold"), wraplength=1080).pack(anchor="w", pady=(0, 12))
        self.metrics = ttk.Frame(self.dashboard)
        self.metrics.pack(fill="x", pady=4)
        self.metric_labels = {}
        for key, label in [("cari","Cari kayıt"),("gelir","Toplam gelir"),("gider","Toplam gider"),
                           ("cek","Bekleyen çek"),("stok","Stok kalemi"),("acik_borc","Açık cari borcu")]:
            card = ttk.LabelFrame(self.metrics, text=label, padding=16, style="Metric.TLabelframe")
            card.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=3)
            val = ttk.Label(card, text="—", style="Metric.TLabel")
            val.pack(anchor="w")
            self.metric_labels[key] = val
        self.debt_alert_label = ttk.Label(self.dashboard, text="", foreground="#A32626", font=("Segoe UI", 10, "bold"), wraplength=980, justify="left")
        self.debt_alert_label.pack(anchor="w", pady=(12, 4))
        ttk.Label(self.dashboard, text="Başlangıç notu", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(18, 6))
        ttk.Label(self.dashboard, text="Beton satışında şantiye seçimi, cari hesap arama, çeklerin cariyle ilişkilendirilmesi, kullanıcı girişi ve fatura taslağı hesaplamaları bu sürümde bulunur. EDM üzerinden resmi e-belge gönderimi entegrasyon erişimi sağlandığında yapılacaktır.",
                  wraplength=970, justify="left").pack(anchor="w")

    def _tree(self, parent, columns, headings, widths=None):
        searchbar = ttk.Frame(parent)
        searchbar.pack(fill="x", pady=(6, 0))
        ttk.Label(searchbar, text="Listede Ara:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(searchbar, textvariable=search_var, width=32)
        search_entry.pack(side="left", padx=(6, 4))
        ttk.Button(searchbar, text="Temizle", command=lambda v=search_var: v.set("")).pack(side="left")
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, pady=(4, 0))
        tree = DateTreeview(frame, columns=columns, show="headings")
        for i, (col, head) in enumerate(zip(columns, headings)):
            tree.heading(col, text=head)
            tree.column(col, width=(widths[i] if widths else 130), anchor="w")
        vs = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        tree.tag_configure("even", background="#FFFFFF")
        tree.tag_configure("odd", background="#F5F8F6")
        tree.tag_configure("debt_unpaid", background="#FCE4E4", foreground="#9B1C1C")
        tree.tag_configure("debt_partial", background="#FFF0D6", foreground="#8A4B08")
        tree.tag_configure("debt_closed", background="#E5F4E8", foreground="#216438")
        tree._search_var = search_var
        search_var.trace_add("write", lambda *_args, w=tree: self._apply_tree_search(w))
        tree.bind("<Double-1>", lambda event, w=tree: self._double_click_edit(w))
        # Sağ tık bağlam menüsü; sol tık ve çift tık davranışları korunur.
        tree.bind("<Button-3>", lambda event, w=tree: self._tree_context_menu(event, w))
        tree.bind("<Control-Button-1>", lambda event, w=tree: self._tree_context_menu(event, w))
        return tree

    def _tree_context_menu(self, event, tree):
        """Tablolarda sağ tıkla görüntüle/düzenle/sil işlemlerini aç."""
        row = tree.identify_row(event.y)
        if row:
            tree.selection_set(row)
            tree.focus(row)
        else:
            return "break"
        mapping={getattr(self,n,None):t for n,t in [
            ("cari_tree","cari"),("hareket_tree","hareket"),("cek_tree","cek"),
            ("stok_tree","stok"),("personel_tree","personel"),("beton_tree","beton_satis"),
            ("fatura_tree","fatura"),("kasa_tree","nakit_kasa"),("santiye_tree","santiye"),
            ("stok_hareket_tree","stok_hareket"),("maas_tree","maas_odeme"),("banka_tree","banka_hareket"),
            ("gelen_fatura_tree","gelen_fatura") ]}
        table=mapping.get(tree)
        menu=tk.Menu(self, tearoff=0)
        if tree is getattr(self,"cari_tree",None):
            menu.add_command(label="Cari hareketlerini görüntüle", command=self.show_cari_hareketleri)
            menu.add_separator()
        if table:
            edit_command = self.edit_gelen_fatura if table == "gelen_fatura" else (lambda:self.edit_selected(tree,table))
            menu.add_command(label="Düzenle", command=edit_command)
            menu.add_command(label="Sil", command=lambda:self.delete_selected(tree,table))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def show_cari_hareketleri(self):
        """Seçili carinin borç/alacak hareketlerini ve bağlı satış/çeklerini göster."""
        sel=self.cari_tree.selection()
        if not sel:
            return messagebox.showinfo("Cari seçin","Hareketlerini görmek için önce bir cari seçin.",parent=self)
        vals=self.cari_tree.item(sel[0],"values")
        cari_id=int(vals[0]); unvan=str(vals[2])
        win=tk.Toplevel(self); win.title(f"Cari Hareketleri - {unvan}"); win.geometry("1100x650"); win.minsize(850,450)
        win.transient(self)
        top=ttk.Frame(win,padding=12); top.pack(fill="x")
        ttk.Label(top,text=unvan,font=("Segoe UI",15,"bold")).pack(side="left")
        with connect() as c:
            totals=c.execute("SELECT toplam_borc,toplam_alacak FROM cari_bakiye WHERE cari_id=?",(cari_id,)).fetchone() or (0,0)
            moves=c.execute("SELECT h.id,h.tarih,h.tur,h.kategori,COALESCE(h.aciklama,''),COALESCE(h.odeme_yontemi,''),h.tutar FROM hareket h WHERE h.cari_id=? ORDER BY h.tarih ASC,h.id ASC",(cari_id,)).fetchall()
            sales=c.execute("SELECT b.id,b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(b.aciklama,''),COALESCE(s.ad,'') FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC",(cari_id,)).fetchall()
            checks=c.execute("SELECT id,yon,COALESCE(banka,''),COALESCE(cek_no,''),vade,tutar,durum FROM cek WHERE cari_id=? ORDER BY vade ASC,id ASC",(cari_id,)).fetchall()
        balance=totals[0]-totals[1]
        total_m3 = sum(safe_float(r[3] or 0) for r in sales)
        sales_total_tl = sum(safe_float(r[6] or 0) for r in sales)
        ttk.Label(top,text=f"Cari Borç: {totals[0]:,.2f} TL    Alacak: {totals[1]:,.2f} TL    Bakiye: {balance:,.2f} TL    |    Beton: {format_m3(total_m3)} m³ / {sales_total_tl:,.2f} TL",font=("Segoe UI",10,"bold")).pack(side="right")
        nb=ttk.Notebook(win); nb.pack(fill="both",expand=True,padx=12,pady=(0,12))
        mf=ttk.Frame(nb,padding=6); nb.add(mf,text="Cari Hareketleri")
        mt=DateTreeview(mf,columns=("id","tarih","tur","kategori","aciklama","odeme","borc","alacak","bakiye"),show="headings")
        heads=("ID","Tarih","Tür","Kategori","Açıklama","Ödeme","Borç TL","Alacak TL","Bakiye TL")
        widths=(45,95,85,120,270,100,100,100,110)
        for col,head,width in zip(mt.cget("columns"),heads,widths): mt.heading(col,text=head); mt.column(col,width=width,anchor="w")
        mvs=ttk.Scrollbar(mf,orient="vertical",command=mt.yview); mt.configure(yscrollcommand=mvs.set); mt.pack(side="left",fill="both",expand=True); mvs.pack(side="right",fill="y")
        bal=0
        sales_by_id = {int(r[0]): r for r in sales}
        import re
        for r in moves:
            debt=r[6] if r[2]=="Borç" else 0
            credit=r[6] if r[2]=="Alacak" else 0
            bal += debt-credit
            description = r[4]
            match = re.search(r"(?:\[SATIŞ:(\d+)\]|Beton satış #?(\d+))", description, re.IGNORECASE)
            if match and r[2] == "Borç" and int(match.group(1) or match.group(2)) in sales_by_id:
                sale = sales_by_id[int(match.group(1) or match.group(2))]
                description = format_beton_aciklama(sale[2], sale[3], sale[4], sale[5], sale[6], sale[9], sale[8], sale[7])
            mt.insert("","end",values=(r[0],r[1],r[2],r[3],description,r[5],f"{debt:,.2f}" if debt else "-",f"{credit:,.2f}" if credit else "-",f"{bal:,.2f}"))
        ttk.Label(mf,text=f"Cari hareket toplamı: Borç {sum(safe_float(r[6] or 0) for r in moves if r[2] == 'Borç'):,.2f} TL  |  Alacak {sum(safe_float(r[6] or 0) for r in moves if r[2] == 'Alacak'):,.2f} TL  |  Bakiye {balance:,.2f} TL",font=("Segoe UI",10,"bold")).pack(side="bottom",anchor="w",pady=(8,2))
        sf=ttk.Frame(nb,padding=6); nb.add(sf,text="Beton Satışları")
        ttk.Label(sf,text=f"BETON SATIŞ ÖZETİ     {len(sales)} satış  |  Toplam {format_m3(total_m3)} m³  |  Satış tutarı {sales_total_tl:,.2f} TL",font=("Segoe UI",11,"bold")).pack(side="top",anchor="w",pady=(0,8))
        ttk.Label(sf,text="Bilgi: Beton satış toplamı ile cari bakiye farklı olabilir. Cari bakiye yalnızca kayıtlı borç/alacak hareketlerinden hesaplanır.",font=("Segoe UI",9)).pack(side="bottom",anchor="w",pady=(8,2))
        st=DateTreeview(sf,columns=("id","tarih","cinsi","m3","fiyat","pompa","toplam","durum","santiye","aciklama"),show="headings")
        for col,head,width in zip(st.cget("columns"),("ID","Tarih","Beton Cinsi","Miktar m³","Birim Fiyat TL/m³","Pompa TL","Toplam TL","Ödeme Durumu","Şantiye","Açıklama"),(42,90,95,85,115,95,110,145,130,250)): st.heading(col,text=head); st.column(col,width=width,anchor="w")
        st.pack(fill="both",expand=True)
        for r in sales:
            detail = format_beton_aciklama(r[2],r[3],r[4],r[5],r[6],r[9],r[8],r[7])
            st.insert("","end",values=(r[0],r[1],r[2],format_m3(r[3]),f"{r[4]:,.2f}",f"{r[5]:,.2f}" if r[5] else "Yok",f"{r[6]:,.2f}",r[7],r[9] or "-",detail))
        cf=ttk.Frame(nb,padding=6); nb.add(cf,text="Bağlı Çekler")
        ct=DateTreeview(cf,columns=("id","yon","banka","no","vade","tutar","durum"),show="headings")
        for col,head,width in zip(ct.cget("columns"),("ID","Yön","Banka","Çek No","Vade","Tutar TL","Durum"),(45,100,160,140,110,140,150)): ct.heading(col,text=head); ct.column(col,width=width,anchor="w")
        ct.pack(fill="both",expand=True)
        for r in checks: ct.insert("","end",values=(r[0],r[1],r[2],r[3],r[4],f"{r[5]:,.2f}",r[6]))
        with connect() as audit_db:
            missing_sales, wrong_sales, ambiguous_sales = audit_beton_cari_amounts(audit_db, cari_id)
        if missing_sales or wrong_sales or ambiguous_sales:
            ttk.Label(sf,text=f"Cari bağlantısı: {len(missing_sales)} eksik, {len(wrong_sales)} tutar hatası, {len(ambiguous_sales)} belirsiz. Kayıtlar otomatik değiştirilmedi.",font=("Segoe UI",9,"bold")).pack(side="bottom",anchor="w")
        actions=ttk.Frame(win,padding=(12,0,12,12)); actions.pack(fill="x")
        ttk.Button(actions,text="Beton-Cari Bağlantısını Kontrol Et",
                   command=lambda:self.repair_cari_beton_links(cari_id,win)).pack(side="left",padx=(0,8))
        ttk.Button(actions,text="Cari Hesap PDF",command=lambda:self.export_single_cari_pdf(cari_id,unvan,"Cari Hesap Ekstresi")).pack(side="left",padx=(0,8))
        ttk.Button(actions,text="Beton Cari PDF",command=lambda:self.export_single_cari_pdf(cari_id,unvan,"Cari Beton Satış Dökümü")).pack(side="left",padx=(0,8))
        ttk.Button(actions,text="Excel Dökümü Al",command=lambda:self.export_single_cari_excel(cari_id,unvan)).pack(side="left",padx=(0,8))
        ttk.Button(actions,text="Seçili cari hareketini düzenle",command=lambda:self._edit_cari_move_from_window(mt)).pack(side="left",padx=(0,8))
        ttk.Button(actions,text="Kapat",command=win.destroy).pack(side="right")

    def repair_cari_beton_links(self, cari_id, window=None):
        """Backed-up, explicit correction of identifiable movements only."""
        with connect() as db:
            missing, mismatched, ambiguous = audit_beton_cari_amounts(db, cari_id)
            if not missing and not mismatched and not ambiguous:
                return messagebox.showinfo("Cari bağlantısı", "Satış bağlantıları ve tutarları uyumlu görünüyor.", parent=window or self)
            summary = (f"Eksik bağlantı: {len(missing)}\n"
                       f"Yanlış tutarlı/bağlantılı hareket: {len(mismatched)}\n"
                       f"Belirsiz kayıt (otomatik değişmez): {len(ambiguous)}\n\n"
                       f"Eksik borç toplamı: {sum(x[2] for x in missing):,.2f} TL\n"
                       f"Bağlı tutar farkı: {sum(x[2]-x[4] for x in mismatched):,.2f} TL")
            if not missing and not mismatched:
                return messagebox.showwarning("Cari bağlantısı inceleme", summary + "\n\nBelirsiz kayıtlar elle incelenmelidir.", parent=window or self)
            if not messagebox.askyesno("Cari bağlantısı onarımı", summary +
                    "\n\nÖNEMLİ: Bu işlem cari bakiyesini değiştirebilir."
                    "\nÖnce veritabanı yedeği alınarak yalnızca kesin eşleşen"
                    " kayıtlar düzeltilecek. Onaylıyor musunuz?", parent=window or self):
                return
            backup = _verified_sqlite_backup(db, "beton_cari_v77_onarim_oncesi")
            db.execute("BEGIN IMMEDIATE")
            try:
                again_missing, again_mismatch, again_ambiguous = audit_beton_cari_amounts(db, cari_id)
                for sale_id, _, _ in again_missing:
                    sync_beton_sale_cari_movement(db, sale_id)
                for sale_id, _, _, _, _, _ in again_mismatch:
                    sync_beton_sale_cari_movement(db, sale_id)
                db.commit()
            except Exception:
                db.rollback()
                raise
        if window and window.winfo_exists():
            window.destroy()
        self.refresh_all()
        messagebox.showinfo("Cari bağlantısı", f"{len(again_missing)} eksik, {len(again_mismatch)} hatalı hareket onarıldı.\n"
                            f"Belirsiz kayıt: {len(again_ambiguous)}\nYedek: {backup}", parent=self)

    def _safe_export_name(self, value):
        cleaned = "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in str(value)).strip()
        return (cleaned[:60] or "Cari")

    def _preview_then_save_pdf(self, report_type, final_path=None, cari_id=None):
        """Önce tek önizleme göster; kullanıcı onaylarsa bir kez hedef dosyayı seçtirip kaydet."""
        preview_path = None
        try:
            fd, preview_path = tempfile.mkstemp(prefix="HOYTAS_PDF_ONIZLEME_", suffix=".pdf")
            os.close(fd)
            self.generate_pdf(report_type, preview_path, cari_id)
            with open(preview_path, "rb") as stream:
                data = stream.read()
            if not data.startswith(b"%PDF-") or b"%%EOF" not in data[-2048:] or len(data) < 500:
                raise RuntimeError("Önizleme PDF'si doğrulanamadı; hiçbir dosya kaydedilmedi.")
            try:
                viewer = open_pdf_in_viewer(preview_path)
            except RuntimeError as exc:
                messagebox.showerror("PDF önizlemesi açılamadı", f"PDF kaydedilmedi. Önce Adobe Acrobat/Reader ile önizleme açılmalıdır.\n\n{exc}", parent=self)
                return False
            if not messagebox.askyesno("PDF Kaydetme Onayı", f"{report_type} önizlemesi {viewer} ile açıldı.\n\nPDF'yi kaydetmek istiyor musunuz?", parent=self):
                return False
            if not final_path:
                initial = "HOYTAS_" + report_type.replace(" ", "_").replace("/", "_") + "_" + datetime.now().strftime("%Y%m%d") + ".pdf"
                final_path = filedialog.asksaveasfilename(parent=self, title="PDF dosyasını kaydet", defaultextension=".pdf", filetypes=[("PDF dosyası", "*.pdf")], initialfile=initial)
            if not final_path:
                return False
            os.makedirs(os.path.dirname(os.path.abspath(final_path)), exist_ok=True)
            shutil.copy2(preview_path, final_path)
            with open(final_path, "rb") as stream:
                saved = stream.read()
            if not saved.startswith(b"%PDF-") or b"%%EOF" not in saved[-2048:] or len(saved) < 500:
                try: os.remove(final_path)
                except OSError: pass
                raise RuntimeError("Kaydedilen PDF doğrulamadan geçmedi; hatalı dosya kaldırıldı.")
            return True
        finally:
            if preview_path and os.path.exists(preview_path):
                try: os.remove(preview_path)
                except OSError: pass

    def export_selected_cari_pdf(self, report_type):
        sel = self.cari_tree.selection()
        if not sel:
            return messagebox.showinfo("Cari seçin", "PDF dökümü almak için önce cari hesap seçin.", parent=self)
        values = self.cari_tree.item(sel[0], "values")
        cari_id = int(values[0])
        try:
            self._preview_then_save_pdf(report_type, None, cari_id)
        except Exception as exc:
            messagebox.showerror("PDF oluşturulamadı", str(exc), parent=self)

    def export_selected_beton_cari_pdf(self):
        """Export only the customer's sales for the selected concrete-sale row."""
        selection = self.beton_tree.selection()
        if not selection:
            return messagebox.showinfo("Beton satışı seçin", "Önce listeden bir beton satışı seçin.", parent=self)
        sale_id = int(self.beton_tree.item(selection[0], "values")[0])
        with connect() as c:
            record = c.execute("SELECT cari_id FROM beton_satis WHERE id=?", (sale_id,)).fetchone()
        if not record or record[0] is None:
            return messagebox.showerror("Cari bulunamadı", "Seçilen satışın bağlı olduğu cari bulunamadı.", parent=self)
        try:
            self._preview_then_save_pdf("Cari Beton Satış Dökümü", None, int(record[0]))
        except Exception as exc:
            messagebox.showerror("Beton Cari PDF", str(exc), parent=self)

    def export_beton_cari_sec_pdf(self):
        """Beton ekranından yalnızca seçilen müşterinin satışlarını dışa aktar."""
        with connect() as c:
            cariler = c.execute("SELECT DISTINCT c.id,c.unvan FROM cari c JOIN beton_satis b ON b.cari_id=c.id ORDER BY c.unvan COLLATE NOCASE").fetchall()
        if not cariler:
            return messagebox.showinfo("Beton satışları", "PDF alınabilecek beton satışı bulunamadı.", parent=self)
        options = [f"{cid} - {name}" for cid,name in cariler]
        choice = self.prompt("Beton Cari PDF - Müşteri Seç", [("cari", "Müşteri / Cari", "combo", options)])
        if not choice:
            return
        cari_id = int(choice["cari"].split(" - ",1)[0])
        try:
            self._preview_then_save_pdf("Cari Beton Satış Dökümü", None, cari_id)
        except Exception as exc:
            messagebox.showerror("Beton Cari PDF", str(exc), parent=self)

    def export_pdf_report(self, report_type):
        """Önce önizleme aç; kaydetme onayından sonra dosya konumunu sor."""
        try:
            self._preview_then_save_pdf(report_type, None)
        except Exception as exc:
            messagebox.showerror("PDF oluşturulamadı", str(exc), parent=self)

    def export_single_cari_pdf(self, cari_id, unvan, report_type="Cari Hesap Ekstresi"):
        """Cari hareket penceresindeki cari için doğrudan doğrulanmış PDF ekstresi üret."""
        initial = f"HOYTAS_{self._safe_export_name(unvan)}_{report_type.replace(' ', '_')}_{datetime.now():%Y%m%d}.pdf"
        try:
            self._preview_then_save_pdf(report_type, None, cari_id)
        except ImportError:
            messagebox.showerror("Eksik bileşen", "PDF oluşturmak için kurulum.bat dosyasını çalıştırın veya reportlab paketini kurun.", parent=self)
        except Exception as exc:
            messagebox.showerror("PDF oluşturulamadı", str(exc), parent=self)

    def export_single_cari_excel(self, cari_id, unvan):
        """Cari hareketleri, beton satışları ve bağlı çekleri tek XLSX dosyasına aktar."""
        initial = f"HOYTAS_Cari_Ekstresi_{self._safe_export_name(unvan)}_{datetime.now():%Y%m%d}.xlsx"
        path = filedialog.asksaveasfilename(parent=self, title="Cari Excel ekstresini kaydet",
            defaultextension=".xlsx", filetypes=[("Excel çalışma kitabı", "*.xlsx")], initialfile=initial)
        if not path:
            return
        try:
            from openpyxl import Workbook, load_workbook
            from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
            from openpyxl.utils import get_column_letter
            with connect() as c:
                cari = c.execute("SELECT unvan,COALESCE(telefon,''),COALESCE(vergi_no,'') FROM cari WHERE id=?", (cari_id,)).fetchone()
                if not cari:
                    raise ValueError("Cari hesap bulunamadı.")
                moves = c.execute("SELECT tarih,tur,kategori,COALESCE(aciklama,''),COALESCE(odeme_yontemi,''),tutar FROM hareket WHERE cari_id=? AND tur IN ('Borç','Alacak') ORDER BY tarih ASC,id ASC", (cari_id,)).fetchall()
                sales = c.execute("SELECT b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(b.aciklama,''),COALESCE(s.ad,'') FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC", (cari_id,)).fetchall()
                checks = c.execute("SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),COALESCE(cek_no,''),vade,tutar,durum,COALESCE(notlar,'') FROM cek WHERE cari_id=? ORDER BY vade ASC,id ASC", (cari_id,)).fetchall()
            wb = Workbook()
            summary = wb.active; summary.title = "Cari Özeti"
            summary.append(["HOYTAŞ HAZIR BETON | CARİ HESAP EKSTRESİ"])
            summary.append(["Cari Ünvanı", cari[0]])
            summary.append(["Telefon", cari[1] or "-"])
            summary.append(["Vergi No", cari[2] or "-"])
            debt = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Borç")
            credit = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Alacak")
            summary.append(["Toplam Borç (TL)", debt])
            summary.append(["Toplam Alacak (TL)", credit])
            summary.append(["Bakiye (Borç - Alacak) (TL)", debt-credit])
            total_m3 = sum(safe_float(r[2] or 0) for r in sales)
            total_sales_amount = sum(safe_float(r[5] or 0) for r in sales)
            total_collected = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Alacak" and ("tahsilat" in str(r[2]).casefold() or "tahsilat" in str(r[3]).casefold()))
            summary.append(["Toplam Beton Satışı (m³)", safe_float(total_m3)])
            summary.append(["Toplam Satış Tutarı (TL)", total_sales_amount])
            summary.append(["Toplam Tahsilat (TL)", total_collected])
            summary.append(["Kalan Cari Bakiye (TL)", debt-credit])
            summary.append(["Rapor Tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")])
            sh = wb.create_sheet("Cari Hareketleri")
            sh.append(["Tarih", "Tür", "Kategori", "Açıklama", "Ödeme Yöntemi", "Borç (TL)", "Alacak (TL)", "Bakiye (TL)"])
            balance = 0
            for r in moves:
                balance += safe_float(r[5] or 0) if r[1] == "Borç" else -safe_float(r[5] or 0)
                sh.append([display_date(r[0]), r[1], r[2], r[3], r[4], safe_float(r[5] or 0) if r[1] == "Borç" else 0, safe_float(r[5] or 0) if r[1] == "Alacak" else 0, balance])
            sh.append(["", "", "", "", "TOPLAM", debt, credit, balance])
            ss = wb.create_sheet("Beton Satışları")
            ss.append(["Tarih", "Beton Cinsi", "Miktar (m³)", "Birim Fiyat (TL)", "Pompa Ücreti (TL)", "Toplam (TL)", "Ödeme Durumu", "Açıklama"])
            for r in sales: ss.append([display_date(r[0]), r[1], safe_float(r[2] or 0), r[3], r[4], r[5], r[6], format_beton_aciklama(r[1],r[2],r[3],r[4],r[5],r[8],r[7],r[6])])
            sales_total_row = ss.max_row + 1
            ss.append(["", "TOPLAM BETON (m³)", f"=SUM(C2:C{sales_total_row-1})", "", "", f"=SUM(F2:F{sales_total_row-1})", "", ""])
            sc = wb.create_sheet("Bağlı Çekler")
            sc.append(["Yön", "Banka", "Banka Kodu", "Çek No", "Vade", "Tutar (TL)", "Durum", "Notlar"])
            for r in checks: sc.append([r[0], r[1], r[2], r[3], display_date(r[4]), r[5], r[6], r[7]])
            header_fill = PatternFill("solid", fgColor="24583D")
            accent_fill = PatternFill("solid", fgColor="F2F7EF")
            yellow_fill = PatternFill("solid", fgColor="F6E8A8")
            white_bold = Font(name="Aptos", bold=True, color="FFFFFF")
            title_font = Font(name="Aptos Display", bold=True, size=15, color="24583D")
            label_font = Font(name="Aptos", bold=True, color="24583D")
            thin = Side(style="thin", color="D6DFD2")
            for ws in wb.worksheets:
                ws.freeze_panes = "A2"
                ws.sheet_view.showGridLines = False
                for row in ws.iter_rows():
                    for cell in row:
                        cell.border = Border(bottom=thin)
                        cell.alignment = Alignment(vertical="center", wrap_text=True)
                if ws.title == "Cari Özeti":
                    ws.merge_cells("A1:B1")
                    ws["A1"].font = title_font
                    ws["A1"].fill = accent_fill
                    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
                    ws.row_dimensions[1].height = 30
                    for rr in range(2, ws.max_row + 1):
                        ws.cell(rr, 1).font = label_font
                        ws.cell(rr, 1).fill = accent_fill
                    ws.column_dimensions["A"].width = 34
                    ws.column_dimensions["B"].width = 48
                else:
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = white_bold
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    ws.row_dimensions[1].height = 32
                    ws.auto_filter.ref = ws.dimensions
                    for rr in range(2, ws.max_row + 1):
                        if rr % 2 == 0:
                            for cell in ws[rr]:
                                cell.fill = accent_fill
                    if ws.max_row >= 2 and str(ws.cell(ws.max_row, 5).value or "").startswith("TOPLAM"):
                        for cell in ws[ws.max_row]:
                            cell.fill = yellow_fill
                            cell.font = Font(name="Aptos", bold=True, color="24583D")
                for col in ws.columns:
                    letter = get_column_letter(col[0].column)
                    max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                    ws.column_dimensions[letter].width = min(max(max_len + 2, 13), 46)
            for row in (5, 6, 7, 9, 10, 11): summary.cell(row, 2).number_format = '#,##0.00 "TL"'
            summary.cell(8, 2).number_format = '#,##0.### "m³"'
            for ws in (sh, ss, sc):
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if cell.column in (6, 7, 8) and ws.title == "Cari Hareketleri": cell.number_format = '#,##0.00'
                        if ws.title == "Beton Satışları" and cell.column in (4, 5, 6): cell.number_format = '#,##0.00'
                        if ws.title == "Bağlı Çekler" and cell.column == 6: cell.number_format = '#,##0.00'
            wb.save(path)
            check = load_workbook(path, read_only=True, data_only=True)
            expected = {"Cari Özeti", "Cari Hareketleri", "Beton Satışları", "Bağlı Çekler"}
            if not expected.issubset(set(check.sheetnames)):
                raise RuntimeError("Excel dosyasının sayfaları doğrulanamadı.")
            check.close()
            with open(path, "rb") as stream:
                if stream.read(2) != b"PK": raise RuntimeError("Dosya geçerli XLSX biçiminde değil.")
            messagebox.showinfo("Excel hazır", f"Cari ekstresi Excel olarak kaydedildi ve doğrulandı:\n{path}", parent=self)
        except ImportError:
            messagebox.showerror("Eksik bileşen", "Excel oluşturmak için openpyxl paketini kurun veya kurulum.bat dosyasını çalıştırın.", parent=self)
        except Exception as exc:
            messagebox.showerror("Excel oluşturulamadı", str(exc), parent=self)

    def _edit_cari_move_from_window(self, tree):
        sel=tree.selection()
        if not sel: return messagebox.showinfo("Hareket seçin","Düzenlemek için bir hareket seçin.",parent=self)
        rid=int(tree.item(sel[0],"values")[0])
        # Hareket ana listesinde aynı ID'li kaydı seçip standart düzenleme formunu kullan.
        if hasattr(self,"hareket_tree"):
            for item in self.hareket_tree.get_children():
                if int(self.hareket_tree.item(item,"values")[0])==rid:
                    self.tabs.select(self.hareket_tab); self.hareket_tree.selection_set(item); self.hareket_tree.focus(item)
                    self.edit_selected(self.hareket_tree,"hareket"); self.refresh_all(); return
        return messagebox.showinfo("Hareket bulunamadı","Bu cari hareketi ana hareket listesinde bulunamadı.",parent=self)

    def _build_cari(self):
        bar = ttk.Frame(self.cari_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Cari Hesaplar", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Cari Ekle", command=self.add_cari).pack(side="right")
        self.cari_search=tk.StringVar()
        ttk.Label(bar,text="Ara:").pack(side="left",padx=(18,4))
        search=ttk.Entry(bar,textvariable=self.cari_search,width=28); search.pack(side="left"); self.cari_search_entry=search
        self.cari_search.trace_add("write",lambda *_: self.refresh_cari())
        legend=ttk.Frame(self.cari_tab); legend.pack(fill="x",pady=(6,0))
        ttk.Label(legend,text="● Ödeme yapılmamış borç",foreground="#9B1C1C",font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,14))
        ttk.Label(legend,text="● Kısmi ödeme / kalan borç",foreground="#8A4B08",font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,14))
        ttk.Label(legend,text="● Borç kapalı",foreground="#216438",font=("Segoe UI",9,"bold")).pack(side="left")
        self.cari_tree = self._tree(self.cari_tab, ("id","tur","unvan","telefon","vergi_no","borc","alacak","bakiye"),
                                    ("ID","Tür","Ünvan","Telefon","Vergi No","Borç (TL)","Alacak (TL)","Kalan Borç (TL)"),
                                    (45,85,220,120,110,115,115,125))
        row=ttk.Frame(self.cari_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kaydı Düzenle",command=lambda:self.edit_selected(self.cari_tree,"cari")).pack(side="left")
        ttk.Button(row,text="Cari Hareketlerini Gör",command=self.show_cari_hareketleri).pack(side="left",padx=(0,8))
        ttk.Button(row,text="Cari Ekstresi PDF",command=lambda:self.export_selected_cari_pdf("Cari Hesap Ekstresi")).pack(side="left",padx=(0,8))
        ttk.Button(row,text="Cari Hareket PDF",command=lambda:self.export_selected_cari_pdf("Cari Hareket Dökümü")).pack(side="left",padx=(0,8))
        ttk.Button(row,text="Tahsilat Al / Cari Borcunu Düş",command=self.cari_tahsilat).pack(side="left",padx=8)
        ttk.Button(row, text="Seçili Kaydı Sil", command=lambda: self.delete_selected(self.cari_tree,"cari")).pack(side="right")

    def _build_borc_takip(self):
        """Borçlu carileri tek listede göster; tutarlar hareketlerden hesaplanır."""
        bar = ttk.Frame(self.borc_takip_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Cari Borç Takip Listesi", font=("Segoe UI", 16, "bold")).pack(side="left")
        self.borc_search = tk.StringVar()
        ttk.Label(bar, text="Cari ara:").pack(side="left", padx=(18, 4))
        search = ttk.Entry(bar, textvariable=self.borc_search, width=28)
        search.pack(side="left")
        self.borc_search.trace_add("write", lambda *_: self.refresh_borc_takip())
        ttk.Button(bar, text="Listeyi Yenile", command=self.refresh_all).pack(side="right", padx=(6, 0))
        ttk.Label(self.borc_takip_tab,
                  text="Yalnızca kalan borcu bulunan cariler listelenir. Ödeme kaydı girildiğinde kalan tutar otomatik güncellenir.",
                  foreground="#53645A").pack(anchor="w", pady=(8, 6))
        self.borc_tree = self._tree(self.borc_takip_tab,
            ("id", "unvan", "borc", "odenen", "kalan", "durum"),
            ("ID", "Cari Ünvanı", "Toplam Borç (TL)", "Ödenen (TL)", "Kalan Borç (TL)", "Durum"),
            (50, 300, 150, 150, 160, 145))
        self.borc_tree.tag_configure("debt_unpaid", foreground="#9B1C1C")
        self.borc_tree.tag_configure("debt_partial", foreground="#8A4B08")
        summary = ttk.Frame(self.borc_takip_tab); summary.pack(fill="x", pady=8)
        self.borc_total_label = ttk.Label(summary, text="Toplam kalan borç: 0,00 TL",
                                          font=("Segoe UI", 12, "bold"), foreground="#9B1C1C")
        self.borc_total_label.pack(side="left")
        actions = ttk.Frame(self.borc_takip_tab); actions.pack(fill="x", pady=(2, 0))
        ttk.Button(actions, text="Borç Listesi PDF Al",
                   command=lambda: self.export_pdf_report("Cari Borç Takip Raporu")).pack(side="left")
        ttk.Button(actions, text="Seçili Carinin Hareketlerini Gör",
                   command=self._open_selected_debt_cari).pack(side="left", padx=8)

    def _debt_rows(self):
        with connect() as c:
            return c.execute("""SELECT cari_id,unvan,toplam_borc,toplam_alacak,bakiye
                FROM cari_bakiye WHERE bakiye > 0.005
                ORDER BY bakiye DESC,unvan COLLATE NOCASE""").fetchall()

    def refresh_borc_takip(self):
        if not hasattr(self, "borc_tree"):
            return
        term = normalize_search_text(self.borc_search.get()) if hasattr(self, "borc_search") else ""
        rows = self._debt_rows()
        visible = []
        for cid, name, debt, paid, remaining in rows:
            if term and term not in normalize_search_text(name):
                continue
            status = "Kısmi ödeme" if safe_float(paid or 0) > 0.005 else "Ödeme bekliyor"
            tag = "debt_partial" if safe_float(paid or 0) > 0.005 else "debt_unpaid"
            visible.append(((cid, name, f"{debt:,.2f}", f"{paid:,.2f}", f"{remaining:,.2f}", status), tag))
        self.borc_tree.delete(*self.borc_tree.get_children())
        for values, tag in visible:
            self.borc_tree.insert("", "end", values=values, tags=(tag,))
        visible_total = sum(safe_float(item[0][4]) for item in visible)
        self.borc_total_label.config(text=f"Toplam kalan borç: {visible_total:,.2f} TL   |   Borçlu cari: {len(visible)}")

    def _open_selected_debt_cari(self):
        sel = self.borc_tree.selection()
        if not sel:
            return messagebox.showinfo("Cari seçin", "Hareketlerini görmek için borç listesinden bir cari seçin.", parent=self)
        values = self.borc_tree.item(sel[0], "values")
        self.tabs.select(self.cari_tab)
        if hasattr(self, "cari_search"):
            self.cari_search.set("")
        self.refresh_cari()
        for item in self.cari_tree.get_children():
            if int(self.cari_tree.item(item, "values")[0]) == int(values[0]):
                self.cari_tree.selection_set(item); self.cari_tree.focus(item)
                self.show_cari_hareketleri()
                break

    def _build_alisveris_cari(self):
        bar=ttk.Frame(self.alisveris_cari_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Çarşı-Pazar ve Sanayi Alışverişleri",font=("Segoe UI",16,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Tedarikçi Cari",command=self.add_tedarikci_cari).pack(side="right")
        ttk.Label(self.alisveris_cari_tab,
            text="Alışverişleri tedarikçi carisine borç yazın; ödeme yaptığınızda aynı carinin bakiyesi otomatik azalır.",
            foreground="#53645A").pack(anchor="w",pady=(8,6))
        self.alisveris_tree=self._tree(self.alisveris_cari_tab,
            ("id","tarih","cari","tur","kategori","aciklama","borc","alacak","bakiye"),
            ("ID","Tarih","Tedarikçi Cari","İşlem","Tür","Açıklama","Borç TL","Ödeme TL","Kalan TL"),
            (45,100,190,90,155,250,105,105,110))
        row=ttk.Frame(self.alisveris_cari_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Çarşı-Pazar Alışverişi Ekle",command=lambda:self.add_supplier_purchase("Çarşı-Pazar Alışverişi")).pack(side="left")
        ttk.Button(row,text="Sanayi Alışverişi Ekle",command=lambda:self.add_supplier_purchase("Sanayi Alışverişi")).pack(side="left",padx=8)
        ttk.Button(row,text="Tedarikçiye Ödeme Yap",command=self.pay_supplier).pack(side="left")
        ttk.Button(row,text="Listeyi Yenile",command=self.refresh_alisveris_cari).pack(side="right")

    def add_tedarikci_cari(self):
        d=self.prompt("Yeni Tedarikçi Cari",[("unvan","Tedarikçi / Esnaf Ünvanı","entry",None),
            ("telefon","Telefon","entry",None),("notlar","Notlar","entry",None)])
        if not d:return
        unvan=(d["unvan"] or "").strip()
        if not unvan:return messagebox.showwarning("Eksik bilgi","Tedarikçi adı boş bırakılamaz.",parent=self)
        with connect() as c:
            if c.execute("SELECT id FROM cari WHERE unvan=? COLLATE NOCASE",(unvan,)).fetchone():
                return messagebox.showwarning("Kayıt mevcut","Bu ünvanla cari zaten kayıtlı.",parent=self)
            c.execute("INSERT INTO cari(tur,unvan,telefon,vergi_no,notlar,created_at) VALUES(?,?,?,?,?,?)",
                ("Tedarikçi",unvan,d["telefon"],"",d["notlar"],datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.refresh_all()
        messagebox.showinfo("Kaydedildi",f"{unvan} tedarikçi carisi oluşturuldu.",parent=self)

    def _supplier_choices(self):
        with connect() as c:
            rows=c.execute("SELECT id,unvan FROM cari WHERE tur='Tedarikçi' ORDER BY unvan COLLATE NOCASE").fetchall()
        return [f"{r[0]} - {r[1]}" for r in rows]

    def add_supplier_purchase(self, category):
        choices=self._supplier_choices()
        if not choices:return messagebox.showinfo("Tedarikçi gerekli","Önce Yeni Tedarikçi Cari düğmesiyle tedarikçi oluşturun.",parent=self)
        d=self.prompt(category,[("tarih","Alışveriş tarihi (GG.AA.YYYY)","entry",None),
            ("cari","Tedarikçi cari","combo_search",choices),("aciklama","Mal / hizmet açıklaması","entry",None),
            ("tutar","Toplam tutar (TL)","entry",None)])
        if not d:return
        try:
            tarih=d["tarih"] or datetime.now().strftime("%Y-%m-%d"); datetime.strptime(tarih,"%Y-%m-%d")
            cari_id=int(d["cari"].split(" - ",1)[0]); tutar=parse_amount(d["tutar"])
            if tutar<=0:raise ValueError()
        except Exception:return messagebox.showerror("Hatalı bilgi","Tarih, tedarikçi ve tutarı kontrol edin.",parent=self)
        acik=(d["aciklama"] or category).strip()
        with connect() as c:
            if not c.execute("SELECT 1 FROM cari WHERE id=?",(cari_id,)).fetchone():
                return messagebox.showerror("Cari bulunamadı","Seçilen tedarikçi bulunamadı.",parent=self)
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Borç",category,acik,tutar,"Vadeli"))
        self.refresh_all()
        messagebox.showinfo("Alışveriş kaydedildi",f"{tutar:,.2f} TL alışveriş tedarikçi borcuna eklendi.",parent=self)

    def pay_supplier(self):
        choices=self._supplier_choices()
        if not choices:return messagebox.showinfo("Tedarikçi yok","Önce tedarikçi cari oluşturun.",parent=self)
        d=self.prompt("Tedarikçi Ödemesi",[("tarih","Ödeme tarihi (GG.AA.YYYY)","entry",None),
            ("cari","Tedarikçi cari","combo_search",choices),("tutar","Ödeme tutarı (TL)","entry",None),
            ("yontem","Ödeme yöntemi","combo",["Nakit","Banka Havalesi / EFT","Kredi Kartı"]),
            ("banka","Banka adı (banka ise)","entry",None),("aciklama","Açıklama","entry",None)])
        if not d:return
        try:
            tarih=d["tarih"] or datetime.now().strftime("%Y-%m-%d"); datetime.strptime(tarih,"%Y-%m-%d")
            cari_id=int(d["cari"].split(" - ",1)[0]); tutar=parse_amount(d["tutar"])
            if tutar<=0:raise ValueError()
        except Exception:return messagebox.showerror("Hatalı bilgi","Tarih, cari ve ödeme tutarını kontrol edin.",parent=self)
        with connect() as c:
            bal=c.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?",(cari_id,)).fetchone()
            if not bal or safe_float(bal[0])<=0.005:
                return messagebox.showwarning("Açık borç yok","Bu tedarikçinin ödenecek açık borcu bulunmuyor.",parent=self)
            if tutar>safe_float(bal[0])+0.005:
                return messagebox.showwarning("Tutar fazla",f"Ödeme kalan borcu aşamaz. Kalan borç: {safe_float(bal[0]):,.2f} TL",parent=self)
            name=c.execute("SELECT unvan FROM cari WHERE id=?",(cari_id,)).fetchone()[0]
            acik=(d["aciklama"] or f"{name} tedarikçi ödemesi").strip()
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Alacak","Tedarikçi Ödemesi",acik,tutar,d["yontem"]))
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Gider","Alışveriş Ödemesi",acik,tutar,d["yontem"]))
            if d["yontem"]=="Nakit":
                c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                    (tarih,"Para Çıkışı","Tedarikçi Ödemesi",acik,tutar,getattr(self,"current_user",{}).get("username","")))
            else:
                c.execute("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                    (tarih,d["banka"] or "Banka","", "Para Çıkışı",acik,tutar))
        self.refresh_all()
        messagebox.showinfo("Ödeme kaydedildi",f"{tutar:,.2f} TL ödeme kaydedildi; tedarikçi cari borcu düşürüldü.",parent=self)

    def refresh_alisveris_cari(self):
        if not hasattr(self,"alisveris_tree"):return
        with connect() as c:
            rows=c.execute("""SELECT h.id,h.tarih,COALESCE(c.unvan,''),h.tur,h.kategori,
                COALESCE(h.aciklama,''),CASE WHEN h.tur='Borç' THEN h.tutar ELSE 0 END,
                CASE WHEN h.tur='Alacak' THEN h.tutar ELSE 0 END,COALESCE(b.bakiye,0)
                FROM hareket h LEFT JOIN cari c ON c.id=h.cari_id
                LEFT JOIN cari_bakiye b ON b.cari_id=h.cari_id
                WHERE h.kategori IN ('Çarşı-Pazar Alışverişi','Sanayi Alışverişi','Tedarikçi Ödemesi')
                ORDER BY h.tarih ASC,h.id ASC""").fetchall()
        self.fill_tree(self.alisveris_tree,[(r[0],display_date(r[1]),r[2],r[3],r[4],r[5],
            f"{r[6]:,.2f}",f"{r[7]:,.2f}",f"{r[8]:,.2f}") for r in rows])

    def _build_hareket(self):
        bar = ttk.Frame(self.hareket_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Gelir / Gider Kayıtları", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Hareket Ekle", command=self.add_hareket).pack(side="right")
        self.hareket_tree = self._tree(self.hareket_tab, ("id","tarih","tur","kategori","cari","aciklama","tutar","odeme"),
            ("ID","Tarih","Tür","Kategori","Cari","Açıklama","Tutar (TL)","Ödeme"), (50,105,90,130,200,220,120,110))
        row=ttk.Frame(self.hareket_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kaydı Düzenle",command=lambda:self.edit_selected(self.hareket_tree,"hareket")).pack(side="left")
        ttk.Button(row, text="Seçili Kaydı Sil", command=lambda: self.delete_selected(self.hareket_tree,"hareket")).pack(side="right")

    def _build_cek(self):
        bar = ttk.Frame(self.cek_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Alınan / Verilen Çekler", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Çek Ekle", command=self.add_cek).pack(side="right")
        self.cek_tree = self._tree(self.cek_tab, ("id","yon","banka","cek_no","taraf","vade","tutar","durum"),
            ("ID","Yön","Banka","Çek No","Keşideci / Lehtar","Vade","Tutar (TL)","Durum"), (50,100,130,120,220,110,120,110))
        row=ttk.Frame(self.cek_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Çeki Düzenle",command=lambda:self.edit_selected(self.cek_tree,"cek")).pack(side="left")
        ttk.Button(row,text="Alınan Çekler PDF",command=lambda:self.export_pdf_report("Alınan Çekler Dökümü")).pack(side="left",padx=(8,0))
        ttk.Button(row,text="Verilen Çekler PDF",command=lambda:self.export_pdf_report("Verilen Çekler Dökümü")).pack(side="left",padx=(8,0))
        ttk.Button(row, text="Seçili Çeki Sil", command=lambda: self.delete_selected(self.cek_tree,"cek")).pack(side="right")

    def _build_stok(self):
        bar = ttk.Frame(self.stok_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Stok / Depo", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Stok Kalemi Ekle", command=self.add_stok).pack(side="right")
        self.stok_tree = self._tree(self.stok_tab, ("id","kod","ad","birim","miktar","fiyat","deger"),
            ("ID","Kod","Malzeme","Birim","Miktar","Birim Fiyat","Stok Değeri (TL)"), (50,100,260,90,110,130,150))
        row=ttk.Frame(self.stok_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kalemi Düzenle",command=lambda:self.edit_selected(self.stok_tree,"stok")).pack(side="left")
        ttk.Button(row, text="Seçili Kalemi Sil", command=lambda: self.delete_selected(self.stok_tree,"stok")).pack(side="right")

    def _build_personel(self):
        bar = ttk.Frame(self.personel_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Personel", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Personel Ekle", command=self.add_personel).pack(side="right")
        self.personel_tree = self._tree(self.personel_tab, ("id","ad","gorev","telefon","maas","notlar"),
            ("ID","Ad Soyad","Görev","Telefon","Aylık Maaş (TL)","Notlar"), (50,220,180,140,150,300))
        row=ttk.Frame(self.personel_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Personeli Düzenle",command=lambda:self.edit_selected(self.personel_tree,"personel")).pack(side="left")
        ttk.Button(row, text="Seçili Personeli Sil", command=lambda: self.delete_selected(self.personel_tree,"personel")).pack(side="right")

    def _build_beton(self):
        bar = ttk.Frame(self.beton_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Beton Satış ve Metreküp Takibi", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Beton Satışı", command=self.add_beton).pack(side="right")
        self.beton_tree = self._tree(self.beton_tab, ("id","tarih","cari","cinsi","m3","fiyat","pompa","toplam","durum","aciklama"),
            ("ID","Tarih","Müşteri","Beton Cinsi","Miktar (m³)","Birim Fiyat (TL)","Pompa Ücreti (TL)","Toplam (TL)","Ödeme Durumu","Açıklama"),
            (45,100,190,100,95,125,125,120,125,200))
        row=ttk.Frame(self.beton_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Satışı Düzenle",command=lambda:self.edit_selected(self.beton_tree,"beton_satis")).pack(side="left")
        ttk.Button(row,text="Beton Satışları PDF Dökümü",command=lambda:self.export_pdf_report("Beton Satışları Dökümü")).pack(side="left",padx=(8,0))
        ttk.Button(row,text="Seçili Satışın Cari PDF",command=self.export_selected_beton_cari_pdf).pack(side="left",padx=(8,0))
        ttk.Button(row,text="Cari Seçerek PDF",command=self.export_beton_cari_sec_pdf).pack(side="left",padx=(8,0))
        ttk.Button(row,text="Listeyi Yenile",command=self.refresh_all).pack(side="left",padx=(8,0))
        ttk.Button(row, text="Seçili Satışı Sil", command=lambda: self.delete_selected(self.beton_tree,"beton_satis")).pack(side="right")

    def _build_kasa(self):
        bar = ttk.Frame(self.kasa_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Nakit Kasa | Elden Gelen ve Çıkan Paralar", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="+ Para Girişi / Çıkışı Ekle", command=self.add_kasa).pack(side="right")
        self.kasa_ozet = ttk.Label(self.kasa_tab, text="Kasa bakiyesi: 0,00 TL", font=("Segoe UI", 12, "bold"))
        self.kasa_ozet.pack(anchor="w", pady=(10, 0))
        self.kasa_tree = self._tree(self.kasa_tab, ("id","tarih","yon","kaynak","aciklama","tutar","kullanici"),
            ("ID","Tarih","İşlem","Kaynak","Açıklama","Tutar (TL)","Kaydeden"), (45,110,100,160,330,130,140))
        row = ttk.Frame(self.kasa_tab); row.pack(fill="x", pady=8)
        ttk.Button(row, text="Seçili Kasa Kaydını Düzenle", command=lambda:self.edit_selected(self.kasa_tree,"nakit_kasa")).pack(side="left")
        ttk.Button(row, text="Seçili Kasa Kaydını Sil", command=lambda:self.delete_selected(self.kasa_tree,"nakit_kasa")).pack(side="right")

    def add_kasa(self):
        d = self.prompt("Nakit Kasa İşlemi", [("tarih","Tarih (GG.AA.YYYY)","entry",None),
            ("yon","İşlem yönü","combo",["Para Girişi","Para Çıkışı"]),
            ("kaynak","Para nereden / kime","combo",["Elden Tahsilat","Müşteri Ödemesi","Ortak / Sermaye","Banka Çekimi","Diğer"]),
            ("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None)])
        if not d: return
        try:
            tarih=d["tarih"] or datetime.now().strftime("%Y-%m-%d")
            datetime.strptime(tarih,"%Y-%m-%d")
            raw_tutar = d["tutar"].strip().replace("₺", "").replace(" ", "")
            # Türkçe 1.234,56 ve standart 1234.56 biçimlerini doğru yorumla.
            if "," in raw_tutar:
                raw_tutar = raw_tutar.replace(".", "").replace(",", ".")
            tutar = safe_float(raw_tutar)
            if not (tutar > 0 and tutar < 1_000_000_000_000): raise ValueError()
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Tarih GG.AA.YYYY biçiminde ve tutar sıfırdan büyük olmalı.", parent=self)
        with connect() as c:
            c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                (tarih,d["yon"],d["kaynak"],d["aciklama"],tutar,self.current_user.get("username","")))
        self.refresh_all()

    def prompt(self, title, fields, initial=None):
        initial = initial or {}
        win = tk.Toplevel(self); win.title(title); win.transient(self); win.grab_set()
        win.resizable(False, False)
        entries = {}
        body = ttk.Frame(win, padding=16); body.pack(fill="both", expand=True)
        ttk.Label(body, text="Bilgileri kontrol edin. Cari alanında serbestçe yazabilir, önerileri ok tuşuyla açabilirsiniz.",
                  foreground="#315B46", wraplength=590).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        for r, (key, label, kind, options) in enumerate(fields):
            r += 1
            ttk.Label(body, text=label).grid(row=r, column=0, sticky="w", padx=(0,12), pady=5)
            if kind == "combo_search":
                # Tek alanda arama + seçim: kullanıcı cari ünvanını yazar, eşleşmeler açılır.
                widget = ttk.Combobox(body, values=options, state="normal", width=38)
                widget.grid(row=r, column=1, columnspan=2, pady=5, sticky="ew")
                widget.set(options[0] if options else "")
                if key in initial and initial[key] is not None:
                    widget.set(str(initial[key]))
                # Kullanıcı metni serbestçe yazar. Açılır pencereyi her tuşta
                # zorla açmak Windows ttk Combobox üzerinde odağı çalıyordu.
                # İpuçları filtrelenir, açılır liste kullanıcı isteğiyle açılır.
                all_customer_options = list(options or [])
                def suggest_customer(event=None, combo=widget, original=all_customer_options):
                    if event is not None and event.keysym in {
                        "Up", "Down", "Left", "Right", "Return", "Escape", "Tab",
                        "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"
                    }:
                        return
                    typed = combo.get()
                    matches = filter_customer_options(original, typed) if typed.strip() else original
                    combo.configure(values=matches)
                    # combo.set(), selection_set(), Post() çağrısı YOK:
                    # klavye odağı ve kullanıcının yazdığı harfler korunur.
                widget.bind("<KeyRelease>", suggest_customer)
                widget.configure(postcommand=lambda combo=widget, original=all_customer_options:
                                 combo.configure(values=(filter_customer_options(original, combo.get())
                                                         if combo.get().strip() else original)))
                entries[key] = widget
                continue
            if kind in ("combo", "combo_edit"):
                widget = ttk.Combobox(body, values=options, state=("normal" if kind == "combo_edit" else "readonly"), width=30)
                widget.set(options[0] if options and kind != "combo_edit" else "")
                if kind == "combo_edit":
                    all_options = list(options or [])
                    def suggest_values(_event=None, combo=widget, original=all_options):
                        typed = combo.get()
                        query = typed.split(" - ", 1)[-1].strip().casefold()
                        if not query:
                            combo.configure(values=original)
                            return
                        matches = [item for item in original if query in item.casefold()]
                        combo.configure(values=matches if matches else original)
                    widget.bind("<KeyRelease>", suggest_values)
            else:
                widget = ttk.Entry(body, width=33)
            if key in initial and initial[key] is not None:
                try:
                    initial_value = display_date(initial[key]) if key.lower() in ("tarih", "vade") else str(initial[key])
                    if isinstance(widget, ttk.Combobox): widget.set(initial_value)
                    else: widget.insert(0, initial_value)
                except tk.TclError: pass
            widget.grid(row=r, column=1, pady=5)
            entries[key] = widget
        result = {}
        def save():
            for key, widget in entries.items():
                result[key] = widget.get().strip()
                if key.lower() in ("tarih", "vade") and result[key]:
                    try: result[key] = normalize_date(result[key])
                    except ValueError:
                        return messagebox.showerror("Hatalı tarih", "Tarihi GG.AA.YYYY biçiminde yazın. Örnek: 26.10.2026", parent=win)
            win.destroy()
        # Alanlar 1..len(fields) satırlarında; düğmeler ayrı SON satırda olmalı.
        # Önceki sürümde row=len(fields) son alanla çakışıyordu.
        btns = ttk.Frame(body)
        btns.grid(row=len(fields) + 1, column=0, columnspan=3,
                  sticky="e", pady=(16, 2))
        win.bind("<Return>", lambda _e: save())
        win.bind("<Escape>", lambda _e: win.destroy())
        ttk.Button(btns, text="Kaydet", command=save).pack(side="left", padx=4)
        ttk.Button(btns, text="İptal", command=win.destroy).pack(side="left")
        self.wait_window(win)
        return result if result else None

    def _double_click_edit(self, tree):
        mapping={getattr(self,n,None):t for n,t in [("cari_tree","cari"),("hareket_tree","hareket"),("cek_tree","cek"),("stok_tree","stok"),("personel_tree","personel"),("beton_tree","beton_satis"),("fatura_tree","fatura"),("kasa_tree","nakit_kasa"),("santiye_tree","santiye"),("stok_hareket_tree","stok_hareket"),("maas_tree","maas_odeme"),("banka_tree","banka_hareket")]}
        table=mapping.get(tree)
        if table:self.edit_selected(tree,table)

    def _resolve_site(self, value):
        value=(value or "").strip()
        if not value: return None
        if " - " in value and value.split(" - ",1)[0].isdigit():
            return int(value.split(" - ",1)[0])
        with connect() as c:
            row=c.execute("SELECT id FROM santiye WHERE lower(ad)=lower(?)",(value,)).fetchone()
            if row: return row[0]
            cur=c.execute("INSERT INTO santiye(ad) VALUES(?)",(value,))
            return cur.lastrowid

    def edit_selected(self, tree, table):
        sel=tree.selection()
        if not sel: return messagebox.showinfo("Kayıt seçin","Düzenlemek için önce bir satır seçin.",parent=self)
        rid=int(tree.item(sel[0],"values")[0])
        if table == "hareket":
            with connect() as c:
                linked = c.execute("SELECT aciklama FROM hareket WHERE id=?", (rid,)).fetchone()
            sale_id = linked_beton_sale_id(linked[0] if linked else "")
            if sale_id:
                return messagebox.showinfo(
                    "Beton satışına bağlı cari hareket",
                    f"Bu cari hareketi Beton Satış #{sale_id} kaydından otomatik oluşuyor. "
                    "Tutar, tarih veya cari değişikliğini Beton Satış sayfasından yapın; "
                    "cari hareketi ve ilgili özetler birlikte güncellenecektir.", parent=self)
        # Formlar doğrudan veritabanındaki kayıt alanlarını düzenler; ID değişmez.
        specs={
          "cari":[("tur","Tür","combo",["Müşteri","Tedarikçi","Diğer"]),("unvan","Ünvan / Ad Soyad","entry",None),("telefon","Telefon","entry",None),("vergi_no","Vergi / T.C. No","entry",None),("notlar","Notlar","entry",None)],
          "hareket":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("tur","Tür","combo",["Gelir","Gider","Borç","Alacak"]),("kategori","Kategori","entry",None),("cari_id","Cari ID (boş bırakılabilir)","entry",None),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None),("odeme_yontemi","Ödeme yöntemi","entry",None)],
          "cek":[("yon","Yön","combo",["Alınan","Verilen"]),("banka","Banka adı","entry",None),("banka_kodu","Banka kodu","entry",None),("cari_id","Cari ID (boş bırakılabilir)","entry",None),("cek_no","Çek No","entry",None),("kesideci_lehtar","Keşideci / Lehtar","entry",None),("vade","Vade (GG.AA.YYYY)","entry",None),("tutar","Tutar (TL)","entry",None),("durum","Durum","combo",["Bekliyor","Tahsil Edildi","Ödendi","İade","Karşılıksız"]),("notlar","Notlar","entry",None)],
          "stok":[("kod","Stok Kodu","entry",None),("ad","Malzeme Adı","entry",None),("birim","Birim","combo",["Adet","Kg","Ton","m³","Litre","Metre","Paket"]),("miktar","Miktar","entry",None),("birim_fiyat","Birim Fiyat (TL)","entry",None)],
          "personel":[("ad","Ad Soyad","entry",None),("gorev","Görev","entry",None),("telefon","Telefon","entry",None),("maas","Aylık Maaş (TL)","entry",None),("notlar","Notlar","entry",None)],
          "beton_satis":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("cari_id","Cari ID (boş bırakılabilir)","entry",None),("santiye_id","Gideceği şantiye (boş bırakılabilir, elle yazılabilir)","combo_edit",[]),("beton_cinsi","Beton Cinsi","combo",["C16","C20","C25","C30","C35","C40","Şap"]),("miktar_m3","Miktar (m³; ör. 30,5)","entry",None),("birim_fiyat","Birim Fiyat (TL/m³)","entry",None),("pompa_ucreti","Pompa Ücreti (TL)","entry",None),("odeme_durumu","Ödeme Durumu","combo",["Vadeli / Tahsil Edilmedi","Tahsil Edildi"]),("aciklama","Açıklama","entry",None)],
          "fatura":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("cari_id","Cari ID (boş bırakılabilir)","entry",None),("belge_no","Belge No","entry",None),("aciklama","Açıklama","entry",None),("miktar_m3","Miktar (m³)","entry",None),("birim_fiyat","Birim Fiyat","entry",None),("kdv_oran","KDV (%)","entry",None),("durum","Durum","combo",["Taslak","Hazır","İptal edildi"])],
          "nakit_kasa":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("kaynak","Kaynak","entry",None),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None)],
          "santiye":[("ad","Şantiye adı","entry",None),("adres","Adres","entry",None),("yetkili","Yetkili","entry",None),("telefon","Telefon","entry",None),("notlar","Notlar","entry",None)],
          "stok_hareket":[("tarih","Tarih","entry",None),("stok_id","Stok ID","entry",None),("yon","Yön","combo",["Giriş","Çıkış"]),("miktar","Miktar","entry",None),("aciklama","Açıklama","entry",None)],
          "maas_odeme":[("tarih","Tarih","entry",None),("personel_id","Personel ID","entry",None),("islem_turu","İşlem Türü","combo",["Maaş","Nakit Avans"]),("donem","Dönem (YYYY-AA)","entry",None),("tutar","Tutar","entry",None),("yontem","Ödeme yöntemi","entry",None),("aciklama","Açıklama","entry",None)],
          "banka_hareket":[("tarih","Tarih","entry",None),("banka","Banka","entry",None),("hesap","Hesap / IBAN sonu","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("aciklama","Açıklama","entry",None),("tutar","Tutar","entry",None)]}
        if table not in specs: return messagebox.showinfo("Bilgi","Bu listedeki düzenleme henüz desteklenmiyor.",parent=self)
        cols=[f[0] for f in specs[table]]
        with connect() as c: row=c.execute("SELECT "+",".join(cols)+f" FROM {table} WHERE id=?",(rid,)).fetchone()
        if not row: return messagebox.showerror("Kayıt bulunamadı","Seçilen kayıt artık mevcut değil.",parent=self)
        fields=[]
        initial=dict(zip(cols,row))
        for spec,val in zip(specs[table],row):
            key,label,kind,opts=spec
            if table=="beton_satis" and key=="santiye_id":
                with connect() as c: sites=c.execute("SELECT id,ad FROM santiye ORDER BY ad").fetchall()
                opts=[f"{sid} - {name}" for sid,name in sites]
                initial[key]="" if val is None else next((f"{sid} - {name}" for sid,name in sites if sid==val), "")
            if kind=="combo" and val not in (opts or []): opts=(opts or [])+[str(val or "")]
            fields.append((key,label,kind,opts))
        # Fatura toplamları alanları kullanıcı girişinden yeniden hesaplanır.
        d=self.prompt("Kayıt Düzenle | "+table,fields,initial)
        if not d: return
        try:
            vals=[]
            for key,_,_,_ in specs[table]:
                v=d[key]
                if table=="beton_satis" and key=="santiye_id":
                    v=self._resolve_site(v)
                elif key in ("cari_id","santiye_id","stok_id","personel_id"):
                    v=int(v) if str(v).strip() else None
                elif key in ("tutar","miktar","miktar_m3","birim_fiyat","pompa_ucreti","maas","kdv_oran"):
                    v=parse_amount(v) if str(v).strip() else 0
                vals.append(v)
            data=dict(zip(cols,vals))
            if "tarih" in data and data["tarih"]: datetime.strptime(data["tarih"],"%Y-%m-%d")
            if table=="beton_satis":
                data["miktar_m3"]=parse_amount(data["miktar_m3"])
                if data["miktar_m3"]<=0: raise ValueError("Metreküp sıfırdan büyük olmalı.")
                if data["birim_fiyat"] < 0: raise ValueError("Birim fiyat negatif olamaz.")
                if data["cari_id"] is not None:
                    with connect() as check_c: 
                        if not check_c.execute("SELECT 1 FROM cari WHERE id=?",(data["cari_id"],)).fetchone(): raise ValueError("Seçilen cari hesap bulunamadı.")
                data["toplam"]=round(data["miktar_m3"]*data["birim_fiyat"]+data.get("pompa_ucreti",0),2)
            if table=="fatura":
                data["miktar_m3"]=int(data["miktar_m3"] or 0); data["ara_toplam"]=round(data["miktar_m3"]*data["birim_fiyat"],2); data["kdv_tutar"]=round(data["ara_toplam"]*data["kdv_oran"]/100,2); data["genel_toplam"]=data["ara_toplam"]+data["kdv_tutar"]
            if table in ("stok","personel") and any(v is None for v in vals): raise ValueError("Zorunlu alanları kontrol edin.")
            for key in ("tutar","miktar","miktar_m3","birim_fiyat","maas"):
                if key in data and data[key] is not None and data[key] < 0: raise ValueError("Tutar, miktar ve fiyat negatif olamaz.")
            with connect() as c:
                if table=="stok_hareket":
                    old=c.execute("SELECT stok_id,yon,miktar FROM stok_hareket WHERE id=?",(rid,)).fetchone()
                    # Once the old movement is reversed, apply the edited movement to the new stock balance.
                    old_delta=old[2] if old[1]=="Giriş" else -old[2]
                    c.execute("UPDATE stok SET miktar=miktar-? WHERE id=?",(old_delta,old[0]))
                    current=c.execute("SELECT miktar FROM stok WHERE id=?",(data["stok_id"],)).fetchone()
                    if not current: raise ValueError("Stok kalemi bulunamadı.")
                    new_delta=data["miktar"] if data["yon"]=="Giriş" else -data["miktar"]
                    if current[0]+new_delta<0: raise ValueError("Bu düzenleme stok miktarını negatife düşürür.")
                    c.execute("UPDATE stok SET miktar=miktar+? WHERE id=?",(new_delta,data["stok_id"]))
                assignments=",".join(f"{col}=?" for col in cols)
                c.execute(f"UPDATE {table} SET {assignments} WHERE id=?",tuple(vals)+(rid,))
                if table=="beton_satis":
                    c.execute("UPDATE beton_satis SET toplam=? WHERE id=?",(data["toplam"],rid))
                    # Cari borç satırını satışın değişmez ID'sine göre tekilleştirip senkronla.
                    sync_beton_sale_cari_movement(c, rid)
                    ref=f"Beton satış tahsilatı #{rid}"
                    paid=c.execute("SELECT id FROM hareket WHERE aciklama=? AND kategori='Tahsilat' LIMIT 1",(ref,)).fetchone()
                    if data["odeme_durumu"]=="Tahsil Edildi" and data["cari_id"] and not paid:
                        c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(data["tarih"],data["cari_id"],"Alacak","Tahsilat",ref,data["toplam"],"Nakit"))
                        c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(data["tarih"],data["cari_id"],"Gelir","Satış",ref,data["toplam"],"Nakit"))
                        c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",(data["tarih"],"Para Girişi","Beton Satışı",ref,data["toplam"],getattr(self,"current_user",{}).get("username","")))
                    elif data["odeme_durumu"]!="Tahsil Edildi" and paid:
                        c.execute("DELETE FROM hareket WHERE aciklama=? AND kategori IN ('Tahsilat','Satış')",(ref,))
                        c.execute("DELETE FROM nakit_kasa WHERE aciklama=?",(ref,))
                        c.execute("DELETE FROM banka_hareket WHERE aciklama=?",(ref,))
                        c.execute("DELETE FROM cek WHERE notlar=?",(ref,))
                    elif paid:
                        c.execute("UPDATE hareket SET tarih=?,cari_id=?,tutar=? WHERE aciklama=? AND kategori IN ('Tahsilat','Satış')",(data["tarih"],data["cari_id"],data["toplam"],ref))
                        c.execute("UPDATE nakit_kasa SET tarih=?,tutar=? WHERE aciklama=?",(data["tarih"],data["toplam"],ref))
                        c.execute("UPDATE banka_hareket SET tarih=?,tutar=? WHERE aciklama=?",(data["tarih"],data["toplam"],ref))
                        c.execute("UPDATE cek SET tutar=? WHERE notlar=?",(data["toplam"],ref))
                if table=="fatura": c.execute("UPDATE fatura SET ara_toplam=?,kdv_tutar=?,genel_toplam=? WHERE id=?",(data["ara_toplam"],data["kdv_tutar"],data["genel_toplam"],rid))
            self.refresh_all()
        except Exception as e: messagebox.showerror("Düzenleme başarısız",f"Kayıt güncellenemedi. Bilgileri kontrol edin.\n\n{e}",parent=self)

    def _build_extra_tabs(self):
        # Stok hareketleri
        bar=ttk.Frame(self.stok_hareket_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Stok Giriş / Çıkış Hareketleri",font=("Segoe UI",15,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Hareket",command=self.add_stok_hareket).pack(side="right")
        self.stok_hareket_tree=self._tree(self.stok_hareket_tab,("id","tarih","stok","yon","miktar","aciklama"),("ID","Tarih","Stok","Yön","Miktar","Açıklama"),(45,110,220,100,120,300))
        row=ttk.Frame(self.stok_hareket_tab); row.pack(fill="x",pady=6)
        ttk.Button(row,text="Düzenle",command=lambda:self.edit_selected(self.stok_hareket_tree,"stok_hareket")).pack(side="left")
        ttk.Button(row,text="Sil",command=lambda:self.delete_selected(self.stok_hareket_tree,"stok_hareket")).pack(side="right")
        # Maaş ödemeleri
        bar=ttk.Frame(self.maas_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Personel Maaş Ödemeleri",font=("Segoe UI",15,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Maaş Ödemesi",command=self.add_maas_odeme).pack(side="right")
        self.maas_tree=self._tree(self.maas_tab,("id","tarih","personel","islem_turu","donem","tutar","yontem","aciklama"),("ID","Tarih","Personel","İşlem Türü","Dönem","Tutar TL","Yöntem","Açıklama"),(45,110,200,120,100,120,120,280))
        row=ttk.Frame(self.maas_tab); row.pack(fill="x",pady=6)
        ttk.Button(row,text="Düzenle",command=lambda:self.edit_selected(self.maas_tree,"maas_odeme")).pack(side="left")
        ttk.Button(row,text="Sil",command=lambda:self.delete_selected(self.maas_tree,"maas_odeme")).pack(side="right")
        # Banka hareketleri
        bar=ttk.Frame(self.banka_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Banka Giriş / Çıkış İşlemleri",font=("Segoe UI",15,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Banka İşlemi",command=self.add_banka_hareket).pack(side="right")
        self.banka_tree=self._tree(self.banka_tab,("id","tarih","banka","hesap","yon","aciklama","tutar"),("ID","Tarih","Banka","Hesap","İşlem","Açıklama","Tutar TL"),(45,110,180,140,110,300,130))
        self.banka_ozet=ttk.Label(self.banka_tab,text="Banka net hareketi: 0,00 TL",font=("Segoe UI",11,"bold")); self.banka_ozet.pack(anchor="w",pady=5)
        row=ttk.Frame(self.banka_tab); row.pack(fill="x",pady=6)
        ttk.Button(row,text="Düzenle",command=lambda:self.edit_selected(self.banka_tree,"banka_hareket")).pack(side="left")
        ttk.Button(row,text="Sil",command=lambda:self.delete_selected(self.banka_tree,"banka_hareket")).pack(side="right")

    def add_stok_hareket(self):
        with connect() as c: items=c.execute("SELECT id,ad FROM stok ORDER BY ad").fetchall()
        if not items: return messagebox.showinfo("Stok yok","Önce Stok / Depo sekmesinden stok kalemi ekleyin.",parent=self)
        d=self.prompt("Stok Hareketi",[("tarih","Tarih (GG.AA.YYYY)","entry",None),("stok","Stok kalemi","combo",[f"{i} - {n}" for i,n in items]),("yon","Yön","combo",["Giriş","Çıkış"]),("miktar","Miktar","entry",None),("aciklama","Açıklama","entry",None)])
        if not d:return
        try:
            dt=d["tarih"] or date.today().isoformat(); datetime.strptime(dt,"%Y-%m-%d"); sid=int(d["stok"].split(" - ",1)[0]); qty=parse_amount(d["miktar"])
            if qty<=0: raise ValueError("Miktar sıfırdan büyük olmalı.")
            with connect() as c:
                current=c.execute("SELECT miktar FROM stok WHERE id=?",(sid,)).fetchone()[0]
                new=current+qty if d["yon"]=="Giriş" else current-qty
                if new<0: raise ValueError("Çıkış miktarı mevcut stoktan fazla olamaz.")
                c.execute("INSERT INTO stok_hareket(tarih,stok_id,yon,miktar,aciklama) VALUES(?,?,?,?,?)",(dt,sid,d["yon"],qty,d["aciklama"]))
                c.execute("UPDATE stok SET miktar=? WHERE id=?",(new,sid))
            self.refresh_all()
        except Exception as e: messagebox.showerror("Stok hareketi kaydedilemedi",str(e),parent=self)

    def add_maas_odeme(self):
        with connect() as c: items=c.execute("SELECT id,ad FROM personel ORDER BY ad").fetchall()
        if not items:return messagebox.showinfo("Personel yok","Önce Personel sekmesinden personel ekleyin.",parent=self)
        d=self.prompt("Maaş / Nakit Avans Ödemesi",[("tarih","Ödeme tarihi","entry",None),("personel","Personel","combo",[f"{i} - {n}" for i,n in items]),("islem_turu","İşlem türü","combo",["Maaş","Nakit Avans"]),("donem","Dönem (YYYY-AA)","entry",None),("tutar","Tutar (TL)","entry",None),("yontem","Ödeme yöntemi","combo",["Banka","Nakit","Diğer"]),("aciklama","Açıklama","entry",None)])
        if not d:return
        try:
            dt=d["tarih"] or date.today().isoformat(); datetime.strptime(dt,"%Y-%m-%d"); pid=int(d["personel"].split(" - ",1)[0]); amount=parse_amount(d["tutar"])
            if amount<=0:raise ValueError("Tutar sıfırdan büyük olmalı.")
            if d["donem"]: datetime.strptime(d["donem"]+"-01","%Y-%m-%d")
            with connect() as c:c.execute("INSERT INTO maas_odeme(tarih,personel_id,donem,tutar,yontem,aciklama,islem_turu) VALUES(?,?,?,?,?,?,?)",(dt,pid,d["donem"],amount,d["yontem"],d["aciklama"],d["islem_turu"]))
            self.refresh_all()
        except Exception as e:messagebox.showerror("Maaş ödemesi kaydedilemedi",str(e),parent=self)

    def add_banka_hareket(self):
        d=self.prompt("Banka İşlemi",[("tarih","Tarih (GG.AA.YYYY)","entry",None),("banka","Banka adı","entry",None),("hesap","Hesap / IBAN sonu","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None)])
        if not d:return
        try:
            dt=d["tarih"] or date.today().isoformat(); datetime.strptime(dt,"%Y-%m-%d"); amount=parse_amount(d["tutar"])
            if not d["banka"].strip() or amount<=0:raise ValueError("Banka adı gerekli, tutar sıfırdan büyük olmalı.")
            with connect() as c:c.execute("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",(dt,d["banka"],d["hesap"],d["yon"],d["aciklama"],amount))
            self.refresh_all()
        except Exception as e:messagebox.showerror("Banka işlemi kaydedilemedi",str(e),parent=self)

    def add_cari(self):
        d = self.prompt("Yeni Cari", [("tur","Tür","combo",["Müşteri","Tedarikçi","Diğer"]),
            ("unvan","Ünvan / Ad Soyad","entry",None),("telefon","Telefon","entry",None),
            ("vergi_no","Vergi / T.C. No","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        if not d["unvan"]: return messagebox.showwarning("Eksik bilgi","Ünvan zorunludur.")
        with connect() as c:
            c.execute("INSERT INTO cari(tur,unvan,telefon,vergi_no,notlar,created_at) VALUES(?,?,?,?,?,?)",
                (d["tur"],d["unvan"],d["telefon"],d["vergi_no"],d["notlar"],datetime.now().isoformat(timespec="seconds")))
        self.refresh_all()

    def show_audit_log(self):
        if getattr(self, "current_user", {}).get("role") != "Yönetici":
            return messagebox.showwarning("Yetki gerekli", "İşlem geçmişini yalnızca yönetici görüntüleyebilir.", parent=self)
        win=tk.Toplevel(self); win.title("HOYTAŞ | İşlem Geçmişi"); win.geometry("1120x560"); win.transient(self)
        frm=ttk.Frame(win,padding=14); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Kayıt ve İşlem Denetim Geçmişi",font=("Segoe UI",15,"bold")).pack(anchor="w")
        ttk.Label(frm,text="Cari, çek, beton satışı, fatura, kasa, banka, stok ve personel kayıtlarındaki ekleme/değiştirme/silme işlemleri listelenir. Şifreler bu kayda alınmaz.",wraplength=1050).pack(anchor="w",pady=(4,10))
        bar=ttk.Frame(frm); bar.pack(fill="x",pady=(0,6))
        ttk.Label(bar,text="Ara:").pack(side="left")
        q=tk.StringVar(); ttk.Entry(bar,textvariable=q,width=38).pack(side="left",padx=6)
        cols=("tarih","kullanici","tablo","islem","kayit","detay")
        tree=ttk.Treeview(frm,columns=cols,show="headings")
        for col,label,width in [("tarih","Tarih / Saat",150),("kullanici","Kullanıcı",130),("tablo","Bölüm",130),("islem","İşlem",90),("kayit","Kayıt ID",75),("detay","Değişiklik Özeti",500)]:
            tree.heading(col,text=label); tree.column(col,width=width,anchor="w")
        tree.pack(fill="both",expand=True)
        def reload_rows(*_):
            tree.delete(*tree.get_children())
            term=q.get().strip().casefold()
            with connect() as c:
                rows=c.execute("SELECT changed_at,actor,table_name,action,record_id,COALESCE(old_values,''),COALESCE(new_values,'') FROM audit_log ORDER BY id DESC LIMIT 1000").fetchall()
            for r in rows:
                detail=(r[5]+(" -> " if r[5] and r[6] else "")+r[6]).replace("\n"," ")
                vals=(r[0],r[1],r[2],r[3],r[4],detail)
                if not term or term in " ".join(map(str,vals)).casefold(): tree.insert("","end",values=vals)
        q.trace_add("write",reload_rows); reload_rows()
        ttk.Label(frm,text="Son 1.000 işlem gösterilir. Denetim kayıtları normal kullanıcı ekranlarından silinemez.",foreground="#66736B").pack(anchor="w",pady=(7,0))

    def manage_users(self):
        win=tk.Toplevel(self); win.title("Kullanıcı Yönetimi"); win.transient(self); win.grab_set(); win.geometry("520x390")
        frm=ttk.Frame(win,padding=16); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Kullanıcı Hesapları",font=("Segoe UI",15,"bold")).pack(anchor="w")
        tree=ttk.Treeview(frm,columns=("id","user","role","active"),show="headings",height=7)
        for col,label,width in [("id","ID",45),("user","Kullanıcı adı",190),("role","Rol",120),("active","Durum",90)]: tree.heading(col,text=label); tree.column(col,width=width)
        tree.pack(fill="both",expand=True,pady=8)
        def reload_users():
            tree.delete(*tree.get_children())
            with connect() as c: rows=c.execute("SELECT id,kullanici_adi,rol,aktif FROM kullanici ORDER BY id").fetchall()
            for row in rows: tree.insert("","end",values=(row[0],row[1],row[2],"Aktif" if row[3] else "Pasif"))
        def add_user():
            d=self.prompt("Yeni Kullanıcı",[("user","Kullanıcı adı","entry",None),("password","Şifre (en az 8 karakter)","entry",None),("role","Rol","combo",["Kullanıcı","Yönetici"])])
            if not d:return
            if not d["user"] or len(d["password"])<8:return messagebox.showerror("Eksik bilgi","Kullanıcı adı gerekli; şifre en az 8 karakter olmalı.",parent=win)
            salt=secrets.token_hex(16)
            try:
                with connect() as c: c.execute("INSERT INTO kullanici(kullanici_adi,parola_hash,salt,rol) VALUES(?,?,?,?)",(d["user"],self._hash_password(d["password"],salt),salt,d["role"]))
            except sqlite3.IntegrityError:return messagebox.showerror("Kullanıcı mevcut","Bu kullanıcı adı zaten kayıtlı.",parent=win)
            reload_users()
        def toggle_user():
            sel=tree.selection()
            if not sel:return messagebox.showinfo("Kayıt seçin","Önce kullanıcı seçin.",parent=win)
            vals=tree.item(sel[0],"values"); uid=int(vals[0])
            if uid==getattr(self,"current_user",{}).get("id"):return messagebox.showwarning("İşlem engellendi","Kendi hesabınızı bu ekrandan devre dışı bırakamazsınız.",parent=win)
            new_active=0 if vals[3]=="Aktif" else 1
            with connect() as c:c.execute("UPDATE kullanici SET aktif=? WHERE id=?",(new_active,uid))
            reload_users()
        row=ttk.Frame(frm); row.pack(fill="x")
        ttk.Button(row,text="Yeni Kullanıcı Ekle",command=add_user).pack(side="left")
        ttk.Button(row,text="Aktif / Pasif Yap",command=toggle_user).pack(side="left",padx=8)
        reload_users()

    def add_santiye(self):
        d=self.prompt("Yeni Şantiye",[("ad","Şantiye adı","entry",None),("adres","Adres","entry",None),("yetkili","Yetkili kişi","entry",None),("telefon","Telefon","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        if not d["ad"]: return messagebox.showwarning("Eksik bilgi","Şantiye adı zorunludur.")
        with connect() as c: c.execute("INSERT INTO santiye(ad,adres,yetkili,telefon,notlar) VALUES(?,?,?,?,?)",(d["ad"],d["adres"],d["yetkili"],d["telefon"],d["notlar"]))
        self.refresh_all()

    def add_fatura(self):
        """Yerel taslak: m³, KDV ve seçilen tevkifat oranından net tahsilatı hesapla."""
        with connect() as c:
            cariler = c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        if not cariler:
            return messagebox.showinfo("Cari gerekli", "Önce cari hesap ekleyin.", parent=self)
        names = [f"{i} - {n}" for i,n in cariler]
        d = self.prompt("Yeni Satış Faturası Taslağı", [
            ("tarih", "Fatura tarihi (GG.AA.YYYY)", "entry", None),
            ("cari", "Faturanın kesileceği cari", "combo_search", names),
            ("belge", "Fatura / belge numarası", "entry", None),
            ("aciklama", "Beton cinsi, şantiye ve teslim açıklaması", "entry", None),
            ("miktar", "Teslim edilen beton (m³; ör. 30,5)", "entry", None),
            ("fiyat", "KDV hariç birim fiyat (TL/m³)", "entry", None),
            ("kdv", "KDV oranı (%)", "combo", ["20","10","1","0"]),
            ("tur", "Fatura türü", "combo", ["Normal", "Tevkifatlı"]),
            ("tevkifat", "KDV tevkifat oranı (ör. 4/10, 5/10; normalde 0)", "combo_edit", ["0","4/10","5/10","7/10","9/10"])
        ])
        if not d:
            return
        try:
            tarih = d["tarih"] or date.today().isoformat()
            datetime.strptime(tarih, "%Y-%m-%d")
            miktar = parse_amount(d["miktar"])
            fiyat = parse_amount(d["fiyat"])
            kdv_oran = parse_amount(d["kdv"])
            if miktar <= 0 or fiyat < 0 or not 0 <= kdv_oran <= 100:
                raise ValueError("Miktar, fiyat veya KDV oranı geçersiz.")
            tur = d["tur"]
            oran_metin = d["tevkifat"].strip()
            if tur == "Tevkifatlı":
                if "/" not in oran_metin:
                    raise ValueError("Tevkifat oranını 4/10 gibi girin.")
                pay, payda = [int(x.strip()) for x in oran_metin.split("/", 1)]
                if payda != 10 or not 0 < pay < payda:
                    raise ValueError("Tevkifat oranı 1/10 ile 9/10 arasında olmalıdır.")
                tevkifat_oran = pay / payda
            else:
                tevkifat_oran = 0.0
            cari_id = int(d["cari"].split(" - ", 1)[0])
            ara = round(miktar * fiyat, 2)
            kdv_tutar = round(ara * kdv_oran / 100, 2)
            toplam = round(ara + kdv_tutar, 2)
            tevkifat_tutar = round(kdv_tutar * tevkifat_oran, 2)
            net = round(toplam - tevkifat_tutar, 2)
            if not d["belge"].strip():
                raise ValueError("Belge numarası boş bırakılamaz.")
            with connect() as c:
                if c.execute("SELECT 1 FROM fatura WHERE cari_id=? AND belge_no=?",
                             (cari_id, d["belge"].strip())).fetchone():
                    raise ValueError("Bu cariye ait aynı belge numarası zaten kayıtlı.")
                c.execute("""INSERT INTO fatura
                    (tarih,cari_id,belge_no,aciklama,miktar_m3,birim_fiyat,ara_toplam,kdv_oran,kdv_tutar,
                     genel_toplam,durum,fatura_turu,tevkifat_oran,tevkifat_tutar,tahsil_edilecek,odenen)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
                    (tarih,cari_id,d["belge"].strip(),d["aciklama"],miktar,fiyat,ara,kdv_oran,
                     kdv_tutar,toplam,"Taslak",tur,tevkifat_oran,tevkifat_tutar,net))
            self.refresh_all()
            messagebox.showinfo("Fatura taslağı kaydedildi",
                f"Ara toplam: {ara:,.2f} TL\nKDV: {kdv_tutar:,.2f} TL\n"
                f"Genel toplam: {toplam:,.2f} TL\nTevkifat: {tevkifat_tutar:,.2f} TL\n"
                f"Tahsil edilecek: {net:,.2f} TL\n\nBu kayıt resmî e-Fatura gönderimi değildir.",
                parent=self)
        except (ValueError, TypeError) as exc:
            messagebox.showerror("Fatura kaydedilemedi", str(exc), parent=self)

    def tahsilat_fatura(self):
        selected = self.fatura_tree.selection()
        if not selected:
            return messagebox.showinfo("Fatura seçin", "Önce bir fatura seçin.", parent=self)
        fid = int(self.fatura_tree.item(selected[0], "values")[0])
        with connect() as c:
            row = c.execute("""SELECT f.cari_id,f.belge_no,f.tahsil_edilecek,f.odenen,
                               COALESCE(c.unvan,'') FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id
                               WHERE f.id=?""", (fid,)).fetchone()
        if not row:
            return
        kalan = round(row[2] - row[3], 2)
        if kalan <= 0:
            return messagebox.showinfo("Tahsil edildi", "Bu faturanın kalan tutarı bulunmuyor.", parent=self)
        d = self.prompt(f"Fatura Tahsilatı - {row[1]}", [
            ("tarih", "Tahsilat tarihi (GG.AA.YYYY)", "entry", None),
            ("tutar", f"Alınan tutar (kalan {kalan:,.2f} TL)", "entry", None),
            ("yontem", "Tahsilat yöntemi", "combo", ["Nakit","Banka","Çek","Diğer"]),
            ("aciklama", "Tahsilat açıklaması", "entry", None)
        ])
        if not d:
            return
        try:
            tarih = d["tarih"] or date.today().isoformat()
            datetime.strptime(tarih, "%Y-%m-%d")
            tutar = parse_amount(d["tutar"])
            if not 0 < tutar <= kalan:
                raise ValueError("Tahsilat tutarı sıfırdan büyük ve kalan borçtan fazla olmamalıdır.")
            with connect() as c:
                c.execute("INSERT INTO fatura_tahsilat(fatura_id,tarih,tutar,yontem,aciklama) VALUES(?,?,?,?,?)",
                          (fid,tarih,tutar,d["yontem"],d["aciklama"]))
                c.execute("UPDATE fatura SET odenen=odenen+? WHERE id=?", (tutar,fid))
                # Fatura taslakları cari borcu ikinci kez artırmaz; cari tahsilat tek defa işlenir.
                c.execute("""INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi)
                             VALUES(?,?,?,?,?,?,?)""",
                          (tarih,row[0],"Alacak","Fatura Tahsilatı",f"Fatura #{fid} / {row[1]} - {d['aciklama']}",tutar,d["yontem"]))
                if d["yontem"] == "Nakit":
                    c.execute("""INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici)
                                 VALUES(?,?,?,?,?,?)""",
                              (tarih,"Para Girişi","Fatura Tahsilatı",f"Fatura #{fid} - {row[4]}",tutar,
                               getattr(self,"current_user",{}).get("username","")))
            self.refresh_all()
            messagebox.showinfo("Tahsilat kaydedildi",
                f"Tahsil edilen: {tutar:,.2f} TL\nKalan: {kalan-tutar:,.2f} TL",parent=self)
        except (ValueError,TypeError) as exc:
            messagebox.showerror("Tahsilat kaydedilemedi",str(exc),parent=self)

    def cari_tahsilat(self):
        """Seçili müşteriden alınan kısmi/tam ödemeyi cari ve para hareketlerine işler."""
        sel = self.cari_tree.selection()
        if not sel:
            return messagebox.showinfo("Cari seçin", "Önce tahsilat alınacak müşteriyi seçin.", parent=self)
        vals = self.cari_tree.item(sel[0], "values")
        cari_id, cari_unvan = int(vals[0]), str(vals[2])
        d = self.prompt(f"Cari Tahsilat - {cari_unvan}", [
            ("tarih", "Tahsilat tarihi (GG.AA.YYYY)", "entry", None),
            ("tutar", "Alınan tutar (TL)", "entry", None),
            ("yontem", "Ödeme şekli", "combo", ["Nakit", "Banka Havalesi / EFT", "Çek"]),
            ("banka", "Banka adı (banka/çek için)", "entry", None),
            ("cek_no", "Çek numarası (çek ise)", "entry", None),
            ("vade", "Çek vadesi GG.AA.YYYY (çek ise)", "entry", None),
            ("aciklama", "Açıklama", "entry", None)])
        if not d: return
        try:
            tarih = d["tarih"] or datetime.now().strftime("%Y-%m-%d")
            datetime.strptime(tarih, "%Y-%m-%d")
            tutar = parse_amount(d["tutar"])
            if tutar <= 0: raise ValueError("Tutar sıfırdan büyük olmalı.")
            if d["yontem"] == "Çek":
                vade = d["vade"] or tarih
                datetime.strptime(vade, "%Y-%m-%d")
        except ValueError:
            return messagebox.showerror("Hatalı bilgi", "Tarih/vade GG.AA.YYYY biçiminde, tutar sıfırdan büyük olmalıdır.", parent=self)
        aciklama = (d["aciklama"] or f"{cari_unvan} cari tahsilatı")
        with connect() as c:
            # Tahsilat, cari borcunu yalnızca alınan tutar kadar azaltır.
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                      (tarih, cari_id, "Alacak", "Tahsilat", aciklama, tutar, d["yontem"]))
            # Tahsilat gelir hareketi olarak raporlanır; satış borcu ayrıca korunur.
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                      (tarih, cari_id, "Gelir", "Tahsilat", aciklama, tutar, d["yontem"]))
            if d["yontem"] == "Nakit":
                c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                          (tarih, "Para Girişi", "Cari Tahsilat", f"{cari_unvan} - {aciklama}", tutar, getattr(self, "current_user", {}).get("username", "")))
            elif d["yontem"] == "Banka Havalesi / EFT":
                c.execute("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                          (tarih, d["banka"] or "Banka", "", "Para Girişi", f"{cari_unvan} - {aciklama}", tutar))
            else:
                cek_no = d["cek_no"] or f"TAH-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                c.execute("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",
                          ("Alınan", d["banka"] or "", "", cari_id, cek_no, cari_unvan, d["vade"] or tarih, tutar, "Bekliyor", aciklama))
        self.refresh_all()
        messagebox.showinfo("Tahsilat kaydedildi", f"{tutar:,.2f} TL tahsilat kaydedildi. Cari bakiye alınan tutar kadar düşürüldü.", parent=self)

    def refresh_cari(self):
        if not hasattr(self,"cari_tree"): return
        term=normalize_search_text(self.cari_search.get()) if hasattr(self,"cari_search") else ""
        with connect() as c:
            raw=c.execute("""SELECT c.id,c.tur,c.unvan,c.telefon,c.vergi_no,
                b.toplam_borc,b.toplam_alacak,b.bakiye
                FROM cari c JOIN cari_bakiye b ON b.cari_id=c.id
                ORDER BY c.unvan COLLATE NOCASE""").fetchall()
        rows=[]; row_tags={}
        for r in raw:
            borc, alacak, bakiye = safe_float(r[5] or 0), safe_float(r[6] or 0), safe_float(r[7] or 0)
            rows.append((r[0],r[1],r[2],r[3],r[4],f"{borc:.2f}",f"{alacak:.2f}",f"{bakiye:.2f}"))
            row_tags[str(r[0])] = ("debt_partial" if alacak > 0.005 else "debt_unpaid") if bakiye > 0.005 else "debt_closed"
        self.cari_tree._row_tags=row_tags
        if term: rows=[r for r in rows if term in normalize_search_text(" ".join(str(x or "") for x in r[:5]))]
        self.fill_tree(self.cari_tree,rows)

    def add_hareket(self):
        with connect() as c: cariler = c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        names = ["(Cari seçilmedi)"] + [f"{i} - {n}" for i,n in cariler]
        d = self.prompt("Yeni Gelir / Gider", [("tarih","Tarih (GG.AA.YYYY)","entry",None),
            ("tur","Tür","combo",["Gelir","Gider","Borç","Alacak"]),("kategori","Kategori","combo",["Satış","Alış","Tahsilat","Ödeme","Maaş","Yakıt","Kira","Diğer"]),
            ("cari","Cari","combo",names),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None),
            ("odeme","Ödeme yöntemi","combo",["Nakit","Banka","Kredi Kartı","Çek","Diğer"])])
        if not d: return
        try:
            tarih = d["tarih"] or datetime.now().strftime("%Y-%m-%d")
            datetime.strptime(tarih,"%Y-%m-%d")
            tutar = parse_amount(d["tutar"])
            if tutar <= 0: raise ValueError()
            cari_id = int(d["cari"].split(" - ",1)[0]) if " - " in d["cari"] else None
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Tarih GG.AA.YYYY biçiminde, tutar ise sıfırdan büyük sayı olmalıdır.")
        with connect() as c:
            c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,d["tur"],d["kategori"],d["aciklama"],tutar,d["odeme"]))
        self.refresh_all()

    def add_cek(self):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        names=["(Cari seçilmedi)"]+[f"{i} - {n}" for i,n in cariler]
        banks=["010 - T.C. Ziraat Bankası","012 - Türkiye Halk Bankası","015 - Türkiye Vakıflar Bankası","032 - Türk Ekonomi Bankası","046 - Akbank","062 - Türkiye Garanti Bankası","064 - Türkiye İş Bankası","067 - Yapı ve Kredi Bankası","111 - QNB Bank","134 - DenizBank","203 - Albaraka Türk","206 - Türkiye Finans","208 - Kuveyt Türk","210 - Vakıf Katılım","211 - Türkiye Emlak Katılım","214 - Şekerbank","223 - Alternatif Bank","999 - Diğer / listede yok"]
        d=self.prompt("Yeni Çek",[("yon","Yön","combo",["Alınan","Verilen"]),("banka","Banka adı ve kodu","combo",banks), ("cari","İlgili cari hesap","combo",names),
            ("cek_no","Çek No","entry",None),("taraf","Keşideci / Lehtar","entry",None),("vade","Vade (GG.AA.YYYY)","entry",None),("tutar","Tutar (TL)","entry",None),
            ("durum","Durum","combo",["Bekliyor","Tahsil Edildi","Ödendi","İade","Karşılıksız"]),("notlar","Notlar","entry",None)])
        if not d:return
        try:
            datetime.strptime(d["vade"],"%Y-%m-%d"); tutar=parse_amount(d["tutar"]);
            if tutar<=0: raise ValueError
            cari_id=int(d["cari"].split(" - ",1)[0]) if " - " in d["cari"] else None
            banka_kodu,banka=(d["banka"].split(" - ",1)+[""])[:2] if " - " in d["banka"] else ("",d["banka"])
        except ValueError:return messagebox.showerror("Hatalı bilgi","Vade GG.AA.YYYY biçiminde ve tutar sıfırdan büyük olmalıdır.")
        with connect() as c:
            c.execute("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",(d["yon"],banka,banka_kodu,cari_id,d["cek_no"],d["taraf"],d["vade"],tutar,d["durum"],d["notlar"]))
            if cari_id:
                # Alınan çek müşteri alacağını tahsilat gibi azaltır; verilen çek
                # ise cari borcu/alacak yönünde ters etki oluşturur.
                move_type = "Alacak" if d["yon"] == "Alınan" else "Borç"
                c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(datetime.now().strftime("%Y-%m-%d"),cari_id,move_type,"Çek",f"{d['yon']} çek #{d['cek_no']} - {banka}",tutar,"Çek"))
        self.refresh_all()

    def add_stok(self):
        d = self.prompt("Stok Kalemi", [("kod","Stok Kodu","entry",None),("ad","Malzeme Adı","entry",None),
            ("birim","Birim","combo",["Adet","Kg","Ton","m³","Litre","Metre","Paket"]),
            ("miktar","Miktar","entry",None),("fiyat","Birim Fiyat (TL)","entry",None)])
        if not d: return
        try:
            miktar=parse_amount(d["miktar"] or "0")
            fiyat=parse_amount(d["fiyat"] or "0")
            if not d["ad"] or miktar < 0 or fiyat < 0: raise ValueError()
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Malzeme adı gerekli; miktar ve fiyat negatif olamaz.")
        with connect() as c:
            c.execute("INSERT INTO stok(kod,ad,birim,miktar,birim_fiyat) VALUES(?,?,?,?,?)",
                      (d["kod"],d["ad"],d["birim"],miktar,fiyat))
        self.refresh_all()

    def add_personel(self):
        d = self.prompt("Personel Ekle", [("ad","Ad Soyad","entry",None),("gorev","Görev","entry",None),
            ("telefon","Telefon","entry",None),("maas","Aylık Maaş (TL)","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        try:
            maas=parse_amount(d["maas"] or "0")
            if not d["ad"] or maas < 0: raise ValueError()
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Ad soyad gerekli; maaş negatif olamaz.")
        with connect() as c:
            c.execute("INSERT INTO personel(ad,gorev,telefon,maas,notlar) VALUES(?,?,?,?,?)",
                      (d["ad"],d["gorev"],d["telefon"],maas,d["notlar"]))
        self.refresh_all()

    def add_beton(self):
        with connect() as c: cariler = c.execute("SELECT id,unvan FROM cari WHERE tur='Müşteri' OR tur='Diğer' ORDER BY unvan").fetchall()
        names = ["(Cari seçilmedi)"] + [f"{i} - {n}" for i,n in cariler]
        with connect() as c: sites=c.execute("SELECT id,ad FROM santiye ORDER BY ad").fetchall()
        site_names=[f"{i} - {n}" for i,n in sites]
        site_names += [n for n in ["Adıyaman", "Besni", "Tut", "Gölbaşı", "Erkenek", "Erkenek Köyü"] if n.casefold() not in [x.casefold() for x in site_names]]
        d = self.prompt("Yeni Beton Satışı", [("tarih","Tarih (GG.AA.YYYY)","entry",None),
            ("cari","Müşteri ara (ünvan yazın) / seçin","combo_search",names),("santiye","Gideceği şantiye (elle yazabilirsiniz)","combo_edit",site_names),("beton_cinsi","Beton Cinsi","combo",["C16","C20","C25","C30","C35","C40","Şap"]),("m3","Miktar (m³; ör. 30,5)","entry",None),
            ("fiyat","Birim Fiyat (TL/m³)","entry",None),
            ("pompa_ucreti","Pompa Ücreti (TL, yoksa 0)","entry",None),
            ("durum","Ödeme Durumu","combo",["Vadeli / Tahsil Edilmedi","Tahsil Edildi"]),
            ("odeme_yontemi","Tahsilat yöntemi (ödendiyse)","combo",["Nakit","Banka Havalesi / EFT","Çek"]),
            ("banka","Banka adı (banka/çek ise)","entry",None),
            ("cek_no","Çek numarası (çek ise)","entry",None),
            ("vade","Çek vadesi GG.AA.YYYY (çek ise)","entry",None),
            ("aciklama","Açıklama","entry",None)])
        if not d: return
        try:
            tarih=d["tarih"] or datetime.now().strftime("%Y-%m-%d")
            datetime.strptime(tarih,"%Y-%m-%d")
            miktar,fiyat,pompa_ucreti,toplam=concrete_sale_totals(d["m3"],d.get("fiyat"),d.get("pompa_ucreti"))
            cari_value = (d.get("cari") or "").strip()
            with connect() as lookup_db:
                cari_id = resolve_customer_id(lookup_db, cari_value)
            santiye_id=self._resolve_site(d["santiye"])
            if d["durum"] == "Tahsil Edildi" and d.get("odeme_yontemi") == "Çek":
                datetime.strptime(d.get("vade") or tarih,"%Y-%m-%d")
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Tarih GG.AA.YYYY olmalı; metreküp sıfırdan büyük olmalı. Fiyatı boş bırakabilirsiniz (0 TL kaydedilir); fiyat ve pompa ücreti negatif olamaz. 30,5 gibi küsurat girebilirsiniz.")
        with connect() as c:
            cur=c.execute("INSERT INTO beton_satis(tarih,cari_id,santiye_id,beton_cinsi,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu,aciklama) VALUES(?,?,?,?,?,?,?,?,?,?)",
                (tarih,cari_id,santiye_id,d["beton_cinsi"],miktar,fiyat,pompa_ucreti,toplam,d["durum"],d["aciklama"]))
            sale_id=cur.lastrowid
            # Her satış için cari bağlantısını tek ortak fonksiyon üzerinden oluştur.
            sync_beton_sale_cari_movement(c, sale_id)
            if cari_id and d["durum"] == "Tahsil Edildi":
                method=d.get("odeme_yontemi") or "Nakit"
                ref=f"Beton satış tahsilatı #{sale_id}"
                c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                    (tarih,cari_id,"Alacak","Tahsilat",ref,toplam,method))
                c.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                    (tarih,cari_id,"Gelir","Satış",ref,toplam,method))
                if method == "Nakit":
                    c.execute("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                        (tarih,"Para Girişi","Beton Satışı",ref,toplam,getattr(self,"current_user",{}).get("username","")))
                elif method == "Banka Havalesi / EFT":
                    c.execute("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                        (tarih,d.get("banka") or "Banka","","Para Girişi",ref,toplam))
                else:
                    vade=d.get("vade") or tarih
                    datetime.strptime(vade,"%Y-%m-%d")
                    cek_no=d.get("cek_no") or f"SAT-{sale_id}"
                    c.execute("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        ("Alınan",d.get("banka") or "","",cari_id,cek_no,"",vade,toplam,"Bekliyor",ref))
        self.refresh_all()

    def delete_selected(self, tree, table):
        sel = tree.selection()
        if not sel: return messagebox.showinfo("Kayıt seçin","Önce silmek istediğiniz satırı seçin.", parent=self)
        row_id = int(tree.item(sel[0],"values")[0])
        if table == "hareket":
            with connect() as c:
                linked = c.execute("SELECT aciklama FROM hareket WHERE id=?", (row_id,)).fetchone()
            sale_id = linked_beton_sale_id(linked[0] if linked else "")
            if sale_id:
                return messagebox.showinfo(
                    "Bağlı kayıt korunuyor",
                    f"Bu cari hareketi Beton Satış #{sale_id} kaydının parçasıdır. "
                    "Bağlantının bozulmaması için buradan silinemez. Satışı Beton Satış sayfasından silin.", parent=self)
        if table == "cari":
            with connect() as c:
                cari = c.execute("SELECT unvan FROM cari WHERE id=?", (row_id,)).fetchone()
                if not cari: return messagebox.showinfo("Bulunamadı", "Cari hesap zaten silinmiş.", parent=self)
                counts = {t:c.execute(f"SELECT COUNT(*) FROM {t} WHERE cari_id=?",(row_id,)).fetchone()[0] for t in ("hareket","cek","beton_satis","fatura","gelen_fatura")}
            summary = ", ".join(f"{k}: {v}" for k,v in counts.items() if v)
            if not messagebox.askyesno("Cari ve tüm kayıtları sil", f"'{cari[0]}' cari hesabı ve bu hesaba bağlı TÜM hareketler, çekler, beton satışları ve faturalar silinecek.\n\n{summary or 'Bağlı kayıt bulunamadı.'}\n\nİşlem öncesi otomatik yedek alınacak. Devam edilsin mi?", parent=self): return
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True)
                backup_path=os.path.join(BACKUP_DIR, f"cari_silme_oncesi_{datetime.now():%Y%m%d_%H%M%S}.db")
                with sqlite3.connect(DB_PATH) as src, sqlite3.connect(backup_path) as dst: src.backup(dst)
                with connect() as c:
                    # Satış kaynaklı kasa/banka/çek kayıtlarını da referanslarından temizle.
                    sales=c.execute("SELECT id FROM beton_satis WHERE cari_id=?",(row_id,)).fetchall()
                    for (sid,) in sales:
                        refs=(f"Beton satış tahsilatı #{sid}", f"Beton satış #{sid}")
                        c.execute("DELETE FROM nakit_kasa WHERE aciklama IN (?,?)",refs)
                        c.execute("DELETE FROM banka_hareket WHERE aciklama IN (?,?)",refs)
                        c.execute("DELETE FROM cek WHERE cari_id=? AND notlar=?",(row_id,refs[0]))
                    # Cariyle ilişkili gelen fatura ödeme kayıtları, hareketleri ve finans yansımaları.
                    invoices=c.execute("SELECT id FROM gelen_fatura WHERE cari_id=?",(row_id,)).fetchall()
                    for (fid,) in invoices:
                        ref=f"Gelen fatura #{fid} ödeme - %"
                        c.execute("DELETE FROM hareket WHERE cari_id=? AND aciklama LIKE ?",(row_id,ref))
                        c.execute("DELETE FROM nakit_kasa WHERE aciklama LIKE ?",(ref,))
                        c.execute("DELETE FROM banka_hareket WHERE aciklama LIKE ?",(ref,))
                    c.execute("DELETE FROM hareket WHERE cari_id=?",(row_id,))
                    c.execute("DELETE FROM cek WHERE cari_id=?",(row_id,))
                    c.execute("DELETE FROM beton_satis WHERE cari_id=?",(row_id,))
                    c.execute("DELETE FROM fatura WHERE cari_id=?",(row_id,))
                    c.execute("DELETE FROM gelen_fatura WHERE cari_id=?",(row_id,))
                    c.execute("DELETE FROM cari WHERE id=?",(row_id,))
                self.refresh_all()
                messagebox.showinfo("Cari silindi", f"Cari ve bağlı kayıtları silindi.\nGeri dönüş yedeği: {backup_path}", parent=self)
            except Exception as e:
                messagebox.showerror("Silme başarısız", f"Cari silinemedi; işlem geri alındı.\n{e}", parent=self)
            return
        if not messagebox.askyesno("Onay","Seçili kayıt silinsin mi?", parent=self): return
        with connect() as c:
            if table == "stok_hareket":
                movement=c.execute("SELECT stok_id,yon,miktar FROM stok_hareket WHERE id=?",(row_id,)).fetchone()
                if movement:
                    delta=movement[2] if movement[1]=="Giriş" else -movement[2]
                    c.execute("UPDATE stok SET miktar=miktar-? WHERE id=?",(delta,movement[0]))
            if table == "beton_satis":
                sale_ref=f"Beton satış tahsilatı #{row_id}"
                c.execute("DELETE FROM hareket WHERE aciklama=? OR aciklama LIKE ?",(sale_ref,f"%[SATIŞ:{row_id}]%"))
                c.execute("DELETE FROM nakit_kasa WHERE aciklama=?",(sale_ref,))
                c.execute("DELETE FROM banka_hareket WHERE aciklama=?",(sale_ref,))
                c.execute("DELETE FROM cek WHERE notlar=?",(sale_ref,))
            if table == "cek":
                check_no=self._selected_check_number(tree)
                if check_no: c.execute("DELETE FROM hareket WHERE kategori='Çek' AND aciklama LIKE ?",(f"%çek #{check_no}%",))
            if table == "gelen_fatura":
                ref=f"Gelen fatura #{row_id} ödeme - %"
                c.execute("DELETE FROM hareket WHERE aciklama LIKE ?",(ref,))
                c.execute("DELETE FROM nakit_kasa WHERE aciklama LIKE ?",(ref,))
                c.execute("DELETE FROM banka_hareket WHERE aciklama LIKE ?",(ref,))
            try: c.execute(f"DELETE FROM {table} WHERE id=?",(row_id,))
            except sqlite3.IntegrityError:
                return messagebox.showerror("Kayıt silinemedi", "Bu kayıt başka işlemlere bağlı. Veri kaybını önlemek için silme durduruldu.", parent=self)
        self.refresh_all()

    def _selected_check_number(self, tree):
        sel=tree.selection()
        return tree.item(sel[0],"values")[3] if sel else ""

    def _apply_tree_search(self, tree):
        rows = getattr(tree, "_all_rows", [])
        term = tree._search_var.get().strip().casefold() if hasattr(tree, "_search_var") else ""
        shown = [r for r in rows if term in " ".join(str(v or "") for v in r).casefold()] if term else rows
        tree.delete(*tree.get_children())
        row_tags = getattr(tree, "_row_tags", {})
        for idx, row in enumerate(shown):
            tag = row_tags.get(str(row[0]))
            if not tag:
                tag = "even" if idx % 2 == 0 else "odd"
            tree.insert("", "end", values=row, tags=(tag,))

    def fill_tree(self, tree, rows):
        tree._all_rows = list(rows)
        self._apply_tree_search(tree)

    def refresh_dashboard(self):
        with connect() as c:
            cari = c.execute("SELECT COUNT(*) FROM cari").fetchone()[0]
            gelir = c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE tur='Gelir'").fetchone()[0]
            gider = c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE tur='Gider'").fetchone()[0]
            cek = c.execute("SELECT COALESCE(SUM(tutar),0) FROM cek WHERE durum='Bekliyor'").fetchone()[0]
            stok = c.execute("SELECT COUNT(*) FROM stok").fetchone()[0]
            debt_rows = c.execute("SELECT cari_id,bakiye,toplam_alacak FROM cari_bakiye WHERE bakiye>0.005").fetchall()
            acik_borc = sum(safe_float(r[1]) for r in debt_rows)
            odemesiz = sum(1 for r in debt_rows if safe_float(r[2]) <= 0)
        self.metric_labels["cari"].config(text=str(cari))
        self.metric_labels["gelir"].config(text=f"{gelir:,.2f} TL")
        self.metric_labels["gider"].config(text=f"{gider:,.2f} TL")
        self.metric_labels["cek"].config(text=f"{cek:,.2f} TL")
        self.metric_labels["stok"].config(text=str(stok))
        self.metric_labels["acik_borc"].config(text=f"{acik_borc:,.2f} TL")
        if hasattr(self, "debt_alert_label"):
            self.debt_alert_label.config(text=f"⚠ ÖDEME BEKLEYEN CARİ: {odemesiz} hesapta hiç tahsilat yok. Toplam açık cari bakiyesi: {acik_borc:,.2f} TL. Cari Hesaplar sekmesinde kırmızı/turuncu uyarılarla işaretlenir.")

    def refresh_all(self):
        with connect() as c:

            self.fill_tree(self.hareket_tree, c.execute("""SELECT h.id,h.tarih,h.tur,h.kategori,COALESCE(c.unvan,''),h.aciklama,printf('%.2f',h.tutar),h.odeme_yontemi FROM hareket h LEFT JOIN cari c ON c.id=h.cari_id ORDER BY h.tarih ASC,h.id ASC""").fetchall())
            self.fill_tree(self.cek_tree, c.execute("SELECT id,yon,COALESCE(banka,'') || CASE WHEN banka_kodu IS NULL OR banka_kodu='' THEN '' ELSE ' (' || banka_kodu || ')' END,cek_no,COALESCE((SELECT unvan FROM cari WHERE cari.id=cek.cari_id),kesideci_lehtar),vade,printf('%.2f',tutar),durum FROM cek ORDER BY vade ASC,id ASC").fetchall())
            self.fill_tree(self.stok_tree, c.execute("SELECT id,kod,ad,birim,printf('%.2f',miktar),printf('%.2f',birim_fiyat),printf('%.2f',miktar*birim_fiyat) FROM stok ORDER BY ad").fetchall())
            self.fill_tree(self.personel_tree, c.execute("SELECT id,ad,gorev,telefon,printf('%.2f',maas),notlar FROM personel ORDER BY ad").fetchall())
            beton_rows = c.execute("""SELECT b.id,b.tarih,COALESCE(c.unvan,''),COALESCE(b.beton_cinsi,''),b.miktar_m3,
                b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,b.odeme_durumu,COALESCE(b.aciklama,''),COALESCE(s.ad,'')
                FROM beton_satis b LEFT JOIN cari c ON c.id=b.cari_id LEFT JOIN santiye s ON s.id=b.santiye_id ORDER BY b.tarih ASC,b.id ASC""").fetchall()
            beton_display = []
            for b in beton_rows:
                detail = format_beton_aciklama(b[3],b[4],b[5],b[6],b[7],b[10],b[9],b[8])
                beton_display.append((b[0],b[1],(b[2] + (" / " + b[10] if b[10] else "")),b[3],format_m3(b[4]),f"{safe_float(b[5]):,.2f}",f"{safe_float(b[6]):,.2f}" if safe_float(b[6]) else "Yok",f"{safe_float(b[7]):,.2f}",b[8],detail))
            self.fill_tree(self.beton_tree, beton_display)
            if hasattr(self, "santiye_tree"):
                try:
                    if self.santiye_tree.winfo_exists():
                        self.fill_tree(self.santiye_tree,c.execute("SELECT id,ad,COALESCE(adres,''),COALESCE(yetkili,''),COALESCE(telefon,''),COALESCE(notlar,'') FROM santiye ORDER BY ad").fetchall())
                except tk.TclError:
                    pass
            self.fill_tree(self.fatura_tree,c.execute("""SELECT f.id,f.tarih,COALESCE(c.unvan,''),COALESCE(f.belge_no,''),f.miktar_m3,
                printf('%.2f',f.ara_toplam),printf('%.2f',f.kdv_tutar),printf('%.2f',f.genel_toplam),
                printf('%.2f',f.tevkifat_tutar),printf('%.2f',f.tahsil_edilecek),
                printf('%.2f',f.odenen),printf('%.2f',f.tahsil_edilecek-f.odenen),
                CASE WHEN f.tahsil_edilecek-f.odenen<=0.005 THEN 'Tahsil edildi'
                     WHEN f.odenen>0 THEN 'Kısmi tahsilat' ELSE 'Tahsil edilmedi' END
                FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id ORDER BY f.tarih ASC,f.id ASC""").fetchall())
            if hasattr(self,"gelen_fatura_tree"):
                self.fill_tree(self.gelen_fatura_tree,c.execute("SELECT gf.id,gf.tarih,COALESCE(gf.vade,''),COALESCE(c.unvan,''),gf.belge_no,printf('%.2f',gf.ara_toplam),printf('%.2f',gf.kdv_tutar),printf('%.2f',gf.genel_toplam),printf('%.2f',gf.odenen),printf('%.2f',gf.genel_toplam-gf.odenen),gf.durum FROM gelen_fatura gf LEFT JOIN cari c ON c.id=gf.cari_id ORDER BY CASE WHEN gf.vade IS NULL OR gf.vade='' THEN gf.tarih ELSE gf.vade END ASC,gf.id ASC").fetchall())
            if hasattr(self, "kasa_tree"):
                self.fill_tree(self.kasa_tree,c.execute("SELECT id,tarih,yon,kaynak,COALESCE(aciklama,''),printf('%.2f',tutar),COALESCE(kullanici,'') FROM nakit_kasa ORDER BY tarih ASC,id ASC").fetchall())
            if hasattr(self,"stok_hareket_tree"):
                self.fill_tree(self.stok_hareket_tree,c.execute("SELECT sh.id,sh.tarih,COALESCE(s.ad,''),sh.yon,printf('%.2f',sh.miktar),COALESCE(sh.aciklama,'') FROM stok_hareket sh LEFT JOIN stok s ON s.id=sh.stok_id ORDER BY sh.tarih ASC,sh.id ASC").fetchall())
            if hasattr(self,"maas_tree"):
                self.fill_tree(self.maas_tree,c.execute("SELECT m.id,m.tarih,COALESCE(p.ad,''),COALESCE(m.islem_turu,'Maaş'),m.donem,printf('%.2f',m.tutar),COALESCE(m.yontem,''),COALESCE(m.aciklama,'') FROM maas_odeme m LEFT JOIN personel p ON p.id=m.personel_id ORDER BY m.tarih ASC,m.id ASC").fetchall())
            if hasattr(self,"banka_tree"):
                self.fill_tree(self.banka_tree,c.execute("SELECT id,tarih,banka,COALESCE(hesap,''),yon,COALESCE(aciklama,''),printf('%.2f',tutar) FROM banka_hareket ORDER BY tarih ASC,id ASC").fetchall())
                banka_net=c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE -tutar END),0) FROM banka_hareket").fetchone()[0]
                self.banka_ozet.config(text=f"Banka net hareketi: {safe_float(banka_net):,.2f} TL")
            kasa_bakiye = c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE -tutar END),0) FROM nakit_kasa").fetchone()[0]
        if hasattr(self, "kasa_ozet"):
            self.kasa_ozet.config(text=f"Kasa bakiyesi: {safe_float(kasa_bakiye):,.2f} TL")
        self.refresh_cari()
        self.refresh_borc_takip()
        self.refresh_dashboard()

        if hasattr(self, 'alisveris_tree'):
            self.refresh_alisveris_cari()

    def pdf_menu(self):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        choices=["Cari Hesap Ekstresi","Cari Hareket Dökümü","Cari Beton Satış Dökümü","Gelir-Gider Raporu","Çek Listesi","Alınan Çekler Dökümü","Verilen Çekler Dökümü","Beton Satış Raporu","Beton Satışları Dökümü","Fatura Taslağı PDF","Gelen Alış Faturaları","Kasa ve Banka Raporu"]
        d=self.prompt("PDF Raporu Oluştur",[("rapor","Rapor Türü","combo",choices)])
        if not d: return
        cari_id=None
        if d["rapor"]=="Fatura Taslağı PDF":
            with connect() as c: invoices=c.execute("SELECT id,COALESCE(belge_no,'Belgesiz'),tarih FROM fatura ORDER BY id DESC").fetchall()
            if not invoices: return messagebox.showinfo("Bilgi","Önce Fatura / Belge bölümünden bir taslak oluşturun.")
            q=self.prompt("Fatura Taslağı Seç",[("fatura","Taslak","combo",[f"{i} - {no} ({tarih})" for i,no,tarih in invoices])])
            if not q:return
            cari_id=int(q["fatura"].split(" - ",1)[0])
        if d["rapor"] in ("Cari Hesap Ekstresi", "Cari Hareket Dökümü", "Cari Beton Satış Dökümü"):
            names=[f"{i} - {n}" for i,n in cariler]
            if not names: return messagebox.showinfo("Bilgi","Cari ekstresi için önce cari hesap ekleyin.")
            q=self.prompt("Cari Seç",[("cari","Cari Hesap","combo",names)])
            if not q: return
            cari_id=int(q["cari"].split(" - ",1)[0])
        try:
            self._preview_then_save_pdf(d["rapor"], None, cari_id)
        except ImportError:
            messagebox.showerror("Eksik bileşen","PDF oluşturmak için paket klasöründeki exe_olustur.bat dosyasını çalıştırın veya Komut İstemi'nde: python -m pip install reportlab")
        except Exception as e:
            messagebox.showerror("PDF hatası",str(e))

    def generate_pdf(self, report_type, path, cari_id=None):
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        # Türkçe karakterleri destekleyen Windows yazı tipini bul; Helvetica son çare.
        font_name="Helvetica"
        font_paths=[os.path.join(APP_DIR,"fonts","DejaVuSans.ttf"),
                    r"C:\Windows\Fonts\arial.ttf",r"C:\Windows\Fonts\Arial.ttf",
                    r"C:\Windows\Fonts\calibri.ttf",r"C:\Windows\Fonts\tahoma.ttf",
                    r"C:\Windows\Fonts\segoeui.ttf"]
        for fp in font_paths:
            if os.path.isfile(fp):
                try:
                    pdfmetrics.registerFont(TTFont("HoytasTurkish",fp))
                    bold_fp=os.path.join(APP_DIR,"fonts","DejaVuSans-Bold.ttf") if "DejaVuSans.ttf" in fp else None
                    if bold_fp and os.path.isfile(bold_fp):
                        pdfmetrics.registerFont(TTFont("HoytasTurkish-Bold",bold_fp))
                        pdfmetrics.registerFontFamily("HoytasTurkish",normal="HoytasTurkish",bold="HoytasTurkish-Bold",italic="HoytasTurkish",boldItalic="HoytasTurkish-Bold")
                    font_name="HoytasTurkish"
                    break
                except Exception:
                    continue
        doc=SimpleDocTemplate(path,pagesize=landscape(A4) if report_type not in ("Cari Hesap Ekstresi","Fatura Taslağı PDF") else A4,
            rightMargin=28,leftMargin=28,topMargin=30,bottomMargin=30,
            title=f"HOYTAŞ - {report_type}", author="HOYTAŞ Hazır Beton", subject="Finans ve İşletme Yönetimi Raporu", pageCompression=1)
        styles=getSampleStyleSheet()
        title_style=ParagraphStyle("HoytasTitle",parent=styles["Title"],fontName=font_name,textColor=colors.HexColor("#76ad42"),alignment=TA_CENTER)
        cari_title_style=ParagraphStyle("HoytasCariTitle",parent=styles["Title"],fontName=font_name,fontSize=14,textColor=colors.HexColor("#30363b"),alignment=TA_CENTER)
        normal=ParagraphStyle("HoytasNormal",parent=styles["BodyText"],fontName=font_name,fontSize=8,leading=10)
        logo_path=os.path.join(APP_DIR, "hoytas_logo.png")
        story=[]
        beton_summary_rows=None
        cari_summary_rows=None
        if os.path.exists(logo_path):
            logo=RLImage(logo_path, width=205, height=82)
            logo.hAlign="CENTER"
            story.extend([logo, Spacer(1,4)])
        if report_type == "Cari Hesap Ekstresi":
            story.extend([
                Paragraph("HOYTAŞ HAZIR BETON", title_style),
                Paragraph("FİNANS VE İŞLETME YÖNETİMİ", normal),
                Paragraph("Rapor Tarihi: "+datetime.now().strftime("%d.%m.%Y %H:%M"),normal),
                Spacer(1,8)
            ])
        else:
            story.extend([Paragraph("HOYTAŞ HAZIR BETON - Finans ve İşletme Yönetimi",title_style),
                          Paragraph(escape(report_type), title_style), Spacer(1,8),
                          Paragraph("Rapor tarihi: "+datetime.now().strftime("%d.%m.%Y %H:%M"),normal),Spacer(1,12)])
        with connect() as c:
            if report_type=="Cari Hesap Ekstresi":
                cari=c.execute("SELECT unvan,telefon,vergi_no FROM cari WHERE id=?",(cari_id,)).fetchone()
                if not cari: raise ValueError("Cari hesap bulunamadı.")
                moves=c.execute("SELECT id,tarih,tur,kategori,COALESCE(aciklama,''),COALESCE(odeme_yontemi,''),tutar FROM hareket WHERE cari_id=? AND tur IN ('Borç','Alacak') ORDER BY tarih ASC,id ASC",(cari_id,)).fetchall()
                first_date=display_date(moves[0][1]) if moves else "-"
                last_date=display_date(moves[-1][1]) if moves else "-"
                story.append(Paragraph("CARİ HESAP EKSTRESİ",cari_title_style))
                story.append(Spacer(1,9))
                info_data=[
                    [Paragraph(f"<b>Cari Ünvanı:</b> {escape(cari[0])}",normal), Paragraph(f"<b>Başlangıç Tarihi:</b> {first_date}",normal)],
                    [Paragraph(f"<b>Telefon:</b> {escape(cari[1] or '-')}",normal), Paragraph(f"<b>Bitiş Tarihi:</b> {last_date}",normal)],
                    [Paragraph(f"<b>Vergi No:</b> {escape(cari[2] or '-')}",normal), Paragraph(f"<b>Rapor Dönemi:</b> {first_date} - {last_date}",normal)]]
                info=Table(info_data,colWidths=[300,220])
                info.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,-1),colors.HexColor("#F2F7EF")),
                    ('BOX',(0,0),(-1,-1),0.8,colors.HexColor("#78A943")),
                    ('INNERGRID',(0,0),(-1,-1),0.35,colors.HexColor("#D8E6CD")),
                    ('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),
                    ('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),6),
                    ('BOTTOMPADDING',(0,0),(-1,-1),6)]))
                story.extend([info,Spacer(1,10)])
                rows=[["Tarih","Fiş No","İşlem Türü","Belge No","Açıklama","Borç (TL)","Alacak (TL)","Bakiye (TL)"]]
                sales_for_pdf = {r[0]: r[1:] for r in c.execute("SELECT b.id,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.aciklama,''),COALESCE(b.odeme_durumu,''),COALESCE(s.ad,'') FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=?", (cari_id,)).fetchall()}
                bal=0.0
                for move_id,tarih,tur,kat,acik,odeme,tutar in moves:
                    import re
                    sale_match = re.search(r"(?:\[SATIŞ:(\d+)\]|Beton satış #?(\d+))", acik, re.IGNORECASE)
                    if sale_match and tur == "Borç":
                        sid = int(sale_match.group(1) or sale_match.group(2)); sale=sales_for_pdf.get(sid)
                        if sale:
                            cinsi,m3,fiyat,pompa,toplam_satis,notlar,durum,santiye=sale
                            acik=format_beton_aciklama(cinsi,m3,fiyat,pompa,toplam_satis,santiye,notlar,durum)
                    bal += tutar if tur=="Borç" else -tutar
                    balance_label=f"{abs(bal):,.2f} ({'B' if bal>=0 else 'A'})"
                    rows.append([display_date(tarih),f"{move_id:07d}",kat,"",Paragraph(escape(acik),normal),f"{tutar:,.2f}" if tur=="Borç" else "",f"{tutar:,.2f}" if tur=="Alacak" else "",balance_label])
                total_debit=sum(r[6] for r in moves if r[2]=='Borç'); total_credit=sum(r[6] for r in moves if r[2]=='Alacak')
                rows.append(["","","","","Rapor Toplamı",f"{total_debit:,.2f}",f"{total_credit:,.2f}",f"{abs(bal):,.2f} ({'B' if bal>=0 else 'A'})"])
                total_collected=c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE cari_id=? AND tur='Alacak' AND (kategori LIKE '%Tahsilat%' OR aciklama LIKE '%tahsilat%')",(cari_id,)).fetchone()[0]
                cari_summary_rows=[
                    ["CARİ HESAP ÖZETİ", "TOPLAM"],
                    ["Toplam Borç", f"{total_debit:,.2f} TL"],
                    ["Toplam Alacak / Tahsilat", f"{total_credit:,.2f} TL"],
                    ["Kalan Cari Bakiye", f"{abs(bal):,.2f} TL ({'Borç' if bal>=0 else 'Alacak'})"],
                ]
            elif report_type=="Cari Beton Satış Dökümü":
                cari=c.execute("SELECT unvan,telefon,vergi_no FROM cari WHERE id=?",(cari_id,)).fetchone()
                if not cari: raise ValueError("Cari hesap bulunamadı.")
                sales=c.execute("SELECT b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(s.ad,''),COALESCE(b.aciklama,'') FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC",(cari_id,)).fetchall()
                story.append(Paragraph(f"<b>Cari:</b> {escape(cari[0])} &nbsp;&nbsp; <b>Telefon:</b> {escape(cari[1] or '-')}",normal))
                rows=[["Tarih","Beton Cinsi","Miktar (m³)","Birim Fiyat (TL)","Pompa (TL)","Toplam (TL)","Ödeme Durumu","Şantiye","Açıklama"]]
                for r in sales:
                    rows.append([display_date(r[0]),r[1],format_m3(r[2]),f"{r[3]:,.2f}",f"{r[4]:,.2f}" if r[4] else "-",f"{r[5]:,.2f}",r[6],Paragraph(escape(r[7] or '-'),normal),Paragraph(escape(r[8]),normal)])
                total_m3=sum(safe_float(r[2] or 0) for r in sales); total_amount=sum(safe_float(r[5] or 0) for r in sales)
                rows.append(["GENEL TOPLAM","",format_m3(total_m3),"","",f"{total_amount:,.2f}","","",f"{len(sales)} satış"])
                story.insert(2,Paragraph(f"Toplam beton miktarı: <b>{format_m3(total_m3)} m³</b> &nbsp;&nbsp; | &nbsp;&nbsp; Toplam satış: <b>{total_amount:,.2f} TL</b>",normal))
            elif report_type=="Cari Hareket Dökümü":
                cari=c.execute("SELECT unvan FROM cari WHERE id=?",(cari_id,)).fetchone()
                if not cari: raise ValueError("Cari hesap bulunamadı.")
                moves=c.execute("SELECT tarih,tur,kategori,COALESCE(aciklama,''),COALESCE(odeme_yontemi,''),tutar FROM hareket WHERE cari_id=? AND tur IN ('Borç','Alacak') ORDER BY tarih ASC,id ASC",(cari_id,)).fetchall()
                rows=[["Tarih","Tür","Kategori","Açıklama","Ödeme Yöntemi","Borç (TL)","Alacak (TL)","Bakiye (TL)"]]
                balance=0.0
                for tarih,tur,kategori,acik,odeme,tutar in moves:
                    balance += tutar if tur == "Borç" else -tutar
                    rows.append([display_date(tarih),tur,kategori,Paragraph(escape(acik),normal),odeme,f"{tutar:,.2f}" if tur == "Borç" else "",f"{tutar:,.2f}" if tur == "Alacak" else "",f"{abs(balance):,.2f}"])
                debit=sum(r[5] for r in moves if r[1] == "Borç"); credit=sum(r[5] for r in moves if r[1] == "Alacak")
                rows.append(["","","","TOPLAM","",f"{debit:,.2f}",f"{credit:,.2f}",f"{abs(balance):,.2f}"])
            elif report_type=="Cari Borç Takip Raporu":
                rows=[["Cari Ünvanı", "Toplam Borç (TL)", "Ödenen (TL)", "Kalan Borç (TL)", "Durum"]]
                debt_rows=c.execute("""SELECT unvan,toplam_borc,toplam_alacak,bakiye
                    FROM cari_bakiye WHERE bakiye > 0.005
                    ORDER BY bakiye DESC,unvan COLLATE NOCASE""").fetchall()
                total_debt=total_paid=total_remaining=0.0
                for name, debt, paid, remaining in debt_rows:
                    status="Kısmi ödeme" if paid > 0.005 else "Ödeme bekliyor"
                    rows.append([Paragraph(escape(name),normal),f"{debt:,.2f}",f"{paid:,.2f}",f"{remaining:,.2f}",status])
                    total_debt += safe_float(debt or 0); total_paid += safe_float(paid or 0); total_remaining += safe_float(remaining or 0)
                rows.append(["TOPLAM",f"{total_debt:,.2f}",f"{total_paid:,.2f}",f"{total_remaining:,.2f}",f"{len(debt_rows)} cari"])
                story.insert(2, Paragraph(f"Borçlu cari sayısı: <b>{len(debt_rows)}</b> &nbsp;&nbsp; | &nbsp;&nbsp; TOPLAM CARİ BORCU: <b>{total_remaining:,.2f} TL</b>", normal))
                if not debt_rows:
                    rows.append(["Borç bakiyesi bulunan cari hesap yok.","","","",""])
            elif report_type=="Gelir-Gider Raporu":
                rows=[["Tarih","Tür","Kategori","Cari","Açıklama","Ödeme","Tutar (TL)"]]
                for r in c.execute("SELECT h.tarih,h.tur,h.kategori,COALESCE(c.unvan,''),COALESCE(h.aciklama,''),COALESCE(h.odeme_yontemi,''),h.tutar FROM hareket h LEFT JOIN cari c ON c.id=h.cari_id WHERE h.tur IN ('Gelir','Gider') ORDER BY h.tarih ASC,h.id ASC"):
                    rows.append([display_date(r[0]),r[1],r[2],Paragraph(escape(r[3]),normal),Paragraph(escape(r[4]),normal),r[5],f"{r[6]:,.2f}"])
                income=c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE tur='Gelir'").fetchone()[0]
                expense=c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE tur='Gider'").fetchone()[0]
                rows.extend([["","","","","Toplam gelir","",f"{income:,.2f}"],["","","","","Toplam gider","",f"{expense:,.2f}"],["","","","","Net","",f"{income-expense:,.2f}"]])
            elif report_type in ("Çek Listesi", "Alınan Çekler Dökümü", "Verilen Çekler Dökümü"):
                rows=[["Yön","Banka","Banka Kodu","Çek No","Keşideci/Lehtar","Vade","Tutar (TL)","Durum","Notlar"]]
                if report_type=="Alınan Çekler Dökümü":
                    cheque_sql="SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar,durum,COALESCE(notlar,'') FROM cek WHERE yon='Alınan' ORDER BY vade ASC,id ASC"
                elif report_type=="Verilen Çekler Dökümü":
                    cheque_sql="SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar,durum,COALESCE(notlar,'') FROM cek WHERE yon='Verilen' ORDER BY vade ASC,id ASC"
                else:
                    cheque_sql="SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar,durum,COALESCE(notlar,'') FROM cek ORDER BY vade ASC,id ASC"
                for r in c.execute(cheque_sql):
                    rows.append([r[0],r[1],r[2],r[3],Paragraph(escape(r[4]),normal),display_date(r[5]),f"{r[6]:,.2f}",r[7],Paragraph(escape(r[8]),normal)])
                cheque_total=c.execute("SELECT COALESCE(SUM(tutar),0),COUNT(*) FROM cek" + (" WHERE yon='Alınan'" if report_type=="Alınan Çekler Dökümü" else " WHERE yon='Verilen'" if report_type=="Verilen Çekler Dökümü" else "")).fetchone()
                rows.append(["","","","","TOPLAM",f"{cheque_total[1]} adet",f"{cheque_total[0]:,.2f}","",""])
            elif report_type=="Fatura Taslağı PDF":
                invoice=c.execute("SELECT f.tarih,COALESCE(c.unvan,''),COALESCE(c.telefon,''),COALESCE(c.vergi_no,''),COALESCE(f.belge_no,''),COALESCE(f.aciklama,''),f.miktar_m3,f.birim_fiyat,f.ara_toplam,f.kdv_oran,f.kdv_tutar,f.genel_toplam,f.durum FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id WHERE f.id=?",(cari_id,)).fetchone()
                if not invoice: raise ValueError("Fatura taslağı bulunamadı.")
                labels=[("Tarih",display_date(invoice[0])),("Cari / Ünvan",invoice[1]),("Telefon",invoice[2] or "-"),("Vergi No",invoice[3] or "-"),("Belge No",invoice[4] or "Taslak"),("Açıklama",invoice[5] or "-"),("Belge durumu",invoice[12])]
                for label,value in labels: story.append(Paragraph(f"<b>{escape(label)}:</b> {escape(str(value))}",normal)); story.append(Spacer(1,4))
                story.append(Spacer(1,10))
                rows=[["Ürün / Hizmet","Miktar (m³)","Birim Fiyat (TL)","Ara Toplam (TL)"],["Hazır beton",format_m3(invoice[6]),f"{invoice[7]:,.2f}",f"{invoice[8]:,.2f}"]]
                rows.extend([["","","KDV (%g)" % invoice[9],f"{invoice[10]:,.2f}"],["","","GENEL TOPLAM",f"{invoice[11]:,.2f}"]])
            elif report_type=="Gelen Alış Faturaları":
                rows=[["Fatura Tarihi","Vade","Tedarikçi","Belge No","Açıklama","Genel Toplam (TL)","Ödenen (TL)","Kalan (TL)","Durum"]]
                for r in c.execute("SELECT gf.tarih,COALESCE(gf.vade,''),COALESCE(c.unvan,''),gf.belge_no,COALESCE(gf.aciklama,''),gf.genel_toplam,gf.odenen,gf.genel_toplam-gf.odenen,gf.durum FROM gelen_fatura gf LEFT JOIN cari c ON c.id=gf.cari_id ORDER BY COALESCE(NULLIF(gf.vade,''),gf.tarih) ASC,gf.id ASC"):
                    rows.append([display_date(r[0]),display_date(r[1]) if r[1] else "-",Paragraph(escape(r[2]),normal),r[3],Paragraph(escape(r[4]),normal),f"{r[5]:,.2f}",f"{r[6]:,.2f}",f"{r[7]:,.2f}",r[8]])
                totals=c.execute("SELECT COALESCE(SUM(genel_toplam),0),COALESCE(SUM(odenen),0),COALESCE(SUM(genel_toplam-odenen),0) FROM gelen_fatura").fetchone()
                rows.append(["","","","","TOPLAM",f"{totals[0]:,.2f}",f"{totals[1]:,.2f}",f"{totals[2]:,.2f}",""])
            elif report_type=="Kasa ve Banka Raporu":
                rows=[["Tarih","Hesap Türü","Banka/Kaynak","İşlem","Açıklama","Giriş (TL)","Çıkış (TL)"]]
                for r in c.execute("SELECT tarih,'Nakit Kasa',kaynak,yon,COALESCE(aciklama,''),tutar FROM nakit_kasa ORDER BY tarih ASC,id ASC"):
                    rows.append([display_date(r[0]),r[1],r[2],r[3],Paragraph(escape(r[4]),normal),f"{r[5]:,.2f}" if r[3]=="Para Girişi" else "-",f"{r[5]:,.2f}" if r[3]!="Para Girişi" else "-"])
                for r in c.execute("SELECT tarih,'Banka',banka,yon,COALESCE(aciklama,''),tutar FROM banka_hareket ORDER BY tarih ASC,id ASC"):
                    rows.append([display_date(r[0]),r[1],r[2],r[3],Paragraph(escape(r[4]),normal),f"{r[5]:,.2f}" if r[3]=="Para Girişi" else "-",f"{r[5]:,.2f}" if r[3]!="Para Girişi" else "-"])
                totals=c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE 0 END),0),COALESCE(SUM(CASE WHEN yon='Para Çıkışı' THEN tutar ELSE 0 END),0) FROM nakit_kasa").fetchone()
                bt=c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE 0 END),0),COALESCE(SUM(CASE WHEN yon='Para Çıkışı' THEN tutar ELSE 0 END),0) FROM banka_hareket").fetchone()
                rows.append(["","","","","TOPLAM GİRİŞ",f"{totals[0]+bt[0]:,.2f}",""])
                rows.append(["","","","","TOPLAM ÇIKIŞ","",f"{totals[1]+bt[1]:,.2f}"])
            else:
                rows=[["Tarih","Müşteri","Beton Cinsi","Miktar (m³)","Birim Fiyat","Pompa Ücreti","Toplam (TL)","Ödeme Durumu","Açıklama"]]
                for r in c.execute("SELECT b.tarih,COALESCE(c.unvan,''),COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,b.odeme_durumu,COALESCE(b.aciklama,''),COALESCE(s.ad,'') FROM beton_satis b LEFT JOIN cari c ON c.id=b.cari_id LEFT JOIN santiye s ON s.id=b.santiye_id ORDER BY b.tarih ASC,b.id ASC"):
                    detail = format_beton_aciklama(r[2],r[3],r[4],r[5],r[6],r[9],r[8],r[7])
                    rows.append([display_date(r[0]),Paragraph(escape(r[1]),normal),r[2],format_m3(r[3]),f"{r[4]:,.2f}",f"{r[5]:,.2f}" if r[5] else "Yok",f"{r[6]:,.2f}",r[7],Paragraph(escape(detail),normal)])
                total=c.execute("SELECT COALESCE(SUM(miktar_m3),0),COALESCE(SUM(toplam),0),COUNT(*) FROM beton_satis").fetchone()
                rows.append(["","GENEL TOPLAM",format_m3(total[0]),"","","",f"{total[1]:,.2f}","",""])
                # Ana listeyi başka modüllerle karıştırmadan beton türü özeti ekle.
                beton_summary_rows=[["Beton Cinsi","Satış Adedi","Toplam Miktar (m³)","Toplam Tutar (TL)"]]
                by_type=c.execute("SELECT CASE WHEN TRIM(COALESCE(beton_cinsi,''))='' THEN 'Belirtilmemiş' ELSE beton_cinsi END, COUNT(*), COALESCE(SUM(miktar_m3),0), COALESCE(SUM(toplam),0) FROM beton_satis GROUP BY CASE WHEN TRIM(COALESCE(beton_cinsi,''))='' THEN 'Belirtilmemiş' ELSE beton_cinsi END ORDER BY 1 COLLATE NOCASE").fetchall()
                for kind,count,m3,amount in by_type:
                    beton_summary_rows.append([kind,str(count),format_m3(m3),f"{amount:,.2f}"])
                beton_summary_rows.append(["GENEL TOPLAM",str(total[2]),format_m3(total[0]),f"{total[1]:,.2f}"])
        # Uzun açıklama/unvanların sayfa dışına taşmasını engelle: tüm hücreleri
        # Paragraph yaparak satır kaydır ve tabloyu kullanılabilir sayfa genişliğine sığdır.
        from reportlab.platypus import Paragraph as RLParagraph
        wrapped_rows=[]
        for row_index, row in enumerate(rows):
            wrapped=[]
            for cell in row:
                if isinstance(cell, RLParagraph):
                    wrapped.append(cell)
                elif isinstance(cell, str):
                    safe = escape(cell).replace("\n", "<br/>")
                    wrapped.append(RLParagraph(safe, normal))
                else:
                    wrapped.append(cell)
            wrapped_rows.append(wrapped)
        page_width = (A4[0] if report_type in ("Cari Hesap Ekstresi", "Fatura Taslağı PDF") else landscape(A4)[0]) - 56
        # Her rapor türüne uygun sütun genişlikleri: tarih/tür gibi kısa alanlar
        # açıklama ve cari unvanı gibi uzun alanları ezmesin.
        width_ratios = {
            "Cari Hesap Ekstresi": [0.72, 0.68, 0.78, 0.72, 2.05, 0.92, 0.92, 1.05],
            "Gelir-Gider Raporu": [0.78, 0.72, 0.95, 1.55, 2.20, 0.90, 1.00],
            "Çek Listesi": [0.65, 0.90, 0.85, 0.95, 1.45, 0.85, 0.95, 0.85, 1.35],
            "Alınan Çekler Dökümü": [0.65, 0.90, 0.85, 0.95, 1.45, 0.85, 0.95, 0.85, 1.35],
            "Verilen Çekler Dökümü": [0.65, 0.90, 0.85, 0.95, 1.45, 0.85, 0.95, 0.85, 1.35],
            "Cari Hareket Dökümü": [0.8, 0.8, 1.0, 2.0, 1.0, 1.0, 1.0, 1.0],
            "Cari Beton Satış Dökümü": [0.75, 0.9, 0.75, 0.9, 0.75, 0.9, 1.0, 1.0, 1.8],
            "Beton Satış Raporu": [0.78, 1.65, 0.90, 0.70, 0.90, 0.90, 1.00, 1.05, 1.60],
            "Beton Satışları Dökümü": [0.78, 1.65, 0.90, 0.70, 0.90, 0.90, 1.00, 1.05, 1.60],
            "Fatura Taslağı PDF": [1.75, 1.00, 1.20, 1.20],
            "Gelen Alış Faturaları": [0.85, 0.85, 1.45, 1.05, 1.55, 1.00, 1.00, 1.00, 0.95],
            "Kasa ve Banka Raporu": [0.85, 1.00, 1.15, 1.05, 2.10, 1.00, 1.00],
            "Cari Borç Takip Raporu": [2.20, 1.05, 1.05, 1.15, 1.20],
        }.get(report_type)
        col_count = max(1, max(len(r) for r in wrapped_rows))
        if not width_ratios or len(width_ratios) != col_count:
            width_ratios = [1.0] * col_count
        ratio_sum = sum(width_ratios)
        col_widths = [page_width * ratio / ratio_sum for ratio in width_ratios]
        # Başlık satırı yalnızca ilk sayfada görünür; her sayfada tarih/tür satırı
        # tekrar etmez ve tablo sütunları sayfa genişliğinin dışına taşmaz.
        table=Table(wrapped_rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT", splitByRow=1)
        header_bg=colors.HexColor("#24583D") if report_type=="Cari Hesap Ekstresi" else colors.HexColor("#477b35")
        header_fg=colors.white
        alternate=colors.HexColor("#F4F8EF") if report_type=="Cari Hesap Ekstresi" else colors.HexColor("#f2f5f8")
        table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),header_bg),("TEXTCOLOR",(0,0),(-1,0),header_fg),
            ("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),7),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),
            ("VALIGN",(0,0),(-1,-1),"TOP"),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,alternate]),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        story.append(table)
        if cari_summary_rows:
            story.extend([Spacer(1,12), Paragraph("CARİ HESAP ÖZETİ", title_style), Spacer(1,5)])
            summary_table=Table(cari_summary_rows,colWidths=[page_width*0.58,page_width*0.42],hAlign="LEFT")
            summary_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#24583D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),8),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F4F8EF")]),("ALIGN",(1,1),(1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
            story.append(summary_table)
        if beton_summary_rows:
            story.extend([Spacer(1,12), Paragraph("BETON M³ ÖZETİ", title_style), Spacer(1,5)])
            summary_table=Table(beton_summary_rows,colWidths=[page_width*0.38,page_width*0.16,page_width*0.22,page_width*0.24],repeatRows=1)
            summary_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#24583D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),8),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F4F8EF")]),("ALIGN",(1,1),(-1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
            story.append(summary_table)
        doc.build(story)

    def check_due_cheques(self):
        today=date.today(); soon=(today+timedelta(days=7)).isoformat(); today_s=today.isoformat()
        with connect() as c:
            rows=c.execute("SELECT yon,COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar FROM cek WHERE durum='Bekliyor' AND vade<=? ORDER BY vade ASC,id ASC",(soon,)).fetchall()
        if rows:
            lines=[]
            for yon,no,taraf,vade,tutar in rows[:15]:
                try: delta=(datetime.strptime(vade,"%Y-%m-%d").date()-today).days
                except ValueError: delta=0
                state="GECİKMİŞ" if delta<0 else ("BUGÜN" if delta==0 else f"{delta} gün kaldı")
                lines.append(f"{state}: {yon} çek {no} | {taraf} | Vade {vade} | {tutar:,.2f} TL")
            messagebox.showwarning("Çek vade hatırlatması","Yaklaşan veya geciken çekler:\n\n"+"\n".join(lines))

    def export_excel(self):
        path = filedialog.asksaveasfilename(title="HOYTAŞ Excel raporunu kaydet", defaultextension=".xlsx",
            filetypes=[("Excel çalışma kitabı", "*.xlsx")], initialfile="HOYTAS_Hazir_Beton_Rapor_" + datetime.now().strftime("%Y%m%d") + ".xlsx")
        if not path:
            return
        try:
            from openpyxl import Workbook, load_workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from openpyxl.drawing.image import Image as XLImage
            tables = [
                ("Cari Hesaplar", "cari"), ("Gelir Gider", "hareket"), ("Çekler", "cek"),
                ("Stok Depo", "stok"), ("Personel", "personel"), ("Beton Satışları", "beton_satis"),
                ("Şantiyeler", "santiye"), ("Fatura Taslakları", "fatura"), ("Gelen Alış Faturaları", "gelen_fatura"), ("Gelen Fatura Ödemeleri", "gelen_fatura_odeme"), ("Nakit Kasa", "nakit_kasa"), ("Stok Hareketleri", "stok_hareket"), ("Maaş Ödemeleri", "maas_odeme"), ("Banka İşlemleri", "banka_hareket")]
            wb = Workbook()
            wb.remove(wb.active)
            with connect() as c:
                for sheet_name, table in tables:
                    ws = wb.create_sheet(sheet_name[:31])
                    # Her Excel sayfasında kurumsal logo ve başlık
                    logo_path = os.path.join(APP_DIR, "hoytas_logo.png")
                    if os.path.exists(logo_path):
                        logo = XLImage(logo_path)
                        logo.width = 300
                        logo.height = 150
                        ws.add_image(logo, "A1")
                    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=5)
                    title_cell = ws.cell(row=6, column=1, value=f"HOYTAŞ HAZIR BETON | FİNANS VE İŞLETME YÖNETİMİ — {sheet_name.upper()}")
                    title_cell.font = Font(name="Calibri", size=14, bold=True, color="548235")
                    title_cell.alignment = Alignment(vertical="center")
                    ws.row_dimensions[6].height = 28
                    columns_info = c.execute(f"PRAGMA table_info({table})").fetchall()
                    column_names = [info[1] for info in columns_info]
                    date_col = next((name for name in ("tarih", "vade") if name in column_names), None)
                    order_sql = f" ORDER BY {date_col} ASC" if date_col else ""
                    if "id" in column_names:
                        order_sql += (", id ASC" if order_sql else " ORDER BY id ASC")
                    cur = c.execute(f"SELECT * FROM {table}{order_sql}")
                    headers = [d[0] for d in cur.description]
                    for idx, value in enumerate(headers, 1):
                        ws.cell(row=8, column=idx, value=value)
                    for ridx, row in enumerate(cur.fetchall(), 9):
                        for cidx, value in enumerate(row, 1):
                            ws.cell(row=ridx, column=cidx, value=value)
                    ws.freeze_panes = "A9"
                    ws.auto_filter.ref = f"A8:{get_column_letter(max(1,len(headers)))}{max(8,ws.max_row)}"
                    for cell in ws[8]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill("solid", fgColor="76AD42")
                        cell.alignment = Alignment(vertical="center", wrap_text=True)
                    ws.row_dimensions[8].height = 28
                    for col in ws.iter_cols(min_row=8, max_row=ws.max_row):
                        letter = get_column_letter(col[0].column)
                        max_len = max((len(str(cell.value or "")) for cell in col[:300]), default=10)
                        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 38)
            # Cari bazında toplam beton m³: Excel raporunda ayrı, kolay bulunur özet.
            ws_m3 = wb.create_sheet("Cari Beton m3 Özeti")
            ws_m3.append(["Cari Ünvan", "Toplam Beton Satışı (m³)", "Satış Tutarı (TL)"])
            with connect() as c:
                m3_rows = c.execute("SELECT COALESCE(c.unvan,'Cariye Bağlanmamış'), COALESCE(SUM(b.miktar_m3),0), COALESCE(SUM(b.toplam),0) FROM beton_satis b LEFT JOIN cari c ON c.id=b.cari_id GROUP BY b.cari_id, c.unvan ORDER BY c.unvan ASC").fetchall()
            for row in m3_rows:
                ws_m3.append([row[0], safe_float(row[1] or 0), safe_float(row[2] or 0)])
            m3_total_row = ws_m3.max_row + 1
            ws_m3.append(["GENEL TOPLAM", f"=SUM(B2:B{m3_total_row-1})", f"=SUM(C2:C{m3_total_row-1})"])
            for cell in ws_m3[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="477B35")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for col in ws_m3.columns:
                letter = get_column_letter(col[0].column)
                ws_m3.column_dimensions[letter].width = min(max(max(len(str(cell.value or "")) for cell in col) + 2, 16), 42)
            ws_m3.freeze_panes = "A2"
            ws_m3.auto_filter.ref = ws_m3.dimensions
            # Beton cinsine göre ayrı m³ ve satış tutarı özeti.
            ws_type = wb.create_sheet("Beton Cinsi m3 Özeti")
            ws_type.append(["Beton Cinsi", "Satış Adedi", "Toplam Miktar (m³)", "Toplam Tutar (TL)"])
            with connect() as c:
                type_rows = c.execute("SELECT CASE WHEN TRIM(COALESCE(beton_cinsi,''))='' THEN 'Belirtilmemiş' ELSE beton_cinsi END, COUNT(*), COALESCE(SUM(miktar_m3),0), COALESCE(SUM(toplam),0) FROM beton_satis GROUP BY CASE WHEN TRIM(COALESCE(beton_cinsi,''))='' THEN 'Belirtilmemiş' ELSE beton_cinsi END ORDER BY 1 COLLATE NOCASE").fetchall()
                type_totals = c.execute("SELECT COUNT(*), COALESCE(SUM(miktar_m3),0), COALESCE(SUM(toplam),0) FROM beton_satis").fetchone()
            for row in type_rows:
                ws_type.append([row[0], int(row[1]), safe_float(row[2] or 0), safe_float(row[3] or 0)])
            ws_type.append(["GENEL TOPLAM", int(type_totals[0]), safe_float(type_totals[1]), safe_float(type_totals[2])])
            for cell in ws_type[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="24583D")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for col in ws_type.columns:
                letter = get_column_letter(col[0].column)
                ws_type.column_dimensions[letter].width = min(max(max(len(str(cell.value or "")) for cell in col) + 2, 16), 42)
            ws_type.freeze_panes = "A2"
            ws_type.auto_filter.ref = ws_type.dimensions
            wb.save(path)
            # Kaydedilen xlsx dosyasini yeniden acarak dosya butunlugunu dogrula.
            from openpyxl import load_workbook
            check_wb = load_workbook(path, read_only=True, data_only=True)
            expected_sheets = [name[:31] for name, _ in tables] + ["Cari Beton m3 Özeti", "Beton Cinsi m3 Özeti"]
            actual_sheets = list(check_wb.sheetnames)
            check_wb.close()
            if actual_sheets != expected_sheets:
                raise RuntimeError("Excel dosyasının sayfa listesi beklenen raporla eşleşmiyor.")
            with open(path, "rb") as xlsx_file:
                if xlsx_file.read(2) != b"PK":
                    raise RuntimeError("Dosya geçerli bir XLSX/ZIP dosyası değil.")
            messagebox.showinfo("Excel hazır", f"Excel raporu oluşturuldu ve doğrulandı:\n{path}")
        except ImportError:
            messagebox.showerror("Eksik bileşen", "Excel çıktısı için requirements.txt içindeki paketleri kurun veya kurulum.bat dosyasını yeniden çalıştırın.")
        except Exception as e:
            messagebox.showerror("Excel dışa aktarma hatası", str(e))

    def export_csv(self):
        folder = filedialog.askdirectory(title="CSV dosyalarının kaydedileceği klasörü seçin")
        if not folder: return
        tables = {"cari":"cari_hesaplar.csv","hareket":"gelir_gider.csv","cek":"cekler.csv","stok":"stok.csv","personel":"personel.csv","beton_satis":"beton_satislari.csv","santiye":"santiyeler.csv","fatura":"fatura_taslaklari.csv","gelen_fatura":"gelen_alis_faturalari.csv","gelen_fatura_odeme":"gelen_fatura_odemeleri.csv","nakit_kasa":"nakit_kasa.csv","stok_hareket":"stok_hareketleri.csv","maas_odeme":"maas_odemeleri.csv","banka_hareket":"banka_islemleri.csv"}
        try:
            with connect() as c:
                for table, filename in tables.items():
                    cur = c.execute(f"SELECT * FROM {table}")
                    rows = cur.fetchall()
                    headers = [d[0] for d in cur.description]
                    with open(os.path.join(folder,filename),"w",newline="",encoding="utf-8-sig") as f:
                        w=csv.writer(f,delimiter=";"); w.writerow(headers); w.writerows(rows)
            messagebox.showinfo("Tamamlandı",f"CSV dosyaları kaydedildi:\n{folder}")
        except Exception as e:
            messagebox.showerror("Hata",str(e))

    def _create_backup_file(self):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        target = os.path.join(BACKUP_DIR, "hoytas_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".db")
        with connect() as c:
            target = _verified_sqlite_backup(c, "hoytas")
        # Kullanıcının geçmiş yedekleri otomatik silinmez; saklama kararı kullanıcıdadır.
        return target

    def backup(self):
        try:
            target = self._create_backup_file()
            messagebox.showinfo("Yedek alındı", f"Yedek dosyası oluşturuldu:\n{target}")
        except Exception as e:
            messagebox.showerror("Yedekleme hatası", str(e))

    def automatic_backup(self):
        """İlk açılışta ve günde bir kez otomatik veritabanı yedeği al."""
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            today = datetime.now().strftime("%Y%m%d")
            exists_today = any(name.startswith("hoytas_" + today) and name.endswith(".db") for name in os.listdir(BACKUP_DIR))
            if not exists_today:
                self._create_backup_file()
        except Exception as e:
            # Yedekleme başarısızlığı sessizce gizlenmesin; uygulama açılmaya devam eder.
            try: messagebox.showwarning("Otomatik yedekleme", f"Otomatik yedek alınamadı. Lütfen 'Şimdi Yedekle' seçeneğini kullanın.\n\n{e}", parent=self)
            except tk.TclError: pass

    def database_health_check(self):
        """Run non-destructive SQLite and FK checks, then create a verified snapshot."""
        try:
            with connect() as c:
                integrity = c.execute("PRAGMA integrity_check").fetchall()
                fk_errors = c.execute("PRAGMA foreign_key_check").fetchall()
                schema_version = c.execute("PRAGMA user_version").fetchone()[0]
                record_counts = {}
                for table in ("cari", "hareket", "cek", "beton_satis", "fatura", "gelen_fatura"):
                    record_counts[table] = c.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            backup_path = self._create_backup_file()
            integrity_ok = len(integrity) == 1 and integrity[0][0] == "ok"
            status = "BAŞARILI" if integrity_ok and not fk_errors else "İNCELEME GEREKİYOR"
            details = "\n".join(f"• {k}: {v:,} kayıt" for k,v in record_counts.items())
            fk_text = "Yabancı anahtar kontrolü: Sorun bulunmadı." if not fk_errors else f"Yabancı anahtar uyarısı: {len(fk_errors)} ilişki kaydı incelenmeli."
            report = (f"Sonuç: {status}\nŞema sürümü: {schema_version}\n"
                      f"SQLite bütünlük kontrolü: {'OK' if integrity_ok else integrity}\n{fk_text}\n\n"
                      f"Kayıt özeti:\n{details}\n\nDoğrulanmış yedek:\n{backup_path}")
            messagebox.showinfo("Veritabanı Sağlık Kontrolü", report, parent=self)
        except Exception as exc:
            messagebox.showerror("Veritabanı kontrolü başarısız",f"Kontrol tamamlanamadı. Mevcut veriler üzerinde onarım/silme yapılmadı.\n\n{exc}",parent=self)

    def show_db_path(self):
        messagebox.showinfo("Veritabanı",f"Veriler burada saklanıyor:\n{DB_PATH}\n\nEski veriler ilk açılışta bu konuma kopyalanır; yedekleri düzenli alın.")

if __name__ == "__main__":
    app = None
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        # If the desktop icon was lost during an earlier update, recreate it on first launch.
        try:
            _restore_desktop_shortcut_after_update(APP_DIR)
        except Exception as shortcut_exc:
            try:
                with open(os.path.join(USER_DATA_DIR, "guncelleme.log"), "a", encoding="utf-8") as log:
                    log.write(f"{datetime.now().isoformat(timespec='seconds')} | Kısayol: {shortcut_exc}\n")
            except OSError:
                pass
        app = HoytasApp()
        # HoytasApp.__init__ already performs the initial refresh; avoid a duplicate startup pass.
        app.after(500, app.automatic_backup)
        app.after(900, app.check_due_cheques)
        app.mainloop()
    except SystemExit:
        raise
    except Exception:
        detail = traceback.format_exc()
        log_path = os.path.join(USER_DATA_DIR, "program_hatalari.log")
        try:
            with open(log_path, "a", encoding="utf-8") as log:
                log.write("\n--- UYGULAMA ACILIS HATASI ---\n" + detail + "\n")
        except OSError:
            pass
        try:
            messagebox.showerror("HOYTAŞ açılış hatası", "Program açılırken bir hata oluştu. Veritabanınızdaki kayıtlar silinmedi.\n\nHata kaydı: " + log_path + "\n\n" + str(detail.splitlines()[-1] if detail.splitlines() else "Bilinmeyen hata"))
        except Exception:
            pass
        try:
            if app is not None:
                app.destroy()
        except Exception:
            pass

