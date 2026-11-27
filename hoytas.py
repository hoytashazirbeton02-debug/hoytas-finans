import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from calendar import monthrange
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import datetime, date, timedelta
from contextlib import nullcontext
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
import threading
import queue
import time
from xml.sax.saxutils import escape

from hybrid_branding import set_brand_logo
from invoice_matching import (GIB_TEBLIG, WITHHOLDING_LABELS, cents, invoice_quote,
                              post_invoice, record_is_posted, selected_sales, stored_details)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_CURRENT_USER = "Sistem"
TARGET_SCHEMA_VERSION = 62
APP_VERSION = 132
# Veritabanini uygulamanin acildigi konumdan bagimsiz, kalici kullanici klasorunde tut.
ISOLATED_PROFILE_DIR = os.environ.get("HOYTAS_PROFILE_DIR", "").strip()
USER_DATA_DIR = ISOLATED_PROFILE_DIR or os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "HOYTASFinans")
os.makedirs(USER_DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(USER_DATA_DIR, "hoytas.db")
LEGACY_DB_PATH = os.path.join(APP_DIR, "hoytas.db")
# Onceki surumdeki veritabanini ilk acilista bir defa aktar; var olan kayitlari ezme.
if not ISOLATED_PROFILE_DIR and not os.path.exists(DB_PATH) and os.path.isfile(LEGACY_DB_PATH):
    shutil.copy2(LEGACY_DB_PATH, DB_PATH)
BACKUP_DIR = os.path.join(USER_DATA_DIR, "yedekler")
UPDATE_CONFIG_PATH = os.path.join(USER_DATA_DIR, "guncelleme_ayarlari.json")
# Canlı yayın adresi doğrulanmış, anonim HTTPS erişimli bir manifest olmalıdır.
# Önceki OneDrive bağlantısı 401 verdiği için hatalı adres varsayılan olmaktan çıkarıldı.
DEFAULT_UPDATE_MANIFEST_URL = ""
LIVE_UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/hoytashazirbeton02-debug/hoytas-finans/refs/heads/v91-hibrit-entegrasyon/canli-manifest.json"
LIVE_UPDATE_FILES = {"Windows_Canli_Baslat.cmd", "hybrid_start_live.py", "hybrid_runtime.py",
                     "hybrid_http_relay.py", "hybrid_exchange.py", "hybrid_inbound.py",
                     "hybrid_identity.py", "hybrid_audit_bridge.py", "HOYTAS.pyw",
                     "hybrid_membership_client.py", "hybrid_desktop_membership.py", "hybrid_enrollment.py"}
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


def prepare_cari_excel_print_layout(workbook, logo_path):
    """Fit each report to A4 width and place one printable logo inside its bounds."""
    from io import BytesIO
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
    from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.units import pixels_to_EMU
    from PIL import Image

    summary = workbook.active
    logo = None
    if os.path.isfile(logo_path):
        with Image.open(logo_path) as source:
            logo = source.convert("RGBA")
        logo.putalpha(logo.getchannel("A").point(lambda alpha: round(alpha * 0.14)))

    for sheet in workbook.worksheets:
        is_summary = sheet is summary
        last_row = sheet.max_row if logo is None else max(sheet.max_row, 10)
        sheet.print_area = f"A1:{get_column_letter(sheet.max_column)}{last_row}"
        sheet.print_options.horizontalCentered = True
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.sheet_properties.pageSetUpPr.autoPageBreaks = False
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT if is_summary else sheet.ORIENTATION_LANDSCAPE
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 1 if is_summary else 0
        if not is_summary:
            sheet.print_title_rows = "1:1"
        if logo is None:
            continue

        width = sum(round(column.width * 7 + 5) for column in sheet.column_dimensions.values())
        header_height = sheet.row_dimensions[1].height * 96 / 72
        row_height = sheet.sheet_format.defaultRowHeight * 96 / 72
        height = min((last_row - 1) * row_height - 12, 190)
        scale = min(height / logo.height, width * 0.45 / logo.width)
        image_width, image_height = logo.width * scale, logo.height * scale
        png = BytesIO()
        logo.save(png, format="PNG")
        watermark = XLImage(png)
        watermark.anchor = AbsoluteAnchor(
            pos=XDRPoint2D(pixels_to_EMU((width - image_width) / 2), pixels_to_EMU(header_height + 6)),
            ext=XDRPositiveSize2D(pixels_to_EMU(image_width), pixels_to_EMU(image_height)),
        )
        sheet.add_image(watermark)


def open_pdf_in_viewer(path, *, print_dialog=False):
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
                    arguments = [exe, "/p", path] if print_dialog else [exe, path]
                    subprocess.Popen(arguments, close_fds=True)
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
CREATE TABLE IF NOT EXISTS personel_odeme_bilgisi (
 personel_id INTEGER PRIMARY KEY REFERENCES personel(id) ON DELETE CASCADE,
 iban TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS beton_satis (
 id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, cari_id INTEGER, santiye_id INTEGER,
 miktar_m3 REAL NOT NULL, birim_fiyat REAL NOT NULL, toplam REAL NOT NULL,
 pompa_ucreti REAL NOT NULL DEFAULT 0,
 mikser_nakliye_ucreti REAL NOT NULL DEFAULT 0,
 fatura_id INTEGER REFERENCES fatura(id) ON DELETE RESTRICT,
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
CREATE TABLE IF NOT EXISTS mastarci_hesap (
    id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, mastarci TEXT NOT NULL,
    santiye TEXT NOT NULL, miktar_m3 TEXT NOT NULL, birim_fiyat TEXT NOT NULL,
    hakedis_kurus INTEGER NOT NULL CHECK(hakedis_kurus>0),
    odenen_kurus INTEGER NOT NULL DEFAULT 0 CHECK(odenen_kurus BETWEEN 0 AND hakedis_kurus),
    aciklama TEXT NOT NULL, musteri TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS personel_puantaj (
    id INTEGER PRIMARY KEY, personel_id INTEGER NOT NULL REFERENCES personel(id),
    tarih TEXT NOT NULL, durum TEXT NOT NULL, baslangic TEXT NOT NULL,
    bitis TEXT NOT NULL, mola INTEGER NOT NULL, normal_dakika INTEGER NOT NULL,
    net_dakika INTEGER NOT NULL, mesai_dakika INTEGER NOT NULL,
    manuel_mesai_dakika INTEGER,
    UNIQUE(personel_id,tarih)
);
CREATE INDEX IF NOT EXISTS idx_puantaj_tarih ON personel_puantaj(tarih,personel_id);
CREATE TABLE IF NOT EXISTS puantaj_onay (
    id INTEGER PRIMARY KEY, personel_id INTEGER NOT NULL REFERENCES personel(id),
    donem TEXT NOT NULL, odeme_id INTEGER NOT NULL UNIQUE REFERENCES maas_odeme(id),
    maas TEXT NOT NULL, bolen TEXT NOT NULL, katsayi TEXT NOT NULL, mesai_dakika INTEGER NOT NULL,
    saatlik_mesai_ucreti TEXT,
    UNIQUE(personel_id,donem)
);
CREATE TRIGGER IF NOT EXISTS puantaj_cift_mesai_ekleme BEFORE INSERT ON maas_odeme
WHEN NEW.islem_turu='Mesai' AND EXISTS (
    SELECT 1 FROM puantaj_onay WHERE personel_id=NEW.personel_id AND donem=NEW.donem)
BEGIN SELECT RAISE(ABORT, 'Bu donemin mesaisi Puantaj ekraninda onaylandi.'); END;
CREATE TRIGGER IF NOT EXISTS puantaj_cift_mesai_degistirme BEFORE UPDATE ON maas_odeme
WHEN NEW.islem_turu='Mesai' AND EXISTS (
    SELECT 1 FROM puantaj_onay WHERE personel_id=NEW.personel_id AND donem=NEW.donem AND odeme_id<>OLD.id)
BEGIN SELECT RAISE(ABORT, 'Bu donemin mesaisi Puantaj ekraninda onaylandi.'); END;
CREATE TRIGGER IF NOT EXISTS puantaj_odeme_koru BEFORE UPDATE ON maas_odeme
WHEN EXISTS (SELECT 1 FROM puantaj_onay WHERE odeme_id=OLD.id)
BEGIN SELECT RAISE(ABORT, 'Bu mesai kaydini Puantaj ekranindan duzenleyin.'); END;

CREATE TABLE IF NOT EXISTS banka_hareket (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL, banka TEXT NOT NULL, hesap TEXT, yon TEXT NOT NULL, aciklama TEXT, tutar REAL NOT NULL CHECK(tutar>0), karsi_taraf TEXT NOT NULL DEFAULT '');
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
    from operation_undo import AUDITED_TABLES

    for table in (*AUDITED_TABLES, "personel_puantaj", "puantaj_onay", "mastarci_hesap"):
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
    hybrid_batch_id = None
    hybrid_slot = None

    def insert_shared(self, sql, values):
        if self.hybrid_slot is None:
            return self.execute(sql, values)
        from hybrid_identity import insert_shared

        return insert_shared(self, self.hybrid_slot, sql, values)

    def capture_batch_id(self):
        if self.hybrid_batch_id is None:
            self.hybrid_batch_id = secrets.token_hex(16)
        return self.hybrid_batch_id

    def commit(self):
        super().commit()
        self.hybrid_batch_id = None

    def rollback(self):
        super().rollback()
        self.hybrid_batch_id = None

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            return super().__exit__(exc_type, exc_value, traceback)
        finally:
            self.close()


def _enable_hybrid_capture(con):
    if os.environ.get("HOYTAS_HYBRID_CAPTURE") == "1":
        try:
            from hybrid_audit_bridge import install_capture
            from hybrid_identity import replica_slot

            con.hybrid_slot = replica_slot(con)
            install_capture(con, con.capture_batch_id)
        except Exception:
            con.close()
            raise
    elif con.execute("SELECT 1 FROM sqlite_master WHERE name='hybrid_replica_identity'").fetchone():
        con.close()
        raise RuntimeError("Bu veritabanı hibrit oturum ayarıyla açılmalıdır.")
    from operation_undo import enable_journal

    enable_journal(con)
    return con


def puantaj_manuel_mesai(saat):
    if not saat.strip():
        return None
    try:
        hours = Decimal(saat.strip().replace(",", "."))
    except ArithmeticError as error:
        raise ValueError("Mesai saatini sayı olarak girin.") from error
    if not hours.is_finite() or not 0 <= hours <= 24:
        raise ValueError("Mesai saati 0 ile 24 arasında olmalı.")
    return int((hours * 60).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def puantaj_sureleri(durum, baslangic, bitis, mola, normal, manuel_mesai=None):
    if durum in ("İzin", "Devamsız", "Tatil"):
        return 0, 0, 0, 0
    if durum != "Çalıştı":
        raise ValueError("Bir çalışma durumu seçin.")
    start, end = (datetime.strptime(value, "%H:%M") for value in (baslangic, bitis))
    duration = int((end - start).total_seconds() // 60)
    if duration < 0:
        duration += 24 * 60
    pause = int(mola)
    regular = int((Decimal(normal.replace(",", ".")) * 60).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if not 0 <= pause < duration or not 0 <= regular <= 24 * 60:
        raise ValueError("Saatleri, molayı ve normal çalışma süresini kontrol edin.")
    net = duration - pause
    if manuel_mesai is not None:
        regular_worked = min(net, regular)
        if regular_worked + manuel_mesai + pause > 24 * 60:
            raise ValueError("Normal çalışma, mola ve mesai toplamı 24 saati aşamaz.")
        return pause, regular, regular_worked + manuel_mesai, manuel_mesai
    return pause, regular, net, max(0, net - regular)


def _puantaj_ucret_degerleri(maas, bolen):
    salary, divisor = (Decimal(str(value).strip().replace(",", ".")) for value in (maas, bolen))
    if any(not value.is_finite() or value <= 0 for value in (salary, divisor)):
        raise ValueError("Maaş ve ücret böleni sıfırdan büyük olmalı.")
    return salary, divisor


def puantaj_ucret_bol(maas, bolen):
    salary, divisor = _puantaj_ucret_degerleri(maas, bolen)
    return salary / divisor


def _puantaj_mesai_katsayisi(katsayi):
    factor = Decimal(str(katsayi).strip().replace(",", "."))
    if not factor.is_finite() or factor <= 0:
        raise ValueError("Mesai katsayısı sıfırdan büyük olmalı.")
    return factor


def puantaj_mesai_saat_ucreti(maas, bolen, katsayi):
    return puantaj_ucret_bol(maas, bolen) * _puantaj_mesai_katsayisi(katsayi)


def puantaj_mesai_tutari(maas, dakika, bolen, katsayi):
    salary, divisor = _puantaj_ucret_degerleri(maas, bolen)
    factor = _puantaj_mesai_katsayisi(katsayi)
    return (salary * dakika * factor / (divisor * 60)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def puantaj_girilen_saat_ucreti(value):
    try:
        rate = Decimal(str(value).strip().replace(",", "."))
        if rate.is_finite() and rate > 0 and rate == rate.quantize(Decimal("0.01")):
            return rate
    except ArithmeticError:
        pass
    raise ValueError("Saatlik mesai ücretini TL olarak girin (ör. 240 veya 240,50).")


def puantaj_girilen_mesai_tutari(saatlik, dakika):
    return (puantaj_girilen_saat_ucreti(saatlik) * dakika / 60).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP)


def save_puantaj_approval(db, personel_id, first, divisor="", factor="", *, hourly=None, revoke=False):
    period = first.strftime("%Y-%m")
    existing = db.execute("SELECT odeme_id FROM puantaj_onay WHERE personel_id=? AND donem=?",
                          (personel_id, period)).fetchone()
    if revoke:
        if existing:
            db.execute("DELETE FROM puantaj_onay WHERE personel_id=? AND donem=?", (personel_id, period))
            db.execute("DELETE FROM maas_odeme WHERE id=?", existing)
        return
    if existing:
        return
    if db.execute("""SELECT 1 FROM maas_odeme WHERE personel_id=? AND donem=? AND islem_turu='Mesai'""",
                  (personel_id, period)).fetchone():
        raise ValueError("Bu dönem için mevcut bir mesai hareketi var. Önce o kaydı kontrol edin.")
    last = first.replace(day=monthrange(first.year, first.month)[1])
    overtime = db.execute("""SELECT COALESCE(SUM(mesai_dakika),0) FROM personel_puantaj
        WHERE personel_id=? AND tarih BETWEEN ? AND ?""",
        (personel_id, first.isoformat(), last.isoformat())).fetchone()[0]
    employee = db.execute("SELECT maas FROM personel WHERE id=?", (personel_id,)).fetchone()
    if employee is None:
        raise ValueError("Personel kaydı bulunamadı.")
    salary = employee[0]
    if hourly is None:
        amount = puantaj_mesai_tutari(salary, overtime, divisor, factor)
    else:
        hourly = str(puantaj_girilen_saat_ucreti(hourly))
        amount = puantaj_girilen_mesai_tutari(hourly, overtime)
    if amount <= 0:
        raise ValueError("Onay için pozitif bir mesai tutarı olmalı.")
    payment_id = db.execute("""INSERT INTO maas_odeme
        (tarih,personel_id,donem,tutar,yontem,aciklama,islem_turu)
        VALUES(?,?,?,?,'',?,'Mesai')""",
        (date.today().isoformat(), personel_id, period, str(amount),
         f"Puantaj · {overtime // 60} sa {overtime % 60} dk")).lastrowid
    db.execute("""INSERT INTO puantaj_onay
        (personel_id,donem,odeme_id,maas,bolen,katsayi,mesai_dakika,saatlik_mesai_ucreti)
        VALUES(?,?,?,?,?,?,?,?)""", (personel_id, period, payment_id, str(salary), divisor, factor, overtime, hourly))


def connect(path=None):
    path = DB_PATH if path is None else path
    existed = os.path.isfile(path) and os.path.getsize(path) > 0
    con = sqlite3.connect(path, timeout=3, factory=ClosingSQLiteConnection)
    con.row_factory = None
    con.create_function("hoytas_actor", 0, lambda: APP_CURRENT_USER)
    con.execute("PRAGMA busy_timeout = 3000")
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
    if existed and current_version == TARGET_SCHEMA_VERSION:
        return _enable_hybrid_capture(con)
    if current_version > TARGET_SCHEMA_VERSION:
        con.close()
        raise RuntimeError("Veritabanı daha yeni bir sürüme ait.")
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
        "cari": [("arsivli", "INTEGER NOT NULL DEFAULT 0 CHECK(arsivli IN (0,1))")],
        "cek": [("banka_kodu", "TEXT"), ("cari_id", "INTEGER")],
        "beton_satis": [("santiye_id", "INTEGER"), ("beton_cinsi", "TEXT NOT NULL DEFAULT ''"), ("pompa_ucreti", "REAL NOT NULL DEFAULT 0"), ("mikser_nakliye_ucreti", "REAL NOT NULL DEFAULT 0"), ("fatura_id", "INTEGER REFERENCES fatura(id) ON DELETE RESTRICT")],
        "kullanici": [("kurtarma_hash", "TEXT"), ("kurtarma_salt", "TEXT")],
        "banka_hareket": [("karsi_taraf", "TEXT NOT NULL DEFAULT ''")],
        "maas_odeme": [("islem_turu", "TEXT NOT NULL DEFAULT 'Maaş'")],
        "mastarci_hesap": [("musteri", "TEXT NOT NULL DEFAULT ''")],
        "personel_puantaj": [("manuel_mesai_dakika", "INTEGER")],
        "puantaj_onay": [("saatlik_mesai_ucreti", "TEXT")],
        "fatura": [("belge_bilgileri", "TEXT NOT NULL DEFAULT '{}'"), ("fatura_turu", "TEXT NOT NULL DEFAULT 'Normal'"), ("tevkifat_oran", "REAL NOT NULL DEFAULT 0"), ("tevkifat_tutar", "REAL NOT NULL DEFAULT 0"), ("tahsil_edilecek", "REAL NOT NULL DEFAULT 0"), ("odenen", "REAL NOT NULL DEFAULT 0")],
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
    if current_version < 52:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_banka_hareket_{action}")
    if current_version < 56:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_mastarci_hesap_{action}")
    if current_version < 57:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_personel_puantaj_{action}")
    if current_version < 58:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_puantaj_onay_{action}")
    if current_version < 59:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_fatura_{action}")
    if current_version < 60:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_cari_{action}")
    con.execute("CREATE INDEX IF NOT EXISTS idx_beton_fatura ON beton_satis(fatura_id)")
    if current_version < 62:
        for action in ("insert", "update", "delete"):
            con.execute(f"DROP TRIGGER IF EXISTS audit_beton_satis_{action}")
    _install_audit_triggers(con)
    from operation_undo import create_journal

    create_journal(con)
    # Şema sürümünü ilan etmeden önce veritabanının fiziksel bütünlüğünü doğrula.
    integrity = con.execute("PRAGMA integrity_check").fetchone()
    if not integrity or integrity[0] != "ok":
        con.close()
        raise RuntimeError(f"HOYTAŞ veritabanı bütünlük kontrolünden geçemedi: {integrity}")
    # Sürüm numarası ancak bütün yükseltmeler ve bütünlük kontrolü başarıyla tamamlandıktan sonra yazılır.
    con.execute(f"PRAGMA user_version = {TARGET_SCHEMA_VERSION}")
    return _enable_hybrid_capture(con)

CHEQUE_STATES = ("Bekliyor", "Portföyde", "Tahsilde", "Ciro Edildi", "Tahsil Edildi", "Ödendi", "İade", "Karşılıksız")


CHEQUE_STATUS_REPORTS = {
    "Tahsildeki Çekler Dökümü": "collection",
    "Kendi Çeklerimiz Dökümü": "own",
    "Ciro Edilen Çekler Dökümü": "endorsed",
    "Portföydeki Çekler Dökümü": "portfolio",
}


def cheque_matches_view(direction, status, view):
    if view == "endorsed":
        return status == "Ciro Edildi"
    if view in ("received", "own"):
        return direction == ("Alınan" if view == "received" else "Verilen")
    if view == "collection":
        return direction == "Alınan" and status == "Tahsilde"
    if view == "portfolio":
        return direction == "Alınan" and status in ("Bekliyor", "Portföyde")
    return True


def normalize_personel_iban(value):
    iban = "".join(value.split()).upper()
    if not iban:
        return ""
    if not re.fullmatch(r"TR[0-9]{24}", iban):
        raise ValueError("IBAN, TR ile başlamalı ve toplam 26 karakter olmalıdır.")
    if int(iban[4:] + "2927" + iban[2:4]) % 97 != 1:
        raise ValueError("IBAN kontrolü başarısız. Bankadaki IBAN'ı kontrol edin.")
    return iban


def save_personel_iban(db, personel_id, iban):
    if iban:
        db.execute("""INSERT INTO personel_odeme_bilgisi(personel_id,iban) VALUES (?,?)
            ON CONFLICT(personel_id) DO UPDATE SET iban=excluded.iban""", (personel_id, iban))
    else:
        db.execute("DELETE FROM personel_odeme_bilgisi WHERE personel_id=?", (personel_id,))


def normalize_amount_text(value):
    """Normalize Turkish and international numeric formats."""
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
    return sign + normalized


def parse_amount(value):
    import math
    amount = float(normalize_amount_text(value))
    if not math.isfinite(amount):
        raise ValueError("Geçersiz tutar.")
    return amount


def mastarci_hakedis(quantity, price):
    quantity, price = (Decimal(normalize_amount_text(value)) for value in (quantity, price))
    total = int((quantity * price * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if quantity <= 0 or price <= 0 or not 0 < total < 2**63:
        raise ValueError("Metreküp, birim fiyat ve hakediş tutarını kontrol edin.")
    return quantity, price, total


def mastarci_odeme_kurus(value):
    amount = Decimal(normalize_amount_text(value)) * 100
    if not 0 < amount < 2**63 or amount != amount.to_integral_value():
        raise ValueError("Ödemeyi sıfırdan büyük ve en fazla iki ondalık hane ile girin.")
    return int(amount)


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
    sql = "SELECT id,unvan FROM cari"
    if allowed_types:
        sql += " WHERE tur IN (" + ",".join("?" for _ in allowed_types) + ")"
    rows = connection.execute(sql + " ORDER BY id", allowed_types or ()).fetchall()
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
    if not match:
        match = re.match(r"Beton satış #(\d+)(?!\d)", str(description or ""), flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def concrete_sale_totals(m3_value, price_value="", pump_value="0", transport_value="0"):
    """M³ is mandatory; an omitted unit price means a zero-value cari sale."""
    miktar = parse_amount(m3_value)
    fiyat = parse_amount(price_value or "0")
    pompa = parse_amount(pump_value or "0")
    nakliye = parse_amount(transport_value or "0")
    if miktar <= 0 or min(fiyat, pompa, nakliye) < 0:
        raise ValueError("Metreküp sıfırdan büyük; fiyat, pompa ve mikser nakliye ücreti geçerli ve negatif olmayan değerler olmalı.")
    return miktar, fiyat, pompa, nakliye, round(miktar * fiyat + pompa + nakliye, 2)


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
                          santiye="", notlar="", odeme_durumu="", mikser_nakliye_ucreti=0):
    """Beton satışını her ekranda aynı, eksiksiz ve okunaklı dille açıkla."""
    m3 = safe_float(miktar or 0)
    fiyat = safe_float(birim_fiyat or 0)
    pompa = safe_float(pompa_ucreti or 0)
    nakliye = safe_float(mikser_nakliye_ucreti or 0)
    genel = safe_float(toplam) if toplam is not None else round(m3 * fiyat + pompa + nakliye, 2)
    parts = [f"{(cinsi or 'Hazır beton').strip()} beton",
             f"{format_m3(m3)} m³",
             f"{fiyat:,.2f} TL/m³"]
    parts.append(f"Pompa: {pompa:,.2f} TL (Var)" if pompa > 0.005 else "Pompa: Yok")
    if nakliye:
        parts.append(f"Mikser nakliye: {nakliye:,.2f} TL")
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
        COALESCE(pompa_ucreti,0),toplam,COALESCE(aciklama,''),COALESCE(odeme_durumu,''),mikser_nakliye_ucreti
        FROM beton_satis WHERE id=?""", (sale_id,)).fetchone()
    if not sale:
        raise ValueError(f"Beton satış kaydı bulunamadı: {sale_id}")
    tarih,cari_id,cinsi,m3,fiyat,pompa,toplam,notlar,durum,nakliye = sale
    linked = [(mid,) for mid, desc in db.execute(
        "SELECT id,aciklama FROM hareket WHERE kategori='Satış' AND tur='Borç' ORDER BY id")
        if linked_beton_sale_id(desc) == sale_id]
    if not cari_id:
        for (movement_id,) in linked:
            db.execute("DELETE FROM hareket WHERE id=?", (movement_id,))
        return None
    description = format_beton_aciklama(cinsi,m3,fiyat,pompa,toplam,notlar=notlar,odeme_durumu=durum,mikser_nakliye_ucreti=nakliye)
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
    cur = db.insert_shared("""INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi)
        VALUES(?,?,?,?,?,?,?)""", (tarih,cari_id,"Borç","Satış",description,toplam,"Vadeli"))
    return cur.lastrowid



def index_beton_movements(db):
    linked, legacy = {}, set()
    for row in db.execute("""SELECT id,cari_id,tarih,tutar,COALESCE(aciklama,'')
        FROM hareket WHERE tur='Borç' AND kategori='Satış' ORDER BY id"""):
        sale_id = linked_beton_sale_id(row[4])
        if sale_id is None:
            legacy.add(row[1])
        else:
            linked.setdefault(sale_id, []).append(row)
    return linked, legacy


def audit_beton_cari_amounts(db, cari_id, movement_index=None):
    """Read-only reconciliation: distinguish absent, mismatched and ambiguous links.

    A sale with a linked zero-amount movement is NOT a missing sale.
    Never infer that an unrelated manual debit belongs to a sale.
    """
    sales = db.execute("SELECT id,tarih,toplam FROM beton_satis WHERE cari_id=? ORDER BY tarih,id", (cari_id,)).fetchall()
    linked, legacy = index_beton_movements(db) if movement_index is None else movement_index
    missing, mismatched, ambiguous = [], [], []
    for sale_id, date, amount in sales:
        matches = linked.get(sale_id, ())
        if len(matches) == 1:
            m = matches[0]
            if m[1] != cari_id or m[2] != date or abs(safe_float(m[3])-safe_float(amount)) > .009:
                mismatched.append((sale_id, date, safe_float(amount), m[0], safe_float(m[3]), m[1]))
        elif len(matches) > 1:
            ambiguous.append((sale_id, date, safe_float(amount), 'Çoklu bağlantı'))
        else:
            if cari_id in legacy:
                ambiguous.append((sale_id, date, safe_float(amount), 'Kimliği belirsiz eski hareket'))
            else:
                missing.append((sale_id, date, safe_float(amount)))
    return missing, mismatched, ambiguous


def find_missing_beton_movements(db, cari_id):
    """Return unlinked sales; avoid creating a duplicate for plausible legacy entries."""
    missing, _, ambiguous = audit_beton_cari_amounts(db, cari_id)
    return missing, [row[:3] for row in ambiguous]


def audit_cari(db, cid, name, movement_index=None):
    missing, wrong, ambiguous = audit_beton_cari_amounts(db, cid, movement_index)
    debt, credit = db.execute("""SELECT COALESCE(SUM(CASE WHEN tur='Borç' THEN tutar ELSE 0 END),0),
        COALESCE(SUM(CASE WHEN tur='Alacak' THEN tutar ELSE 0 END),0)
        FROM hareket WHERE cari_id=?""", (cid,)).fetchone()
    view = db.execute("SELECT toplam_borc,toplam_alacak,bakiye FROM cari_bakiye WHERE cari_id=?", (cid,)).fetchone()
    balance_error = not view or any(abs(safe_float(a)-safe_float(b)) > .009
        for a,b in zip(view or (), (debt,credit,debt-credit)))
    delta = round(sum(x[2] for x in missing)+sum(x[2]-x[4] for x in wrong), 2)
    return cid, name, missing, wrong, ambiguous, balance_error, delta


def audit_all_cari(db):
    """Read-only bulk audit, including closed and credit-balance accounts."""
    movement_index = index_beton_movements(db)
    return [audit_cari(db, cid, name, movement_index)
            for cid, name in db.execute("SELECT id,unvan FROM cari ORDER BY unvan,id")]


def repairable_cari_movements(audit):
    cid, _, _, wrong, _, balance_error, _ = audit
    return [movement for movement in wrong if movement[5] == cid and not balance_error]


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
        if os.name == "nt":
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        super().__init__()
        self.title(f"HOYTAŞ Hazır Beton | Finans ve İşletme Yönetimi v{APP_VERSION}")
        icon = os.path.join(APP_DIR, "HOYTAS.ico")
        if os.name == "nt" and os.path.isfile(icon):
            try:
                self.iconbitmap(default=icon)
            except tk.TclError:
                pass
        self.geometry(f"{min(1440,self.winfo_screenwidth()-32)}x{min(900,self.winfo_screenheight()-80)}")
        self.minsize(800, 550)
        self.configure(bg="#F3F5F4")
        self.membership_client = None
        endpoint = os.environ.get("HOYTAS_MEMBERSHIP_URL", "").strip()
        if endpoint and ISOLATED_PROFILE_DIR:
            from hybrid_membership_client import MembershipClient
            self.membership_client = MembershipClient(endpoint)
            self.withdraw()
            self._central_login()
        else:
            self._ensure_admin_and_login()
        self.style = ttk.Style(self)
        try: self.style.theme_use("clam")
        except tk.TclError: pass
        # Kurumsal HOYTAŞ arayüzü: sakin gri zemin, koyu yeşil ve kontrollü sarı vurgu.
        self.style.configure("Brand.TLabel", background="#F7FAF1", foreground="#19220E")
        self.style.configure("TFrame", background="#F4F7ED")
        self.style.configure("TLabel", background="#F4F7ED", foreground="#24352C", font=("Segoe UI", 10))
        self.style.configure("TButton", padding=(12, 8), font=("Segoe UI", 9, "bold"),
                             background="#315C16", foreground="#FFFFFF", borderwidth=0, focusthickness=2)
        self.style.map("TButton", background=[("active", "#264810"), ("pressed", "#1F3412"), ("disabled", "#D5DCD7")],
                       foreground=[("disabled", "#7C8780")])
        self.style.configure("TMenubutton", padding=(10, 7), font=("Segoe UI", 9, "bold"))
        self.style.configure("Selected.TMenubutton", background="#DEF2C4", foreground="#1F3412")
        self.style.configure("Selected.TButton", background="#DEF2C4", foreground="#1F3412")
        self.style.configure("TEntry", padding=7, fieldbackground="#FFFFFF", foreground="#24352C", bordercolor="#C9D5CD")
        self.style.configure("TCombobox", padding=7, fieldbackground="#FFFFFF", foreground="#24352C", arrowsize=14)
        self.style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], foreground=[("readonly", "#24352C")])
        self.style.configure("TNotebook", background="#E8EEEA", borderwidth=0, tabmargins=(3, 7, 3, 0))
        self.style.configure("TNotebook.Tab", padding=(7, 8), font=("Segoe UI", 9, "bold"),
                             background="#E5EEDB", foreground="#40564A", borderwidth=0)
        self.style.map("TNotebook.Tab", background=[("selected", "#FFFFFF"), ("active", "#D6EAC5")],
                       foreground=[("selected", "#1D5B39"), ("active", "#1F3028")], padding=[("selected", (7, 8))])
        self.style.configure("Treeview", rowheight=34, font=("Segoe UI", 9), background="#FFFFFF",
                             fieldbackground="#FFFFFF", foreground="#24352C", borderwidth=0)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                             background="#DCE8E0", foreground="#214B35", padding=(9, 11), relief="flat")
        self.style.map("Treeview", background=[("selected", "#DEF2C4")], foreground=[("selected", "#1F3412")])
        self.style.configure("TLabelframe", background="#FFFFFF", bordercolor="#D5E0D8", relief="solid", padding=8)
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#315C16", font=("Segoe UI", 9, "bold"))
        self.style.configure("TLabelframe", background="#FFFFFF", bordercolor="#D8E0DB", relief="solid")
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#315B46", font=("Segoe UI", 9, "bold"))
        self.style.configure("Metric.TLabelframe", background="#FFFFFF", bordercolor="#D8E0DB", relief="solid")
        self.style.configure("Metric.TLabel", background="#FFFFFF", foreground="#1F3028", font=("Segoe UI", 18, "bold"))
        self.style.map("CariCard.TNotebook.Tab", background=[("selected", "#DCEBDB")],
                       foreground=[("selected", "#174D30")])
        self.style.configure("CariSummary.TFrame", background="#E9F0E5")
        self.style.configure("CariSummary.TLabel", background="#E9F0E5", foreground="#174D30")
        self._build()
        self._adapt_action_rows(self)
        # Other modules load when opened. Avoid repeating the initial dashboard refresh.
        self._last_selected_tab = self.tabs.select()
        self.tabs.bind("<<NotebookTabChanged>>", self._refresh_on_tab_change, add="+")
        # Klavye kısayolları: sekmeler, cari arama ve kayıt formları.
        self.bind_all("<Control-f>", self._shortcut_cari_search)
        self.bind_all("<Control-F>", self._shortcut_cari_search)
        self.bind_all("<F5>", self._refresh_from_button)
        self.bind_all("<Control-r>", self._refresh_from_button)
        self.bind_all("<Control-R>", self._refresh_from_button)
        for _i in range(1, 10):
            self.bind_all(f"<Alt-{_i}>", lambda e, i=_i: self._shortcut_tab(i))
        self._refresh_selected_module()
        if self.membership_client is not None:
            self.deiconify()
        if os.name == "nt":
            self.state("zoomed")
        # Sunucu tanımlıysa açılışta sessizce yeni sürüm kontrolü yap.
        self.after(1800, self.check_for_updates_silent)

    def _adapt_action_rows(self, root):
        """Wrap button-only packed toolbars without altering their commands."""
        # Reserve all fixed-height controls before allocating remaining space to tables.
        packed = root.pack_slaves()
        first = next((w for w in packed if w.pack_info()["expand"] and
                      w.pack_info()["fill"] == "both"), None)
        if first is not None:
            for widget in reversed(packed[packed.index(first) + 1:]):
                info = widget.pack_info()
                if not info["expand"] and info["side"] in ("top", "bottom"):
                    widget.pack_configure(side="bottom", before=first)
        for frame in root.winfo_children():
            if isinstance(frame, ttk.Frame):
                children = frame.winfo_children()
                if len(children) >= 2 and all(isinstance(w, (ttk.Button, ttk.Menubutton)) and
                                              w.winfo_manager() == "pack" for w in children):
                    for button in children:
                        button.pack_forget()

                    def arrange(event=None, bar=frame, buttons=tuple(children)):
                        width = bar.winfo_width()
                        if width <= 1:
                            return
                        row_height = max(button.winfo_reqheight() for button in buttons) + 4
                        x = y = 0
                        for button in buttons:
                            need = button.winfo_reqwidth() + 6
                            if x and x + need > width:
                                x = 0
                                y += row_height
                            button.place(x=x + 3, y=y + 2)
                            x += need
                        height = y + row_height
                        if int(bar.cget("height")) != height:
                            bar.configure(height=height)

                    frame.configure(height=max(w.winfo_reqheight() for w in children) + 4)
                    frame.bind("<Configure>", arrange, add="+")
                    frame.after_idle(arrange)
                else:
                    self._adapt_action_rows(frame)
            elif isinstance(frame, ttk.Notebook):
                self._adapt_action_rows(frame)

    def _refresh_on_tab_change(self, event=None):
        """Açılan modülü güncel veritabanı kayıtlarıyla doldur."""
        selected = self.tabs.select()
        if selected == self._last_selected_tab:
            return
        self._last_selected_tab = selected
        pending = getattr(self, "_tab_refresh_job", None)
        if pending is not None:
            self.after_cancel(pending)
        self._tab_refresh_job = self.after(180, self._run_tab_refresh)

    def _run_tab_refresh(self):
        self._tab_refresh_job = None
        if getattr(self, "_refreshing_tab_change", False):
            return
        self._refreshing_tab_change = True
        try:
            self._refresh_selected_module()
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
                self._open_module(self.cari_tab)
                self.cari_search_entry.focus_set()
                self.cari_search_entry.selection_range(0, "end")
                return "break"
            except tk.TclError:
                pass

    def _shortcut_tab(self, number):
        if hasattr(self, "tabs") and number <= len(self.tabs.tabs()):
            self._open_module(self.tabs.tabs()[number - 1])
            return "break"

    def _hash_password(self, password, salt):
        return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 180000).hex()

    def _central_login(self):
        global APP_CURRENT_USER
        from hybrid_desktop_membership import sign_in

        from pathlib import Path
        from hybrid_enrollment import install_baseline
        from hybrid_identity import replica_slot
        from hybrid_start_live import readonly

        database = Path(DB_PATH)
        setup_admin = False
        if database.is_file():
            with readonly(database) as db:
                setup_admin = replica_slot(db) == 1
        def prepare(account):
            from hybrid_http_relay import HTTPRelay

            if not database.is_file():
                legacy = database.parent.parent / "hoytas.db"
                session = account["session"]
                relay = HTTPRelay(self.membership_client.url, session)
                install_baseline(database, legacy, self.membership_client, session, connect, relay)
            with connect() as db:
                if replica_slot(db) is None:
                    raise RuntimeError("Canlı profil kimliği bulunamadı.")

        account = sign_in(self, self.membership_client, setup_admin=setup_admin, prepare=prepare)
        if account is None:
            self.destroy()
            raise SystemExit
        self.membership_session = account["session"]
        role = "Yönetici" if account["role"] == "admin" else "Kullanıcı"
        self.current_user = {"username": account["username"], "role": role}
        APP_CURRENT_USER = account["username"]

    def _ensure_admin_and_login(self):
        from login_memory import LoginFields

        with connect() as c:
            count = c.execute("SELECT COUNT(*) FROM kullanici").fetchone()[0]
        if count == 0:
            while True:
                win = tk.Toplevel(self); win.title("HOYTAŞ Hazır Beton - İlk Kurulum"); win.transient(self); win.grab_set(); win.resizable(False, False)
                frm = ttk.Frame(win, padding=22); frm.pack()
                brand = ttk.Label(frm, text="Yönetici Hesabı", font=("Segoe UI", 15, "bold"))
                set_brand_logo(brand, compound="top")
                brand.grid(row=0, column=0, columnspan=2, pady=(0, 12))
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
        failed_attempts = 0
        locked_until = 0.0
        while True:
            win=tk.Toplevel(self); win.title("HOYTAŞ Hazır Beton - Giriş"); win.transient(self); win.grab_set(); win.resizable(False,False)
            frm=ttk.Frame(win,padding=24); frm.pack()
            brand = ttk.Label(frm, text="HOYTAŞ Hazır Beton", font=("Segoe UI", 22, "bold"))
            set_brand_logo(brand)
            brand.grid(row=0, column=0, columnspan=2, pady=(0, 3))
            ttk.Label(frm,text="Finans Yönetimi - Güvenli Giriş").grid(row=1,column=0,columnspan=2,pady=(0,14))
            ttk.Label(frm,text="Kullanıcı adı").grid(row=2,column=0,sticky="w",pady=5)
            ttk.Label(frm,text="Şifre").grid(row=3,column=0,sticky="w",pady=5)
            fields = LoginFields(frm, 2, "yerel")
            user, pw = fields.username, fields.password
            result={"ok":False}
            def login():
                global APP_CURRENT_USER
                nonlocal failed_attempts, locked_until
                pw.configure(show="*")
                if time.monotonic() < locked_until:
                    messagebox.showwarning("Biraz bekleyin", "Çok fazla başarısız giriş oldu. 30 saniye sonra tekrar deneyin.", parent=win)
                    return
                try:
                    with connect() as c:
                        row=c.execute("SELECT id,kullanici_adi,parola_hash,salt,rol,aktif FROM kullanici WHERE kullanici_adi=?",(user.get().strip(),)).fetchone()
                    if row and hmac.compare_digest(self._hash_password(pw.get(),row[3]),row[2]):
                        if row[5] != 1:
                            status = "Üyeliğiniz yönetici onayı bekliyor." if row[5] == -1 else "Hesabınız devre dışı."
                            messagebox.showwarning("Giriş yapılamadı", status, parent=win)
                            return
                        self.current_user={"id":row[0],"username":row[1],"role":row[4]}
                        APP_CURRENT_USER = row[1]
                        fields.save(row[1], pw.get(), fields.remember.get())
                        result["ok"]=True
                        win.destroy()
                    else:
                        failed_attempts += 1
                        if failed_attempts == 5:
                            failed_attempts = 0
                            locked_until = time.monotonic() + 30
                        messagebox.showerror("Giriş başarısız","Kullanıcı adı veya şifre yanlış.",parent=win)
                except Exception:
                    detail = traceback.format_exc()
                    try:
                        with open(os.path.join(USER_DATA_DIR, "program_hatalari.log"), "a", encoding="utf-8") as log:
                            log.write("\n--- GIRIS ISLEMI HATASI ---\n" + detail + "\n")
                    except OSError:
                        pass
                    messagebox.showerror("Giriş sırasında hata", "Giriş işlemi tamamlanamadı. Verileriniz silinmedi.\nAyrıntı: " + str(detail.splitlines()[-1] if detail.splitlines() else "Bilinmeyen hata"), parent=win)
            ttk.Button(frm,text="Giriş Yap",command=login).grid(row=5,column=0,columnspan=2,pady=(14,0))
            links=ttk.Frame(frm); links.grid(row=6,column=0,columnspan=2,pady=(8,0))
            ttk.Button(links,text="Yeni Üye Ol",command=self.signup_dialog).pack(side="left",padx=4)
            ttk.Button(links,text="Şifremi Unuttum",command=self.forgot_password_dialog).pack(side="left",padx=4)
            pw.bind("<Return>",lambda e:login()); user.focus_set(); self.wait_window(win)
            if result["ok"]: break
            messagebox.showwarning("Giriş gerekli","Giriş yapmadan program açılamaz."); self.destroy(); raise SystemExit

    def signup_dialog(self):
        d=self.prompt("HOYTAŞ Üyelik Oluştur",[("user","Kullanıcı adı","entry",None),("password","Şifre (en az 12 karakter)","secret",None),("recovery","Kurtarma kelimesi (en az 12 karakter)","secret",None)])
        if not d: return
        if not d["user"].strip() or len(d["password"])<12 or len(d["recovery"].strip())<12:
            return messagebox.showerror("Eksik bilgi","Kullanıcı adı, en az 12 karakter şifre ve en az 12 karakter kurtarma bilgisi gerekli.")
        salt=secrets.token_hex(16); rs=secrets.token_hex(16)
        try:
            with connect() as c: c.execute("INSERT INTO kullanici(kullanici_adi,parola_hash,salt,rol,kurtarma_hash,kurtarma_salt,aktif) VALUES(?,?,?,?,?,?,-1)",(d["user"].strip(),self._hash_password(d["password"],salt),salt,"Kullanıcı",self._hash_password(d["recovery"].strip().casefold(),rs),rs))
        except sqlite3.IntegrityError: return messagebox.showerror("Kullanıcı mevcut","Bu kullanıcı adı zaten kayıtlı.")
        messagebox.showinfo("Üyelik oluşturuldu","Üyelik isteği kaydedildi. Yönetici onayından sonra giriş yapabilirsiniz.")

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

    def add_gelen_fatura(self, cari_context=None):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari WHERE tur IN ('Tedarikçi','Diğer') OR tur='Müşteri' ORDER BY unvan").fetchall()
        if not cariler: return messagebox.showinfo("Tedarikçi gerekli","Önce Cari Hesaplar bölümünden tedarikçi ekleyin.",parent=self)
        names=[f"{i} - {n}" for i,n in cariler]
        d=self._prompt_for_cari(cari_context, "Gelen Alış Faturası",[("tarih","Fatura tarihi (GG.AA.YYYY)","entry",None),("vade","Ödeme vadesi (GG.AA.YYYY)","entry",None),("cari","Tedarikçi","combo",names),("belge","Fatura numarası","entry",None),("aciklama","Açıklama / ürün-hizmet","entry",None),("ara","KDV hariç tutar (TL)","entry",None),("kdv","KDV oranı (%)","combo",["0","1","10","20"])])
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
                c.insert_shared("INSERT INTO gelen_fatura(tarih,vade,cari_id,belge_no,aciklama,ara_toplam,kdv_oran,kdv_tutar,genel_toplam,odenen,durum) VALUES(?,?,?,?,?,?,?,?,?,0,'Ödenmedi')",(tarih,vade,cid,d["belge"].strip(),d["aciklama"],ara,oran,kdv,toplam))
            self._refresh_selected_module(); messagebox.showinfo("Kaydedildi",f"Gelen fatura kaydedildi.\\nGenel toplam: {toplam:,.2f} TL\\nÖdeme durumu: Ödenmedi",parent=self)
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
                c.insert_shared("INSERT INTO gelen_fatura_odeme(fatura_id,tarih,tutar,yontem,aciklama) VALUES(?,?,?,?,?)",(rid,dt,amount,d["yontem"],d["aciklama"]))
                c.execute("UPDATE gelen_fatura SET odenen=?,durum=? WHERE id=?",(newpaid,status,rid))
                c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(dt,row[0],"Gider","Alış Faturası Ödemesi",ref,amount,d["yontem"]))
                if d["yontem"]=="Nakit": c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",(dt,"Para Çıkışı","Gelen Fatura",ref,amount,getattr(self,"current_user",{}).get("username","")))
                elif d["yontem"]=="Banka": c.insert_shared("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",(dt,"Banka","", "Para Çıkışı",ref,amount))
            self._refresh_selected_module(); messagebox.showinfo("Ödeme kaydedildi",f"Ödeme: {amount:,.2f} TL\\nKalan: {round(kalan-amount,2):,.2f} TL\\nDurum: {status}",parent=self)
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
            self._refresh_selected_module()
        except Exception as e:messagebox.showerror("Düzenleme başarısız",str(e),parent=self)

    def _build_fatura(self):
        bar=ttk.Frame(self.fatura_tab); bar.pack(fill="x")
        ttk.Label(bar,text="Fatura / Belge Taslakları",font=("Segoe UI",16,"bold")).pack(side="left")
        ttk.Button(bar,text="Yeni Fatura Taslağı",command=self.add_fatura).pack(side="right")
        ttk.Label(self.fatura_tab, text="Satış taslakları cari karta bağlıdır. Önizleme ve cari kontrolü yerel olarak yapılır.",
                  foreground="#7a4b00", wraplength=900).pack(anchor="w", pady=(8, 0))
        row = ttk.Frame(self.fatura_tab)
        row.pack(fill="x", pady=8)
        self.fatura_tree=self._tree(self.fatura_tab,("id","tarih","cari","belge","miktar","ara","kdv","toplam","tevkifat","net","odenen","kalan","durum"),("ID","Tarih","Cari","Belge No","m³","Ara Toplam TL","KDV TL","Genel Toplam TL","Tevkifat TL","Net TL","Tahsil TL","Kalan TL","Durum"),(45,100,200,120,70,120,100,120,110,110,110,110,125))
        ttk.Button(row, text="Cari Kontrolü",
                   command=lambda: self.selected_invoice_document(self.fatura_tree, review=True)).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="Taslak PDF", command=lambda: self.selected_invoice_document(self.fatura_tree)).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="Belge Bilgileri", command=lambda: self.selected_invoice_document(self.fatura_tree, edit=True)).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="Taslağı Düzenle", command=lambda: self.edit_selected(self.fatura_tree, "fatura")).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="Tahsilat Al", command=self.tahsilat_fatura).pack(side="left")
        ttk.Button(row, text="Taslağı Sil", command=lambda: self.delete_selected(self.fatura_tree, "fatura")).pack(side="right")

    def _build(self):
        header = tk.Frame(self, bg="#174D30")
        header.pack(fill="x")
        logo = ttk.Label(header, text="HOYTAŞ", style="Brand.TLabel")
        set_brand_logo(logo, max_size=42)
        logo.pack(side="left", padx=(20, 14), pady=6)
        tk.Label(header, text="HOYTAŞ | Finans ve İşletme", bg="#174D30", fg="#FFFFFF",
                 font=("Segoe UI", 13, "bold")).pack(side="left", pady=8)
        who = getattr(self, "current_user", {})
        user_text = f"KULLANICI  {who.get('username', '—')}   |   {who.get('role', '')}"
        tk.Label(header, text=user_text, bg="#174D30", fg="#FFFFFF",
                 font=("Segoe UI", 9)).pack(side="right", padx=24)
        tk.Frame(self, bg="#56BA00", height=3).pack(fill="x")
        self.style.layout("Workspace.TNotebook.Tab", [])
        self.tabs = ttk.Notebook(self, style="Workspace.TNotebook")
        self.tabs.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.dashboard = ttk.Frame(self.tabs, padding=18)
        self.cari_tab = ttk.Frame(self.tabs, padding=12)
        self.borc_takip_tab = ttk.Frame(self.tabs, padding=12)
        self.alisveris_cari_tab = ttk.Frame(self.tabs, padding=12)
        self.hareket_tab = ttk.Frame(self.tabs, padding=12)
        self.cek_tab = ttk.Frame(self.tabs, padding=12)
        self.stok_tab = ttk.Frame(self.tabs, padding=12)
        self.personel_tab = ttk.Frame(self.tabs, padding=12)
        self.beton_tab = ttk.Frame(self.tabs, padding=12)
        self.mastarci_tab = ttk.Frame(self.tabs, padding=12)
        self.fatura_tab = ttk.Frame(self.tabs, padding=12)
        self.gelen_fatura_tab = ttk.Frame(self.tabs, padding=12)
        self.kasa_tab = ttk.Frame(self.tabs, padding=12)
        self.stok_hareket_tab = ttk.Frame(self.tabs, padding=12)
        self.maas_tab = ttk.Frame(self.tabs, padding=12)
        self.banka_tab = ttk.Frame(self.tabs, padding=12)
        self.cari_card_page = ttk.Frame(self.tabs)
        # Keep tab indices stable for the existing keyboard shortcuts.
        modules = {
            self.dashboard: "Genel Bakış", self.cari_tab: "Cari Hesaplar",
            self.borc_takip_tab: "Borç Takip", self.alisveris_cari_tab: "Tedarikçi Cari",
            self.hareket_tab: "Gelir / Gider", self.cek_tab: "Çek İşlemleri",
            self.stok_tab: "Stok / Depo", self.personel_tab: "Personel",
            self.beton_tab: "Beton Satışları", self.mastarci_tab: "Mastarcı Hesabı",
            self.fatura_tab: "Satış Faturaları", self.gelen_fatura_tab: "Alış Faturaları",
            self.kasa_tab: "Nakit Kasa", self.stok_hareket_tab: "Stok Hareketleri",
            self.maas_tab: "Maaş Ödemeleri", self.banka_tab: "Banka İşlemleri",
            self.cari_card_page: "Cari Kartı",
        }
        for frame, title in modules.items():
            self.tabs.add(frame, text=title)
            if frame != self.dashboard:
                self.tabs.hide(frame)
        nav = ttk.Frame(self)
        nav.pack(fill="x", before=self.tabs, padx=10, pady=(8, 0))
        self.style.configure("Navigation.TButton", padding=(8, 4))
        self.style.configure("Selected.Navigation.TButton", background="#DEF2C4", foreground="#1F3412")
        self.style.configure("Navigation.TMenubutton", padding=(8, 4), arrowsize=9)
        self.style.configure("Selected.Navigation.TMenubutton", background="#DEF2C4", foreground="#1F3412")
        overview = ttk.Button(nav, text="Ana Sayfa", style="Navigation.TButton",
                              command=lambda: self._open_module(self.dashboard))
        self.module_navigation = nav
        nav_items = [overview]
        nav_groups = {str(self.dashboard): overview}
        for label, frames in (
            ("Cari", (self.cari_tab, self.borc_takip_tab, self.alisveris_cari_tab)),
            ("Beton", (self.beton_tab, self.mastarci_tab)),
            ("Çekler", (self.cek_tab,)),
            ("Kasa ve Banka", (self.kasa_tab, self.banka_tab, self.hareket_tab)),
            ("Faturalar", (self.fatura_tab, self.gelen_fatura_tab)),
            ("Personel", (self.personel_tab, self.maas_tab)),
            ("Diğer", (self.stok_tab, self.stok_hareket_tab)),
        ):
            button = self._action_menu(nav, label, [
                (modules[frame], lambda page=frame: self._open_module(page)) for frame in frames
            ])
            button.configure(direction="below", style="Navigation.TMenubutton", width=0,
                             takefocus=True)
            nav_items.append(button)
            nav_groups.update((str(frame), button) for frame in frames)

        nav_groups[str(self.cari_card_page)] = nav_groups[str(self.cari_tab)]
        nav.configure(height=overview.winfo_reqheight() + 4)
        nav.bind("<Configure>", lambda _event: self._layout_module_navigation(nav, nav_items))

        self.module_strip = ttk.Frame(self)
        self.module_strip.pack(fill="x", before=self.tabs, padx=10, pady=(8, 0))
        self.module_controls = {}
        for frame, title in modules.items():
            tab = ttk.Frame(self.module_strip)
            select = ttk.Button(tab, text=title, style="Workspace.TButton",
                                command=lambda page=frame: self._open_module(page))
            select.grid(row=0, column=0, sticky="ns")
            close = None
            if frame != self.dashboard:
                close = ttk.Button(tab, text="×", width=1, style="Close.Workspace.TButton",
                                   command=lambda page=frame: self._close_module(page=page))
                close.grid(row=0, column=1, sticky="ns")
            self.module_controls[str(frame)] = (tab, select, close)
        self.style.configure("Workspace.TButton", padding=(8, 5), background="#E5EEDB", foreground="#263519")
        self.style.configure("Current.Workspace.TButton", padding=(8, 5), background="#315C16", foreground="#FFFFFF")
        self.style.map("Workspace.TButton", background=[("active", "#D6EAC5")])
        self.style.map("Current.Workspace.TButton", background=[("active", "#264810")])
        self.style.configure("Close.Workspace.TButton", font=("Segoe UI", 10, "bold"),
                             padding=(5, 0), background="#E5EEDB", foreground="#263519")
        self.style.configure("Current.Close.Workspace.TButton", font=("Segoe UI", 10, "bold"),
                             padding=(5, 0), background="#315C16", foreground="#FFFFFF")
        self.style.map("Close.Workspace.TButton", background=[("active", "#D6EAC5")])
        self.style.map("Current.Close.Workspace.TButton", background=[("active", "#264810")])
        self.module_strip.bind("<Configure>", self._layout_workspace_tabs)
        self.bind("<Control-w>", self._close_module, add="+")

        def highlight_module(_event=None):
            current_group = nav_groups.get(self.tabs.select())
            self._layout_workspace_tabs()
            for button in set(nav_groups.values()):
                prefix = "Selected." if button is current_group else ""
                button.configure(style=f"{prefix}Navigation.{button.winfo_class()}")

        self.tabs.bind("<<NotebookTabChanged>>", highlight_module, add="+")
        highlight_module()
        self._build_dashboard()
        self._build_cari()
        self._build_borc_takip()
        self._build_alisveris_cari()
        self._build_hareket()
        self._build_cek()
        self._build_stok()
        self._build_personel()
        self._build_beton()
        self._build_mastarci()
        self._build_fatura()
        self._build_gelen_faturalar()
        self._build_kasa()
        self._build_extra_tabs()
        bottom = ttk.Frame(self, padding=(18, 4, 18, 12))
        bottom.pack(side="bottom",fill="x",before=self.tabs)
        ttk.Button(bottom, text="Yenile (F5)", command=self._refresh_from_button).pack(side="left", padx=(0, 6))
        other = nav_groups[str(self.stok_tab)]
        other_menu = other.nametowidget(other.cget("menu"))
        reports = tk.Menu(other_menu, tearoff=False)
        for label, action in (
            ("PDF Raporları", self.pdf_menu),
            ("Excel Raporu", self.export_excel),
            ("CSV Dışa Aktar", self.export_csv),
        ):
            reports.add_command(label=label, command=action)
        other_menu.add_separator()
        other_menu.add_cascade(label="Raporlar", menu=reports)
        settings = [
            ("Şimdi Yedekle", self.backup),
            ("Veritabanı Kontrolü", self.database_health_check),
            ("Veritabanı Konumu", self.show_db_path),
            ("Güncellemeleri Denetle", self.check_for_updates),
            ("Güncelleme URL Ayarları", self.configure_update_url),
            ("ZIP ile Güncelle", self.install_update_package),
        ]
        if self.membership_client is None:
            settings.append(("Şifre Değiştir", self.change_password_dialog))
        if getattr(self, "current_user", {}).get("role") == "Yönetici":
            settings.extend([("Kullanıcı Yönetimi", self.manage_users),
                             ("İşlem Geçmişi", self.show_audit_log)])
        system_menu = tk.Menu(other_menu, tearoff=False)
        for label, action in settings:
            system_menu.add_command(label=label, command=action)
        other_menu.add_cascade(label="Sistem", menu=system_menu)

    def _layout_workspace_tabs(self, _event=None):
        width = self.module_strip.winfo_width()
        if width <= 1:
            return
        selected = self.tabs.select()
        x = y = row_height = 0
        for page, (tab, select, close) in self.module_controls.items():
            if self.tabs.tab(page, "state") == "hidden":
                tab.place_forget()
                continue
            style = "Current.Workspace.TButton" if page == selected else "Workspace.TButton"
            select.configure(style=style)
            tab_width = select.winfo_reqwidth()
            tab_height = select.winfo_reqheight()
            if close is not None:
                close.configure(style="Current.Close.Workspace.TButton" if page == selected else "Close.Workspace.TButton")
                tab_width += close.winfo_reqwidth()
                tab_height = max(tab_height, close.winfo_reqheight())
            if x and x + tab_width > width:
                y += row_height + 4
                x = row_height = 0
            tab.place(x=x, y=y, width=tab_width, height=tab_height)
            x += tab_width + 5
            row_height = max(row_height, tab_height)
        height = y + row_height + 4
        if int(self.module_strip.cget("height")) != height:
            self.module_strip.configure(height=height)

    def _open_module(self, page):
        self.tabs.add(page)
        self.tabs.select(page)
        self._layout_workspace_tabs()

    def _close_module(self, _event=None, *, page=None):
        page = str(page) if page is not None else self.tabs.select()
        if page == str(self.cari_card_page) and self.tabs.select() == page:
            self._return_from_cari_card()
        elif page and page != str(self.dashboard):
            self.tabs.hide(page)
            self._layout_workspace_tabs()
        return "break"

    def _layout_module_navigation(self, nav, items):
        width = nav.winfo_width()
        if width <= 1:
            return
        row_height = max(item.winfo_reqheight() for item in items) + 4
        x = y = 0
        for item in items:
            item_width = item.winfo_reqwidth()
            if x and x + item_width > width:
                x = 0
                y += row_height
            item.place(x=width - item_width if item is items[-1] else x,
                       y=y, width=item_width, height=row_height - 4)
            x += item_width + 6
        height = y + row_height
        if int(nav.cget("height")) != height:
            nav.configure(height=height)

    def _action_menu(self, parent, text, actions):
        button = ttk.Menubutton(parent, text=text)
        menu = tk.Menu(button, tearoff=False)
        for label, command in actions:
            menu.add_command(label=label, command=command)
        button.configure(menu=menu)
        return button

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
        source = "HOYTAŞ Canlı" if ISOLATED_PROFILE_DIR else ("HTTPS manifesti" if current_url or DEFAULT_UPDATE_MANIFEST_URL else "GitHub Releases")
        ttk.Label(frame, text=f"Kanal: {source}   |   Kurulu sürüm: v{APP_VERSION}").pack(anchor="w", pady=(0, 8))
        if ISOLATED_PROFILE_DIR:
            ttk.Label(frame, text=LIVE_UPDATE_MANIFEST_URL, wraplength=540).pack(anchor="w", pady=(0, 8))
            ttk.Button(frame, text="Güncellemeleri Denetle", command=self.check_for_updates).pack(anchor="w")
            return
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
        """Only fetch version metadata in a worker; never invoke Tk or install there."""
        results = queue.Queue()
        def worker():
            try:
                repo = None
                if ISOLATED_PROFILE_DIR:
                    url = LIVE_UPDATE_MANIFEST_URL
                else:
                    repo = self._update_config()
                    with open(UPDATE_CONFIG_PATH, "r", encoding="utf-8") as f:
                        config = json.load(f)
                    url = str(config.get("manifest_url", "")).strip() or DEFAULT_UPDATE_MANIFEST_URL
                use_manifest = bool(url)
                if url:
                    url = normalize_dropbox_download_url(url)
                    if urllib.parse.urlparse(url).scheme != "https":
                        raise ValueError("Manifest HTTPS olmalıdır.")
                elif repo:
                    url = f"https://api.github.com/repos/{repo}/releases/latest"
                else:
                    results.put(None)
                    return
                req = urllib.request.Request(url, headers={"User-Agent": "HOYTAS-Updater/1.0"})
                with urllib.request.urlopen(req, timeout=10) as response:
                    raw = response.read(2 * 1024 * 1024 + 1)
                if len(raw) > 2 * 1024 * 1024:
                    raise ValueError("Güncelleme bilgisi çok büyük.")
                data = json.loads(raw.decode("utf-8"))
                if ISOLATED_PROFILE_DIR and data.get("channel") != "hoytas-canli":
                    raise RuntimeError("Canlı güncelleme kanalı doğrulanamadı.")
                version = int(data["version"]) if use_manifest else int(re.fullmatch(r"v?(\d+)", str(data["tag_name"])).group(1))
                results.put(version)
            except Exception as exc:
                self._log_update_error(exc)
                results.put(None)
        def poll():
            try:
                version = results.get_nowait()
            except queue.Empty:
                self.after(200, poll)
                return
            if version is not None and version > APP_VERSION:
                messagebox.showinfo("HOYTAŞ güncellemesi", f"Yeni sürüm v{version} mevcut. Güncellemeleri Denetle düğmesinden kurabilirsiniz.", parent=self)
        threading.Thread(target=worker, daemon=True).start()
        self.after(200, poll)

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

    def _update_status_text(self, source, version):
        return f"Kanal: {source}\nKurulu sürüm: v{APP_VERSION}\nSunucudaki sürüm: v{version}"

    def check_for_updates(self, silent=False):
        """Önce bağımsız HTTPS manifestini, gerekirse eski GitHub kaynağını denetle."""
        try:
            if ISOLATED_PROFILE_DIR:
                self._check_manifest_update(LIVE_UPDATE_MANIFEST_URL, silent=silent)
                return
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
            status = self._update_status_text(f"GitHub Releases ({repo})", version)
            if version <= APP_VERSION:
                if not silent:
                    messagebox.showinfo("HOYTAŞ güncelleme durumu", f"{status}\n\nDaha yeni bir sürüm bulunamadı.", parent=self)
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
            if not messagebox.askyesno("HOYTAŞ güncellemesi bulundu", f"{status}\n\nŞimdi indirip güvenli şekilde kurayım mı?", parent=self):
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
        if ISOLATED_PROFILE_DIR and manifest.get("channel") != "hoytas-canli":
            raise RuntimeError("Canlı güncelleme kanalı doğrulanamadı.")
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
        if ISOLATED_PROFILE_DIR and not re.fullmatch(
            r"https://github\.com/hoytashazirbeton02-debug/hoytas-finans/archive/[0-9a-f]{40}\.zip", package_url
        ):
            raise RuntimeError("Canlı güncelleme paketi HOYTAŞ yayın deposuna ait olmalıdır.")
        digest = str(manifest.get("sha256", "")).strip().lower()
        package_parsed = urllib.parse.urlparse(package_url)
        # Kaynağın güvenliğini sürüm karşılaştırmasından önce doğrula.
        if package_parsed.scheme != "https" or not package_parsed.netloc or package_parsed.username or package_parsed.password:
            raise RuntimeError("Paket adresi kullanıcı bilgisi içermeyen geçerli bir HTTPS adresi olmalıdır.")
        if not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise RuntimeError("Manifest içinde geçerli SHA-256 bilgisi bulunamadı.")
        channel = "HOYTAŞ ortak kullanım" if ISOLATED_PROFILE_DIR else f"Yerel yayın ({parsed.hostname})"
        status = self._update_status_text(channel, version)
        if version <= APP_VERSION:
            if not silent:
                messagebox.showinfo("HOYTAŞ güncelleme durumu", f"{status}\n\nDaha yeni bir sürüm bulunamadı.", parent=self)
            return
        # Sessiz kontrol yalnızca "güncel" ve ağ hatası bildirimlerini gizler;
        # yeni sürüm bulunduğunda kullanıcı onayı olmadan kurulum başlatılmaz.
        if not messagebox.askyesno("HOYTAŞ güncellemesi bulundu", f"{status}\n\nŞimdi indirip güvenli şekilde kurayım mı?", parent=self):
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
            if not ISOLATED_PROFILE_DIR:
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
                    if ISOLATED_PROFILE_DIR and clean_name.lower().endswith((".db", ".sqlite", ".sqlite3")):
                        raise RuntimeError("Canlı güncelleme paketi veritabanı dosyası içeremez.")
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

                if ISOLATED_PROFILE_DIR:
                    package_root = prefix + "/" if prefix else ""
                    members_by_path = {i.filename.replace("\\", "/") for i in members}
                    if not {package_root + name for name in LIVE_UPDATE_FILES}.issubset(members_by_path):
                        raise RuntimeError("Canlı HOYTAŞ paketi eksik.")
                    try:
                        marker = json.loads(zf.read(package_root + "CANLI_GUNCELLEME.json"))
                    except (KeyError, ValueError) as error:
                        raise RuntimeError("Bu paket canlı HOYTAŞ kanalına ait değil.") from error
                    if marker != {"channel": "hoytas-canli", "version": version}:
                        raise RuntimeError("Canlı HOYTAŞ paket kimliği doğrulanamadı.")

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
                    with open(os.path.join(source_dir, "hoytas.py"), encoding="utf-8") as app_file:
                        if not re.search(r"^APP_VERSION\s*=\s*" + str(version) + r"\s*$",
                                         app_file.read(4096), re.MULTILINE):
                            raise RuntimeError("Paketin uygulama dosyasındaki sürüm numarası uyuşmuyor.")
                    source_db = sqlite3.connect(DB_PATH, timeout=30)
                    try:
                        backup_path = _verified_sqlite_backup(source_db, f"guncelleme_oncesi_v{version}")
                    finally:
                        source_db.close()
                    os.rename(source_dir, destination)
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
                if not ISOLATED_PROFILE_DIR:
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
        if os.name == "nt":
            ttk.Button(self.dashboard, text="Yazdır", command=self._print_dashboard).pack(anchor="w")
        self.dashboard_dates = ("0001-01-01", "9999-12-31")
        self.dashboard_period = tk.StringVar(value="Tüm zamanlar")
        filters = ttk.Frame(self.dashboard)
        filters.pack(fill="x", pady=(8, 4))
        self._action_menu(filters, "Gelir / Gider Dönemi", [
            (label, lambda period=label: self._set_dashboard_period(period))
            for label in ("Bugün", "Bu Ay", "Tarih Aralığı", "Tüm zamanlar")
        ]).pack(side="left")
        ttk.Label(filters, textvariable=self.dashboard_period).pack(side="left", padx=8)
        self.metrics = ttk.Frame(self.dashboard)
        self.metrics.pack(fill="x", pady=4)
        self.metric_labels = {}
        for column in range(3):
            self.metrics.columnconfigure(column, weight=1)
        for index, (key, label) in enumerate([
            ("cari", "Cari kayıt"), ("gelir", "Toplam gelir"), ("gider", "Toplam gider"),
            ("cek", "Bekleyen çek"), ("stok", "Stok kalemi"), ("acik_borc", "Açık cari borcu"),
        ]):
            card = ttk.LabelFrame(self.metrics, text=label, padding=10, style="Metric.TLabelframe")
            card.grid(row=index // 3, column=index % 3, sticky="nsew", padx=(0, 10), pady=3)
            val = ttk.Label(card, text="—", style="Metric.TLabel")
            val.pack(anchor="w")
            self.metric_labels[key] = val
        self.debt_alert_label = ttk.Label(self.dashboard, text="", foreground="#A32626", font=("Segoe UI", 10, "bold"), wraplength=980, justify="left")
        self.debt_alert_label.pack(anchor="w", pady=(12, 4))
        self.dashboard.bind("<Configure>", lambda event: self.debt_alert_label.configure(
            wraplength=max(event.width - 36, 1)))

    def _tree(self, parent, columns, headings, widths=None, *, show_id=True, search_label="Listede Ara:", print_button=True, print_title=None):
        searchbar = ttk.Frame(parent)
        searchbar.pack(fill="x", pady=(6, 0))
        ttk.Label(searchbar, text=search_label).pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(searchbar, textvariable=search_var, width=32)
        search_entry.pack(side="left", padx=(6, 4))
        ttk.Button(searchbar, text="Temizle", command=lambda v=search_var: v.set("")).pack(side="left")
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, pady=(4, 0))
        tree = DateTreeview(frame, columns=columns, show="headings", height=6)
        if not show_id:
            tree.configure(displaycolumns=tuple(col for col in columns if col != "id"))
        numeric = {"borc", "alacak", "bakiye", "tutar", "toplam", "kalan", "odenen",
                   "pompa", "nakliye", "fiyat", "m3", "miktar", "maas", "ara", "kdv", "tevkifat", "net"}
        for i, (col, head) in enumerate(zip(columns, headings)):
            anchor = "e" if col in numeric else "w"
            tree.heading(col, text=head, anchor=anchor)
            tree.column(col, width=(widths[i] if widths else 130), anchor=anchor)
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
        tree._search_entry = search_entry
        tree._search_var = search_var
        search_var.trace_add("write", lambda *_args, w=tree: self._apply_tree_search(w))
        tree.bind("<Double-1>", lambda event, w=tree: self._double_click_edit(w))
        # Sağ tık bağlam menüsü; sol tık ve çift tık davranışları korunur.
        tree.bind("<Button-3>", lambda event, w=tree: self._tree_context_menu(event, w))
        tree.bind("<Control-Button-1>", lambda event, w=tree: self._tree_context_menu(event, w))
        if print_button and os.name == "nt":
            if print_title:
                title = print_title
            elif str(parent) in self.tabs.tabs():
                title = self.tabs.tab(parent, "text").replace("\n", " ")
            else:
                title = parent.winfo_toplevel().title()
            ttk.Button(searchbar, text="Yazdır", command=lambda: self._print_tree(tree, title)).pack(
                side="left", padx=(8, 0))
        return tree

    def _print_dashboard(self):
        rows = [("Gösterge", "Değer"), ("Gelir / Gider Dönemi", self.dashboard_period.get())]
        rows.extend((label.master.cget("text"), label.cget("text"))
                    for label in self.metric_labels.values())
        self._print_report_rows("Genel Bakış", rows)

    def _print_tree(self, tree, title):
        items = tree.get_children()
        if not items:
            return messagebox.showinfo("Yazdır", "Bu listede yazdırılacak kayıt yok.", parent=tree)
        columns = tree["columns"]
        visible = tree["displaycolumns"]
        visible = columns if visible == ("#all",) else visible
        indices = [columns.index(column) for column in visible]
        rows = [tuple(tree.heading(column, "text") for column in visible)]
        values = [tree.item(item, "values") for item in items]
        rows.extend(tuple(row[index] for index in indices) for row in values)
        if "m3" in visible and "nakliye" in visible:
            totals = {column: sum(float(str(row[columns.index(column)]).replace(",", "")) for row in values
                                  if row[columns.index(column)] not in ("-", "Yok", ""))
                      for column in ("m3", "nakliye", "toplam")}
            footer = {"tarih": "GENEL TOPLAM", "m3": f"{format_m3(totals['m3'])} m³",
                      "nakliye": f"{totals['nakliye']:,.2f}", "toplam": f"{totals['toplam']:,.2f}"}
            rows.append(tuple(footer.get(column, "") for column in visible))
        self._print_report_rows(title, rows)

    def _print_report_rows(self, title, rows):
        try:
            self._preview_then_save_pdf(title + " Listesi", rows=rows, print_dialog=True)
        except (OSError, ValueError, RuntimeError, ImportError) as error:
            messagebox.showerror("Yazdırma açılamadı", str(error), parent=self)

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
        undo_table = getattr(tree, "_undo_table", None) or table
        menu=tk.Menu(self, tearoff=0)
        if tree is getattr(self,"cari_tree",None):
            menu.add_command(label="Cari hareketlerini görüntüle", command=self.show_cari_hareketleri)
            menu.add_separator()
        if table:
            edit_command = self.edit_gelen_fatura if table == "gelen_fatura" else (lambda:self.edit_selected(tree,table))
            menu.add_command(label="Düzenle", command=edit_command)
            menu.add_command(label="Sil", command=lambda:self.delete_selected(tree,table))
        if undo_table:
            menu.add_command(label="Seçili İşlemi Geri Al", command=lambda:self.undo_selected_operation(tree, undo_table))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def undo_selected_operation(self, tree, table):
        from operation_undo import UndoUnavailable, reverse_operation, selected_operation

        picked = tree.selection()
        if not picked:
            return
        values = tree.item(picked[0], "values")
        record_id = int(values[0])
        if table == "cari_invoice":
            table = "fatura" if values[1] == "Satış" else "gelen_fatura"
        parent = tree.winfo_toplevel()
        account = self.current_user
        username = account["username"]
        admin = account["role"] == "Yönetici"
        try:
            with connect() as db:
                operation_id = selected_operation(db, table, record_id, username, admin=admin)
            if not self.confirm_undo(parent):
                return
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                if selected_operation(db, table, record_id, username, admin=admin) != operation_id:
                    raise UndoUnavailable("Seçilen kayıt onay sırasında değişmiş.")
                reverse_operation(db, operation_id)
        except (UndoUnavailable, sqlite3.Error, ValueError) as error:
            messagebox.showwarning("İşlem geri alınamadı", str(error), parent=parent)
            return
        self._refresh_selected_module()
        refresh = getattr(tree, "_after_undo", None)
        if refresh:
            refresh()
        messagebox.showinfo("İşlem geri alındı", "Seçili işlem ve bağlı kayıtları geri alındı.", parent=parent)

    def confirm_undo(self, parent):
        window = tk.Toplevel(parent)
        window.title("İşlemi geri al")
        window.transient(parent)
        window.resizable(False, False)
        previous_grab = parent.grab_current()
        window.grab_set()
        frame = ttk.Frame(window, padding=16)
        frame.pack()
        ttk.Label(frame, text="Seçili işlem ve aynı işlemde kaydedilen bağlı hareketler geri alınacak.",
                  wraplength=390).pack(pady=(0, 12))
        confirmed = False

        def approve():
            nonlocal confirmed
            confirmed = True
            window.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e")
        ttk.Button(buttons, text="Geri Al", command=approve).pack(side="left", padx=4)
        cancel = ttk.Button(buttons, text="Vazgeç", command=window.destroy)
        cancel.pack(side="left")
        cancel.focus_set()
        window.bind("<Escape>", lambda _: window.destroy())
        parent.wait_window(window)
        if previous_grab is not None and previous_grab.winfo_exists():
            previous_grab.grab_set()
        return confirmed

    def _prompt_for_cari(self, cari_context, title, fields, initial=None):
        if cari_context is not None:
            cari_id, _ = cari_context
            with connect() as db:
                current = db.execute("SELECT unvan, tur FROM cari WHERE id=?", (cari_id,)).fetchone()
            if current is None:
                messagebox.showwarning("Cari bulunamadı", "Bu cari artık kayıtlı değil.", parent=self)
                return None
            if title == "Yeni Beton Satışı" and current[1] not in ("Müşteri", "Diğer"):
                messagebox.showwarning("Cari türü", "Beton satışı için müşteri veya diğer türünde cari seçin.", parent=self)
                return None
            fields = [(key, "Cari", "combo", [f"{cari_id} - {current[0]}"]) if key == "cari"
                      else (key, label, kind, options) for key, label, kind, options in fields]
        return self.prompt(title, fields, initial) if initial else self.prompt(title, fields)

    def _return_from_cari_card(self):
        self._open_module(self.cari_tab)
        self.tabs.hide(self.cari_card_page)
        self._layout_workspace_tabs()
        self.cari_tree.focus_set()

    def show_cari_hareketleri(self, cari_context=None, selected_tab=0):
        """Open the customer card in the main workspace."""
        if cari_context is None:
            sel = self.cari_tree.selection()
            if not sel:
                return messagebox.showinfo("Cari seçin", "Önce bir cari seçin.", parent=self)
            vals = self.cari_tree.item(sel[0], "values")
            cari_context = int(vals[0]), str(vals[2])
        self._render_cari_card(cari_context, selected_tab)
        self._open_module(self.cari_card_page)

    def _render_cari_card(self, cari_context, selected_tab):
        cari_id, unvan = cari_context
        card = self.cari_card_page
        sales_search = ""
        selected_site = "Tüm şantiyeler"
        if getattr(card, "_customer_id", None) == cari_id and hasattr(card, "_sales_filter"):
            sales_search, selected_site = (variable.get() for variable in card._sales_filter)
        for child in card.winfo_children():
            child.destroy()
        card._customer_id = cari_id
        top = ttk.Frame(card, padding=(12, 8, 12, 8))
        top.pack(fill="x")
        header_actions = ttk.Frame(top)
        header_actions.pack(fill="x", pady=(0, 4))
        ttk.Button(header_actions, text="‹ Carilere Dön", command=self._return_from_cari_card).pack(side="left")
        name = ttk.Label(top, text=unvan, font=("Segoe UI", 21, "bold"), foreground="#174D30")
        name.pack(anchor="w")
        top.bind("<Configure>", lambda event: name.configure(wraplength=max(event.width - 24, 1)))
        with connect() as c:
            totals=c.execute("SELECT toplam_borc,toplam_alacak FROM cari_bakiye WHERE cari_id=?",(cari_id,)).fetchone() or (0,0)
            moves=c.execute("SELECT h.id,h.tarih,h.tur,h.kategori,COALESCE(h.aciklama,''),COALESCE(h.odeme_yontemi,''),h.tutar FROM hareket h WHERE h.cari_id=? ORDER BY h.tarih ASC,h.id ASC",(cari_id,)).fetchall()
            sales=c.execute("SELECT b.id,b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(b.aciklama,''),COALESCE(s.ad,''),b.mikser_nakliye_ucreti,f.belge_no FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id LEFT JOIN fatura f ON f.id=b.fatura_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC",(cari_id,)).fetchall()
            checks=c.execute("SELECT id,yon,COALESCE(banka,''),COALESCE(cek_no,''),vade,tutar,durum FROM cek WHERE cari_id=? ORDER BY vade ASC,id ASC",(cari_id,)).fetchall()
        balance=totals[0]-totals[1]
        total_m3 = sum(safe_float(r[3] or 0) for r in sales)
        summary = ttk.Frame(top, style="CariSummary.TFrame")
        summary.pack(fill="x", pady=(6, 0))
        for column, (label, value) in enumerate((
            ("Toplam Borç", f"{totals[0]:,.2f} TL"),
            ("Toplam Alacak", f"{totals[1]:,.2f} TL"),
            ("Cari Bakiye", f"{balance:,.2f} TL"),
            ("Toplam Beton", f"{format_m3(total_m3)} m³"),
        )):
            summary.columnconfigure(column, weight=1, uniform="cari")
            metric = ttk.Frame(summary, padding=(10, 6), style="CariSummary.TFrame")
            metric.grid(row=0, column=column, sticky="ew", padx=(0, 4))
            ttk.Label(metric, text=label, style="CariSummary.TLabel").pack(anchor="w")
            ttk.Label(metric, text=value, style="CariSummary.TLabel",
                      font=("Segoe UI", 12, "bold")).pack(anchor="w")
        nb = ttk.Notebook(card, style="CariCard.TNotebook")
        nb.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        mf = ttk.Frame(nb, padding=6)
        nb.add(mf, text="Cari Hareketleri")
        movement_actions = ttk.Frame(mf)
        movement_actions.pack(fill="x", pady=(0, 4))
        mt = self._tree(mf, ("id", "tarih", "tur", "kategori", "aciklama", "odeme", "borc", "alacak", "bakiye"),
                        ("ID", "Tarih", "Tür", "Kategori", "Açıklama", "Ödeme", "Borç TL", "Alacak TL", "Bakiye TL"),
                        (45, 94, 70, 98, 270, 85, 95, 95, 105), show_id=False,
                        print_button=False)
        ttk.Button(movement_actions, text="Hareketi Düzenle",
                   command=lambda: refresh_card(lambda _cari: self.edit_selected(mt, "hareket"))).pack(side="left")
        ttk.Button(movement_actions, text="Tahsilat Al",
                   command=lambda: refresh_card(self.cari_tahsilat, 0)).pack(side="left", padx=6)
        self._action_menu(movement_actions, "Çıktı Al", [
            ("Cari Hareketleri PDF", lambda: self.export_single_cari_pdf(cari_id, "Cari Hareket Dökümü")),
            ("Cari Hesap Ekstresi PDF", lambda: self.export_single_cari_pdf(cari_id, "Cari Hesap Ekstresi")),
            ("Cari Hesap Ekstresi Excel", lambda: self.export_single_cari_excel(cari_id, unvan, preview=True)),
            ("Yazdır", lambda: self._print_tree(mt, f"{unvan} – Cari Hareketleri")),
        ]).pack(side="right")
        bal=0
        sales_by_id = {int(r[0]): r for r in sales}
        import re
        movement_rows = []
        for r in moves:
            debt=r[6] if r[2]=="Borç" else 0
            credit=r[6] if r[2]=="Alacak" else 0
            bal += debt-credit
            description = r[4]
            match = re.search(r"(?:\[SATIŞ:(\d+)\]|Beton satış #?(\d+))", description, re.IGNORECASE)
            if match and r[2] == "Borç" and int(match.group(1) or match.group(2)) in sales_by_id:
                sale = sales_by_id[int(match.group(1) or match.group(2))]
                description = format_beton_aciklama(sale[2], sale[3], sale[4], sale[5], sale[6], sale[9], sale[8], sale[7], sale[10])
            movement_rows.append((r[0],r[1],r[2],r[3],description,r[5],f"{debt:,.2f}" if debt else "-",f"{credit:,.2f}" if credit else "-",f"{bal:,.2f}"))
        self.fill_tree(mt, movement_rows)
        sf = ttk.Frame(nb, padding=6)
        nb.add(sf, text="Beton Satışları")
        sales_total = ttk.Label(sf, font=("Segoe UI", 10, "bold"), style="CariSummary.TLabel", padding=(10, 6))
        sales_total.pack(side="bottom", fill="x", pady=(5, 0))
        sales_status = ttk.Label(sf, font=("Segoe UI", 9, "bold"))
        st = self._tree(sf, ("id", "tarih", "santiye", "cinsi", "m3", "fiyat", "pompa", "nakliye", "toplam", "durum", "aciklama"),
                        ("ID", "Tarih", "Şantiye", "Beton", "Miktar (m³)", "Birim Fiyat", "Pompa", "Mikser Nakliye", "Toplam", "Ödeme Durumu", "Açıklama"),
                        (42, 88, 135, 90, 85, 97, 85, 110, 100, 135, 250), show_id=False,
                        search_label="Satışlarda Ara", print_button=False)
        toolbar = st._search_entry.master
        sites = ("Tüm şantiyeler", *sorted({r[9] for r in sales if r[9]}))
        site = tk.StringVar(card, value=selected_site if selected_site in sites else sites[0])
        card._sales_filter = (st._search_var, site)
        st._search_var.set(sales_search)
        site_picker = ttk.Combobox(toolbar, textvariable=site, state="readonly", width=15,
                                   values=sites)
        site_picker.pack(side="left", padx=5)
        ttk.Button(toolbar, text="Yeni Beton Satışı",
                   command=lambda: refresh_card(self.add_beton, 1)).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Tahsilat Al",
                   command=lambda: refresh_card(self.cari_tahsilat, 0)).pack(side="left", padx=4)

        def visible_sales():
            return [st.item(item, "values") for item in st.get_children()]

        def sales_pdf():
            ids = tuple(int(row[0]) for row in visible_sales())
            self.export_single_cari_pdf(cari_id, "Cari Beton Satış Dökümü", sale_ids=ids)

        self._action_menu(toolbar, "Çıktı Al", [
            ("Beton Satışları PDF", sales_pdf),
            ("Beton Satışları Excel", lambda: self.export_concrete_rows_excel(st, unvan)),
            ("Yazdır", lambda: self._print_tree(st, f"{unvan} – Beton Satışları")),
        ]).pack(side="right", padx=(4, 0))
        st._search_entry.configure(width=13)
        st._row_filter = lambda row: site.get() == "Tüm şantiyeler" or row[2] == site.get()
        st._on_search = lambda rows: sales_total.configure(text=(
            f"{len(rows)} satış   |   Toplam {format_m3(sum(float(str(row[4]).replace(',', '')) for row in rows))} m³"
            f"   |   {sum(safe_float(row[8]) for row in rows):,.2f} TL"))
        site_picker.bind("<<ComboboxSelected>>", lambda _event: self._apply_tree_search(st))
        sales_rows = []
        for r in sales:
            detail = format_beton_aciklama(r[2],r[3],r[4],r[5],r[6],r[9],r[8],r[7],r[10])
            if r[11]:
                detail = f"Fatura {r[11]} · {detail}"
            sales_rows.append((r[0],r[1],r[9] or "-",r[2],format_m3(r[3]),f"{r[4]:,.2f}",f"{r[5]:,.2f}" if r[5] else "Yok",f"{r[10]:,.2f}" if r[10] else "-",f"{r[6]:,.2f}",r[7],detail))
        self.fill_tree(st, sales_rows)
        cf = ttk.Frame(nb, padding=6)
        nb.add(cf, text="Bağlı Çekler")
        ct = self._tree(cf, ("id", "yon", "banka", "no", "vade", "tutar", "durum"),
                        ("ID", "Yön", "Banka", "Çek No", "Vade", "Tutar TL", "Durum"),
                        (45, 100, 160, 140, 110, 140, 150), show_id=False)
        self.fill_tree(ct, [(cid, direction, bank, number, due, f"{amount:,.2f}", status)
                            for cid, direction, bank, number, due, amount, status in checks])
        with connect() as db:
            invoices = db.execute("""SELECT id, 'Satış', tarih, belge_no, genel_toplam,
                odenen, tahsil_edilecek-odenen, durum,
                CASE WHEN json_valid(belge_bilgileri) AND json_extract(belge_bilgileri,'$.cari_isleme') IS NOT NULL
                     THEN 'Cariye işlendi · ' || COALESCE(aciklama,'') ELSE aciklama END
                FROM fatura WHERE cari_id=?
                UNION ALL
                SELECT id, 'Alış', tarih, belge_no, genel_toplam, odenen,
                genel_toplam-odenen, durum, aciklama FROM gelen_fatura WHERE cari_id=?
                ORDER BY 3, 2, 4""", (cari_id, cari_id)).fetchall()
        invoice_frame = ttk.Frame(nb, padding=6)
        nb.add(invoice_frame, text="Faturalar")
        invoice_actions = ttk.Frame(invoice_frame)
        invoice_actions.pack(fill="x", pady=(0, 6))
        ttk.Button(invoice_actions, text="Beton Satışını Faturaya Bağla",
                   command=lambda: refresh_card(self.match_concrete_invoice, 3)).pack(side="left", padx=(0, 8))
        ttk.Button(invoice_actions, text="Fatura Türleri ve Mevzuat",
                   command=self.show_invoice_tax_options).pack(side="left", padx=(0, 8))
        ttk.Button(invoice_actions, text="Cari Kontrolü",
                   command=lambda: self.selected_invoice_document(invoice_tree, review=True)).pack(side="left", padx=(0, 8))
        ttk.Button(invoice_actions, text="Satış Taslağı PDF",
                   command=lambda: self.selected_invoice_document(invoice_tree)).pack(side="left", padx=(0, 8))
        ttk.Button(invoice_actions, text="Belge Bilgileri",
                   command=lambda: self.selected_invoice_document(invoice_tree, edit=True)).pack(side="left")
        invoice_tree = self._tree(invoice_frame, ("id", "tur", "tarih", "belge", "toplam", "odenen", "kalan", "durum", "aciklama"),
                                  ("ID", "Tür", "Tarih", "Belge No", "Toplam TL", "Ödenen TL", "Kalan TL", "Durum", "Açıklama"),
                                  (55, 75, 95, 115, 115, 115, 115, 110, 225), show_id=False)
        self.fill_tree(invoice_tree, [
            (invoice_id, kind, day, number or "", f"{safe_float(total):,.2f}",
             f"{safe_float(paid):,.2f}", f"{safe_float(remaining):,.2f}", status, detail or "")
            for invoice_id, kind, day, number, total, paid, remaining, status, detail in invoices
        ])
        nb.select(selected_tab)

        def refresh_card(operation=None, tab=None):
            current_tab = nb.index(nb.select()) if tab is None else tab
            if operation is not None:
                operation(cari_context)
            self._render_cari_card(cari_context, current_tab)

        self._refresh_open_cari_card = refresh_card

        for tree, table in ((mt, "hareket"), (st, "beton_satis"), (ct, "cek"), (invoice_tree, "cari_invoice")):
            tree._undo_table = table
            tree._after_undo = refresh_card
            tree.bind("<Button-3>", lambda event, widget=tree: self._tree_context_menu(event, widget))

        def edit_sale(event):
            if self._edit_tree_row(event, st, "beton_satis"):
                refresh_card(tab=1)
                return "break"

        st.bind("<Double-1>", edit_sale)

        self._action_menu(header_actions, "Cari İşlemleri", [
            ("Çek Girişi", lambda: refresh_card(self.add_cek, 2)),
            ("Satış Faturası Taslağı", lambda: refresh_card(self.add_fatura, 3)),
            ("Beton Satışını Faturaya Bağla", lambda: refresh_card(self.match_concrete_invoice, 3)),
            ("Cariye Yeni Fatura Borcu İşle", lambda: refresh_card(lambda customer: self.add_fatura(customer, post=True), 3)),
            ("Alış Faturası", lambda: refresh_card(self.add_gelen_fatura, 3)),
            ("Cari Hareket Ekle", lambda: refresh_card(self.add_hareket, 0)),
            ("Beton-Cari Bağlantısını Kontrol Et", lambda: self.repair_cari_beton_links(cari_id, refreshed=refresh_card)),
        ]).pack(side="right", padx=(4, 0))
        ttk.Button(header_actions, text="Yenile", command=refresh_card).pack(side="right")
        with connect() as audit_db:
            missing_sales, wrong_sales, ambiguous_sales = audit_beton_cari_amounts(audit_db, cari_id)
        if missing_sales or wrong_sales or ambiguous_sales:
            sales_status.configure(text=f"Cari bağlantısı: {len(missing_sales)} eksik, "
                                   f"{len(wrong_sales)} tutar hatası, {len(ambiguous_sales)} belirsiz.")
            sales_status.pack(side="bottom", anchor="w", before=st.master, pady=(6, 0))

    def repair_cari_beton_links(self, cari_id, *, refreshed):
        """Backed-up, explicit correction of identifiable movements only."""
        with connect() as db:
            missing, mismatched, ambiguous = audit_beton_cari_amounts(db, cari_id)
            if not missing and not mismatched and not ambiguous:
                return messagebox.showinfo("Cari bağlantısı", "Satış bağlantıları ve tutarları uyumlu görünüyor.", parent=self)
            summary = (f"Eksik bağlantı: {len(missing)}\n"
                       f"Yanlış tutarlı/bağlantılı hareket: {len(mismatched)}\n"
                       f"Belirsiz kayıt (otomatik değişmez): {len(ambiguous)}\n\n"
                       f"Eksik borç toplamı: {sum(x[2] for x in missing):,.2f} TL\n"
                       f"Bağlı tutar farkı: {sum(x[2]-x[4] for x in mismatched):,.2f} TL")
            if not missing and not mismatched:
                return messagebox.showwarning("Cari bağlantısı inceleme", summary + "\n\nBelirsiz kayıtlar elle incelenmelidir.", parent=self)
            if not messagebox.askyesno("Cari bağlantısı onarımı", summary +
                    "\n\nÖNEMLİ: Bu işlem cari bakiyesini değiştirebilir."
                    "\nÖnce veritabanı yedeği alınarak yalnızca kesin eşleşen"
                    " kayıtlar düzeltilecek. Onaylıyor musunuz?", parent=self):
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
        refreshed()
        messagebox.showinfo("Cari bağlantısı", f"{len(again_missing)} eksik, {len(again_mismatch)} hatalı hareket onarıldı.\n"
                            f"Belirsiz kayıt: {len(again_ambiguous)}\nYedek: {backup}", parent=self)

    def _safe_export_name(self, value):
        cleaned = "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in str(value)).strip()
        return (cleaned[:60] or "Cari")

    def _report_filename(self, report_type, extension, customer=None):
        parts = ["HOYTAS", self._safe_export_name(report_type)]
        if customer:
            parts.append(self._safe_export_name(customer))
        parts.append(datetime.now().strftime("%Y%m%d"))
        return "_".join(parts).replace(" ", "_") + extension

    def _pdf_preview_action(self, report_type, viewer, preview_path):
        """Keep the preview available while the user prints or chooses to save."""
        window = tk.Toplevel(self)
        window.title("PDF Önizleme")
        window.transient(self)
        window.resizable(False, False)
        frame = ttk.Frame(window, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"{report_type} önizlemesi {viewer} ile açıldı.",
                  wraplength=450).pack(anchor="w", pady=(0, 12))
        status = tk.StringVar(window)
        ttk.Label(frame, textvariable=status, wraplength=450).pack(anchor="w")
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        save = False

        def choose_save():
            nonlocal save
            save = True
            window.destroy()

        def print_pdf():
            try:
                open_pdf_in_viewer(preview_path, print_dialog=True)
            except (OSError, RuntimeError) as error:
                messagebox.showerror("Yazdırma açılamadı", str(error), parent=window)
            else:
                status.set("Acrobat'ta yazıcıyı ve kopya sayısını seçin. İşiniz bitince bu pencereyi kapatın.")

        if os.name == "nt":
            ttk.Button(buttons, text="Yazdır", command=print_pdf).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="PDF Kaydet", command=choose_save).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Kapat", command=window.destroy).pack(side="right")
        window.bind("<Escape>", lambda event: window.destroy())
        previous_grab = self.grab_current()
        window.grab_set()
        try:
            self.wait_window(window)
        finally:
            if previous_grab is not None and previous_grab.winfo_exists():
                previous_grab.grab_set()
        return save

    def _preview_then_save_pdf(self, report_type, final_path=None, cari_id=None, *, rows=None, print_dialog=False, sale_ids=None):
        """Önce tek önizleme göster; kullanıcı onaylarsa bir kez hedef dosyayı seçtirip kaydet."""
        preview_path = None
        try:
            fd, preview_path = tempfile.mkstemp(prefix="HOYTAS_PDF_ONIZLEME_", suffix=".pdf")
            os.close(fd)
            if sale_ids is not None:
                self.generate_pdf(report_type, preview_path, cari_id, sale_ids=sale_ids)
            elif rows is None:
                self.generate_pdf(report_type, preview_path, cari_id)
            else:
                self.generate_pdf(report_type, preview_path, cari_id, display_rows=rows)
            with open(preview_path, "rb") as stream:
                data = stream.read()
            if not data.startswith(b"%PDF-") or b"%%EOF" not in data[-2048:] or len(data) < 500:
                raise RuntimeError("Önizleme PDF'si doğrulanamadı; hiçbir dosya kaydedilmedi.")
            try:
                viewer = open_pdf_in_viewer(preview_path, print_dialog=print_dialog)
            except RuntimeError as exc:
                messagebox.showerror("PDF önizlemesi açılamadı", f"PDF kaydedilmedi. Önce Adobe Acrobat/Reader ile önizleme açılmalıdır.\n\n{exc}", parent=self)
                return False
            customer = None
            if cari_id is not None and report_type in (
                "Cari Hesap Ekstresi", "Cari Hareket Dökümü", "Cari Beton Satış Dökümü"
            ):
                with connect() as db:
                    record = db.execute("SELECT unvan FROM cari WHERE id=?", (cari_id,)).fetchone()
                if record:
                    customer = record[0]
            if cari_id is not None and report_type == "Fatura Taslağı PDF":
                with connect() as db:
                    record = db.execute("""SELECT c.unvan FROM fatura f JOIN cari c ON c.id=f.cari_id
                        WHERE f.id=?""", (cari_id,)).fetchone()
                if record:
                    customer = record[0]
            title = f"{report_type} – {customer}" if customer else report_type
            if not self._pdf_preview_action(title, viewer, preview_path):
                return False
            if not final_path:
                final_path = filedialog.asksaveasfilename(
                    parent=self, title=f"{title} – PDF Kaydet", defaultextension=".pdf",
                    filetypes=[("PDF dosyası", "*.pdf")],
                    initialfile=self._report_filename(report_type, ".pdf", customer))
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

    def export_selected_cari_excel(self):
        selection = self.cari_tree.selection()
        if not selection:
            return messagebox.showinfo("Cari seçin", "Excel dökümü almak için önce cari hesap seçin.", parent=self)
        values = self.cari_tree.item(selection[0], "values")
        self.export_single_cari_excel(int(values[0]), str(values[2]))

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

    def export_single_cari_pdf(self, cari_id, report_type="Cari Hesap Ekstresi", *, sale_ids=None):
        """Cari hareket penceresindeki cari için doğrudan doğrulanmış PDF ekstresi üret."""
        try:
            self._preview_then_save_pdf(report_type, None, cari_id, sale_ids=sale_ids)
        except ImportError:
            messagebox.showerror("Eksik bileşen", "PDF oluşturmak için kurulum.bat dosyasını çalıştırın veya reportlab paketini kurun.", parent=self)
        except Exception as exc:
            messagebox.showerror("PDF oluşturulamadı", str(exc), parent=self)

    def _preview_excel(self, workbook, parent):
        """Show the export's worksheets before choosing a destination."""
        window = tk.Toplevel(parent)
        window.title("Excel Önizleme")
        window.geometry("880x540")
        window.transient(parent)
        notebook = ttk.Notebook(window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        for sheet in workbook.worksheets:
            page = ttk.Frame(notebook, padding=4)
            notebook.add(page, text=sheet.title)
            titles = [str(cell.value or "") for cell in sheet[1]]
            columns = tuple(f"col{i}" for i in range(len(titles)))
            table = self._tree(page, columns, titles, print_button=False)
            self.fill_tree(table, [tuple("" if value is None else value for value in row)
                                   for row in sheet.iter_rows(min_row=2, values_only=True)])
        approved = False

        def approve():
            nonlocal approved
            approved = True
            window.destroy()

        buttons = ttk.Frame(window, padding=(10, 0, 10, 10))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Excel Kaydet", command=approve).pack(side="right")
        ttk.Button(buttons, text="Vazgeç", command=window.destroy).pack(side="right", padx=6)
        window.bind("<Escape>", lambda _event: window.destroy())
        previous = self.grab_current()
        window.grab_set()
        try:
            self.wait_window(window)
        finally:
            if previous is not None and previous.winfo_exists():
                previous.grab_set()
        return approved

    def export_concrete_rows_excel(self, tree, customer):
        """Export the concrete rows currently displayed in the customer card."""
        parent = tree.winfo_toplevel()
        try:
            from openpyxl import Workbook

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Beton Satışları"
            columns = tree["displaycolumns"]
            indices = [tree["columns"].index(column) for column in columns]
            numeric = {"m3", "fiyat", "pompa", "nakliye", "toplam"}
            sheet.append([tree.heading(column, "text") for column in columns])
            for item in tree.get_children():
                row = tree.item(item, "values")
                values = [row[index] for index in indices]
                sheet.append([(float(str(value).replace(",", "")) if value not in ("-", "Yok", "") else 0)
                              if column in numeric else value for column, value in zip(columns, values)])
            totals = {"tarih": "GENEL TOPLAM"}
            for column in ("m3", "pompa", "nakliye", "toplam"):
                index = columns.index(column) + 1
                totals[column] = sum(sheet.cell(row, index).value for row in range(2, sheet.max_row + 1))
            sheet.append([totals.get(column, "") for column in columns])
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if cell.data_type == "f":
                        cell.data_type = "s"
            sheet.freeze_panes = "A2"
            for cells in sheet.columns:
                sheet.column_dimensions[cells[0].column_letter].width = max(14, min(45, max(
                    len(str(cell.value or "")) + 2 for cell in cells)))
            if not self._preview_excel(workbook, parent):
                return
            path = filedialog.asksaveasfilename(parent=parent, title="Beton Satışları Excel Kaydet",
                defaultextension=".xlsx", filetypes=[("Excel çalışma kitabı", "*.xlsx")],
                initialfile=self._report_filename("Beton Satışları", ".xlsx", customer))
            if path:
                workbook.save(path)
        except (ImportError, OSError, ValueError) as exc:
            messagebox.showerror("Excel oluşturulamadı", str(exc), parent=parent)

    def export_single_cari_excel(self, cari_id, unvan, *, preview=False):
        """Cari hareketleri, beton satışları ve bağlı çekleri tek XLSX dosyasına aktar."""
        initial = self._report_filename("Cari Hesap Ekstresi", ".xlsx", unvan)
        try:
            from openpyxl import Workbook, load_workbook
            from openpyxl.drawing.image import Image as XLImage
            from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
            from openpyxl.drawing.xdr import XDRPositiveSize2D
            from openpyxl.utils.units import pixels_to_EMU
            from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
            from openpyxl.utils import get_column_letter
            with connect() as c:
                cari = c.execute("SELECT unvan,COALESCE(telefon,''),COALESCE(vergi_no,'') FROM cari WHERE id=?", (cari_id,)).fetchone()
                if not cari:
                    raise ValueError("Cari hesap bulunamadı.")
                moves = c.execute("SELECT tarih,tur,kategori,COALESCE(aciklama,''),COALESCE(odeme_yontemi,''),tutar FROM hareket WHERE cari_id=? AND tur IN ('Borç','Alacak') ORDER BY tarih ASC,id ASC", (cari_id,)).fetchall()
                sales = c.execute("SELECT b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(b.aciklama,''),COALESCE(s.ad,''),b.mikser_nakliye_ucreti FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC", (cari_id,)).fetchall()
                checks = c.execute("SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),COALESCE(cek_no,''),vade,tutar,durum,COALESCE(notlar,'') FROM cek WHERE cari_id=? ORDER BY vade ASC,id ASC", (cari_id,)).fetchall()
            wb = Workbook()
            summary = wb.active; summary.title = "Cari Özeti"
            logo_path = os.path.join(APP_DIR, "hoytas_logo.png")
            has_logo = os.path.isfile(logo_path)
            if has_logo:
                logo = XLImage(logo_path)
                scale = min(54 / logo.width, 48 / logo.height)
                logo.anchor = OneCellAnchor(
                    _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(8), rowOff=pixels_to_EMU(5)),
                    ext=XDRPositiveSize2D(pixels_to_EMU(logo.width * scale),
                                          pixels_to_EMU(logo.height * scale)),
                )
                summary.add_image(logo)
            summary.append(["HOYTAŞ HAZIR BETON | CARİ HESAP EKSTRESİ"])
            summary.append(["Cari Ünvanı", cari[0]])
            summary.append(["Telefon", cari[1] or "-"])
            summary.append(["Vergi No", cari[2] or "-"])
            debt = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Borç")
            credit = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Alacak")
            summary.append(["Toplam Borç (TL)", debt])
            summary.append(["Toplam Alacak (TL)", credit])
            summary.append(["Bakiye (Borç - Alacak) (TL)", debt-credit])
            quantities_and_prices = [(safe_float(row[2] or 0), safe_float(row[3] or 0)) for row in sales]
            total_m3 = sum(quantity for quantity, _ in quantities_and_prices)
            prices = {price for _, price in quantities_and_prices}
            unit_price = (sum(quantity * price for quantity, price in quantities_and_prices) / total_m3
                          if total_m3 > 0 and prices and min(prices) > 0 else None)
            price_label = ("Ort. Birim Fiyat (pompa ve nakliye hariç)" if len(prices) > 1
                           else "Birim Fiyat (pompa ve nakliye hariç)")
            total_pump = sum(safe_float(row[4] or 0) for row in sales)
            total_sales_amount = sum(safe_float(r[5] or 0) for r in sales)
            total_collected = sum(safe_float(r[5] or 0) for r in moves if r[1] == "Alacak" and r[2] not in ("Fatura Farkı", "Fatura") and ("tahsilat" in str(r[2]).casefold() or "tahsilat" in str(r[3]).casefold()))
            summary.append(["Toplam Beton Satışı (m³)", total_m3])
            quantity_format = '#,##0' if '.' not in format_m3(total_m3) else '#,##0.###'
            summary.cell(summary.max_row, 2).number_format = quantity_format + ' "m³"'
            summary.append([price_label, unit_price if unit_price is not None else "—"])
            summary.cell(summary.max_row, 2).number_format = '#,##0.00 "TL/m³"'
            money_rows = [("Toplam Satış Tutarı (TL)", total_sales_amount),
                          ("Toplam Tahsilat (TL)", total_collected)]
            total_transport = sum(row[9] for row in sales)
            if total_transport:
                money_rows.insert(0, ("Toplam Mikser Nakliye (TL)", total_transport))
            if total_pump:
                money_rows.insert(0, ("Toplam Pompa Ücreti (TL)", total_pump))
            for label, amount in money_rows:
                summary.append([label, amount])
                summary.cell(summary.max_row, 2).number_format = '#,##0.00 "TL"'
            summary.append(["Rapor Tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")])
            sh = wb.create_sheet("Cari Hareketleri")
            sh.append(["Tarih", "Tür", "Kategori", "Açıklama", "Ödeme Yöntemi", "Borç (TL)", "Alacak (TL)", "Bakiye (TL)"])
            balance = 0
            for r in moves:
                balance += safe_float(r[5] or 0) if r[1] == "Borç" else -safe_float(r[5] or 0)
                sh.append([display_date(r[0]), r[1], r[2], r[3], r[4], safe_float(r[5] or 0) if r[1] == "Borç" else 0, safe_float(r[5] or 0) if r[1] == "Alacak" else 0, balance])
            sh.append(["", "", "", "", "TOPLAM", debt, credit, balance])
            ss = wb.create_sheet("Beton Satışları")
            ss.append(["Tarih", "Beton Cinsi", "Miktar (m³)", "Birim Fiyat (TL/m³)", "Pompa Ücreti (TL)", "Mikser Nakliye (TL)", "Toplam (TL)", "Ödeme Durumu", "Açıklama"])
            for r in sales: ss.append([display_date(r[0]), r[1], safe_float(r[2] or 0), r[3], r[4], r[9], r[5], r[6], format_beton_aciklama(r[1],r[2],r[3],r[4],r[5],r[8],r[7],r[6],r[9])])
            ss.append(["", "TOPLAM BETON (m³)", total_m3,
                       "", "", sum(row[9] for row in sales), total_sales_amount, "", ""])
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
                    ws.row_dimensions[1].height = 43 if has_logo else 30
                    for rr in range(2, ws.max_row + 1):
                        ws.cell(rr, 1).font = label_font
                        if not has_logo:
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
                        if not has_logo and rr % 2 == 0:
                            for cell in ws[rr]:
                                cell.fill = accent_fill
                    if ws.max_row >= 2 and str(ws.cell(ws.max_row, 5).value or "").startswith("TOPLAM"):
                        for cell in ws[ws.max_row]:
                            cell.fill = yellow_fill
                            cell.font = Font(name="Aptos", bold=True, color="24583D")
                if ws is summary:
                    continue
                for col in ws.columns:
                    letter = get_column_letter(col[0].column)
                    max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                    ws.column_dimensions[letter].width = min(max(max_len + 2, 13), 46)
            for row in (5, 6, 7): summary.cell(row, 2).number_format = '#,##0.00 "TL"'
            for ws in (sh, ss, sc):
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if cell.column in (6, 7, 8) and ws.title == "Cari Hareketleri": cell.number_format = '#,##0.00'
                        if ws.title == "Beton Satışları" and cell.column in (4, 5, 6, 7): cell.number_format = '#,##0.00'
                        if ws.title == "Bağlı Çekler" and cell.column == 6: cell.number_format = '#,##0.00'
            prepare_cari_excel_print_layout(wb, logo_path)
            if preview and not self._preview_excel(wb, self):
                return
            path = filedialog.asksaveasfilename(parent=self, title=f"Cari Hesap Ekstresi – {unvan} – Excel Kaydet",
                defaultextension=".xlsx", filetypes=[("Excel çalışma kitabı", "*.xlsx")], initialfile=initial)
            if not path:
                return
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

    def bulk_cari_audit(self):
        try:
            with connect() as db:
                db.execute("BEGIN")
                rows = audit_all_cari(db)
                db.rollback()
            issues = [r for r in rows if r[2] or r[3] or r[4] or r[5]]
            if not issues:
                return messagebox.showinfo("Toplu cari kontrol", f"{len(rows)} cari kontrol edildi. Beton bağlantısı veya bakiye hesap hatası bulunmadı.", parent=self)
            win = tk.Toplevel(self)
            win.title("Toplu cari kontrol")
            win.geometry("920x480")
            win.minsize(650,360)
            footer=ttk.Frame(win,padding=8); footer.pack(side="bottom",fill="x")
            ttk.Label(win,text=f"{len(rows)} cari kontrol edildi; {len(issues)} inceleme gerekli.").pack(anchor="w",padx=12,pady=8)
            frame=ttk.Frame(win); frame.pack(fill="both",expand=True,padx=12)
            cols=("name","missing","wrong","ambiguous","balance","delta")
            tree=ttk.Treeview(frame,columns=cols,show="headings",selectmode="browse")
            for col,label,width in zip(cols,("Cari","Eksik","Tutar hatası","Belirsiz","Bakiye hatası","Olası fark TL"),(260,80,100,85,100,130)):
                tree.heading(col,text=label); tree.column(col,width=width)
            scroll=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
            tree.configure(yscrollcommand=scroll.set)
            scroll.pack(side="right",fill="y"); tree.pack(fill="both",expand=True)
            by_id = {str(row[0]): row for row in issues}
            for cid, name, missing, wrong, ambiguous, err, delta in issues:
                tree.insert("", "end", iid=str(cid),
                            values=(name, len(missing), len(wrong), len(ambiguous),
                                    "Var" if err else "Yok", f"{delta:,.2f}"))
            ttk.Label(footer, text="Bir cari seçin. Kesin tutar ve tarih hataları seçili cari için onarılır. "
                      "Eksik, belirsiz ve diğer hatalar elle incelenir.",
                      wraplength=620).pack(anchor="w", pady=4)

            def repair():
                selection = tree.selection()
                if not selection:
                    return
                selected = by_id[selection[0]]
                cid, name = selected[:2]
                planned = repairable_cari_movements(selected)
                if not planned:
                    return
                difference = sum(row[2] - row[4] for row in planned)
                if not messagebox.askyesno(
                    "Seçili cari için yedekli onarım",
                    f"{name}\n\nDüzeltilecek hareket: {len(planned)}\n"
                    f"Net borç farkı: {difference:,.2f} TL\n\nOnaylıyor musunuz?",
                    parent=win,
                ):
                    return
                try:
                    with connect() as db:
                        backup = _verified_sqlite_backup(db, "cari_onarim_v107")
                        db.execute("BEGIN IMMEDIATE")
                        current = db.execute("SELECT unvan FROM cari WHERE id=?", (cid,)).fetchone()
                        if current != (name,) or repairable_cari_movements(audit_cari(db, cid, name)) != planned:
                            raise RuntimeError("Seçilen carinin kayıtları değişti. Kontrolü yeniden çalıştırın.")
                        for _, date, amount, movement_id, _ in planned:
                            db.execute("UPDATE hareket SET tutar=?,tarih=? WHERE id=? AND cari_id=?",
                                       (amount, date, movement_id, cid))
                    win.destroy()
                    self._refresh_selected_module()
                    messagebox.showinfo("Onarım", f"{name} için {len(planned)} hareket düzeltildi.\nYedek: {backup}",
                                        parent=self)
                    self.bulk_cari_audit()
                except Exception as exc:
                    messagebox.showerror("Onarım hatası", str(exc), parent=self)

            def update_repair_button(_event):
                selection = tree.selection()
                enabled = selection and repairable_cari_movements(by_id[selection[0]])
                repair_button.state(["!disabled" if enabled else "disabled"])

            repair_button = ttk.Button(footer, text="Seçili Cariyi Onar", command=repair, state="disabled")
            repair_button.pack(side="left", padx=12, pady=8)
            tree.bind("<<TreeviewSelect>>", update_repair_button)
            ttk.Button(footer,text="Kapat",command=win.destroy).pack(side="right",padx=12,pady=8)
        except Exception as exc:
            messagebox.showerror("Cari kontrol hatası",str(exc),parent=self)

    def _build_cari(self):
        bar = ttk.Frame(self.cari_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Cari Hesaplar", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Cari Ekle", command=self.add_cari).pack(side="right")
        ttk.Button(bar,text="Tüm Cari Borçları Kontrol Et",command=self.bulk_cari_audit).pack(side="right",padx=6)
        legend=ttk.Frame(self.cari_tab); legend.pack(fill="x",pady=(6,0))
        ttk.Label(legend,text="● Alacak kaydı olmayan borç",foreground="#9B1C1C",font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,14))
        ttk.Label(legend,text="● Alacak kaydı olan açık borç",foreground="#8A4B08",font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,14))
        ttk.Label(legend,text="● Borç kapalı",foreground="#216438",font=("Segoe UI",9,"bold")).pack(side="left")
        self.cari_tree = self._tree(self.cari_tab, ("id","tur","unvan","telefon","vergi_no","borc","alacak","bakiye"),
                                    ("ID","Tür","Ünvan","Telefon","Vergi No","Borç (TL)","Alacak (TL)","Kalan Borç (TL)"),
                                    (45,85,220,120,110,115,115,125))
        self.cari_search = self.cari_tree._search_var
        self.cari_search_entry = self.cari_tree._search_entry
        self.cari_arsiv_goster = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="Arşivdekiler", variable=self.cari_arsiv_goster,
                        command=self.refresh_cari).pack(side="left", padx=10)
        self.cari_tree.bind("<Double-1>", lambda event: self.show_cari_hareketleri()
                            if self.cari_tree.identify_row(event.y) else None)
        row=ttk.Frame(self.cari_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kaydı Düzenle",command=lambda:self.edit_selected(self.cari_tree,"cari")).pack(side="left")
        ttk.Button(row,text="Cari Kartını Aç",command=self.show_cari_hareketleri).pack(side="left",padx=(0,8))
        self._action_menu(row, "Cari Ekstresi", [
            ("PDF", lambda: self.export_selected_cari_pdf("Cari Hesap Ekstresi")),
            ("Excel", self.export_selected_cari_excel),
        ]).pack(side="left", padx=(0, 8))
        ttk.Button(row,text="Tahsilat Al / Cari Borcunu Düş",command=self.cari_tahsilat).pack(side="left",padx=8)
        self._action_menu(row, "Cari İşlemleri", [
            ("Arşivle / Arşivden Çıkar", self.toggle_cari_archive),
            ("Kalıcı Sil (Yönetici)", lambda: self.delete_selected(self.cari_tree, "cari")),
        ]).pack(side="right")

    def _build_borc_takip(self):
        """Borçlu carileri tek listede göster; tutarlar hareketlerden hesaplanır."""
        bar = ttk.Frame(self.borc_takip_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Cari Borç Takip Listesi", font=("Segoe UI", 16, "bold")).pack(side="left")
        self.borc_tree = self._tree(self.borc_takip_tab,
            ("id", "unvan", "borc", "odenen", "kalan", "durum"),
            ("ID", "Cari Ünvanı", "Toplam Borç (TL)", "Toplam Alacak (TL)", "Kalan Borç (TL)", "Durum"),
            (50, 270, 130, 115, 135, 125), show_id=False, search_label="Cari ara:", print_button=False)
        self.borc_tree.bind("<Double-1>", self._open_debt_row)
        self.borc_tree.bind("<Return>", lambda _event: self._open_selected_debt_cari())
        self.borc_tree.tag_configure("debt_unpaid", foreground="#9B1C1C")
        self.borc_tree.tag_configure("debt_partial", foreground="#8A4B08")
        summary = ttk.Frame(self.borc_takip_tab); summary.pack(fill="x", pady=8)
        self.borc_total_label = ttk.Label(summary, text="Toplam kalan borç: 0,00 TL",
                                          font=("Segoe UI", 12, "bold"), foreground="#9B1C1C")
        self.borc_total_label.pack(side="left")
        actions = ttk.Frame(self.borc_takip_tab); actions.pack(fill="x", pady=(2, 0))
        ttk.Button(actions, text="Cari Kartını Aç", command=self._open_selected_debt_cari).pack(side="left")
        ttk.Button(actions, text="Açık Borç Listesi PDF", command=self._export_debt_list).pack(side="left")
        if os.name == "nt":
            ttk.Button(actions, text="Yazdır", command=lambda: self._export_debt_list(print_dialog=True)).pack(side="left")
        ttk.Button(actions, text="Listeyi Yenile", command=self._refresh_from_button).pack(side="left")

    def _open_debt_row(self, event):
        item = self.borc_tree.identify_row(event.y)
        if item and self.borc_tree.identify_region(event.x, event.y) == "cell":
            self.borc_tree.selection_set(item)
            self._open_selected_debt_cari()

    def _export_debt_list(self, *, print_dialog=False):
        tree = self.borc_tree
        rows = [tuple(tree.item(item, "values"))[1:] for item in tree.get_children()]
        if not rows:
            return messagebox.showinfo("Borç Takip", "Listede döküm alınacak kayıt yok.", parent=self)
        amounts = [sum(Decimal(row[index].replace(",", "")) for row in rows) for index in (1, 2, 3)]
        headings = tuple(tree.heading(column, "text") for column in tree["displaycolumns"])
        summary = ("TOPLAM", *(f"{amount:,.2f}" for amount in amounts), f"{len(rows)} cari")
        try:
            self._preview_then_save_pdf("Açık Borç Listesi", rows=[headings, *rows, summary],
                                        print_dialog=print_dialog)
        except (OSError, ValueError, RuntimeError, ImportError) as error:
            messagebox.showerror("Borç Takip Raporu", str(error), parent=self)

    def _debt_rows(self):
        with connect() as c:
            return c.execute("""SELECT cari_id,unvan,toplam_borc,toplam_alacak,bakiye
                FROM cari_bakiye WHERE bakiye > 0.005
                ORDER BY bakiye DESC,unvan COLLATE NOCASE""").fetchall()

    def refresh_borc_takip(self):
        if not hasattr(self, "borc_tree"):
            return
        rows = self._debt_rows()
        visible = []
        tags = {}
        for cid, name, debt, credit, remaining in rows:
            partial = safe_float(credit or 0) > 0.005
            status = "Alacak kaydı var" if partial else "Alacak kaydı yok"
            tags[str(cid)] = "debt_partial" if partial else "debt_unpaid"
            visible.append((cid, name, f"{debt:,.2f}", f"{credit:,.2f}", f"{remaining:,.2f}", status))
        self.borc_tree._row_tags = tags
        self.fill_tree(self.borc_tree, visible)

    def _open_selected_debt_cari(self):
        sel = self.borc_tree.selection()
        if not sel:
            return messagebox.showinfo("Cari seçin", "Hareketlerini görmek için borç listesinden bir cari seçin.", parent=self)
        values = self.borc_tree.item(sel[0], "values")
        self._open_module(self.cari_tab)
        if hasattr(self, "cari_search"):
            self.cari_search.set("")
            with connect() as db:
                archived = db.execute("SELECT arsivli FROM cari WHERE id=?", (int(values[0]),)).fetchone()
            self.cari_arsiv_goster.set(bool(archived and archived[0]))
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
            c.insert_shared("INSERT INTO cari(tur,unvan,telefon,vergi_no,notlar,created_at) VALUES(?,?,?,?,?,?)",
                ("Tedarikçi",unvan,d["telefon"],"",d["notlar"],datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self._refresh_selected_module()
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
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Borç",category,acik,tutar,"Vadeli"))
        self._refresh_selected_module()
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
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Alacak","Tedarikçi Ödemesi",acik,tutar,d["yontem"]))
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,"Gider","Alışveriş Ödemesi",acik,tutar,d["yontem"]))
            if d["yontem"]=="Nakit":
                c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                    (tarih,"Para Çıkışı","Tedarikçi Ödemesi",acik,tutar,getattr(self,"current_user",{}).get("username","")))
            else:
                c.insert_shared("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                    (tarih,d["banka"] or "Banka","", "Para Çıkışı",acik,tutar))
        self._refresh_selected_module()
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
            ("ID","Tarih","Tür","Kategori","Cari","Açıklama","Tutar (TL)","Ödeme"), (50,95,80,105,155,185,105,90), show_id=False)
        row=ttk.Frame(self.hareket_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kaydı Düzenle",command=lambda:self.edit_selected(self.hareket_tree,"hareket")).pack(side="left")
        ttk.Button(row, text="Seçili Kaydı Sil", command=lambda: self.delete_selected(self.hareket_tree,"hareket")).pack(side="right")

    def _build_cek(self):
        bar = ttk.Frame(self.cek_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Alınan / Verilen Çekler", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Çek Ekle", command=self.add_cek).pack(side="right")
        filters = ttk.Frame(self.cek_tab)
        filters.pack(fill="x", pady=(4, 8))
        self.cheque_view = tk.StringVar(self.cek_tab, "all")
        for label, value in (("Tüm Çekler", "all"), ("Kendi Çeklerimiz", "own"),
                             ("Tahsildeki Çekler", "collection"), ("Portföydeki Çekler", "portfolio")):
            ttk.Radiobutton(filters, text=label, value=value, variable=self.cheque_view,
                            command=self._change_cheque_view).pack(side="left", padx=(0, 12))
        reports = ttk.Frame(self.cek_tab)
        reports.pack(fill="x", pady=(0, 8))
        for title in CHEQUE_STATUS_REPORTS:
            ttk.Button(reports, text=title.removesuffix(" Dökümü") + " PDF",
                       command=lambda report=title: self.export_pdf_report(report)).pack(side="left", padx=(0, 6))
        self.cek_tree = self._tree(self.cek_tab, ("id","yon","banka","cek_no","taraf","vade","tutar","durum"),
            ("ID","Yön","Banka","Çek No","Keşideci / Lehtar","Vade","Tutar (TL)","Durum"), (50,100,130,120,220,110,120,110))
        row=ttk.Frame(self.cek_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Çeki Düzenle",command=lambda:self.edit_selected(self.cek_tree,"cek")).pack(side="left")
        ttk.Button(row,text="Alınan Çekler PDF",command=lambda:self.export_pdf_report("Alınan Çekler Dökümü")).pack(side="left",padx=(8,0))
        ttk.Button(row,text="Verilen Çekler PDF",command=lambda:self.export_pdf_report("Verilen Çekler Dökümü")).pack(side="left",padx=(8,0))
        ttk.Button(row, text="Seçili Çeki Sil", command=lambda: self.delete_selected(self.cek_tree,"cek")).pack(side="right")

    def _change_cheque_view(self):
        with connect() as db:
            self._refresh_cheque_view(db)

    def _refresh_cheque_view(self, db):
        rows = db.execute("""SELECT id,yon,COALESCE(banka,'') ||
            CASE WHEN COALESCE(banka_kodu,'')='' THEN '' ELSE ' (' || banka_kodu || ')' END,
            cek_no,COALESCE((SELECT unvan FROM cari WHERE cari.id=cek.cari_id),kesideci_lehtar),
            vade,printf('%.2f',tutar),durum FROM cek ORDER BY vade,id""")
        view = self.cheque_view.get()
        self.fill_tree(self.cek_tree, (row for row in rows if cheque_matches_view(row[1], row[7], view)))

    def _build_stok(self):
        bar = ttk.Frame(self.stok_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Stok / Depo", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Stok Kalemi Ekle", command=self.add_stok).pack(side="right")
        self.stok_tree = self._tree(self.stok_tab, ("id","kod","ad","birim","miktar","fiyat","deger"),
            ("ID","Kod","Malzeme","Birim","Miktar","Birim Fiyat","Stok Değeri (TL)"), (50,100,260,90,110,130,150))
        row=ttk.Frame(self.stok_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Kalemi Düzenle",command=lambda:self.edit_selected(self.stok_tree,"stok")).pack(side="left")
        ttk.Button(row, text="Seçili Kalemi Sil", command=lambda: self.delete_selected(self.stok_tree,"stok")).pack(side="right")

    def show_personel_karti(self, event=None):
        selected = self.personel_tree.selection()
        if not selected:
            return messagebox.showinfo("Personel seçin", "Önce bir personel seçin.", parent=self)
        personel_id, name = self.personel_tree.item(selected[0], "values")[:2]
        personel_id = int(personel_id)
        win = tk.Toplevel(self)
        win.title(f"Personel Kartı - {name}")
        win.geometry("940x460")
        win.transient(self)
        bar = ttk.Frame(win, padding=8)
        bar.pack(fill="x")
        payment = ttk.Frame(win, padding=(8, 0, 8, 8))
        payment.pack(fill="x")
        iban_value = tk.StringVar(win)
        ttk.Label(payment, text="IBAN").pack(side="left", padx=(0, 6))
        ttk.Entry(payment, textvariable=iban_value, width=30, state="readonly").pack(side="left")
        def copy_payment_iban():
            win.clipboard_clear()
            win.clipboard_append(iban_value.get())

        copy_iban = ttk.Button(payment, text="IBAN'ı Kopyala", command=copy_payment_iban)
        copy_iban.pack(side="left", padx=6)
        tree = self._tree(win, ("id", "tarih", "tur", "donem", "tutar", "yontem", "aciklama"),
                          ("ID", "Tarih", "İşlem Türü", "Dönem", "Tutar TL", "Ödeme Yöntemi", "Açıklama"),
                          (55, 95, 115, 95, 100, 120, 240))

        def refresh():
            with connect() as db:
                account = db.execute("""SELECT p.ad, COALESCE(o.iban,'') FROM personel p
                    LEFT JOIN personel_odeme_bilgisi o ON o.personel_id=p.id WHERE p.id=?""",
                    (personel_id,)).fetchone()
                rows = db.execute("""SELECT id,tarih,COALESCE(islem_turu,'Maaş'),donem,
                    printf('%.2f',tutar),COALESCE(yontem,''),COALESCE(aciklama,'')
                    FROM maas_odeme WHERE personel_id=? ORDER BY tarih,id""", (personel_id,)).fetchall()
            if account:
                win.title(f"Personel Kartı - {account[0]}")
                iban_value.set(account[1])
            copy_iban.state(["!disabled"] if iban_value.get() else ["disabled"])
            self.fill_tree(tree, rows)

        def edit_payment():
            self.edit_personel_payment(personel_id)
            refresh()

        ttk.Button(payment, text="Ad Soyad / IBAN Düzenle", command=edit_payment).pack(side="left")

        def add(kind):
            self.add_maas_odeme(personel_id, kind)
            refresh()

        for label, kind in (("Maaş", "Maaş"), ("Avans", "Nakit Avans"), ("Mesai", "Mesai"), ("Kesinti", "Kesinti")):
            ttk.Button(bar, text=label, command=lambda selected_kind=kind: add(selected_kind)).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Puantaj", command=lambda: self.show_puantaj(personel_id, refresh)).pack(side="left")
        ttk.Button(bar, text="Yenile", command=refresh).pack(side="right")
        tree._undo_table = "maas_odeme"
        tree._after_undo = refresh
        self._adapt_action_rows(win)
        refresh()

    def show_puantaj(self, personel_id=None, after_change=None):
        win = tk.Toplevel(self)
        win.title("HOYTAŞ Finans | Personel Puantajı")
        win.geometry("1150x620")
        win.minsize(900, 530)
        win.transient(self)
        first = date.today().replace(day=1)
        week = 0
        selected = None
        roster = {}
        daily = {}
        approved = {}
        draft_rates = {}
        rate = tk.StringVar(win)
        paid_days = tk.StringVar(win, "30")
        status = tk.StringVar(win, "Çalıştı")
        schedules = {"Kış 08.00–18.00": ("08:00", "18:00"),
                     "Yaz 07.00–17.00": ("07:00", "17:00")}
        schedule = tk.StringVar(win, next(iter(schedules)))
        default_start, default_end = schedules[schedule.get()]
        start = tk.StringVar(win, default_start)
        end = tk.StringVar(win, default_end)
        pause = tk.StringVar(win, "0")
        normal_hours = "10"
        manual = tk.StringVar(win)
        current_day = tk.StringVar(win, "Tablodan bir gün seçin")
        inline_entry = None
        period_label = tk.StringVar(win)
        summary = tk.StringVar(win)
        preview = tk.StringVar(win)
        statuses = ("Çalıştı", "İzin", "Devamsız", "Tatil")
        colors = {"Çalıştı": "#E2F3E5", "İzin": "#FFF0C6", "Devamsız": "#FADEE1", "Tatil": "#E9EDF3"}

        root = ttk.Frame(win, padding=14)
        root.pack(fill="both", expand=True)
        toolbar = ttk.Frame(root)
        toolbar.pack(fill="x", pady=(0, 12))
        ttk.Label(toolbar, text="Personel Puantajı", font=("Segoe UI", 19, "bold")).pack(side="left")
        schedule_box = ttk.Combobox(toolbar, textvariable=schedule, values=tuple(schedules),
                                    state="readonly", width=19)
        schedule_box.pack(side="right")
        ttk.Label(toolbar, text="Çalışma düzeni").pack(side="right", padx=8)
        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)
        panel = ttk.LabelFrame(body, text="Günlük Puantaj", padding=12)
        panel.pack(side="right", fill="y", padx=(12, 0))
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        navigation = ttk.Frame(left)
        navigation.pack(fill="x", pady=(0, 6))
        grid_area = ttk.Frame(left)
        grid_area.pack(fill="both", expand=True)
        ttk.Label(left, text="Tarihin altına mesai saatini yazın ve Enter’a basın.").pack(
            before=grid_area, anchor="w", pady=(0, 5))
        canvas = tk.Canvas(grid_area, highlightthickness=0, background="white")
        scroll = ttk.Scrollbar(grid_area, orient="vertical", command=canvas.yview)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=scroll.set)
        grid = tk.Frame(canvas, background="white")
        grid_window = canvas.create_window(0, 0, window=grid, anchor="nw")
        grid.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(grid_window, width=event.width))

        monthly = ttk.LabelFrame(left, text="Aylık Özet", padding=10)
        monthly.pack(fill="x", pady=(10, 0))
        ttk.Label(monthly, textvariable=summary, wraplength=620, justify="left").pack(anchor="w")
        settings = ttk.Frame(monthly)
        settings.pack(fill="x", pady=(8, 0))
        ttk.Label(settings, text="Saatlik mesai ücreti (TL)").pack(side="left")
        rate_entry = ttk.Entry(settings, textvariable=rate, width=7)
        rate_entry.pack(side="left", padx=5)
        daily_settings = ttk.Frame(monthly)
        daily_settings.pack(fill="x", pady=(5, 0))
        ttk.Label(daily_settings, text="Günlük ücret için maaş gün böleni").pack(side="left")
        ttk.Entry(daily_settings, textvariable=paid_days, width=7).pack(side="left", padx=5)

        ttk.Label(panel, textvariable=current_day, font=("Segoe UI", 11, "bold"),
                  wraplength=210).pack(anchor="w", pady=(0, 10))
        fields = ttk.Frame(panel)
        fields.pack(fill="x")
        for row, (title, variable) in enumerate((
            ("Durum", status), ("Başlangıç (SS:DD)", start), ("Bitiş (SS:DD)", end),
            ("Mola (dk)", pause),
            ("Manuel mesai (saat)", manual),
        )):
            ttk.Label(fields, text=title).grid(row=row, column=0, sticky="w", pady=5)
            if variable is status:
                widget = ttk.Combobox(fields, textvariable=variable, values=statuses, state="readonly", width=11)
            else:
                widget = ttk.Entry(fields, textvariable=variable, width=13)
            widget.grid(row=row, column=1, padx=(6, 0), sticky="ew")
            if variable is manual:
                manual_entry = widget
        ttk.Label(panel, text="Normal çalışma süresi 10 saat").pack(anchor="w", pady=(4, 0))
        ttk.Label(panel, text="Mesai yoksa boş bırakın.",
                  wraplength=210).pack(anchor="w", pady=(4, 0))
        ttk.Label(panel, textvariable=preview, wraplength=210).pack(anchor="w", pady=12)

        def minutes(value):
            hours, remainder = divmod(value, 60)
            if not remainder:
                return f"{hours} sa"
            return f"{hours} sa {remainder} dk" if hours else f"{remainder} dk"

        def day_range():
            return monthrange(first.year, first.month)[1]

        def month_key():
            return first.strftime("%Y-%m")

        def totals(pid):
            rows = [row for (employee, _), row in daily.items() if employee == pid]
            return (sum(row[0] == "Çalıştı" for row in rows), sum(row[0] == "İzin" for row in rows),
                    sum(row[0] == "Devamsız" for row in rows), sum(row[5] for row in rows),
                    sum(row[6] for row in rows))

        def show_summary(*_):
            if not selected:
                summary.set("Aylık özeti görmek için bir personel seçin.")
                return
            pid = selected[0]
            worked, leave, absent, net, overtime = totals(pid)
            name, salary = roster[pid]
            known = approved.get(pid)
            hourly = None
            if known:
                salary, divisor, multiplier, amount, saved_hourly = known
                amount = Decimal(str(amount))
                try:
                    hourly = (Decimal(saved_hourly) if saved_hourly is not None
                              else puantaj_mesai_saat_ucreti(salary, divisor, multiplier))
                except (ValueError, ArithmeticError):
                    pass
            else:
                try:
                    hourly = puantaj_girilen_saat_ucreti(rate.get())
                    amount = puantaj_girilen_mesai_tutari(hourly, overtime)
                except (ValueError, ArithmeticError):
                    amount = None

            def tl(value):
                rounded = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                return f"{rounded:,.2f}".translate(str.maketrans(",.", ".,")) + " TL"

            def rate_label(divisor, unit):
                try:
                    return tl(puantaj_ucret_bol(salary, divisor)) + f" / {unit}"
                except (ValueError, ArithmeticError):
                    return "—"

            if amount is None and overtime:
                payment = "Aylık mesai tutarı için saatlik mesai ücretini girin."
            else:
                amount = amount if amount is not None else Decimal(0)
                label = "Onaylı fazla mesai" if known else "Aylık fazla mesai ücreti"
                formula = ""
                if hourly is not None:
                    displayed_rate = hourly.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    displayed_total = (displayed_rate * overtime / 60).quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP)
                    equals = "=" if displayed_total == amount else "≈"
                    formula = f"{tl(hourly)} / sa × {minutes(overtime)} {equals} "
                payment = f"{label}: {formula}{tl(amount)}\nMaaş + mesai {tl(Decimal(str(salary)) + amount)}"
            summary.set(
                f"{name}  ·  Aylık maaş {tl(salary)}\n"
                f"Toplam çalışma {minutes(net)}  ·  Fazla mesai {minutes(overtime)}\n"
                f"Çalışma {worked} gün  ·  İzin {leave} gün  ·  Devamsızlık {absent} gün\n"
                f"Günlük ücret {rate_label(paid_days.get(), 'gün')}\n"
                f"{payment}"
            )
            approve_button.configure(state="disabled" if known or not overtime or amount is None else "normal")
            revoke_button.configure(state="normal" if known else "disabled")

        def daily_durations():
            override = puantaj_manuel_mesai(manual.get()) if status.get() == "Çalıştı" else None
            durations = puantaj_sureleri(status.get(), start.get(), end.get(), pause.get(), normal_hours, override or 0)
            return durations, override

        def daily_preview(*_):
            locked = selected and selected[0] in approved
            state = "disabled" if locked or status.get() != "Çalıştı" else "normal"
            manual_entry.configure(state=state)
            if inline_entry is not None:
                inline_entry.configure(state=state)
            try:
                (_, _, net, overtime), _ = daily_durations()
                preview.set(f"Net çalışma {minutes(net)}\nFazla mesai {minutes(overtime)}")
            except ValueError as error:
                preview.set(str(error))
            except ArithmeticError:
                preview.set("Saatleri ve mola süresini kontrol edin.")

        def choose(pid, day, *, focus_cell=False):
            nonlocal selected
            previous_pid = selected[0] if selected else None
            if previous_pid is not None and previous_pid != pid:
                if previous_pid not in approved:
                    draft_rates[previous_pid] = rate.get()
            selected = (pid, day)
            if previous_pid != pid:
                rate.set(draft_rates.get(pid, ""))
            current_day.set(f"{roster[pid][0]} · {first.replace(day=day):%d.%m.%Y}")
            row = daily.get(selected)
            default_start, default_end = schedules[schedule.get()]
            if row and row[0] == "Çalıştı":
                values = (*row[:3], str(row[3]))
            else:
                values = (row[0] if row else "Çalıştı", default_start, default_end, "0")
            for variable, value in zip((status, start, end, pause), values):
                variable.set(value)
            stored_minutes = (row[7] if row[7] is not None else row[6]) if row else 0
            manual.set(str(Decimal(stored_minutes) / 60).replace(".", ",") if stored_minutes else "")
            locked = pid in approved
            if locked:
                rate.set(approved[pid][4] or "")
            rate_entry.configure(state="readonly" if locked else "normal")
            save_button.configure(state="disabled" if locked else "normal")
            delete_button.configure(state="normal" if row and not locked else "disabled")
            show_summary()
            render()
            if focus_cell and inline_entry is not None:
                inline_entry.focus_set()
                inline_entry.selection_range(0, "end")

        def render():
            nonlocal inline_entry
            inline_entry = None
            for child in grid.winfo_children():
                child.destroy()
            days = list(range(week * 7 + 1, min(week * 7 + 7, day_range()) + 1))
            months = ("Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
                      "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık")
            period_label.set(f"{days[0]}–{days[-1]} {months[first.month - 1]} {first.year}")
            for col in range(9):
                grid.columnconfigure(col, weight=0, minsize=0)
            headings = ["Personel", *(f"{day:02}" for day in days), "Aylık Mesai"]
            for col, title in enumerate(headings):
                grid.columnconfigure(col, weight=2 if col == 0 else 1, minsize=64 if col else 146)
                tk.Label(grid, text=title, bg="#102D53", fg="white", pady=8).grid(
                    row=0, column=col, sticky="nsew", padx=1, pady=1)
            for index, (pid, (name, _)) in enumerate(roster.items(), 1):
                tk.Label(grid, text=name, anchor="w", padx=7, bg="#ECF1F8", wraplength=145).grid(
                    row=index, column=0, sticky="nsew", pady=1)
                for col, day in enumerate(days, 1):
                    row = daily.get((pid, day))
                    label, bg = "—", "#F3F6FA"
                    if row:
                        label = row[0] if row[0] != "Çalıştı" else minutes(row[5] - row[6])
                        if row[6]:
                            label += f" + {minutes(row[6])}"
                        bg = "#DCEAFF" if row[6] else colors[row[0]]
                    active = (pid, day) == selected
                    if active and pid not in approved and (not row or row[0] == "Çalıştı"):
                        inline_entry = ttk.Entry(grid, textvariable=manual, width=6, justify="center")
                        inline_entry.bind("<Return>", lambda event: record())
                        inline_entry.bind("<KP_Enter>", lambda event: record())
                        cell = inline_entry
                    else:
                        cell = tk.Button(grid, text=label, wraplength=85, bg=bg,
                                         relief="solid" if active else "flat", bd=int(active),
                                         command=lambda p=pid, d=day: choose(p, d, focus_cell=True))
                    cell.grid(row=index, column=col, sticky="nsew", padx=1, pady=1, ipady=6)
                tk.Label(grid, text=minutes(totals(pid)[4]), bg="#ECF1F8").grid(
                    row=index, column=len(days) + 1, sticky="nsew", pady=1)

        def refresh():
            nonlocal roster, daily, approved, selected
            until = first.replace(day=day_range()).isoformat()
            with connect() as db:
                roster = {pid: (name, salary) for pid, name, salary in db.execute(
                    "SELECT id,ad,maas FROM personel WHERE ? IS NULL OR id=? ORDER BY ad",
                    (personel_id, personel_id))}
                daily = {(pid, date.fromisoformat(day).day): tuple(values)
                         for pid, day, *values in db.execute("""SELECT personel_id,tarih,durum,baslangic,bitis,mola,
                             normal_dakika,net_dakika,mesai_dakika,manuel_mesai_dakika FROM personel_puantaj
                             WHERE tarih BETWEEN ? AND ?""", (first.isoformat(), until))}
                approved = {pid: tuple(values) for pid, *values in db.execute("""
                    SELECT a.personel_id,a.maas,a.bolen,a.katsayi,m.tutar,a.saatlik_mesai_ucreti
                    FROM puantaj_onay a JOIN maas_odeme m ON m.id=a.odeme_id WHERE a.donem=?""", (month_key(),))}
            if roster:
                pid = selected[0] if selected and selected[0] in roster else next(iter(roster))
                day = min(week * 7 + 7, day_range(), max(week * 7 + 1, selected[1] if selected else date.today().day))
                choose(pid, day)
            else:
                selected = None
                current_day.set("Personel kaydı bulunamadı")
                for button in (save_button, delete_button, approve_button, revoke_button):
                    button.configure(state="disabled")
                render()
                show_summary()

        def move_month(delta):
            nonlocal first, week, selected
            target = first + timedelta(days=day_range() if delta > 0 else -1)
            first, week, selected = target.replace(day=1), 0, None
            draft_rates.clear()
            refresh()

        def move_week(delta):
            nonlocal week
            week = max(0, min((day_range() - 1) // 7, week + delta))
            refresh()

        def record(remove=False):
            if not selected:
                return
            pid, day = selected
            day_key = first.replace(day=day).isoformat()
            try:
                with connect() as db:
                    db.execute("BEGIN IMMEDIATE")
                    if db.execute("SELECT 1 FROM puantaj_onay WHERE personel_id=? AND donem=?",
                                  (pid, month_key())).fetchone():
                        raise ValueError("Dönem onaylandı. Düzenlemek için önce mesai onayını kaldırın.")
                    if remove:
                        db.execute("DELETE FROM personel_puantaj WHERE personel_id=? AND tarih=?", (pid, day_key))
                    else:
                        durations, override = daily_durations()
                        times = (start.get(), end.get()) if status.get() == "Çalıştı" else ("", "")
                        db.execute("""INSERT INTO personel_puantaj
                            (personel_id,tarih,durum,baslangic,bitis,mola,normal_dakika,net_dakika,mesai_dakika,manuel_mesai_dakika)
                            VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(personel_id,tarih) DO UPDATE SET
                            durum=excluded.durum,baslangic=excluded.baslangic,bitis=excluded.bitis,
                            mola=excluded.mola,normal_dakika=excluded.normal_dakika,net_dakika=excluded.net_dakika,
                            mesai_dakika=excluded.mesai_dakika, manuel_mesai_dakika=excluded.manuel_mesai_dakika""",
                            (pid, day_key, status.get(), *times, *durations, override))
                refresh()
            except (ValueError, ArithmeticError, sqlite3.Error) as error:
                messagebox.showerror("Puantaj kaydedilemedi", str(error), parent=win)

        def approve(revoke=False):
            if not selected:
                return
            pid = selected[0]
            if not messagebox.askyesno("Mesai onayı", "Onay ve bağlı maaş mesai kaydı kaldırılsın mı?" if revoke
                                      else "Hesaplanan fazla mesai maaş işlemlerine eklensin mi?", parent=win):
                return
            try:
                with connect() as db:
                    db.execute("BEGIN IMMEDIATE")
                    save_puantaj_approval(db, pid, first, hourly=rate.get(), revoke=revoke)
                refresh()
                self._refresh_selected_module()
                if after_change:
                    after_change()
            except (ValueError, ArithmeticError, sqlite3.Error) as error:
                messagebox.showerror("Mesai onayı tamamlanamadı", str(error), parent=win)

        def report(kind):
            rows = [("Personel", "Tarih", "Durum", "Başlangıç", "Bitiş", "Net", "Fazla mesai")]
            for (pid, day), row in sorted(daily.items(), key=lambda item: (item[0][1], item[0][0])):
                if pid in roster:
                    rows.append((roster[pid][0], f"{first.replace(day=day):%d.%m.%Y}", *row[:3],
                                 minutes(row[5]), minutes(row[6])))
            for pid, (name, salary) in roster.items():
                worked, leave, absent, net, overtime = totals(pid)
                known = approved.get(pid)
                if known:
                    amount = f"Onaylı {known[3]:.2f} TL"
                else:
                    try:
                        hourly = rate.get() if selected and selected[0] == pid else draft_rates.get(pid, "")
                        amount = f"Taslak {puantaj_girilen_mesai_tutari(hourly, overtime):.2f} TL"
                    except (ValueError, ArithmeticError):
                        amount = ""
                rows.append((name, "Aylık toplam", f"{worked} çalışma / {leave} izin", f"{absent} devamsız",
                             amount, minutes(net), minutes(overtime)))
            title = f"Personel Puantajı {month_key()}"
            try:
                if kind == "Yazdır":
                    self._print_report_rows(title, rows)
                elif kind == "PDF":
                    self._preview_then_save_pdf(title, rows=rows)
                else:
                    from openpyxl import Workbook
                    path = filedialog.asksaveasfilename(parent=win, defaultextension=".xlsx",
                                                        filetypes=[("Excel", "*.xlsx")],
                                                        initialfile=f"HOYTAS_Puantaj_{month_key()}.xlsx")
                    if path:
                        book = Workbook()
                        sheet = book.active
                        sheet.title = "Puantaj"
                        sheet.freeze_panes = "B2"
                        for values in rows:
                            sheet.append(values)
                            for cell in sheet[sheet.max_row]:
                                if cell.data_type == "f":
                                    cell.data_type = "s"
                        for column in "ABCDEFG":
                            sheet.column_dimensions[column].width = 24
                        book.save(path)
            except (OSError, ValueError, RuntimeError, ImportError) as error:
                messagebox.showerror("Puantaj raporu oluşturulamadı", str(error), parent=win)

        for label, command in (("Önceki Ay", lambda: move_month(-1)), ("Sonraki Ay", lambda: move_month(1)),
                               ("Yenile", refresh)):
            ttk.Button(toolbar, text=label, command=command).pack(side="left", padx=(8, 0))
        for kind in ("Yazdır", "PDF", "Excel"):
            ttk.Button(toolbar, text=kind, command=lambda output=kind: report(output)).pack(side="right", padx=3)
        ttk.Button(navigation, text="‹ Önceki Hafta", command=lambda: move_week(-1)).pack(side="left")
        ttk.Button(navigation, text="Sonraki Hafta ›", command=lambda: move_week(1)).pack(side="right")
        ttk.Label(navigation, textvariable=period_label, font=("Segoe UI", 12, "bold")).pack()
        save_button = ttk.Button(panel, text="Kaydet", command=record)
        save_button.pack(fill="x")
        delete_button = ttk.Button(panel, text="Günü Sil", command=lambda: (
            messagebox.askyesno("Puantaj", "Günlük kayıt silinsin mi?", parent=win) and record(True)))
        delete_button.pack(fill="x", pady=5)
        ttk.Button(panel, text="İptal", command=lambda: selected and choose(*selected)).pack(fill="x")
        approve_button = ttk.Button(settings, text="Mesaiyi Onayla", command=approve)
        approve_button.pack(side="right")
        revoke_button = ttk.Button(settings, text="Onayı Kaldır", command=lambda: approve(True))
        revoke_button.pack(side="right", padx=5)
        def apply_schedule(_event):
            if selected and selected[0] in approved:
                return
            shift_start, shift_end = schedules[schedule.get()]
            start.set(shift_start)
            end.set(shift_end)

        schedule_box.bind("<<ComboboxSelected>>", apply_schedule)
        for variable in (status, start, end, pause, manual):
            variable.trace_add("write", daily_preview)
        for variable in (rate, paid_days):
            variable.trace_add("write", show_summary)
        week = (date.today().day - 1) // 7
        refresh()

    def _build_personel(self):
        bar = ttk.Frame(self.personel_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Personel", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Puantaj", command=self.show_puantaj).pack(side="left", padx=12)
        ttk.Button(bar, text="Personel Ekle", command=self.add_personel).pack(side="right")
        self.personel_tree = self._tree(self.personel_tab, ("id","ad","gorev","telefon","maas","notlar"),
            ("ID","Ad Soyad","Görev","Telefon","Aylık Maaş (TL)","Notlar"), (50,220,180,140,150,300))
        row=ttk.Frame(self.personel_tab); row.pack(fill="x",pady=8)
        self.personel_tree.bind("<Double-1>", self.show_personel_karti)
        ttk.Button(row, text="Personel Kartını Aç", command=self.show_personel_karti).pack(side="left", padx=(0, 8))
        ttk.Button(row,text="Seçili Personeli Düzenle",command=lambda:self.edit_selected(self.personel_tree,"personel")).pack(side="left")
        ttk.Button(row, text="Seçili Personeli Sil", command=lambda: self.delete_selected(self.personel_tree,"personel")).pack(side="right")

    def _build_mastarci(self):
        bar = ttk.Frame(self.mastarci_tab)
        bar.pack(fill="x")
        ttk.Label(bar, text="Mastarcı Hesabı", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Mastar İşi", command=self._edit_mastarci).pack(side="right")
        ttk.Label(self.mastarci_tab, text="Dışarıdan alınan mastar işleri, hakediş ve ödemeler",
                  foreground="#315B46").pack(anchor="w", pady=(2, 8))
        self.mastarci_total = ttk.Label(self.mastarci_tab, font=("Segoe UI", 11, "bold"))
        self.mastarci_total.pack(anchor="w", pady=6)
        actions = ttk.Frame(self.mastarci_tab)
        actions.pack(fill="x", side="bottom", pady=(8, 0))
        ttk.Button(actions, text="Düzenle", command=lambda: self._edit_mastarci("edit")).pack(side="left")
        ttk.Button(actions, text="Ödeme Ekle", command=lambda: self._edit_mastarci("pay")).pack(side="left", padx=6)
        ttk.Button(actions, text="Yenile", command=self.refresh_all).pack(side="left")
        ttk.Button(actions, text="Sil", command=lambda: self._edit_mastarci("delete")).pack(side="right")
        self.mastarci_tree = self._tree(self.mastarci_tab,
            ("id", "tarih", "mastarci", "musteri", "santiye", "m3", "fiyat", "hakedis", "odenen", "kalan", "aciklama"),
            ("ID", "İş Tarihi", "Mastarcı / Ekip", "Müşteri", "Şantiye", "Miktar (m³)", "TL / m³",
             "Hakediş (TL)", "Ödenen (TL)", "Kalan (TL)", "Açıklama"),
            (38, 96, 144, 147, 125, 90, 90, 106, 102, 102, 136))
        self.mastarci_tree.bind("<Double-1>", lambda _event: self._edit_mastarci("edit"))
        self.mastarci_tree.unbind("<Button-3>")
        self.mastarci_tree.unbind("<Control-Button-1>")

    def _edit_mastarci(self, action="add"):
        tree = self.mastarci_tree
        record_id = None
        try:
            if action != "add":
                selected = tree.selection()
                if not selected:
                    return messagebox.showinfo("İş seçin", "Önce listeden bir mastar işi seçin.", parent=self)
                record_id = int(tree.item(selected[0], "values")[0])
                with connect() as db:
                    row = db.execute("""SELECT tarih,mastarci,santiye,miktar_m3,birim_fiyat,hakedis_kurus,
                        odenen_kurus,aciklama,musteri FROM mastarci_hesap WHERE id=?""", (record_id,)).fetchone()
                if row is None:
                    return self._refresh_selected_module()
            if action == "delete":
                if not messagebox.askyesno("Mastar işini sil", "Seçili iş ve ödeme bilgisi silinsin mi?", parent=self):
                    return
                with connect() as db:
                    db.execute("DELETE FROM mastarci_hesap WHERE id=?", (record_id,))
            elif action == "pay":
                remaining = Decimal(row[5] - row[6]) / 100
                if remaining <= 0:
                    return messagebox.showinfo("Ödeme tamamlandı", "Bu işin kalan bakiyesi yok.", parent=self)
                data = self.prompt(f"Mastarcı Ödemesi · Kalan {remaining:.2f} TL",
                                   [("tutar", "Ödeme (TL)", "entry", None)])
                if not data:
                    return
                amount = mastarci_odeme_kurus(data["tutar"])
                with connect() as db:
                    changed = db.execute("""UPDATE mastarci_hesap SET odenen_kurus=odenen_kurus+?
                        WHERE id=? AND odenen_kurus+?<=hakedis_kurus""", (amount, record_id, amount))
                    if changed.rowcount != 1:
                        raise ValueError("Ödeme güncel kalan bakiyeyi aşamaz.")
            else:
                with connect() as db:
                    names = [name for name, in db.execute(
                        "SELECT DISTINCT mastarci FROM mastarci_hesap ORDER BY mastarci")]
                    customers = [name for name, in db.execute("""SELECT unvan FROM cari
                        WHERE tur='Müşteri' AND TRIM(unvan)<>''
                        UNION SELECT musteri FROM mastarci_hesap WHERE TRIM(musteri)<>''
                        ORDER BY 1""")]
                fields = [
                    ("tarih", "İş tarihi (GG.AA.YYYY)", "entry", None),
                    ("mastarci", "Mastarcı / Ekip", "combo_edit", names),
                    ("musteri", "Müşteri", "combo_edit", customers),
                    ("santiye", "Şantiye", "entry", None),
                    ("m3", "Miktar (m³)", "entry", None),
                    ("fiyat", "Birim fiyat (TL/m³)", "entry", None),
                    ("aciklama", "Açıklama", "entry", None),
                ]
                initial = {"tarih": date.today().isoformat()}
                if record_id is not None:
                    initial = dict(zip(("tarih", "mastarci", "santiye", "m3", "fiyat", "aciklama", "musteri"),
                                       (*row[:5], row[7], row[8])))
                data = self.prompt("Mastarcı İşi", fields, initial)
                if not data:
                    return
                if not data["mastarci"]:
                    raise ValueError("Mastarcı veya ekip adı girin.")
                if not data["musteri"] or not data["tarih"]:
                    raise ValueError("Müşteri adı ve iş tarihi girin.")
                quantity, price, total = mastarci_hakedis(data["m3"], data["fiyat"])
                values = (data["tarih"], data["mastarci"], data["santiye"], str(quantity), str(price),
                          total, data["aciklama"], data["musteri"])
                with connect() as db:
                    if record_id is None:
                        db.execute("""INSERT INTO mastarci_hesap
                            (tarih,mastarci,santiye,miktar_m3,birim_fiyat,hakedis_kurus,aciklama,musteri)
                            VALUES(?,?,?,?,?,?,?,?)""", values)
                    else:
                        changed = db.execute("""UPDATE mastarci_hesap SET tarih=?,mastarci=?,santiye=?,
                            miktar_m3=?,birim_fiyat=?,hakedis_kurus=?,aciklama=?,musteri=?
                            WHERE id=? AND odenen_kurus<=?""", (*values, record_id, total))
                        if changed.rowcount != 1:
                            raise ValueError("Hakediş güncel ödenen tutarın altına inemez.")
            self._refresh_selected_module()
        except (ValueError, ArithmeticError, sqlite3.Error) as error:
            messagebox.showerror("Mastarcı işlemi kaydedilemedi", str(error), parent=self)

    def _build_beton(self):
        bar = ttk.Frame(self.beton_tab); bar.pack(fill="x")
        ttk.Label(bar, text="Beton Satış ve Metreküp Takibi", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(bar, text="Yeni Beton Satışı", command=self.add_beton).pack(side="right")
        self.beton_tree = self._tree(self.beton_tab, ("id","tarih","cari","cinsi","m3","fiyat","pompa","nakliye","toplam","durum","aciklama"),
            ("ID","Tarih","Müşteri","Beton Cinsi","Miktar (m³)","Birim Fiyat (TL)","Pompa Ücreti (TL)","Nakliye (TL)","Toplam (TL)","Ödeme Durumu","Açıklama"),
            (45,100,190,100,95,125,125,110,120,125,200))
        self.beton_tree.bind("<Double-1>", lambda event: self._edit_tree_row(event, self.beton_tree, "beton_satis"))
        row=ttk.Frame(self.beton_tab); row.pack(fill="x",pady=8)
        ttk.Button(row,text="Seçili Satışı Düzenle",command=lambda:self.edit_selected(self.beton_tree,"beton_satis")).pack(side="left")
        self._action_menu(row, "Beton Satış Dökümü", [
            ("Tüm Beton Satışları PDF", lambda: self.export_pdf_report("Beton Satışları Dökümü")),
            ("Seçili Satışın Müşterisi PDF", self.export_selected_beton_cari_pdf),
            ("Müşteri Seçerek PDF", self.export_beton_cari_sec_pdf),
        ]).pack(side="left", padx=(8, 0))
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
            c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                (tarih,d["yon"],d["kaynak"],d["aciklama"],tutar,self.current_user.get("username","")))
        self._refresh_selected_module()

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
                widget = ttk.Entry(body, width=33, show="*" if kind == "secret" else "")
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

    def _edit_tree_row(self, event, tree, table):
        item = tree.identify_row(event.y)
        if item and tree.identify_region(event.x, event.y) in ("cell", "tree"):
            tree.selection_set(item)
            self.edit_selected(tree, table)
            return "break"

    def _resolve_site(self, value):
        value=(value or "").strip()
        if not value: return None
        if " - " in value and value.split(" - ",1)[0].isdigit():
            return int(value.split(" - ",1)[0])
        with connect() as c:
            row=c.execute("SELECT id FROM santiye WHERE lower(ad)=lower(?)",(value,)).fetchone()
            if row: return row[0]
            cur=c.insert_shared("INSERT INTO santiye(ad) VALUES(?)",(value,))
            return cur.lastrowid

    def edit_selected(self, tree, table):
        sel=tree.selection()
        if not sel: return messagebox.showinfo("Kayıt seçin","Düzenlemek için önce bir satır seçin.",parent=self)
        rid=int(tree.item(sel[0],"values")[0])
        if table in ("fatura", "beton_satis", "hareket"):
            with connect() as db:
                protected = record_is_posted(db, table, rid)
        else:
            protected = False
        if protected:
            return messagebox.showinfo("Bağlı fatura kaydı", "Faturaya bağlanmış tutar doğrudan düzenlenemez. "
                "Tahsilat yoksa eşleştirme işlemini İşlem Geri Al ile geri alabilirsiniz.", parent=self)
        if table == "fatura":
            with connect() as check_c:
                invoice = check_c.execute("SELECT odenen,tevkifat_oran FROM fatura WHERE id=?", (rid,)).fetchone()
            if invoice and (invoice[0] > 0.005 or invoice[1] > 0.005):
                return messagebox.showinfo("Fatura koruması",
                    "Tahsilat yapılmış veya tevkifatlı faturalar bu genel düzenleme ekranından değiştirilemez. "
                    "Finansal hesapları korumak için düzeltme işlemi ayrı bir fatura revizyonuyla yapılmalıdır.",parent=self)
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
          "hareket":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("tur","Tür","combo",["Gelir","Gider","Borç","Alacak"]),("kategori","Kategori","entry",None),("cari_id","Cari hesap ara / seçin","combo_search",None),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None),("odeme_yontemi","Ödeme yöntemi","entry",None)],
          "cek":[("yon","Yön","combo",["Alınan","Verilen"]),("banka","Banka adı","entry",None),("banka_kodu","Banka kodu","entry",None),("cari_id","Cari hesap ara / seçin","combo_search",None),("cek_no","Çek No","entry",None),("kesideci_lehtar","Keşideci / Lehtar","entry",None),("vade","Vade (GG.AA.YYYY)","entry",None),("tutar","Tutar (TL)","entry",None),("durum","Durum","combo",list(CHEQUE_STATES)),("notlar","Notlar","entry",None)],
          "stok":[("kod","Stok Kodu","entry",None),("ad","Malzeme Adı","entry",None),("birim","Birim","combo",["Adet","Kg","Ton","m³","Litre","Metre","Paket"]),("miktar","Miktar","entry",None),("birim_fiyat","Birim Fiyat (TL)","entry",None)],
          "personel":[("ad","Ad Soyad","entry",None),("gorev","Görev","entry",None),("telefon","Telefon","entry",None),("maas","Aylık Maaş (TL)","entry",None),("notlar","Notlar","entry",None)],
          "beton_satis":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("cari_id","Cari hesap ara / seçin","combo_search",None),("santiye_id","Gideceği şantiye (boş bırakılabilir, elle yazılabilir)","combo_edit",[]),("beton_cinsi","Beton Cinsi","combo",["C16","C20","C25","C30","C35","C40","Şap"]),("miktar_m3","Miktar (m³; ör. 30,5)","entry",None),("birim_fiyat","Birim Fiyat (TL/m³)","entry",None),("pompa_ucreti","Pompa Ücreti (TL)","entry",None),("mikser_nakliye_ucreti","Mikser Nakliye Ücreti (TL)","entry",None),("odeme_durumu","Ödeme Durumu","combo",["Vadeli / Tahsil Edilmedi","Tahsil Edildi"]),("aciklama","Açıklama","entry",None)],
          "fatura":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("cari_id","Cari hesap ara / seçin","combo_search",None),("belge_no","Belge No","entry",None),("aciklama","Açıklama","entry",None),("miktar_m3","Miktar (m³)","entry",None),("birim_fiyat","Birim Fiyat","entry",None),("kdv_oran","KDV (%)","entry",None),("durum","Durum","combo",["Taslak","Hazır","İptal edildi"])],
          "nakit_kasa":[("tarih","Tarih (GG.AA.YYYY)","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("kaynak","Kaynak","entry",None),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None)],
          "santiye":[("ad","Şantiye adı","entry",None),("adres","Adres","entry",None),("yetkili","Yetkili","entry",None),("telefon","Telefon","entry",None),("notlar","Notlar","entry",None)],
          "stok_hareket":[("tarih","Tarih","entry",None),("stok_id","Stok ID","entry",None),("yon","Yön","combo",["Giriş","Çıkış"]),("miktar","Miktar","entry",None),("aciklama","Açıklama","entry",None)],
          "maas_odeme":[("tarih","Tarih","entry",None),("personel_id","Personel ID","entry",None),("islem_turu","İşlem Türü","combo",["Maaş","Nakit Avans","Mesai","Kesinti"]),("donem","Dönem (YYYY-AA)","entry",None),("tutar","Tutar","entry",None),("yontem","Ödeme yöntemi","entry",None),("aciklama","Açıklama","entry",None)],
          "banka_hareket":[("tarih","Tarih","entry",None),("banka","Banka","entry",None),("hesap","Hesap / IBAN sonu","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("karsi_taraf","Gönderici / Alıcı Ad Soyad veya Ünvan","entry",None),("aciklama","Açıklama","entry",None),("tutar","Tutar","entry",None)]}
        if table not in specs: return messagebox.showinfo("Bilgi","Bu listedeki düzenleme henüz desteklenmiyor.",parent=self)
        cols=[f[0] for f in specs[table]]
        with connect() as c: row=c.execute("SELECT "+",".join(cols)+f" FROM {table} WHERE id=?",(rid,)).fetchone()
        if not row: return messagebox.showerror("Kayıt bulunamadı","Seçilen kayıt artık mevcut değil.",parent=self)
        fields=[]
        initial=dict(zip(cols,row))
        for spec,val in zip(specs[table],row):
            key,label,kind,opts=spec
            if key == "cari_id":
                with connect() as c:
                    customers = c.execute("SELECT id,unvan FROM cari ORDER BY unvan COLLATE NOCASE,id").fetchall()
                choices = {cid: f"{cid} - {name}" for cid, name in customers}
                empty = "(Cari seçilmedi)"
                opts = [empty, *choices.values()]
                initial[key] = choices.get(val, empty)
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
                elif key == "cari_id":
                    with connect() as c:
                        v = resolve_customer_id(c, v, allowed_types=None)
                elif key in ("santiye_id","stok_id","personel_id"):
                    v=int(v) if str(v).strip() else None
                elif key in ("tutar","miktar","miktar_m3","birim_fiyat","pompa_ucreti","mikser_nakliye_ucreti","maas","kdv_oran"):
                    v=parse_amount(v) if str(v).strip() else 0
                vals.append(v)
            data=dict(zip(cols,vals))
            if "tarih" in data and data["tarih"]: datetime.strptime(data["tarih"],"%Y-%m-%d")
            if table=="beton_satis":
                *_, data["toplam"] = concrete_sale_totals(
                    data["miktar_m3"], data["birim_fiyat"], data["pompa_ucreti"], data["mikser_nakliye_ucreti"])
                if data["cari_id"] is not None:
                    with connect() as check_c: 
                        if not check_c.execute("SELECT 1 FROM cari WHERE id=?",(data["cari_id"],)).fetchone(): raise ValueError("Seçilen cari hesap bulunamadı.")
            if table=="fatura":
                data["miktar_m3"]=parse_amount(data["miktar_m3"]); data["ara_toplam"]=round(data["miktar_m3"]*data["birim_fiyat"],2); data["kdv_tutar"]=round(data["ara_toplam"]*data["kdv_oran"]/100,2); data["genel_toplam"]=round(data["ara_toplam"]+data["kdv_tutar"],2); data["tahsil_edilecek"]=data["genel_toplam"]
            if table in ("stok","personel") and any(v is None for v in vals): raise ValueError("Zorunlu alanları kontrol edin.")
            for key in ("tutar","miktar","miktar_m3","birim_fiyat","maas"):
                if key in data and data[key] is not None and data[key] < 0: raise ValueError("Tutar, miktar ve fiyat negatif olamaz.")
            with connect() as c:
                if table in ("fatura", "beton_satis", "hareket"):
                    c.execute("BEGIN IMMEDIATE")
                    if record_is_posted(c, table, rid):
                        raise ValueError("Kayıt bu sırada faturaya bağlandı. Ekranı yenileyin.")
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
                if table=="fatura":
                    c.execute("UPDATE fatura SET ara_toplam=?,kdv_tutar=?,genel_toplam=?,tahsil_edilecek=? WHERE id=?", (data["ara_toplam"],data["kdv_tutar"],data["genel_toplam"],data["tahsil_edilecek"],rid))
                if table=="beton_satis":
                    c.execute("UPDATE beton_satis SET toplam=? WHERE id=?",(data["toplam"],rid))
                    # Cari borç satırını satışın değişmez ID'sine göre tekilleştirip senkronla.
                    sync_beton_sale_cari_movement(c, rid)
                    ref=f"Beton satış tahsilatı #{rid}"
                    paid=c.execute("SELECT id FROM hareket WHERE aciklama=? AND kategori='Tahsilat' LIMIT 1",(ref,)).fetchone()
                    if data["odeme_durumu"]=="Tahsil Edildi" and data["cari_id"] and not paid:
                        c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(data["tarih"],data["cari_id"],"Alacak","Tahsilat",ref,data["toplam"],"Nakit"))
                        c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(data["tarih"],data["cari_id"],"Gelir","Satış",ref,data["toplam"],"Nakit"))
                        c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",(data["tarih"],"Para Girişi","Beton Satışı",ref,data["toplam"],getattr(self,"current_user",{}).get("username","")))
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
            self._refresh_selected_module()
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
        self.banka_tree=self._tree(self.banka_tab,("id","tarih","banka","hesap","yon","karsi_taraf","aciklama","tutar"),("ID","Tarih","Banka","Hesap","İşlem","Gönderici / Alıcı","Açıklama","Tutar TL"),(45,100,130,100,105,200,175,110))
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
            self._refresh_selected_module()
        except Exception as e: messagebox.showerror("Stok hareketi kaydedilemedi",str(e),parent=self)

    def add_maas_odeme(self, personel_id=None, kind=None):
        with connect() as db:
            if personel_id is None:
                items = db.execute("SELECT id,ad FROM personel ORDER BY ad").fetchall()
            else:
                items = db.execute("SELECT id,ad FROM personel WHERE id=?", (personel_id,)).fetchall()
        if not items:
            return messagebox.showinfo("Personel yok", "Kayıtlı bir personel seçin.", parent=self)
        fields = [
            ("tarih", "İşlem tarihi", "entry", None),
            ("personel", "Personel", "combo", [f"{i} - {n}" for i, n in items]),
            ("islem_turu", "İşlem türü", "combo", [kind] if kind else list(("Maaş", "Nakit Avans", "Mesai", "Kesinti"))),
            ("donem", "Dönem (YYYY-AA)", "entry", None),
            ("tutar", "Tutar (TL)", "entry", None),
        ]
        if kind not in ("Mesai", "Kesinti"):
            fields.append(("yontem", "Ödeme yöntemi (maaş/avans)", "combo", ["Banka", "Nakit", "Diğer"]))
        fields.append(("aciklama", "Açıklama", "entry", None))
        data = self.prompt(f"Personel İşlemi - {kind}" if kind else "Personel İşlemi", fields)
        if not data:
            return
        try:
            day = data["tarih"] or date.today().isoformat()
            datetime.strptime(day, "%Y-%m-%d")
            pid = int(data["personel"].split(" - ", 1)[0])
            amount = parse_amount(data["tutar"])
            if amount <= 0:
                raise ValueError("Tutar sıfırdan büyük olmalı.")
            if data["donem"]:
                datetime.strptime(data["donem"] + "-01", "%Y-%m-%d")
            method = data.get("yontem", "") if data["islem_turu"] in ("Maaş", "Nakit Avans") else ""
            with connect() as db:
                db.execute("""INSERT INTO maas_odeme(tarih,personel_id,donem,tutar,yontem,aciklama,islem_turu)
                    VALUES(?,?,?,?,?,?,?)""", (day, pid, data["donem"], amount,
                    method, data["aciklama"], data["islem_turu"]))
            self._refresh_selected_module()
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Personel işlemi kaydedilemedi", str(error), parent=self)

    def add_banka_hareket(self):
        d=self.prompt("Banka İşlemi",[("tarih","Tarih (GG.AA.YYYY)","entry",None),("banka","Banka adı","entry",None),("hesap","Hesap / IBAN sonu","entry",None),("yon","İşlem","combo",["Para Girişi","Para Çıkışı"]),("karsi_taraf","Gönderici / Alıcı Ad Soyad veya Ünvan","entry",None),("aciklama","Açıklama","entry",None),("tutar","Tutar (TL)","entry",None)])
        if not d:return
        try:
            dt=d["tarih"] or date.today().isoformat(); datetime.strptime(dt,"%Y-%m-%d"); amount=parse_amount(d["tutar"])
            if not d["banka"].strip() or amount<=0:raise ValueError("Banka adı gerekli, tutar sıfırdan büyük olmalı.")
            with connect() as c:c.insert_shared("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar,karsi_taraf) VALUES(?,?,?,?,?,?,?)",(dt,d["banka"],d["hesap"],d["yon"],d["aciklama"],amount,d["karsi_taraf"]))
            self._refresh_selected_module()
        except Exception as e:messagebox.showerror("Banka işlemi kaydedilemedi",str(e),parent=self)

    def add_cari(self):
        d = self.prompt("Yeni Cari", [("tur","Tür","combo",["Müşteri","Tedarikçi","Diğer"]),
            ("unvan","Ünvan / Ad Soyad","entry",None),("telefon","Telefon","entry",None),
            ("vergi_no","Vergi / T.C. No","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        if not d["unvan"]: return messagebox.showwarning("Eksik bilgi","Ünvan zorunludur.")
        with connect() as c:
            c.insert_shared("INSERT INTO cari(tur,unvan,telefon,vergi_no,notlar,created_at) VALUES(?,?,?,?,?,?)",
                (d["tur"],d["unvan"],d["telefon"],d["vergi_no"],d["notlar"],datetime.now().isoformat(timespec="seconds")))
        self._refresh_selected_module()

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
        if getattr(self, "current_user", {}).get("role") != "Yönetici":
            return messagebox.showwarning("Yetki gerekli", "Üyelikleri yalnızca yönetici yönetebilir.", parent=self)
        if self.membership_client is not None:
            from hybrid_desktop_membership import manage_pending
            from pathlib import Path
            from hybrid_enrollment import read_baseline

            def publish():
                return self.membership_client.publish_baseline(
                    self.membership_session, read_baseline(Path(DB_PATH)))

            with connect() as db:
                publish_available = db.hybrid_slot == 1
            return manage_pending(self, self.membership_client, self.membership_session,
                                  publish if publish_available else None)
        win=tk.Toplevel(self); win.title("Kullanıcı Yönetimi"); win.transient(self); win.grab_set(); win.geometry("520x390")
        frm=ttk.Frame(win,padding=16); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Kullanıcı Hesapları",font=("Segoe UI",15,"bold")).pack(anchor="w")
        tree=ttk.Treeview(frm,columns=("id","user","role","active"),show="headings",height=7)
        for col,label,width in [("id","ID",45),("user","Kullanıcı adı",190),("role","Rol",120),("active","Durum",90)]: tree.heading(col,text=label); tree.column(col,width=width)
        tree.pack(fill="both",expand=True,pady=8)
        def reload_users():
            tree.delete(*tree.get_children())
            with connect() as c: rows=c.execute("SELECT id,kullanici_adi,rol,aktif FROM kullanici ORDER BY id").fetchall()
            for row in rows: tree.insert("","end",values=(row[0],row[1],row[2],{1: "Aktif", -1: "Onay bekliyor"}.get(row[3], "Pasif")))
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
        ttk.Button(row,text="Onayla / Aktif-Pasif Yap",command=toggle_user).pack(side="left",padx=8)
        reload_users()

    def add_santiye(self):
        d=self.prompt("Yeni Şantiye",[("ad","Şantiye adı","entry",None),("adres","Adres","entry",None),("yetkili","Yetkili kişi","entry",None),("telefon","Telefon","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        if not d["ad"]: return messagebox.showwarning("Eksik bilgi","Şantiye adı zorunludur.")
        with connect() as c: c.insert_shared("INSERT INTO santiye(ad,adres,yetkili,telefon,notlar) VALUES(?,?,?,?,?)",(d["ad"],d["adres"],d["yetkili"],d["telefon"],d["notlar"]))
        self._refresh_selected_module()

    def selected_invoice_document(self, tree, *, edit=False, review=False):
        selection = tree.selection()
        if not selection:
            return messagebox.showinfo("Fatura seçin", "Önce bir satış faturası taslağı seçin.", parent=tree)
        values = tree.item(selection[0], "values")
        if tree["columns"][1] == "tur" and values[1] != "Satış":
            return messagebox.showinfo("Satış taslağı", "Bu işlem için bir satış faturası taslağı seçin.", parent=tree)
        invoice_id = int(values[0])
        if review:
            return self._review_invoice_customer(invoice_id, tree)
        if not edit:
            try:
                return self._preview_then_save_pdf("Fatura Taslağı PDF", cari_id=invoice_id)
            except (OSError, ValueError, RuntimeError, ImportError) as exc:
                return messagebox.showerror("Fatura PDF açılamadı", str(exc), parent=tree)

        with connect() as db:
            invoice = db.execute("""SELECT f.cari_id, f.belge_bilgileri, c.unvan FROM fatura f
                LEFT JOIN cari c ON c.id=f.cari_id WHERE f.id=?""", (invoice_id,)).fetchone()
        if not invoice:
            return messagebox.showinfo("Fatura bulunamadı", "Seçili taslak artık kayıtlı değil.", parent=tree)
        customer_id, stored, customer = invoice
        fields = [
            ("satici_unvan", "Satıcı ticaret unvanı", "entry", None),
            ("satici_adres", "Satıcı iş adresi", "entry", None),
            ("satici_daire", "Satıcı vergi dairesi", "entry", None),
            ("satici_vkn", "Satıcı VKN / TCKN", "entry", None),
            ("alici_adres", "Alıcı fatura adresi", "entry", None),
            ("alici_daire", "Alıcı vergi dairesi (varsa)", "entry", None),
            ("teslim", "Teslim tarihi (GG.AA.YYYY)", "entry", None),
            ("irsaliye", "İrsaliye numarası", "entry", None),
            ("senaryo", "Taslak senaryosu", "combo", list(INVOICE_DRAFT_PROFILES)),
            ("tevkifat_kodu", "Tevkifat kodu ve işlem adı (varsa)", "entry", None),
        ]
        initial = invoice_document_details(stored, customer_id)
        if initial.get("tevkifat_islem_kodu"):
            fields = [field for field in fields if field[0] != "tevkifat_kodu"]
        initial.setdefault("satici_unvan", "HOYTAŞ Hazır Beton")
        values = self.prompt(f"Fatura Belge Bilgileri | {customer or 'Cari seçilmemiş'}", fields, initial)
        if not values:
            return
        try:
            if values["teslim"]:
                values["teslim"] = normalize_fatura_tarihi(values["teslim"])
            tax_id = values["satici_vkn"]
            if tax_id and not re.fullmatch(r"[0-9]{10,11}", tax_id):
                raise ValueError("Satıcı VKN 10, TCKN 11 rakam olmalıdır.")
        except ValueError as exc:
            return messagebox.showerror("Belge bilgileri", str(exc), parent=tree)
        values = {**initial, **values, "cari_id": customer_id}
        with connect() as db:
            result = db.execute("""UPDATE fatura SET belge_bilgileri=?
                WHERE id=? AND cari_id IS ? AND belge_bilgileri=?""",
                (json.dumps(values, ensure_ascii=False), invoice_id, customer_id, stored))
        if not result.rowcount:
            return messagebox.showwarning("Taslak değişti", "Faturayı yenileyin ve belge bilgilerini tekrar açın.", parent=tree)
        messagebox.showinfo("Belge bilgileri", "Fatura belge bilgileri kaydedildi.", parent=tree)

    def _review_invoice_customer(self, invoice_id, parent):
        with connect() as db:
            row = db.execute("""SELECT f.cari_id, f.belge_no, f.belge_bilgileri, c.unvan,
                c.vergi_no, c.telefon FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id
                WHERE f.id=?""", (invoice_id,)).fetchone()
        if not row:
            return messagebox.showinfo("Fatura bulunamadı", "Taslağı yenileyip tekrar seçin.", parent=parent)
        customer_id, number, stored, name, tax_id, phone = row
        details = invoice_document_details(stored, customer_id)
        fields = (
            ("Cari unvanı", name),
            ("VKN / TCKN", tax_id),
            ("Fatura adresi", details.get("alici_adres")),
            ("Vergi dairesi", details.get("alici_daire")),
            ("Telefon", phone),
            ("Taslak senaryosu", details.get("senaryo") or "Seçilmedi"),
        )
        issues = []
        if not name:
            issues.append("Cari unvanı eksik.")
        if not re.fullmatch(r"[0-9]{10,11}", tax_id or ""):
            issues.append("VKN / TCKN biçimini kontrol edin.")
        if not details.get("alici_adres"):
            issues.append("Fatura adresi eksik.")
        if not details.get("senaryo"):
            issues.append("Belge senaryosu seçilmemiş.")
        window = tk.Toplevel(self)
        window.title("Fatura Cari Kontrolü")
        window.transient(self)
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Cari kontrolü · {number or 'Fatura taslağı'}",
                  font=("Segoe UI", 14, "bold")).grid(row=0, columnspan=2, sticky="w", pady=(0, 12))
        for index, (label, value) in enumerate(fields, start=1):
            ttk.Label(body, text=label).grid(row=index, column=0, sticky="nw", padx=(0, 18), pady=4)
            ttk.Label(body, text=value or "—", wraplength=410).grid(row=index, column=1, sticky="w", pady=4)
        ttk.Label(body, text=" ".join(issues) if issues else "Cari alanları yerel kontrolden geçti.",
                  foreground="#815224" if issues else "#315B46", wraplength=550).grid(
                      row=7, columnspan=2, sticky="w", pady=(12, 4))
        ttk.Label(body, text="GİB mükellef durumu sorgulanmadı.").grid(row=8, columnspan=2, sticky="w")
        ttk.Button(body, text="Kapat", command=window.destroy).grid(row=9, column=1, sticky="e", pady=(14, 0))
        return window

    def match_concrete_invoice(self, cari_context):
        customer, name = cari_context
        with connect() as db:
            available = db.execute("""SELECT b.id,b.tarih,COALESCE(s.ad,'—'),b.miktar_m3,b.toplam
                FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id
                WHERE b.cari_id=? AND b.fatura_id IS NULL ORDER BY b.tarih,b.id""", (customer,)).fetchall()
            invoices = db.execute("""SELECT id,belge_no,tahsil_edilecek,belge_bilgileri FROM fatura
                WHERE cari_id=? AND durum<>'İptal edildi' ORDER BY tarih,id""", (customer,)).fetchall()
        if not available:
            return messagebox.showinfo("Satış bulunamadı", "Bu caride eşleştirilebilir beton satışı yok.", parent=self)
        window = tk.Toplevel(self)
        window.title("Beton Satışını Faturaya Bağla")
        window.geometry("980x480")
        window.transient(self)
        body = ttk.Frame(window, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Cari: {name}", font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(0, 8))
        options = ttk.LabelFrame(body, text="Fatura", padding=10)
        options.pack(side="right", fill="y", padx=(10, 0))
        mode = tk.StringVar(window, "new")
        for label, option in (("Yeni Fatura Taslağı", "new"), ("Mevcut Faturayı Seç", "existing")):
            ttk.Radiobutton(options, text=label, value=option, variable=mode).pack(anchor="w", pady=4)
        invoice_choices = {f"{number} · {amount:,.2f} TL · #{fid}": fid
                           for fid, number, amount, details in invoices
                           if not stored_details(details).get("cari_isleme")}
        picker = ttk.Combobox(options, values=list(invoice_choices), width=30, state="readonly")
        picker.pack(fill="x", pady=6)
        ttk.Label(options, text="Normal ve tevkifatlı satış faturaları.\nİşlem kodu ve dayanak yeni taslak formunda.",
                  wraplength=240).pack(anchor="w", pady=6)
        ttk.Label(body, text="Beton Satışları · Birden fazla satır için Ctrl tuşunu kullanın").pack(anchor="w")
        table = self._tree(body, ("id", "tarih", "santiye", "m3", "borc"),
                           ("ID", "Tarih", "Şantiye", "Miktar (m³)", "Satış Borcu"),
                           (50, 90, 175, 100, 100), show_id=False, print_button=False)
        table.unbind("<Double-1>")
        table.unbind("<Button-3>")
        table.unbind("<Control-Button-1>")
        self.fill_tree(table, [(sid, day, site, format_m3(qty), f"{value:,.2f}")
                               for sid, day, site, qty, value in available])
        selection = None

        def choose():
            nonlocal selection
            sales = tuple(int(table.set(item, "id")) for item in table.selection())
            invoice = invoice_choices.get(picker.get())
            if not sales or (mode.get() == "existing" and invoice is None):
                return messagebox.showinfo("Seçim gerekli", "Satışları ve gerekiyorsa faturayı seçin.", parent=window)
            selection = sales, invoice if mode.get() == "existing" else None
            window.destroy()

        ttk.Button(options, text="Devam Et", command=choose).pack(fill="x", pady=(12, 4))
        ttk.Button(options, text="Vazgeç", command=window.destroy).pack(fill="x")
        window.bind("<Escape>", lambda _event: window.destroy())
        window.grab_set()
        self.wait_window(window)
        if selection is None:
            return
        sale_ids, invoice_id = selection
        if invoice_id is None:
            return self.add_fatura(cari_context, sale_ids=sale_ids, post=True)
        try:
            with connect() as db:
                sales = selected_sales(db, customer, sale_ids, index_beton_movements(db)[0])
                invoice = invoice_quote(db, customer, invoice_id)
            if abs(invoice[5] - sum(row[2] for row in sales)) > 0.000001:
                raise ValueError("Faturanın miktarı seçili beton satışlarının toplamına eşit olmalıdır.")
            if not self._confirm_invoice_posting(cari_context, sales, invoice[2], invoice[3]):
                return
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                if (invoice_quote(db, customer, invoice_id) != invoice or
                        selected_sales(db, customer, sale_ids, index_beton_movements(db)[0]) != sales):
                    raise ValueError("Satış veya fatura değişti. İşlemi yeniden açın.")
                post_invoice(db, customer, invoice_id, invoice, sales)
            messagebox.showinfo("Fatura eşleştirildi", "Fatura bağlantısı ve cari farkı kaydedildi.", parent=self)
        except (ValueError, sqlite3.Error) as exc:
            messagebox.showerror("Fatura eşleştirilemedi", str(exc), parent=self)

    def _confirm_invoice_posting(self, customer, sales, payable, paid=0, operation="", taxes=None):
        sale_debt = sum(cents(row[3]) for row in sales)
        delta = cents(payable) - sale_debt
        message = (f"{customer[1]}\n\n"
                   f"Mevcut satış borcu: {invoice_money(sale_debt / 100)} TL\n"
                   f"Fatura tahsil edilecek: {invoice_money(payable)} TL\n"
                   f"Faturaya kaydedilmiş tahsilat: {invoice_money(paid)} TL\n"
                   f"Cariye işlenecek fark: {'+' if delta > 0 else ''}{invoice_money(delta / 100)} TL")
        if taxes is not None:
            base, vat, withheld = taxes
            message += (f"\n\nMatrah: {invoice_money(base)} TL · KDV: {invoice_money(vat)} TL"
                        f"\nTevkifat: {invoice_money(withheld)} TL")
        if operation:
            message += f"\n{operation}"
        message += "\n\nBu fatura için ayrıca borç kaydı girdiyseniz önce cari hareketleri kontrol edin."
        return messagebox.askyesno("Fatura ve Cari Farkı Önizleme", message, parent=self)

    def show_invoice_tax_options(self):
        import webbrowser
        window = tk.Toplevel(self)
        window.title("Fatura Türleri ve GİB İşlem Kodları")
        window.geometry("860x560")
        window.transient(self)
        body = ttk.Frame(window, padding=12)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Normal (SATIŞ) · Tevkifatlı (TEVKIFAT)",
                  font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(body, text=("4/10 hesaplanan KDV'nin kesinti oranıdır. Tebliğ I/C-2.1.3.1/b, belirlenmiş alıcılar bölümüdür. "
                              "İşlem ve alıcı koşullarını dayanak metinden kontrol edin."),
                  wraplength=810).pack(anchor="w", pady=6)
        table = self._tree(body, ("kod", "islem", "oran", "dayanak"),
                           ("Kod", "İşlem", "Oran", "Tebliğ I/C"), (60, 480, 65, 120), print_button=False)
        self.fill_tree(table, [(code, label.split(" · ")[1], f"{rate}/10", basis)
                               for label, (code, rate, basis) in WITHHOLDING_LABELS.items()])
        ttk.Button(body, text="GİB Tebliğini Aç", command=lambda: webbrowser.open(GIB_TEBLIG)).pack(anchor="e", pady=6)

    def add_fatura(self, cari_context=None, *, sale_ids=(), post=False):
        """Yerel taslak: m³, KDV ve seçilen tevkifat oranından net tahsilatı hesapla."""
        with connect() as c:
            cariler = c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        if not cariler:
            return messagebox.showinfo("Cari gerekli", "Önce cari hesap ekleyin.", parent=self)
        names = [f"{i} - {n}" for i,n in cariler]
        sales = ()
        if sale_ids:
            try:
                with connect() as db:
                    sales = selected_sales(db, cari_context[0], sale_ids, index_beton_movements(db)[0])
            except (ValueError, sqlite3.Error) as exc:
                return messagebox.showerror("Satış seçimi", str(exc), parent=self)
        initial = {"miktar": str(sum(Decimal(str(row[2])) for row in sales))} if sales else None
        d = self._prompt_for_cari(cari_context, "Yeni Satış Faturası Taslağı", [
            ("tarih", "Fatura tarihi (GG.AA.YYYY)", "entry", None),
            ("cari", "Faturanın kesileceği cari", "combo_search", names),
            ("belge", "Fatura / belge numarası", "entry", None),
            ("aciklama", "Mal/hizmet, şantiye ve teslim açıklaması", "entry", None),
            ("miktar", "Teslim edilen beton (m³; ör. 30,5)", "entry", None),
            ("fiyat", "KDV hariç birim fiyat (TL/m³, ücretler dahilse belirtin)", "entry", None),
            ("kdv", "KDV oranı (%)", "combo", ["20","10","1","0"]),
            ("tur", "Fatura türü", "combo", ["Normal", "Tevkifatlı"]),
            ("kod", "GİB işlem kodu · işlem adı · oran", "combo_edit", ["Seçiniz", *WITHHOLDING_LABELS]),
            ("tevkifat", "Kod seçilmediyse tevkifat oranı (eski taslak)", "combo_edit", ["0","4/10","5/10","7/10","9/10"])
        ], initial)
        if not d:
            return
        try:
            tarih = d["tarih"] or date.today().isoformat()
            tarih = normalize_fatura_tarihi(tarih)
            miktar = parse_amount(d["miktar"])
            fiyat = parse_amount(d["fiyat"])
            kdv_oran = parse_amount(d["kdv"])
            if miktar <= 0 or fiyat < 0 or not 0 <= kdv_oran <= 100:
                raise ValueError("Miktar, fiyat veya KDV oranı geçersiz.")
            tur = d["tur"]
            oran_metin = d["tevkifat"].strip()
            operation = WITHHOLDING_LABELS.get(d.get("kod", ""))
            if d.get("kod") not in ("", "Seçiniz", None) and operation is None:
                raise ValueError("GİB işlem kodunu açıklamalı listeden seçin.")
            if tur == "Normal" and operation:
                raise ValueError("GİB tevkifat kodu için fatura türünü Tevkifatlı seçin.")
            if tur == "Tevkifatlı" and operation:
                tevkifat_oran = operation[1] / 10
            elif tur == "Tevkifatlı":
                if post:
                    raise ValueError("Cariye işlenecek tevkifatlı fatura için GİB işlem kodunu listeden seçin.")
                if "/" not in oran_metin:
                    raise ValueError("Tevkifat oranını 4/10 gibi girin.")
                pay, payda = [int(x.strip()) for x in oran_metin.split("/", 1)]
                if payda != 10 or not 0 < pay < payda:
                    raise ValueError("Tevkifat oranı 1/10 ile 9/10 arasında olmalıdır.")
                tevkifat_oran = pay / payda
            else:
                tevkifat_oran = 0.0
            if d["cari"] not in names and cari_context is None:
                raise ValueError("Listeden kayıtlı bir cari hesap seçin.")
            cari_id = cari_context[0] if cari_context is not None else int(d["cari"].split(" - ", 1)[0])
            base_cents = cents(Decimal(str(miktar)) * Decimal(str(fiyat)))
            vat_cents = cents(Decimal(base_cents) * Decimal(str(kdv_oran)) / 10000)
            withholding_cents = cents(Decimal(vat_cents) * Decimal(str(tevkifat_oran)) / 100)
            ara, kdv_tutar, toplam, tevkifat_tutar, net = (
                value / 100 for value in (base_cents, vat_cents, base_cents + vat_cents,
                                         withholding_cents, base_cents + vat_cents - withholding_cents))
            if sales and abs(miktar - sum(row[2] for row in sales)) > 0.000001:
                raise ValueError("Fatura miktarı seçili beton satışlarının toplamına eşit olmalıdır.")
            if not d["belge"].strip():
                raise ValueError("Belge numarası boş bırakılamaz.")
            operation_label = f"{d['kod']} · KDVGUT I/C-{operation[2]}" if operation else ""
            details = ({"cari_id": cari_id, "tevkifat_kodu": operation_label,
                        "tevkifat_islem_kodu": operation[0]} if operation else {})
            if post:
                if net <= 0:
                    raise ValueError("Cariye işlenecek faturanın tahsil edilecek tutarı pozitif olmalıdır.")
                if not self._confirm_invoice_posting(cari_context, sales, net, operation=operation_label,
                                                     taxes=(ara, kdv_tutar, tevkifat_tutar)):
                    return
            with connect() as c:
                c.execute("BEGIN IMMEDIATE")
                if sales and selected_sales(c, cari_id, sale_ids, index_beton_movements(c)[0]) != sales:
                    raise ValueError("Satış değişti. İşlemi yeniden açın.")
                if not c.execute("SELECT 1 FROM cari WHERE id=?", (cari_id,)).fetchone():
                    raise ValueError("Seçilen cari hesap bulunamadı.")
                if c.execute("SELECT 1 FROM fatura WHERE cari_id=? AND belge_no=?",
                             (cari_id, d["belge"].strip())).fetchone():
                    raise ValueError("Bu cariye ait aynı belge numarası zaten kayıtlı.")
                cur = c.insert_shared("""INSERT INTO fatura
                    (tarih,cari_id,belge_no,aciklama,miktar_m3,birim_fiyat,ara_toplam,kdv_oran,kdv_tutar,
                     genel_toplam,durum,fatura_turu,tevkifat_oran,tevkifat_tutar,tahsil_edilecek,odenen,belge_bilgileri)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0,?)""",
                    (tarih,cari_id,d["belge"].strip(),d["aciklama"],miktar,fiyat,ara,kdv_oran,
                     kdv_tutar,toplam,"Taslak",tur,tevkifat_oran,tevkifat_tutar,net, json.dumps(details, ensure_ascii=False)))
                if post:
                    posted = invoice_quote(c, cari_id, cur.lastrowid)
                    post_invoice(c, cari_id, cur.lastrowid, posted, sales)
            self._refresh_selected_module()
            messagebox.showinfo("Fatura taslağı kaydedildi",
                f"Ara toplam: {ara:,.2f} TL\nKDV: {kdv_tutar:,.2f} TL\n"
                f"Genel toplam: {toplam:,.2f} TL\nTevkifat: {tevkifat_tutar:,.2f} TL\n"
                f"Tahsil edilecek: {net:,.2f} TL\n\nBu kayıt resmî e-Fatura gönderimi değildir.",
                parent=self)
        except (ValueError, TypeError, sqlite3.Error) as exc:
            messagebox.showerror("Fatura kaydedilemedi", str(exc), parent=self)

    def tahsilat_fatura(self):
        selected = self.fatura_tree.selection()
        if not selected:
            return messagebox.showinfo("Fatura seçin", "Önce bir fatura seçin.", parent=self)
        invoice_id = int(self.fatura_tree.item(selected[0], "values")[0])
        with connect() as db:
            customer = db.execute("""SELECT c.id,c.unvan FROM fatura f JOIN cari c ON c.id=f.cari_id
                WHERE f.id=? AND f.durum<>'İptal edildi' AND f.tahsil_edilecek-f.odenen>0.005""", (invoice_id,)).fetchone()
        if not customer:
            return messagebox.showinfo("Tahsilat", "Bu faturanın açık tahsilatı bulunmuyor.", parent=self)
        self.cari_tahsilat(customer, invoice_id=invoice_id)

    def cari_tahsilat(self, cari_context=None, *, invoice_id=None):
        """Seçili müşteriden alınan kısmi/tam ödemeyi cari ve para hareketlerine işler."""
        if cari_context is None:
            sel = self.cari_tree.selection()
            if not sel:
                return messagebox.showinfo("Cari seçin", "Önce tahsilat alınacak müşteriyi seçin.", parent=self)
            vals = self.cari_tree.item(sel[0], "values")
            cari_context = int(vals[0]), str(vals[2])
        cari_id, cari_unvan = cari_context
        with connect() as c:
            open_invoices = c.execute("""SELECT id, COALESCE(belge_no,''), tahsil_edilecek-odenen
                FROM fatura WHERE cari_id=? AND durum<>'İptal edildi' AND tahsil_edilecek-odenen>0.005
                ORDER BY tarih,id""", (cari_id,)).fetchall()
        invoice_options = [f"{fid} - {belge} (Kalan {kalan:.2f} TL)"
                           for fid, belge, kalan in open_invoices if invoice_id is None or fid == invoice_id]
        if invoice_id is None:
            invoice_options.insert(0, "Faturaya bağlama")
        elif not invoice_options:
            return messagebox.showinfo("Tahsilat", "Faturanın açık tahsilatı kalmadı.", parent=self)
        d = self.prompt(f"Cari Tahsilat - {cari_unvan}", [
            ("fatura", "Tahsilatın bağlanacağı fatura (isteğe bağlı)", "combo", invoice_options),
            ("tarih", "Tahsilat tarihi (GG.AA.YYYY)", "entry", None),
            ("tutar", "Alınan tutar (TL)", "entry", None),
            ("yontem", "Ödeme şekli", "combo", ["Nakit", "Banka Havalesi / EFT", "Çek"]),
            ("banka", "Banka adı (banka/çek için)", "entry", None),
            ("cek_no", "Çek numarası (çek ise)", "entry", None),
            ("vade", "Çek vadesi GG.AA.YYYY (çek ise)", "entry", None),
            ("aciklama", "Açıklama", "entry", None)])
        if not d: return
        try:
            tarih = normalize_fatura_tarihi(d["tarih"])
            amount = Decimal(str(parse_amount(d["tutar"]))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if not amount.is_finite() or amount <= 0:
                raise ValueError("Tutar sıfırdan büyük olmalı.")
            tutar = float(amount)
            if d["yontem"] == "Çek":
                vade = d["vade"] or tarih
                vade = normalize_fatura_tarihi(vade)
        except (ValueError, InvalidOperation):
            return messagebox.showerror("Hatalı bilgi", "Tarih/vade GG.AA.YYYY biçiminde, tutar sıfırdan büyük olmalıdır.", parent=self)
        aciklama = (d["aciklama"] or f"{cari_unvan} cari tahsilatı")
        selected_id = d.get("fatura", "").split(" - ", 1)[0]
        fatura_id = invoice_id if invoice_id is not None else (int(selected_id) if selected_id.isdigit() else None)
        with connect() as c:
            if fatura_id is not None:
                updated = c.execute("""UPDATE fatura SET odenen=odenen+?
                    WHERE id=? AND cari_id=? AND durum<>'İptal edildi'
                      AND ? <= ROUND(tahsil_edilecek-odenen, 2)+0.005
                      AND ? <= COALESCE((SELECT ROUND(bakiye,2) FROM cari_bakiye WHERE cari_id=?),0)+0.005""",
                    (tutar, fatura_id, cari_id, tutar, tutar, cari_id))
                if updated.rowcount != 1:
                    return messagebox.showerror("Fatura tutarı", "Tutar, açık fatura tutarını veya cari borç bakiyesini aşıyor ya da fatura tahsilata kapalı.", parent=self)
                c.insert_shared("INSERT INTO fatura_tahsilat(fatura_id,tarih,tutar,yontem,aciklama) VALUES(?,?,?,?,?)",
                          (fatura_id,tarih,tutar,d["yontem"],aciklama))
            # Tahsilat, cari borcunu yalnızca alınan tutar kadar azaltır.
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                      (tarih, cari_id, "Alacak", "Tahsilat", aciklama, tutar, d["yontem"]))
            # Tahsilat gelir hareketi olarak raporlanır; satış borcu ayrıca korunur.
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                      (tarih, cari_id, "Gelir", "Tahsilat", aciklama, tutar, d["yontem"]))
            if d["yontem"] == "Nakit":
                c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                          (tarih, "Para Girişi", "Cari Tahsilat", f"{cari_unvan} - {aciklama}", tutar, getattr(self, "current_user", {}).get("username", "")))
            elif d["yontem"] == "Banka Havalesi / EFT":
                c.insert_shared("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                          (tarih, d["banka"] or "Banka", "", "Para Girişi", f"{cari_unvan} - {aciklama}", tutar))
            else:
                cek_no = d["cek_no"] or f"TAH-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                c.insert_shared("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",
                          ("Alınan", d["banka"] or "", "", cari_id, cek_no, cari_unvan, vade, tutar, "Bekliyor", aciklama))
        self._refresh_selected_module()
        messagebox.showinfo("Tahsilat kaydedildi", f"{tutar:,.2f} TL tahsilat kaydedildi. Cari bakiye alınan tutar kadar düşürüldü.", parent=self)

    def toggle_cari_archive(self):
        selected = self.cari_tree.selection()
        if not selected:
            return messagebox.showinfo("Cari seçin", "Önce bir cari hesap seçin.", parent=self)
        values = self.cari_tree.item(selected[0], "values")
        cari_id = int(values[0])
        archived = int(self.cari_arsiv_goster.get())
        action = "arşivden çıkarılsın" if archived else "arşive alınsın"
        if not messagebox.askyesno("Cari arşivi", f"{values[2]} {action} mı? Cari geçmişi korunur.", parent=self):
            return
        with connect() as db:
            db.execute("UPDATE cari SET arsivli=? WHERE id=? AND arsivli=?", (1-archived, cari_id, archived))
        self.refresh_cari()

    def refresh_cari(self):
        if not hasattr(self,"cari_tree"): return
        with connect() as c:
            raw=c.execute("""SELECT c.id,c.tur,c.unvan,c.telefon,c.vergi_no,
                b.toplam_borc,b.toplam_alacak,b.bakiye
                FROM cari c JOIN cari_bakiye b ON b.cari_id=c.id
                WHERE c.arsivli=? ORDER BY c.unvan COLLATE NOCASE""",
                (int(self.cari_arsiv_goster.get()),)).fetchall()
        rows=[]; row_tags={}
        for r in raw:
            borc, alacak, bakiye = safe_float(r[5] or 0), safe_float(r[6] or 0), safe_float(r[7] or 0)
            rows.append((r[0],r[1],r[2],r[3],r[4],f"{borc:.2f}",f"{alacak:.2f}",f"{bakiye:.2f}"))
            row_tags[str(r[0])] = ("debt_partial" if alacak > 0.005 else "debt_unpaid") if bakiye > 0.005 else "debt_closed"
        self.cari_tree._row_tags=row_tags
        self.fill_tree(self.cari_tree,rows)

    def add_hareket(self, cari_context=None):
        with connect() as c: cariler = c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        names = ["(Cari seçilmedi)"] + [f"{i} - {n}" for i,n in cariler]
        d = self._prompt_for_cari(cari_context, "Yeni Gelir / Gider", [("tarih","Tarih (GG.AA.YYYY)","entry",None),
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
            c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                (tarih,cari_id,d["tur"],d["kategori"],d["aciklama"],tutar,d["odeme"]))
        self._refresh_selected_module()

    def add_cek(self, cari_context=None):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        names=["(Cari seçilmedi)"]+[f"{i} - {n}" for i,n in cariler]
        banks=["010 - T.C. Ziraat Bankası","012 - Türkiye Halk Bankası","015 - Türkiye Vakıflar Bankası","032 - Türk Ekonomi Bankası","046 - Akbank","062 - Türkiye Garanti Bankası","064 - Türkiye İş Bankası","067 - Yapı ve Kredi Bankası","111 - QNB Bank","134 - DenizBank","203 - Albaraka Türk","206 - Türkiye Finans","208 - Kuveyt Türk","210 - Vakıf Katılım","211 - Türkiye Emlak Katılım","214 - Şekerbank","223 - Alternatif Bank","999 - Diğer / listede yok"]
        d=self._prompt_for_cari(cari_context, "Yeni Çek",[("yon","Yön","combo",["Alınan","Verilen"]),("banka","Banka adı ve kodu","combo",banks), ("cari","İlgili cari hesap","combo",names),
            ("cek_no","Çek No","entry",None),("taraf","Keşideci / Lehtar","entry",None),("vade","Vade (GG.AA.YYYY)","entry",None),("tutar","Tutar (TL)","entry",None),
            ("durum","Durum","combo",list(CHEQUE_STATES)),("notlar","Notlar","entry",None)])
        if not d:return
        try:
            datetime.strptime(d["vade"],"%Y-%m-%d"); tutar=parse_amount(d["tutar"]);
            if tutar<=0: raise ValueError
            cari_id=int(d["cari"].split(" - ",1)[0]) if " - " in d["cari"] else None
            banka_kodu,banka=(d["banka"].split(" - ",1)+[""])[:2] if " - " in d["banka"] else ("",d["banka"])
        except ValueError:return messagebox.showerror("Hatalı bilgi","Vade GG.AA.YYYY biçiminde ve tutar sıfırdan büyük olmalıdır.")
        with connect() as c:
            c.insert_shared("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",(d["yon"],banka,banka_kodu,cari_id,d["cek_no"],d["taraf"],d["vade"],tutar,d["durum"],d["notlar"]))
            if cari_id:
                # Alınan çek müşteri alacağını tahsilat gibi azaltır; verilen çek
                # ise cari borcu/alacak yönünde ters etki oluşturur.
                move_type = "Alacak" if d["yon"] == "Alınan" else "Borç"
                c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",(datetime.now().strftime("%Y-%m-%d"),cari_id,move_type,"Çek",f"{d['yon']} çek #{d['cek_no']} - {banka}",tutar,"Çek"))
        self._refresh_selected_module()

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
        self._refresh_selected_module()

    def edit_personel_payment(self, personel_id):
        with connect() as db:
            row = db.execute("""SELECT p.ad, COALESCE(o.iban,'') FROM personel p
                LEFT JOIN personel_odeme_bilgisi o ON o.personel_id=p.id WHERE p.id=?""",
                (personel_id,)).fetchone()
        if row is None:
            return
        data = self.prompt("Personel Ödeme Bilgileri",
            [("ad", "Ad Soyad", "entry", None), ("iban", "IBAN", "entry", None)],
            {"ad": row[0], "iban": row[1]})
        if not data:
            return
        try:
            if not data["ad"]:
                raise ValueError("Ad soyad zorunludur.")
            iban = normalize_personel_iban(data["iban"])
            with connect() as db:
                if data["ad"] != row[0]:
                    db.execute("UPDATE personel SET ad=? WHERE id=?", (data["ad"], personel_id))
                save_personel_iban(db, personel_id, iban)
            self._refresh_selected_module()
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Ödeme bilgileri kaydedilemedi", str(error), parent=self)

    def add_personel(self):
        d = self.prompt("Personel Ekle", [("ad","Ad Soyad","entry",None),("gorev","Görev","entry",None),
            ("telefon","Telefon","entry",None),("iban","IBAN","entry",None),("maas","Aylık Maaş (TL)","entry",None),("notlar","Notlar","entry",None)])
        if not d: return
        try:
            maas=parse_amount(d["maas"] or "0")
            if not d["ad"] or maas < 0:
                raise ValueError("Ad soyad gerekli; maaş negatif olamaz.")
            iban = normalize_personel_iban(d["iban"])
        except ValueError as error:
            return messagebox.showerror("Hatalı bilgi", str(error), parent=self)
        with connect() as c:
            personel = c.execute("INSERT INTO personel(ad,gorev,telefon,maas,notlar) VALUES(?,?,?,?,?)",
                      (d["ad"],d["gorev"],d["telefon"],maas,d["notlar"])).lastrowid
            save_personel_iban(c, personel, iban)
        self._refresh_selected_module()

    def add_beton(self, cari_context=None):
        with connect() as c: cariler = c.execute("SELECT id,unvan FROM cari WHERE tur='Müşteri' OR tur='Diğer' ORDER BY unvan").fetchall()
        names = ["(Cari seçilmedi)"] + [f"{i} - {n}" for i,n in cariler]
        with connect() as c: sites=c.execute("SELECT id,ad FROM santiye ORDER BY ad").fetchall()
        site_names=[f"{i} - {n}" for i,n in sites]
        site_names += [n for n in ["Adıyaman", "Besni", "Tut", "Gölbaşı", "Erkenek", "Erkenek Köyü"] if n.casefold() not in [x.casefold() for x in site_names]]
        d = self._prompt_for_cari(cari_context, "Yeni Beton Satışı", [("tarih","Tarih (GG.AA.YYYY)","entry",None),
            ("cari","Müşteri ara (ünvan yazın) / seçin","combo_search",names),("santiye","Gideceği şantiye (elle yazabilirsiniz)","combo_edit",site_names),("beton_cinsi","Beton Cinsi","combo",["C16","C20","C25","C30","C35","C40","Şap"]),("m3","Miktar (m³; ör. 30,5)","entry",None),
            ("fiyat","Birim Fiyat (TL/m³)","entry",None),
            ("pompa_ucreti","Pompa Ücreti (TL, yoksa 0)","entry",None),
            ("mikser_nakliye_ucreti","Mikser Nakliye Ücreti (TL, isteğe bağlı)","entry",None),
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
            miktar,fiyat,pompa_ucreti,nakliye,toplam=concrete_sale_totals(
                d["m3"],d.get("fiyat"),d.get("pompa_ucreti"),d.get("mikser_nakliye_ucreti"))
            cari_value = (d.get("cari") or "").strip()
            with connect() as lookup_db:
                cari_id = resolve_customer_id(lookup_db, cari_value)
            santiye_id=self._resolve_site(d["santiye"])
            if d["durum"] == "Tahsil Edildi" and d.get("odeme_yontemi") == "Çek":
                datetime.strptime(d.get("vade") or tarih,"%Y-%m-%d")
        except ValueError:
            return messagebox.showerror("Hatalı bilgi","Tarih GG.AA.YYYY olmalı; metreküp sıfırdan büyük olmalı. Fiyatı boş bırakabilirsiniz (0 TL kaydedilir); fiyat, pompa ve mikser nakliye ücreti negatif olamaz. 30,5 gibi küsurat girebilirsiniz.")
        with connect() as c:
            cur=c.insert_shared("INSERT INTO beton_satis(tarih,cari_id,santiye_id,beton_cinsi,miktar_m3,birim_fiyat,pompa_ucreti,mikser_nakliye_ucreti,toplam,odeme_durumu,aciklama) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (tarih,cari_id,santiye_id,d["beton_cinsi"],miktar,fiyat,pompa_ucreti,nakliye,toplam,d["durum"],d["aciklama"]))
            sale_id=cur.lastrowid
            # Her satış için cari bağlantısını tek ortak fonksiyon üzerinden oluştur.
            sync_beton_sale_cari_movement(c, sale_id)
            if cari_id and d["durum"] == "Tahsil Edildi":
                method=d.get("odeme_yontemi") or "Nakit"
                ref=f"Beton satış tahsilatı #{sale_id}"
                c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                    (tarih,cari_id,"Alacak","Tahsilat",ref,toplam,method))
                c.insert_shared("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)",
                    (tarih,cari_id,"Gelir","Satış",ref,toplam,method))
                if method == "Nakit":
                    c.insert_shared("INSERT INTO nakit_kasa(tarih,yon,kaynak,aciklama,tutar,kullanici) VALUES(?,?,?,?,?,?)",
                        (tarih,"Para Girişi","Beton Satışı",ref,toplam,getattr(self,"current_user",{}).get("username","")))
                elif method == "Banka Havalesi / EFT":
                    c.insert_shared("INSERT INTO banka_hareket(tarih,banka,hesap,yon,aciklama,tutar) VALUES(?,?,?,?,?,?)",
                        (tarih,d.get("banka") or "Banka","","Para Girişi",ref,toplam))
                else:
                    vade=d.get("vade") or tarih
                    datetime.strptime(vade,"%Y-%m-%d")
                    cek_no=d.get("cek_no") or f"SAT-{sale_id}"
                    c.insert_shared("INSERT INTO cek(yon,banka,banka_kodu,cari_id,cek_no,kesideci_lehtar,vade,tutar,durum,notlar) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        ("Alınan",d.get("banka") or "","",cari_id,cek_no,"",vade,toplam,"Bekliyor",ref))
        self._refresh_selected_module()

    def delete_selected(self, tree, table):
        sel = tree.selection()
        if not sel: return messagebox.showinfo("Kayıt seçin","Önce silmek istediğiniz satırı seçin.", parent=self)
        row_id = int(tree.item(sel[0],"values")[0])
        if table in ("fatura", "beton_satis", "hareket"):
            with connect() as db:
                protected = record_is_posted(db, table, row_id)
        else:
            protected = False
        if protected:
            return messagebox.showinfo("Bağlı fatura kaydı", "Fatura, satış ve cari farkı birlikte korunur. "
                "Tahsilat yoksa eşleştirme işlemini İşlem Geri Al ile geri alabilirsiniz.", parent=self)
        if table == "fatura":
            with connect() as check_c:
                paid = check_c.execute("SELECT COUNT(*) FROM fatura_tahsilat WHERE fatura_id=?", (row_id,)).fetchone()[0]
            if paid:
                return messagebox.showerror("Tahsilatlı fatura korunuyor",
                    "Bu faturaya ait tahsilat kaydı var. Muhasebe geçmişini korumak için silme engellendi.",parent=self)
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
            if getattr(self, "current_user", {}).get("role") != "Yönetici":
                return messagebox.showinfo("Yetki gerekli", "Kalıcı cari silme işlemi yöneticiye açıktır.", parent=self)
            with connect() as c:
                receipt_count = c.execute("""SELECT COUNT(*) FROM fatura_tahsilat ft
                    JOIN fatura f ON f.id=ft.fatura_id WHERE f.cari_id=?""", (row_id,)).fetchone()[0]
            if receipt_count:
                return messagebox.showerror("Finansal kayıt koruması",
                    "Bu cariye bağlı fatura tahsilatları var. İşlem geçmişini ve tahsilat kayıtlarını "
                    "korumak için cari silinemez. Önce kayıtları muhasebe yetkilisiyle mutabık hâle getirin.",
                    parent=self)
            with connect() as c:
                cari = c.execute("SELECT unvan FROM cari WHERE id=?", (row_id,)).fetchone()
                if not cari: return messagebox.showinfo("Bulunamadı", "Cari hesap zaten silinmiş.", parent=self)
                counts = {t:c.execute(f"SELECT COUNT(*) FROM {t} WHERE cari_id=?",(row_id,)).fetchone()[0] for t in ("hareket","cek","beton_satis","fatura","gelen_fatura")}
            summary = ", ".join(f"{k}: {v}" for k,v in counts.items() if v)
            if not messagebox.askyesno("Cari ve tüm kayıtları sil", f"'{cari[0]}' cari hesabı ve bu hesaba bağlı TÜM hareketler, çekler, beton satışları ve faturalar silinecek.\n\n{summary or 'Bağlı kayıt bulunamadı.'}\n\nİşlem öncesi otomatik yedek alınacak. Devam edilsin mi?", parent=self): return
            try:
                with connect() as src:
                    backup_path = _verified_sqlite_backup(src, "cari_silme_oncesi")
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
                self._refresh_selected_module()
                messagebox.showinfo("Cari silindi", f"Cari ve bağlı kayıtları silindi.\nGeri dönüş yedeği: {backup_path}", parent=self)
            except Exception as e:
                messagebox.showerror("Silme başarısız", f"Cari silinemedi; işlem geri alındı.\n{e}", parent=self)
            return
        if not messagebox.askyesno("Onay","Seçili kayıt silinsin mi?", parent=self): return
        with connect() as c:
            if table in ("fatura", "beton_satis", "hareket"):
                c.execute("BEGIN IMMEDIATE")
                if record_is_posted(c, table, row_id):
                    return messagebox.showinfo("Bağlı fatura kaydı", "Kayıt bu sırada faturaya bağlandı.", parent=self)
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
                c.rollback()
                return messagebox.showerror("Kayıt silinemedi", "Bu kayıt başka işlemlere bağlı. Veri kaybını önlemek için silme durduruldu.", parent=self)
        self._refresh_selected_module()

    def _selected_check_number(self, tree):
        sel=tree.selection()
        return tree.item(sel[0],"values")[3] if sel else ""

    def _apply_tree_search(self, tree):
        rows = getattr(tree, "_all_rows", [])
        term = normalize_search_text(tree._search_var.get()) if hasattr(tree, "_search_var") else ""
        shown = [r for r in rows if term in normalize_search_text(" ".join(str(v or "") for v in r))] if term else rows
        if hasattr(tree, "_row_filter"):
            shown = [row for row in shown if tree._row_filter(row)]
        if hasattr(tree, "_on_search"):
            tree._on_search(shown)
        if tree is getattr(self, "borc_tree", None):
            shown = [row for row in rows if term in normalize_search_text(row[1])] if term else rows
            total = sum(Decimal(row[4].replace(",", "")) for row in shown)
            self.borc_total_label.config(text=f"Toplam kalan borç: {total:,.2f} TL   |   Borçlu cari: {len(shown)}")
        if tree is getattr(self, "mastarci_tree", None):
            volume, earned, paid = (sum(Decimal(row[index].replace(",", ".")) for row in shown)
                                    for index in (5, 7, 8))
            self.mastarci_total.configure(text=(
                f"Listelenen işler {len(shown)}   |   Toplam {volume:f} m³   |   "
                f"Hakediş {earned:.2f} TL   |   Ödenen {paid:.2f} TL   |   Kalan {earned - paid:.2f} TL"
            ).replace(".", ","))
        row_tags = getattr(tree, "_row_tags", {})
        signature = (tuple(tuple(r) for r in shown), tuple(sorted(row_tags.items())))
        children = tree.get_children()
        if getattr(tree, "_render_signature", None) == signature and len(children) == len(shown):
            return
        selected_ids = {str(tree.item(i,"values")[0]) for i in tree.selection() if tree.item(i,"values")}
        scroll_position = tree.yview()[0] if children else 0
        tree.delete(*children)
        for idx, row in enumerate(shown):
            tag = row_tags.get(str(row[0]))
            if not tag:
                tag = "even" if idx % 2 == 0 else "odd"
            item = tree.insert("", "end", values=row, tags=(tag,))
            if str(row[0]) in selected_ids:
                tree.selection_add(item)
        tree.yview_moveto(scroll_position)
        tree._render_signature = signature

    def fill_tree(self, tree, rows):
        tree._all_rows = list(rows)
        self._apply_tree_search(tree)

    def _set_dashboard_period(self, period):
        today = date.today()
        if period == "Tarih Aralığı":
            last_start, last_end = self.dashboard_dates
            fields = [("start", "Başlangıç (GG.AA.YYYY)", "entry", None),
                      ("end", "Bitiş (GG.AA.YYYY)", "entry", None)]
            initial = {"start": display_date(last_start) if last_start != "0001-01-01" else display_date(today.replace(day=1)),
                       "end": display_date(last_end) if last_end != "9999-12-31" else display_date(today)}
            values = self.prompt("Gelir / Gider Tarih Aralığı", fields, initial)
            if not values:
                return
            try:
                start = normalize_date(values["start"])
                end = normalize_date(values["end"])
                if not start or not end or start > end:
                    raise ValueError()
            except ValueError:
                return messagebox.showerror("Tarih aralığı", "Başlangıç ve bitiş tarihlerini GG.AA.YYYY biçiminde ve sırasıyla girin.", parent=self)
        elif period == "Bugün":
            start = end = today.isoformat()
        elif period == "Bu Ay":
            start = today.replace(day=1).isoformat()
            end = today.replace(day=monthrange(today.year, today.month)[1]).isoformat()
        else:
            start, end = "0001-01-01", "9999-12-31"
        self.dashboard_dates = (start, end)
        label = period if period != "Tarih Aralığı" else f"{display_date(start)} – {display_date(end)}"
        self.dashboard_period.set(label)
        self.refresh_dashboard()

    def refresh_dashboard(self):
        with connect() as c:
            cari = c.execute("SELECT COUNT(*) FROM cari").fetchone()[0]
            gelir, gider = c.execute("""SELECT
                COALESCE(SUM(CASE WHEN tur='Gelir' THEN tutar END),0),
                COALESCE(SUM(CASE WHEN tur='Gider' THEN tutar END),0)
                FROM hareket WHERE tarih BETWEEN ? AND ? AND tur IN ('Gelir','Gider')""",
                self.dashboard_dates).fetchone()
            cek = c.execute("SELECT COALESCE(SUM(tutar),0) FROM cek WHERE durum IN ('Bekliyor','Portföyde','Tahsilde')").fetchone()[0]
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
            self.debt_alert_label.config(text=f"Açık borçlu cari: {len(debt_rows)}. Alacak kaydı olmayan: {odemesiz}. Toplam açık cari bakiyesi: {acik_borc:,.2f} TL. Cari Hesaplar sekmesinde kırmızı/turuncu uyarılarla işaretlenir.")

    def _refresh_selected_module(self):
        page = self.tabs.select()
        if page == str(self.cari_tab):
            self.refresh_cari()
        elif page == str(self.borc_takip_tab):
            self.refresh_borc_takip()
        elif page == str(self.dashboard):
            self.refresh_dashboard()
        elif page == str(self.alisveris_cari_tab):
            self.refresh_alisveris_cari()
        else:
            with connect() as c:
                self._refresh_module_tables(c, page)

    def _refresh_module_tables(self, c, page=None):
        if page is None or page == str(self.hareket_tab):
            self.fill_tree(self.hareket_tree, c.execute("""SELECT h.id,h.tarih,h.tur,h.kategori,COALESCE(c.unvan,''),h.aciklama,printf('%.2f',h.tutar),h.odeme_yontemi FROM hareket h LEFT JOIN cari c ON c.id=h.cari_id ORDER BY h.tarih ASC,h.id ASC""").fetchall())
        if page is None or page == str(self.cek_tab):
            self._refresh_cheque_view(c)
        if page is None or page == str(self.mastarci_tab):
            mastar_rows = c.execute("""SELECT id,tarih,mastarci,musteri,santiye,miktar_m3,birim_fiyat,
                hakedis_kurus,odenen_kurus,aciklama FROM mastarci_hesap ORDER BY tarih,id""")
            self.fill_tree(self.mastarci_tree, [
                (rid, day, name, customer, site, format(Decimal(qty).normalize(), "f").replace(".", ","),
                 format(Decimal(price).normalize(), "f").replace(".", ","),
                 *(f"{Decimal(cents) / 100:.2f}".replace(".", ",") for cents in (total, paid, total - paid)), note)
                for rid, day, name, customer, site, qty, price, total, paid, note in mastar_rows
            ])
        if page is None or page == str(self.stok_tab):
            self.fill_tree(self.stok_tree, c.execute("SELECT id,kod,ad,birim,printf('%.2f',miktar),printf('%.2f',birim_fiyat),printf('%.2f',miktar*birim_fiyat) FROM stok ORDER BY ad").fetchall())
        if page is None or page == str(self.personel_tab):
            self.fill_tree(self.personel_tree, c.execute("SELECT id,ad,gorev,telefon,printf('%.2f',maas),notlar FROM personel ORDER BY ad").fetchall())
        if page is None or page == str(self.beton_tab):
            beton_rows = c.execute("""SELECT b.id,b.tarih,COALESCE(c.unvan,''),COALESCE(b.beton_cinsi,''),b.miktar_m3,
                b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,b.odeme_durumu,COALESCE(b.aciklama,''),COALESCE(s.ad,''),b.mikser_nakliye_ucreti
                FROM beton_satis b LEFT JOIN cari c ON c.id=b.cari_id LEFT JOIN santiye s ON s.id=b.santiye_id ORDER BY b.tarih ASC,b.id ASC""").fetchall()
            beton_display = []
            for b in beton_rows:
                detail = format_beton_aciklama(b[3],b[4],b[5],b[6],b[7],b[10],b[9],b[8],b[11])
                beton_display.append((b[0],b[1],(b[2] + (" / " + b[10] if b[10] else "")),b[3],format_m3(b[4]),f"{safe_float(b[5]):,.2f}",f"{safe_float(b[6]):,.2f}" if safe_float(b[6]) else "Yok",f"{b[11]:,.2f}" if b[11] else "-",f"{safe_float(b[7]):,.2f}",b[8],detail))
            self.fill_tree(self.beton_tree, beton_display)
            if hasattr(self, "santiye_tree"):
                try:
                    if self.santiye_tree.winfo_exists():
                        self.fill_tree(self.santiye_tree,c.execute("SELECT id,ad,COALESCE(adres,''),COALESCE(yetkili,''),COALESCE(telefon,''),COALESCE(notlar,'') FROM santiye ORDER BY ad").fetchall())
                except tk.TclError:
                    pass
        if page is None or page == str(self.fatura_tab):
            self.fill_tree(self.fatura_tree,c.execute("""SELECT f.id,f.tarih,COALESCE(c.unvan,''),COALESCE(f.belge_no,''),f.miktar_m3,
                printf('%.2f',f.ara_toplam),printf('%.2f',f.kdv_tutar),printf('%.2f',f.genel_toplam),
                printf('%.2f',f.tevkifat_tutar),printf('%.2f',f.tahsil_edilecek),
                printf('%.2f',f.odenen),printf('%.2f',f.tahsil_edilecek-f.odenen),
                CASE WHEN f.tahsil_edilecek-f.odenen<=0.005 THEN 'Tahsil edildi'
                     WHEN f.odenen>0 THEN 'Kısmi tahsilat' ELSE 'Tahsil edilmedi' END
                FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id ORDER BY f.tarih ASC,f.id ASC""").fetchall())
        if page is None or page == str(self.gelen_fatura_tab):
            if hasattr(self,"gelen_fatura_tree"):
                self.fill_tree(self.gelen_fatura_tree,c.execute("SELECT gf.id,gf.tarih,COALESCE(gf.vade,''),COALESCE(c.unvan,''),gf.belge_no,printf('%.2f',gf.ara_toplam),printf('%.2f',gf.kdv_tutar),printf('%.2f',gf.genel_toplam),printf('%.2f',gf.odenen),printf('%.2f',gf.genel_toplam-gf.odenen),gf.durum FROM gelen_fatura gf LEFT JOIN cari c ON c.id=gf.cari_id ORDER BY CASE WHEN gf.vade IS NULL OR gf.vade='' THEN gf.tarih ELSE gf.vade END ASC,gf.id ASC").fetchall())
        if page is None or page == str(self.kasa_tab):
            if hasattr(self, "kasa_tree"):
                self.fill_tree(self.kasa_tree,c.execute("SELECT id,tarih,yon,kaynak,COALESCE(aciklama,''),printf('%.2f',tutar),COALESCE(kullanici,'') FROM nakit_kasa ORDER BY tarih ASC,id ASC").fetchall())
            kasa_bakiye = c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE -tutar END),0) FROM nakit_kasa").fetchone()[0]
            if hasattr(self, "kasa_ozet"):
                self.kasa_ozet.config(text=f"Kasa bakiyesi: {safe_float(kasa_bakiye):,.2f} TL")
        if page is None or page == str(self.stok_hareket_tab):
            if hasattr(self,"stok_hareket_tree"):
                self.fill_tree(self.stok_hareket_tree,c.execute("SELECT sh.id,sh.tarih,COALESCE(s.ad,''),sh.yon,printf('%.2f',sh.miktar),COALESCE(sh.aciklama,'') FROM stok_hareket sh LEFT JOIN stok s ON s.id=sh.stok_id ORDER BY sh.tarih ASC,sh.id ASC").fetchall())
        if page is None or page == str(self.maas_tab):
            if hasattr(self,"maas_tree"):
                self.fill_tree(self.maas_tree,c.execute("SELECT m.id,m.tarih,COALESCE(p.ad,''),COALESCE(m.islem_turu,'Maaş'),m.donem,printf('%.2f',m.tutar),COALESCE(m.yontem,''),COALESCE(m.aciklama,'') FROM maas_odeme m LEFT JOIN personel p ON p.id=m.personel_id ORDER BY m.tarih ASC,m.id ASC").fetchall())
        if page is None or page == str(self.banka_tab):
            if hasattr(self,"banka_tree"):
                self.fill_tree(self.banka_tree,c.execute("SELECT id,tarih,banka,COALESCE(hesap,''),yon,karsi_taraf,COALESCE(aciklama,''),printf('%.2f',tutar) FROM banka_hareket ORDER BY tarih ASC,id ASC").fetchall())
                banka_net=c.execute("SELECT COALESCE(SUM(CASE WHEN yon='Para Girişi' THEN tutar ELSE -tutar END),0) FROM banka_hareket").fetchone()[0]
                self.banka_ozet.config(text=f"Banka net hareketi: {safe_float(banka_net):,.2f} TL")

    def refresh_all(self):
        started = time.monotonic()
        with connect() as c:
            self._refresh_module_tables(c)
        self.refresh_cari()
        self.refresh_borc_takip()
        self.refresh_dashboard()
        if self.tabs.select() == str(self.cari_card_page):
            self._refresh_open_cari_card()

        if hasattr(self, 'alisveris_tree'):
            self.refresh_alisveris_cari()
        elapsed = time.monotonic() - started
        if elapsed > 1.0:
            try:
                with open(os.path.join(USER_DATA_DIR, "performans.log"), "a", encoding="utf-8") as log:
                    log.write(f"{datetime.now().isoformat(timespec='seconds')} | refresh_all | {elapsed:.3f}s\n")
            except OSError:
                pass

    def pdf_menu(self):
        with connect() as c: cariler=c.execute("SELECT id,unvan FROM cari ORDER BY unvan").fetchall()
        choices=["Cari Hesap Ekstresi","Gelir-Gider Raporu","Çek Listesi","Alınan Çekler Dökümü","Verilen Çekler Dökümü",*CHEQUE_STATUS_REPORTS,"Fatura Taslağı PDF","Gelen Alış Faturaları","Kasa ve Banka Raporu"]
        d=self.prompt("PDF Raporu Oluştur",[("rapor","Rapor Türü","combo",choices)])
        if not d: return
        cari_id=None
        if d["rapor"]=="Fatura Taslağı PDF":
            with connect() as c: invoices=c.execute("SELECT id,COALESCE(belge_no,'Belgesiz'),tarih FROM fatura ORDER BY id DESC").fetchall()
            if not invoices: return messagebox.showinfo("Bilgi","Önce Fatura / Belge bölümünden bir taslak oluşturun.")
            q=self.prompt("Fatura Taslağı Seç",[("fatura","Taslak","combo",[f"{i} - {no} ({tarih})" for i,no,tarih in invoices])])
            if not q:return
            cari_id=int(q["fatura"].split(" - ",1)[0])
        if d["rapor"] == "Cari Hesap Ekstresi":
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

    def generate_pdf(self, report_type, path, cari_id=None, *, display_rows=None, sale_ids=None):
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_RIGHT
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        # Türkçe karakterleri destekleyen Windows yazı tipini bul; Helvetica son çare.
        font_name="Helvetica"
        font_paths=[os.path.join(APP_DIR,"fonts","DejaVuSans.ttf"),
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
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
        invoice_footer = []
        invoice_totals = []
        beton_summary_rows=None
        cari_summary_rows=None
        has_logo = os.path.isfile(logo_path)
        cari_watermark = has_logo and report_type in ("Cari Hesap Ekstresi", "Cari Hareket Dökümü")
        if has_logo and report_type != "Fatura Taslağı PDF":
            logo=RLImage(logo_path, width=64, height=64, kind="proportional", mask="auto")
            logo.hAlign="CENTER"
            story.extend([logo, Spacer(1,4)])
        if report_type == "Cari Hesap Ekstresi":
            story.extend([
                Paragraph("HOYTAŞ HAZIR BETON", title_style),
                Paragraph("FİNANS VE İŞLETME YÖNETİMİ", normal),
                Paragraph("Rapor Tarihi: "+datetime.now().strftime("%d.%m.%Y %H:%M"),normal),
                Spacer(1,8)
            ])
        elif report_type != "Fatura Taslağı PDF":
            story.extend([Paragraph("HOYTAŞ HAZIR BETON - Finans ve İşletme Yönetimi",title_style),
                          Paragraph(escape(report_type), title_style), Spacer(1,8),
                          Paragraph("Rapor tarihi: "+datetime.now().strftime("%d.%m.%Y %H:%M"),normal),Spacer(1,12)])
        with (nullcontext() if display_rows is not None else connect()) as c:
            if display_rows is not None:
                rows = [[str(cell) for cell in row] for row in display_rows]
            elif report_type=="Cari Hesap Ekstresi":
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
                sales_for_pdf = {r[0]: r[1:] for r in c.execute("SELECT b.id,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.aciklama,''),COALESCE(b.odeme_durumu,''),COALESCE(s.ad,''),b.mikser_nakliye_ucreti FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=?", (cari_id,)).fetchall()}
                bal=0.0
                for move_id,tarih,tur,kat,acik,odeme,tutar in moves:
                    import re
                    sale_match = re.search(r"(?:\[SATIŞ:(\d+)\]|Beton satış #?(\d+))", acik, re.IGNORECASE)
                    if sale_match and tur == "Borç":
                        sid = int(sale_match.group(1) or sale_match.group(2)); sale=sales_for_pdf.get(sid)
                        if sale:
                            cinsi,m3,fiyat,pompa,toplam_satis,notlar,durum,santiye,nakliye=sale
                            acik=format_beton_aciklama(cinsi,m3,fiyat,pompa,toplam_satis,santiye,notlar,durum,nakliye)
                    bal += tutar if tur=="Borç" else -tutar
                    balance_label=f"{abs(bal):,.2f} ({'B' if bal>=0 else 'A'})"
                    rows.append([display_date(tarih),f"{move_id:07d}",kat,"",Paragraph(escape(acik),normal),f"{tutar:,.2f}" if tur=="Borç" else "",f"{tutar:,.2f}" if tur=="Alacak" else "",balance_label])
                total_debit=sum(r[6] for r in moves if r[2]=='Borç'); total_credit=sum(r[6] for r in moves if r[2]=='Alacak')
                rows.append(["","","","","Rapor Toplamı",f"{total_debit:,.2f}",f"{total_credit:,.2f}",f"{abs(bal):,.2f} ({'B' if bal>=0 else 'A'})"])
                total_collected=c.execute("SELECT COALESCE(SUM(tutar),0) FROM hareket WHERE cari_id=? AND tur='Alacak' AND kategori NOT IN ('Fatura Farkı','Fatura') AND (kategori LIKE '%Tahsilat%' OR aciklama LIKE '%tahsilat%')",(cari_id,)).fetchone()[0]
                cari_summary_rows=[
                    ["CARİ HESAP ÖZETİ", "TOPLAM"],
                    ["Toplam Borç", f"{total_debit:,.2f} TL"],
                    ["Toplam Alacak", f"{total_credit:,.2f} TL"],
                    ["Toplam Tahsilat", f"{total_collected:,.2f} TL"],
                    ["Kalan Cari Bakiye", f"{abs(bal):,.2f} TL ({'Borç' if bal>=0 else 'Alacak'})"],
                ]
            elif report_type=="Cari Beton Satış Dökümü":
                cari=c.execute("SELECT unvan,telefon,vergi_no FROM cari WHERE id=?",(cari_id,)).fetchone()
                if not cari: raise ValueError("Cari hesap bulunamadı.")
                sales=c.execute("SELECT b.id,b.tarih,COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,COALESCE(b.odeme_durumu,''),COALESCE(s.ad,''),COALESCE(b.aciklama,''),b.mikser_nakliye_ucreti FROM beton_satis b LEFT JOIN santiye s ON s.id=b.santiye_id WHERE b.cari_id=? ORDER BY b.tarih ASC,b.id ASC",(cari_id,)).fetchall()
                selected = set(sale_ids) if sale_ids is not None else None
                sales = [row[1:] for row in sales if selected is None or row[0] in selected]
                story.append(Paragraph(f"<b>Cari:</b> {escape(cari[0])} &nbsp;&nbsp; <b>Telefon:</b> {escape(cari[1] or '-')}",normal))
                rows=[["Tarih","Beton Cinsi","Miktar (m³)","Birim Fiyat (TL)","Pompa (TL)","Mikser Nakliye (TL)","Toplam (TL)","Ödeme Durumu","Şantiye","Açıklama"]]
                for r in sales:
                    rows.append([display_date(r[0]),r[1],format_m3(r[2]),f"{r[3]:,.2f}",f"{r[4]:,.2f}" if r[4] else "-",f"{r[9]:,.2f}" if r[9] else "-",f"{r[5]:,.2f}",r[6],Paragraph(escape(r[7] or '-'),normal),Paragraph(escape(r[8]),normal)])
                total_m3=sum(safe_float(r[2] or 0) for r in sales); total_amount=sum(safe_float(r[5] or 0) for r in sales)
                rows.append(["GENEL TOPLAM","",f"{format_m3(total_m3)} m³","",f"{sum(r[4] for r in sales):,.2f}",
                             f"{sum(r[9] for r in sales):,.2f}",f"{total_amount:,.2f}","","",f"{len(sales)} satış"])
                story.append(Paragraph(f"Toplam beton miktarı: <b>{format_m3(total_m3)} m³</b> &nbsp;&nbsp; | &nbsp;&nbsp; Toplam satış: <b>{total_amount:,.2f} TL</b>",normal))
            elif report_type=="Cari Hareket Dökümü":
                cari=c.execute("SELECT unvan FROM cari WHERE id=?",(cari_id,)).fetchone()
                if not cari: raise ValueError("Cari hesap bulunamadı.")
                story.extend([Paragraph(escape(cari[0]), cari_title_style), Spacer(1, 6)])
                moves=c.execute("SELECT tarih,tur,kategori,COALESCE(aciklama,''),COALESCE(odeme_yontemi,''),tutar FROM hareket WHERE cari_id=? AND tur IN ('Borç','Alacak') ORDER BY tarih ASC,id ASC",(cari_id,)).fetchall()
                rows=[["Tarih","Tür","Kategori","Açıklama","Ödeme Yöntemi","Borç (TL)","Alacak (TL)","Bakiye (TL)"]]
                balance=0.0
                for tarih,tur,kategori,acik,odeme,tutar in moves:
                    balance += tutar if tur == "Borç" else -tutar
                    rows.append([display_date(tarih),tur,kategori,Paragraph(escape(acik),normal),odeme,f"{tutar:,.2f}" if tur == "Borç" else "",f"{tutar:,.2f}" if tur == "Alacak" else "",f"{balance:,.2f}"])
                debit=sum(r[5] for r in moves if r[1] == "Borç"); credit=sum(r[5] for r in moves if r[1] == "Alacak")
                rows.append(["","","","TOPLAM","",f"{debit:,.2f}",f"{credit:,.2f}",f"{balance:,.2f}"])
            elif report_type=="Açık Borç Listesi":
                rows=[["Cari Ünvanı", "Toplam Borç (TL)", "Toplam Alacak (TL)", "Kalan Borç (TL)", "Durum"]]
                debt_rows=c.execute("""SELECT unvan,toplam_borc,toplam_alacak,bakiye
                    FROM cari_bakiye WHERE bakiye > 0.005
                    ORDER BY bakiye DESC,unvan COLLATE NOCASE""").fetchall()
                total_debt=total_credit=total_remaining=0.0
                for name, debt, credit, remaining in debt_rows:
                    status="Alacak kaydı var" if credit > 0.005 else "Alacak kaydı yok"
                    rows.append([Paragraph(escape(name),normal),f"{debt:,.2f}",f"{credit:,.2f}",f"{remaining:,.2f}",status])
                    total_debt += safe_float(debt or 0); total_credit += safe_float(credit or 0); total_remaining += safe_float(remaining or 0)
                rows.append(["TOPLAM",f"{total_debt:,.2f}",f"{total_credit:,.2f}",f"{total_remaining:,.2f}",f"{len(debt_rows)} cari"])
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
            elif report_type in ("Çek Listesi", "Alınan Çekler Dökümü", "Verilen Çekler Dökümü", *CHEQUE_STATUS_REPORTS):
                rows = [["Yön","Banka","Banka Kodu","Çek No","Keşideci/Lehtar","Vade","Tutar (TL)","Durum","Notlar"]]
                view = CHEQUE_STATUS_REPORTS.get(report_type, {
                    "Alınan Çekler Dökümü": "received", "Verilen Çekler Dökümü": "own",
                }.get(report_type, "all"))
                checks = [row for row in c.execute("""SELECT yon,COALESCE(banka,''),COALESCE(banka_kodu,''),
                    COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar,durum,COALESCE(notlar,'')
                    FROM cek ORDER BY vade,id""") if cheque_matches_view(row[0], row[7], view)]
                for r in checks:
                    rows.append([r[0],r[1],r[2],r[3],Paragraph(escape(r[4]),normal),display_date(r[5]),f"{r[6]:,.2f}",r[7],Paragraph(escape(r[8]),normal)])
                rows.append(["","","","","TOPLAM",f"{len(checks)} adet",f"{sum(r[6] for r in checks):,.2f}","",""])
            elif report_type == "Fatura Taslağı PDF":
                invoice = c.execute("""SELECT f.tarih, f.belge_no, f.aciklama, f.miktar_m3, f.birim_fiyat,
                    f.ara_toplam, f.kdv_oran, f.kdv_tutar, f.genel_toplam, f.fatura_turu,
                    f.tevkifat_oran, f.tevkifat_tutar, f.tahsil_edilecek, f.belge_bilgileri,
                    f.cari_id, c.unvan, c.vergi_no FROM fatura f LEFT JOIN cari c ON c.id=f.cari_id
                    WHERE f.id=?""", (cari_id,)).fetchone()
                if not invoice:
                    raise ValueError("Fatura taslağı bulunamadı.")
                (day, number, description, quantity, price, base, vat_rate, vat, gross,
                 kind, withholding_rate, withholding, payable, stored, customer_id, customer, tax_id) = invoice
                info = invoice_document_details(stored, customer_id)
                def invoice_paragraph(value, style=normal):
                    return Paragraph(escape(str(value or "—")).replace("\n", "<br/>"), style)

                block_title = ParagraphStyle("InvoiceSection", parent=normal, fontSize=10,
                                             leading=14, textColor=colors.HexColor("#315C16"))
                columns = (A4[0] - 68) / 2
                def detail_block(left, right):
                    block = Table([[left, right]], colWidths=[columns, columns], hAlign="LEFT")
                    block.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F7F5")),
                        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0DAD4")),
                        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0DAD4")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 12),
                        ("TOPPADDING", (0, 0), (-1, -1), 10),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                    ]))
                    return block

                brand_style = ParagraphStyle("InvoiceBrand", parent=normal, fontSize=14, leading=18,
                                             textColor=colors.HexColor("#1F3412"))
                heading_style = ParagraphStyle("InvoiceHeading", parent=brand_style, fontSize=13,
                                               alignment=TA_RIGHT)
                draft_style = ParagraphStyle("InvoiceDraftNotice", parent=normal, alignment=TA_RIGHT)
                brand = [invoice_paragraph("HOYTAŞ", brand_style), invoice_paragraph("HAZIR BETON")]
                heading = [invoice_paragraph("SATIŞ FATURASI TASLAĞI", heading_style),
                           invoice_paragraph("Resmî belge yerine kullanılamaz.", draft_style)]
                icon = RLImage(logo_path, width=60, height=60, kind="proportional", mask="auto") if os.path.isfile(logo_path) else ""
                header = Table([[icon, brand, heading]], colWidths=[72, 124, columns * 2 - 196],
                               hAlign="LEFT")
                header.setStyle(TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LINEBELOW", (0, 0), (-1, -1), 1.2, colors.HexColor("#FFB800")),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ]))
                story.extend([
                    header, Spacer(1, 14),
                    detail_block(
                        [invoice_paragraph("SATICI", block_title), invoice_paragraph(info.get("satici_unvan") or "HOYTAŞ Hazır Beton"),
                         invoice_paragraph(info.get("satici_adres")),
                         invoice_paragraph(f"Vergi dairesi: {info.get('satici_daire') or '—'}"),
                         invoice_paragraph(f"VKN / TCKN: {info.get('satici_vkn') or '—'}")],
                        [invoice_paragraph("ALICI", block_title), invoice_paragraph(customer), invoice_paragraph(info.get("alici_adres")),
                         invoice_paragraph(f"Vergi dairesi: {info.get('alici_daire') or '—'}"),
                         invoice_paragraph(f"VKN / TCKN: {tax_id or '—'}")],
                    ),
                    Spacer(1, 10),
                    detail_block(
                        [invoice_paragraph("BELGE BİLGİLERİ", block_title),
                         invoice_paragraph(f"Taslak numarası: {number or '—'}"),
                         invoice_paragraph(f"Düzenleme tarihi: {display_date(day)}"),
                         invoice_paragraph(f"Senaryo: {info.get('senaryo') or 'Seçilmedi'}"),
                         invoice_paragraph(f"Fatura tipi: {'TEVKIFAT' if kind == 'Tevkifatlı' else 'SATIS'}")],
                        [invoice_paragraph("TESLİMAT BİLGİLERİ", block_title),
                         invoice_paragraph(f"Teslim tarihi: {display_date(info['teslim']) if info.get('teslim') else '—'}"),
                         invoice_paragraph(f"İrsaliye numarası: {info.get('irsaliye') or '—'}"),
                         *([invoice_paragraph(f"Tevkifat işlemi: {info.get('tevkifat_kodu') or '—'}")]
                           if withholding_rate else [])],
                    ),
                    Spacer(1, 14),
                ])
                rows = [["Sıra", "Mal / Hizmet", "Miktar", "Birim fiyat (TL)", "KDV oranı", "Mal / hizmet tutarı (TL)"],
                        ["1", f"Hazır beton\n{description}" if description else "Hazır beton",
                         format_m3(quantity).translate(str.maketrans(",.", ".,")) + " m³", invoice_money(price), f"%{vat_rate:g}", invoice_money(base)]]
                invoice_totals = [("Mal / hizmet toplamı", base), ("KDV matrahı", base),
                                  (f"Hesaplanan KDV (%{vat_rate:g})", vat), ("Vergiler dahil toplam", gross)]
                if withholding_rate:
                    invoice_totals.append((f"KDV tevkifatı ({withholding_rate * 10:.0f}/10)", withholding))
                invoice_totals.append(("Ödenecek tutar", payable))
                invoice_footer = [
                    f"Yalnız {invoice_amount_in_words(payable)}",
                    "Fatura bilgileri: 213 sayılı Vergi Usul Kanunu, madde 230.",
                    "Elektronik belge düzenleme: 509 Sıra No.lu Vergi Usul Kanunu Genel Tebliği.",
                ]
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
                rows=[["Tarih","Müşteri","Beton Cinsi","Miktar (m³)","Birim Fiyat","Pompa Ücreti","Mikser Nakliye (TL)","Toplam (TL)","Ödeme Durumu","Açıklama"]]
                for r in c.execute("SELECT b.tarih,COALESCE(c.unvan,''),COALESCE(b.beton_cinsi,''),b.miktar_m3,b.birim_fiyat,COALESCE(b.pompa_ucreti,0),b.toplam,b.odeme_durumu,COALESCE(b.aciklama,''),COALESCE(s.ad,''),b.mikser_nakliye_ucreti FROM beton_satis b LEFT JOIN cari c ON c.id=b.cari_id LEFT JOIN santiye s ON s.id=b.santiye_id ORDER BY b.tarih ASC,b.id ASC"):
                    detail = format_beton_aciklama(r[2],r[3],r[4],r[5],r[6],r[9],r[8],r[7],r[10])
                    rows.append([display_date(r[0]),Paragraph(escape(r[1]),normal),r[2],format_m3(r[3]),f"{r[4]:,.2f}",f"{r[5]:,.2f}" if r[5] else "Yok",f"{r[10]:,.2f}" if r[10] else "-",f"{r[6]:,.2f}",r[7],Paragraph(escape(detail),normal)])
                total=c.execute("SELECT COALESCE(SUM(miktar_m3),0),COALESCE(SUM(toplam),0),COUNT(*),COALESCE(SUM(pompa_ucreti),0),COALESCE(SUM(mikser_nakliye_ucreti),0) FROM beton_satis").fetchone()
                rows.append(["","GENEL TOPLAM","",f"{format_m3(total[0])} m³","",f"{total[3]:,.2f}",
                             f"{total[4]:,.2f}",f"{total[1]:,.2f}","",""])
                story.append(Paragraph(f"Toplam beton miktarı: <b>{format_m3(total[0])} m³</b>"
                                       f" &nbsp;&nbsp; | &nbsp;&nbsp; Toplam satış: <b>{total[1]:,.2f} TL</b>",normal))
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
        invoice_header = ParagraphStyle("InvoiceTableHeader", parent=normal, textColor=colors.white)
        for row_index, row in enumerate(rows):
            row_style = invoice_header if row_index == 0 else normal
            wrapped=[]
            for cell in row:
                if isinstance(cell, RLParagraph):
                    wrapped.append(cell)
                elif isinstance(cell, str):
                    safe = escape(cell).replace("\n", "<br/>")
                    wrapped.append(RLParagraph(safe, row_style))
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
            **{title: [0.65, 0.90, 0.85, 0.95, 1.45, 0.85, 0.95, 0.85, 1.35] for title in CHEQUE_STATUS_REPORTS},
            "Cari Hareket Dökümü": [0.8, 0.8, 1.0, 2.0, 1.0, 1.0, 1.0, 1.0],
            "Cari Beton Satış Dökümü": [0.75, 0.9, 0.75, 0.9, 0.75, 0.85, 0.9, 1.0, 1.0, 1.6],
            "Beton Satış Raporu": [0.78, 1.45, 0.90, 0.70, 0.90, 0.85, 0.85, 1.00, 1.05, 1.45],
            "Beton Satışları Dökümü": [0.78, 1.45, 0.90, 0.70, 0.90, 0.85, 0.85, 1.00, 1.05, 1.45],
            "Fatura Taslağı PDF": [0.45, 2.20, 0.85, 1.20, 0.80, 1.35],
            "Gelen Alış Faturaları": [0.85, 0.85, 1.45, 1.05, 1.55, 1.00, 1.00, 1.00, 0.95],
            "Kasa ve Banka Raporu": [0.85, 1.00, 1.15, 1.05, 2.10, 1.00, 1.00],
            "Açık Borç Listesi": [2.20, 1.05, 1.05, 1.15, 1.20],
        }.get(report_type)
        col_count = max(1, max(len(r) for r in wrapped_rows))
        if not width_ratios or len(width_ratios) != col_count:
            width_ratios = [1.0] * col_count
        ratio_sum = sum(width_ratios)
        col_widths = [page_width * ratio / ratio_sum for ratio in width_ratios]
        # Başlık satırı yalnızca ilk sayfada görünür; her sayfada tarih/tür satırı
        # tekrar etmez ve tablo sütunları sayfa genişliğinin dışına taşmaz.
        table=Table(wrapped_rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT", splitByRow=1)
        header_bg=colors.HexColor("#315C16") if report_type=="Cari Hesap Ekstresi" else colors.HexColor("#315C16")
        header_fg=colors.white
        alternate=colors.HexColor("#F4F8EF") if report_type=="Cari Hesap Ekstresi" else colors.HexColor("#f2f5f8")
        body_backgrounds = [None, colors.Color(244 / 255, 248 / 255, 239 / 255, alpha=0.35)] if cari_watermark else [colors.white, alternate]
        table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),header_bg),("TEXTCOLOR",(0,0),(-1,0),header_fg),
            ("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),7),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),
            ("VALIGN",(0,0),(-1,-1),"TOP"),("ROWBACKGROUNDS",(0,1),(-1,-1),body_backgrounds),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        story.append(table)
        if invoice_totals:
            totals_table = Table([[invoice_paragraph(label), invoice_paragraph(invoice_money(amount))]
                                  for label, amount in invoice_totals],
                                 colWidths=[178, 93], hAlign="RIGHT")
            totals_table.setStyle(TableStyle([
                ("LINEABOVE", (0, 0), (-1, -1), 0.4, colors.HexColor("#D0DAD4")),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DFEDE4")),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            story.extend([Spacer(1, 12), totals_table])
        if invoice_footer:
            story.append(Spacer(1, 10))
            story.extend(Paragraph(escape(line), normal) for line in invoice_footer)
        if cari_summary_rows:
            story.extend([Spacer(1,12), Paragraph("CARİ HESAP ÖZETİ", title_style), Spacer(1,5)])
            summary_table=Table(cari_summary_rows,colWidths=[page_width*0.58,page_width*0.42],hAlign="LEFT")
            summary_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#315C16")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),8),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),("ROWBACKGROUNDS",(0,1),(-1,-1),body_backgrounds),("ALIGN",(1,1),(1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
            story.append(summary_table)
        if beton_summary_rows:
            story.extend([Spacer(1,12), Paragraph("BETON M³ ÖZETİ", title_style), Spacer(1,5)])
            summary_table=Table(beton_summary_rows,colWidths=[page_width*0.38,page_width*0.16,page_width*0.22,page_width*0.24],repeatRows=1)
            summary_table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#315C16")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),font_name),("FONTSIZE",(0,0),(-1,-1),8),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#cbd0d4")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F4F8EF")]),("ALIGN",(1,1),(-1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
            story.append(summary_table)
        def draw_cari_watermark(canvas, document):
            page_width, page_height = document.pagesize
            width = page_width * 0.48
            canvas.saveState()
            canvas.setFillAlpha(0.09)
            canvas.drawImage(logo_path, (page_width - width) / 2, (page_height - width) / 2,
                             width, width, preserveAspectRatio=True, anchor="c", mask="auto")
            canvas.restoreState()

        if cari_watermark:
            doc.build(story, onFirstPage=draw_cari_watermark, onLaterPages=draw_cari_watermark)
        else:
            doc.build(story)

    def check_due_cheques(self):
        today=date.today(); soon=(today+timedelta(days=7)).isoformat(); today_s=today.isoformat()
        with connect() as c:
            rows=c.execute("SELECT yon,COALESCE(cek_no,''),COALESCE(kesideci_lehtar,''),vade,tutar FROM cek WHERE durum IN ('Bekliyor','Portföyde','Tahsilde') AND vade<=? ORDER BY vade ASC,id ASC",(soon,)).fetchall()
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
                        scale = 85 / max(logo.width, logo.height)
                        logo.width *= scale
                        logo.height *= scale
                        ws.add_image(logo, "A1")
                    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=5)
                    title_cell = ws.cell(row=6, column=1, value=f"HOYTAŞ HAZIR BETON | FİNANS VE İŞLETME YÖNETİMİ — {sheet_name.upper()}")
                    title_cell.font = Font(name="Calibri", size=14, bold=True, color="315C16")
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
                        cell.fill = PatternFill("solid", fgColor="315C16")
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
        """Run daily verified backup outside the Tk event loop."""
        results = queue.Queue()
        def worker():
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True)
                today = datetime.now().strftime("%Y%m%d")
                if not any(name.startswith("hoytas_" + today) and name.endswith(".db") for name in os.listdir(BACKUP_DIR)):
                    self._create_backup_file()
                results.put(None)
            except Exception as exc:
                results.put(str(exc))
        def poll():
            try:
                error = results.get_nowait()
            except queue.Empty:
                self.after(250, poll)
                return
            if error:
                messagebox.showwarning("Otomatik yedekleme", "Otomatik yedek alınamadı. Lütfen Şimdi Yedekle seçeneğini kullanın.\n\n" + error, parent=self)
        threading.Thread(target=worker, daemon=True).start()
        self.after(250, poll)

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



INVOICE_DRAFT_PROFILES = ("", "TEMELFATURA", "TICARIFATURA", "EARSIVFATURA")


def invoice_document_details(stored, customer_id):
    details = json.loads(stored or "{}")
    if details.get("cari_id") != customer_id:
        details.pop("alici_adres", None)
        details.pop("alici_daire", None)
    return details


def invoice_money(value):
    return f"{value:,.2f}".translate(str.maketrans(",.", ".,"))


def invoice_amount_in_words(value):
    units = ("", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz")
    tens = ("", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Seksen", "Doksan")
    scales = ("", "Bin", "Milyon", "Milyar", "Trilyon", "Katrilyon")

    def words(number):
        if number >= 1000 ** len(scales):
            return f"{number:,}".replace(",", ".")
        groups = []
        for scale in scales:
            number, group = divmod(number, 1000)
            if not group:
                continue
            hundreds, remainder = divmod(group, 100)
            prefix = (units[hundreds] if hundreds > 1 else "") + ("Yüz" if hundreds else "")
            triplet = prefix + tens[remainder // 10] + units[remainder % 10]
            groups.append("Bin" if scale == "Bin" and group == 1 else triplet + scale)
        return "".join(reversed(groups)) or "Sıfır"

    cents = int((Decimal(str(value)) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    lira, kurus = divmod(abs(cents), 100)
    return f"{'Eksi ' if cents < 0 else ''}{words(lira)} Türk Lirası {words(kurus)} Kuruş"


def normalize_fatura_tarihi(value):
    """Accept Turkish dotted date or ISO; store ISO consistently."""
    value = (value or "").strip()
    if not value:
        return date.today().isoformat()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("Fatura tarihi GG.AA.YYYY biçiminde olmalıdır.")


if __name__ == "__main__":
    app = None
    if os.name == "nt" and sys.version_info[:2] == (3, 14):
        bundled = os.path.join(os.environ["LOCALAPPDATA"], "HOYTASFinans", "Python314-Kitaplik")
        if os.path.isdir(bundled):
            sys.path.append(bundled)
    if (os.path.basename(os.path.normpath(ISOLATED_PROFILE_DIR)) == "canli-kullanim"
            and os.environ.get("HOYTAS_HYBRID_CAPTURE") == "1"):
        from hybrid_windows_profile import ENDPOINT

        os.environ["HOYTAS_MEMBERSHIP_URL"] = ENDPOINT
        os.environ["HOYTAS_SYNC_URL"] = ENDPOINT
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        # If the desktop icon was lost during an earlier update, recreate it on first launch.
        try:
            if not ISOLATED_PROFILE_DIR:
                _restore_desktop_shortcut_after_update(APP_DIR)
        except Exception as shortcut_exc:
            try:
                with open(os.path.join(USER_DATA_DIR, "guncelleme.log"), "a", encoding="utf-8") as log:
                    log.write(f"{datetime.now().isoformat(timespec='seconds')} | Kısayol: {shortcut_exc}\n")
            except OSError:
                pass
        import faulthandler
        _hang_log = open(os.path.join(USER_DATA_DIR, "donma_teshis.log"), "a", encoding="utf-8")
        faulthandler.enable(file=_hang_log)
        faulthandler.dump_traceback_later(20, repeat=True, file=_hang_log)
        app = HoytasApp()
        if ISOLATED_PROFILE_DIR and os.name == "nt":
            try:
                from hybrid_shortcut_setup import install_after_update

                install_after_update(ISOLATED_PROFILE_DIR, APP_VERSION, app.membership_client is not None)
            except Exception as shortcut_exc:
                app._log_update_error(shortcut_exc)
                messagebox.showwarning("HOYTAŞ kısayolu", f"v{APP_VERSION} açıldı, ancak masaüstü simgesi yenilenemedi. Ayrıntı güncelleme kaydına yazıldı.", parent=app)
        if os.environ.get("HOYTAS_SYNC_URL", "").strip():
            from hybrid_runtime import start_background_sync

            start_background_sync(app, connect, USER_DATA_DIR)
        # HoytasApp.__init__ already performs the initial refresh; avoid a duplicate startup pass.
        app.after(500, app.automatic_backup)
        app.after(900, app.check_due_cheques)
        def heartbeat():
            faulthandler.cancel_dump_traceback_later()
            faulthandler.dump_traceback_later(20, repeat=True, file=_hang_log)
            app.after(1000, heartbeat)
        app.after(1000, heartbeat)
        app.mainloop()
        faulthandler.cancel_dump_traceback_later()
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

