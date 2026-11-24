from datetime import date
from decimal import Decimal
from pathlib import Path
import sqlite3
import tempfile
import unittest

from hoytas import connect, puantaj_mesai_tutari, puantaj_sureleri, puantaj_ucret_bol, save_puantaj_approval


class PuantajTests(unittest.TestCase):
    def test_work_time_and_pay(self):
        self.assertEqual(puantaj_sureleri("Çalıştı", "08:00", "19:00", "60", "8"), (60, 480, 600, 120))
        self.assertEqual(puantaj_sureleri("Çalıştı", "22:00", "06:30", "30", "7,5"), (30, 450, 480, 30))
        self.assertEqual(puantaj_sureleri("İzin", "", "", "", ""), (0, 0, 0, 0))
        with self.assertRaises(ValueError):
            puantaj_sureleri("Çalıştı", "08:00", "09:00", "60", "8")
        self.assertEqual(puantaj_mesai_tutari(30000, 150, "225", "1,5"), Decimal("500.00"))
        with self.assertRaises(ValueError):
            puantaj_mesai_tutari(30000, 150, "0", "1,5")

    def test_daily_and_hourly_rates_are_independent_of_overtime_multiplier(self):
        self.assertEqual(puantaj_ucret_bol("30000", "30"), Decimal("1000"))
        self.assertEqual(puantaj_ucret_bol("30000", "225").quantize(Decimal(".01")), Decimal("133.33"))
        self.assertEqual(puantaj_ucret_bol("30000", "23,5"), Decimal("30000") / Decimal("23.5"))
        with self.assertRaises(ValueError):
            puantaj_ucret_bol("30000", "0")

    def test_approval_is_monthly_and_does_not_duplicate_payment(self):
        with tempfile.TemporaryDirectory() as folder, connect(Path(folder) / "puantaj.db") as db:
            employee = db.execute("INSERT INTO personel(ad,maas) VALUES('Örnek Personel',30000)").lastrowid
            for day, overtime in (("2026-11-01", 120), ("2026-11-30", 30), ("2026-12-01", 60)):
                db.execute("""INSERT INTO personel_puantaj
                    (personel_id,tarih,durum,baslangic,bitis,mola,normal_dakika,net_dakika,mesai_dakika)
                    VALUES(?,?,'Çalıştı','08:00','19:00',60,480,?,?)""",
                    (employee, day, 480 + overtime, overtime))
            month = date(2026, 11, 1)
            save_puantaj_approval(db, employee, month, "225", "1,5")
            save_puantaj_approval(db, employee, month, "200", "2")
            self.assertEqual(db.execute("SELECT donem,tutar,islem_turu FROM maas_odeme").fetchall(),
                             [("2026-11", 500, "Mesai")])
            with self.assertRaises(sqlite3.IntegrityError):
                db.execute("UPDATE maas_odeme SET tutar=1")
            with self.assertRaises(sqlite3.IntegrityError):
                db.execute("DELETE FROM maas_odeme")
            save_puantaj_approval(db, employee, month, "", "", revoke=True)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM maas_odeme").fetchone(), (0,))
            self.assertEqual(db.execute("SELECT COUNT(*) FROM personel_puantaj").fetchone(), (3,))
            self.assertIsNone(db.execute("PRAGMA foreign_key_check").fetchone())
