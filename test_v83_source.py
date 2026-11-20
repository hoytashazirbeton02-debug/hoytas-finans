"""v83 source-only regression checks; no customer database access."""
import ast
import datetime
from pathlib import Path

source = Path("hoytas.py").read_text(encoding="utf-8")
tree = ast.parse(source)
assert "APP_VERSION = 94" in source
assert source.index("def normalize_fatura_tarihi") < source.index('if __name__ == "__main__":')
node = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "normalize_fatura_tarihi")
namespace = {"date": datetime.date, "datetime": datetime.datetime}
exec(compile(ast.Module(body=[node], type_ignores=[]), "hoytas_date_helper", "exec"), namespace)
normalize = namespace["normalize_fatura_tarihi"]
assert normalize("18.11.2026") == "2026-11-18"
assert normalize("2026-11-18") == "2026-11-18"
for invalid in ("31.02.2026", "2026/11/18"):
    try:
        normalize(invalid)
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid date accepted: " + invalid)
assert "INSERT INTO fatura_tahsilat" not in source.split("    def tahsilat_fatura(self):", 1)[1].split("    def cari_tahsilat(self):", 1)[0]
print("v83 static date and receipt guard checks passed")

# Trigger PR validation on v83 candidate branch.

# Invoice allocation must use one receipt and one invoice ledger entry, with a balance cap.
section = source.split("    def cari_tahsilat(self):", 1)[1].split("    def refresh_cari(self):", 1)[0]
assert "fatura_tahsilat(fatura_id,tarih,tutar,yontem,aciklama)" in section
assert "UPDATE fatura SET odenen=odenen+?" in section
assert "tutar > round(safe_float(invoice[0])-safe_float(invoice[1]), 2)+0.005" in section
assert section.count('INSERT INTO nakit_kasa(') == 1
assert section.count('INSERT INTO banka_hareket(') == 1

# v88: exact legacy IDs, mismatched values and ambiguous legacy rows.
import sqlite3, re
ns = {"re": re, "safe_float": float}
for name in ("linked_beton_sale_id", "audit_beton_cari_amounts", "sync_beton_sale_cari_movement"):
    fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == name)
    exec(compile(ast.Module(body=[fn], type_ignores=[]), "v88", "exec"), ns)
ns["format_beton_aciklama"] = lambda *a, **k: "Beton"
db = sqlite3.connect(":memory:")
db.executescript("""
CREATE TABLE beton_satis(id INTEGER,tarih TEXT,cari_id INTEGER,beton_cinsi TEXT,miktar_m3 REAL,birim_fiyat REAL,pompa_ucreti REAL,toplam REAL,aciklama TEXT,odeme_durumu TEXT);
CREATE TABLE hareket(id INTEGER PRIMARY KEY,cari_id INTEGER,tarih TEXT,tutar REAL,aciklama TEXT,tur TEXT,kategori TEXT,odeme_yontemi TEXT);
INSERT INTO beton_satis VALUES(12,'2026-11-18',1,'C30',28,4200,5000,122600,'','Vadeli');
INSERT INTO hareket VALUES(1,1,'2026-11-18',111200,'[SATIŞ:12]','Borç','Satış','Vadeli');
INSERT INTO hareket VALUES(2,2,'2026-11-18',100,'Beton satış #123 - eski','Borç','Satış','Vadeli');
INSERT INTO hareket VALUES(3,1,'2026-11-18',200000,'Tahsilat','Alacak','Tahsilat','Banka');
""")
audit = ns["audit_beton_cari_amounts"]
assert audit(db,1)[1][0][2] - audit(db,1)[1][0][4] == 11400
ns["sync_beton_sale_cari_movement"](db,12)
assert audit(db,1) == ([],[],[])
assert db.execute("SELECT tutar FROM hareket ORDER BY id").fetchall() == [(122600,), (100,), (200000,)]
db.execute("UPDATE hareket SET aciklama='Eski kayit',tutar=111200 WHERE id=1")
assert audit(db,1)[0] == [] and len(audit(db,1)[2]) == 1
print("v88 reconciliation checks passed")

fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "audit_all_cari")
exec(compile(ast.Module(body=[fn],type_ignores=[]),"v89","exec"),ns)
db.executescript("""
CREATE TABLE cari(id INTEGER,unvan TEXT);
INSERT INTO cari VALUES(1,'A'),(2,'B');
CREATE VIEW cari_bakiye AS SELECT c.id cari_id, 0 toplam_borc, 0 toplam_alacak, 0 bakiye FROM cari c;
""")
rows=ns["audit_all_cari"](db)
assert len(rows)==2 and rows[0][4] and rows[0][5]
assert db.execute("SELECT COUNT(*) FROM hareket").fetchone()[0]==3
print("v89 bulk read-only audit checks passed")

assert 'self.after(180, self._run_tab_refresh)' in source
assert 'footer=ttk.Frame(win,padding=8)' in source
assert 'self._adapt_action_rows(self)' in source
assert 'tree.yview_moveto(scroll_position)' in source

assert 'w.pack_configure(before=first)' in source
assert 'show="headings", height=6' in source
