"""v83 source-only regression checks; no customer database access."""
import ast
import datetime
from pathlib import Path

source = Path("hoytas.py").read_text(encoding="utf-8")
tree = ast.parse(source)
assert "APP_VERSION = 86" in source
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
