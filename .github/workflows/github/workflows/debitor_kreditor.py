#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Debitor-Kreditor proqramı
=========================
- Qaimələr   -> qaimeler/   papkası (şablon: sablonlar/sablon_qaime.xlsx)
- Ödənişlər  -> odenisler/  papkası (şablon: sablonlar/sablon_odenis.xlsx)
- "Papkaları yenilə" düyməsi yalnız YENİ və ya DƏYİŞMİŞ faylları oxuyur.
- Məlumat bazası: SQLite (tək fayl, quraşdırma tələb etmir). Məbləğlər qəpiklə
  (tam ədəd) saxlanır - toplamalarda yuvarlaqlaşdırma xətası olmur.

İstifadə:
    python debitor_kreditor.py                      # pəncərəli proqram
    python debitor_kreditor.py --sablon             # Excel şablonlarını yarat
    python debitor_kreditor.py --yenile             # papkaları oxu
    python debitor_kreditor.py --export hesabat.xlsx [--tarix 31.12.2026]
"""
import sys
import json
import time
import sqlite3
import argparse
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "config.json"
VAT_DEFAULT = 18  # ƏDV dərəcəsi (%), şablonda "ƏDV %" sütunu boşdursa
DEFAULTS = {
    "qaime": str(BASE / "qaimeler"),
    "odenis": str(BASE / "odenisler"),
    "db": str(BASE / "debitor_kreditor.db"),
}

KIND_LABEL = {
    "S": "Satış qaiməsi",
    "SQ": "Satış qaytarılması",
    "A": "Alış qaiməsi",
    "AQ": "Alış qaytarılması",
    "D": "Daxil olan ödəniş",
    "C": "Çıxan ödəniş",
}

# ----------------------------------------------------------------------------
# Konfiqurasiya
# ----------------------------------------------------------------------------

def load_cfg():
    cfg = dict(DEFAULTS)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def save_cfg(cfg):
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------------------------------------------------------
# Verilənlər bazası (SQLite)
# ----------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(path TEXT PRIMARY KEY, mtime REAL, size INTEGER);
CREATE TABLE IF NOT EXISTS parties(voen TEXT PRIMARY KEY, name TEXT);
CREATE TABLE IF NOT EXISTS txn(
    id INTEGER PRIMARY KEY,
    src TEXT NOT NULL,
    row_no INTEGER,
    kind TEXT NOT NULL,          -- S satış, SQ satış qaytarılma, A alış, AQ alış qaytarılma, D daxil olan, C çıxan
    date TEXT NOT NULL,          -- ISO: YYYY-MM-DD
    doc_no TEXT,
    voen TEXT NOT NULL,
    amount INTEGER NOT NULL,     -- qəpik (qaytarılmalar da müsbət saxlanır)
    note TEXT,
    vat INTEGER DEFAULT 18,      -- ƏDV % (yalnız 'part' boşdursa istifadə olunur)
    part TEXT DEFAULT ''         -- E = məbləğ əsasdır, V = məbləğ ƏDV-dir, boş = ƏDV daxil
);
CREATE INDEX IF NOT EXISTS ix_txn_voen_date ON txn(voen, date);
CREATE INDEX IF NOT EXISTS ix_txn_src ON txn(src);
"""


def connect(path):
    con = sqlite3.connect(str(path))
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA temp_store=MEMORY")
    con.execute("PRAGMA cache_size=-65536")
    con.executescript(SCHEMA)
    cols = [r[1] for r in con.execute("PRAGMA table_info(txn)")]
    if "part" not in cols:
        con.execute("ALTER TABLE txn ADD COLUMN part TEXT DEFAULT ''")
    if "vat" not in cols:
        con.execute("ALTER TABLE txn ADD COLUMN vat INTEGER DEFAULT 18")
    # köhnə bazada mənfi qaimələr = qaytarılma -> ayrı növə keçir
    with con:
        con.execute("UPDATE txn SET kind='SQ', amount=-amount WHERE kind='S' AND amount<0")
        con.execute("UPDATE txn SET kind='AQ', amount=-amount WHERE kind='A' AND amount<0")
    return con


# ----------------------------------------------------------------------------
# Dəyərlərin təhlili
# ----------------------------------------------------------------------------

_TR = str.maketrans("əıöüçşğ", "eioucsg")


def norm(s):
    """Azərbaycan hərflərini sadələşdirib kiçik hərfə çevirir (müqayisə üçün)."""
    return str(s).strip().lower().replace("i\u0307", "i").translate(_TR)


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float)):
        if 30000 < v < 80000:
            return from_excel(v).date().isoformat()
        return None
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def parse_amount(v):
    """Məbləği qəpiklə (int) qaytarır. '1 234,56' və '1,234.56' formatlarını anlayır."""
    if v is None or str(v).strip() == "":
        raise ValueError("məbləğ boşdur")
    if isinstance(v, (int, float)):
        d = Decimal(str(v))
    else:
        s = str(v).strip().replace(" ", "").replace("\u00a0", "")
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            d = Decimal(s)
        except InvalidOperation:
            raise ValueError("məbləğ rəqəm deyil")
    return int((d * 100).to_integral_value(rounding=ROUND_HALF_UP))


def parse_vat(v):
    """ƏDV dərəcəsi (%). Boşdursa 18. '18', '18%', '0', 'ƏDV-siz' anlaşılır."""
    if v is None or str(v).strip() == "":
        return VAT_DEFAULT
    s = str(v).strip().replace("%", "").replace(",", ".")
    try:
        d = Decimal(s)
        if d < 1 and d > 0:   # Excel-də 18% = 0.18
            d *= 100
        return int(d)
    except InvalidOperation:
        return 0 if norm(s).startswith(("yox", "siz", "ed")) else VAT_DEFAULT


def vat_part(amount, rate):
    """ƏDV daxil məbləğdən ƏDV hissəsi (qəpiklə)."""
    if not rate:
        return 0
    return int((Decimal(amount) * rate / (100 + rate)).to_integral_value(rounding=ROUND_HALF_UP))


def parse_part(v):
    """'Əsas' -> 'E', 'ƏDV' -> 'V', digər -> ''."""
    if v is None or isinstance(v, (int, float)):
        return ""
    n = norm(v)
    if n.startswith("es"):
        return "E"
    if n.startswith("edv") or n.startswith("ed"):
        return "V"
    return ""


def split_amount(amount, part, rate):
    """(əsas, ƏDV) qəpiklə. part: E/V -> məbləğ tam o hissədir; boş -> ƏDV daxil sayılıb ayrılır."""
    if part == "E":
        return amount, 0
    if part == "V":
        return 0, amount
    vt = vat_part(amount, rate if rate is not None else 0)
    return amount - vt, vt


def parse_voen(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def kind_of(ftype, s):
    n = norm(s or "")
    if ftype == "Q":
        ret = "qayt" in n or "iade" in n
        if n.startswith("sat"):
            return "SQ" if ret else "S"
        if n.startswith("al"):
            return "AQ" if ret else "A"
    else:
        if n.startswith(("dax", "med", "in")):
            return "D"
        if n.startswith(("cix", "mex", "out")):
            return "C"
    return None


# ----------------------------------------------------------------------------
# Papkaların oxunması
# ----------------------------------------------------------------------------

def import_file(con, path, ftype):
    """Bir Excel faylını oxuyub bazaya yazır. (sətir sayı, xəta siyahısı) qaytarır."""
    errors, rows, names = [], [], {}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return 0, [(path.name, 0, f"Fayl açılmadı: {e}")], False
    try:
        ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        for i, r in enumerate(ws.iter_rows(min_row=1, max_col=9, values_only=True), start=1):
            r = tuple(r) + (None,) * (9 - len(r))
            if all(x is None or str(x).strip() == "" for x in r):
                continue
            tarix, doc, voen, name, nov, mebleg, qeyd, h8, h9 = r
            if i == 1 and tarix is not None and norm(tarix).startswith("tarix"):
                continue  # başlıq sətri
            d = parse_date(tarix)
            if not d:
                errors.append((path.name, i, "Tarix düzgün deyil"))
                continue
            k = kind_of(ftype, nov)
            if not k:
                errors.append((path.name, i, f"Növ tanınmadı: {nov!r}"))
                continue
            v = parse_voen(voen)
            if not v:
                errors.append((path.name, i, "VÖEN boşdur"))
                continue
            try:
                amt = parse_amount(mebleg)
            except ValueError as e:
                errors.append((path.name, i, str(e)))
                continue
            if k in ("S", "A") and amt < 0:      # mənfi yazılıbsa = qaytarılma
                k, amt = k + "Q", -amt
            elif k in ("SQ", "AQ"):              # qaytarılma həmişə müsbət saxlanır
                amt = abs(amt)
            part = parse_part(h8)
            # köhnə şablon: 8-ci sütun ƏDV % idisə
            rate_raw = h9 if h9 not in (None, "") else (h8 if isinstance(h8, (int, float)) else None)
            if rate_raw is None:
                rate = VAT_DEFAULT if ftype == "Q" else 0
            else:
                rate = parse_vat(rate_raw)
            rows.append((str(path), i, k, d,
                         "" if doc is None else str(doc).strip(), v, amt,
                         "" if qeyd is None else str(qeyd).strip(),
                         rate, part))
            if name is not None and str(name).strip():
                names[v] = str(name).strip()
    finally:
        wb.close()

    st = path.stat()
    with con:
        con.execute("DELETE FROM txn WHERE src=?", (str(path),))
        con.executemany(
            "INSERT INTO txn(src,row_no,kind,date,doc_no,voen,amount,note,vat,part) VALUES (?,?,?,?,?,?,?,?,?,?)",
            rows)
        con.executemany(
            "INSERT INTO parties(voen,name) VALUES(?,?) "
            "ON CONFLICT(voen) DO UPDATE SET name=excluded.name", names.items())
        con.execute("INSERT OR REPLACE INTO files VALUES(?,?,?)",
                    (str(path), st.st_mtime, st.st_size))
    return len(rows), errors, True


def scan(con, cfg, progress=None):
    """Papkaları yoxlayır: yalnız yeni/dəyişmiş faylları oxuyur, silinmişləri bazadan çıxarır."""
    t0 = time.time()
    files = {}
    for ftype, key in (("Q", "qaime"), ("O", "odenis")):
        d = Path(cfg[key])
        d.mkdir(parents=True, exist_ok=True)
        for p in sorted(d.rglob("*")):
            if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"):
                files[str(p)] = (ftype, p)

    known = {r[0]: (r[1], r[2]) for r in con.execute("SELECT path,mtime,size FROM files")}
    stats = {"yeni": 0, "yenilenen": 0, "kecilen": 0, "silinen": 0, "setir": 0, "xetalar": []}

    for sp in set(known) - set(files):
        with con:
            con.execute("DELETE FROM txn WHERE src=?", (sp,))
            con.execute("DELETE FROM files WHERE path=?", (sp,))
        stats["silinen"] += 1

    for n, (sp, (ftype, p)) in enumerate(files.items(), start=1):
        st = p.stat()
        if known.get(sp) == (st.st_mtime, st.st_size):
            stats["kecilen"] += 1
            continue
        if progress:
            progress(f"Oxunur ({n}/{len(files)}): {p.name}")
        cnt, errs, ok = import_file(con, p, ftype)
        stats["xetalar"].extend(errs)
        if ok:
            stats["yenilenen" if sp in known else "yeni"] += 1
            stats["setir"] += cnt

    changed = stats["yeni"] + stats["yenilenen"] + stats["silinen"]
    stats["dublikat"] = duplicates(con) if changed else 0
    stats["muddet"] = time.time() - t0
    return stats


def duplicates(con):
    """Eyni (növ, VÖEN, tarix, sənəd №, məbləğ) təkrarlanan qrupların sayı."""
    return con.execute(
        "SELECT COUNT(*) FROM (SELECT 1 FROM txn "
        "GROUP BY kind, voen, date, doc_no, amount HAVING COUNT(*)>1)").fetchone()[0]


# ----------------------------------------------------------------------------
# Hesablama
# ----------------------------------------------------------------------------

SIGN = {"S": 1, "SQ": -1, "D": -1, "A": -1, "AQ": 1, "C": 1}  # xalis: müsbət = tərəf bizə borcludur


def balances(con, as_of=None):
    """Hər tərəf üzrə: {növ: [cəm, əsas, ƏDV]} (qəpiklə). Qaytarıldı: [(voen, ad, {növ: [...]})]."""
    data = {}
    for voen, k, amt, vat, part in con.execute(
            "SELECT voen, kind, amount, vat, part FROM txn WHERE (? IS NULL OR date <= ?)", (as_of, as_of)):
        b, v = split_amount(amt, part, vat)
        e = data.setdefault(voen, {}).setdefault(k, [0, 0, 0])
        e[0] += amt; e[1] += b; e[2] += v
    names = dict(con.execute("SELECT voen, name FROM parties"))
    out = [(v, names.get(v) or "", d) for v, d in data.items()]
    out.sort(key=lambda x: (x[1], x[0]))
    return out


def _side(d, plus, minus, pay):
    g = lambda k, i: d.get(k, [0, 0, 0])[i]
    inv, ret, pd = g(plus, 0), g(minus, 0), g(pay, 0)
    q = [g(plus, i) - g(minus, i) - g(pay, i) for i in range(3)]
    return inv, ret, pd, q


def debitor_rows(bal):
    """(VÖEN, ad, satış, satış qaytarılma, daxil olan, qalıq, qalıq əsas, qalıq ƏDV) - qalıq != 0."""
    out = []
    for v, n, d in bal:
        s_, sq, dd, q = _side(d, "S", "SQ", "D")
        if q[0] != 0:
            out.append((v, n, s_, sq, dd, q[0], q[1], q[2]))
    return out


def kreditor_rows(bal):
    """(VÖEN, ad, alış, alış qaytarılma, çıxan, qalıq, qalıq əsas, qalıq ƏDV) - qalıq != 0."""
    out = []
    for v, n, d in bal:
        a, aq, c, q = _side(d, "A", "AQ", "C")
        if q[0] != 0:
            out.append((v, n, a, aq, c, q[0], q[1], q[2]))
    return out


def statement(con, voen, as_of=None):
    """Hərəkət (çıxarış). Hər sətir:
    (tarix, sənəd, əməliyyat,
     qaimə əsas, qaimə ƏDV, qaytarılma əsas, qaytarılma ƏDV, ödəniş əsas, ödəniş ƏDV,
     qalıq əsas, qalıq ƏDV, qeyd)
    Qalıq xalisdir: müsbət = tərəf bizə borcludur (debitor), mənfi = biz borcluyuq (kreditor)."""
    rows = con.execute(
        "SELECT date, doc_no, kind, amount, note, vat, part FROM txn "
        "WHERE voen=? AND (? IS NULL OR date<=?) ORDER BY date, id", (voen, as_of, as_of)).fetchall()
    out, qb, qv = [], 0, 0
    for date, doc, k, amt, note, vat, part in rows:
        b, v = split_amount(amt, part, vat)
        qb += SIGN[k] * b
        qv += SIGN[k] * v
        cells = [0] * 6      # qaimə, qaytarılma, ödəniş  x (əsas, ƏDV)
        pos = {"S": 0, "A": 0, "SQ": 2, "AQ": 2, "D": 4, "C": 4}[k]
        cells[pos], cells[pos + 1] = b, v
        out.append((date, doc, KIND_LABEL[k], *cells, qb, qv, note))
    return out


def fmt_money(q):
    return f"{q / 100:,.2f}".replace(",", " ")


def iso_to_az(iso):
    return dt.date.fromisoformat(iso).strftime("%d.%m.%Y")


# ----------------------------------------------------------------------------
# Excel: şablonlar və ixrac
# ----------------------------------------------------------------------------

FONT = "Arial"
HEAD_FILL = PatternFill("solid", start_color="1F3864")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '#,##0.00;[Red]-#,##0.00;"-"'


def _style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def make_templates(folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    specs = [
        ("sablon_qaime.xlsx", "Qaimə №", "Satış,Alış,Satış qaytarılma,Alış qaytarılma", ["Satış", "Alış"],
         ["Satış = müştəriyə kəsilən qaimə (debitor artır).",
          "Alış = təchizatçıdan alınan qaimə (kreditor artır).",
          "Satış qaytarılma = müştərinin qaytardığı mal (debitor azalır). Alış qaytarılma = təchizatçıya qaytardığımız mal (kreditor azalır).",
          "Qaytarılmada məbləği MÜSBƏT yazın (mənfi yazsanız da proqram özü qaytarılma sayır).",
          "Əsas / ƏDV sütunu: məbləğ əsas məbləğdirsə \"Əsas\", ƏDV-dirsə \"ƏDV\" seçin (bir sənəd = iki sətir).",
          "Bu sütun boşdursa məbləğ ƏDV daxil sayılır və ƏDV % (boşdursa 18) ilə ayrılır."],
         (dt.date(2026, 10, 15), "ƏDV-Q-1042", "1234567890", "Nümunə MMC", "Satış", 1180.00, "nümunə sətir")),
        ("sablon_odenis.xlsx", "Sənəd №", "Daxil olan,Çıxan", ["Daxil olan", "Çıxan"],
         ["Daxil olan = müştəridən bizə daxil olan ödəniş (debitor azalır).",
          "Çıxan = bizdən təchizatçıya çıxan ödəniş (kreditor azalır).",
          "Əsas / ƏDV sütunu: ödənişin əsas hissəsi \"Əsas\", ƏDV hissəsi \"ƏDV\" (iki sətir). Boşdursa hamısı əsas sayılır.",
          "Bank çıxarışının sətirlərini bu şablona köçürüb papkaya atın."],
         (dt.date(2026, 10, 20), "MO-77", "1234567890", "Nümunə MMC", "Daxil olan", 500.00, "nümunə sətir")),
    ]
    for fname, docname, validation, types, notes, example in specs:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        headers = ["Tarix", docname, "VÖEN", "Tərəfin adı", "Növ", "Məbləğ", "Qeyd"]
        headers.append("Əsas / ƏDV")
        headers.append("ƏDV %")
        ws.append(headers)
        _style_header(ws, len(headers))
        for col, w in zip("ABCDEFGHI", (13, 16, 15, 34, 20, 16, 30, 13, 9)):
            ws.column_dimensions[col].width = w
        for r in range(2, 2002):
            ws.cell(row=r, column=1).number_format = "DD.MM.YYYY"
            ws.cell(row=r, column=3).number_format = "@"
            ws.cell(row=r, column=6).number_format = "#,##0.00"
            for c in range(1, 8):
                ws.cell(row=r, column=c).font = Font(name=FONT)
        dv = DataValidation(type="list", formula1=f'"{validation}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add("E2:E5000")
        dv2 = DataValidation(type="list", formula1='"Əsas,ƏDV"', allow_blank=True)
        ws.add_data_validation(dv2)
        dv2.add("H2:H5000")
        ws.freeze_panes = "A2"

        info = wb.create_sheet("Təlimat")
        info.column_dimensions["A"].width = 100
        lines = ["Bu şablonu doldurub müvafiq papkaya qoyun, sonra proqramda \"Papkaları yenilə\" basın.",
                 "Məlumat yalnız \"Data\" vərəqindən oxunur; sütunların sırasını dəyişməyin.",
                 "Məcburi sütunlar: Tarix, VÖEN, Növ, Məbləğ. Tərəfin adı VÖEN üzrə yadda saxlanır.",
                 ""] + notes + ["", "Nümunə sətir (bunu Data vərəqinə köçürməyin):"]
        for ln in lines:
            info.append([ln])
        info.append(["   |   ".join([example[0].strftime("%d.%m.%Y"), example[1], example[2],
                                     example[3], example[4], f"{example[5]:.2f}", example[6]])])
        for row in info.iter_rows():
            for cell in row:
                cell.font = Font(name=FONT, bold=(cell.row == 1))
        wb.save(folder / fname)
    return [folder / s[0] for s in specs]


def _write_table(ws, headers, rows, money_cols, formula_cols=None, text_cols=(), total=True):
    """rows - siyahı; formula_cols = {sütun_nömrəsi: lambda sətir_nömrəsi -> formula}"""
    formula_cols = formula_cols or {}
    ws.append(headers)
    _style_header(ws, len(headers))
    for r_i, row in enumerate(rows, start=2):
        vals = list(row)
        for c, fn in formula_cols.items():
            vals[c - 1] = fn(r_i)
        ws.append(vals)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r_i, column=c)
            cell.font = Font(name=FONT)
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = MONEY_FMT
            if c in text_cols:
                cell.number_format = "@"
    last = len(rows) + 1
    if total and rows:
        tr = last + 1
        ws.cell(row=tr, column=1, value="YEKUN")
        for c in money_cols:
            L = get_column_letter(c)
            ws.cell(row=tr, column=c, value=f"=SUM({L}2:{L}{last})")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = TOTAL_FILL
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"


def export_xlsx(path, con, as_of=None):
    bal = balances(con, as_of)
    wb = Workbook()
    ws = wb.active
    ws.title = "Debitor"
    f = lambda r: tuple(x / 100 if isinstance(x, int) else x for x in r)
    drows = [f(r) for r in debitor_rows(bal)]
    _write_table(ws, ["VÖEN", "Tərəfin adı", "Satış qaimələri", "Satış qaytarılma", "Daxil olan ödənişlər",
                      "Debitor qalıq", "Qalıq əsas", "Qalıq ƏDV"],
                 drows, money_cols=(3, 4, 5, 6, 7, 8), formula_cols={6: lambda r: f"=C{r}-D{r}-E{r}"},
                 text_cols=(1,))
    ws2 = wb.create_sheet("Kreditor")
    krows = [f(r) for r in kreditor_rows(bal)]
    _write_table(ws2, ["VÖEN", "Tərəfin adı", "Alış qaimələri", "Alış qaytarılma", "Çıxan ödənişlər",
                       "Kreditor qalıq", "Qalıq əsas", "Qalıq ƏDV"],
                 krows, money_cols=(3, 4, 5, 6, 7, 8), formula_cols={6: lambda r: f"=C{r}-D{r}-E{r}"},
                 text_cols=(1,))
    for w in (ws, ws2):
        for col, width in zip("ABCDEFGH", (15, 40, 20, 20, 22, 20, 18, 16)):
            w.column_dimensions[col].width = width
    info = wb.create_sheet("Məlumat")
    info.append(["Hesabat tarixi", iso_to_az(as_of) if as_of else "Bütün əməliyyatlar"])
    info.append(["Yaradılma vaxtı", dt.datetime.now().strftime("%d.%m.%Y %H:%M")])
    info.append(["Qeyd", "Qalıq = qaimələr − qaytarılmalar − ödənişlər. Mənfi qalıq: qarşı tərəfə artıq ödəniş edilib."])
    info.column_dimensions["A"].width = 18
    info.column_dimensions["B"].width = 80
    for row in info.iter_rows():
        for cell in row:
            cell.font = Font(name=FONT, bold=(cell.column == 1))
    wb.save(path)


STMT_HEADERS = ["Tarix", "Sənəd №", "Əməliyyat", "E-qaimə əsas", "E-qaimə ƏDV",
                "Qaytarılma e-qaimə əsas", "Qaytarılma e-qaimə ƏDV", "Ödəniş əsas", "Ödəniş ƏDV",
                "Qalıq əsas", "Qalıq ƏDV", "Qeyd"]


def export_statement(path, con, voen, name, as_of=None):
    rows = statement(con, voen, as_of)
    wb = Workbook()
    ws = wb.active
    ws.title = "Çıxarış"
    data = [(iso_to_az(r[0]), r[1], r[2], *[x / 100 for x in r[3:11]], r[11]) for r in rows]
    _write_table(ws, STMT_HEADERS, data, money_cols=tuple(range(4, 12)), total=False)
    n = len(data)
    if n:   # yekun sətri: yalnız hərəkət sütunları (qalıq sütunları yekunlanmır)
        tr = n + 2
        ws.cell(row=tr, column=1, value="YEKUN")
        for c in range(4, 10):
            L = get_column_letter(c)
            ws.cell(row=tr, column=c, value=f"=SUM({L}4:{L}{n + 3})")  # 2 sətir yuxarıdan əlavə olunacaq.number_format = MONEY_FMT
        for c in range(1, len(STMT_HEADERS) + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = TOTAL_FILL
            cell.border = BORDER
    for col, width in zip("ABCDEFGHIJKL", (13, 18, 22, 15, 15, 17, 17, 15, 15, 15, 15, 30)):
        ws.column_dimensions[col].width = width
    ws.insert_rows(1, amount=2)
    ws["A1"] = f"{name}  (VÖEN: {voen})"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws["A2"] = "Tarixə qədər: " + (iso_to_az(as_of) if as_of else "bütün əməliyyatlar")
    ws["A2"].font = Font(name=FONT, italic=True)
    ws.freeze_panes = "A4"
    wb.save(path)


# ----------------------------------------------------------------------------
# Pəncərəli proqram (Tkinter)
# ----------------------------------------------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    cfg = load_cfg()
    con = connect(cfg["db"])
    root = tk.Tk()
    root.title("Debitor-Kreditor")
    root.geometry("1150x680")

    state = {"bal": [], "as_of": None}
    status = tk.StringVar(value="Hazır")
    asof_var = tk.StringVar()
    search_var = tk.StringVar()

    class Table:
        def __init__(self, parent, headers):
            frame = ttk.Frame(parent)
            self.frame = frame
            self.tree = ttk.Treeview(frame, columns=list(range(len(headers))), show="headings")
            sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.rows, self.rev = [], False
            for i, h in enumerate(headers):
                self.tree.heading(i, text=h, command=lambda i=i: self.sort(i))
                self.tree.column(i, width=130 if i != 1 else 360, anchor="e" if i >= 2 else "w")
            self.total_label = None

        def fill(self, rows, money_from=2):
            self.rows = rows
            self.show(rows, money_from)

        def show(self, rows, money_from=2):
            self.tree.delete(*self.tree.get_children())
            for r in rows:
                self.tree.insert("", "end",
                                 values=[fmt_money(x) if i >= money_from else x for i, x in enumerate(r)])

        def sort(self, i):
            self.rev = not self.rev
            self.rows.sort(key=lambda r: r[i], reverse=self.rev)
            apply_filter()

    nb = ttk.Notebook(root)
    deb_t = Table(nb, ["VÖEN", "Tərəfin adı", "Satış qaimələri", "Satış qaytarılma",
                       "Daxil olan ödənişlər", "Debitor qalıq", "Qalıq əsas", "Qalıq ƏDV"])
    kred_t = Table(nb, ["VÖEN", "Tərəfin adı", "Alış qaimələri", "Alış qaytarılma",
                        "Çıxan ödənişlər", "Kreditor qalıq", "Qalıq əsas", "Qalıq ƏDV"])
    nb.add(deb_t.frame, text="  Debitor  ")
    nb.add(kred_t.frame, text="  Kreditor  ")
    totals = tk.StringVar()

    def apply_filter(*_):
        q = norm(search_var.get())
        for t in (deb_t, kred_t):
            rows = [r for r in t.rows if not q or q in norm(r[0]) or q in norm(r[1])]
            t.show(rows)
        d = sum(r[5] for r in deb_t.rows)
        k = sum(r[5] for r in kred_t.rows)
        totals.set(f"Debitor yekun: {fmt_money(d)}      Kreditor yekun: {fmt_money(k)}      "
                   f"Xalis: {fmt_money(d - k)}")

    def refresh():
        txt = asof_var.get().strip()
        as_of = parse_date(txt) if txt else None
        if txt and not as_of:
            messagebox.showerror("Tarix", "Tarixi gün.ay.il formatında yazın, məs. 31.12.2026")
            return
        state["as_of"] = as_of
        bal = balances(con, as_of)
        state["bal"] = bal
        deb_t.fill(debitor_rows(bal))
        kred_t.fill(kreditor_rows(bal))
        apply_filter()

    def do_scan():
        root.config(cursor="watch")
        try:
            def prog(m):
                status.set(m)
                root.update_idletasks()
            s = scan(con, load_cfg(), prog)
        finally:
            root.config(cursor="")
        refresh()
        msg = (f"Yeni fayl: {s['yeni']}, yenilənən: {s['yenilenen']}, dəyişməyən: {s['kecilen']}, "
               f"silinən: {s['silinen']}, oxunan sətir: {s['setir']}  ({s['muddet']:.1f} san)")
        status.set(msg)
        warn = []
        if s["dublikat"]:
            warn.append(f"Diqqət: {s['dublikat']} qrup təkrarlanan əməliyyat var "
                        "(eyni növ, VÖEN, tarix, sənəd № və məbləğ). Eyni qaimə iki faylda ola bilər.")
        if s["xetalar"]:
            lines = [f"{f}, sətir {r}: {m}" for f, r, m in s["xetalar"][:25]]
            more = len(s["xetalar"]) - len(lines)
            warn.append(f"{len(s['xetalar'])} sətir oxunmadı:\n" + "\n".join(lines) +
                        (f"\n... və daha {more}" if more > 0 else ""))
        if warn:
            messagebox.showwarning("Yeniləmə nəticəsi", "\n\n".join(warn))

    def do_export():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"debitor_kreditor_{dt.date.today():%Y-%m-%d}.xlsx")
        if path:
            export_xlsx(path, con, state["as_of"])
            status.set(f"İxrac edildi: {path}")

    def do_templates():
        files = make_templates(BASE / "sablonlar")
        messagebox.showinfo("Şablonlar", "Yaradıldı:\n" + "\n".join(str(f) for f in files))

    def pick_folder(key, title):
        d = filedialog.askdirectory(title=title, initialdir=load_cfg()[key])
        if d:
            c = load_cfg()
            c[key] = d
            save_cfg(c)
            do_scan()

    def open_statement(_event=None):
        tree = deb_t.tree if nb.index(nb.select()) == 0 else kred_t.tree
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0])["values"]
        voen, name = str(vals[0]), str(vals[1])
        rows = statement(con, voen, state["as_of"])
        win = tk.Toplevel(root)
        win.title(f"Çıxarış - {name} ({voen})")
        win.geometry("1000x500")
        t = Table(win, STMT_HEADERS)
        t.frame.pack(fill="both", expand=True)
        win.geometry("1600x500")
        for i in range(3, 11):
            t.tree.column(i, width=110)
        t.tree.column(1, width=110)
        t.tree.column(2, width=150, anchor="w")
        t.tree.column(11, width=180, anchor="w")
        z = lambda x: fmt_money(x) if x else ""
        for r in rows:
            t.tree.insert("", "end", values=[iso_to_az(r[0]), r[1], r[2]] + [z(x) for x in r[3:9]] +
                                            [fmt_money(r[9]), fmt_money(r[10]), r[11]])

        def exp():
            p = filedialog.asksaveasfilename(parent=win, defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             initialfile=f"cixaris_{voen}.xlsx")
            if p:
                export_statement(p, con, voen, name, state["as_of"])

        ttk.Button(win, text="Excelə ixrac", command=exp).pack(pady=6)

    bar = ttk.Frame(root, padding=6)
    bar.pack(fill="x")
    ttk.Button(bar, text="Papkaları yenilə", command=do_scan).pack(side="left", padx=2)
    ttk.Button(bar, text="Excelə ixrac", command=do_export).pack(side="left", padx=2)
    ttk.Button(bar, text="Şablonları yarat", command=do_templates).pack(side="left", padx=2)
    ttk.Button(bar, text="Qaimə papkası…",
               command=lambda: pick_folder("qaime", "Qaimələr papkası")).pack(side="left", padx=2)
    ttk.Button(bar, text="Ödəniş papkası…",
               command=lambda: pick_folder("odenis", "Ödənişlər / çıxarışlar papkası")).pack(side="left", padx=2)
    ttk.Label(bar, text="   Tarixə qalıq:").pack(side="left")
    e = ttk.Entry(bar, textvariable=asof_var, width=11)
    e.pack(side="left")
    e.bind("<Return>", lambda _e: refresh())
    ttk.Button(bar, text="Göstər", command=refresh).pack(side="left", padx=2)
    ttk.Label(bar, text="   Axtar:").pack(side="left")
    ttk.Entry(bar, textvariable=search_var, width=22).pack(side="left")
    search_var.trace_add("write", apply_filter)

    nb.pack(fill="both", expand=True, padx=6)
    ttk.Label(root, textvariable=totals, font=("Arial", 10, "bold")).pack(anchor="w", padx=8, pady=2)
    ttk.Label(root, textvariable=status, relief="sunken", anchor="w").pack(fill="x", side="bottom")
    deb_t.tree.bind("<Double-1>", open_statement)
    kred_t.tree.bind("<Double-1>", open_statement)

    root.after(100, do_scan)
    root.mainloop()


# ----------------------------------------------------------------------------
# Əmr sətri
# ----------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Debitor-Kreditor")
    ap.add_argument("--sablon", action="store_true", help="Excel şablonlarını yarat")
    ap.add_argument("--yenile", action="store_true", help="papkaları oxu")
    ap.add_argument("--export", metavar="FAYL.xlsx", help="debitor/kreditor cədvəllərini Excelə ixrac et")
    ap.add_argument("--tarix", metavar="GG.AA.İİİİ", help="qalıq tarixi (boşdursa hamısı)")
    a = ap.parse_args()

    if not (a.sablon or a.yenile or a.export):
        run_gui()
        return
    cfg = load_cfg()
    if a.sablon:
        for f in make_templates(BASE / "sablonlar"):
            print("Yaradıldı:", f)
    if a.yenile or a.export:
        con = connect(cfg["db"])
        if a.yenile or a.export:
            s = scan(con, cfg, print)
            print(f"Yeni: {s['yeni']}, yenilənən: {s['yenilenen']}, dəyişməyən: {s['kecilen']}, "
                  f"silinən: {s['silinen']}, sətir: {s['setir']}, xəta: {len(s['xetalar'])}, "
                  f"dublikat qrup: {s['dublikat']}, müddət: {s['muddet']:.2f} san")
            for f, r, m in s["xetalar"][:50]:
                print(f"  XƏTA {f}, sətir {r}: {m}")
        if a.export:
            as_of = parse_date(a.tarix) if a.tarix else None
            if a.tarix and not as_of:
                sys.exit("Tarix formatı səhvdir (GG.AA.İİİİ).")
            export_xlsx(a.export, con, as_of)
            print("İxrac edildi:", a.export)


if __name__ == "__main__":
    main()
