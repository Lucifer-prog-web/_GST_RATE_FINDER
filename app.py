import os, re, io, json, glob, hashlib, sqlite3, logging, datetime as dt
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageFilter
from rapidfuzz import process, fuzz

# =====================================================================
# CONFIG
# =====================================================================
APP_VERSION = "2.1.0"
MAX_ROWS = 1000
CBIC_URL = "https://cbic-gst.gov.in/gst-goods-services-rates.html"
DEFAULT_XLSX = os.environ.get("GST_HSN_FILE", "/content/UNIVERSAL_HSN_CODES.xlsx")
DATA_DIR = os.environ.get("GST_DATA_DIR", "/content/gst_data")
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except Exception:
    DATA_DIR = os.path.abspath("gst_data")
    os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "finder.db")
UPLOADED_XLSX = os.path.join(DATA_DIR, "hsn_workbook_uploaded.xlsx")

OFFICIAL_LINKS = [
    ("CBIC - GST goods & services rates", CBIC_URL),
    ("Official GST Portal", "https://www.gst.gov.in"),
    ("CBIC", "https://www.cbic.gov.in"),
]

logging.basicConfig(filename=os.path.join(DATA_DIR, "finder.log"), level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("finder")   # event codes only - never search terms or cell values

NOT_STATED = "Not stated in workbook"
VERIF = "Per uploaded workbook - not checked against the current notification"

# Common-name hints -> standard HS heading/chapter. They only widen the SEARCH: every result still shows the
# official description and the workbook's own rate. A hint is never treated as a classification.
ALIASES = {
    "laptop": ["8471"], "notebook computer": ["8471"], "desktop computer": ["8471"], "computer": ["8471"],
    "mobile phone": ["8517"], "cell phone": ["8517"], "smartphone": ["8517"], "mobile": ["8517"],
    "tv": ["8528"], "led tv": ["8528"], "television": ["8528"],
    "fridge": ["8418"], "refrigerator": ["8418"], "ac": ["8415"], "air conditioner": ["8415"],
    "washing machine": ["8450"], "printer": ["8443"],
    "car": ["8703"], "motorcycle": ["8711"], "scooter": ["8711"], "bicycle": ["8712"], "cycle": ["8712"],
    "tyre": ["4011"], "battery": ["8507"], "power bank": ["8507"], "bulb": ["8539"], "led bulb": ["8539"],
    "shirt": ["61", "62"], "tshirt": ["61"], "t shirt": ["61"], "jeans": ["61", "62"], "trouser": ["61", "62"],
    "dress": ["61", "62"], "garment": ["61", "62"], "clothes": ["61", "62"],
    "shoe": ["64"], "footwear": ["64"], "chappal": ["64"], "slipper": ["64"],
    "pen": ["9608"], "umbrella": ["6601"], "mattress": ["9404"], "furniture": ["94"],
    "medicine": ["30"], "tablet": ["3004"], "medicament": ["30"],
    "rice": ["1006"], "atta": ["1101"], "wheat flour": ["1101"], "maida": ["1101"], "sugar": ["1701"],
    "tea": ["0902"], "coffee": ["0901"], "milk": ["0401"], "curd": ["0403"], "dahi": ["0403"],
    "biscuit": ["1905"], "chocolate": ["1806"], "ice cream": ["2105"], "cement": ["2523"],
    "soap": ["3401"], "shampoo": ["3305"], "perfume": ["3303"], "toothpaste": ["3306"],
    "cigarette": ["2402"], "gold": ["7108"], "diamond": ["7102"], "jewellery": ["7113"], "ring": ["7113"],
    "necklace": ["7113"], "bangle": ["7113"], "earring": ["7113"],
}


# =====================================================================
# TEXT HELPERS
# =====================================================================
def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace(" ", " ")
    return re.sub(r"\s+", " ", s).strip()


def normalize(s):
    s = _clean(s).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    toks = []
    for t in s.split():
        if len(t) > 3 and t.endswith("ies"):
            t = t[:-3] + "y"
        elif len(t) > 3 and t.endswith("s") and not t.endswith("ss"):
            t = t[:-1]
        toks.append(t)
    return " ".join(toks)


def pct(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    return f"{x:g}%"


def confidence(score):
    return "High" if score >= 90 else ("Medium" if score >= 80 else "Low")


def safe_cell(v):
    """Neutralise spreadsheet formula injection in exported cells."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@") and not re.fullmatch(r"-?\d+(\.\d+)?", v):
        return "'" + v
    return v


# =====================================================================
# WORKBOOK PARSER  (CBIC goods rate schedule layout, 8 columns)
#   Schedules | S. No. | HSN / Heading / Tariff item | Description | CGST | SGST/UTGST | IGST | Cess
# =====================================================================
EXCL_RE = re.compile(r"(other\s+than|except|excluding)", re.I)
ANYCH_RE = re.compile(r"any\s+(other\s+)?chapter", re.I)


def _codes_from_text(text):
    codes = []
    for m in re.finditer(r"\b(\d{4})\s+to\s+(\d{4})\b", text):
        a, b = int(m.group(1)), int(m.group(2))
        if 0 <= b - a <= 300:
            codes += [f"{x:04d}" for x in range(a, b + 1)]
        else:
            codes += [m.group(1), m.group(2)]
    text = re.sub(r"\b\d{4}\s+to\s+\d{4}\b", " ", text)
    for part in re.split(r"[,;/]|\bor\b|\band\b", text, flags=re.I):
        cur = ""
        for g in re.findall(r"\d+", part):
            if not cur:
                cur = g
            elif len(g) == 2 and len(cur) in (4, 6):
                cur += g
            else:
                codes.append(cur)
                cur = g
        if cur:
            codes.append(cur)
    # valid HS levels only (2/4/6/8 digits); malformed 3/5/7-digit tokens are dropped, never guessed
    return [c for c in codes if len(c) in (2, 4, 6, 8)]


def parse_hsn_cell(raw):
    """-> (codes, excluded_codes, any_chapter)"""
    raw = _clean(raw)
    any_ch = bool(ANYCH_RE.search(raw))
    m = EXCL_RE.search(raw)
    main, excl = (raw[:m.start()], raw[m.end():]) if m else (raw, "")
    main = ANYCH_RE.sub(" ", main)
    codes = list(dict.fromkeys(_codes_from_text(main)))
    ex = list(dict.fromkeys(_codes_from_text(excl)))
    return codes, ex, any_ch


def _flags(hsn_raw, desc):
    t = f"{hsn_raw} {desc}".lower()
    out = []
    if re.search(r"pre-?\s?packaged|labelled", t):
        out.append("Pre-packaged & labelled")
    if re.search(r"other than|except|excluding", t):
        out.append("Has exclusions")
    if re.search(r"sale value|exceeding|not exceeding|rs\.?\s?\d", t):
        out.append("Value / size limit")
    if "brand" in t:
        out.append("Brand condition")
    if re.search(r"unbranded|not pre-?\s?packaged|other than pre", t):
        out.append("Packaging condition")
    return list(dict.fromkeys(out))


def _rate(v, scale):
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    if pd.isna(x):
        return None
    return round(x * scale, 4)


def parse_workbook(source):
    """source: path or bytes. Returns (entries DataFrame, meta dict). Raises ValueError on bad layout."""
    buf = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    xl = pd.ExcelFile(buf, engine="openpyxl")
    raw = None
    sheet = None
    for sn in xl.sheet_names:
        d = xl.parse(sn, header=None, dtype=object)
        if d.shape[1] >= 8 and d.shape[0] > 20:
            raw, sheet = d, sn
            break
    if raw is None:
        raise ValueError("No sheet with the expected 8-column CBIC rate-schedule layout was found.")
    first = _clean(raw.iloc[0, 0]).lower()
    start = 1 if first.startswith("schedule") else 0
    if start == 0 and not re.fullmatch(r"[ivx]+|nil rate", first):
        raise ValueError("First column should hold schedule names (I, II, ... / Nil rate / Compensation Cess).")
    body = raw.iloc[start:, :8].copy()
    body.columns = ["sched", "sno", "hsn", "desc", "cgst", "sgst", "igst", "cess"]

    igst_num = pd.to_numeric(body["igst"], errors="coerce")
    if igst_num.notna().sum() < 10:
        raise ValueError("Could not find numeric IGST rates in column 7.")
    scale = 100 if igst_num.max() <= 1.0 else 1

    rows, skipped = [], {"omitted": 0, "blank_or_note": 0, "no_rate": 0}
    for i, r in body.iterrows():
        s = _clean(r["sched"])
        desc = _clean(r["desc"]).rstrip('"; ')
        if not s or s.lower().startswith("note") or not desc and not _clean(r["hsn"]):
            skipped["blank_or_note"] += 1
            continue
        is_cess = s.lower().startswith("compensation")
        is_nil = s.lower().startswith("nil")
        igst = _rate(r["igst"], scale)
        if re.fullmatch(r"\[?\s*omitted\s*\]?", desc.lower()):
            skipped["omitted"] += 1
            continue
        if not is_cess and igst is None:
            skipped["no_rate"] += 1
            continue
        sched = "Compensation Cess" if is_cess else ("Nil rate" if is_nil else s.upper())
        sno = _clean(r["sno"])
        hsn_raw = _clean(r["hsn"])
        codes, ex, any_ch = parse_hsn_cell(hsn_raw)
        cgst, sgst = _rate(r["cgst"], scale), _rate(r["sgst"], scale)
        if is_cess:
            kind, tax = "cess", "Compensation cess levy"
            ref = f"Compensation Cess schedule, S. No. {sno}"
        else:
            kind = "rate"
            tax = "Nil rate (CBIC 'Nil rate' list)" if is_nil else ("Nil rate (0%)" if igst == 0 else "Taxable")
            ref = (f"Nil rate schedule, S. No. {sno}" if is_nil else f"Schedule {sched}, S. No. {sno}")
        rows.append(dict(
            entry_id=len(rows), kind=kind, schedule=sched, sno=sno, ref=ref, hsn_raw=hsn_raw,
            codes=codes, excl=ex, any_chapter=any_ch, desc=desc, cgst=cgst, sgst=sgst, igst=igst,
            cess_text=_clean(r["cess"]) if is_cess else "", taxability=tax,
            conditions="; ".join(_flags(hsn_raw, desc)), wb_row=int(i) + 1, desc_norm=normalize(desc)))
    entries = pd.DataFrame(rows)
    if entries.empty:
        raise ValueError("No usable rate entries were found in the workbook.")
    rate = entries[entries.kind == "rate"]
    slabs = rate["igst"].round(3).value_counts().sort_index()
    meta = dict(
        sheet=sheet, scale=scale, n_entries=len(entries), n_rate=len(rate),
        n_cess=int((entries.kind == "cess").sum()), n_nil=int((rate.schedule == "Nil rate").sum()),
        n_no_codes=int(((entries.codes.str.len() == 0) & ~entries.any_chapter).sum()),
        skipped=skipped, slabs={float(k): int(v) for k, v in slabs.items()},
        n12=int((rate.igst == 12).sum()), n28=int((rate.igst == 28).sum()),
        schedules=sorted(rate.schedule.unique().tolist()),
    )
    meta["legacy"] = (meta["n12"] + meta["n28"]) > 0
    return entries, meta


# =====================================================================
# FINDER ENGINE
# =====================================================================
class Finder:
    def __init__(self, entries, meta, label):
        self.entries, self.meta, self.label = entries, meta, label
        meta.setdefault("file", label)
        meta.setdefault("file_mtime", "-")
        meta.setdefault("loaded_at", now_iso())
        self.e = entries.set_index("entry_id", drop=False)
        rate = entries[entries.kind == "rate"]
        self.rate_ids = rate.entry_id.tolist()
        self.rate_norm = rate.desc_norm.tolist()
        self.pairs = {"rate": [], "cess": []}
        for r in entries.itertuples():
            for c in r.codes:
                self.pairs[r.kind].append((c, r.entry_id))
        self.excl = {r.entry_id: r.excl for r in entries.itertuples()}
        from collections import Counter
        freq = Counter()
        for d in self.rate_norm:
            freq.update(t for t in d.split() if len(t) >= 4 and t.isalpha())
        self.freq = freq
        vocab = set(freq)
        for k in ALIASES:
            vocab.update(t for t in normalize(k).split() if len(t) >= 4)
        self.vocab = vocab

    # ---------- HSN ----------
    def hsn_matches(self, q, kind="rate"):
        best = {}
        for c, eid in self.pairs[kind]:
            if q.startswith(c):
                rel = 0 if q == c else 1
            elif c.startswith(q):
                rel = 2
            else:
                continue
            ex = self.excl[eid]
            if rel < 2 and any(q.startswith(x) for x in ex):
                continue
            if rel == 2 and any(c.startswith(x) for x in ex):
                continue
            key = (rel, -len(c))
            if eid not in best or key < best[eid][0]:
                best[eid] = (key, rel, c)
        out = [(eid, v[1], v[2]) for eid, v in best.items()]
        out.sort(key=lambda t: (t[1], -len(t[2]) if t[1] < 2 else 0, t[2], t[0]))
        return out

    def _row(self, eid, **extra):
        e = self.e.loc[eid]
        d = dict(entry_id=int(eid), ref=e.ref, schedule=e.schedule, hsn_raw=e.hsn_raw, desc=e.desc,
                 cgst=e.cgst, sgst=e.sgst, igst=e.igst, taxability=e.taxability,
                 conditions=e.conditions, excl=", ".join(e.excl), wb_row=e.wb_row)
        d.update(extra)
        return d

    def cess_for(self, eids):
        seen, out = set(), []
        for eid in eids:
            for c in self.e.loc[eid].codes:
                if len(c) < 4:
                    continue
                for k, ceid in self.pairs["cess"]:
                    if len(k) >= 4 and (c.startswith(k) or k.startswith(c)) and ceid not in seen:
                        seen.add(ceid)
                        ce = self.e.loc[ceid]
                        out.append(dict(ref=ce.ref, hsn_raw=ce.hsn_raw, desc=ce.desc, cess=ce.cess_text))
        return out

    def residual_ids(self):
        e = self.entries
        m = e[(e.kind == "rate") & e.any_chapter & e.desc.str.contains(r"not specified in", case=False, regex=True)]
        return m.entry_id.tolist()

    def search_hsn(self, digits):
        ms = self.hsn_matches(digits)
        res = dict(mode="HSN", term=digits, suggestion=None, alias_note=None)
        if not any(m[1] < 2 for m in ms) and not ms:
            rid = self.residual_ids()
            if rid:
                rows = [self._row(i, rel=3, matched_code="-", match_basis="Residual entry: goods not specified elsewhere",
                                  confidence="Low", score=50) for i in rid]
                df = pd.DataFrame(rows)
                res.update(rows=df, level="warn", specific=rid, cess=[],
                           verdict=("No specific entry for this HSN in the workbook. The workbook's residual entry for goods "
                                    "'not specified' elsewhere may apply - confirm the goods are not covered by any other "
                                    "schedule entry or the 'Nil rate' list."))
                return res
            res.update(rows=pd.DataFrame(), verdict="No entry for this HSN in the workbook.", level="none",
                       specific=[], cess=[])
            return res
        basis = {0: "Exact HSN match", 1: "Workbook entry covers your HSN (broader heading/chapter)",
                 2: "Sub-entry under the HSN you entered"}
        conf = {0: "High", 1: "Medium", 2: "Low"}
        rows = [self._row(eid, rel=rel, matched_code=c, match_basis=basis[rel], confidence=conf[rel],
                          score={0: 100, 1: 95, 2: 80}[rel]) for eid, rel, c in ms]
        df = pd.DataFrame(rows)
        direct = df[df.rel < 2]
        if not direct.empty:
            top_len = direct.matched_code.str.len().max()
            spec = direct[direct.matched_code.str.len() == top_len]
        else:
            spec = df
        rates = sorted(spec.igst.dropna().unique())
        if len(spec) == 1 and len(direct):
            verdict, level = f"One entry applies in the workbook: {pct(rates[0])} IGST.", "ok"
        elif len(rates) == 1:
            verdict, level = (f"{len(spec)} entries match; all carry the same rate: {pct(rates[0])} IGST. "
                              "Conditions in the descriptions may still differ.", "ok")
        else:
            verdict = (f"{len(spec)} entries match with DIFFERENT rates ({', '.join(pct(x) for x in rates)}). "
                       "Multiple classifications may apply - read each description and confirm.")
            level = "warn"
        if direct.empty:
            verdict = ("Only more specific sub-entries exist under this HSN; the rate depends on the exact "
                       "sub-classification. " + (f"All show {pct(rates[0])}." if len(rates) == 1 else
                                                 f"Rates found: {', '.join(pct(x) for x in rates)}."))
            level = "ok" if len(rates) == 1 else "warn"
        res.update(rows=df, verdict=verdict, level=level, specific=spec.entry_id.tolist(),
                   cess=self.cess_for(spec.entry_id.tolist()[:20]))
        return res

    # ---------- text ----------
    def suggest(self, norm):
        out, changed = [], False
        for t in norm.split():
            if t in self.vocab or len(t) < 4 or not t.isalpha():
                out.append(t)
                continue
            cands = process.extract(t, self.vocab, scorer=fuzz.ratio, score_cutoff=80, limit=8)
            if cands:
                top = max(c[1] for c in cands)
                pick = max((c for c in cands if c[1] >= top - 0.01), key=lambda c: self.freq.get(c[0], 0))
                out.append(pick[0])
                changed = True
            else:
                out.append(t)
        return " ".join(out) if changed else None

    def search_text(self, term, limit=25):
        norm = normalize(term)
        res = dict(mode="Product", term=term, suggestion=None, alias_note=None)
        if not norm:
            res.update(rows=pd.DataFrame(), verdict="Enter a product name or HSN code.", level="none",
                       specific=[], cess=[])
            return res
        best = {}
        for _, sc, pos in process.extract(norm, self.rate_norm, scorer=fuzz.token_set_ratio, limit=60, score_cutoff=72):
            best[self.rate_ids[pos]] = (sc, "Text similarity to official description", False)
        hinted = False
        for k, hsns in ALIASES.items():
            if not re.search(rf"(^| ){re.escape(normalize(k))}( |$)", norm):
                continue
            for h in hsns:
                direct_ms = [m for m in self.hsn_matches(h) if m[1] < 2]
                if not direct_ms:
                    continue
                top_len = max(len(m[2]) for m in direct_ms)
                for eid, rel, c in direct_ms:
                    if len(c) == top_len:
                        hinted = True
                        best[eid] = (99, f"Common-name hint: '{k}' -> HSN {h} (verify)", True)
        if hinted:
            res["alias_note"] = ("Search was widened with common-name hints (marked per result). A hint only helps find "
                                 "candidate entries - it is not an official classification.")
        sug = self.suggest(norm)
        if sug and sug != norm:
            res["suggestion"] = sug
        if not best:
            res.update(rows=pd.DataFrame(), level="none", specific=[], cess=[],
                       verdict="No entry in the workbook matched. No rate is shown because none can be determined. "
                               "Try the HSN code or different wording.")
            return res
        order = sorted(best.items(), key=lambda kv: (not kv[1][2], -kv[1][0], len(self.e.loc[kv[0]].desc)))
        total = len(order)
        order = order[:limit]
        rows = [self._row(eid, rel=None, matched_code=", ".join(self.e.loc[eid].codes[:4]) or "-",
                          match_basis=b, confidence=("Medium" if hint else confidence(sc)), score=round(sc, 1))
                for eid, (sc, b, hint) in order]
        df = pd.DataFrame(rows)
        top = df.score.max()
        grp = df[df.score >= top - 5]
        rates = sorted(grp.igst.dropna().unique())
        heads = {(self.e.loc[i].codes[0][:4] if self.e.loc[i].codes else "-") for i in grp.entry_id}
        if len(grp) == 1 and top >= 90:
            verdict, level = f"One strong candidate: {pct(rates[0])} IGST. Confirm the description fits your product.", "ok"
        elif len(rates) == 1:
            verdict = (f"{len(grp)} candidate entries" + (f" across {len(heads)} HSN headings" if len(heads) > 1 else "")
                       + f" - all at {pct(rates[0])} IGST. Confirm which description fits your product.")
            level = "ok"
        else:
            verdict = (f"Multiple classifications may apply, with DIFFERENT rates ({', '.join(pct(x) for x in rates)}). "
                       "Review the descriptions and confirm the correct classification based on the product's actual "
                       "characteristics. Professional classification advice may be necessary.")
            level = "warn"
        if total > limit:
            verdict += f"  (Showing top {limit} of {total} matches.)"
        res.update(rows=df, verdict=verdict, level=level, specific=grp.entry_id.tolist(),
                   cess=self.cess_for(df.entry_id.tolist()[:10]))
        return res

    def search(self, term):
        t = _clean(term)
        d = re.sub(r"[\s.\-]", "", t)
        if re.fullmatch(r"\d{2,8}", d):
            return self.search_hsn(d)
        if re.fullmatch(r"\d+", d):
            return dict(mode="HSN", term=t, rows=pd.DataFrame(), level="none", specific=[], cess=[],
                        suggestion=None, alias_note=None,
                        verdict="An HSN code must be 2 to 8 digits.")
        return self.search_text(t)


# =====================================================================
# DATA SOURCE
# =====================================================================
def _db_get(key, default=None):
    with db() as c:
        r = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def _db_set(key, value):
    with db() as c:
        c.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                  (key, str(value)))


def find_workbook():
    p = _db_get("data_file")
    if p and os.path.exists(p):
        return p
    for c in [DEFAULT_XLSX] + sorted(glob.glob("/content/*HSN*.xlsx")) + sorted(glob.glob("/content/**/*HSN*.xlsx", recursive=True)):
        if c and os.path.exists(c):
            return c
    return None


@st.cache_resource(show_spinner="Reading the HSN workbook ...")
def load_finder(path, mtime):
    entries, meta = parse_workbook(path)
    meta["file"] = os.path.basename(path)
    meta["file_mtime"] = dt.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
    meta["loaded_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return Finder(entries, meta, f"{meta['file']} (modified {meta['file_mtime']})")


# =====================================================================
# SQLITE (history / favourites / settings)
# =====================================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS history(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, term TEXT, mode TEXT, verdict TEXT, ref TEXT, hsn TEXT,
  description TEXT, igst REAL, cgst REAL, sgst REAL, taxability TEXT, n_results INTEGER, data_version TEXT);
CREATE TABLE IF NOT EXISTS favourites(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ref TEXT, hsn TEXT, description TEXT, added_ts TEXT);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
"""


def db():
    from contextlib import contextmanager

    @contextmanager
    def _cm():
        conn = sqlite3.connect(DB_PATH, timeout=30)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()
    return _cm()


def init_db():
    with db() as c:
        c.executescript(SCHEMA)


def now_iso():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def purge_history():
    days = int(_db_get("retention_days", "90"))
    if days > 0:
        cut = (dt.datetime.now() - dt.timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        with db() as c:
            c.execute("DELETE FROM history WHERE ts < ?", (cut,))


def save_history(F, term, res):
    if _db_get("save_history", "1") != "1":
        return
    df = res["rows"]
    row = {}
    if not df.empty:
        spec = df[df.entry_id.isin(res["specific"])] if res["specific"] else df
        row = (spec if not spec.empty else df).iloc[0].to_dict()
    with db() as c:
        c.execute("""INSERT INTO history(ts,term,mode,verdict,ref,hsn,description,igst,cgst,sgst,taxability,
                     n_results,data_version) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (now_iso(), term, res["mode"], res["verdict"], row.get("ref"), row.get("hsn_raw"),
                   row.get("desc"), row.get("igst"), row.get("cgst"), row.get("sgst"), row.get("taxability"),
                   len(df), F.label))
    log.info("event=search mode=%s n=%d", res["mode"], len(df))


def add_favourite(ref, hsn, desc):
    with db() as c:
        if c.execute("SELECT 1 FROM favourites WHERE ref=?", (ref,)).fetchone():
            return False
        c.execute("INSERT INTO favourites(ref,hsn,description,added_ts) VALUES(?,?,?,?)", (ref, hsn, desc, now_iso()))
    return True


def goto(page, term=None):
    st.session_state["page"] = page
    if term is not None:
        st.session_state["finder_term"] = term
        st.session_state["auto_search"] = True


# =====================================================================
# CALCULATOR
# =====================================================================
def q2(x):
    return x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def calc_gst(amount, rate, cess_pct, inclusive, intra):
    a, r, c = Decimal(str(amount)), Decimal(str(rate)), Decimal(str(cess_pct))
    base = a / (1 + (r + c) / 100) if inclusive else a
    gst, cs = base * r / 100, base * c / 100
    out = {"Taxable value": q2(base)}
    if intra:
        out["CGST"] = q2(gst / 2)
        out["SGST/UTGST"] = q2(gst / 2)
    else:
        out["IGST"] = q2(gst)
    out["Compensation cess"] = q2(cs)
    out["Total GST (incl. cess)"] = q2(gst + cs)
    out["Invoice amount"] = q2(base + gst + cs)
    return out


# =====================================================================
# RESULT EXPORT / BULK
# =====================================================================
EXPORT_COLS = ["HSN code (as in schedule)", "Description of goods", "GST rate (IGST)", "CGST", "SGST/UTGST", "IGST",
               "Compensation cess", "Taxability", "Conditions", "Exclusions", "Legal reference", "Source",
               "Effective date", "Verification status", "Match basis", "Matching confidence"]


def export_frame(F, res):
    df = res["rows"]
    if df.empty:
        return pd.DataFrame(columns=EXPORT_COLS)
    cess = "; ".join(f"{c['cess']} ({c['ref']})" for c in res["cess"]) or "None listed in workbook for this HSN"
    out = pd.DataFrame({
        "HSN code (as in schedule)": df.hsn_raw, "Description of goods": df.desc,
        "GST rate (IGST)": df.igst.map(pct), "CGST": df.cgst.map(pct), "SGST/UTGST": df.sgst.map(pct),
        "IGST": df.igst.map(pct), "Compensation cess": cess, "Taxability": df.taxability,
        "Conditions": df.conditions, "Exclusions": df.excl,
        "Legal reference": df.ref + " (workbook row " + df.wb_row.astype(str) + ")",
        "Source": "CBIC rate schedule workbook: " + F.meta["file"] + " | " + CBIC_URL,
        "Effective date": NOT_STATED, "Verification status": VERIF,
        "Match basis": df.match_basis, "Matching confidence": df.confidence})
    return out


def _safe_df(df):
    """Neutralise formula injection in BOTH cell values and column headers before writing to Excel."""
    out = df.apply(lambda col: col.map(safe_cell))
    out.columns = [safe_cell(str(c)) for c in out.columns]
    return out


def to_xlsx(df, sheet="Results"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        _safe_df(df).to_excel(xw, sheet_name=sheet, index=False)
    return buf.getvalue()


def read_upload(f):
    n = f.name.lower()
    if n.endswith(".csv"):
        df = pd.read_csv(f, dtype=str, keep_default_na=False)
    elif n.endswith(".xlsx"):
        df = pd.read_excel(f, dtype=str, engine="openpyxl").fillna("")
    else:
        raise ValueError("Unsupported format. Upload .xlsx or .csv")
    if df.empty:
        raise ValueError("The file is empty.")
    if len(df) > MAX_ROWS:
        raise ValueError(f"File has {len(df)} rows; the limit is {MAX_ROWS}.")
    return df


def guess_col(cols, keys):
    for c in cols:
        if any(k in str(c).lower() for k in keys):
            return c
    return None


def resolve_row(F, name, hsn):
    h = re.sub(r"\D", "", str(hsn or ""))
    remarks = []
    if str(hsn or "").strip() and not 2 <= len(h) <= 8:
        remarks.append("Invalid HSN format.")
        h = ""
    res = F.search_hsn(h) if h else None
    if (res is None or res["rows"].empty) and str(name or "").strip():
        if h:
            remarks.append("HSN not found in workbook; matched by product name instead.")
        res = F.search_text(str(name))
    if res is None or res["rows"].empty:
        return "Unresolved", None, res, "; ".join(remarks + ["No matching entry in workbook."])
    df = res["rows"]
    spec = df[df.entry_id.isin(res["specific"])]
    rates = spec.igst.dropna().unique()
    top = spec.iloc[0]
    if res["mode"] == "Product" and top.score < 80:
        return "Unresolved", top, res, "; ".join(remarks + ["Match confidence too low; possible: " + top.desc[:80]])
    if len(rates) > 1:
        return "Ambiguous", top, res, "; ".join(remarks + [f"Entries with different rates apply ({', '.join(pct(x) for x in sorted(rates))}); "
                                                           "conditions in descriptions decide."])
    if len(spec) > 1:
        remarks.append(f"{len(spec)} entries/HSN codes match; all carry the same rate.")
    return "Matched", top, res, "; ".join(remarks)


def _blank_record(name, hsn):
    return {"Original product name": name, "Normalized product name": normalize(name), "Original HSN code": hsn,
            "Matched HSN code": "", "Official HSN description": "", "GST rate": None, "CGST rate": None,
            "SGST rate": None, "IGST rate": None, "Cess rate": "", "Taxability classification": "",
            "Official notification or legal reference": "", "Source": "", "Effective date": "",
            "Verification status": "", "Matching confidence": "", "Remarks": "", "Processing status": ""}


def _fill_record(F, rec, name, hsn):
    """Look the row up and fill `rec` in place. Returns (status, resolve_result)."""
    status, top, res, remark = resolve_row(F, name, hsn)
    rec["Processing status"], rec["Remarks"] = status, remark
    if status in ("Matched", "Ambiguous") and top is not None:
        if status == "Matched":
            rec.update({"GST rate": top.igst, "CGST rate": top.cgst, "SGST rate": top.sgst, "IGST rate": top.igst,
                        "Taxability classification": top.taxability})
        rec.update({"Matched HSN code": top.hsn_raw, "Official HSN description": top.desc,
                    "Official notification or legal reference": top.ref,
                    "Source": "CBIC rate schedule workbook (" + F.meta["file"] + ")",
                    "Effective date": NOT_STATED,
                    "Verification status": VERIF if status == "Matched" else "Needs review",
                    "Matching confidence": top.confidence})
        cs = res["cess"]
        rec["Cess rate"] = "; ".join(c["cess"] for c in cs) if cs else "None listed in workbook"
    else:
        rec["Verification status"] = "Unresolved"
    return status, res


def process_bulk(F, df, ncol, hcol, progress=None):
    out, n = [], len(df)
    for i, (_, r) in enumerate(df.iterrows()):
        name = r[ncol] if ncol else ""
        hsn = r[hcol] if hcol else ""
        rec = _blank_record(name, hsn)
        try:
            _fill_record(F, rec, name, hsn)
        except Exception as e:
            rec["Processing status"], rec["Remarks"] = "Error", f"Row could not be processed ({type(e).__name__})."
            log.error("event=row_error type=%s", type(e).__name__)
        out.append(rec)
        if progress and (i % 10 == 0 or i == n - 1):
            progress.progress((i + 1) / n, text=f"Processed {i + 1} of {n} rows")
    return pd.DataFrame(out)


def build_bulk_xlsx(orig, res, fname, extra_note=""):
    out = orig.copy().reset_index(drop=True)
    for c in res.columns:
        out[c if c not in out.columns else c + " (GST tool)"] = res[c].values
    flag = res["Processing status"].isin(["Unresolved", "Ambiguous", "Error"]).values
    cnt = res["Processing status"].value_counts()
    summary = pd.DataFrame({"Item": ["Source file", "Processed at", "Rows", "Matched", "Ambiguous", "Unresolved", "Errors", "Note"],
                            "Value": [fname, now_iso(), len(res), cnt.get("Matched", 0), cnt.get("Ambiguous", 0),
                                      cnt.get("Unresolved", 0), cnt.get("Error", 0),
                                      "Rates are as listed in the uploaded CBIC schedule workbook and have not been "
                                      "checked against the current notification. Original columns are preserved; text "
                                      "starting with = + - @ (cells and headers) is prefixed with an apostrophe for safety."
                                      + ((" " + extra_note) if extra_note else "")]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        _safe_df(out).to_excel(xw, sheet_name="Results", index=False)
        _safe_df(out[flag]).to_excel(xw, sheet_name="Unresolved", index=False)
        summary.to_excel(xw, sheet_name="Processing report", index=False)
    return buf.getvalue()


# =====================================================================
# INVOICE OCR  -  runs locally with Tesseract; the image never leaves this server.
# Public contract (also used by the mobile API, api_template.py):
#   ocr_status()            -> (available: bool, note: str)
#   ocr_languages()         -> ["eng", ...]
#   prep_image(bytes)       -> (img RGB PIL, gray PIL "L")      raises ValueError with a user-safe message
#   run_ocr(gray, lang)     -> lines (list of dicts: text, conf, x0, y0, x1, y1, words)
#   analyse_lines(F, lines) -> DataFrame[Use, Item, HSN, HSN check, Invoice GST %, Read quality, OCR line]
# OCR is only reliable on PRINTED text. Every extracted value must be checked by the user.
# =====================================================================
MAX_IMAGE_BYTES = 15 * 1024 * 1024
MAX_IMAGE_PIXELS = 60_000_000
OCR_MIN_WIDTH = 1800
OCR_MAX_SIDE = 3400
LOW_CONF = 70            # Tesseract word confidence (0-100) below which a row is flagged "Check"
INVOICE_COLS = ["Use", "Item", "HSN", "HSN check", "Invoice GST %", "Read quality", "OCR line"]

_OCR_LANGS = []
_HEIF_DONE = False


def ocr_status():
    """(available, note). Never raises."""
    try:
        import pytesseract
    except ImportError:
        return False, "The Python package 'pytesseract' is not installed (pip install pytesseract)."
    try:
        v = pytesseract.get_tesseract_version()
    except Exception:
        return False, "The Tesseract OCR program was not found (Colab/Ubuntu: apt-get install -y tesseract-ocr)."
    return True, f"Tesseract {v}"


def ocr_languages():
    """Installed Tesseract language codes, 'eng' first. Empty list if OCR is unavailable."""
    global _OCR_LANGS
    if _OCR_LANGS:
        return list(_OCR_LANGS)
    try:
        import pytesseract
        langs = [l for l in pytesseract.get_languages(config="") if l not in ("osd", "equ", "snum")]
    except Exception:
        return []
    langs = sorted(set(langs), key=lambda l: (l != "eng", l))
    _OCR_LANGS = langs
    return list(langs)


def _register_heif():
    """Lets Pillow open iPhone HEIC photos when pillow-heif is installed (optional)."""
    global _HEIF_DONE
    if _HEIF_DONE:
        return
    _HEIF_DONE = True
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except Exception:
        pass


def _otsu(arr):
    hist = np.bincount(arr.ravel(), minlength=256).astype(np.float64)
    p = hist / max(hist.sum(), 1.0)
    omega = np.cumsum(p)
    mu = np.cumsum(p * np.arange(256))
    sigma = (mu[-1] * omega - mu) ** 2 / (omega * (1.0 - omega) + 1e-12)
    return int(np.argmax(sigma))


def _normalise_gray(im):
    """Grayscale + remove uneven lighting / shadows by dividing by an estimate of the paper background."""
    g = ImageOps.grayscale(im)
    w, h = g.size
    small = g.resize((max(8, w // 4), max(8, h // 4)), Image.BILINEAR)
    bg = small.filter(ImageFilter.MaxFilter(9)).filter(ImageFilter.GaussianBlur(radius=6))
    bg = bg.resize((w, h), Image.BILINEAR)
    a = np.asarray(g, dtype=np.float32)
    b = np.maximum(np.asarray(bg, dtype=np.float32), 1.0)
    flat = np.clip(a / b * 255.0, 0, 255).astype(np.uint8)
    return ImageOps.autocontrast(Image.fromarray(flat, "L"), cutoff=1)


def _estimate_skew(gray):
    """Rotation (degrees, PIL convention) that makes text lines horizontal. 0.0 if unsure."""
    w, h = gray.size
    s = 700.0 / w if w > 700 else 1.0
    small = gray.resize((max(1, int(w * s)), max(1, int(h * s))), Image.BILINEAR) if s != 1.0 else gray
    arr = np.asarray(small)
    thr = _otsu(arr)
    ink = np.where(arr <= thr, 0, 255).astype(np.uint8)
    if (ink == 0).mean() < 0.005:
        return 0.0
    base = Image.fromarray(ink, "L")

    def score(angle):
        r = base.rotate(angle, resample=Image.BILINEAR, fillcolor=255)
        prof = (np.asarray(r) < 128).sum(axis=1).astype(np.float64)
        return float(np.sum(np.diff(prof) ** 2))

    best = max(np.arange(-12.0, 12.01, 1.0), key=score)
    best = max(np.arange(best - 1.0, best + 1.01, 0.25), key=score)
    if abs(best) < 0.4 or score(best) < score(0.0) * 1.03:
        return 0.0
    return float(best)


def _line_cover(ink, L):
    """Boolean mask of pixels lying on a vertical straight run of ink at least L pixels long (axis 0)."""
    H, W = ink.shape
    if H < L:
        return np.zeros_like(ink)
    cp = np.zeros((H + 1, W), dtype=np.int32)
    np.cumsum(ink, axis=0, dtype=np.int32, out=cp[1:])
    run = (cp[L:] - cp[:-L]) == L                      # run[s]: ink on rows s .. s+L-1
    ce = np.zeros((run.shape[0] + 1, W), dtype=np.int32)
    np.cumsum(run, axis=0, dtype=np.int32, out=ce[1:])
    ys = np.arange(H)
    hi = np.minimum(ys, run.shape[0] - 1) + 1
    lo = np.maximum(ys - L + 1, 0)
    return (ce[hi] - ce[lo]) > 0


def _erase_rules(gray):
    """Paint out long straight table rules. Left in, Tesseract reads them as stray '|', '1', 'l' characters glued to
    the neighbouring numbers (e.g. HSN 6205 becoming 162052). Letters are far shorter than the rules, so they survive."""
    a = np.asarray(gray)
    H, W = a.shape
    ink = a <= min(255, _otsu(a) + 20)
    v = _line_cover(ink, max(40, int(0.034 * W)))
    h = _line_cover(np.ascontiguousarray(ink.T), max(60, int(0.06 * W))).T
    m = v | h
    for _ in range(2):                                   # widen over anti-aliased edges
        m[:, 1:] |= m[:, :-1]
        m[:, :-1] |= m[:, 1:]
        m[1:, :] |= m[:-1, :]
        m[:-1, :] |= m[1:, :]
    out = a.copy()
    out[m] = 255
    return Image.fromarray(out, "L")


def prep_image(data, rotate=0):
    """Raw image bytes -> (img, gray). img = upright RGB for display; gray = flattened, deskewed, OCR-ready.
    `rotate` = extra clockwise degrees (0/90/180/270) chosen by the user. Raises ValueError (safe to show)."""
    if not data:
        raise ValueError("The image is empty.")
    if len(data) > MAX_IMAGE_BYTES:
        raise ValueError(f"The image is larger than {MAX_IMAGE_BYTES // (1024 * 1024)} MB.")
    _register_heif()
    try:
        probe = Image.open(io.BytesIO(data))
        probe.verify()
        im = Image.open(io.BytesIO(data))
        if im.size[0] * im.size[1] > MAX_IMAGE_PIXELS:
            raise ValueError("The image resolution is too large; please use a smaller photo.")
        try:
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGB")
    except ValueError:
        raise
    except Exception:
        raise ValueError("This file could not be read as an image (supported: JPG, PNG, WEBP, HEIC, BMP, TIFF).")
    if int(rotate) % 360:
        im = im.rotate(-int(rotate), expand=True, fillcolor=(255, 255, 255))
    w, h = im.size
    if min(w, h) < 300:
        raise ValueError("The image is too small to read (it needs at least 300 pixels on its short side).")
    scale = 1.0
    if w < OCR_MIN_WIDTH:
        scale = min(OCR_MIN_WIDTH / w, 3.0)
    if max(w, h) * scale > OCR_MAX_SIDE:
        scale = OCR_MAX_SIDE / max(w, h)
    if abs(scale - 1.0) > 0.05:
        im = im.resize((max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS)
    gray = _normalise_gray(im)
    angle = _estimate_skew(gray)
    if angle:
        im = im.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor=(255, 255, 255))
        gray = gray.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor=255)
    return im, _erase_rules(gray)


_LANG_RE = re.compile(r"^[A-Za-z_]{2,20}(\+[A-Za-z_]{2,20}){0,3}$")


def run_ocr(gray, lang="eng"):
    """OCR with Tesseract (LSTM, one uniform block so table rows stay together). Returns a list of lines,
    each {text, conf, x0, y0, x1, y1, words:[{text, conf, x0, y0, x1, y1}]}, top to bottom."""
    import pytesseract
    if not _LANG_RE.match(str(lang)):
        lang = "eng"
    d = pytesseract.image_to_data(gray, lang=lang, config="--oem 1 --psm 6",
                                  output_type=pytesseract.Output.DICT, timeout=120)
    groups = {}
    for i, t in enumerate(d["text"]):
        t = (t or "").strip()
        try:
            conf = float(d["conf"][i])
        except (TypeError, ValueError):
            conf = -1.0
        if not t or conf < 0:
            continue
        x0, y0 = int(d["left"][i]), int(d["top"][i])
        w = dict(text=t, conf=conf, x0=x0, y0=y0, x1=x0 + int(d["width"][i]), y1=y0 + int(d["height"][i]))
        groups.setdefault((d["block_num"][i], d["par_num"][i], d["line_num"][i]), []).append(w)
    lines = []
    for ws in groups.values():
        ws.sort(key=lambda w: w["x0"])
        lines.append(dict(text=" ".join(w["text"] for w in ws), conf=sum(w["conf"] for w in ws) / len(ws),
                          x0=min(w["x0"] for w in ws), y0=min(w["y0"] for w in ws),
                          x1=max(w["x1"] for w in ws), y1=max(w["y1"] for w in ws), words=ws))
    lines.sort(key=lambda l: (l["y0"], l["x0"]))
    return lines


# ---------------------------------------------------------------- invoice table understanding
_BORDER_RE = re.compile(r"^[|¦\[\]{}_~=—–\\!]{1,8}$")
_NUM_RE = re.compile(r"^(?:₹|Rs\.?|INR)?\d[\d,]*(?:\.\d+)?%?$", re.I)
_PCT_RE = re.compile(r"(\d{1,2}(?:\.\d{1,2})?)\s*%")
_STOP_RE = re.compile(r"^\W*(sub\s*-?\s*total|grand\s*total|total\b|amount\s+(?:in\s+)?words|rupees|round\s*-?\s*off|"
                      r"bank\b|terms\b|declaration|e\s*&\s*o\s*e|tax\s+summary|cgst\b|sgst\b|igst\b|utgst\b|cess\b)", re.I)
_NOISE_RE = re.compile(r"\b(invoice|gstin|gst\s*no|bill\s*to|ship\s*to|buyer|seller|vendor|supplier|customer|date|dated|"
                       r"phone|tel|mobile|mob|e-?mail|www|address|pan|state|place\s+of\s+supply|total|sub\s*total|"
                       r"grand|round|bank|a/c|ifsc|branch|terms|declaration|signature|authori[sz]ed|thank|receipt|"
                       r"cashier|change|cash|upi|paid|balance|discount|description|particulars|qty|quantity)\b", re.I)


def _col_type(text):
    t = re.sub(r"[^a-z%@ ]", " ", str(text).lower())
    t = " ".join(t.split())
    if not t:
        return None
    if re.search(r"\b(hsn|sac)\b", t):
        return "hsn"
    if re.search(r"\b(amt|amount|value|total)\b", t):
        return "amount"
    if re.search(r"\bcgst\b", t):
        return "cgst"
    if re.search(r"\b(sgst|utgst)\b", t):
        return "sgst"
    if re.search(r"\bigst\b", t):
        return "igst"
    if re.search(r"\b(gst|tax)\b", t):
        return "gst"
    if re.search(r"descript|particular|\bitems?\b|product|goods|commodity|name of", t):
        return "desc"
    if re.search(r"\bqty\b|quantity|\bpcs\b|\bnos\b", t):
        return "qty"
    if re.search(r"\b(rate|price|mrp)\b", t):
        return "rate"
    if re.search(r"^(s|sr|sl)( no)?$|^no$|^sn$", t):
        return "sno"
    return None


def _wtext(ws):
    return " ".join(w["text"] for w in ws)


def _mconf(ws):
    return sum(w["conf"] for w in ws) / len(ws) if ws else 0.0


def _letters(s):
    return len(re.findall(r"[A-Za-z]", s or ""))


def _words_from(lines):
    out = []
    for l in lines or []:
        for w in l.get("words", []):
            t = str(w.get("text", "")).strip()
            if not t or _BORDER_RE.match(t):
                continue
            t = t.strip("|¦_")
            if t:
                out.append(dict(w, text=t))
    return out


def _cluster_rows(words):
    """Group words into visual rows by vertical centre (independent of Tesseract's own line split)."""
    if not words:
        return [], 0
    hs = sorted(w["y1"] - w["y0"] for w in words)
    hmed = max(hs[len(hs) // 2], 8)
    rows = []
    for w in sorted(words, key=lambda w: (w["y0"] + w["y1"]) / 2):
        yc = (w["y0"] + w["y1"]) / 2
        if rows and abs(yc - rows[-1]["yc"]) <= 0.6 * hmed:
            rows[-1]["words"].append(w)
            ws = rows[-1]["words"]
            rows[-1]["yc"] = sum((x["y0"] + x["y1"]) / 2 for x in ws) / len(ws)
        else:
            rows.append(dict(words=[w], yc=yc))
    for r in rows:
        r["words"].sort(key=lambda w: w["x0"])
        r["y0"] = min(w["y0"] for w in r["words"])
        r["y1"] = max(w["y1"] for w in r["words"])
        r["text"] = _wtext(r["words"])
    return rows, hmed


_KW_RE = re.compile(r"cgst|sgst|igst|utgst|hsn|sac|qty|quantity|rate|price|mrp|amount|amt|gst|descript\w*|particulars|items?", re.I)


def _split_glued(w):
    """OCR sometimes glues neighbouring header cells into one word ('HSN/SACQty'). Split at the keyword boundary."""
    t = w["text"]
    cuts, prev = [], None
    for m in _KW_RE.finditer(t):
        ty = _col_type(m.group(0))
        if ty and prev and ty != prev and m.start() > 0:
            cuts.append(m.start())
        if ty:
            prev = ty
    if not cuts:
        return [w]
    pts = [0] + cuts + [len(t)]
    span = w["x1"] - w["x0"]
    out = []
    for a, b in zip(pts, pts[1:]):
        out.append(dict(w, text=t[a:b], x0=w["x0"] + round(span * a / len(t)), x1=w["x0"] + round(span * b / len(t))))
    return out


def _header_columns(row, hmed):
    """Header words -> columns. 'Description of Goods' stays ONE column; a new column starts when a different known
    column type begins or when there is a wide gap."""
    cols = []
    for w in [p for w0 in row["words"] for p in _split_glued(w0)]:
        t = _col_type(w["text"])
        if cols:
            cur = cols[-1]
            far = w["x0"] - cur["x1"] > 3.0 * hmed
            if not far and (t is None or cur["type"] is None or t == cur["type"]):
                cur["words"].append(w)
                cur["x1"] = w["x1"]
                if cur["type"] is None:
                    cur["type"] = t
                continue
        cols.append(dict(type=t, words=[w], x0=w["x0"], x1=w["x1"]))
    for c in cols:
        c["text"] = _wtext(c["words"])
        c["pct_header"] = bool(re.search(r"%|rate|@", c["text"].lower()))
    return cols


def _find_header(rows, hmed):
    for i, r in enumerate(rows):
        cols = _header_columns(r, hmed)
        types = {c["type"] for c in cols if c["type"]}
        if "desc" in types and len(types) >= 3:
            return i, cols
    return None, None


def _pct_value(text, header_has_pct):
    text = text or ""
    m = _PCT_RE.search(text)
    if m:
        v = float(m.group(1))
    else:
        m2 = re.fullmatch(r"\s*(\d{1,2}(?:\.\d{1,2})?)\s*", text) if header_has_pct else None
        v = float(m2.group(1)) if m2 else None
    return v if v is not None and 0 <= v <= 50 else None


def _clean_item(s):
    s = re.sub(r"[|¦_~]+", " ", str(s or ""))
    s = re.sub(r"^\s*\d{1,3}[.)]\s+(?=\S)", "", s)
    return " ".join(s.split()).strip(" -:;,.")


def _clean_hsn(text):
    """-> (code, ocr_corrected). Fixes the usual digit look-alikes only when the cell is mostly digits."""
    s = re.sub(r"[\s.\-/]", "", str(text or ""))
    if not s or s.isdigit():
        return s, False
    digits = sum(ch.isdigit() for ch in s)
    t = s.translate(str.maketrans("OoIl|SB", "0011158"))
    if len(t) >= 4 and t.isdigit() and digits / len(s) >= 0.6:
        return t, True
    return s, False


def _hsn_check(F, code, corrected=False):
    if not code:
        return ""
    if not re.fullmatch(r"\d{4}(\d{2}(\d{2})?)?", code):
        return "Invalid HSN format"
    if code.startswith("99"):
        return "Service (SAC) code - not covered by this goods workbook"
    rels = {m[1] for m in F.hsn_matches(code)} if F is not None else set()
    if 0 in rels:
        s = "Exact entry in workbook"
    elif 1 in rels:
        s = "Covered by a broader heading/chapter entry"
    elif 2 in rels:
        s = "Only more specific sub-entries in workbook"
    else:
        s = "Not in workbook"
    return s + (" (OCR-corrected - verify)" if corrected else "")


def _table_items(rows, hi, cols):
    cols = sorted(cols, key=lambda c: c["x0"])
    bounds = [(a["x1"] + b["x0"]) / 2 for a, b in zip(cols, cols[1:])]
    first = {}
    for i, c in enumerate(cols):
        if c["type"] and c["type"] not in first:
            first[c["type"]] = i
    items, prev = [], None
    for r in rows[hi + 1:]:
        by_idx = {}
        for w in r["words"]:
            xc = (w["x0"] + w["x1"]) / 2
            by_idx.setdefault(sum(1 for b in bounds if xc >= b), []).append(w)
        txt = {t: _wtext(by_idx.get(i, [])) for t, i in first.items()}
        desc_ws = by_idx.get(first["desc"], [])
        desc = _wtext(desc_ws)
        if _STOP_RE.search(r["text"]) or _STOP_RE.search(desc):
            break
        numeric = any(re.search(r"\d", txt.get(t, "")) for t in ("qty", "rate", "amount", "hsn", "gst", "cgst", "sgst", "igst"))
        if _letters(desc) >= 3 and numeric:
            ph = {t: cols[first[t]]["pct_header"] for t in ("gst", "cgst", "sgst", "igst") if t in first}
            cg = _pct_value(txt.get("cgst"), ph.get("cgst", False))
            sg = _pct_value(txt.get("sgst"), ph.get("sgst", False))
            ig = _pct_value(txt.get("igst"), ph.get("igst", False))
            gs = _pct_value(txt.get("gst"), ph.get("gst", False))
            if cg is not None and sg is not None:
                rate = cg + sg
            elif ig is not None:
                rate = ig
            elif cg is not None and (gs is None or gs == cg):
                rate = 2 * cg   # intra-state: CGST and SGST are equal halves (SGST header may be unreadable)
            elif sg is not None and gs is None:
                rate = 2 * sg
            else:
                rate = gs
            hws = by_idx.get(first.get("hsn", -1), [])
            code, fixed = "", False
            for hw in hws:                              # first digit-like word wins (a neighbour's digit may spill in)
                c, f = _clean_hsn(hw["text"])
                if len(c) >= 4 and c.isdigit():
                    code, fixed = c, f
                    break
            if not code:
                code, fixed = _clean_hsn(txt.get("hsn", ""))
            if not code and "hsn" in first:
                m = re.search(r"\b(\d{8}|\d{6}|\d{4})\s*$", desc)
                if m:                                   # HSN merged into the description cell
                    code, desc = m.group(1), desc[:m.start()]
            items.append(dict(desc=desc, hsn=code, fixed=fixed, rate=rate, conf=_mconf(desc_ws),
                              line=r["text"], y1=r["y1"]))
        elif items and _letters(desc) >= 2 and not numeric and prev is not None \
                and (r["y0"] - prev["y1"]) < 1.2 * max(prev["y1"] - prev["y0"], 1):
            items[-1]["desc"] += " " + desc            # wrapped description line
            items[-1]["line"] += " / " + r["text"]
        prev = r
    return items


def _line_items(rows):
    """Fallback for receipts / invoices without a recognisable table header: 'name  qty  price  amount' rows."""
    doc = " ".join(r["text"].lower() for r in rows)
    split_tax = "cgst" in doc or "sgst" in doc
    items = []
    for r in rows:
        text = " ".join(r["text"].split())
        if _NOISE_RE.search(text) or _STOP_RE.search(text):
            continue
        text = re.sub(r"^\(?\d{1,3}[.)\]]?\s+(?=[A-Za-z])", "", text)
        toks = text.split()
        k = next((i for i, t in enumerate(toks) if i > 0 and _NUM_RE.match(t)), None)
        if k is None:
            continue
        desc, nums = " ".join(toks[:k]), [t for t in toks[k:] if _NUM_RE.match(t)]
        if _letters(desc) < 3 or not (len(nums) >= 2 or any("." in n for n in nums)):
            continue
        pcts = [float(p) for p in _PCT_RE.findall(text)]
        rate = pcts[0] if len(pcts) == 1 else (pcts[0] + pcts[1] if len(pcts) == 2 and pcts[0] == pcts[1] and split_tax else None)
        m = re.search(r"\b(?:hsn|sac)\b[\s:/\-]*(\d{4,8})", text, re.I)
        code = m.group(1) if m else ""
        items.append(dict(desc=desc, hsn=code, fixed=False, rate=rate if rate is None or rate <= 50 else None,
                          conf=_mconf(r["words"]), line=r["text"], y1=r["y1"]))
    return items


def _analyse(F, lines):
    words = _words_from(lines)
    rows, hmed = _cluster_rows(words)
    info = dict(mode="none", rows=len(rows))
    items = []
    if rows:
        hi, cols = _find_header(rows, hmed)
        if hi is not None:
            items, info["mode"] = _table_items(rows, hi, cols), "table"
        if not items:
            items, info["mode"] = _line_items(rows), "lines"
    recs = []
    for it in items:
        name = _clean_item(it["desc"])
        recs.append({"Use": _letters(name) >= 3, "Item": name, "HSN": it["hsn"],
                     "HSN check": _hsn_check(F, it["hsn"], it["fixed"]),
                     "Invoice GST %": it["rate"], "Read quality": "OK" if it["conf"] >= LOW_CONF else "Check",
                     "OCR line": it["line"][:200]})
    df = pd.DataFrame(recs, columns=INVOICE_COLS)
    df = df.astype({"Use": bool, "Item": str, "HSN": str, "HSN check": str, "Read quality": str, "OCR line": str})
    info["items"] = len(df)
    return df, info


def analyse_lines(F, lines):
    """OCR lines -> editable item table (see contract at the top of this section)."""
    df, info = _analyse(F, lines)
    df.attrs["info"] = info
    return df


def compare_invoice_rate(inv, status, rec, res):
    """Plain-language check of the rate printed on the invoice against the workbook. Never changes the rate shown."""
    if inv is None or pd.isna(inv):
        return "Rate not read from the invoice"
    inv = float(inv)
    if status == "Matched" and rec.get("IGST rate") is not None and not pd.isna(rec.get("IGST rate")):
        wb = float(rec["IGST rate"])
        if abs(wb - inv) < 0.01:
            return f"Consistent: invoice {inv:g}% = workbook {wb:g}%"
        note = f"DIFFERS: invoice {inv:g}% vs workbook {wb:g}%"
        if wb in (12, 28):
            note += " (the workbook still has a pre-22-Sep-2025 slab - check the current notification)"
        return note
    if status == "Ambiguous" and res is not None:
        spec = res["rows"][res["rows"].entry_id.isin(res["specific"])]
        rates = sorted({float(x) for x in spec.igst.dropna()})
        lst = ", ".join(f"{x:g}%" for x in rates)
        if any(abs(x - inv) < 0.01 for x in rates):
            return f"Invoice {inv:g}% is one of the possible workbook rates ({lst})"
        return f"Invoice {inv:g}% is not among the possible workbook rates ({lst}) - verify the classification"
    return "Cannot compare - the item was not resolved"


def invoice_lookup(F, table, progress=None):
    """Look up every ticked row of the (user-corrected) item table. Same record layout as Excel bulk processing,
    plus 'Invoice GST %' and 'Invoice rate check'. Rows are never dropped silently."""
    rows = []
    for _, r in table.iterrows():
        use = r.get("Use")
        use = True if use is None or pd.isna(use) else bool(use)
        item, hsn = str(r.get("Item") or "").strip(), str(r.get("HSN") or "").strip()
        if use and (item or hsn):
            rows.append((item, hsn, r.get("Invoice GST %")))
    out = []
    for i, (item, hsn, inv) in enumerate(rows):
        rec = _blank_record(item, hsn)
        res, status = None, "Error"
        try:
            status, res = _fill_record(F, rec, item, hsn)
        except Exception as e:
            rec["Processing status"], rec["Remarks"] = "Error", f"Row could not be processed ({type(e).__name__})."
            log.error("event=row_error type=%s", type(e).__name__)
        inv = None if inv is None or pd.isna(inv) else float(inv)
        rec["Invoice GST %"] = inv
        rec["Invoice rate check"] = compare_invoice_rate(inv, status, rec, res)
        out.append(rec)
        if progress:
            progress.progress((i + 1) / len(rows), text=f"Looked up {i + 1} of {len(rows)} items")
    return pd.DataFrame(out)


def scan_invoice_images(F, blobs, lang="eng", rotate=0):
    """blobs: [(name, bytes)]. Returns dict(table, pages=[...], errors=[...]). Nothing is written to disk by the app."""
    tables, pages, errors = [], [], []
    for name, data in blobs:
        try:
            orig = Image.open(io.BytesIO(data)).size
        except Exception:
            orig = (0, 0)
        try:
            img, gray = prep_image(data, rotate)
            lines = run_ocr(gray, lang)
        except ValueError as e:
            errors.append(f"{name}: {e}")
            continue
        except Exception as e:
            errors.append(f"{name}: the text could not be read from this image ({type(e).__name__}).")
            log.error("event=ocr_error type=%s", type(e).__name__)
            continue
        df, info = _analyse(F, lines)
        confs = [w["conf"] for l in lines for w in l["words"]]
        thumb = img.copy()
        thumb.thumbnail((1000, 1000))
        pages.append(dict(name=name, orig_size=orig, info=info, n_lines=len(lines),
                          mean_conf=(sum(confs) / len(confs)) if confs else 0.0, thumb=thumb,
                          raw="\n".join(l["text"] for l in lines)))
        tables.append(df)
    table = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame(columns=INVOICE_COLS)
    log.info("event=invoice_scan images=%d items=%d", len(pages), len(table))
    return dict(table=table, pages=pages, errors=errors)


# =====================================================================
# UI
# =====================================================================
def data_banner(F):
    m = F.meta
    if m["legacy"]:
        st.warning(f"**Check the data vintage before relying on any rate.** This workbook still has the older 12% / 28% slabs "
                   f"({m['n12']} entries at 12%, {m['n28']} at 28%). Confirm each rate on the [CBIC rates page]({CBIC_URL}).")
        with st.expander("Why this matters"):
            st.write("From 22 Sep 2025 most goods moved to 5%, 18% or 40% under Notification No. 9/2025-Central Tax (Rate), as "
                     "amended. Rates are shown **exactly as listed in your workbook**, so confirm each one on the CBIC page "
                     "before using it on an invoice or return.")


def rate_flag(igst):
    if igst in (12, 28):
        return "12% / 28% slab - removed for most goods from 22 Sep 2025; likely superseded. Verify."
    return ""


def render_cards(F, df, key, limit=25):
    if df.empty:
        return
    if len(df) > limit:
        st.caption(f"Showing the first {limit} of {len(df)} entries here - the download contains all of them.")
    for r in df.head(limit).itertuples():
        tag = {0: "Exact match", 1: "Covers your HSN", 2: "Sub-entry", 3: "Residual entry"}.get(r.rel, "") if r.rel == r.rel and r.rel is not None else ""
        with st.container(border=True):
            c = st.columns([1.2, 1, 1, 1, 1.4])
            c[0].metric("GST rate", pct(r.igst))
            c[1].metric("CGST", pct(r.cgst))
            c[2].metric("SGST / UTGST", pct(r.sgst))
            c[3].metric("IGST", pct(r.igst))
            cs = F.cess_for([r.entry_id])
            c[4].metric("Compensation cess", cs[0]["cess"] if len(cs) == 1 else ("see cess list" if cs else "None listed"))
            st.markdown(f"**{r.desc[:600]}{' ...' if len(r.desc) > 600 else ''}**")
            st.caption(f"HSN as written in schedule: `{r.hsn_raw[:120] or 'any chapter / not specified'}`"
                       + (f"  |  {tag}" if tag else ""))
            fl = rate_flag(r.igst)
            if fl:
                st.warning(fl, icon="⚠️")
            a, b = st.columns(2)
            a.markdown(f"**Taxability:** {r.taxability}  \n"
                       f"**Conditions:** {r.conditions or 'None detected - read the description'}  \n"
                       f"**Exclusions:** {r.excl or 'None'}  \n"
                       f"**Effective date:** {NOT_STATED}")
            b.markdown(f"**Legal reference:** {r.ref} (workbook row {r.wb_row})  \n"
                       f"**Source:** CBIC rate schedule workbook, [CBIC rates page]({CBIC_URL})  \n"
                       f"**Verification:** {VERIF}  \n"
                       f"**Match:** {r.match_basis} - confidence **{r.confidence}**")
            if st.button("Add to favourites", key=f"{key}_fav_{r.entry_id}"):
                st.toast("Added to favourites." if add_favourite(r.ref, r.hsn_raw, r.desc) else "Already in favourites.")


def _inv_clear():
    """Wipe everything held for the current invoice (called by the Clear button). Also resets the file widgets."""
    nonce = st.session_state.get("inv_nonce", 0)
    for k in [k for k in st.session_state if str(k).startswith("inv_") and k != "inv_nonce"]:
        del st.session_state[k]
    st.session_state["inv_nonce"] = nonce + 1


def page_invoice(F):
    st.title("Invoice Scan")
    data_banner(F)
    ok, note = ocr_status()
    if not ok:
        st.error("Reading text from images is not available on this server. " + note)
        st.code("!apt-get install -y tesseract-ocr\n!pip install pytesseract")
        return
    st.info("Take a photo of a **printed** invoice or bill, or upload one from your gallery. The app reads the item lines, "
            "you check them, and each item is then looked up in your HSN workbook. Best results: flat page, good light, "
            "sharp and upright photo. Handwriting is not supported.")
    with st.expander("Privacy - please read before uploading"):
        st.markdown(
            "- The picture is read **on the machine running this app** (Tesseract OCR). It is **not sent to any AI or OCR "
            "service**.\n"
            "- The app does **not save** the image or the text it reads: no history entry, no log of contents, no file "
            "kept (Tesseract briefly uses a temporary file that is deleted straight after reading).\n"
            "- What you see stays only in **your browser session** until you press *Clear invoice data* or close the tab.\n"
            "- Invoices usually carry business names, addresses, GSTINs and phone numbers, sometimes of individuals. "
            "Upload only invoices you are entitled to process. On Google Colab the link is public and traffic passes "
            "through Cloudflare - do not share the link and do not upload invoices with personal data you are not "
            "authorised to handle.")

    nonce = st.session_state.setdefault("inv_nonce", 0)
    src = st.radio("Image source", ["Upload / gallery", "Take a photo"], horizontal=True, key="inv_src")
    blobs = []
    if src == "Upload / gallery":
        files = st.file_uploader("Invoice image(s) - JPG, PNG, WEBP or HEIC. Up to 5 pages.",
                                 type=["jpg", "jpeg", "png", "webp", "heic", "heif", "bmp", "tif", "tiff"],
                                 accept_multiple_files=True, key=f"inv_up_{nonce}")
        for f in (files or [])[:5]:
            blobs.append((f.name, f.getvalue()))
        if files and len(files) > 5:
            st.warning("Only the first 5 images are used.")
        st.caption("On a phone, this button also lets you pick the camera app, which usually gives a sharper picture "
                   "than the in-page camera below.")
    else:
        shot = st.camera_input("Point the camera at the invoice and take the photo", key=f"inv_cam_{nonce}")
        if shot:
            blobs.append(("camera photo", shot.getvalue()))

    langs = ocr_languages() or ["eng"]
    opts = langs + [f"eng+{l}" for l in langs if l != "eng"] if "eng" in langs else langs
    c1, c2 = st.columns(2)
    lang = c1.selectbox("Language of the invoice text", opts, key="inv_lang") if len(opts) > 1 else opts[0]
    rot = c2.selectbox("Rotate the image if the text is sideways", [0, 90, 180, 270], key="inv_rot",
                       format_func=lambda d: f"{d} degrees")
    if not blobs:
        st.caption("No image yet.")
        return

    sig = hashlib.sha256(b"".join(b for _, b in blobs) + f"|{lang}|{rot}".encode()).hexdigest()
    if st.session_state.get("inv_sig") != sig:
        for k in ("inv_res", "inv_scan"):
            st.session_state.pop(k, None)
        with st.spinner("Reading the invoice ..."):
            st.session_state["inv_scan"] = scan_invoice_images(F, blobs, lang, rot)
        st.session_state["inv_sig"] = sig
    scan = st.session_state["inv_scan"]
    st.button("Clear invoice data", on_click=_inv_clear)

    for e in scan["errors"]:
        st.error(e)
    for p in scan["pages"]:
        w = p["orig_size"][0]
        if w and w < 1000:
            st.warning(f"{p['name']}: the picture is small ({w} px wide). Small photos are read less accurately - "
                       "retake it closer or at higher quality.")
        if p["n_lines"] == 0:
            st.warning(f"{p['name']}: no text was found. Check the photo and the rotation setting.")
        elif p["mean_conf"] < 70:
            st.warning(f"{p['name']}: the text is hard to read (average confidence {p['mean_conf']:.0f}/100). The photo "
                       "may be blurry, dim or at an angle. Check every row carefully.")
    if scan["pages"]:
        with st.expander("Show the picture(s) and the raw text that was read"):
            for p in scan["pages"]:
                a, b = st.columns(2)
                a.image(p["thumb"], caption=p["name"])
                b.text(p["raw"][:4000] or "(no text)")

    table = scan["table"]
    st.subheader("1. Check the items read from the invoice")
    if table.empty and not scan["errors"]:
        st.warning("No item lines were recognised. Try a sharper photo or the rotation setting, or add the items by hand "
                   "in the table below (click the empty row at the bottom).")
    st.caption("OCR can misread letters and digits. Correct the names and HSN codes, untick rows that are not items, "
               "and add anything that is missing. 'Invoice GST %' is the rate printed on the invoice (if it was found); "
               "it is only used to cross-check, never to set the rate.")
    edited = st.data_editor(
        table, key=f"inv_editor_{sig[:10]}", num_rows="dynamic", hide_index=True,
        disabled=["HSN check", "Read quality", "OCR line"],
        column_config={
            "Use": st.column_config.CheckboxColumn("Use", default=True, width="small"),
            "Item": st.column_config.TextColumn("Item / description", width="large"),
            "HSN": st.column_config.TextColumn("HSN", width="small"),
            "HSN check": st.column_config.TextColumn("HSN check", width="medium"),
            "Invoice GST %": st.column_config.NumberColumn("Invoice GST %", min_value=0, max_value=100, format="%g"),
            "Read quality": st.column_config.TextColumn("Read quality", width="small"),
            "OCR line": st.column_config.TextColumn("As read from the image", width="large"),
        })

    st.subheader("2. Look up the GST rates")
    if st.button("Find GST for the ticked items", type="primary", key="inv_go"):
        n = int(sum(1 for _, r in edited.iterrows()
                    if (r.get("Use") is None or pd.isna(r.get("Use")) or bool(r.get("Use")))
                    and (str(r.get("Item") or "").strip() or str(r.get("HSN") or "").strip())))
        if n == 0:
            st.error("Tick at least one item (or add one) first.")
        else:
            bar = st.progress(0.0, text="Starting ...")
            res = invoice_lookup(F, edited, bar)
            bar.empty()
            kept = edited[[bool(True if (u is None or pd.isna(u)) else u) and bool(str(i or "").strip() or str(h or "").strip())
                           for u, i, h in zip(edited["Use"], edited["Item"], edited["HSN"])]]
            st.session_state["inv_res"] = dict(sig=sig, res=res, items=kept[["Item", "HSN", "HSN check", "Read quality", "OCR line"]])
    R = st.session_state.get("inv_res")
    if not R or R["sig"] != sig:
        return
    res = R["res"]
    k = res["Processing status"].value_counts()
    m = st.columns(4)
    m[0].metric("Matched", int(k.get("Matched", 0)))
    m[1].metric("Ambiguous - needs your call", int(k.get("Ambiguous", 0)))
    m[2].metric("Unresolved", int(k.get("Unresolved", 0)))
    diff = int(res["Invoice rate check"].str.startswith("DIFFERS").sum() + res["Invoice rate check"].str.contains("not among").sum())
    m[3].metric("Rate differs from invoice", diff)
    st.caption("Rates are exactly as listed in your workbook. 'Ambiguous' means several entries with different rates fit the "
               "description - pick the item in section 3 below to read the conditions. Edit the table above and press the button again to refresh.")
    show = res[["Original product name", "Matched HSN code", "IGST rate", "CGST rate", "SGST rate", "Cess rate",
                "Processing status", "Invoice GST %", "Invoice rate check", "Remarks"]].rename(
        columns={"Original product name": "Item", "IGST rate": "GST rate (IGST)"})
    st.dataframe(show, hide_index=True)
    st.download_button("Download Excel (items + GST details)",
                       build_bulk_xlsx(R["items"], res, "invoice scan",
                                       "Items were read from an invoice image by OCR and may have been corrected by the user. "
                                       "'Invoice GST %' is the rate printed on the invoice and is used only to cross-check."),
                       file_name="gst_invoice_scan.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("3. Full details for an item")
    st.caption("Conditions, legal references and cess for the item you pick. Shown here - you do not need to leave this page. "
               "These lookups are not saved to your search history.")
    sel = st.selectbox("Item", list(range(len(res))), key="inv_open_sel",
                       format_func=lambda i: f"{i + 1}. {str(res.iloc[i]['Original product name'])[:60] or res.iloc[i]['Original HSN code']}")
    row = res.iloc[sel]
    code = re.sub(r"\D", "", str(row["Original HSN code"]))
    term = code if 4 <= len(code) <= 8 else str(row["Original product name"])
    if not term.strip():
        st.info("This row has no item name or HSN code to look up.")
        return
    one = F.search(term)
    st.caption(f"Looked up: **{term}**  ({one['mode']} search)")
    if one.get("suggestion"):
        st.info(f"Did you mean **{one['suggestion']}**? Correct the item in the table above and press the button again.")
    {"ok": st.success, "warn": st.warning, "none": st.error}[one["level"]](one["verdict"])
    inv = row.get("Invoice GST %")
    if inv is not None and not pd.isna(inv):
        st.caption("Invoice check: " + str(row["Invoice rate check"]))
    render_search_result(F, one, kp="invd", downloads=False)


def page_home(F):
    st.title("GST Rate Finder")
    data_banner(F)
    m = F.meta
    c = st.columns(4)
    c[0].metric("Rate entries loaded", m["n_rate"])
    c[1].metric("Cess entries", m["n_cess"])
    c[2].metric("Workbook", m["file"])
    c[3].metric("File modified", m["file_mtime"])
    st.subheader("Search")
    st.text_input("Product name or HSN code", key="home_q", placeholder="e.g. 8471, refrigerator, wheat flour")
    st.button("Find GST rate", type="primary",
              on_click=lambda: goto("Rate Finder", st.session_state.get("home_q", "").strip())
              if st.session_state.get("home_q", "").strip() else None)
    st.button("Scan an invoice (camera or gallery)", on_click=goto, args=("Invoice Scan",))
    left, right = st.columns(2)
    with left:
        st.subheader("Rate slabs in the workbook (IGST)")
        s = pd.Series(m["slabs"])
        st.bar_chart(pd.DataFrame({"entries": s.values}, index=[f"{k:g}%" for k in s.index]))
    with right:
        with db() as cn:
            recent = pd.read_sql_query("SELECT ts, term, hsn, igst, verdict FROM history ORDER BY id DESC LIMIT 8", cn)
            favs = pd.read_sql_query("SELECT hsn, description FROM favourites ORDER BY id DESC LIMIT 6", cn)
        st.subheader("Recent searches")
        if recent.empty:
            st.caption("No searches yet.")
        else:
            st.dataframe(recent, hide_index=True)
        st.subheader("Favourites")
        if favs.empty:
            st.caption("No favourites yet.")
        else:
            st.dataframe(favs, hide_index=True)
    cols = st.columns(len(OFFICIAL_LINKS))
    for col, (lbl, url) in zip(cols, OFFICIAL_LINKS):
        col.link_button(lbl, url)


def render_search_result(F, res, kp="f", downloads=True):
    """Full detail view of a search result: cess, best-match cards, broader entries. kp = widget-key prefix."""
    if res["rows"].empty:
        return
    st.caption("An entry applies only to goods that match its **description** and conditions - a chapter-level HSN "
               "match alone does not mean the whole chapter carries that rate.")
    ex = export_frame(F, res) if downloads else None
    if downloads:
        d1, d2 = st.columns(2)
        d1.download_button("Download results (Excel)", to_xlsx(ex), file_name="gst_rate_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("Download results (CSV)", ex.to_csv(index=False).encode("utf-8-sig"), file_name="gst_rate_results.csv")
    if res["cess"]:
        with st.expander(f"Compensation cess entries related to these HSN codes ({len(res['cess'])})", expanded=True):
            st.caption("Cess is a separate levy and is NOT included in the GST rate above.")
            st.dataframe(pd.DataFrame(res["cess"]).rename(columns={"ref": "Legal reference", "hsn_raw": "HSN",
                                                                  "desc": "Description", "cess": "Cess (as listed)"}),
                         hide_index=True)
    df = res["rows"]
    if res["mode"] == "HSN":
        spec_mask = df.entry_id.isin(res["specific"])
        st.subheader("Best match" if (df.rel < 2).any() else "Entries under this HSN")
        render_cards(F, df[spec_mask], kp + "s")
        sub = df[(df.rel == 2) & ~spec_mask]
        broad = df[(df.rel < 2) & ~spec_mask]
        if len(sub):
            with st.expander(f"More specific entries under this HSN ({len(sub)})"):
                render_cards(F, sub, kp + "sub")
        if len(broad):
            with st.expander(f"Broader chapter/heading entries that also mention this HSN ({len(broad)}) - the description decides"):
                render_cards(F, broad, kp + "b")
    else:
        render_cards(F, df, kp)


def page_finder(F):
    st.title("GST Rate Finder")
    data_banner(F)
    st.text_input("Product name, common name or HSN code (full or partial)", key="finder_term",
                  placeholder="e.g. 8507 60 00  |  lithium-ion battery  |  refrigerator")
    go = st.button("Find GST rate", type="primary")
    auto = st.session_state.pop("auto_search", False)
    term = st.session_state.get("finder_term", "").strip()
    if (go or auto) and term:
        res = F.search(term)
        save_history(F, term, res)
        st.session_state["last"] = {"term": term, "res": res}
    L = st.session_state.get("last")
    if not L:
        st.info("Type a product or HSN code and press **Find GST rate**.")
        return
    res = L["res"]
    st.caption(f"Searched for: **{L['term']}**  ({res['mode']} search)")
    if res.get("suggestion"):
        st.info(f"Did you mean **{res['suggestion']}**? Your original search was not changed.")
        st.button("Search the suggestion", on_click=goto, args=("Rate Finder", res["suggestion"]))
    if res.get("alias_note"):
        st.caption(res["alias_note"])
    {"ok": st.success, "warn": st.warning, "none": st.error}[res["level"]](res["verdict"])
    render_search_result(F, res)
    st.session_state["calc_rates"] = res["rows"][["igst", "desc"]].dropna().head(15).values.tolist()


def page_calc(F):
    st.title("GST Calculator")
    data_banner(F)
    amount = st.number_input("Amount (INR)", min_value=0.0, value=1000.0, step=100.0, format="%.2f")
    opts = ["Enter manually"] + (["Use a rate from my last search"] if st.session_state.get("calc_rates") else [])
    src = st.radio("Rate source", opts, horizontal=True)
    rate, cess = 0.0, 0.0
    if src == "Enter manually":
        c1, c2 = st.columns(2)
        rate = c1.number_input("GST rate (%)", min_value=0.0, max_value=100.0, value=18.0, step=0.5)
        cess = c2.number_input("Compensation cess (% ad valorem)", min_value=0.0, max_value=500.0, value=0.0, step=0.5)
    else:
        lst = st.session_state["calc_rates"]
        lab = [f"{r[0]:g}% - {r[1][:90]}" for r in lst]
        pick = st.selectbox("Entry", lab)
        rate = float(lst[lab.index(pick)][0])
        cess = st.number_input("Compensation cess (% ad valorem, if any)", min_value=0.0, max_value=500.0, value=0.0, step=0.5)
    mode = st.radio("Amount is", ["Exclusive of GST", "Inclusive of GST"], horizontal=True)
    supply = st.selectbox("Type of supply (not assumed)", ["Intra-state (CGST + SGST)", "Inter-state (IGST)"],
                          index=None, placeholder="Select place-of-supply outcome")
    if st.button("Calculate", type="primary"):
        if supply is None:
            st.error("Select intra-state or inter-state. The tool does not assume place of supply.")
            return
        inc = mode.startswith("Inclusive")
        res = calc_gst(amount, rate, cess, inc, supply.startswith("Intra"))
        st.table(pd.DataFrame({"Component": list(res.keys()), "Amount (INR)": [f"{v:,.2f}" for v in res.values()]}))
        st.code(("Taxable value = Amount / (1 + (GST% + Cess%)/100)" if inc else "Taxable value = Amount")
                + f"\nGST = Taxable x {rate:g}%   Cess = Taxable x {cess:g}%   Invoice = Taxable + GST + Cess")
        st.caption("Ad valorem cess only; specific (per-unit) cess is not supported. Rounded half-up to 2 decimals.")


def page_bulk(F):
    st.title("Excel Bulk Processing")
    data_banner(F)
    st.caption(f"Up to {MAX_ROWS} rows. Processed in memory; nothing is sent to external services.")
    f = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "csv"], key="bulk_up")
    if not f:
        return
    try:
        data = read_upload(f)
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return
    st.write(f"**File:** {f.name}  |  **Rows:** {len(data)}")
    st.dataframe(data.head(10))
    cols = ["(none)"] + list(data.columns)
    gn, gh = guess_col(data.columns, ["product", "item", "name", "description"]), guess_col(data.columns, ["hsn", "sac"])
    c1, c2 = st.columns(2)
    nc = c1.selectbox("Product name column", cols, index=cols.index(gn) if gn else 0)
    hc = c2.selectbox("HSN column", cols, index=cols.index(gh) if gh else 0)
    if nc == "(none)" and hc == "(none)":
        st.info("Select at least a product-name or HSN column.")
        return
    if st.button("Process file", type="primary"):
        bar = st.progress(0.0, text="Starting ...")
        res = process_bulk(F, data, None if nc == "(none)" else nc, None if hc == "(none)" else hc, bar)
        st.session_state["bulk"] = {"orig": data, "res": res, "name": f.name}
    B = st.session_state.get("bulk")
    if B and B["name"] == f.name:
        res = B["res"]
        k = res["Processing status"].value_counts()
        m = st.columns(4)
        m[0].metric("Matched", int(k.get("Matched", 0)))
        m[1].metric("Ambiguous", int(k.get("Ambiguous", 0)))
        m[2].metric("Unresolved", int(k.get("Unresolved", 0)))
        m[3].metric("Errors", int(k.get("Error", 0)))
        st.dataframe(res)
        st.download_button("Download Excel output", build_bulk_xlsx(B["orig"], res, B["name"]),
                           file_name="gst_bulk_results.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_history(F):
    st.title("Search History")
    with db() as c:
        h = pd.read_sql_query("SELECT * FROM history ORDER BY id DESC", c)
    if h.empty:
        st.info("No search history.")
        return
    q = st.text_input("Search within history")
    v = h[h.apply(lambda r: q.lower() in (str(r["term"]) + str(r["description"]) + str(r["hsn"])).lower(), axis=1)] if q else h
    st.dataframe(v, hide_index=True)
    if v.empty:
        return
    st.caption("Each row is the result as retrieved at that time. Compare it with what the current workbook says below.")
    sel = st.selectbox("Select entry", v["id"].tolist(),
                       format_func=lambda i: f"#{i} | {v[v.id == i].iloc[0]['ts']} | {v[v.id == i].iloc[0]['term']}")
    row = v[v.id == sel].iloc[0]
    c1, c2 = st.columns(2)
    c1.button("Reopen search", on_click=goto, args=("Rate Finder", row["term"]))
    if c2.button("Delete this entry"):
        with db() as c:
            c.execute("DELETE FROM history WHERE id=?", (int(sel),))
        st.rerun()
    if row["ref"]:
        cur = F.entries[F.entries.ref == row["ref"]]
        a, b = st.columns(2)
        a.markdown(f"**Saved ({row['ts']})**  \nIGST {pct(row['igst'])} | data: {row['data_version']}")
        if cur.empty:
            b.markdown("**Current workbook**  \nEntry not found in the current workbook.")
        else:
            b.markdown(f"**Current workbook**  \nIGST {pct(cur.iloc[0].igst)} | {F.label}")
    d1, d2 = st.columns(2)
    d1.download_button("Export history (Excel)", to_xlsx(h, "History"), file_name="gst_search_history.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        ok = st.checkbox("Confirm clear ALL history")
        if st.button("Clear all history", disabled=not ok):
            with db() as c:
                c.execute("DELETE FROM history")
            st.rerun()


def page_fav(F):
    st.title("Favourites")
    with db() as c:
        f = pd.read_sql_query("SELECT * FROM favourites ORDER BY id DESC", c)
    if f.empty:
        st.info("No favourites yet. Add them from the Rate Finder results.")
        return
    st.dataframe(f, hide_index=True)
    sel = st.selectbox("Select favourite", f["id"].tolist(),
                       format_func=lambda i: f"{f[f.id == i].iloc[0]['ref']} - {f[f.id == i].iloc[0]['description'][:70]}")
    row = f[f.id == sel].iloc[0]
    c1, c2 = st.columns(2)
    first = (row["hsn"] or "").split(",")[0].split(" ")[0]
    c1.button("Open in finder", on_click=goto, args=("Rate Finder", first if first.isdigit() else row["description"][:60]))
    if c2.button("Remove"):
        with db() as c:
            c.execute("DELETE FROM favourites WHERE id=?", (int(sel),))
        st.rerun()


def page_data(F):
    st.title("Data & Sources")
    m = F.meta
    data_banner(F)
    st.subheader("Connected workbook")
    st.write(f"**File:** {m['file']}  |  **Modified:** {m['file_mtime']}  |  **Sheet:** {m['sheet']}  |  **Loaded:** {m['loaded_at']}")
    c = st.columns(4)
    c[0].metric("GST rate entries", m["n_rate"])
    c[1].metric("of which 'Nil rate' list", m["n_nil"])
    c[2].metric("Cess entries", m["n_cess"])
    c[3].metric("Rows skipped ([Omitted]/blank)", m["skipped"]["omitted"] + m["skipped"]["blank_or_note"])
    st.write("**Schedules found:** " + ", ".join(m["schedules"]))
    st.write("**IGST slabs in workbook:** " + ", ".join(f"{k:g}% ({v})" for k, v in m["slabs"].items()))
    st.caption("This is the CBIC goods rate schedule: it lists scheduled entries, not every HSN code. An HSN with no "
               "entry (and no covering heading/chapter entry) returns 'no entry' - the tool never guesses a rate. "
               "Services (SAC) are not covered. The workbook carries no effective dates or notification numbers, "
               "so those fields say 'Not stated in workbook'.")
    st.subheader("Replace the workbook")
    st.write("Upload a newer CBIC goods rate schedule in the **same 8-column layout** "
             "(Schedules, S. No., HSN, Description, CGST, SGST/UTGST, IGST, Cess). It is validated before it is used.")
    up = st.file_uploader("New workbook (.xlsx)", type=["xlsx"], key="wb_up")
    if up and st.button("Validate and use this workbook", type="primary"):
        try:
            data = up.getvalue()
            _, meta = parse_workbook(data)
            with open(UPLOADED_XLSX, "wb") as fh:
                fh.write(data)
            _db_set("data_file", UPLOADED_XLSX)
            load_finder.clear()
            st.success(f"Loaded {meta['n_rate']} rate entries. Reloading ...")
            st.rerun()
        except Exception as e:
            st.error(f"Workbook rejected: {e}")
    st.subheader("Official references")
    for lbl, url in OFFICIAL_LINKS:
        st.link_button(lbl, url)


def page_about(F):
    st.title("About & Privacy")
    st.write(f"GST Rate Finder v{APP_VERSION}. A reference tool - not legal or tax advice. Fuzzy and common-name matching "
             "only *suggest* candidate entries; the rate shown is always the one written in the workbook entry.")
    st.markdown("""
- **Stored locally:** search history (term, matched entry, rate, time), favourites, settings.
- **Uploaded files** are processed in memory (a replacement workbook is saved to the session so it can be reused).
- **Invoice photos** are read locally with Tesseract OCR. The image and the text read from it are not saved, not put in the
  search history and not written to logs (logs keep only counts). They stay in your browser session until you clear them.
- **Logs** hold event codes only - never search terms or cell values. No telemetry or external AI calls.
""")
    st.warning("On Google Colab the data lives on a temporary VM and the public link works for anyone who has it (no "
               "login). Do not share the link, and stop the runtime when finished.")
    st.subheader("Settings")
    sv = st.toggle("Save search history", value=_db_get("save_history", "1") == "1")
    days = st.number_input("Auto-delete history older than (days, 0 = keep)", 0, 3650, int(_db_get("retention_days", "90")))
    if st.button("Save settings"):
        _db_set("save_history", "1" if sv else "0")
        _db_set("retention_days", days)
        purge_history()
        st.success("Saved.")
    ok = st.checkbox("I understand wiping local data cannot be undone")
    if st.button("Delete history and favourites", disabled=not ok):
        with db() as c:
            c.execute("DELETE FROM history")
            c.execute("DELETE FROM favourites")
        st.success("Deleted.")


PAGES = {"Home": page_home, "Rate Finder": page_finder, "GST Calculator": page_calc,
         "Excel Bulk Processing": page_bulk, "Invoice Scan": page_invoice, "Search History": page_history, "Favourites": page_fav,
         "Data & Sources": page_data, "About & Privacy": page_about}


RESPONSIVE_CSS = """<style>
.st-key-topnav { display: none; }
@media (max-width: 820px) {
  [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"],
  [data-testid="stSidebarCollapseButton"] { display: none !important; }
  .st-key-topnav { display: block; }
  .block-container { padding: 3.2rem 0.85rem 2rem 0.85rem !important; max-width: 100% !important; }
  h1 { font-size: 1.6rem !important; line-height: 1.25 !important; }
  h2, h3 { font-size: 1.2rem !important; }
  [data-testid="stMetricValue"] { font-size: 1.5rem !important; }
  [data-testid="stMetricLabel"] { font-size: 0.8rem !important; }
  .stButton, .stDownloadButton { width: 100% !important; }
  .stButton > button, .stDownloadButton > button { min-height: 2.9rem; width: 100%; }
  [data-testid="stSelectbox"] div[data-baseweb="select"] > div { min-height: 2.9rem; }
  input, textarea { font-size: 16px !important; }
  [data-testid="stFileUploaderDropzone"] { padding: 0.8rem; }
  [data-testid="stCameraInput"] video, [data-testid="stCameraInput"] img { max-width: 100% !important; height: auto !important; }
  [data-testid="stDataFrame"], [data-testid="stDataEditor"] { max-width: 100%; }
  [data-testid="stAlert"] { font-size: 0.92rem; }
}
@media (max-width: 480px) {
  [data-testid="stHorizontalBlock"] { gap: 0.4rem !important; }
}
</style>"""


def _nav_top_changed():
    st.session_state["page"] = st.session_state["nav_top"]


def main():
    st.set_page_config(page_title="GST Rate Finder", page_icon="🧾", layout="wide")
    init_db()
    if "purged" not in st.session_state:
        purge_history()
        st.session_state["purged"] = True
    st.session_state.setdefault("page", "Home")
    st.markdown(RESPONSIVE_CSS, unsafe_allow_html=True)
    # Small screens: the sidebar is hidden by CSS and this compact menu is shown instead (hidden on wide screens).
    st.session_state["nav_top"] = st.session_state.get("page", "Home")
    with st.container(key="topnav"):
        st.markdown("**🧾 GST Rate Finder**")
        st.selectbox("Menu - go to", list(PAGES), key="nav_top", on_change=_nav_top_changed)
    with st.sidebar:
        st.markdown("## 🧾 GST Rate Finder")
        st.radio("Navigation", list(PAGES), key="page")
        st.caption("No login. Local storage only.")
    path = find_workbook()
    F = None
    if path:
        try:
            F = load_finder(path, os.path.getmtime(path))
        except Exception as e:
            st.error(f"The HSN workbook could not be read: {e}")
    if F is None:
        st.title("Connect your HSN workbook")
        st.write("Upload `UNIVERSAL_HSN_CODES.xlsx` (CBIC goods rate schedule layout).")
        up = st.file_uploader("HSN workbook (.xlsx)", type=["xlsx"], key="first_up")
        if up and st.button("Use this workbook", type="primary"):
            try:
                data = up.getvalue()
                parse_workbook(data)
                with open(UPLOADED_XLSX, "wb") as fh:
                    fh.write(data)
                _db_set("data_file", UPLOADED_XLSX)
                load_finder.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Workbook rejected: {e}")
        return
    PAGES[st.session_state["page"]](F)


if __name__ == "__main__":
    main()
