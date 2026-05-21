import os
import platform
import subprocess
import uuid
import re
import copy
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


LETTER = r"[^\W\d_]"   # alla Unicode-bokstäver, inga siffror/underscore

# -------------------------------
# Autocomplete för Lärosäte
# -------------------------------
INSTITUTIONS = [
    "Högskolan i Gävle - HiG",
    "Karlstads universitet - KAU",
    "Mittuniversitetet - MIUN",
    "Luleå tekniska universitet - LTU",
    "Högskolan Dalarna - HD",
    "Kungliga Tekniska högskolan - KTH",
    "FEI",
    "Linköpings universitet - LiU",
    "Uppsala universitet",
    "Lunds universitet",
    "Göteborgs universitet",
    "Umeå universitet",
    "Linnéuniversitetet",
    "Jönköping University",
    "Högskolan i Halmstad",
    "Högskolan i Borås",
    "Högskolan Kristianstad",
    "Högskolan i Skövde",
]

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, master=None, **kwargs):
        self._all_values = kwargs.pop("values", [])
        self._case_sensitive = kwargs.pop("case_sensitive", False)
        super().__init__(master, values=self._all_values, **kwargs)
        self.bind("<KeyRelease>", self._on_keyrelease)

    def _filter(self, text: str):
        if not text:
            return self._all_values
        if self._case_sensitive:
            return [v for v in self._all_values if v.startswith(text)]
        t = text.lower()
        return [v for v in self._all_values if v.lower().startswith(t)]

    def _on_keyrelease(self, event):
        cur = self.get()
        new_vals = self._filter(cur)
        self["values"] = new_vals
        if event.keysym not in ("BackSpace", "Left", "Right", "Up", "Down", "Escape"):
            try:
                self.event_generate("<Down>")
            except Exception:
                pass

# -------------------------------
# Datamodell
# -------------------------------
@dataclass
class ExamInfo:
    lärosäte: str
    kurs: str
    kurskod: str
    tid: str
    senastbörja: str
    betala: str          
    hjälpmedel: str
    tahemtenta: str         
    tentander: List[str]
    anonymkod: str
    övriginfo: str

# -------------------------------
# Hjälpfunktioner
# -------------------------------

RE_PNR = re.compile(r"(\d{8}-\d{4})\s*$")
RE_LEADING_NO = re.compile(r"^\s*\d+\s+")
RE_ANON = re.compile(r"^\s*\d-\s*\d{4}-[A-ZÅÄÖ]{3}\s+", re.IGNORECASE)

SKIP_WORDS = (
    "deltagarlista", "tentamen i", "kurskod", "provkod", "antal anmälda",
    "tentasal", "lokal", "datum", "tid", "anm.", "sida"
)

def parse_hig_line(line: str) -> str | None:
    s = (line or "").strip()
    if not s:
        return None

    low = s.lower()
    if any(w in low for w in SKIP_WORDS):
        return None

    m = RE_PNR.search(s)
    if not m:
        return None
    pnr = m.group(1)

    left = s[:m.start()].strip()              # allt före personnumret
    left = RE_LEADING_NO.sub("", left).strip()# ta bort platsnumret

    # ta bort anonymkod om den finns (OMG401-varianten)
    left = RE_ANON.sub("", left).strip()

    name = re.sub(r"\s{2,}", " ", left).strip()
    if not name:
        return None

    # vill du bara visa namn i listbox: returnera name
    return name

def _append_first_n_pages(writer: PdfWriter, pdf_path: str, n: int):
    """Lägg till de första n sidorna från pdf_path i en befintlig PdfWriter."""
    if not pdf_path or not os.path.exists(pdf_path) or n <= 0:
        return
    rdr = PdfReader(pdf_path)
    total = len(rdr.pages)
    count = min(max(0, n), total)
    for i in range(count):
        writer.add_page(rdr.pages[i])

def _append_all_pages(writer, pdf_path: str) -> None:
    """Lägg till ALLA sidor från en PDF till writer."""
    rdr = PdfReader(pdf_path)
    for pg in rdr.pages:
        writer.add_page(copy.copy(pg))


def extract_names_kau(pdf_path: str, page_from: int = 1, page_to: int | None = None) -> list[str]:
    if not pdf_path:
        return []
    names: list[str] = []
    import re, pdfplumber
    DASH = r"[-–—]"
    anon_pat = re.compile(rf"[A-ZÅÄÖ0-9]{{3,10}}{DASH}\d{{4}}{DASH}[A-Z0-9ÅÄÖ]{{2,8}}")

    def clean_norm(n: str) -> str:
        n = (n or "").replace("\u00a0", " ").strip()
        n = re.sub(r"^\s*\d+\s+", "", n)
        n = re.sub(r"\b\d{6,12}(-\d{4})?\b", "", n)
        n = " ".join(n.split())
        if "," in n:
            last, rest = [p.strip() for p in n.split(",", 1)]
            return f"{rest} {last}".strip()
        return n

    with pdfplumber.open(pdf_path) as pdf:
        last = len(pdf.pages)
        p1 = max(1, page_from)
        p2 = min(last, page_to or last)
        for idx, page in enumerate(pdf.pages, start=1):
            if not (p1 <= idx <= p2):
                continue
            text = page.extract_text() or ""
            for raw in text.splitlines():
                line = " ".join((raw or "").split())
                if not line:
                    continue
                m = anon_pat.search(line)
                if not m:
                    continue
                left = clean_norm(line[:m.start()].strip())
                right = clean_norm(line[m.end():].strip())
                cand = left if left else right
                if cand and cand not in names:
                    names.append(cand)
    return names

def _extract_names_from_table(pdf_path: str, page_from: int = 1, page_to: int | None = None) -> list[str]:
    import pdfplumber, re
    names: list[str] = []

    def _norm(n: str) -> str:
        n = (n or "").replace("\u00a0", " ").strip()
        n = re.sub(r"^\s*\d+\s+", "", n)
        n = re.sub(r"\b\d{6,12}(-\d{4})?\b", "", n)
        n = " ".join(n.split())
        if not n: return ""
        if "," in n:
            last, rest = [p.strip() for p in n.split(",", 1)]
            return f"{rest} {last}".strip()
        return n

    def grab_with(settings):
        out = []
        with pdfplumber.open(pdf_path) as pdf:
            last = len(pdf.pages)
            p1 = max(1, page_from)
            p2 = min(last, page_to or last)
            for idx, page in enumerate(pdf.pages, start=1):
                if not (p1 <= idx <= p2):
                    continue
                try:
                    tables = page.extract_tables(table_settings=settings) or []
                except Exception:
                    tables = []
                for tbl in tables:
                    if not tbl: continue
                    ntbl = [[(c or "").strip() for c in row] for row in tbl]
                    maxw = max(len(r) for r in ntbl)
                    merged = [""] * maxw
                    namn_col = None
                    header_row = 0
                    for i in range(min(3, len(ntbl))):
                        row = ntbl[i] + [""] * (maxw - len(ntbl[i]))
                        for j, c in enumerate(row):
                            merged[j] = (merged[j] + " " + c).strip().lower()
                        # exakt "namn"/"name", uteslut 'provnamn'/'test name'
                        if any(re.search(r"\bnamn\b", c) or re.search(r"\bname\b", c) for c in merged):
                            for j, c in enumerate(merged):
                                if (re.search(r"\bnamn\b", c) or re.search(r"\bname\b", c)) and not re.search(r"provnamn|test name", c):
                                    namn_col = j
                                    break
                            header_row = i
                            break
                    if namn_col is None:
                        # fallback: kolumn som mest liknar namn
                        scores = []
                        for j in range(maxw):
                            vals = []
                            for r in ntbl[header_row+1:]:
                                r = r + [""] * (maxw - len(r))
                                vals.append(r[j].strip())
                            score = sum(1 for v in vals if len(_norm(v).split()) >= 2 and not re.search(r"\d{3,}", v))
                            scores.append((score, j))
                        scores.sort(reverse=True)
                        if scores and scores[0][0] >= 2:
                            namn_col = scores[0][1]
                    if namn_col is None:
                        continue
                    for r in ntbl[header_row+1:]:
                        r = r + [""] * (maxw - len(r))
                        nm = _norm(r[namn_col])
                        if nm and nm not in out:
                            out.append(nm)
        return out

    # 1) linjer
    names = grab_with({
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 5,
        "snap_tolerance": 3,
        "edge_min_length": 30,
        "min_words_vertical": 1,
        "join_tolerance": 2,
        "text_tolerance": 2,
    })
    if names:
        return names
    # 2) text
    names = grab_with({
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "intersection_tolerance": 5,
        "snap_tolerance": 3,
        "min_words_vertical": 1,
        "text_x_tolerance": 2,
        "text_y_tolerance": 2,
        "join_tolerance": 2,
    })
    return names

def extract_names_miun_simple(pdf_path: str, page_from: int = 1, page_to: int | None = None) -> list[str]:
    """
    MIUN: plocka 'Efternamn, Förnamn' mellan anonymkoden 'A-0001-OFR' och födelsedatum 'YYMMDD'.
    Ex: '1 A-0001-OFR Gül, Alpaslan 890228' -> 'Alpaslan Gül'
    """
    if not pdf_path or not os.path.exists(pdf_path):
        return []

    import re, pdfplumber

    # plats  anonymkod (A-0001-OFR)   NAMN         YYMMDD
    pat = re.compile(
        r"^\s*\d+\s+[A-ZÅÄÖ]\s*[-–—]\s*\d{4}\s*[-–—]\s*[A-ZÅÄÖ]{3,6}\s+(?P<name>.+?)\s+\d{6}\s*$"
    )

    def norm(n: str) -> str:
        n = (n or "").replace("\u00a0", " ").strip()
        n = re.sub(r"\b\d{6,12}(-\d{4})?\b", "", n)  # pnr bort
        n = " ".join(n.split())
        if "," in n:
            last, rest = [p.strip() for p in n.split(",", 1)]
            return f"{rest} {last}".strip()
        return n

    names: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        last = len(pdf.pages)
        p1 = max(1, page_from)
        p2 = min(last, page_to or last)
        for idx in range(p1, p2 + 1):
            page = pdf.pages[idx - 1]
            text = page.extract_text() or ""
            for raw in text.splitlines():
                line = " ".join((raw or "").split())
                m = pat.match(line)
                if not m:
                    continue
                nm = norm(m.group("name"))
                if nm and nm not in names:
                    names.append(nm)
    return names

def extract_names_hig_chars(pdf_path: str, page_from: int = 1, page_to: int | None = None) -> list[str]:
    """
    HiG: Tecken-baserad rekonstruktion av rader.
    Klarar två strukturer:
      A) plats + anonymkod + 'Efternamn, Förnamn' + personnummer
      B) plats + (ev 'Efternamn, Förnamn' eller 'Förnamn Efternamn') + personnummer
    Robust för Unicode (t.ex. ü), flerdelade namn och '...' efter förnamn.
    """
    if not pdf_path or not os.path.exists(pdf_path):
        return []

    import re, statistics, pdfplumber

    # Unicode-bokstäver (utan siffror/underscore)
    LETTER = r"[^\W\d_]"
    NAMEPART = rf"{LETTER}+(?:[-']{LETTER}+)*"  # ex: O'Neill, Jean-Claude
    DASH = r"[-–—]"

    # Personnummer: YYYYMMDD-XXXX (ibland utan bindestreck)
    PNR = rf"(?P<pnr>\d{{8}}{DASH}?\d{{4}})\b"

    # Anonymkod: 5- 0004-HHH (tillåt whitespace runt streck)
    ANON_HIG = rf"(?P<anon>\d{{1,2}}\s*{DASH}\s*\d{{4}}\s*{DASH}\s*[A-ZÅÄÖ]{{2,8}})\b"

    # Strikt A: plats + anonymkod + Efternamn, Förnamn + pnr
    STRICT_A = re.compile(
        rf"^\s*\d+\s+{ANON_HIG}\s+"
        rf"(?P<last>{NAMEPART}(?:\s+{NAMEPART})*)\s*,\s*"
        rf"(?P<first>{NAMEPART}(?:\s+{NAMEPART})*)"
        rf"(?:\.{{3,}}|…)?\s+{PNR}\s*$",
        re.UNICODE
    )

    # Strikt B: plats + (namn) + pnr, där namn helst är "Efternamn, Förnamn"
    # (Vi tillåter även utan kommatecken och försöker tolka sista ordet som efternamn i fallback)
    STRICT_B_COMMA = re.compile(
        rf"^\s*\d+\s+"
        rf"(?P<last>{NAMEPART}(?:\s+{NAMEPART})*)\s*,\s*"
        rf"(?P<first>{NAMEPART}(?:\s+{NAMEPART})*)"
        rf"(?:\.{{3,}}|…)?\s+{PNR}\s*$",
        re.UNICODE
    )

    # Super-robust fallback B: plats + (några ord) + pnr (utan krav på kommatecken)
    STRICT_B_FLEX = re.compile(
        rf"^\s*\d+\s+(?P<name>.+?)\s+{PNR}\s*$",
        re.UNICODE
    )

    # Rader som vi aldrig vill plocka som "namn" (rubriker/metadata)
    SKIP_HINTS = (
        "tentamen", "tenta", "antal", "anmäl", "deltagar", "prov", "kurs",
        "datum", "tid", "sal", "lokal", "sida", "institution", "norrtälje",
        "individuell", "salstentamen"
    )

    def norm_display(last: str, first: str) -> str:
        # Städa ellips och NBSP
        last = (last or "").replace("\u00a0", " ").strip()
        first = (first or "").replace("\u00a0", " ").strip()
        last = re.sub(r"(?:\.{3,}|…)+", "", last).strip()
        first = re.sub(r"(?:\.{3,}|…)+", "", first).strip()
        full = f"{first} {last}".strip()
        full = re.sub(r"\s{2,}", " ", full)
        return full

    def looks_like_name(n: str) -> bool:
        # Minst två "ord", inga siffror, inte rubrik-ord
        if not n or re.search(r"\d", n):
            return False
        low = n.lower()
        if any(h in low for h in SKIP_HINTS):
            return False
        parts = n.split()
        return len(parts) >= 2

    names: list[str] = []
    seen = set()

    with pdfplumber.open(pdf_path) as pdf:
        last_page = len(pdf.pages)
        p1 = max(1, page_from)
        p2 = min(last_page, page_to or last_page)

        for idx in range(p1, p2 + 1):
            page = pdf.pages[idx - 1]
            chars = page.chars or []
            if not chars:
                continue

            # Grupp efter Y (rad) med tolerans – funkar bra för HiG
            Y_TOL = 3.5
            rows: dict[int, list[dict]] = {}
            for ch in chars:
                ykey = round(ch["top"] / Y_TOL)
                rows.setdefault(ykey, []).append(ch)

            for ykey in sorted(rows.keys()):
                line_chars = sorted(rows[ykey], key=lambda c: c["x0"])

                # Beräkna typiskt teckenavstånd för dynamisk tröskel
                gaps = [(b["x0"] - a["x1"]) for a, b in zip(line_chars, line_chars[1:])]
                med_gap = statistics.median(gaps) if gaps else 2.0
                WORD_GAP = max(3.0, med_gap * 2.2)

                # Bygg raden: lägg mellanslag bara vid "stora" gap
                rebuilt = []
                prev = None
                for ch in line_chars:
                    t = ch.get("text", "")
                    if not t:
                        continue
                    if prev is not None:
                        gap = ch["x0"] - prev["x1"]
                        if gap > WORD_GAP:
                            rebuilt.append(" ")
                    rebuilt.append(t)
                    prev = ch

                line = "".join(rebuilt).replace("\u00a0", " ").strip()
                line = re.sub(r"\s{2,}", " ", line)
                if not line:
                    continue

                # Sanera ellips i namndel (Christine... / Christine…)
                line = re.sub(r"(\S)\s*(?:\.{3,}|…)\b", r"\1", line)

                # 1) Försök strict A (med anonymkod)
                m = STRICT_A.match(line)
                if m:
                    nm = norm_display(m.group("last"), m.group("first"))
                    if looks_like_name(nm) and nm not in seen:
                        seen.add(nm)
                        names.append(nm)
                    continue

                # 2) Försök strict B med kommatecken (utan anonymkod)
                m = STRICT_B_COMMA.match(line)
                if m:
                    nm = norm_display(m.group("last"), m.group("first"))
                    if looks_like_name(nm) and nm not in seen:
                        seen.add(nm)
                        names.append(nm)
                    continue

                # 3) Flex-fallback: plats + name + pnr (utan krav på komma)
                m = STRICT_B_FLEX.match(line)
                if not m:
                    continue

                # Filtrera bort uppenbara icke-namn innan vi ens normaliserar
                name_raw = (m.group("name") or "").strip()
                if not name_raw:
                    continue
                low = name_raw.lower()
                if any(h in low for h in SKIP_HINTS):
                    continue

                # Om det råkar börja med anonymkod ändå (pga spacing), ta bort
                name_raw = re.sub(rf"^{ANON_HIG}\s+", "", name_raw, flags=re.UNICODE).strip()

                # Om det finns komma, använd det
                if "," in name_raw:
                    last, first = [p.strip() for p in name_raw.split(",", 1)]
                    nm = norm_display(last, first)
                else:
                    # Fallback: tolka sista ordet som efternamn, resten som förnamn
                    parts = name_raw.split()
                    if len(parts) < 2:
                        continue
                    first = " ".join(parts[:-1])
                    last = parts[-1]
                    nm = norm_display(last, first)

                if looks_like_name(nm) and nm not in seen:
                    seen.add(nm)
                    names.append(nm)

    return names

def extract_names_miun_strict(pdf_path: str, page_from: int = 1, page_to: int | None = None) -> list[str]:
    if not pdf_path or not os.path.exists(pdf_path):
        return []
    import re, pdfplumber
    DASH = r"[-–—]"
    anon_code = rf"\b\d{{4}}\s*{DASH}\s*[A-ZÅÄÖ0-9]{{2,8}}\s*{DASH}\s*\d\b"
    pat = re.compile(
        rf"(?P<dob>\b\d{{6}}\b)\s*(?P<name>[A-Za-zÅÄÖåäö ,.'\-]+?)\s*(?P<anon>{anon_code})",
        re.DOTALL
    )

    def norm(n: str) -> str:
        n = (n or "").replace("\u00a0", " ").strip()
        n = re.sub(r"\b\d{6,12}(-\d{4})?\b", "", n)
        n = " ".join(n.split())
        if "," in n:
            last, rest = [p.strip() for p in n.split(",", 1)]
            return f"{rest} {last}".strip()
        return n

    names: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        last = len(pdf.pages)
        p1 = max(1, page_from)
        p2 = min(last, page_to or last)
        for idx, page in enumerate(pdf.pages, start=1):
            if not (p1 <= idx <= p2):
                continue
            text = (page.extract_text() or "").replace("\r", "\n").replace("\u00a0", " ")
            for m in pat.finditer(text):
                nm = norm(m.group("name"))
                if nm and len(nm.split()) >= 2 and not re.search(r"\d{3,}", nm):
                    if nm not in names:
                        names.append(nm)
    return names

def extract_names_miun(pdf_path: str) -> list[str]:
    return _extract_names_from_table(pdf_path)

def extract_names_hig(pdf_path: str) -> list[str]:
    return _extract_names_from_table(pdf_path)

def extract_names_by_institution(
    pdf_path: str,
    lärosäte: str,
    page_from: int = 1,
    page_to: int | None = None
) -> list[str]:
    s = (lärosäte or "").strip().lower()
    if not s:
        return []

    # --- KAU ---
    if "karlstad" in s:
        return extract_names_kau(pdf_path, page_from, page_to)

    # --- MIUN ---
    if "mittuniversitetet" in s or "miun" in s:
        # 1) MIUN-radparser (ingen tabell, inga rubriker)
        names = extract_names_miun_simple(pdf_path, page_from, page_to)
        if names:
            return names

        # 2) tabell
        names = _extract_names_from_table(pdf_path, page_from, page_to)
        if names:
            return names

        # 3) fallback
        return extract_names_kau(pdf_path, page_from, page_to)

    # --- HiG (Högskolan i Gävle) ---
    if ("gävle" in s) or ("gavle" in s) or ("högskolan i gävle" in s) or ("hig" in s):
        # 1) tecken-baserad layoutparser (klarar båda HiG-varianterna)
        names = extract_names_hig_chars(pdf_path, page_from, page_to)
        if names:
            return names

        # 2) tabell (om någon PDF råkar vara tabell)
        names = _extract_names_from_table(pdf_path, page_from, page_to)
        if names:
            return names

        # 3) fallback
        return extract_names_kau(pdf_path, page_from, page_to)

    # --- Okänd/övrig: försök tabell, annars KAU-fallback ---
    names = _extract_names_from_table(pdf_path, page_from, page_to)
    if names:
        return names

    return extract_names_kau(pdf_path, page_from, page_to)


def make_cover_sheet_pdf(out_path: str, info: ExamInfo) -> None:
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT

    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm
    )

    styles = getSampleStyleSheet()
    story = []

    # Rubrik
    title_style = ParagraphStyle(
        'CustomTitle', parent=styles['Title'],
        fontName='Helvetica-Bold', fontSize=30, leading=34, alignment=1
    )
    story.append(Paragraph(f"{info.lärosäte}", title_style))
    story.append(Spacer(1, 8))

    # Cellstilar (etikett + värde). wordWrap='CJK' ger aggressivare radbrytning.
    label_style = ParagraphStyle(
        'MetaLabel', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=18, leading=22,
        alignment=TA_LEFT, wordWrap='CJK', spaceAfter=0, spaceBefore=0
    )
    value_style = ParagraphStyle(
        'MetaValue', parent=styles['Normal'],
        fontName='Helvetica', fontSize=18, leading=22,
        alignment=TA_LEFT, wordWrap='CJK', spaceAfter=0, spaceBefore=0
    )

    def L(x): return Paragraph(x if x else "-", label_style)
    def V(x): return Paragraph((x or "-"), value_style)

    meta_table_data = [
        [L("Kurs"),               V(info.kurs)],
        [L("Kurskod"),            V(info.kurskod)],
        [L("Tid"),                V(info.tid)],
        [L("Får senast börja"),   V(info.senastbörja or "09:00")],
        [L("Betala"),             V(info.betala)],
        [L("Hjälpmedel"),         V(info.hjälpmedel or "Inga")],
        [L("Ta hem tentafrågorna"), V(info.tahemtenta)],
        [L("Antal tentander"),    V(str(len(info.tentander)))],
        [L("Anonymkod"),          V(info.anonymkod)],
        [L("Övrig info"),         V(info.övriginfo)],
    ]

    # Ge värdekolumnen lite större bredd för långa texter.
    meta_table = Table(meta_table_data, colWidths=[80*mm, 100*mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(meta_table)
    meta_table.hAlign = "LEFT"
    story.append(Spacer(1, 12))

    PAD = 2  # samma som name_table LEFTPADDING

    heading = ParagraphStyle(
        "ParticipantsHeading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=22,
        leading=24,
        alignment=TA_LEFT,
        leftIndent=PAD,          # 🔑 synka med tabellen
        firstLineIndent=0,
        spaceBefore=0,
        spaceAfter=6,
    )
    story.append(Paragraph("Deltagare", heading))
    
    

    name_style = ParagraphStyle(
        'NameRow',
        parent=styles['Normal'],
        fontSize=20,
        leading=18,
        wordWrap='CJK'
    )

    names = info.tentander or ["(Inga valda)"]
    n = len(names)

    # Antal rader per kolumn (vänster fylls helt först)
    rows_per_col = 12

    table_data = []
    for r in range(rows_per_col):
        left_idx = r
        right_idx = r + rows_per_col

        left = names[left_idx] if left_idx < n else ""
        right = names[right_idx] if right_idx < n else ""

        table_data.append([
            Paragraph(left, name_style) if left else Paragraph("", name_style),
            Paragraph(right, name_style) if right else Paragraph("", name_style),
        ])

    # Anpassa bredder efter ditt dokument (du hade ~160mm tidigare)
    name_table = Table(table_data, colWidths=[90*mm, 90*mm])

    name_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),

        # Valfritt: en diskret linje mellan kolumnerna
        # ('LINEBEFORE', (1, 0), (1, -1), 0.25, colors.lightgrey),
    ]))

    story.append(name_table)
    name_table.hAlign = "LEFT"

    doc.build(story)

def append_pdf_copies(writer: PdfWriter, input_pdf: str, copies: int):
    for _ in range(copies):
        reader = PdfReader(input_pdf)

        for page in reader.pages:
            writer.add_page(copy.copy(page))

def duplicate_pdf(input_pdf: str, copies: int) -> PdfWriter:
    writer = PdfWriter()
    exam_reader = PdfReader(input_pdf)

    for _ in range(copies):
        for page in exam_reader.pages:
            writer.add_page(page)

def save_writer(writer: PdfWriter, out_path: str) -> None:
    with open(out_path, 'wb') as f:
        writer.write(f)

def print_pdf(file_path: str) -> None:
    system = platform.system()
    try:
        if system == 'Windows':
            os.startfile(file_path, 'print')
        else:
            subprocess.run(['lp', file_path], check=True)
    except Exception as e:
        messagebox.showwarning("Utskrift misslyckades", f"Fel: {e}")

# -------------------------------
# GUI
# -------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tentamensförberedaren")
        self.geometry("1200x900")

        # Filval
        self.exam_var = tk.StringVar()
        self.deltagar_var = tk.StringVar()

        # Metadata
        self.lärosäte_var = tk.StringVar()
        self.kurs_var = tk.StringVar()
        self.kurskod_var = tk.StringVar()
        self.tid_var = tk.StringVar(value="09:00-14:00")
        self.senast_börja_var = tk.StringVar(value="09:00")
        self.betala_var = tk.StringVar(value="Nej")
        self.hjälpmedel_var = tk.StringVar()
        self.tahemtenta_var = tk.StringVar(value="Nej")
        self.anonymkod_var = tk.StringVar(value="Ja")
        self.övriginfo_var = tk.StringVar()

        # Deltagare
        self.all_names: List[str] = []      # alla inlästa
        self.selected_names: List[str] = [] # valda
        self.copies_var = tk.IntVar(value=0)
        self.part_from_var = tk.IntVar(value=1)
        self.part_to_var   = tk.IntVar(value=1)

        # "Försättsblad till tentan" – sidval i deltagarlistan
        self.fsb_page_var = tk.IntVar(value=1)  # 1-indexerad sida
        self.deltagar_pages = 0                 # antal sidor i deltagarlistan

        self.partlist_pages_var = tk.IntVar(value=1)  # hur många sidor av deltagarlistan ska läggas in i paketet


        self._build_ui()

        # Vilka knappar ska synas för respektive lärosäte (matchar på substring i lower-case)
    #BUTTON_RULES = {
        #"karlstads universitet": {"cover", "make_print_pkg", "kvitto"},
        #"mittuniversitetet":     {"cover", "make_print_pkg", "exam_fronts_make"},
        #"högskolan i gävle":     {"cover", "make_print_pkg", "exam_fronts_make"},
        # fallback om inget matchar:
        #"*": {"cover", "make_print_pkg", "kvitto", "exam_fronts_make", "exam_fronts_print"},
    #}

    def open_manual_participants_window(self):
        win = tk.Toplevel(self)
        win.title("Lägg till deltagare manuellt")
        win.geometry("520x420")
        win.transient(self)
        win.grab_set()

        ttk.Label(win, text="Skriv/klistra in en deltagare per rad:").pack(anchor="w", padx=12, pady=(12, 6))

        txt = tk.Text(win, wrap="word", height=14)
        txt.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        # Hjälptext
        help_lbl = ttk.Label(
            win,
            text="Tips: Du kan klistra in från Excel. Tomma rader ignoreras och dubbletter läggs inte till."
        )
        help_lbl.pack(anchor="w", padx=12, pady=(0, 10))

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(0, 12))

        def _add():
            raw = txt.get("1.0", "end-1c")
            self.add_manual_participants(raw)
            win.destroy()

        ttk.Button(btn_row, text="Avbryt", command=win.destroy).pack(side="right")
        ttk.Button(btn_row, text="Lägg till", command=_add).pack(side="right", padx=(0, 8))


    def add_manual_participants(self, raw_text: str):
        import re

        if not hasattr(self, "all_names") or self.all_names is None:
            self.all_names = []

        # Normalisera input: en deltagare per rad
        lines = [ln.strip() for ln in (raw_text or "").splitlines()]
        lines = [ln for ln in lines if ln]

        if not lines:
            self.status_var.set("Inga namn att lägga till.")
            return

        # Rensa lite: ta bort eventuella extra mellanslag
        cleaned = []
        for ln in lines:
            ln = ln.replace("\u00a0", " ")
            ln = re.sub(r"\s{2,}", " ", ln).strip()
            cleaned.append(ln)

        # Undvik dubbletter (case-insensitive jämförelse)
        existing_norm = {n.strip().lower() for n in self.all_names if n}
        added = 0

        for name in cleaned:
            key = name.lower()
            if key in existing_norm:
                continue
            existing_norm.add(key)
            self.all_names.append(name)
            self.listbox.insert(tk.END, name)
            added += 1

        self.status_var.set(f"Lade till {added} deltagare manuellt. Totalt: {len(self.all_names)}.")

    def _apply_step4_layout(self):
        s = (self.lärosäte_var.get() or "").strip().lower()
        is_hig = ("högskolan i gävle" in s) or ("hig" in s)

        if is_hig:
            # Visa steg 4 ramen (om den mot förmodan blivit bortpackad)
            if not self.step4_frame.winfo_ismapped():
                self.step4_frame.pack(fill="x", padx=8, pady=6)

            # Dölj allt som inte ska synas
            if self.step4_row_a.winfo_ismapped():
                self.step4_row_a.pack_forget()
            if self.step4_sep.winfo_ismapped():
                self.step4_sep.pack_forget()

            # Se till att rad B syns
            if not self.step4_row_b.winfo_ismapped():
                self.step4_row_b.pack(fill="x", pady=(6, 6))

            # Dölj andra knappar i rad B
            if self.btn_exam_fronts_make.winfo_ismapped():
                self.btn_exam_fronts_make.pack_forget()
            if self.btn_exam_fronts_print.winfo_ismapped():
                self.btn_exam_fronts_print.pack_forget()

            # Visa paketknappen (på rätt plats)
            if not self.btn_pkg.winfo_ismapped():
                self.btn_pkg.pack(side="left", padx=6)

        else:
            # Normal layout: se till att alla delar syns
            if not self.step4_frame.winfo_ismapped():
                self.step4_frame.pack(fill="x", padx=8, pady=6)

            if not self.step4_row_a.winfo_ismapped():
                self.step4_row_a.pack(fill="x", pady=(4, 2))
            if not self.step4_sep.winfo_ismapped():
                self.step4_sep.pack(side="left", fill="y", padx=10)
            if not self.step4_row_b.winfo_ismapped():
                self.step4_row_b.pack(fill="x", pady=(2, 6))

            # Återställ knappar
            if not self.btn_pkg.winfo_ismapped():
                self.btn_pkg.pack(side="left", padx=6)
            if not self.btn_exam_fronts_make.winfo_ismapped():
                self.btn_exam_fronts_make.pack(side="left", padx=6)
            if not self.btn_exam_fronts_print.winfo_ismapped():
                self.btn_exam_fronts_print.pack(side="left", padx=6)


    def open_kvitto_window(self):
        """Öppnar ett fönster där man kan klistra in underlag till kvittot."""
        win = tk.Toplevel(self)
        win.title("Skapa kvitto – underlag")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()  # modal

        ttk.Label(win, text="Klistra in underlaget till kvittot här:").pack(
            anchor="w", padx=12, pady=(12, 6)
        )

        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(0, 12))

        def _run():
            content = txt.get("1.0", "end-1c").strip()
            if not content:
                messagebox.showwarning("Saknas", "Klistra in underlag först.")
                return
            self.run_kvitto_bat(content)
            win.destroy()

        ttk.Button(btn_row, text="Avbryt", command=win.destroy).pack(side="right")
        ttk.Button(btn_row, text="Skapa kvitto", command=_run).pack(side="right", padx=(0, 8))


    def run_kvitto_bat(self, kvitto_text=None, *args, **kwargs):
        """
        Kör Kvitto.bat.
        - kvitto_text: text från popup (om None körs utan fil)
        - *args/**kwargs: sväljer ev. event-argument om du råkat använda bind()
        """
        import os
        import subprocess
        import tempfile
        from tkinter import messagebox

        current_dir = os.path.dirname(os.path.abspath(__file__))
        bat_path = os.path.join(current_dir, "Kvitto.bat")

        if not os.path.exists(bat_path):
            messagebox.showerror("Fel", f"Kunde inte hitta {bat_path}")
            return

        args_cmd = ["cmd.exe", "/c", bat_path]

        if kvitto_text:
            fd, tmp_path = tempfile.mkstemp(prefix="kvitto_input_", suffix=".txt")
            os.close(fd)
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(kvitto_text)
            args_cmd.append(tmp_path)  # -> %1 i bat

        subprocess.Popen(args_cmd, cwd=current_dir)
        self.status_var.set("Kvitto: skickade underlag till Kvitto.bat.")



    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        self.status_var = tk.StringVar(value="Redo.")
        ttk.Label(self, textvariable=self.status_var).pack(
        side="bottom", fill="x", anchor="w", padx=10, pady=6)

        # ========== 1) Filer ==========
        files = ttk.LabelFrame(self, text="1) Välj filer")
        files.pack(fill="x", **pad)

        row1 = ttk.Frame(files)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="Välj tentamensfil:").pack(side="left")
        ttk.Entry(row1, textvariable=self.exam_var, width=50).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row1, text="Bläddra", command=lambda: self.choose_file(self.exam_var)).pack(side="left", padx=4)

        row2 = ttk.Frame(files)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="Välj deltagarlista:").pack(side="left")
        ttk.Entry(row2, textvariable=self.deltagar_var, width=50).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row2, text="Bläddra", command=lambda: self.choose_file(self.deltagar_var)).pack(side="left", padx=4)

        # ========== 2) Metainfo ==========
        meta = ttk.LabelFrame(self, text="2) Informationsblad åt tentavakt")
        meta.pack(fill="x", **pad)

        ttk.Label(meta, text="Lärosäte:").grid(row=0, column=0, sticky="e")
        cmb = AutocompleteCombobox(meta, textvariable=self.lärosäte_var, values=INSTITUTIONS)
        cmb.grid(row=0, column=1, sticky="we", padx=6)

        ttk.Label(meta, text="Kurs:").grid(row=1, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.kurs_var).grid(row=1, column=1, sticky="we", padx=6)

        ttk.Label(meta, text="Kurskod:").grid(row=2, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.kurskod_var).grid(row=2, column=1, sticky="we", padx=6)

        ttk.Label(meta, text="Tid:").grid(row=3, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.tid_var).grid(row=3, column=1, sticky="we", padx=6)

        ttk.Label(meta, text="Får senast starta:").grid(row=4, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.senast_börja_var).grid(row=4, column=1, sticky="we", padx=6)

        # Betala-rad + hjälpmedel
        row = 5
        ttk.Label(meta, text="Betala:").grid(row=row, column=0, sticky="e")
        row_frame = ttk.Frame(meta)
        row_frame.grid(row=row, column=1, sticky="we")
        ttk.Radiobutton(row_frame, text="Ja", value="Ja", variable=self.betala_var).pack(side="left")
        ttk.Radiobutton(row_frame, text="Nej", value="Nej", variable=self.betala_var).pack(side="left")

        row = 6
        ttk.Label(meta, text="Anonymkod:").grid(row=row, column=0, sticky="e")
        row_frame = ttk.Frame(meta)
        row_frame.grid(row=row, column=1, sticky="we")
        ttk.Radiobutton(row_frame, text="Ja", value="Ja", variable=self.anonymkod_var).pack(side="left")
        ttk.Radiobutton(row_frame, text="Nej", value="Nej", variable=self.anonymkod_var).pack(side="left")

        row = 7
        ttk.Label(meta, text="Hjälpmedel:").grid(row=row, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.hjälpmedel_var).grid(row=row, column=1, sticky="we", padx=6)

        row = 8
        ttk.Label(meta, text="Övrig info:").grid(row=row, column=0, sticky="e")
        ttk.Entry(meta, textvariable=self.övriginfo_var).grid(row=row, column=1, sticky="we", padx=6)

        meta.grid_columnconfigure(1, weight=1)

        # ========== 3) Deltagare ==========
        dl = ttk.LabelFrame(self, text="3) Deltagare – markera vilka som skriver hos oss")
        dl.pack(fill="both", expand=True, **pad)

        top_row = ttk.Frame(dl)
        top_row.pack(fill="x", padx=6, pady=4)

        ttk.Button(top_row, text="Läs in deltagarlista", command=self.read_participants).pack(side="left")

        # Sidor med deltagare (Från/Till)
        ttk.Label(top_row, text="Sidor med deltagare:").pack(side="left", padx=(12, 4))
        self.part_from_spin = tk.Spinbox(top_row, from_=1, to=1, width=4, textvariable=self.part_from_var)
        self.part_from_spin.pack(side="left")
        ttk.Label(top_row, text="till").pack(side="left", padx=4)
        self.part_to_spin = tk.Spinbox(top_row, from_=1, to=1, width=4, textvariable=self.part_to_var)
        self.part_to_spin.pack(side="left")

        # Listan med deltagare
        lb_font = tkfont.nametofont("TkDefaultFont").copy()
        lb_font.configure(size=13)

        self.listbox = tk.Listbox(dl, selectmode=tk.EXTENDED, height=12, font=lb_font)
        self.listbox.pack(fill="both", expand=True, **pad)

        btn_row = ttk.Frame(dl)
        btn_row.pack(fill="x", padx=6, pady=4)

        ttk.Button(btn_row,text="Använd markerade deltagare",command=self.use_selected).pack(side="left")

        ttk.Button(btn_row,text="Skapa informationsblad åt tentavakt",command=self.make_cover).pack(side="left", padx=(8, 0))

        ttk.Button(btn_row, text="Lägg till deltagare manuellt",command=self.open_manual_participants_window).pack(side="left", padx=(8, 0))

        # ========== 4) Skapa/skriv ut ==========
        self.step4_frame = ttk.LabelFrame(self, text="4) Skapa & utskrift")
        self.step4_frame.pack(fill="x", **pad)

        self.step4_row_a = ttk.Frame(self.step4_frame)
        self.step4_row_a.pack(fill="x", pady=(4, 2))

        self.step4_row_b = ttk.Frame(self.step4_frame)
        self.step4_row_b.pack(fill="x", pady=(2, 6))

        ttk.Label(self.step4_row_a, text="Hur många sidor är deltagarlistan?").pack(side="left")
        self.partlist_pages_spin = tk.Spinbox(self.step4_row_a, from_=1, to=1, width=4, textvariable=self.partlist_pages_var)
        self.partlist_pages_spin.pack(side="left", padx=4)

        # Knappen som ska vara enda synliga för HiG
        self.btn_pkg = ttk.Button(self.step4_row_b, text="Skapa och skriv ut paket", command=self.make_and_print_package)
        self.btn_pkg.pack(side="left", padx=6)

        # Separator och DL-försättsblad
        self.step4_sep = ttk.Separator(self.step4_frame, orient="vertical")
        self.step4_sep.pack(side="left", fill="y", padx=10)

        ttk.Label(self.step4_row_a, text="Vilken sida finns försättsbladet till tentan på?:").pack(side="left")
        self.fsb_spin = tk.Spinbox(self.step4_row_a, from_=1, to=1, textvariable=self.fsb_page_var, width=4)
        self.fsb_spin.pack(side="left", padx=4)

        ttk.Label(self.step4_row_a, text="Antal kopior:").pack(side="left")
        ttk.Label(self.step4_row_a, textvariable=self.copies_var).pack(side="left", padx=4)

        self.btn_exam_fronts_make = ttk.Button(self.step4_row_b, text="Skapa försättsblad", command=self.make_exam_fronts)
        self.btn_exam_fronts_make.pack(side="left", padx=6)

        self.btn_exam_fronts_print = ttk.Button(self.step4_row_b, text="Skriv ut försättsbladen", command=self.print_exam_fronts)
        self.btn_exam_fronts_print.pack(side="left", padx=6)

        # ========== 5) Skapa kvitto ==========
        self.step5_frame = ttk.LabelFrame(self, text="5) Skapa kvitto")
        self._step5_pack = dict(fill="x", **pad)          # spara pack-argument så vi kan visa igen snyggt
        self.step5_frame.pack(**self._step5_pack)

        btn_kvitto = ttk.Button(self.step5_frame, text="Skapa och skriv ut kvitto", command=self.open_kvitto_window)
        btn_kvitto.pack(side="left", padx=6)


        # Spara referenser och standard-pack-parametrar
        self._buttons = {
            "make_print_pkg":   {"w": self.btn_pkg, "pack": {"in_": self.step4_row_b, "side": "left", "padx": 6}},
            "kvitto":           {"w": btn_kvitto,   "pack": {"in_": self.step5_frame, "side": "left", "padx": 6}},
            "exam_fronts_make": {"w": self.btn_exam_fronts_make, "pack": {"in_": self.step4_row_b, "side": "left", "padx": 6}},
            "exam_fronts_print":{"w": self.btn_exam_fronts_print,"pack": {"in_": self.step4_row_b, "side": "left", "padx": 6}},
        }


        # Regler: vilka knappar som syns för respektive lärosäte (substring-match i lower-case)
        self.BUTTON_RULES = {
            "karlstads universitet": {"make_print_pkg", "exam_fronts_make", "kvitto"},
            "mittuniversitetet":     {"make_print_pkg", "exam_fronts_make", "kvitto"},
            "högskolan i gävle":     {"make_print_pkg", },
            # Fallback (om ingen match):
            "*": {"make_print_pkg", "exam_fronts_make", "exam_fronts_print"},
        }

        #cmb.bind("<<ComboboxSelected>>", lambda e: self._update_buttons_for_institution())
        #self.lärosäte_var.trace_add("write", lambda *_: self._update_buttons_for_institution())

        # kör en gång vid start
        #self._apply_step4_layout()

        # Intern updaterare (binder både write-trace och combobox-selected)
        def _update_buttons_for_institution(*_):
            s = (self.lärosäte_var.get() or "").strip().lower()
            is_hig = ("högskolan i gävle" in s) or ("hig" in s)

            # 1) Avgör vilka knappar som ska synas (enligt regler)
            visible = None
            for key, show in self.BUTTON_RULES.items():
                if key == "*":
                    continue
                if key in s:
                    visible = show
                    break
            if visible is None:
                visible = self.BUTTON_RULES.get("*", set())

            # 2) Steg 5 (kvitto): visa/dölj hela ramen
            if "kvitto" in visible:
                if not self.step5_frame.winfo_ismapped():
                    self.step5_frame.pack(**self._step5_pack)
            else:
                if self.step5_frame.winfo_ismapped():
                    self.step5_frame.pack_forget()

            # 3) Steg 4 specialfall: HiG = endast paketknappen synlig
            if is_hig:
                # Dölj allt i steg 4 (men inte själva steg4_frame)
                if self.step4_row_a.winfo_ismapped():
                    self.step4_row_a.pack_forget()
                if self.step4_sep.winfo_ismapped():
                    self.step4_sep.pack_forget()

                if not self.step4_row_b.winfo_ismapped():
                    self.step4_row_b.pack(fill="x", pady=(6, 6))

                # Dölj andra knappar
                for k in ("exam_fronts_make", "exam_fronts_print"):
                    if k in self._buttons:
                        self._buttons[k]["w"].pack_forget()

                # Visa bara paket
                rec = self._buttons["make_print_pkg"]
                rec["w"].pack(**rec["pack"])
                return 4

            # 4) Normal läge (icke-HiG): se till att steg 4 byggstenar syns
            if not self.step4_row_a.winfo_ismapped():
                self.step4_row_a.pack(fill="x", pady=(4, 2))
            if not self.step4_sep.winfo_ismapped():
                self.step4_sep.pack(side="left", fill="y", padx=10)
            if not self.step4_row_b.winfo_ismapped():
                self.step4_row_b.pack(fill="x", pady=(2, 6))

            # 5) Dölj alla knappar i _buttons
            for key, rec in self._buttons.items():
                rec["w"].pack_forget()

            # 6) Visa de knappar som ska synas
            for key in visible:
                if key in self._buttons:
                    rec = self._buttons[key]
                    rec["w"].pack(**rec["pack"])
            # Exponera på instansen
        self._update_buttons_for_institution = _update_buttons_for_institution

        # Kör initialt
        self._update_buttons_for_institution()

        # Bind uppdatering när lärosäte ändras
        cmb.bind("<<ComboboxSelected>>", lambda e: self._update_buttons_for_institution())
        self.lärosäte_var.trace_add("write", lambda *_: self._update_buttons_for_institution())


    # --- callbacks ---
    def choose_file(self, var: tk.StringVar):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            var.set(path)

    def make_and_print_package(self):
        """Skapa paketet och skriv ut det direkt om skapandet lyckas."""
        # 1) Skapa paket (återanvänd din befintliga logik)
        #    Den här visar redan felmeddelanden om något går snett.
        self.make_package()

        # 2) Hitta den förväntade filen och skriv ut (med bekräftelse)
        folder = os.path.dirname(self.exam_var.get()) or os.getcwd()
        pkg = os.path.join(folder, "tentamen_paket.pdf")

        # Om make_package använde fallback-namn (p.g.a. PermissionError) kan filnamnet vara annorlunda.
        # Leta upp senaste skapade “tentamen_paket*.pdf” i mappen som en extra säkerhet:
        if not os.path.exists(pkg):
            try:
                candidates = [f for f in os.listdir(folder) if f.startswith("tentamen_paket") and f.endswith(".pdf")]
                if candidates:
                    # välj senast ändrade
                    candidates.sort(key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)
                    pkg = os.path.join(folder, candidates[0])
            except Exception:
                pass

        if not os.path.exists(pkg):
            messagebox.showwarning("Saknas", "Kunde inte hitta det skapade paketet.\nKontrollera behörigheter och försök igen.")
            return

        # 3) Fråga och skriv ut
        if messagebox.askyesno("Skriv ut", f"Skriv ut paketet?\n\n{pkg}"):
            print_pdf(pkg)
            self.status_var.set(f"Sände till skrivare: {pkg}")



    def read_participants(self):
        pdf = self.deltagar_var.get()
        if not pdf:
            messagebox.showwarning("Saknas", "Välj en deltagarlista först.")
            return

        if not self.lärosäte_var.get().strip():
            messagebox.showwarning("Lärosäte saknas", "Välj lärosäte under 2) Information.")
            return

        maxpages = max(1, self.deltagar_pages)
        self.partlist_pages_spin.config(to=maxpages)
        if not (1 <= self.partlist_pages_var.get() <= maxpages):
            self.partlist_pages_var.set(1)
        self.part_to_spin.config(to=maxpages)

        # se till att intervallet är giltigt
        if not (1 <= self.part_from_var.get() <= maxpages):
            self.part_from_var.set(1)
        if not (1 <= self.part_to_var.get() <= maxpages):
            self.part_to_var.set(min(1, maxpages))
        if self.part_to_var.get() < self.part_from_var.get():
            self.part_to_var.set(self.part_from_var.get())

        # uppdatera sidantalet (behövs för din spinbox-del)
        try:
            rdr = PdfReader(pdf)
            self.deltagar_pages = len(rdr.pages)
            self.fsb_spin.config(to=max(1, self.deltagar_pages))
            if not (1 <= self.fsb_page_var.get() <= self.deltagar_pages):
                self.fsb_page_var.set(1)
        except Exception:
            self.deltagar_pages = 0
            self.fsb_spin.config(to=1)
            self.fsb_page_var.set(1)

        # --- NYTT: välj extraktion baserat på lärosäte ---
        names = extract_names_by_institution(
        pdf,
        self.lärosäte_var.get(),
        page_from=self.part_from_var.get(),
        page_to=self.part_to_var.get(),
        )

        self.listbox.delete(0, tk.END)
        for n in names:
            self.listbox.insert(tk.END, n)

        self.all_names = names
        self.selected_names = []
        self.copies_var.set(0)
        self.status_var.set(f"{len(names)} namn inlästa – sidor {self.part_from_var.get()}–{self.part_to_var.get()}. Markera de som ska användas.")

        

    def _normalize_name(self, n: str) -> str:

        try:
            # använd globala hjälpfunktionen om den finns
            return _normalize_name(n)  # type: ignore[name-defined]
        except NameError:
            # minimal fallback: "Efternamn, Förnamn" -> "Förnamn Efternamn"
            n = (n or "").strip()
            if not n:
                return ""
            if "," in n:
                last, rest = [p.strip() for p in n.split(",", 1)]
                return f"{rest} {last}".strip()
            return n

    def use_selected(self):
        sel_raw = [self.listbox.get(i) for i in self.listbox.curselection()]
        if not sel_raw:
            messagebox.showinfo("Inget valt", "Markera minst en deltagare.")
            return

        # Rena namn redan i listboxen → bara normalisera
        selected = [self._normalize_name(r) for r in sel_raw if r.strip()]

        self.selected_names = selected
        self.copies_var.set(len(selected))
        self.status_var.set(f"{len(selected)} deltagare valda.")


    def _collect_info(self) -> Optional[ExamInfo]:
        tentander = self.selected_names if self.selected_names else self.all_names
        tentander = [self._normalize_name(n) for n in tentander if n.strip()]
        return ExamInfo(
            lärosäte=self.lärosäte_var.get().strip(),
            kurs=self.kurs_var.get().strip(),
            kurskod=self.kurskod_var.get().strip(),
            tid=self.tid_var.get().strip(),
            senastbörja=self.senast_börja_var.get(),
            betala=self.betala_var.get(),
            hjälpmedel=self.hjälpmedel_var.get().strip(),
            tahemtenta=self.tahemtenta_var.get(),
            tentander=tentander.copy(),
            anonymkod=self.anonymkod_var.get(),
            övriginfo=self.övriginfo_var.get(),
        )

    def make_cover(self):
        info = self._collect_info()
        if not info:
            return
        folder = os.path.dirname(self.exam_var.get()) or os.getcwd()
        out = os.path.join(folder, "forsattsblad.pdf")
        try:
            make_cover_sheet_pdf(out, info)
            messagebox.showinfo("Klart", f"Försättsblad skapat:\n{out}")
        except PermissionError:
            alt = os.path.join(folder, f"forsattsblad_{uuid.uuid4().hex[:6]}.pdf")
            try:
                make_cover_sheet_pdf(alt, info)
                messagebox.showinfo(
                    "Klart",
                    f"Försättsblad skapat (alternativ sökväg):\n{alt}\n"
                    "(Originalfilen kan vara låst eller öppnad i en läsare.)"
                )
            except Exception as e:
                messagebox.showerror("Fel", f"Kunde inte skapa försättsblad:\n{e}")
        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte skapa försättsblad:\n{e}")

    def make_package(self):
        import os
        from tkinter import messagebox
        from PyPDF2 import PdfReader, PdfWriter  # eller PyPDF2
        # from PyPDF2 import PdfReader, PdfWriter  # om du använder PyPDF2

        # ---- Hämta paths från UI (anpassa variabelnamn om dina heter annorlunda) ----
        dl_path = (self.deltagar_var.get() or "").strip()      # Steg 1: "Välj deltagarlista"
        exam_path = (self.exam_var.get() or "").strip()        # Steg 1: "Välj tentamensfil"
        base_dir = os.path.dirname(dl_path)
        out_path = os.path.join(base_dir, "tentamen_paket.pdf")
        info = self._collect_info()                       # eller hur du bygger ExamInfo från steg 2

        # Antal kopior = antal valda deltagare (Steg 3)
        try:
            copies = len(getattr(info, "tentander", []) or [])
        except Exception:
            copies = 0

        if not dl_path or not os.path.exists(dl_path):
            messagebox.showerror("Saknas", "Välj en deltagarlista (PDF) i steg 1.")
            return

        if not exam_path or not os.path.exists(exam_path):
            messagebox.showerror("Saknas", "Välj en tentamensfil (PDF) i steg 1.")
            return

        if copies <= 0:
            messagebox.showwarning("Inga deltagare", "Inga deltagare är valda i steg 3.")
            return

        if not out_path:
            messagebox.showerror("Saknas", "Ange var paketet ska sparas (utfil).")
            return

        # ---- Lärosätesdetektion ----
        lärosäte = (getattr(info, "lärosäte", "") or "").strip().lower()
        is_hig = ("högskolan i gävle" in lärosäte) or ("hig" in lärosäte)

        writer = PdfWriter()

        # ---- 1) Informationsblad ----
        base_dir = os.path.dirname(dl_path)
        cover_path = os.path.join(base_dir, "forsattsblad.pdf")

        if not os.path.exists(cover_path):
            messagebox.showerror(
                "Saknas",
                f"Hittar inte informationsbladet:\n{cover_path}"
            )
            return

        _append_all_pages(writer, cover_path)


        # ---- 2) Deltagarlista (steg 1) ----
        if is_hig:
            # HiG: hela PDF:en (deltagarlista + HiG-försättsblad)
            _append_all_pages(writer, dl_path)
        else:
            # Övriga: om du vill fortsätta ta första N sidor:
            try:
                dl_pages = int(self.partlist_pages_var.get() or 0)
            except Exception:
                dl_pages = 0

            if dl_pages > 0:
                # Behåll din befintliga funktion om du redan har en:
                # _append_first_n_pages(writer, dl_path, dl_pages)
                rdr = PdfReader(dl_path)
                for i, pg in enumerate(rdr.pages):
                    if i >= dl_pages:
                        break
                    writer.add_page(pg)
            else:
                # fallback: ta hela om inget angivet
                _append_all_pages(writer, dl_path)

                # --- Extra: försättsblad till tentan (om filen finns i mappen) ---
        extra_front_path = os.path.join(base_dir, "forsattsblad_till_tentan.pdf")

        if os.path.exists(extra_front_path):
            try:
                _append_all_pages(writer, extra_front_path)
                self.status_var.set("Extra försättsblad till tentan hittades och lades till.")
            except Exception as e:
                self.status_var.set(f"Kunde inte lägga till forsattsblad_till_tentan.pdf: {e}")

        # ---- 3) X kopior av tentan (steg 1) ----
        # VIKTIGT: Lägg bara in tentan EN gång här (inte två block!)
        for _ in range(copies):
            # Ny PdfReader varje varv = säkrare för vissa PDF:er
            exam_reader = PdfReader(exam_path)

            for pg in exam_reader.pages:
                # Klona sidan innan den läggs till
                writer.add_page(copy.copy(pg))

        # ---- Spara paket ----
        try:
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
        except Exception:
            pass

        with open(out_path, "wb") as f:
            writer.write(f)

        if hasattr(self, "status"):
            self.status_var.set(f"Paket skapat: {out_path} (HiG={is_hig}, kopior={copies})")


    def print_package(self):
        folder = os.path.dirname(self.exam_var.get()) or os.getcwd()
        pkg = os.path.join(folder, "tentamen_paket.pdf")
        if not os.path.exists(pkg):
            messagebox.showwarning("Saknas", "Skapa paketet först.")
            return
        if messagebox.askyesno("Skriv ut", "Skicka paketet till standardskrivaren?"):
            print_pdf(pkg)
    
    # Antal deltagare (valda om finns, annars alla inlästa)
    def _participant_count(self) -> int:
        return len(self.selected_names) if self.selected_names else len(self.all_names)
    
    def make_exam_fronts(self):
        pdf = (self.deltagar_var.get() or "").strip()
        if not pdf or not os.path.exists(pdf):
            messagebox.showwarning("Saknas", "Välj en deltagarlista (PDF).")
            return
        copies = self._participant_count()
        if copies <= 0:
            messagebox.showwarning("Inga deltagare", "Läs in och välj deltagare först.")
            return

        # Sida (1-indexerad i spinbox) -> 0-index för PyPDF2
        if self.deltagar_pages <= 0:
            try:
                self.deltagar_pages = len(PdfReader(pdf).pages)
            except Exception:
                self.deltagar_pages = 0
        page_idx = max(1, min(self.fsb_page_var.get(), max(1, self.deltagar_pages))) - 1

        try:
            rdr = PdfReader(pdf)
            if page_idx >= len(rdr.pages):
                messagebox.showwarning("Ogiltig sida", f"PDF:en har bara {len(rdr.pages)} sidor.")
                return
            writer = PdfWriter()
            page = rdr.pages[page_idx]
            for _ in range(copies):
                writer.add_page(page)

            folder = os.path.dirname(pdf) or os.getcwd()
            out = os.path.join(folder, "forsattsblad_till_tentan.pdf")
            try:
                save_writer(writer, out)
            except PermissionError:
                out = os.path.join(folder, f"forsattsblad_till_tentan_{uuid.uuid4().hex[:6]}.pdf")
                save_writer(writer, out)

            messagebox.showinfo("Klart", f"Försättsblad till tentan skapat:\n{out}\n(Kopior: {copies})")

            # Öppna filen för tydlighet
            try:
                system = platform.system()
                if system == 'Darwin':
                    subprocess.run(['open', out], check=False)
                elif system == 'Windows':
                    os.startfile(out)
                else:
                    subprocess.run(['xdg-open', out], check=False)
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte skapa försättsblad till tentan:\n{e}")

    def print_exam_fronts(self):
        pdf = (self.deltagar_var.get() or "").strip()
        if not pdf:
            messagebox.showwarning("Saknas", "Välj en deltagarlista (PDF).")
            return
        folder = os.path.dirname(pdf) or os.getcwd()
        out = os.path.join(folder, "forsattsblad_till_tentan.pdf")
        if not os.path.exists(out):
            messagebox.showwarning("Saknas", "Skapa 'Försättsblad till tentan' först.")
            return
        if messagebox.askyesno("Skriv ut", "Skriv ut 'Försättsblad till tentan'?"):
            print_pdf(out)

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
