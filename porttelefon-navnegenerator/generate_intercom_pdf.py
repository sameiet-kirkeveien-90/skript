#!/usr/bin/env python3
"""
Generer en pdf med teksten for panelet på porttelefonen fra en CSV med følgende kolonner:

    leilighetsnummer,navn

Input:
- "leilighetsnummer" feltet har format "OPPGANG|HNNMM", f.eks. "E|H0201".
  * OPPGANG ∈ {A, B, C, D, E} (én bokstav).
  * HNNMM: NN = etasje, MM ∈ {01,02} (01=venstre, 02=høyre i utgangspunktet).

Viktige begreper:
- SeksjonsID = "oppgang|leilighetsnummer" (hele feltet). Denne er unik.
- Leilighetsnummer alene er ikke unikt på tvers av oppganger.

Regler:
- En side per OPPGANG med header/bunntekst "Oppgang X".
- Toppbokser: etasjene 8–5, bunnbokser: 4–1. Alltid 4 "slisser" per boks.
- 8. etasje-spesial: Hvis en OPPGANG har nøyaktig én leilighet i etasje 8,
  plasser den i høyre kolonne uansett 01/02. Hvis to, bruk vanlig 01→L / 02→R.
- Oppgang A-spesial: 1. etasje finnes ikke -> bunnboksen bunn-justeres
  (020x nederst, 030x over, 040x over, øverste slisse tom).
- Tekst i VERSALER. Hvis "navn" mangler, vis SeksjonsID (oppgang|leil).

Bruk:
    pip install reportlab
    python generate_intercom_pdf.py input.csv output.pdf
"""

import re, sys, csv, io
from typing import List, Tuple, Dict
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth

# =========================
# Justerbare konstanter
# =========================
PAGE_SIZE = A4
PAGE_MARGIN_MM   = 15
COLUMN_GAP_MM    = 14
BOX_W_MM         = 62
BOX_H_MM         = 76
BOX_H_SPACING_MM = 16

INNER_PAD_MM        = 2
TOP_INNER_MARGIN_MM = 6
BOT_INNER_MARGIN_MM = 6

FONT_NAME   = "Times-Bold"
FONT_MAX_PT = 16
FONT_MIN_PT = 14

HEADER_FONT_NAME = "Helvetica-Bold"
HEADER_FONT_PT   = 12
HEADER_GAP_MM    = 4

FLOORS_TOP = [8, 7, 6, 5]
FLOORS_BOT = [4, 3, 2, 1]
# =========================

def mm2pt(x): 
    return x * mm

def parse_seksjonsid(seksjonsid: str) -> Tuple[str, int, str, str]:
    """
    Parse 'E|H0201' -> (oppgang='E', etasje=2, unit='01'/'02', side_by_unit='L'/'R')
    NB: ingen 8.-etg.-tvang her; det avgjøres pr. oppgang senere.
    """
    if not seksjonsid:
        raise ValueError("Tom SeksjonsID")
    parts = [p.strip() for p in seksjonsid.split("|", 1)]
    if len(parts) != 2:
        raise ValueError(f"Mangler '|': {seksjonsid!r}")
    oppgang = parts[0].upper()
    right   = parts[1]
    m = re.search(r'(\d{4})\s*$', right)
    if not m:
        raise ValueError(f"Fant ikke 4 sifre i: {seksjonsid!r}")
    last4 = m.group(1)
    etasje = int(last4[:2])
    unit   = last4[2:]
    side_by_unit = 'L' if unit == '01' else 'R'
    return oppgang, etasje, unit, side_by_unit

def _sniff_delimiter(sample: str) -> str:
    counts = {sep: sample.count(sep) for sep in (';', ',', '\t')}
    return max(counts, key=counts.get) if max(counts.values()) > 0 else ','

def read_rows(path: str):
    """
    Les CSV og returner liste av dicts:
      {
        'seksjonsid': 'E|H0201',
        'oppgang': 'E',
        'etasje': 2,
        'unit': '01',
        'side_unit': 'L'/'R',   # ut fra unit
        'display': '...'        # VERSALER
      }
    """
    with open(path, 'rb') as fb:
        raw = fb.read()
    text = raw.decode('utf-8-sig', errors='replace')

    lines = [ln for ln in text.splitlines() if ln.strip()]
    head = '\n'.join(lines[:5])
    delimiter = _sniff_delimiter(head)

    f = io.StringIO(text)
    raw_reader = csv.reader(f, delimiter=delimiter)
    fieldnames = next(raw_reader)
    norm = [h.strip().lower() for h in fieldnames]

    def rowdict(row):
        return {norm[i]: (row[i] if i < len(row) else '') for i in range(len(norm))}

    rows = []
    for row in raw_reader:
        rec  = rowdict(row)
        seksjonsid = (rec.get('leilighetsnummer') or rec.get('seksjonsid') or rec.get('apt') or '').strip()
        navn = (rec.get('navn') or rec.get('name') or '').strip()
        if not seksjonsid:
            continue
        oppgang, etasje, unit, side_by_unit = parse_seksjonsid(seksjonsid)
        display = (navn if navn else seksjonsid.split('|')[1]).upper()
        rows.append({
            'seksjonsid': seksjonsid,
            'oppgang': oppgang,
            'etasje': etasje,
            'unit': unit,
            'side_unit': side_by_unit,
            'display': display,
        })
    return rows

def build_boxes_for_oppgang(rows_for_oppgang: List[Dict]):
    """
    Global slisseplassering per oppgang:
      - 8 faste slisser (0..7, topp→bunn).
      - base_slot = {8:0, 7:1, 6:2, 5:3, 4:4, 3:5, 2:6, 1:7}
      - offset = antall manglende etasjer i {1,2,3,4} for denne oppgangen.
        (Oppgang A mangler 1. etg ⇒ offset=1 ⇒ alt flyttes én slisse ned.)
      - Enhets-/side-regel: 01→L, 02→R, men hvis oppgangen har nøyaktig én
        leilighet i 8. etg, plasser den i R uansett.
    Splitter så slisser 0..3 → TOP, 4..7 → BOT, per side.
    Returnerer fire bokser (de uten innhold filtreres bort).
    """
    if not rows_for_oppgang:
        return []

    oppgang = rows_for_oppgang[0]['oppgang']

    # 1) Finn hvilke etasjer som finnes i oppgangen (uavhengig av side)
    present_floors = {r['etasje'] for r in rows_for_oppgang if 1 <= r['etasje'] <= 8}

    # 2) Global offset = manglende nederste etasjer i {1,2,3,4}
    bottom_set = {1, 2, 3, 4}
    missing_bottom = sorted(bottom_set - present_floors)  # info, om ønskelig
    offset = len(bottom_set - present_floors)

    # 3) 8. etasje-spesial per oppgang
    floor8 = [r for r in rows_for_oppgang if r['etasje'] == 8]
    single8 = (len(floor8) == 1)

    # 4) Slissekart og datastruktur for linjer
    base_slot = {8:0, 7:1, 6:2, 5:3, 4:4, 3:5, 2:6, 1:7}
    lines = {
        ('L', 'TOP'): [None]*4,
        ('R', 'TOP'): [None]*4,
        ('L', 'BOT'): [None]*4,
        ('R', 'BOT'): [None]*4,
    }

    # 5) Plasser hver rad i riktig global slisse + side
    for r in rows_for_oppgang:
        floor = r['etasje']
        if floor not in base_slot:
            continue
        # side ut fra unit, evt. overstyr for 8. etasje ved single8
        side = r['side_unit']
        if floor == 8 and single8:
            side = 'R'

        idx = base_slot[floor] + offset
        if not (0 <= idx <= 7):
            continue  # utenfor synlig 8-slissers vindu

        band = 'TOP' if idx <= 3 else 'BOT'
        pos  = idx if band == 'TOP' else idx - 4

        # "Siste vinner" hvis duplikate inputlinjer mot samme posisjon
        lines[(side, band)][pos] = r['display']

    # 6) Bygg bokser (filtrer bort helt tomme)
    boxes = [
        {'column': 'L', 'order': 1, 'lines': lines[('L','TOP')]},
        {'column': 'R', 'order': 1, 'lines': lines[('R','TOP')]},
        {'column': 'L', 'order': 2, 'lines': lines[('L','BOT')]},
        {'column': 'R', 'order': 2, 'lines': lines[('R','BOT')]},
    ]
    return [b for b in boxes if any(x for x in b['lines'] if x)]


def draw_oppgang_page(c: canvas.Canvas, oppgang: str, boxes):
    """Tegn én side for én oppgang, med header 'Oppgang X' oppe og nede."""
    width_pt, height_pt = PAGE_SIZE
    page_margin = mm2pt(PAGE_MARGIN_MM)
    col_gap     = mm2pt(COLUMN_GAP_MM)
    box_w       = mm2pt(BOX_W_MM)
    box_h       = mm2pt(BOX_H_MM)
    box_v_gap   = mm2pt(BOX_H_SPACING_MM)
    inner_pad   = mm2pt(INNER_PAD_MM)
    top_band    = mm2pt(TOP_INNER_MARGIN_MM)
    bot_band    = mm2pt(BOT_INNER_MARGIN_MM)
    header_gap  = mm2pt(HEADER_GAP_MM)

    header_text = f"Oppgang {oppgang}"
    c.setFont(HEADER_FONT_NAME, HEADER_FONT_PT)
    header_height = HEADER_FONT_PT + header_gap

    total_cols_w = 2 * box_w + col_gap
    page_inner_w = width_pt - 2 * page_margin
    x_left  = page_margin + max(0, (page_inner_w - total_cols_w) / 2.0)
    x_right = x_left + box_w + col_gap

    y_top   = height_pt - page_margin - header_height
    y_bottom_reserved = page_margin + header_height

    left  = [b for b in boxes if b['column']=='L']
    right = [b for b in boxes if b['column']=='R']
    left.sort(key=lambda b: b['order'])
    right.sort(key=lambda b: b['order'])

    def draw_header(y):
        c.setFillColorRGB(0,0,0)
        c.setFont(HEADER_FONT_NAME, HEADER_FONT_PT)
        w = stringWidth(header_text, HEADER_FONT_NAME, HEADER_FONT_PT)
        c.drawString((width_pt - w)/2.0, y, header_text)

    def draw_column(col_boxes, x):
        y_top   = height_pt - page_margin - header_height  # <- som før utenfor
        y_cursor = y_top
        for b in col_boxes:
            top_y, bottom_y = y_cursor, y_cursor - box_h
            if bottom_y < y_bottom_reserved:
                break

            # ramme
            c.setStrokeColorRGB(0.9,0.9,0.9)
            c.setFillColorRGB(1,1,1)
            c.rect(x, bottom_y, box_w, box_h, stroke=1, fill=1)

            # tekst
            c.setStrokeColorRGB(0,0,0)

            # Innvendig band (horisontalt)
            content_left, content_right = x + inner_pad, x + box_w - inner_pad
            content_width = content_right - content_left
            x_center = (content_left + content_right) / 2.0

            # === NYTT: "trygt" vertikalbånd for baselines ===
            # Vi reserverer plass tilsvarende ~0.6*FONT_MAX_PT over/under hver baseline,
            # slik at selv ved 16 pt havner tekst innenfor boksen når padding=0.
            # (0.6 og 0.35 er konservative verdier for Helvetica-Bold)
            F = FONT_MAX_PT
            safe_top    = top_y    - mm2pt(TOP_INNER_MARGIN_MM)  - F * 0.6
            safe_bottom = bottom_y + mm2pt(BOT_INNER_MARGIN_MM) + F * 0.6
            # fallback om bandet kollapser:
            if safe_bottom >= safe_top:
                mid = (top_y + bottom_y) / 2.0
                safe_top = mid + mm2pt(5)
                safe_bottom = mid - mm2pt(5)

            lines = b['lines']               # alltid 4 slisser (kan være None)
            n = len(lines)

            # Fordel baselines jevnt i safe-bandet (topp→bunn)
            if n == 1:
                baselines = [(safe_top + safe_bottom) / 2.0]
            else:
                step = (safe_top - safe_bottom) / (n - 1)
                baselines = [safe_top - i * step for i in range(n)]

            # Tegn hver slisse (hopp over None → tom slisse, men behold posisjon)
            for text, baseline in zip(lines, baselines):
                if not text:
                    continue

                # Finn font size som passer bredden (16→14 pt), ellers horisontal skalering
                font_size = FONT_MAX_PT
                while font_size > FONT_MIN_PT:
                    w = stringWidth(text, FONT_NAME, font_size)
                    if w <= content_width:
                        break
                    font_size -= 0.5

                c.setFillColorRGB(0, 0, 0)
                c.setFont(FONT_NAME, font_size)

                w = stringWidth(text, FONT_NAME, font_size)

                # Vertikal plassering: baseline litt under “senterlinjen”
                # 0.35*font_size funker bra for Helvetica-Bold som grov baseline-offset
                y_text = baseline - font_size * 0.35

                if w <= content_width:
                    c.drawString(x_center - w / 2.0, y_text, text)
                else:
                    scale_x = content_width / w if w > 0 else 1.0
                    c.saveState()
                    c.translate(x_center - (w * scale_x) / 2.0, 0)
                    c.scale(scale_x, 1.0)
                    c.drawString(0, y_text, text)
                    c.restoreState()

            y_cursor = bottom_y - box_v_gap


    # Header oppe og nede
    draw_header(height_pt - page_margin - HEADER_FONT_PT)
    draw_column(left, x_left)
    draw_column(right, x_right)
    draw_header(page_margin)
    c.showPage()

def main(argv):
    if len(argv) != 3:
        print("Usage: python generate_intercom_pdf.py input.csv output.pdf", file=sys.stderr)
        return 2

    rows = read_rows(argv[1])  # liste av dicts med SeksjonsID + felt
    by_oppgang: Dict[str, List[Dict]] = defaultdict(list)
    for r in rows:
        by_oppgang[r['oppgang']].append(r)

    c = canvas.Canvas(argv[2], pagesize=PAGE_SIZE)

    any_pages = False
    for oppgang in sorted(by_oppgang.keys()):
        boxes = build_boxes_for_oppgang(by_oppgang[oppgang])
        if not boxes:
            continue
        draw_oppgang_page(c, oppgang, boxes)
        any_pages = True

    if not any_pages:
        # Tom placeholder-side
        draw_oppgang_page(c, "?", [])

    c.save()
    print(f"Wrote {argv[2]}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
