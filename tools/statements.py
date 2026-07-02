"""Historicals engine: parse property income statements (summary or transaction
detail) and assemble a workbook — one source tab per statement + a combined tab
where every value is a formula linking back to its source, categories are aligned
across years (union), missing-year lines shown in red, and each line classified."""

import io
import re
import datetime
import difflib

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CUR = '#,##0.00'
RED = Font(color='C00000'); REDB = Font(color='C00000', bold=True)
BOLD = Font(bold=True); PLAIN = Font()
HDR = Font(bold=True, color='FFFFFF'); HFILL = PatternFill('solid', fgColor='4F46E5')

_ROW = re.compile(r'^(.*?)\s+(-?\$[\d,]+\.\d{2})(?:\s+-?\$[\d,]+\.\d{2})?\s*$')
_TXN = re.compile(r'^([A-Za-z]{2,12})\s+(\d{1,2}/\d{1,2}/\d{4})\s+.*?(-?\$[\d,]+\.\d{2})\s+(-?\$[\d,]+\.\d{2})\s*$')
_HEADER = re.compile(r'^(.+?)\s*\((\d{3,6})\)\s*$')
_TOTAL = re.compile(r'^Total\s+(.+?)\s+(-?\$[\d,]+\.\d{2})\s*$')
_MONEY = re.compile(r'-?\$[\d,]+\.\d{2}')
_PERIOD = re.compile(r'(\d{1,2})/(\d{1,2})/(\d{4})\s*-\s*(\d{1,2})/(\d{1,2})/(\d{4})')


def _money(m):
    return float(m.replace('$', '').replace(',', ''))


def _doubled(s):
    return len(s) >= 2 and len(s) % 2 == 0 and all(s[i] == s[i + 1] for i in range(0, len(s), 2))


def _text(data):
    with pdfplumber.open(io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data) as pdf:
        return '\n'.join((p.extract_text() or '') for p in pdf.pages)


def detect_kind(data):
    t = _text(data)
    return 'detail' if ('Cash Flow Detail' in t or _doubled('CCaasshh') and 'DDeettaaiill' in t or
                        sum(bool(_TXN.match(l.strip())) for l in t.splitlines()) > 15) else 'summary'


# ---------------------------------------------------------------- parsers
def parse_summary(data):
    rows = []
    for raw in _text(data).splitlines():
        s = raw.strip()
        if not s:
            continue
        m = _ROW.match(s)
        if m:
            label = m.group(1).strip()
            rows.append({'label': label, 'amount': _money(m.group(2)),
                         'total': label.lower().startswith('total'), 'net': label.lower().startswith('net')})
        elif s.lower() == 'income':
            rows.append({'label': s, 'amount': None, 'section': True})
    return rows


def parse_detail(data):
    t = _text(data)
    cats, totals, cur = {}, {}, None
    seen_dates = []
    for raw in t.splitlines():
        s = raw.strip()
        if not s or _doubled(s):
            continue
        mt = _TOTAL.match(s)
        if mt:
            totals[mt.group(1).strip()] = _money(mt.group(2)); cur = None; continue
        mtx = _TXN.match(s)
        if mtx and cur:
            d = datetime.datetime.strptime(mtx.group(2), '%m/%d/%Y').date()
            seen_dates.append(d)
            cats.setdefault(cur, {}).setdefault(f'{d.year}-{d.month:02d}', 0.0)
            cats[cur][f'{d.year}-{d.month:02d}'] += _money(mtx.group(4)); continue
        mh = _HEADER.match(s)
        if mh and not _MONEY.search(s):
            cur = mh.group(1).strip()
    # month columns from the actual transaction date range (robust to formatting)
    months = []
    if seen_dates:
        d = datetime.date(min(seen_dates).year, min(seen_dates).month, 1)
        end = datetime.date(max(seen_dates).year, max(seen_dates).month, 1)
        while d <= end:
            months.append((f'{d.year}-{d.month:02d}', d.strftime('%b-%y')))
            d = (d.replace(day=28) + datetime.timedelta(days=7)).replace(day=1)
    return {'cats': cats, 'totals': totals, 'months': months}


# ---------------------------------------------------------------- normalize + classify
def canon(s):
    s = (s or '').lower().replace('–', '-').replace(':', ' ').replace('.', '')
    s = s.replace('-', ' ').replace('/', ' ').replace('&', 'and')
    s = re.sub(r'\s+', ' ', s).strip()
    out = []
    for w in s.split():
        if not out or out[-1] != w:
            out.append(w)
    s = ' '.join(out)
    al = {'city registration fee': 'city reg', 'dwp general': 'dwp',
          'plumbing drain stoppage': 'drain stoppage', 'repairs repair supplies': 'repair supplies'}
    return al.get(s, s)


_CLASS_RULES = [
    (('management fee',), 'OpEx - Management Fee'),
    (('property tax', 're tax'), 'OpEx - RE Taxes'),
    (('registration', 'permits', 'business license', 'admin'), 'OpEx - G&A'),
    (('mortgage', 'debt service', 'interest paid'), 'Non-OpEx - Mortgage Interest'),
    (('legal', 'professional', 'accounting', 'forensic', 'engineering'), 'Non-OpEx - Legal & Professional'),
    (('escrow', 'prepayment', 'deposit', 'security', 'settlement', 'gain', 'cash back'), 'Non-OpEx - Other'),
    (('leasing fee', 'leasing commission'), 'CapEx - Leasing Commissions'),
    (('dwp', 'gas', 'water', 'sewage', 'electric', 'trash', 'utilit', 'comcast'), 'OpEx - Utilities'),
    (('rent',), 'Income - Rental Income'),
    (('scep', 'city reg', 'move out', 'other income'), 'Income - Other Income'),
    (('depreciation', 'amortization'), 'Non-OpEx - Other'),
]


def classify(label):
    n = canon(label)
    for keys, cls in _CLASS_RULES:
        if any(k in n for k in keys):
            return cls
    return 'OpEx - R&M'


# ---------------------------------------------------------------- assemble
def build_workbook(summaries, detail):
    """summaries: [{'label','rows'}] oldest->newest. detail: {'label','cats','totals','months'} or None."""
    spine_src = summaries[-1]['rows'] if summaries else []
    spine = [dict(r) for r in spine_src]
    spine_items = {canon(r['label']) for r in spine_src if r.get('amount') is not None and not r['total'] and not r.get('net')}

    extras = []
    seen = set(spine_items)
    for sm in summaries[:-1]:
        for r in sm['rows']:
            if r.get('amount') is not None and not r['total'] and not r.get('net') and canon(r['label']) not in seen:
                extras.append(dict(r, _only=sm['label'])); seen.add(canon(r['label']))
    cats = detail['cats'] if detail else {}
    for name in cats:
        if canon(name) not in seen:
            extras.append({'label': name, 'amount': 0.0, '_only': detail['label']}); seen.add(canon(name))
    ins = next((i for i, r in enumerate(spine) if r.get('label', '').lower() == 'total expense'), len(spine))
    spine = spine[:ins] + extras + spine[ins:]

    months = detail['months'] if detail else []
    nM = len(months)
    cats_c = {canon(k): k for k in cats}
    tot_c = {canon(k): k for k in (detail['totals'] if detail else {})}

    def ytd_item(label):
        k = canon(label)
        if k in cats_c:
            return cats_c[k]
        m = difflib.get_close_matches(k, list(cats_c), n=1, cutoff=0.86)
        return cats_c[m[0]] if m else None

    def ytd_total(label):
        name = re.sub(r'^total\s+', '', canon(label))
        if name in tot_c:
            return tot_c[name]
        m = difflib.get_close_matches(name, list(tot_c), n=1, cutoff=0.86)
        return tot_c[m[0]] if m else None

    wb = openpyxl.Workbook()
    comb = wb.active; comb.title = 'Combined'
    stabs = []
    for sm in summaries:
        ws = wb.create_sheet(sm['label'][:31]); ws.append(['Line', sm['label']]); stabs.append((sm, ws))
    dtab = None
    if detail:
        dtab = wb.create_sheet(detail['label'][:31]); dtab.append(['Line'] + [l for _, l in months] + ['YTD Total'])

    # per-summary lookups (canon -> (orig label, amount))
    smaps = [{canon(r['label']): (r['label'], r['amount']) for r in sm['rows'] if r.get('amount') is not None} for sm in summaries]

    def put(ws, row, col, val, red=False, bold=False):
        c = ws.cell(row, col, val)
        c.font = REDB if (red and bold) else RED if red else BOLD if bold else PLAIN
        return c

    # column layout
    CLS = 1
    col = 2
    sum_val_cols = []
    sum_blocks = []          # (summary index, Lcol, Vcol)
    for si, sm in enumerate(summaries):
        sum_blocks.append((si, col, col + 1)); sum_val_cols.append(col + 1); col += 3  # label,val,spacer
    det_L = det_M0 = det_T = None
    if detail:
        det_L, det_M0 = col, col + 1
        det_T = det_M0 + nM
        col = det_T + 2                       # + spacer
    TOTAL = col; col += 2
    RECON = col

    # header
    comb.cell(1, CLS, 'Classification')
    for si, Lc, Vc in sum_blocks:
        comb.cell(1, Lc, summaries[si]['label'])
    if detail:
        comb.cell(1, det_L, detail['label'])
        for j, (_, lbl) in enumerate(months):
            comb.cell(1, det_M0 + j, lbl)
        comb.cell(1, det_T, 'YTD Total')
    comb.cell(1, TOTAL, 'Total (all)')
    comb.cell(1, RECON, 'Recon')

    for i, r in enumerate(spine):
        row = i + 2
        label = r['label']
        is_tot = bool(r.get('total') or r.get('net')); is_sec = bool(r.get('section'))

        # write source tabs
        for si, (sm, ws) in enumerate(stabs):
            m = smaps[si].get(canon(label))
            ws.cell(row, 1, m[0] if m else label)
            ws.cell(row, 2, m[1] if m else 0)
        cm = None if is_sec else (ytd_item(label) if detail else None)
        tm = ytd_total(label) if (detail and is_tot) else None
        if detail:
            dtab.cell(row, 1, cm if cm else (tm if tm else label))
            if is_tot and tm:
                dtab.cell(row, 1 + nM + 1, detail['totals'].get(tm))
            elif cm:
                for j, (mk, _) in enumerate(months):
                    dtab.cell(row, 2 + j, round(cats[cm].get(mk, 0.0), 2))
                dtab.cell(row, 1 + nM + 1, round(sum(cats[cm].values()), 2))

        if is_sec:
            for si, Lc, Vc in sum_blocks:
                put(comb, row, Lc, label, bold=True)
            if detail:
                put(comb, row, det_L, label, bold=True)
            continue
        if not is_tot:
            comb.cell(row, CLS, classify(label))

        # summary blocks
        for si, Lc, Vc in sum_blocks:
            present = canon(label) in smaps[si] or r.get('_only') == summaries[si]['label']
            put(comb, row, Lc, f"='{summaries[si]['label'][:31]}'!A{row}", red=not present, bold=is_tot)
            put(comb, row, Vc, f"='{summaries[si]['label'][:31]}'!B{row}", red=not present, bold=is_tot).number_format = CUR
        # detail block
        if detail:
            p26 = bool(cm) or bool(tm) or r.get('_only') == detail['label']
            dl = detail['label'][:31]
            put(comb, row, det_L, f"='{dl}'!A{row}", red=not p26, bold=is_tot)
            for j in range(nM):
                put(comb, row, det_M0 + j, f"='{dl}'!{get_column_letter(2 + j)}{row}", red=not p26, bold=is_tot).number_format = CUR
            put(comb, row, det_T, f"='{dl}'!{get_column_letter(2 + nM)}{row}", red=not p26, bold=is_tot).number_format = CUR
        # total across blocks
        parts = [f"{get_column_letter(vc)}{row}" for vc in sum_val_cols] + ([f"{get_column_letter(det_T)}{row}"] if detail else [])
        put(comb, row, TOTAL, '=' + '+'.join(parts), bold=is_tot).number_format = CUR
        if not is_tot and detail:
            comb.cell(row, RECON,
                      f'=IF(ABS({get_column_letter(det_T)}{row}-SUM({get_column_letter(det_M0)}{row}:{get_column_letter(det_M0+nM-1)}{row}))<0.01,"ok","CHECK")')

    _format(comb, sum_blocks, det_L, det_M0, det_T, nM, TOTAL, RECON, detail is not None)
    for _, ws in stabs:
        _format_src(ws, [2])
    if detail:
        _format_src(dtab, list(range(2, 2 + nM + 1)))

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _format_src(ws, vcols):
    for c in range(1, ws.max_column + 1):
        ws.cell(1, c).font = HDR; ws.cell(1, c).fill = HFILL
    ws.freeze_panes = 'B2'; ws.column_dimensions['A'].width = 30
    for r in range(2, ws.max_row + 1):
        for c in vcols:
            ws.cell(r, c).number_format = CUR
    for c in vcols:
        ws.column_dimensions[get_column_letter(c)].width = 11


def _format(comb, sum_blocks, det_L, det_M0, det_T, nM, TOTAL, RECON, has_detail):
    spacers = set()
    for si, Lc, Vc in sum_blocks:
        spacers.add(Vc + 1)
    if has_detail:
        spacers.add(det_T + 1)
    spacers.add(TOTAL + 1)
    for c in range(1, comb.max_column + 1):
        hc = comb.cell(1, c)
        if c in spacers:
            hc.fill = PatternFill(fill_type=None)
            comb.column_dimensions[get_column_letter(c)].width = 2.5
        else:
            hc.font = HDR; hc.fill = HFILL; hc.alignment = Alignment(horizontal='center')
            comb.column_dimensions[get_column_letter(c)].width = 11
    comb.freeze_panes = 'B2'
    comb.column_dimensions['A'].width = 18
    for si, Lc, Vc in sum_blocks:
        comb.column_dimensions[get_column_letter(Lc)].width = 22
    if has_detail:
        comb.column_dimensions[get_column_letter(det_L)].width = 24
