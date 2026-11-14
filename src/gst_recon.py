#!/usr/bin/env python3
"""
gst_recon.py - Monthly GSTR-2B vs Tally Purchase Register reconciliation.

Usage:
    python gst_recon.py --tally PURCHASE_IN_TELLY.xlsx --gstr2b GSTR_2B_PORTAL.xlsx --out RECO_OCT26.xlsx
Options:
    --tolerance 1.0     max Rs difference treated as rounding (per component)
    --date-window 10    max days between Tally booking date and supplier invoice date for fuzzy matches
    --force             write the report even if sanity gates fail

Only dependency: openpyxl.  Exit codes: 0 ok, 2 sanity gate failed (parsing problem), 3 parse error.

DESIGN RULES (see ANTIGRAVITY_PROMPT.md)
 1. Map columns by HEADER NAME, never by position.
 2. Validate every Tally row (value + taxes + round-off = gross). Too many failures => abort (parsing error).
 3. Clean Tally data first: drop cancelled vouchers, route RCM/expense vouchers separately, invoice no. as text.
 4. Match in passes: exact -> contains -> amount/date (same PAN) -> taxable-only.  Record basis of each match.
 5. 2B side: split RCM, net credit notes (debit notes add), warn on amendment/ISD/import sheets, tie to portal summary.
 6. Report money (GST bridge), status, action; formulas for totals.
 7. Never publish a reconciliation if match rate is implausibly low - show a parsing warning instead.
"""
import argparse, re, sys, datetime as dt
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

class ParseError(Exception): pass

# ------------------------------------------------------------------ helpers
def low(s): return re.sub(r'\s+', ' ', str(s)).strip().lower() if s is not None else ''
def norm_inv(s):
    s = re.sub(r'[^A-Z0-9]', '', str(s).upper())
    return s.lstrip('0') or s
def fnum(x):
    if x in (None, '', 'NA', '-'): return 0.0
    try: return float(x)
    except Exception: return 0.0
def pan(g): return g[2:12] if g and len(g) >= 12 else ''
def clean_name(s): return re.sub(r'_x000D_|\r|\n', ' ', str(s or '')).strip()
def clean_vno(s): return str(s or '').strip().rstrip('.,').strip()
def parse_date(x):
    if isinstance(x, dt.datetime): return x.date()
    if isinstance(x, dt.date): return x
    for f in ('%d/%m/%Y', '%d-%m-%Y', '%d-%b-%y', '%d-%b-%Y', '%Y-%m-%d'):
        try: return dt.datetime.strptime(str(x).strip(), f).date()
        except Exception: pass
    return None
GSTIN_RE = re.compile(r'^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]$')

# ------------------------------------------------------------------ Tally
def parse_tally(path, tol_row=1.0):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(rows[:60]):
        L = [low(c) for c in r]
        if 'date' in L and 'particulars' in L and any(x.startswith('voucher no') for x in L):
            hi = i; break
    if hi is None: raise ParseError("Tally: header row (Date / Particulars / Voucher No.) not found")
    hdr = rows[hi]; H = {}
    for j, c in enumerate(hdr):
        if c is not None and low(c) not in H: H[low(c)] = j
    def col(*aliases):
        for a in aliases:
            for k, j in H.items():
                if k == a or k.startswith(a): return j
        raise ParseError(f"Tally: column {aliases} not found. Headers seen: {list(H)}")
    c_date, c_party, c_vt, c_vno = col('date'), col('particulars'), col('voucher type'), col('voucher no')
    c_inv = col('supplier invoice no', 'supplier inv', 'reference no', 'ref no')
    c_gstin = col('gstin/uin', 'gstin')
    c_val, c_gross = col('value'), col('gross total')
    def many(tax):  # tax input columns, e.g. 'CGST INPUT'
        return [j for k, j in H.items() if tax in k and 'input' in k]
    cg, sg, ig = many('cgst'), many('sgst') + many('utgst'), many('igst')
    if not (cg and sg and ig): raise ParseError(f"Tally: CGST/SGST/IGST INPUT columns not all found. Headers: {list(H)}")
    rnd = [j for k, j in H.items() if 'round' in k]
    used = {c_date, c_party, c_vt, c_vno, c_inv, c_gstin, c_val, c_gross, *cg, *sg, *ig, *rnd}
    ledgers = {j: str(hdr[j]).strip() for j in range(len(hdr)) if hdr[j] is not None and j not in used}
    S = lambda r, js: sum(fnum(r[j]) for j in js)
    docs, cancelled, rcm, dq, gtot = [], [], [], [], None
    bad_rows = 0; checked = 0
    for r in rows[hi + 1:]:
        if all(c is None for c in r): continue
        party = clean_name(r[c_party])
        if low(party) == 'grand total':
            gtot = dict(taxable=fnum(r[c_val]), gross=fnum(r[c_gross]), cgst=S(r, cg), sgst=S(r, sg), igst=S(r, ig)); continue
        if r[c_date] is None: continue
        d = dict(date=parse_date(r[c_date]), party=party, vtype=str(r[c_vt] or ''), vno=clean_vno(r[c_vno]),
                 inv=str(r[c_inv] or '').strip(), gstin=str(r[c_gstin] or '').strip().upper(),
                 taxable=fnum(r[c_val]), gross=fnum(r[c_gross]), cgst=S(r, cg), sgst=S(r, sg), igst=S(r, ig), rnd=S(r, rnd),
                 ledgers={n: fnum(r[j]) for j, n in ledgers.items() if fnum(r[j])})
        d['tax'] = d['cgst'] + d['sgst'] + d['igst']; d['ninv'] = norm_inv(d['inv']); d['pan'] = pan(d['gstin'])
        if 'cancel' in low(party): cancelled.append(d); continue
        if 'exp' in low(d['vtype']) or d['vno'].upper().startswith('RCM'): rcm.append(d); continue
        checked += 1
        calc = d['taxable'] + d['tax'] + d['rnd']
        if abs(calc - d['gross']) > tol_row:
            bad_rows += 1; dq.append(('Row arithmetic', d['vno'], f"{party}: value+tax+round-off = {calc:,.2f} but gross = {d['gross']:,.2f}"))
        if not d['gstin']: dq.append(('Missing GSTIN', d['vno'], f"{party}: purchase without supplier GSTIN"))
        elif not GSTIN_RE.match(d['gstin']): dq.append(('Invalid GSTIN format', d['vno'], f"{party}: {d['gstin']}"))
        if not d['inv']: dq.append(('Missing invoice no.', d['vno'], party))
        ig_led = [n for n in d['ledgers'] if 'igst' in n.lower()]
        gst_led = [n for n in d['ledgers'] if 'gst' in n.lower() and 'igst' not in n.lower()]
        if ig_led and (d['cgst'] or d['sgst']): dq.append(('Ledger/tax mismatch', d['vno'], f"{party}: booked in {ig_led[0]} but CGST/SGST charged"))
        if gst_led and d['igst']: dq.append(('Ledger/tax mismatch', d['vno'], f"{party}: booked in {gst_led[0]} but IGST charged"))
        odd = [n for n in d['ledgers'] if n.lower() not in ('purchase gst', 'purchase igst')]
        if odd: dq.append(('Non-standard purchase ledger', d['vno'], f"{party}: amount in '{odd[0]}' - confirm classification"))
        if d['taxable'] > 0 and d['tax'] == 0: dq.append(('Zero tax on taxable purchase', d['vno'], party))
        docs.append(d)
    seen = defaultdict(list)
    for d in docs: seen[(d['gstin'], d['ninv'])].append(d)
    for k, v in seen.items():
        if len(v) > 1 and k[1]: dq.append(('Duplicate invoice in Tally', ', '.join(x['vno'] for x in v), f"{v[0]['party']} inv {v[0]['inv']} booked {len(v)} times"))
    if gtot:
        tot = dict(cgst=sum(d['cgst'] for d in docs), sgst=sum(d['sgst'] for d in docs), igst=sum(d['igst'] for d in docs), taxable=sum(d['taxable'] for d in docs))
        for k in ('cgst', 'sgst', 'igst'):
            if abs(tot[k] - gtot[k]) > 1: dq.append(('Control total', k.upper(), f"Parsed {tot[k]:,.2f} vs Tally Grand Total {gtot[k]:,.2f} (RCM/other vouchers may explain)"))
    return docs, cancelled, rcm, dq, dict(rows=checked, bad=bad_rows)

# ------------------------------------------------------------------ GSTR-2B
def _find_header(rows, needle):
    for i, r in enumerate(rows[:40]):
        if any(low(c).startswith(needle) for c in r): return i
    return None
def _colmap(rows, h):
    main, sub = rows[h], rows[h + 1] if h + 1 < len(rows) else ()
    def find(pat, prefer_sub=True, nth=0):
        hits = []
        for j in range(len(main)):
            for src in ((sub, main) if prefer_sub else (main, sub)):
                if j < len(src) and src[j] is not None and re.search(pat, low(src[j])):
                    hits.append(j); break
        if len(hits) <= nth: raise ParseError(f"GSTR-2B: column /{pat}/ not found")
        return hits[nth]
    return find
def parse_2b(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    meta, warns = {}, []
    if 'Read me' in wb.sheetnames:
        for r in list(wb['Read me'].iter_rows(values_only=True))[:15]:
            if r[0] and any(r[1:3]): meta[low(r[0])] = [x for x in r[1:3] if x][0]
    def sheet_rows(n): return list(wb[n].iter_rows(values_only=True)) if n in wb.sheetnames else None
    inv, cdn = [], []
    rows = sheet_rows('B2B')
    if rows is None: raise ParseError("GSTR-2B: sheet 'B2B' not found")
    h = _find_header(rows, 'gstin of supplier'); f = _colmap(rows, h)
    cols = dict(g=f(r'gstin of supplier', False), name=f(r'trade/legal'), inv=f(r'invoice number'), typ=f(r'invoice type'), date=f(r'invoice date'),
                gross=f(r'invoice value'), rc=f(r'reverse charge', False), tv=f(r'taxable value', False), ig=f(r'integrated tax'), cg=f(r'central tax'),
                sg=f(r'state/ut tax'), itc=f(r'itc availability', False))
    for r in rows[h + 2:]:
        if not r[cols['g']] or not GSTIN_RE.match(str(r[cols['g']]).strip().upper()): continue
        d = dict(gstin=str(r[cols['g']]).strip().upper(), party=clean_name(r[cols['name']]), inv=str(r[cols['inv']]).strip(), date=parse_date(r[cols['date']]),
                 gross=fnum(r[cols['gross']]), taxable=fnum(r[cols['tv']]), igst=fnum(r[cols['ig']]), cgst=fnum(r[cols['cg']]), sgst=fnum(r[cols['sg']]),
                 rcm=low(r[cols['rc']]) == 'yes', itc=str(r[cols['itc']] or ''), kind='INV')
        d['tax'] = d['igst'] + d['cgst'] + d['sgst']; d['ninv'] = norm_inv(d['inv']); d['pan'] = pan(d['gstin']); d['used'] = False
        inv.append(d)
    rows = sheet_rows('B2B-CDNR')
    if rows:
        h = _find_header(rows, 'gstin of supplier')
        if h is not None:
            f = _colmap(rows, h)
            c = dict(g=f(r'gstin of supplier', False), name=f(r'trade/legal'), no=f(r'note number'), typ=f(r'note type'), date=f(r'note date'), gross=f(r'note value'),
                     tv=f(r'taxable value', False), ig=f(r'integrated tax'), cg=f(r'central tax'), sg=f(r'state/ut tax'))
            for r in rows[h + 2:]:
                if not r[c['g']] or not GSTIN_RE.match(str(r[c['g']]).strip().upper()): continue
                sign = -1 if 'credit' in low(r[c['typ']]) else 1
                d = dict(gstin=str(r[c['g']]).strip().upper(), party=clean_name(r[c['name']]), inv=str(r[c['no']]).strip(), date=parse_date(r[c['date']]),
                         gross=sign * fnum(r[c['gross']]), taxable=sign * fnum(r[c['tv']]), igst=sign * fnum(r[c['ig']]), cgst=sign * fnum(r[c['cg']]),
                         sgst=sign * fnum(r[c['sg']]), rcm=False, itc='Yes', kind='CN' if sign < 0 else 'DN', note=str(r[c['typ']]))
                d['tax'] = d['igst'] + d['cgst'] + d['sgst']; d['pan'] = pan(d['gstin']); cdn.append(d)
    # other sheets that should be empty -> warn if they carry data
    for n in ('B2BA', 'B2B-CDNRA', 'ISD', 'ISDA', 'IMPG', 'IMPGA', 'IMPGSEZ', 'IMPGSEZA', 'ECO', 'ECOA'):
        rws = sheet_rows(n)
        if rws:
            k = sum(1 for r in rws if r and r[0] and GSTIN_RE.match(str(r[0]).strip().upper()))
            if k: warns.append(f"GSTR-2B sheet '{n}' has {k} record(s) - NOT reconciled by this script; review manually")
    for n in ('ITC not available', 'ITC Reversal', 'ITC Rejected'):
        rws = sheet_rows(n)
        if rws:
            amt = sum(abs(fnum(x)) for r in rws if r and str(r[0]).strip() in ('I', 'II', 'III', 'IV') for x in r[3:7])
            if amt: warns.append(f"GSTR-2B sheet '{n}' shows Rs {amt:,.2f} of tax - review (not included in ITC bridge)")
    # portal control total from 'ITC Available'
    control, part = None, None
    rws = sheet_rows('ITC Available')
    if rws:
        a = b = 0.0
        for r in rws:
            k = str(r[0]).strip() if r and r[0] is not None else ''
            if k.startswith('Part A'): part = 'A'
            elif k.startswith('Part B'): part = 'B'
            elif k in ('I', 'II', 'III', 'IV') and part:
                v = sum(fnum(x) for x in r[3:7])
                if part == 'A': a += v
                else: b += v
        control = round(a - b, 2)
    return inv, cdn, meta, warns, control

# ------------------------------------------------------------------ matching
S_M, S_MI, S_MA = 'Matched', 'Matched - invoice no. differs', 'Matched - amount differs'
S_TO, S_BE, S_BO = 'In Tally, not in GSTR-2B', 'In 2B, not in Tally - existing supplier', 'In 2B, not in Tally - other vendor'
S_RC, S_CN = 'In 2B - reverse charge (RCM)', 'Credit/debit note in 2B'
STATUSES = [S_M, S_MI, S_MA, S_TO, S_BE, S_BO, S_RC, S_CN]
def reconcile(tally, b2b, tol, window):
    pairs = []  # (t, b, basis)
    un = list(tally)
    avail = [b for b in b2b if not b['rcm']]
    # pass 1 exact
    idx = defaultdict(list)
    for b in avail: idx[(b['gstin'], b['ninv'])].append(b)
    rest = []
    for t in un:
        c = [b for b in idx.get((t['gstin'], t['ninv']), []) if not b['used']]
        if c and t['ninv']: c[0]['used'] = True; pairs.append((t, c[0], 'Exact GSTIN + invoice no.'))
        else: rest.append(t)
    # pass 2 contains
    un, rest = rest, []
    for t in un:
        c = [b for b in avail if not b['used'] and b['gstin'] == t['gstin'] and min(len(b['ninv']), len(t['ninv'])) >= 3
             and (t['ninv'] in b['ninv'] or b['ninv'] in t['ninv']) and abs(t['tax'] - b['tax']) <= tol * 2]
        if c: c[0]['used'] = True; pairs.append((t, c[0], 'Invoice no. partial match (same GSTIN, tax agrees)'))
        else: rest.append(t)
    def dd(t, b): return abs((t['date'] - b['date']).days) if t['date'] and b['date'] else 0
    # pass 3 same PAN, tax+gross agree, date window
    un, rest = rest, []
    for t in un:
        c = [b for b in avail if not b['used'] and b['pan'] == t['pan'] and t['pan'] and abs(t['tax'] - b['tax']) <= tol
             and abs(t['gross'] - b['gross']) <= tol + 0.5 and dd(t, b) <= window]
        c.sort(key=lambda b: (b['gstin'] != t['gstin'], dd(t, b), abs(t['tax'] - b['tax'])))
        if c: c[0]['used'] = True; pairs.append((t, c[0], 'Amount + date match (same PAN) - invoice no. differs'))
        else: rest.append(t)
    # pass 4 same PAN, taxable agrees, tax differs
    un, rest = rest, []
    for t in un:
        c = [b for b in avail if not b['used'] and b['pan'] == t['pan'] and t['pan'] and abs(t['taxable'] - b['taxable']) <= tol and dd(t, b) <= window]
        c.sort(key=lambda b: (b['gstin'] != t['gstin'], dd(t, b)))
        if c: c[0]['used'] = True; pairs.append((t, c[0], 'Taxable value match only (same PAN) - tax differs'))
        else: rest.append(t)
    return pairs, rest

def build_rows(tally, b2b, cdn, tol, window, month_end):
    pairs, tally_only = reconcile(tally, b2b, tol, window)
    rows = []
    def R(status, basis, party, gstin, t, b, remark):
        rows.append(dict(status=status, basis=basis, party=party, gstin=gstin, tdate=t['date'] if t else None, vno=t['vno'] if t else '', tinv=t['inv'] if t else '',
                         binv=b['inv'] if b else '', bdate=b['date'] if b else None, ttv=t['taxable'] if t else 0, ttax=t['tax'] if t else 0,
                         btv=b['taxable'] if b else 0, btax=b['tax'] if b else 0, remark=remark))
    for t, b, basis in pairs:
        ok = all(abs(x - y) <= tol for x, y in ((t['taxable'], b['taxable']), (t['igst'], b['igst']), (t['cgst'], b['cgst']), (t['sgst'], b['sgst'])))
        exact = basis.startswith('Exact')
        st = S_MA if not ok else (S_M if exact else S_MI); rem = []
        if not exact: rem.append(f"Tally inv no. '{t['inv']}' vs portal '{b['inv']}' - correct Tally entry")
        if not ok:
            if abs(t['igst'] - b['igst']) > tol or abs(t['cgst'] - b['cgst']) > tol: rem.append('GST split (IGST vs CGST/SGST) or amount differs - verify place of supply / invoice')
            else: rem.append('Taxable value differs - verify invoice')
        if b['itc'].lower() != 'yes': rem.append(f"ITC availability per 2B = '{b['itc']}'")
        R(st, basis, t['party'], t['gstin'], t, b, '; '.join(rem))
    for t in tally_only:
        R(S_TO, '', t['party'], t['gstin'], t, None, 'Supplier has not reported in GSTR-1 - ITC NOT claimable until it appears; follow up with supplier')
    tally_pans = {t['pan'] for t in tally if t['pan']}
    cn_by = defaultdict(list)
    for c in cdn: cn_by[c['gstin']].append(c)
    for b in b2b:
        if b['used']: continue
        late = b['date'] and (month_end - b['date']).days <= 3
        if b['rcm']: R(S_RC, '', b['party'], b['gstin'], None, b, 'Reverse-charge: pay tax in GSTR-3B 3.1(d); ITC only after cash payment; book in Tally'); continue
        off = [c for c in cn_by.get(b['gstin'], []) if abs(abs(c['gross']) - b['gross']) <= tol + 0.5]
        if b['pan'] in tally_pans:
            rem = 'Existing supplier - invoice not booked in this month\'s register: check goods receipt / book it' + (' (dated within last 3 days of month: likely cut-off)' if late else '')
            st = S_BE
        else:
            rem = 'Supplier not in Purchase Register - check expense/journal/payment/capital vouchers; book if genuine'; st = S_BO
        if off: rem = f"Offset by credit note {off[0]['inv']} of same value - book both purchase and return"
        if b['itc'].lower() != 'yes': rem += f"; ITC availability per 2B = '{b['itc']}'"
        R(st, '', b['party'], b['gstin'], None, b, rem)
    for c in cdn:
        R(S_CN, '', c['party'], c['gstin'], None, c, f"{c['note']} in 2B {'reduces' if c['tax'] < 0 else 'increases'} ITC; no matching entry in Purchase Register - verify debit/credit note booked")
    order = {s: i for i, s in enumerate(STATUSES)}
    rows.sort(key=lambda r: (order[r['status']], r['party'].upper(), r['binv'] or r['tinv']))
    return rows

# ------------------------------------------------------------------ report
F = 'Arial'; NUM = '#,##0.00;(#,##0.00);-'
HDR = PatternFill('solid', fgColor='1F3864')
FILL = {S_M: 'E2EFDA', S_MI: 'FFF2CC', S_MA: 'FFE699', S_TO: 'F8CBAD', S_BE: 'DDEBF7', S_BO: 'DDEBF7', S_RC: 'EDEDED', S_CN: 'EDEDED'}
def hdr(ws, r, vals):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.font = Font(name=F, bold=True, color='FFFFFF'); c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical='center')
def write_report(path, rows, tally, cancelled, rcm, dq, meta, warns, control, stats, tol):
    wb = openpyxl.Workbook(); S = wb.active; S.title = 'Summary'; D = wb.create_sheet('Reco Detail'); Q = wb.create_sheet('Data Quality')
    hdr(D, 1, ['Status', 'Match basis', 'Supplier', 'GSTIN', 'Tally Date', 'Tally Voucher', 'Tally Inv No.', 'GSTR-2B Inv No.', 'GSTR-2B Date', 'Tally Taxable', 'Tally GST',
               '2B Taxable', '2B GST', 'Diff Taxable (Tally-2B)', 'Diff GST (Tally-2B)', 'Remark / Action'])
    for r, w in enumerate(rows, 2):
        vals = [w['status'], w['basis'], w['party'], w['gstin'], w['tdate'], w['vno'], w['tinv'], w['binv'], w['bdate'], w['ttv'], w['ttax'], w['btv'], w['btax']]
        for i, v in enumerate(vals, 1):
            c = D.cell(r, i, v if v not in ('', None) else None); c.font = Font(name=F, size=10)
            if i in (5, 9): c.number_format = 'dd-mmm-yy'
            if i >= 10: c.number_format = NUM
        D.cell(r, 14, f'=J{r}-L{r}').number_format = NUM; D.cell(r, 15, f'=K{r}-M{r}').number_format = NUM; D.cell(r, 16, w['remark'])
        for i in (14, 15, 16): D.cell(r, i).font = Font(name=F, size=10)
        D.cell(r, 1).fill = PatternFill('solid', fgColor=FILL[w['status']])
    n = len(rows) + 1
    D.cell(n + 1, 3, 'TOTAL').font = Font(name=F, bold=True)
    for col in 'JKLMNO':
        c = D[f'{col}{n + 1}']; c.value = f'=SUM({col}2:{col}{n})'; c.number_format = NUM; c.font = Font(name=F, bold=True)
    for i, w in enumerate([36, 40, 34, 18, 11, 16, 20, 24, 12, 14, 13, 14, 13, 14, 13, 90], 1): D.column_dimensions[CL(i)].width = w
    D.freeze_panes = 'D2'; D.auto_filter.ref = f'A1:P{n}'
    rg = lambda c: f"'Reco Detail'!${c}$2:${c}${n}"
    S['A1'] = f"{meta.get('legal name', 'Company')} - GST Reconciliation: Tally Purchase Register vs GSTR-2B"; S['A1'].font = Font(name=F, size=14, bold=True)
    S['A2'] = f"GSTIN {meta.get('gstin', '')} | Period: {meta.get('tax period', '')} {meta.get('financial year', '')} | Rounding tolerance Rs {tol:g}"; S['A2'].font = Font(name=F, size=10, italic=True)
    hdr(S, 4, ['A. Status-wise reconciliation', 'No. of docs', 'Tally Taxable', 'Tally GST', '2B Taxable', '2B GST', 'Diff GST (Tally-2B)'])
    r0 = 5
    for k, s in enumerate(STATUSES):
        r = r0 + k; S.cell(r, 1, s); S.cell(r, 2, f'=COUNTIF({rg("A")},A{r})')
        for col, src in zip('CDEF', 'JKLM'): S[f'{col}{r}'] = f'=SUMIFS({rg(src)},{rg("A")},$A{r})'
        S[f'G{r}'] = f'=D{r}-F{r}'
    rt = r0 + len(STATUSES); S.cell(rt, 1, 'TOTAL')
    for col in 'BCDEFG': S[f'{col}{rt}'] = f'=SUM({col}{r0}:{col}{rt-1})'
    for r in range(r0, rt + 1):
        for col in 'ABCDEFG':
            c = S[f'{col}{r}']; c.font = Font(name=F, size=10, bold=(r == rt))
            if col != 'A': c.number_format = '#,##0' if col == 'B' else NUM
    b0 = rt + 3; hdr(S, b0 - 1, ['B. ITC bridge (GST amount)', '', '', '', '', 'Amount (Rs)', ''])
    i = {s: r0 + k for k, s in enumerate(STATUSES)}
    br = [('GST per Tally Purchase Register', f'=D{rt}'),
          ('Less: in Tally, not in GSTR-2B (supplier default)', f'=-D{i[S_TO]}'),
          ('Add: existing suppliers in 2B, not booked in Tally', f'=F{i[S_BE]}'),
          ('Add: other vendors in 2B, not in Purchase Register', f'=F{i[S_BO]}'),
          ('Add: reverse-charge items in 2B', f'=F{i[S_RC]}'),
          ('Add/Less: differences on matched invoices (2B - Tally)', f'=SUM(F{i[S_M]}:F{i[S_MA]})-SUM(D{i[S_M]}:D{i[S_MA]})'),
          ('Add/Less: credit & debit notes in 2B', f'=F{i[S_CN]}'),
          ('GST per GSTR-2B (computed)', f'=SUM(F{b0}:F{b0+6})')]
    for k, (a, fm) in enumerate(br): S.cell(b0 + k, 1, a); S.cell(b0 + k, 6, fm)
    S.cell(b0 + 8, 1, 'GST per GSTR-2B portal summary (ITC Available sheet, Part A less Part B)'); S.cell(b0 + 8, 6, control if control is not None else 'n/a')
    S.cell(b0 + 9, 1, 'Difference (must be nil)'); S.cell(b0 + 9, 6, f'=IF(ISNUMBER(F{b0+8}),F{b0+7}-F{b0+8},"n/a")')
    for r in range(b0, b0 + 10):
        bold = r in (b0 + 7, b0 + 9); S.cell(r, 1).font = Font(name=F, size=10, bold=bold)
        S.cell(r, 6).font = Font(name=F, size=10, bold=bold, color='0000FF' if r == b0 + 8 else '000000'); S.cell(r, 6).number_format = NUM
    # auto findings
    cnt = lambda s: [w for w in rows if w['status'] == s]
    sm = lambda s: sum(w['btax'] - w['ttax'] if s != S_TO else w['ttax'] for w in cnt(s))
    f0 = b0 + 12; hdr(S, f0 - 1, ['C. Findings & actions (auto-generated)', '', '', '', '', '', ''])
    nm = len(cnt(S_M)) + len(cnt(S_MI)) + len(cnt(S_MA))
    F_ = [f"1. {nm} of {len(tally)} Tally purchase vouchers matched to GSTR-2B ({len(cnt(S_MI))} with invoice-no. differences, {len(cnt(S_MA))} with amount differences)."]
    if cnt(S_TO): F_.append(f"2. ITC AT RISK Rs {sum(w['ttax'] for w in cnt(S_TO)):,.2f}: {len(cnt(S_TO))} Tally invoice(s) not in GSTR-2B - " + '; '.join(f"{w['party']} {w['tinv']} (Rs {w['ttax']:,.2f})" for w in cnt(S_TO)) + '. Do not claim until supplier reports.')
    if cnt(S_BE): F_.append(f"3. {len(cnt(S_BE))} invoice(s) of existing suppliers are in 2B but not booked in Tally (GST Rs {sum(w['btax'] for w in cnt(S_BE)):,.2f}) - check goods receipt / month-end cut-off and book.")
    if cnt(S_BO): F_.append(f"4. {len(cnt(S_BO))} invoice(s) of suppliers not in the Purchase Register (GST Rs {sum(w['btax'] for w in cnt(S_BO)):,.2f}) - check expense/journal vouchers; book if genuine.")
    if cnt(S_CN): F_.append(f"5. {len(cnt(S_CN))} credit/debit note(s) in 2B (net GST Rs {sum(w['btax'] for w in cnt(S_CN)):,.2f}) - confirm purchase returns are booked.")
    if cnt(S_RC): F_.append(f"6. {len(cnt(S_RC))} reverse-charge document(s) in 2B (GST Rs {sum(w['btax'] for w in cnt(S_RC)):,.2f}) - pay under 3.1(d) before taking ITC.")
    F_.append(f"7. Data quality: {len(dq)} item(s) on 'Data Quality' sheet; {len(cancelled)} cancelled voucher(s); {len(rcm)} RCM/expense voucher(s) excluded from matching.")
    for wmsg in warns: F_.append('WARNING: ' + wmsg)
    F_.append('Matching: GSTIN + normalised invoice no.; then partial invoice no.; then same-PAN amount/date; then taxable-only. Basis is shown per row.')
    for k, t in enumerate(F_):
        c = S.cell(f0 + k, 1, t); c.font = Font(name=F, size=10, color='C00000' if t.startswith('WARNING') else '000000'); c.alignment = Alignment(wrap_text=True, vertical='top')
        S.merge_cells(start_row=f0 + k, start_column=1, end_row=f0 + k, end_column=7); S.row_dimensions[f0 + k].height = 44
    S.column_dimensions['A'].width = 62
    for c in 'BCDEFG': S.column_dimensions[c].width = 16
    hdr(Q, 1, ['Type', 'Reference', 'Detail'])
    items = [(('Cancelled voucher', d['vno'], f"Supplier inv {d['inv'] or '-'} - excluded") for d in cancelled), (('RCM / expense voucher', d['vno'], f"{d['party']} inv {d['inv']} amount {d['gross']:,.2f} GSTIN {d['gstin'] or 'none'} - RCM tax to be paid in 3.1(d)") for d in rcm), iter(dq), ((('Warning', '', w)) for w in warns)]
    r = 2
    for it in items:
        for row in it:
            for j, v in enumerate(row, 1): Q.cell(r, j, v).font = Font(name=F, size=10)
            r += 1
    for j, w in enumerate([30, 30, 120], 1): Q.column_dimensions[CL(j)].width = w
    wb.save(path)

# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description='GSTR-2B vs Tally purchase register reconciliation')
    ap.add_argument('--tally', required=True); ap.add_argument('--gstr2b', required=True); ap.add_argument('--out', required=True)
    ap.add_argument('--tolerance', type=float, default=1.0); ap.add_argument('--date-window', type=int, default=10)
    ap.add_argument('--force', action='store_true'); a = ap.parse_args()
    try:
        tally, cancelled, rcm, dq, st = parse_tally(a.tally)
        b2b, cdn, meta, warns, control = parse_2b(a.gstr2b)
    except ParseError as e:
        print('PARSE ERROR:', e, file=sys.stderr); return 3
    print(f"Tally: {len(tally)} purchase vouchers, {len(cancelled)} cancelled, {len(rcm)} RCM/expense | GSTR-2B: {len(b2b)} invoices, {len(cdn)} notes")
    gates = []
    if st['rows'] and st['bad'] / st['rows'] > 0.2: gates.append(f"{st['bad']} of {st['rows']} Tally rows fail value+tax=gross check -> column mapping/parsing problem")
    months = [d['date'] for d in tally if d['date']]
    ref = max(months) if months else dt.date.today()
    month_end = (ref.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    rows = build_rows(tally, b2b, cdn, a.tolerance, a.date_window, month_end)
    matched = sum(1 for w in rows if w['status'].startswith('Matched'))
    base = min(len(tally), len([b for b in b2b if not b['rcm']]))
    if base and matched / base < 0.2: gates.append(f"match rate only {matched}/{base} - almost certainly a parsing/period/company mismatch, not a real reconciliation result")
    if gates and not a.force:
        print('SANITY GATE FAILED (no report written):\n  - ' + '\n  - '.join(gates) + '\nRe-run with --force to override.', file=sys.stderr); return 2
    write_report(a.out, rows, tally, cancelled, rcm, dq, meta, warns + [f'SANITY: {g}' for g in gates], control, st, a.tolerance)
    tot = defaultdict(float); cntd = defaultdict(int)
    for w in rows: cntd[w['status']] += 1
    print('Status counts:', dict(cntd)); print('Report written:', a.out)
    return 0
if __name__ == '__main__': sys.exit(main())
