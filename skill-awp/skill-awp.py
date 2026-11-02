"""
skill-awp.py  v2.0 - internal-audit work program (AWP) builder, composer, validator and learner.

Layer 1 of the skill (changed only through /awp-release). Commands:

    python skill-awp.py selftest                                        proves the environment and prints the versions (run first)
    python skill-awp.py template  OUT.xlsx [--tests N] [--walkthroughs N]   blank AWP with [placeholders] (defaults 3 tests, 2 walkthroughs)
    python skill-awp.py compose   SELECTION.json SPEC.json              core + packs + tokens -> engagement spec (prints content warnings)
    python skill-awp.py validate  FILE.json [FILE.json ...]             checks a spec or a library (content, references, tokens, duplicates)
    python skill-awp.py build     SPEC.json OUT.xlsx                    spec -> workbook, stamped
    python skill-awp.py export    AWP.xlsx SPEC.json                    workbook -> spec (header-driven, tolerant of fieldwork edits)
    python skill-awp.py learn     AWP.xlsx PROPOSALS.json [--tokens T.json]   anonymised proposals vs the libraries (+ PROPOSALS.md)
    python skill-awp.py apply     PROPOSALS.json DECISIONS.json         merge accepted proposals into the libraries (version bump, changelog)
    python skill-awp.py diff      A.xlsx B.xlsx                         cell-by-cell comparison (release regression)
    python skill-awp.py check     [--deny names.txt]                   release acceptance: selftest, validate, reference build, learn idempotence, leak scan
    python skill-awp.py list      [PACK ...]                            ids a selection can use: sheets, modules, risks, tests, controls, questions, tokens
    python skill-awp.py verify    AWP.xlsx                              step-7 checks on a built workbook (references, step columns, placeholders, formulas)

    --lib FOLDER on compose / validate / learn / apply / list / selftest / check adds a folder of libraries (e.g. a draft pack); core is always found next to this script.

Libraries are discovered by name: skill-awp-lib-core.json and skill-awp-lib-pack-*.json. Only dependency: openpyxl.
Formatting is fixed here - never restyle a built workbook by hand.
"""
import json, math, re, sys, os, glob, datetime as dt, difflib
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

VERSION = '2.0'
HERE = os.path.dirname(os.path.abspath(__file__))
try: sys.stdout.reconfigure(errors='replace')          # never crash on a console that cannot print a character
except Exception: pass
TOKENS = ['{CLIENT}', '{CLIENT_SHORT}', '{PROCESS}', '{DEPARTMENT}', '{HEAD_OF_DEPARTMENT}', '{MANAGER}', '{STAFF}', '{FINANCE}', '{FINANCE_PREPARER}', '{FINANCE_REVIEWER}', '{HEAD_OF_FINANCE}', '{CEO}', '{IT}', '{SYSTEM}', '{ERP}', '{PERIOD}',
          '{COUNTERPARTY}', '{COUNTERPARTIES}', '{COUNTERPARTY_TITLE}', '{AGREEMENT}', '{MASTER_FILE}', '{CURRENCY}', '{RECON_BALANCES}', '{INTERFACE_FREQUENCY}']
# packs may declare further tokens in their token_defaults; the validator accepts those too.

# =============================================================================== styles (house format)
FN = 'Arial'; SZ = 8
def F(bold=False, color=None, size=SZ, italic=False): return Font(name=FN, size=size, bold=bold, color=color, italic=italic)
RED = 'FFFF0000'; WHITE = 'FFFFFFFF'; GREY_TXT = 'FF808080'
FILL = lambda rgb: PatternFill('solid', start_color=rgb)
F_HDR = FILL('D9D9D9'); F_SECTION = FILL('DCE6F1'); F_BAND = FILL('808080'); F_DETAIL = FILL('B7DEE8'); F_STEPS = FILL('92D050'); F_REF = FILL('FFFF00'); F_BLACK = FILL('000000'); F_ACCENT = FILL('4BACC6')
THIN = Side(style='thin'); DOUBLE = Side(style='double')
B_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN); B_DOUBLE = Border(left=DOUBLE, right=DOUBLE, top=DOUBLE, bottom=DOUBLE); B_BOTTOM = Border(bottom=THIN)
TOP = Alignment(vertical='top'); TOPW = Alignment(vertical='top', wrap_text=True); LEFTW = Alignment(horizontal='left', vertical='top', wrap_text=True)
CTR = Alignment(horizontal='center', vertical='top'); CTRW = Alignment(horizontal='center', vertical='top', wrap_text=True); CTRM = Alignment(horizontal='center', vertical='center', wrap_text=True)
TAB_GREEN = '92D050'; TAB_BLACK = '000000'; LINE_PT = 11.25

def ph(v, text): return v if (v is not None and str(v).strip() != '') else f'[{text}]'
def est_lines(text, width_chars):
    if text is None: return 1
    return sum(max(1, math.ceil(len(p) / max(width_chars * 1.15, 1))) for p in str(text).split('\n'))
def fit_height(ws, row, cells_widths, min_h=24.75):
    lines = max([est_lines(v, w) for v, w in cells_widths] + [1]); ws.row_dimensions[row].height = max(min_h, min(409, lines * LINE_PT + 3))
def put(ws, ref, value, font=None, fill=None, align=None, border=None):
    c = ws[ref]; c.value = value; c.font = font or F()
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = border
    return c
def ref_box(ws, col_letter, ref): put(ws, f'{col_letter}1', ref, F(True), F_REF, CTRM, B_DOUBLE); ws.row_dimensions[1].height = 25.5
def widths(ws, d):
    for k, v in d.items(): ws.column_dimensions[k].width = v
def finish(ws, tab=TAB_GREEN): ws.sheet_view.showGridLines = False; ws.sheet_properties.tabColor = tab

def header(ws, spec, index_sheet, labels=('Client', 'Process', 'Audit period '), signoff_col=None):
    for i, lab in enumerate(labels):
        put(ws, f'A{3+i}', lab, F(True))
        if ws.title == index_sheet:
            put(ws, f'B{3+i}', [spec.get('client', ''), spec.get('process', ''), spec.get('period', '')][i] or f'[{lab.strip()}]', F(), align=TOP)
        else:
            put(ws, f'B{3+i}', f"='{index_sheet}'!B{3+i}", F(), align=TOP)
    if signoff_col:
        c1 = column_index_from_string(signoff_col)
        put(ws, f'{get_column_letter(c1)}3', 'Prepared by:', F(True)); put(ws, f'{get_column_letter(c1+2)}3', 'Date:', F(True))
        put(ws, f'{get_column_letter(c1)}5', 'Reviewed by:', F(True)); put(ws, f'{get_column_letter(c1+2)}5', 'Date:', F(True))
        put(ws, f'{get_column_letter(c1+1)}3', spec.get('prepared_by', ''), F()); put(ws, f'{get_column_letter(c1+1)}5', spec.get('reviewed_by', ''), F())

# =============================================================================== sheet builders
def build_scope(wb, spec):
    ws = wb.create_sheet('Scope')
    for col, h in zip('ABCD', ['No', 'Key Process', 'Indicative area of focus', 'Covered?']):
        put(ws, f'{col}1', h, F(True, WHITE if col != 'D' else None), F_BLACK if col != 'D' else F_ACCENT, CTRW, B_THIN)
    r = 2
    for i, kp in enumerate(spec.get('scope') or [{'key_process': None, 'areas': [None, None], 'covered': 'Yes'}], 1):
        areas = kp.get('areas') or [None]; start = r
        for a in areas:
            put(ws, f'A{r}', kp.get('no', i), F(), align=CTRW, border=B_THIN); put(ws, f'B{r}', ph(kp.get('key_process'), 'Key process'), F(), align=LEFTW, border=B_THIN)
            put(ws, f'C{r}', ph(a, 'Indicative area of focus - one line per area'), F(), align=LEFTW, border=B_THIN); put(ws, f'D{r}', kp.get('covered', 'Yes'), F(), align=CTR, border=B_THIN)
            fit_height(ws, r, [(ws[f'C{r}'].value, 80)]); r += 1
        if len(areas) > 1: ws.merge_cells(f'A{start}:A{r-1}'); ws.merge_cells(f'B{start}:B{r-1}')
    widths(ws, {'A': 6, 'B': 49.1, 'C': 80, 'D': 42.4}); finish(ws, TAB_BLACK)

def build_index(wb, spec, P, name, title, items, ref_col, w, stamp=None):
    ws = wb.create_sheet(name); ref_box(ws, ref_col, name)
    header(ws, spec, P, signoff_col='D' if (name == P and spec.get('signoff', True)) else None)
    put(ws, 'B7', title, F(True), align=TOPW, border=B_BOTTOM); put(ws, f'{ref_col}7', 'Ref', F(True), align=CTRW, border=B_BOTTOM)
    r = 9
    for i, (txt, ref) in enumerate(items, 1):
        put(ws, f'A{r}', i, F(), align=CTR); put(ws, f'B{r}', txt, F(), align=LEFTW); put(ws, f'{ref_col}{r}', ref, F(True, RED), align=CTRM)
        fit_height(ws, r, [(txt, w['B'])], min_h=13.5); r += 2
    if stamp:
        put(ws, f'B{r+1}', stamp, F(color=GREY_TXT, size=7, italic=True))
    widths(ws, w); finish(ws); return ws

def build_rcm(wb, spec, P, test_conclusion_rows):
    name = f'{P}100'; ws = wb.create_sheet(name); ref_box(ws, 'H', name); header(ws, spec, P, labels=('Client:', 'Process:', 'Review period:'))
    hdrs = ['S/N', 'Sub-process', 'Risk', 'Control No.', 'Finding for Design Effectiveness', 'Audit Strategy', 'Testing Result', 'WP Ref']
    for j, h in enumerate(hdrs): put(ws, f'{get_column_letter(j+1)}7', h, F(True), F_HDR, Alignment(horizontal='left' if j == 1 else 'center', vertical='center', wrap_text=True), B_THIN)
    ws.row_dimensions[7].height = 30.75; r = 8; sn = 1
    for k, sec in enumerate(spec.get('rcm', {}).get('sections') or [{'title': None, 'rows': [{}, {}]}], 1):
        put(ws, f'A{r}', ph(sec.get('title'), f'{k}. Section title'), F(True), F_SECTION, TOP, B_THIN)
        for j in range(2, 9): ws[f'{get_column_letter(j)}{r}'].border = B_THIN; ws[f'{get_column_letter(j)}{r}'].fill = F_SECTION
        ws.merge_cells(f'A{r}:H{r}'); ws.row_dimensions[r].height = 22.5; r += 1
        for row in sec.get('rows') or [{}]:
            wp = row.get('wp_ref') or ''
            tr = row.get('testing_result') or ''
            refs = [x.strip() for x in str(wp).split('\n') if x.strip() in test_conclusion_rows]
            if not tr and spec.get('link_results', True) and refs:
                cells = [f"'{x}'!B{test_conclusion_rows[x]}" for x in refs]
                tr = f'=IF({cells[0]}="","",{cells[0]})' if len(cells) == 1 else '=' + '&CHAR(10)&'.join(cells)
            vals = [row.get('sn', sn), ph(row.get('sub_process'), 'Sub-process'), ph(row.get('risk'), 'Risk - what could go wrong and the consequence'), ph(row.get('control_no'), f'{P}2xx-n'),
                    row.get('de_finding') or 'To be concluded after the walkthrough.', ph(row.get('audit_strategy'), 'Test n: title\nCheck that:\n1. ...'), tr, ph(wp, f'{P}3xx')]
            aligns = [CTR, LEFTW, LEFTW, CTRW, LEFTW, LEFTW, LEFTW, CTRW]
            for j, (v, al) in enumerate(zip(vals, aligns)): put(ws, f'{get_column_letter(j+1)}{r}', v, F(True, RED) if j in (3, 7) else F(), align=al, border=B_THIN)
            fit_height(ws, r, [(vals[2], 69.6), (vals[4], 67.3), (vals[5], 83.1), (vals[1], 44.9)]); r += 1; sn += 1
    widths(ws, {'A': 14.4, 'B': 44.9, 'C': 69.6, 'D': 15.9, 'E': 67.3, 'F': 83.1, 'G': 70.9, 'H': 13.6}); ws.freeze_panes = 'C8'; finish(ws)

def build_walkthrough(wb, spec, P, wt):
    name = wt['ref']; ws = wb.create_sheet(name); ref_box(ws, 'K', name); header(ws, spec, P)
    put(ws, 'A7', f"Subprocess title: {ph(wt.get('title'), 'Sub-process title')}", F(True, WHITE), F_BAND, TOP, B_THIN)
    for col in 'BCDEFGHIJK': ws[f'{col}7'].fill = F_BAND; ws[f'{col}7'].border = B_THIN
    ws.merge_cells('A7:K7')
    bg = 'Background:\n\n' + ph(wt.get('background'), 'Describe the process, the systems used, the people involved and any change during the period that matters for this sub-process.')
    put(ws, 'A8', bg, F(), align=LEFTW, border=B_THIN); ws.merge_cells('A8:K8'); fit_height(ws, 8, [(bg, 260)], min_h=40)
    hdrs = ['No.', 'Activity', 'Performed By', 'Frequency', 'Control Reference No.', 'Existing controls / mitigating measures', 'Control Owner', 'Control Classification', 'Control Type', 'Finding for design effectiveness', 'Walkthrough Document']
    for j, h in enumerate(hdrs): put(ws, f'{get_column_letter(j+1)}9', h, F(True), F_HDR, TOPW if j < 2 else CTRW, B_THIN)
    ws.row_dimensions[9].height = 22.5; r = 10; n = 1; cref = 1
    for s, sec in enumerate(wt.get('sections') or [{'title': None, 'activities': [{}]}], 1):
        put(ws, f'A{r}', ph(sec.get('title'), f'{s}.1 Sub-section title'), F(True), align=LEFTW, border=B_THIN)
        for col in 'BCDEFGHIJK': ws[f'{col}{r}'].border = B_THIN
        ws.merge_cells(f'A{r}:K{r}'); ws.row_dimensions[r].height = 13.5; r += 1
        for act in sec.get('activities') or [{}]:
            if act.get('refer_to'): vals = [act.get('no', n), act['refer_to']] + [''] * 9
            else:
                cr = act.get('control_ref')
                if cr is None and (act.get('control') or act.get('owner')): cr = f'{name}-{cref}'; cref += 1
                vals = [act.get('no', n), ph(act.get('activity'), 'Activity - who does what, when, in which system'), ph(act.get('performed_by'), 'Performed by'), ph(act.get('frequency'), 'Frequency'), cr or '',
                        ph(act.get('control'), 'Control - what prevents or detects the risk, and the evidence'), ph(act.get('owner'), 'Control owner'), ph(act.get('classification'), 'Preventive / Detective'), ph(act.get('type'), 'Manual / IT'), act.get('de_finding') or '', ph(act.get('document'), 'Document to sight')]
            aligns = [CTR, TOPW, CTRW, CTRW, CTR, LEFTW, CTRW, CTRW, CTRW, LEFTW, CTRW]
            for j, (v, al) in enumerate(zip(vals, aligns)): put(ws, f'{get_column_letter(j+1)}{r}', v, F(True, RED) if j == 4 else F(), align=al, border=B_THIN)
            fit_height(ws, r, [(vals[1], 42.6), (vals[5], 50.4), (vals[9], 59.1), (vals[2], 19), (vals[10], 16)]); r += 1; n += 1
    r += 1
    put(ws, f'A{r}', wt.get('questions_title') or 'Walkthrough questions for the team lead (to be used during the walkthrough session)', F(True, WHITE), F_BAND, TOP, B_THIN)
    for col in 'BCDEFGHIJK': ws[f'{col}{r}'].fill = F_BAND; ws[f'{col}{r}'].border = B_THIN
    ws.merge_cells(f'A{r}:K{r}'); r += 1
    for col, h in [('A', 'No.'), ('B', 'Question'), ('E', 'Why we are asking, including the prior-year context'), ('G', 'Documents to request during the session'), ('I', 'Proposed interviewee'), ('J', 'Response noted during the walkthrough'), ('K', 'Control Reference No.')]: put(ws, f'{col}{r}', h, F(True), F_HDR, CTRW, B_THIN)
    for col in 'ABCDEFGHIJK': ws[f'{col}{r}'].border = B_THIN; ws[f'{col}{r}'].fill = F_HDR
    ws.merge_cells(f'B{r}:D{r}'); ws.merge_cells(f'E{r}:F{r}'); ws.merge_cells(f'G{r}:H{r}'); ws.row_dimensions[r].height = 22.5; r += 1
    for i, q in enumerate(wt.get('questions') or [{}, {}, {}], 1):
        vals = {'A': q.get('no', i), 'B': ph(q.get('question'), 'Open question that makes the interviewee describe the step'), 'E': ph(q.get('why'), 'Why we ask - the risk, the prior-year point or the change we want to understand'), 'G': ph(q.get('documents'), 'Documents to request'), 'I': ph(q.get('interviewee'), 'Interviewee'), 'J': q.get('response') or '', 'K': q.get('control_ref') or ''}
        for col, v in vals.items(): put(ws, f'{col}{r}', v, F(True, RED) if col == 'K' else F(), align=CTR if col == 'A' else (CTRW if col in 'IK' else LEFTW), border=B_THIN)
        for col in 'ABCDEFGHIJK': ws[f'{col}{r}'].border = B_THIN
        ws.merge_cells(f'B{r}:D{r}'); ws.merge_cells(f'E{r}:F{r}'); ws.merge_cells(f'G{r}:H{r}')
        fit_height(ws, r, [(vals['B'], 80), (vals['E'], 66), (vals['G'], 32), (vals['J'], 59)]); r += 1
    widths(ws, {'A': 12.6, 'B': 42.6, 'C': 19, 'D': 19.1, 'E': 16, 'F': 50.4, 'G': 16.1, 'H': 16.6, 'I': 12.9, 'J': 59.1, 'K': 16.1}); ws.freeze_panes = 'A10'; finish(ws)

def test_layout(t):
    detail = t.get('detail_columns') or ['Document / item', 'Reference', 'Date', 'Amount (S$)']
    steps_txt = t.get('steps') if isinstance(t.get('steps'), str) else '\n'.join(f"{i}. {s}" for i, s in enumerate(t.get('steps') or [], 1))
    n_steps = int(t.get('n_steps') or max(1, len(re.findall(r'^\s*\d+\.', steps_txt or '', re.M)) or 6))
    n_rows = int(t.get('n_rows') or (len(t['rows']) if t.get('rows') else 10))
    first_step = 3 + len(detail); last = first_step + n_steps
    hdr_row = 22; conclusion_row = hdr_row + 1 + n_rows + 1 + 4
    return detail, steps_txt, n_steps, n_rows, first_step, last, hdr_row, conclusion_row

def build_test(wb, spec, P, t):
    name = t['ref']; ws = wb.create_sheet(name)
    detail, steps_txt, n_steps, n_rows, first_step, last, hdr_row, _ = test_layout(t)
    L = get_column_letter; last_l = L(last)
    ref_box(ws, last_l, name); ws.row_dimensions[1].height = 14.25; header(ws, spec, P)
    put(ws, 'A7', f"Test {t.get('no', '')}".strip(), F(True, WHITE), F_BAND, TOP, B_THIN); put(ws, 'B7', ph(t.get('title'), 'Test title'), F(True, WHITE), F_BAND, TOP)
    for j in range(2, last + 1): ws[f'{L(j)}7'].fill = F_BAND
    merge_to = L(min(last - 1, 9)) if last - 1 >= 3 else 'C'
    fields = [('Objective: ', ph(t.get('objective'), 'To check that ...')), ('Basis of Selection:', ph(t.get('basis'), 'Haphazard / judgemental, with emphasis on ... / Not applicable')), ('Population:', ph(t.get('population'), 'Population and its source listing')),
              ('Sample Size:', ph(t.get('sample_size'), 'n, or all items if the population is smaller')), ('Source:', ph(t.get('source'), 'Documents and listings used')), ('Audit Steps:', ph(steps_txt, f"Test {t.get('no', 'n')}: title\nCheck that:\n1. ...\n2. ..."))]
    r = 9
    for lab, val in fields:
        put(ws, f'A{r}', lab, F(True), align=TOPW); put(ws, f'B{r}', val, F(), align=LEFTW); ws.merge_cells(f'B{r}:{merge_to}{r}'); fit_height(ws, r, [(val, 150)], min_h=14.25); r += 2
    put(ws, f'A{r}', 'Work done:', F(True), align=TOP); put(ws, f'B{r}', ph(t.get('detail_title'), 'Sample details'), F(True), F_DETAIL, CTR, B_THIN)
    for j in range(3, first_step): ws[f'{L(j)}{r}'].fill = F_DETAIL; ws[f'{L(j)}{r}'].border = B_THIN
    ws.merge_cells(f'B{r}:{L(first_step-1)}{r}')
    put(ws, f'{L(first_step)}{r}', 'Work Done', F(True), F_STEPS, CTRW, B_THIN)
    for j in range(first_step + 1, last): ws[f'{L(j)}{r}'].fill = F_STEPS; ws[f'{L(j)}{r}'].border = B_THIN
    if n_steps > 1: ws.merge_cells(f'{L(first_step)}{r}:{L(last-1)}{r}')
    put(ws, f'{last_l}{r}', 'Remarks', F(True), F_DETAIL, CTR, B_THIN); ws.row_dimensions[r].height = 15; r += 1
    put(ws, f'B{r}', 'S/N', F(True), F_DETAIL, CTRW, B_THIN)
    for j, h in enumerate(detail): put(ws, f'{L(3+j)}{r}', h, F(True), F_DETAIL, CTRW, B_THIN)
    for s in range(n_steps): put(ws, f'{L(first_step+s)}{r}', s + 1, F(True), F_STEPS, CTRW, B_THIN)
    put(ws, f'{last_l}{r}', 'Exceptions and remarks', F(True), F_DETAIL, CTRW, B_THIN); ws.row_dimensions[r].height = 45; r += 1
    rows = t.get('rows') or [{} for _ in range(n_rows)]; formulas = t.get('detail_formulas') or {}
    for i, row in enumerate(rows, 1):
        put(ws, f'B{r}', i, F(), align=CTR, border=B_THIN); vals = row.get('values') or []
        for j in range(len(detail)):
            v = vals[j] if j < len(vals) else None
            if v is None and detail[j] in formulas: v = formulas[detail[j]].replace('{r}', str(r))
            put(ws, f'{L(3+j)}{r}', v, F(), align=TOPW, border=B_THIN)
        steps = row.get('steps') or []
        for s in range(n_steps): put(ws, f'{L(first_step+s)}{r}', steps[s] if s < len(steps) else None, F(), align=CTRW, border=B_THIN)
        put(ws, f'{last_l}{r}', row.get('remarks'), F(), align=LEFTW, border=B_THIN); ws.row_dimensions[r].height = 33.75; r += 1
    r += 1
    for lab, key, default in [('Notes:', 'notes', 'Nil'), ('Findings:', 'findings', 'Nil'), ('Conclusion:', 'conclusion', 'To be concluded on completion of fieldwork.')]:
        put(ws, f'A{r}', lab, F(True), align=LEFTW); put(ws, f'B{r}', t.get(key) or default, F(), align=TOP); r += 2
    w = {'A': 20.6, 'B': 7.6, last_l: 47.4}
    for j, h in enumerate(detail): w[L(3+j)] = max(14, min(42, int(len(h) * 1.4) + 8))
    for s in range(n_steps): w[L(first_step+s)] = 8.9
    widths(ws, w); ws.freeze_panes = f'C{hdr_row+1}'; finish(ws)

def build_da(wb, spec, P):
    name = spec.get('da', {}).get('ref') or f'{P}400'
    ws = wb.create_sheet(name); ref_box(ws, 'G', name); header(ws, spec, P)
    put(ws, 'B7', 'Data analytics', F(True), align=TOPW); put(ws, 'G7', 'Ref ', F(True), align=CTRW)
    for j, h in enumerate(['S/N', 'Analytic', 'Data source', 'What the analytic is designed to identify', 'Feeds', 'Result', 'Status']): put(ws, f'{get_column_letter(j+1)}9', h, F(True), F_HDR, CTRW, B_THIN)
    ws.row_dimensions[9].height = 22.5; r = 10
    for i, a in enumerate(spec.get('da', {}).get('analytics') or [{}, {}, {}], 1):
        vals = [a.get('sn', i), ph(a.get('analytic'), 'Analytic - what is compared with what'), ph(a.get('source'), 'Listings required'), ph(a.get('identifies'), 'What an exception looks like'), ph(a.get('feeds'), f'{P}3xx'), a.get('result') or '', a.get('status') or 'Planned']
        for j, (v, al) in enumerate(zip(vals, [CTR, LEFTW, LEFTW, LEFTW, CTRW, LEFTW, CTRW])): put(ws, f'{get_column_letter(j+1)}{r}', v, F(j == 4), align=al, border=B_THIN)
        fit_height(ws, r, [(vals[1], 64), (vals[3], 60), (vals[2], 18.9)]); r += 1
    r += 1
    note = spec.get('da', {}).get('note') or 'The analytics are to be run on the population listings requested in advance of fieldwork, so that the results can be used to select the samples for the tests above before the start of fieldwork.'
    put(ws, f'A{r}', note, F(), align=LEFTW); ws.merge_cells(f'A{r}:G{r}'); fit_height(ws, r, [(note, 200)], min_h=24)
    widths(ws, {'A': 13.6, 'B': 64, 'C': 18.9, 'D': 60, 'E': 9.7, 'F': 30, 'G': 11}); finish(ws); return name

def build_stamp_sheet(wb, spec, stamp_text):
    ws = wb.create_sheet('_build'); ws['A1'] = 'Build stamp'; ws['A1'].font = F(True)
    rows = [('skill', 'skill-awp'), ('skill_version', VERSION), ('built_on', dt.date.today().isoformat()), ('library_versions', json.dumps(spec.get('library_versions') or {})), ('modules', ', '.join(spec.get('module_ids') or [])), ('tokens', json.dumps(spec.get('tokens') or {}, ensure_ascii=False)), ('stamp', stamp_text)]
    for i, (k, v) in enumerate(rows, 2): ws[f'A{i}'] = k; ws[f'B{i}'] = v; ws[f'A{i}'].font = F(); ws[f'B{i}'].font = F()
    ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 90; ws.sheet_state = 'hidden'

# =============================================================================== build
def normalise(spec):
    """Fills refs and defaults so that an incomplete, hand-written spec still builds. Drops null entries."""
    P = spec.get('prefix') or 'A'
    spec['walkthroughs'] = [w for w in (spec.get('walkthroughs') or []) if isinstance(w, dict)]
    spec['tests'] = [t for t in (spec.get('tests') or []) if isinstance(t, dict)]
    spec['scope'] = [k for k in (spec.get('scope') or []) if isinstance(k, dict)]
    if not isinstance(spec.get('rcm'), dict): spec['rcm'] = {}
    spec['rcm']['sections'] = [s for s in (spec['rcm'].get('sections') or []) if isinstance(s, dict)]
    for s in spec['rcm']['sections']: s['rows'] = [r for r in (s.get('rows') or []) if isinstance(r, dict)]
    for w in spec['walkthroughs']:
        w['sections'] = [s for s in (w.get('sections') or []) if isinstance(s, dict)]
        for s in w['sections']: s['activities'] = [a for a in (s.get('activities') or []) if isinstance(a, dict)]
        w['questions'] = [q for q in (w.get('questions') or []) if isinstance(q, dict)]
    if not isinstance(spec.get('da'), dict): spec['da'] = {}
    spec['da']['analytics'] = [a for a in (spec['da'].get('analytics') or []) if isinstance(a, dict)]
    wts = spec['walkthroughs'] or [{'title': None}, {'title': None}]
    for i, wt in enumerate(wts, 1): wt.setdefault('ref', f'{P}{200+i}')
    tests = spec['tests'] or [{'title': None} for _ in range(3)]
    for i, t in enumerate(tests, 1):
        if not t.get('ref'): t['ref'] = f'{P}{300+i}'
        if not t.get('no'): t['no'] = i
    for i, wt in enumerate(wts, 1):
        if not wt.get('ref'): wt['ref'] = f'{P}{200+i}'
    spec['walkthroughs'] = wts; spec['tests'] = tests; spec['prefix'] = P
    if not spec['da'].get('ref'): spec['da']['ref'] = f'{P}400'
    return spec

def stamp_text(spec):
    lv = spec.get('library_versions') or {}
    parts = [f'skill-awp v{VERSION}'] + [f'{k} v{v}' for k, v in lv.items()] + [dt.date.today().isoformat()]
    return 'Built with ' + ' · '.join(parts)

def build_awp(spec, out_path):
    spec = normalise(json.loads(json.dumps(spec, default=str))); P = spec['prefix']
    wb = Workbook(); wb.remove(wb.active)
    st = stamp_text(spec) if spec.get('stamp', True) else None
    build_scope(wb, spec)
    da_ref = spec['da']['ref']
    build_index(wb, spec, P, P, 'Index', [('Risk and control matrix', f'{P}100'), ('Process walkthrough', f'{P}200'), ('Testing of controls', f'{P}300'), ('Data Analytics', da_ref)], 'G', {'A': 12, 'B': 64, 'C': 8.6, 'D': 12.1, 'E': 7.4, 'F': 8.4, 'G': 14.6}, stamp=st)
    concl = {t['ref']: test_layout(t)[7] for t in spec['tests']}
    build_rcm(wb, spec, P, concl)
    build_index(wb, spec, P, f'{P}200', 'Process walkthrough', [(wt.get('index_text') or ph(wt.get('title'), 'Sub-process'), wt['ref']) for wt in spec['walkthroughs']], 'G', {'A': 13.6, 'B': 64, 'C': 18.9, 'D': 7, 'E': 8.4, 'F': 7.4, 'G': 11})
    for wt in spec['walkthroughs']: build_walkthrough(wb, spec, P, wt)
    build_index(wb, spec, P, f'{P}300', 'Testing', [(ph(t.get('title'), 'Test title'), t['ref']) for t in spec['tests']], 'E', {'A': 13.6, 'B': 64, 'C': 18.9, 'D': 7, 'E': 9.7, 'F': 11})
    for t in spec['tests']: build_test(wb, spec, P, t)
    build_da(wb, spec, P)
    if st: build_stamp_sheet(wb, spec, st)
    wb.save(out_path); return out_path

def template_spec(n_walkthroughs=2, n_tests=3, n_analytics=3, prefix='A'):
    return {'prefix': prefix, 'client': '', 'process': '', 'period': '', 'signoff': True, 'stamp': True,
            'scope': [{'no': 1, 'key_process': '', 'areas': ['', ''], 'covered': 'Yes'}, {'no': 2, 'key_process': '', 'areas': ['', ''], 'covered': 'Yes'}],
            'rcm': {'sections': [{'title': '', 'rows': [{}, {}]}, {'title': '', 'rows': [{}]}]},
            'walkthroughs': [{'title': '', 'background': '', 'sections': [{'title': '', 'activities': [{}, {}]}], 'questions': [{}, {}, {}]} for _ in range(n_walkthroughs)],
            'tests': [{'title': '', 'n_steps': 6, 'n_rows': 10} for _ in range(n_tests)], 'da': {'analytics': [{} for _ in range(n_analytics)]}}

# =============================================================================== libraries and compose
def find_libraries(folder=None):
    """Core and packs in FOLDER (default: next to this script). With --lib, core falls back to the one next to the script, and packs are the union (the folder wins on a name clash)."""
    folder = folder or HERE
    core = next((p for p in (os.path.join(folder, 'skill-awp-lib-core.json'), os.path.join(HERE, 'skill-awp-lib-core.json')) if os.path.exists(p)), None)
    packs = {os.path.basename(p): p for p in glob.glob(os.path.join(HERE, 'skill-awp-lib-pack-*.json'))}
    packs.update({os.path.basename(p): p for p in glob.glob(os.path.join(folder, 'skill-awp-lib-pack-*.json'))})
    return core, [packs[k] for k in sorted(packs)]

def load_json(p):
    with open(p, encoding='utf-8') as f: return json.load(f)
def save_json(obj, p):
    with open(p, 'w', encoding='utf-8') as f: json.dump(obj, f, indent=1, ensure_ascii=False, default=str)

def subst(obj, tokens):
    """Replace {TOKEN} placeholders in every string of a structure. A token's value may itself contain tokens (up to three levels)."""
    if isinstance(obj, str):
        for _ in range(3):
            before = obj
            for k, v in (tokens or {}).items():
                if v: obj = obj.replace(k if k.startswith('{') else '{' + k + '}', str(v))
            if obj == before: break
        return obj
    if isinstance(obj, list): return [subst(x, tokens) for x in obj]
    if isinstance(obj, dict): return {subst(k, tokens): subst(v, tokens) for k, v in obj.items()}
    return obj

def tidy(obj):
    """Grammar after token substitution: doubled articles ("the the system"), lowercase sentence starts ("the system generates"), doubled spaces."""
    if isinstance(obj, str):
        s = re.sub(r'\b([Tt]he|[Aa]n?) (the|an?) ', lambda m: m.group(1) + ' ', obj)
        s = re.sub(r'(^|[.!?]\s+|\n)([a-z])', lambda m: m.group(1) + m.group(2).upper(), s)
        return re.sub(r'[ \t]{2,}', ' ', s)
    if isinstance(obj, list): return [tidy(x) for x in obj]
    if isinstance(obj, dict): return {k: (v if k in ('tokens', 'detail_formulas', 'module_ids', 'library_versions', 'id', 'ref', 'module', 'control_ref', 'wp_ref', 'control_no', 'feeds', 'sheet', 'prefix', 'client', 'process', 'period', 'prepared_by', 'reviewed_by') else tidy(v)) for k, v in obj.items()}
    return obj

STAGE_ORDER = {'control_environment': 0, 'cycle': 1, 'closing': 2}

def load_libraries(selection, folder=None):
    """core + the packs named in the selection, in that order."""
    core_p, _ = find_libraries(folder); libs = []
    if core_p: libs.append(load_json(core_p))
    _, packs = find_libraries(folder); by_name = {os.path.basename(p)[len('skill-awp-lib-pack-'):-5]: p for p in packs}
    for name in selection.get('packs') or []:
        if name not in by_name: raise SystemExit(f'pack not found: skill-awp-lib-pack-{name}.json (available: {", ".join(sorted(by_name)) or "none"})')
        libs.append(load_json(by_name[name]))
    if selection.get('extra_modules') or selection.get('extra_sheets'):     # engagement-specific modules written in library format inside the selection
        libs.append({'library': 'engagement', 'process': 'engagement', 'version': 0, 'sheets': selection.get('extra_sheets') or [], 'modules': selection.get('extra_modules') or [], 'token_defaults': {}})
    return libs

def resolve_tokens(selection, libs):
    """core defaults < pack defaults < derived from the selection < selection.tokens"""
    tokens = {}; origin = {}
    for lib in libs:
        for k, v in (lib.get('token_defaults') or {}).items(): tokens[k] = v; origin[k] = lib.get('library')
    client = selection.get('client') or ''; process = selection.get('process') or ''
    derived = {'{CLIENT}': client, '{CLIENT_SHORT}': client, '{PROCESS}': process.lower() if process else tokens.get('{PROCESS}', ''), '{DEPARTMENT}': process, '{PERIOD}': selection.get('period') or ''}
    for k, v in derived.items():
        if v: tokens[k] = v
    tokens.update({(k if k.startswith('{') else '{' + k + '}'): v for k, v in (selection.get('tokens') or {}).items() if v})
    dept = tokens.get('{DEPARTMENT}') or 'the department'
    tokens.setdefault('{HEAD_OF_DEPARTMENT}', f'Head of {dept}'); tokens.setdefault('{MANAGER}', f'Manager ({dept})'); tokens.setdefault('{STAFF}', f'{dept} staff')
    tokens['_origin'] = {k: origin.get(k, 'derived') for k in tokens if k != '_origin' and k not in {(x if x.startswith('{') else '{' + x + '}') for x in (selection.get('tokens') or {})}}
    return tokens

def compose(selection, folder=None):
    """selection = {client, process, period, prefix?, packs[], tokens{}, modules?[] (ids; default all), exclude?[], prior_year{risk_id: text}, background_additions{sheet_id: text}, extra{}}
    Returns an engagement spec with every reference resolved and every token substituted."""
    libs = load_libraries(selection, folder); tokens = resolve_tokens(selection, libs)
    if not libs: raise SystemExit(f'no libraries found in {folder or HERE}: keep skill-awp-lib-core.json and the packs next to skill-awp.py, or pass --lib FOLDER')
    notes = []
    P = selection.get('prefix') or 'A'
    wanted = set(selection.get('modules') or []); excl = set(selection.get('exclude') or [])
    sheets = {}; modules = []; sheet_map = {}
    for lib in libs:
        for sh in lib.get('sheets', []): sheets[sh['id']] = dict(sh)          # later libraries may override a sheet's text
        sheet_map.update(lib.get('sheet_map') or {})
        for m in lib.get('modules', []):
            if (wanted and m['id'] not in wanted) or m['id'] in excl: continue
            m = json.loads(json.dumps(m)); m['_lib'] = lib.get('library'); modules.append(m)
    sheet_map.update(selection.get('sheet_map') or {})
    ov = selection.get('overrides') or {}; hit = set()
    for m in modules:                                                        # overrides: {"test.x": {"basis": ".."}, "q.y": {"why": ".."}, "ctrl.z": {"text": "..", "activity": ".."}, "risk.r": {"risk": ".."}, "da.a": {..}, "mod.m": {"scope_areas": [..]}}
        if m['id'] in ov: m.update({k: v for k, v in ov[m['id']].items() if k != 'id'}); hit.add(m['id'])
        for coll in ('tests', 'risks', 'analytics'):
            for x in m.get(coll, []):
                if x.get('id') in ov: x.update({k: v for k, v in ov[x['id']].items() if k != 'id'}); hit.add(x['id'])
        for qq in m.get('walkthrough', {}).get('questions', []):
            if qq.get('id') in ov: qq.update({k: v for k, v in ov[qq['id']].items() if k != 'id'}); hit.add(qq['id'])
        for a in m.get('walkthrough', {}).get('activities', []):
            c = a.get('control')
            if isinstance(c, dict) and c.get('id') in ov:
                for k, v in ov[c['id']].items():
                    if k in ('activity', 'performed_by', 'frequency', 'document'): a[k] = v
                    elif k != 'id': c[k] = v
                hit.add(c['id'])
    for k in ov:
        if k not in hit: notes.append(f'overrides key {k} matches nothing in the selection (run: python skill-awp.py list)')
    all_ids = {m['id'] for lib in libs for m in lib.get('modules', [])}
    for mid in sorted((wanted | excl) - all_ids): notes.append(f'module {mid} is not in the loaded libraries (check the id or the packs list)')
    if not modules: raise SystemExit('no modules selected - check "modules" / "exclude" in the selection')
    for m in modules:
        if m['id'] in sheet_map: m['walkthrough']['sheet'] = sheet_map[m['id']]
    key = lambda m: (STAGE_ORDER.get(m.get('stage', 'cycle'), 1), m.get('order', 50), m['id'])
    modules.sort(key=key)
    sheet_order = []
    for m in modules:
        sid = m.get('walkthrough', {}).get('sheet')
        if sid and sid not in sheet_order: sheet_order.append(sid)
    modules.sort(key=lambda m: (sheet_order.index(m['walkthrough']['sheet']),) + key(m))   # everything follows the sheet sequence
    sheet_ref = {sid: f'{P}{201+i}' for i, sid in enumerate(sheet_order)}
    tests_all = [(t, m) for m in modules for t in m.get('tests', [])]
    test_ref = {t['id']: f'{P}{301+i}' for i, (t, _) in enumerate(tests_all)}
    test_no = {t['id']: i + 1 for i, (t, _) in enumerate(tests_all)}
    ctrl_ref = {}; counters = {}
    for m in modules:
        sid = m['walkthrough']['sheet']
        for a in m['walkthrough'].get('activities', []):
            c = a.get('control')
            if isinstance(c, dict) and c.get('id'):
                counters[sid] = counters.get(sid, 0) + 1; ctrl_ref[c['id']] = f'{sheet_ref[sid]}-{counters[sid]}'
    adds = selection.get('background_additions') or {}
    for k in adds:
        if k not in sheet_order: notes.append(f'background_additions key {k} is not a walkthrough sheet in this selection (sheets: {", ".join(sheet_order)})')
    all_risks = {rk['id'] for m in modules for rk in m.get('risks', [])}
    for k in (selection.get('prior_year') or {}):
        if k not in all_risks: notes.append(f'prior_year key {k} is not a risk in this selection (run: python skill-awp.py list)')
    # ---- walkthrough sheets
    wts = []
    for i, sid in enumerate(sheet_order, 1):
        sh = sheets.get(sid, {'title': sid}); bg = sh.get('background') or ''
        if adds.get(sid): bg = (bg + '\n\n' if bg else '') + adds[sid]
        wt = {'ref': sheet_ref[sid], 'id': sid, 'title': sh.get('title'), 'index_text': sh.get('index_text'), 'background': bg, 'questions_title': sh.get('questions_title'), 'sections': [], 'questions': []}
        k = 0; n = 1; qn = 1
        for m in modules:
            w = m['walkthrough']
            if w['sheet'] != sid: continue
            if not w.get('activities'):                                      # question-only module (e.g. general closing questions): no sub-section row
                for qq in w.get('questions', []):
                    wt['questions'].append({'no': qn, 'id': qq.get('id'), 'question': qq.get('question'), 'why': qq.get('why'), 'documents': qq.get('documents'), 'interviewee': qq.get('interviewee'), 'response': '', 'control_ref': ctrl_ref.get(qq.get('control'), '') if qq.get('control') else ''}); qn += 1
                continue
            k += 1; sec = {'title': f"{i}.{k} {w.get('section_title') or m.get('title')}", 'module': m['id'], 'activities': []}
            for a in w.get('activities', []):
                if a.get('refer_to'):
                    rt = a['refer_to']
                    for tid, ref in test_ref.items(): rt = rt.replace('<' + tid + '>', '<' + ref + '>')
                    sec['activities'].append({'no': n, 'refer_to': rt}); n += 1; continue
                c = a.get('control') or {}
                sec['activities'].append({'no': n, 'activity': a.get('activity'), 'performed_by': a.get('performed_by'), 'frequency': a.get('frequency'), 'control_ref': ctrl_ref.get(c.get('id'), '') if c else '',
                                          'control': c.get('text', '') if c else '', 'owner': c.get('owner', '') if c else '', 'classification': c.get('classification', '') if c else '', 'type': c.get('type', '') if c else '', 'de_finding': '', 'document': a.get('document') or (c.get('document', '') if c else '')})
                n += 1
            wt['sections'].append(sec)
            for qq in w.get('questions', []):
                wt['questions'].append({'no': qn, 'id': qq.get('id'), 'question': qq.get('question'), 'why': qq.get('why'), 'documents': qq.get('documents'), 'interviewee': qq.get('interviewee'), 'response': '', 'control_ref': ctrl_ref.get(qq.get('control'), '') if qq.get('control') else ''}); qn += 1
        wts.append(wt)
    # ---- tests
    tests = []
    for t, m in tests_all:
        steps = t.get('steps') or []
        steps_txt = f"Test {test_no[t['id']]}: {t.get('title')}\n" + (t.get('steps_intro') or 'Check that:') + '\n' + '\n'.join(f'{i}. {s}' for i, s in enumerate(steps, 1))
        tests.append({'ref': test_ref[t['id']], 'id': t['id'], 'no': test_no[t['id']], 'title': t.get('title'), 'objective': t.get('objective'), 'basis': t.get('basis'), 'population': t.get('population'), 'sample_size': t.get('sample_size'), 'source': t.get('source'),
                      'steps': steps_txt, 'detail_title': t.get('detail_title'), 'detail_columns': t.get('detail_columns'), 'detail_formulas': t.get('detail_formulas'), 'n_steps': len(steps) or t.get('n_steps'), 'n_rows': t.get('n_rows') or 10})
    by_id = {t['id']: t for t in tests}
    def steps_subset(tid, nums):
        lines = by_id[tid]['steps'].split('\n')
        return '\n'.join(l for l in lines[2:] if re.match(r'^\s*(\d+)\.', l) and int(re.match(r'^\s*(\d+)\.', l).group(1)) in nums)
    # ---- RCM: one section per walkthrough sheet, one row per risk; the first risk that uses a test carries its steps
    sections = []; sn = 1; given = set()
    for i, sid in enumerate(sheet_order, 1):
        sh = sheets.get(sid, {'title': sid}); sec = {'title': f"{i}. {sh.get('title')}", 'rows': []}
        for m in modules:
            if m['walkthrough']['sheet'] != sid: continue
            for rk in m.get('risks', []):
                cn = '\n'.join(ctrl_ref[c] for c in rk.get('controls', []) if c in ctrl_ref) or '-'
                tids = [t for t in rk.get('tests', []) if t in test_ref]
                parts = []; refs = []
                for tid in tids:
                    if tid not in given:
                        given.add(tid); full = by_id[tid]['steps']
                        if rk.get('strategy_steps'): full = '\n'.join(full.split('\n')[:2]) + '\n' + steps_subset(tid, rk['strategy_steps'])
                        parts.append(full)
                    elif rk.get('strategy_steps'): parts.append(f'Refer to Test {test_no[tid]}.\nCheck that:\n' + steps_subset(tid, rk['strategy_steps']))
                    else: refs.append(test_no[tid])
                if refs: parts.insert(0, 'Refer to Test' + ('s ' if len(refs) > 1 else ' ') + ', '.join(str(n) for n in refs[:-1]) + (' and ' if len(refs) > 1 else '') + str(refs[-1]) + '.')
                missing = [t for t in rk.get('tests', []) if t not in test_ref]
                if missing: notes.append(f"risk {rk['id']} refers to test(s) {', '.join(missing)} that are not in the selection - add the module or write a strategy by hand")
                de = 'To be concluded after the walkthrough.'
                py = (selection.get('prior_year') or {}).get(rk['id'])
                if py:
                    de += '\n\n' + py
                    obs = re.search(r'Observation\s+\d+(?:\.\d+)*', py, re.I); label = obs.group(0) if obs else 'observation on ' + subst(rk.get('sub_process') or '', tokens).lower()
                    for tid in tids:                                    # test objective
                        t = by_id[tid]; line = f' This test also follows up on prior-year {label}.'
                        if line.strip() not in (t.get('objective') or ''): t['objective'] = (t.get('objective') or '').rstrip() + line
                    done_ctrl = set()                                   # "why we ask": the first question on each of this risk's controls
                    for wt in wts:
                        for qq in wt['questions']:
                            cr = qq.get('control_ref')
                            if cr and cr in [ctrl_ref.get(c) for c in rk.get('controls', [])] and cr not in done_ctrl and py not in (qq.get('why') or ''):
                                qq['why'] = (qq.get('why') or '').rstrip() + '\n\n' + py; done_ctrl.add(cr)
                sec['rows'].append({'sn': sn, 'id': rk['id'], 'sub_process': rk.get('sub_process'), 'risk': rk.get('risk'), 'control_no': cn, 'de_finding': de, 'audit_strategy': '\n\n'.join(parts), 'testing_result': '', 'wp_ref': '\n'.join(test_ref[t] for t in tids)}); sn += 1
        sections.append(sec)
    # ---- scope: one key process per walkthrough sheet
    scope = []
    for i, sid in enumerate(sheet_order, 1):
        sh = sheets.get(sid, {'title': sid}); areas = [a for m in modules if m['walkthrough']['sheet'] == sid for a in m.get('scope_areas', [])]
        scope.append({'no': i, 'key_process': sh.get('title'), 'areas': areas or [None], 'covered': 'Yes'})
    # ---- DA
    analytics = []
    for m in modules:
        for a in m.get('analytics', []):
            analytics.append({'sn': len(analytics) + 1, 'id': a['id'], 'analytic': a.get('analytic'), 'source': a.get('source'), 'identifies': a.get('identifies'), 'feeds': '\n'.join(test_ref[t] for t in a.get('tests', []) if t in test_ref), 'result': '', 'status': 'Planned'})
    lv = {('core' if lib.get('library') == 'core' else f"pack {lib.get('process')}"): (f"{lib.get('version')} (draft)" if not lib.get('version') else lib.get('version')) for lib in libs if lib.get('library') != 'engagement'}
    spec = {'prefix': P, 'client': selection.get('client', ''), 'process': selection.get('process', ''), 'period': selection.get('period', ''), 'prepared_by': selection.get('prepared_by', ''), 'reviewed_by': selection.get('reviewed_by', ''), 'signoff': True, 'stamp': True, 'link_results': True,
            'library_versions': lv, 'module_ids': [m['id'] for m in modules], 'tokens': tokens,
            'scope': scope, 'rcm': {'sections': sections}, 'walkthroughs': wts, 'tests': tests, 'da': {'ref': f'{P}400', 'analytics': analytics, 'note': None}}
    keep = spec['tokens']; spec['tokens'] = {}
    raw = json.dumps(spec, ensure_ascii=False)
    given = {(k if k.startswith('{') else '{' + k + '}') for k in (selection.get('tokens') or {})} | {'{CLIENT}', '{CLIENT_SHORT}', '{PROCESS}', '{DEPARTMENT}', '{PERIOD}'}
    origin = keep.pop('_origin', {})                                       # pack defaults are the process's own words; core defaults and derived roles are generic and get confirmed
    defaulted = sorted(k for k in keep if k in raw and k not in given and origin.get(k) in ('core', 'derived'))
    if defaulted: notes.append('info: tokens left at their generic default - confirm them against the material or set them in the selection: ' + ', '.join(f'{k}={keep[k]}' for k in defaulted))
    spec = subst(spec, keep); spec['tokens'] = keep
    spec = tidy(spec)
    cap = lambda s: (s[0].upper() + s[1:]) if isinstance(s, str) and s else s
    for wt in spec['walkthroughs']:
        wt['title'] = cap(wt['title']); wt['index_text'] = cap(wt['index_text'])
        for sec in wt['sections']:
            head, _, rest = sec['title'].partition(' '); sec['title'] = head + ' ' + cap(rest)
    for t in spec['tests']:
        t['title'] = cap(t['title']); t['detail_title'] = cap(t['detail_title']); t['steps'] = re.sub(r'^(Test \d+: )(.)', lambda m: m.group(1) + m.group(2).upper(), t['steps'])
    for sec in spec['rcm']['sections']:
        for r in sec['rows']:
            r['sub_process'] = cap(r['sub_process']); r['audit_strategy'] = re.sub(r'(Test \d+: )(.)', lambda m: m.group(1) + m.group(2).upper(), r['audit_strategy'])
    for k in spec['scope']: k['key_process'] = cap(k['key_process'])
    for wt in spec['walkthroughs']:
        for sec in wt['sections']:
            for a in sec['activities']:
                for f in ('frequency', 'performed_by', 'owner', 'classification', 'type'):
                    if a.get(f): a[f] = cap(a[f])
    for k, v in (selection.get('extra') or {}).items(): spec[k] = v          # engagement-specific overrides, taken verbatim
    for sec in spec['walkthroughs']:
        for s in sec['sections']:
            for a in s['activities']:
                if a.get('refer_to') and re.search(r'<test\.[a-z_0-9]+>', a['refer_to']): notes.append(f"{sec['ref']}: '{a['refer_to'][:60]}' points at a test that is not in the selection")
    spec['compose_notes'] = notes
    return spec

# =============================================================================== validate
RISK_WORDS = ('may result', 'resulting in', 'so that', 'leading to', 'which would', 'this may', 'could result', 'exposing', 'leaving', 'without this being detected', 'not detected', 'misstated', 'unsecured', 'loss', 'disputes', 'unauthorised', 'incorrect', 'late', 'incomplete')
def validate(obj, kind=None):
    """Returns list of (severity, message). kind: 'spec' | 'library' (auto-detected)."""
    issues = []
    kind = kind or ('library' if 'modules' in obj else 'spec')
    if kind == 'spec': obj = normalise(json.loads(json.dumps(obj, default=str)))
    txt = json.dumps({k: v for k, v in obj.items() if k not in ('tokens', 'build_stamp', 'token_defaults', 'denylist')}, ensure_ascii=False, default=str)
    allowed = set(TOKENS) | set((obj.get('token_defaults') or {}).keys())
    if kind == 'library' and obj.get('library') != 'core':
        core_p0, _ = find_libraries()
        if core_p0: allowed |= set((load_json(core_p0).get('token_defaults') or {}).keys())      # core's tokens are standard for every pack
    for tok in sorted(set(re.findall(r'\{[A-Z_0-9]+\}', txt.replace('{r}', '')))):
        if kind == 'spec': issues.append(('warn', f'unresolved token {tok}'))
        elif tok not in allowed: issues.append(('warn', f'non-standard token {tok} (declare it in token_defaults)'))
    if kind == 'spec':
        for m in sorted(set(re.findall(r'\b(?:the the|a the|an the|the a|the an)\b', txt)))[:3]: issues.append(('warn', f'doubled article "{m}" - check the text around tokens'))
        tests = {t['ref']: t for t in obj.get('tests', [])}; wts = obj.get('walkthroughs', [])
        refs = [t['ref'] for t in obj.get('tests', [])] + [w['ref'] for w in wts]
        for r in sorted({x for x in refs if refs.count(x) > 1}): issues.append(('error', f'duplicate sheet ref {r}'))
        crefs = [a.get('control_ref') for w in wts for s in w.get('sections', []) for a in s.get('activities', []) if a.get('control_ref')]
        for r in sorted({x for x in crefs if crefs.count(x) > 1}): issues.append(('error', f'duplicate control ref {r}'))
        ctrls = {a.get('control_ref') for wt in wts for s in wt.get('sections', []) for a in s.get('activities', []) if a.get('control_ref')}
        for sec in obj.get('rcm', {}).get('sections', []):
            for r in sec.get('rows', []):
                rk = (r.get('risk') or '').lower()
                if not rk or rk.startswith('['): issues.append(('error', f"RCM {r.get('sn')}: risk missing"))
                elif not any(w in rk for w in RISK_WORDS): issues.append(('warn', f"RCM {r.get('sn')} ({r.get('sub_process')}): risk does not state a consequence"))
                for ref in str(r.get('wp_ref') or '').split('\n'):
                    if ref and ref not in tests and not ref.startswith('['): issues.append(('error', f"RCM {r.get('sn')}: WP ref {ref} has no test sheet"))
                for c in str(r.get('control_no') or '').split('\n'):
                    if c and c != '-' and c not in ctrls and not c.startswith('['): issues.append(('error', f"RCM {r.get('sn')}: control {c} not on any walkthrough sheet"))
                if not r.get('audit_strategy') or str(r.get('audit_strategy')).startswith('['): issues.append(('warn', f"RCM {r.get('sn')}: audit strategy missing"))
        for wt in wts:
            for s in wt.get('sections', []):
                for a in s.get('activities', []):
                    if a.get('refer_to') and re.search(r'<test\.[a-z_0-9]+>', str(a['refer_to'])): issues.append(('error', f"{wt['ref']}: cross-reference '{str(a['refer_to'])[:50]}' not resolved to a sheet"))
            acts = [a for s in wt.get('sections', []) for a in s.get('activities', []) if not a.get('refer_to')]
            nctrl = sum(1 for a in acts if a.get('control') and not str(a['control']).startswith('['))
            if not wt.get('background') or str(wt.get('background')).startswith('['): issues.append(('warn', f"{wt['ref']}: background missing"))
            if nctrl < 2: issues.append(('warn', f"{wt['ref']}: only {nctrl} controls documented (expect >= 2)"))
            nq = sum(1 for q in wt.get('questions', []) if q.get('question') and not str(q['question']).startswith('['))
            if nq < 5: issues.append(('warn', f"{wt['ref']}: only {nq} walkthrough questions (expect >= 5)"))
            for q in wt.get('questions', []):
                if q.get('control_ref') and q['control_ref'] not in ctrls: issues.append(('error', f"{wt['ref']} Q{q.get('no')}: control ref {q['control_ref']} unknown"))
        for sec in obj.get('rcm', {}).get('sections', []):
            for r in sec.get('rows', []):
                m = re.match(r'\s*Test (\d+):', str(r.get('audit_strategy') or ''))
                if m:
                    t = next((x for x in obj.get('tests', []) if str(x.get('no')) == m.group(1)), None)
                    if t:
                        quoted = re.findall(r'^\s*(\d+)\.\s*(.+)$', str(r['audit_strategy']), re.M); actual = dict(re.findall(r'^\s*(\d+)\.\s*(.+)$', str(t.get('steps') or ''), re.M))
                        bad = [n for n, txt in quoted if actual.get(n, '').strip() != txt.strip()]
                        if bad: issues.append(('warn', f"RCM {r.get('sn')}: audit strategy quotes Test {m.group(1)} steps {', '.join(bad)} differently from the test sheet - edit the test's steps and let the RCM follow, or align both"))
        for ref, t in tests.items():
            n = len(re.findall(r'^\s*\d+\.', str(t.get('steps') or ''), re.M))
            if n < 3: issues.append(('warn', f'{ref}: only {n} numbered audit steps (expect >= 3)'))
            if t.get('n_steps') and n and int(t['n_steps']) != n: issues.append(('warn', f"{ref}: {t['n_steps']} step columns but {n} numbered steps"))
            for k in ('objective', 'basis', 'population', 'sample_size', 'source'):
                if not t.get(k) or str(t[k]).startswith('['): issues.append(('warn', f'{ref}: {k} missing'))
        for a in obj.get('da', {}).get('analytics', []):
            for ref in str(a.get('feeds') or '').split('\n'):
                if ref and ref not in tests and not ref.startswith('['): issues.append(('error', f"DA {a.get('sn')}: feeds {ref} has no test sheet"))
    else:
        ids = {}
        known_tests = {x['id'] for mm in obj.get('modules', []) for x in mm.get('tests', [])}; core_ctrls = set()
        if obj.get('library') in ('pack', 'engagement'):      # packs may point at core tests and controls (e.g. the reconciliation test, the DOA control)
            core_p, _ = find_libraries()
            if core_p:
                core = load_json(core_p); known_tests |= {x['id'] for mm in core.get('modules', []) for x in mm.get('tests', [])}
                core_ctrls = {a['control']['id'] for mm in core.get('modules', []) for a in mm.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict)}
        for m in obj.get('modules', []):
            for coll, key in (('tests', 'id'), ('risks', 'id'), ('analytics', 'id')):
                for x in m.get(coll, []): ids.setdefault(x['id'], 0); ids[x['id']] += 1
            for a in m.get('walkthrough', {}).get('activities', []):
                c = a.get('control')
                if isinstance(c, dict) and c.get('id'): ids.setdefault(c['id'], 0); ids[c['id']] += 1
            for rk in m.get('risks', []):
                for t in rk.get('tests', []):
                    if t not in known_tests: issues.append(('error', f"{m['id']}: risk {rk['id']} refers to unknown test {t}"))
            for a in m.get('analytics', []):
                for t in a.get('tests', []):
                    if t not in known_tests: issues.append(('error', f"{m['id']}: analytic {a['id']} refers to unknown test {t}"))
                if not any(w in (rk.get('risk') or '').lower() for w in RISK_WORDS): issues.append(('warn', f"{m['id']}: risk {rk['id']} does not state a consequence"))
        for k, v in ids.items():
            if v > 1: issues.append(('error', f'duplicate id {k}'))
        ctrl_ids = {a['control']['id'] for m in obj.get('modules', []) for a in m.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict) and a['control'].get('id')} | core_ctrls
        for m in obj.get('modules', []):
            for qq in m.get('walkthrough', {}).get('questions', []):
                if qq.get('control') and qq['control'] not in ctrl_ids: issues.append(('error', f"{m['id']}: question {qq.get('id')} refers to unknown control {qq['control']}"))
            for rk in m.get('risks', []):
                for c in rk.get('controls', []):
                    if c not in ctrl_ids: issues.append(('error', f"{m['id']}: risk {rk['id']} refers to unknown control {c}"))
            for c in [a['control']['id'] for a in m.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict) and a['control'].get('id')]:
                if not any(c in rk.get('controls', []) for mm in obj.get('modules', []) for rk in mm.get('risks', [])): issues.append(('warn', f"{m['id']}: control {c} is not linked to any risk (it will not appear in the RCM)"))
        # near-duplicate controls / questions (library bloat guard)
        texts = [(m['id'], a['control']['text']) for m in obj.get('modules', []) for a in m.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict) and a['control'].get('text')]
        for i in range(len(texts)):
            for j in range(i + 1, len(texts)):
                if difflib.SequenceMatcher(None, texts[i][1], texts[j][1]).ratio() > 0.9: issues.append(('warn', f'near-duplicate controls in {texts[i][0]} and {texts[j][0]}'))
        for w in obj.get('denylist', []):
            if w and w.lower() in txt.lower(): issues.append(('error', f'denylisted term present: {w}'))
    return issues

# =============================================================================== export (header-driven, tolerant)
def _find(ws, text, col=None, start=1, exact=False):
    for row in ws.iter_rows(min_row=start):
        for c in row:
            if col and get_column_letter(c.column) != col: continue
            v = c.value
            if isinstance(v, str) and ((v.strip().lower() == text.lower()) if exact else v.strip().lower().startswith(text.lower())): return c.row, c.column
    return None, None

def _hdr_map(ws, row):
    return {str(c.value).strip().lower(): c.column for c in ws[row] if c.value is not None and not isinstance(c.value, (int, float))}

def export_spec(path, report=None):
    report = report if report is not None else []
    wbf = load_workbook(path); wb = load_workbook(path, data_only=True)       # values for content, formulas only for the detail-column formulas
    idx = [s for s in wb.sheetnames if re.fullmatch(r'[A-Z]', s)]; P = idx[0] if idx else 'A'; A = wb[P] if idx else None
    spec = {'prefix': P, 'client': A['B3'].value if A else None, 'process': A['B4'].value if A else None, 'period': A['B5'].value if A else None, 'scope': [], 'rcm': {'sections': []}, 'walkthroughs': [], 'tests': [], 'da': {'analytics': []}}
    if '_build' in wb.sheetnames:
        bs = wb['_build']; spec['build_stamp'] = {bs[f'A{i}'].value: bs[f'B{i}'].value for i in range(2, bs.max_row + 1) if bs[f'A{i}'].value}
        try: spec['tokens'] = json.loads(spec['build_stamp'].get('tokens') or '{}')
        except Exception: spec['tokens'] = {}
    if 'Scope' in wb.sheetnames:
        ws = wb['Scope']; hr, _ = _find(ws, 'Indicative area', exact=False); cur = None
        if hr:
            h = _hdr_map(ws, hr); cK, cA, cC = h.get('key process', 2), [v for k, v in h.items() if k.startswith('indicative')][0], h.get('covered?', 4)
            for r in range(hr + 1, ws.max_row + 1):
                if ws.cell(r, cA).value is None: continue
                if ws.cell(r, cK).value: cur = {'no': ws.cell(r, 1).value, 'key_process': ws.cell(r, cK).value, 'areas': [], 'covered': ws.cell(r, cC).value}; spec['scope'].append(cur)
                if cur: cur['areas'].append(ws.cell(r, cA).value)
    rcm = f'{P}100'
    if rcm in wb.sheetnames:
        ws = wb[rcm]; hr, _ = _find(ws, 'Risk', exact=True)
        if hr:
            h = _hdr_map(ws, hr); col = lambda *names: next((h[n] for n in names if n in h), None)
            cB, cC, cD, cE, cF, cG, cH = col('sub-process'), col('risk'), col('control no.', 'control no'), col('finding for design effectiveness'), col('audit strategy'), col('testing result'), col('wp ref')
            sec = None
            for r in range(hr + 1, ws.max_row + 1):
                a = ws.cell(r, 1).value
                if a is None: continue
                if ws.cell(r, cC).value is None and ws.cell(r, cB).value is None: sec = {'title': a, 'rows': []}; spec['rcm']['sections'].append(sec); continue
                if sec is None: sec = {'title': '', 'rows': []}; spec['rcm']['sections'].append(sec)
                g = lambda c: ws.cell(r, c).value if c else None
                sec['rows'].append({'sn': a, 'sub_process': g(cB), 'risk': g(cC), 'control_no': g(cD), 'de_finding': g(cE), 'audit_strategy': g(cF), 'testing_result': g(cG), 'wp_ref': g(cH)})
        else: report.append(f'{rcm}: header row with "Risk" not found')
    for name in wb.sheetnames:
        if re.fullmatch(P + r'2\d\d', name) and name != f'{P}200':
            ws = wb[name]; wt = {'ref': name, 'sections': [], 'questions': []}
            tr, _ = _find(ws, 'Subprocess title', col='A'); wt['title'] = str(ws[f'A{tr}'].value).split(':', 1)[1].strip() if tr else None
            br, _ = _find(ws, 'Background', col='A'); wt['background'] = re.sub(r'^Background:\s*', '', str(ws[f'A{br}'].value)).strip() if br else None
            idxws = wb.get(f'{P}200') if hasattr(wb, 'get') else (wb[f'{P}200'] if f'{P}200' in wb.sheetnames else None)
            if idxws:
                for r in range(1, idxws.max_row + 1):
                    if idxws[f'G{r}'].value == name: wt['index_text'] = idxws[f'B{r}'].value
            hr, _ = _find(ws, 'Activity', col='B', exact=True)
            if not hr: hr, _ = _find(ws, 'Procedures', col='B', exact=True)
            qr, _ = _find(ws, 'Question', col='B', exact=True)
            if hr:
                h = _hdr_map(ws, hr); col = lambda *names: next((h[n] for n in names if n in h), None)
                cols = dict(performed_by=col('performed by'), frequency=col('frequency'), control_ref=col('control reference no.', 'control ref. no.', 'control reference no'), control=col('existing controls / mitigating measures', 'controls identified'), owner=col('control owner'), classification=col('control classification'), type=col('control type'), de_finding=col('finding for design effectiveness', 'weakness in controls design'), document=col('walkthrough document'))
                sec = None; end = (qr - 2) if qr else ws.max_row
                for r in range(hr + 1, end + 1):
                    a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
                    if a is not None and b is None and not isinstance(a, (int, float)): sec = {'title': a, 'activities': []}; wt['sections'].append(sec)
                    elif a is not None and b is not None:
                        if sec is None: sec = {'title': '', 'activities': []}; wt['sections'].append(sec)
                        act = {'no': a, 'activity': b}; act.update({k: (ws.cell(r, c).value if c else None) for k, c in cols.items()})
                        if all(act.get(k) in (None, '') for k in ('performed_by', 'control', 'owner')): act = {'no': a, 'refer_to': b}
                        sec['activities'].append(act)
            else: report.append(f'{name}: activity table not found')
            if qr:
                tr2, _ = _find(ws, 'Walkthrough questions', col='A'); wt['questions_title'] = ws[f'A{tr2}'].value if tr2 else None
                h = _hdr_map(ws, qr); col = lambda pref: next((v for k, v in h.items() if k.startswith(pref)), None)
                cw, cd, ci, crsp, cc = col('why'), col('documents'), col('proposed interviewee'), col('response'), col('control reference')
                for r in range(qr + 1, ws.max_row + 1):
                    a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
                    if a is not None and b is not None: wt['questions'].append({'no': a, 'question': b, 'why': ws.cell(r, cw).value if cw else None, 'documents': ws.cell(r, cd).value if cd else None, 'interviewee': ws.cell(r, ci).value if ci else None, 'response': ws.cell(r, crsp).value if crsp else None, 'control_ref': ws.cell(r, cc).value if cc else None})
            spec['walkthroughs'].append(wt)
    for name in wb.sheetnames:
        if re.fullmatch(P + r'3\d\d', name) and name != f'{P}300':
            ws = wb[name]; t = {'ref': name}
            tr, _ = _find(ws, 'Test ', col='A'); t['no'] = int(re.sub(r'\D', '', str(ws[f'A{tr}'].value)) or 0) if tr else None; t['title'] = ws[f'B{tr}'].value if tr else None
            lab = {'objective': 'objective', 'basis of selection': 'basis', 'population': 'population', 'sample size': 'sample_size', 'source': 'source', 'audit steps': 'steps', 'notes': 'notes', 'findings': 'findings', 'conclusion': 'conclusion'}
            wd = None
            for r in range(1, ws.max_row + 1):
                a = ws[f'A{r}'].value
                if not isinstance(a, str): continue
                key = a.strip().rstrip(':').strip().lower()
                if key in lab: t[lab[key]] = ws[f'B{r}'].value
                if key == 'work done': wd = r
            if wd:
                t['detail_title'] = ws[f'B{wd}'].value; hdr = wd + 1; cols = []; steps = 0
                for j in range(3, ws.max_column + 1):
                    v = ws.cell(row=hdr, column=j).value
                    if v is None: continue
                    if isinstance(v, (int, float)) or (isinstance(v, str) and v.isdigit()): steps += 1
                    elif str(v).startswith('Exceptions'): pass
                    else: cols.append(v)
                t['detail_columns'] = cols; t['n_steps'] = steps
                n = 0; r = hdr + 1; rows = []
                while ws[f'B{r}'].value is not None:
                    vals = [ws.cell(r, 3 + j).value for j in range(len(cols))]; stp = [ws.cell(r, 3 + len(cols) + s).value for s in range(steps)]; rem = ws.cell(r, 3 + len(cols) + steps).value
                    if any(v not in (None, '') for v in vals + stp + [rem]): rows.append({'values': vals, 'steps': stp, 'remarks': rem})
                    n += 1; r += 1
                t['n_rows'] = n
                if rows: t['rows'] = rows
                formulas = {}; wsf = wbf[name]
                for j, hname in enumerate(cols):
                    v = wsf.cell(row=hdr + 1, column=3 + j).value
                    if isinstance(v, str) and v.startswith('='): formulas[hname] = re.sub(r'([A-Z]{1,2})' + str(hdr + 1) + r'(?!\d)', r'\1{r}', v)
                if formulas:
                    t['detail_formulas'] = formulas
                    for row in rows:
                        row['values'] = [None if (isinstance(v, str) and v.startswith('=')) else v for v in row['values']]
            else: report.append(f'{name}: work-done table not found')
            spec['tests'].append(t)
    da_name = next((s for s in wb.sheetnames if s in (f'{P}400', 'DA')), None)
    if da_name:
        ws = wb[da_name]; spec['da']['ref'] = da_name; hr, _ = _find(ws, 'Analytic', col='B', exact=True)
        if hr:
            for r in range(hr + 1, ws.max_row + 1):
                if ws[f'B{r}'].value and ws[f'A{r}'].value is not None: spec['da']['analytics'].append({'sn': ws[f'A{r}'].value, 'analytic': ws[f'B{r}'].value, 'source': ws[f'C{r}'].value, 'identifies': ws[f'D{r}'].value, 'feeds': ws[f'E{r}'].value, 'result': ws[f'F{r}'].value, 'status': ws[f'G{r}'].value})
                elif ws[f'A{r}'].value and ws[f'B{r}'].value is None and len(str(ws[f'A{r}'].value)) > 40: spec['da']['note'] = ws[f'A{r}'].value
    return spec

# =============================================================================== learn (harvest -> proposals) and apply
STOP = {'that', 'with', 'from', 'this', 'which', 'were', 'have', 'been', 'their', 'before', 'after', 'during', 'where', 'there', 'other', 'such', 'what', 'when', 'into', 'will', 'also', 'each', 'than', 'then', 'them', 'they', 'these', 'those', 'being', 'does', 'please', 'refer'}
def _stem(w):
    for suf in ('ations', 'ation', 'ings', 'ing', 'ies', 'ied', 'ed', 'es', 'al', 's'):
        if w.endswith(suf) and len(w) - len(suf) >= 4: return w[:-len(suf)]
    return w
def _words(s): return {_stem(w) for w in re.findall(r'[a-z]{4,}', (s or '').lower()) if w not in STOP}
def _sim(a, b):
    """Blend of character similarity and shared-vocabulary containment, so that a longer or more specific version of the same control still matches it."""
    a = a or ''; b = b or ''
    if not a.strip() or not b.strip(): return 0
    r = difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()
    wa, wb = _words(a), _words(b); cont = len(wa & wb) / min(len(wa), len(wb)) if wa and wb else 0
    first = difflib.SequenceMatcher(None, a.lower().split('.')[0], b.lower().split('.')[0]).ratio()
    return max(r, 0.6 * cont + 0.4 * first)

def build_denylist(spec, tokens):
    """Names that must not reach a library and are not covered by a token: the client name, abbreviations defined in the text ("XYZ"), company names in work-done rows."""
    words = set()
    for k in ('client',):
        if spec.get(k): words.add(str(spec[k]))
    words.update(re.findall(r'\("([A-Z]{2,8})"\)', json.dumps(spec, default=str)))
    for t in spec.get('tests', []):
        for row in t.get('rows', []) or []:
            for v in row.get('values') or []:
                if isinstance(v, str) and re.search(r'\b(PTE|LTD|LIMITED|INC|LLC|BHD|CORP|GMBH|PLC)\b', v.upper()): words.add(v.strip())
    covered = {str(v) for v in (tokens or {}).values() if v}
    return sorted(w for w in words if w and w not in covered)

def _pattern(value, ci=False):
    v = re.escape(str(value)); flags = re.I if (ci or str(value).islower()) else 0
    return re.compile(r'(?<![A-Za-z0-9])' + v + r'(s|es)?(?![A-Za-z0-9])', flags)

PERSON = re.compile(r'\b(?:Mr|Mrs|Ms|Dr|Mdm|Miss)\.?\s+[A-Z][a-z]+')

NUMBER_WORDS = {'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'twenty', 'thirty', 'sixty', 'ninety'}
def _substitutable(value):
    """Token values that are numbers or number words ('3', '90', 'three') are thresholds; replacing them blindly corrupts unrelated text, so they are left for the reviewer."""
    v = str(value).strip()
    return len(v) >= 2 and not re.fullmatch(r'[\d.,%]+', v) and v.lower() not in NUMBER_WORDS

def anonymise(text, tokens, denylist):
    if not isinstance(text, str): return text
    out = text
    pairs = sorted([(str(v), k if k.startswith('{') else '{' + k + '}') for k, v in (tokens or {}).items() if v and k not in ('{PERIOD}', 'PERIOD') and _substitutable(v)], key=lambda x: -len(x[0]))
    for v, k in pairs: out = _pattern(v).sub(lambda m, k=k: k + (m.group(1) or ''), out)
    for w in sorted(denylist, key=len, reverse=True):
        out = _pattern(w).sub('{CLIENT_SHORT}' if (len(w) <= 8 and w.isupper()) else '[REDACTED]', out)
    return out

def anonymise_obj(o, tk, dl):
    if isinstance(o, dict): return {k: anonymise_obj(v, tk, dl) for k, v in o.items()}
    if isinstance(o, list): return [anonymise_obj(v, tk, dl) for v in o]
    return anonymise(o, tk, dl)

def _empty(v): return v is None or str(v).strip().lower().rstrip('.') in ('', 'n.a', 'na', 'nil', 'none', '-', 'not applicable')

def known_words(libs):
    """Capitalised words already used in the libraries are generic by definition; anything else capitalised in a harvested text is flagged for review."""
    txt = json.dumps([lib.get('modules') for lib in libs] + [lib.get('sheets') for lib in libs], ensure_ascii=False, default=str)
    return set(re.findall(r'\b[A-Z][A-Za-z]+\b', txt)) | {'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'}

def engagement_words(tokens, libs, denylist=()):
    """Token values that are the engagement's own (not a library default) plus the denylist, and the client name without its corporate suffix: these must never reach a library."""
    defaults = {str(v) for lib in libs for v in (lib.get('token_defaults') or {}).values()} | {'CEO', 'IT', 'Finance'}
    words = {str(v) for v in (tokens or {}).values() if v and _substitutable(v) and str(v) not in defaults} | set(denylist)
    for k in ('{CLIENT}', '{CLIENT_SHORT}'):
        v = str((tokens or {}).get(k) or '')
        short = re.sub(r'\s*\b(Pte\.?|Ltd\.?|Limited|Inc\.?|LLC|Bhd\.?|Sdn\.?|Corp\.?|GmbH|PLC|Co\.?|Private|Company)\b\.?', '', v, flags=re.I).strip(' ,')
        if len(short) >= 4: words.add(short)
    return sorted(words, key=len, reverse=True)

def review_terms(text, known, words=()):
    """Words and figures in an anonymised text that may still identify the engagement: the engagement's own words, unknown capitalised words, amounts, years, FY labels, observation numbers, person titles, day/month counts."""
    if not isinstance(text, str): return []
    text = re.sub(r'\{[A-Z_0-9]+\}', ' ', text)                    # token names are not words to review
    found = set()
    for v in words:
        if _pattern(v).search(text): found.add(v)
    for m in re.finditer(r'\b[A-Z][A-Za-z]+\b', text):
        w = m.group(0); i = m.start()
        at_start = i == 0 or text[max(0, i - 2):i].strip() in ('.', '?', '!', ':', '"', '(') or text[i-1] in '\n"('
        if w not in known and not at_start: found.add(w)
    for pat in (r'[A-Z]{0,3}\$\s?[\d,]+(?:\.\d+)?', r'\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b', r'\b\d{5,}\b', r'\b(?:19|20)\d\d(?:/\d\d)?\b', r'\bFY\s?\d{2,4}(?:/\d{2,4})?\b', r'\bobservation\s*(?:no\.?|number|#)?\s*\d+(?:\.\d+)?\b', r'\((?:High|Moderate|Medium|Low|Significant|Major|Minor)\)', r'\b(?:Mr|Mrs|Ms|Dr|Mdm|Miss)\.?\s+[A-Z][a-z]+', r'\b\d+(?:\.\d+)?\s?(?:days?|months?|weeks?|years?|%|per cent)\b', r'\b(?:v|version)\s?\d+(?:\.\d+)?\b', r'#\d{2}-\d+'):
        found.update(re.findall(pat, text, flags=re.I))
    return sorted(found)

def _hash(o): 
    import hashlib; return hashlib.sha1(json.dumps(o, sort_keys=True, ensure_ascii=False, default=str).encode('utf-8')).hexdigest()[:12]

def _status(sim, harvested_text, library_text):
    """'same' only when the texts really are the same; a longer text that contains the library text is a refinement (an extension), not a duplicate; a weak match is flagged as uncertain."""
    r = difflib.SequenceMatcher(None, (harvested_text or '').lower(), (library_text or '').lower()).ratio()
    if r > 0.85: return 'same'
    if sim > 0.45: return 'refinement (extends the library text)' if sim >= 0.95 else ('refinement' if sim >= 0.6 else 'refinement (uncertain match - check the objective)')
    return 'new'

def learn(awp_path, out_path, tokens_in=None, folder=None):
    """Harvest an AWP and compare with the libraries. Writes a proposals file (and a readable .md next to it)."""
    report = []; spec = export_spec(awp_path, report)
    tokens = dict(spec.get('tokens') or {}); tokens.update({(k if k.startswith('{') else '{' + k + '}'): v for k, v in (tokens_in or {}).items()})
    core_p, pack_ps = find_libraries(folder); libs = [load_json(p) for p in ([core_p] if core_p else []) + pack_ps]
    learned = {x.get('source_hash') for lib in libs for m in lib.get('modules', []) for coll in (m.get('tests', []), m.get('risks', []), m.get('analytics', []), m.get('walkthrough', {}).get('questions', []), [a.get('control') for a in m.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict)]) for x in coll if x and x.get('source_hash')}
    known = known_words(libs); raw = {}                                   # vocabulary and raw (tokenised) text come from the libraries as shipped
    for lib in libs:
        for m in lib.get('modules', []):
            for a in m.get('walkthrough', {}).get('activities', []):
                if isinstance(a.get('control'), dict): raw[a['control'].get('id')] = a['control'].get('text')
            for coll, key in (('questions', 'question'), ('tests', 'objective'), ('risks', 'risk'), ('analytics', 'analytic')):
                for x in (m.get('walkthrough', {}).get(coll, []) if coll == 'questions' else m.get(coll, [])): raw[x.get('id')] = x.get(key)
    full = {}                                                            # compare like with like: core defaults < pack defaults < this engagement's words, applied to every library
    for lib in libs: full.update(lib.get('token_defaults') or {})
    full.update(tokens)
    for lib in libs: lib['modules'] = subst(lib['modules'], full)
    deny = build_denylist(spec, tokens); words = engagement_words(tokens, libs, deny)
    if not tokens: report.append('no tokens: the workbook has no _build sheet and --tokens was not given, so only the client name was redacted; role, department and system names are NOT tokenised - pass --tokens {"CLIENT_SHORT": .., "SYSTEM": .., "HEAD_OF_DEPARTMENT": .., ...} and rerun')
    lib_ctrls, lib_qs, lib_tests, lib_risks, lib_das = [], [], [], [], []
    for lib in libs:
        for m in lib.get('modules', []):
            for a in m.get('walkthrough', {}).get('activities', []):
                c = a.get('control')
                if isinstance(c, dict) and c.get('text'): lib_ctrls.append((lib.get('library'), m['id'], c['id'], c['text'], a.get('activity')))
            for q in m.get('walkthrough', {}).get('questions', []): lib_qs.append((lib.get('library'), m['id'], q.get('id'), q.get('question'), q.get('why')))
            for t in m.get('tests', []): lib_tests.append((lib.get('library'), m['id'], t['id'], t.get('title'), (t.get('objective') or '') + '\n' + (t.get('steps_intro') or 'Check that:') + '\n' + '\n'.join(f'{i}. {x}' for i, x in enumerate(t.get('steps') or [], 1))))
            for rk in m.get('risks', []): lib_risks.append((lib.get('library'), m['id'], rk['id'], rk.get('risk'), rk.get('sub_process')))
            for a in m.get('analytics', []): lib_das.append((lib.get('library'), m['id'], a['id'], a.get('analytic'), a.get('identifies')))
    props = []; skipped = 0
    def add(kind, harvested, text, best, sim, matched_text):
        nonlocal skipped
        status = _status(sim, text, matched_text)
        if status == 'same': return
        kind_status = status.split(' ')[0]
        h = anonymise_obj(harvested, tokens, deny); hh = _hash(h)
        if hh in learned: skipped += 1; return               # this exact item was already learned from this workbook
        props.append({'id': f'P{len(props)+1:03d}', 'kind': kind, 'status': status, 'similarity': round(sim, 2), 'matches': (best[1] + ' / ' + str(best[2])) if best else None, 'matched_text': (raw.get(best[2]) or matched_text or '')[:500] if best else None,
                      'matched_library': best[0] if best else None, 'harvested': h, 'source_hash': hh, 'review_terms': review_terms(json.dumps(h, ensure_ascii=False, default=str), known, words), 'decision': ''})
    for wt in spec['walkthroughs']:
        for s in wt.get('sections', []):
            for a in s.get('activities', []):
                if a.get('refer_to') or _empty(a.get('control')): continue
                score = lambda x: max(_sim(x[3], a['control']), _sim(x[4], a.get('activity')))      # same control slot if either the control or its activity matches
                best = max(lib_ctrls, key=score, default=None); sim = score(best) if best else 0
                add('control', {'activity': a.get('activity'), 'performed_by': a.get('performed_by'), 'frequency': a.get('frequency'), 'control': a.get('control'), 'owner': a.get('owner'), 'classification': a.get('classification'), 'type': a.get('type'), 'document': a.get('document'), 'sheet': wt.get('title'), 'section': s.get('title')}, a['control'], best, sim, best[3] if best else '')
        for q in wt.get('questions', []):
            if _empty(q.get('question')): continue
            best = max(lib_qs, key=lambda x: _sim(x[3], q['question']), default=None); sim = _sim(best[3], q['question']) if best else 0
            add('question', {'question': q.get('question'), 'why': q.get('why'), 'documents': q.get('documents'), 'interviewee': q.get('interviewee'), 'sheet': wt.get('title')}, q['question'], best, sim, best[3] if best else '')
    for t in spec['tests']:
        if _empty(t.get('title')): continue
        body = (t.get('objective') or '') + '\n' + '\n'.join(str(t.get('steps') or '').split('\n')[1:])      # objective + steps without the "Test n: title" line
        score = lambda x: max(_sim(x[3], t['title']), _sim(x[4], body))
        best = max(lib_tests, key=score, default=None); sim = score(best) if best else 0
        add('test', {k: t.get(k) for k in ('title', 'objective', 'basis', 'population', 'sample_size', 'source', 'steps', 'detail_title', 'detail_columns', 'detail_formulas', 'n_rows')}, body, best, sim, (best[4] if best else ''))
    for sec in spec['rcm']['sections']:
        for r in sec.get('rows', []):
            if _empty(r.get('risk')): continue
            best = max(lib_risks, key=lambda x: _sim(x[3], r['risk']), default=None); sim = _sim(best[3], r['risk']) if best else 0
            add('risk', {'sub_process': r.get('sub_process'), 'risk': r.get('risk'), 'section': sec.get('title')}, r['risk'], best, sim, best[3] if best else '')
    for a in spec.get('da', {}).get('analytics', []):
        if _empty(a.get('analytic')): continue
        best = max(lib_das, key=lambda x: _sim(x[3], a['analytic']), default=None); sim = _sim(best[3], a['analytic']) if best else 0
        add('analytic', {'analytic': a.get('analytic'), 'source': a.get('source'), 'identifies': a.get('identifies')}, a['analytic'], best, sim, best[3] if best else '')
    built = spec.get('build_stamp') or {}
    try: built_versions = json.loads(built.get('library_versions') or '{}')
    except Exception: built_versions = {}
    for lib in libs:
        key = 'core' if lib.get('library') == 'core' else f"pack {lib.get('process')}"
        if key in built_versions and str(built_versions[key]) != str(lib.get('version')): report.append(f"the workbook was built with {key} v{built_versions[key]}; the library is now v{lib.get('version')} - its old wording will come back as refinements, reject those unless they add substance")
    out = {'source_awp': os.path.basename(awp_path), 'harvested_on': dt.date.today().isoformat(), 'tokens_used': tokens, 'engagement_words': words, 'denylist': deny, 'export_report': report, 'already_learned': skipped,
           'summary': {k: sum(1 for p in props if p['status'].startswith(k)) for k in ('new', 'refinement')}, 'proposals': props}
    save_json(out, out_path)
    md = [f"# /awp-learn proposals from {os.path.basename(awp_path)}",
          f"{out['summary']['new']} new, {out['summary']['refinement']} refinements" + (f", {skipped} already learned earlier" if skipped else '') + '.',
          'Tokens used: ' + (', '.join(f'{k}={v}' for k, v in tokens.items()) or 'none (pass --tokens)') + '.',
          "'refinement' = a library item with the same objective exists (its text is shown): accept with \"replace\" only if the harvested version is better in substance; check that the matched item really has the same objective before replacing. 'new' = no match (the nearest item is shown for orientation only). Words under 'review before accepting' may identify the engagement: generalise or remove them in \"edits\" of the decision, or set \"redaction_checked\": true if they are plainly generic.", '']
    for p in props:
        h = p['harvested']
        md.append(f"## {p['id']} — {p['kind']} · {p['status']}" + (f" ({'vs' if p['status'].startswith('refinement') else 'nearest'} {p['matches']}, similarity {p['similarity']})" if p['matches'] else ''))
        for k, v in h.items():
            if v not in (None, '', []): md.append(f"- **{k}**: {v}")
        if p.get('matched_text'): md.append(f"- *library text (as stored, with its tokens)*: {p['matched_text']}")
        if p.get('review_terms'): md.append(f"- **review before accepting**: {', '.join(p['review_terms'])}")
        md.append('')
    with open(os.path.splitext(out_path)[0] + '.md', 'w', encoding='utf-8') as f: f.write('\n'.join(md))
    return out

def apply_proposals(proposals_path, decisions_path, folder=None):
    """decisions = {"P001": {"accept": true, "library": "core"|"pack:<name>", "module": "mod.x", "replace": "ctrl.y", "edits": {"control": "..."}, "risk": "risk.z", "control": "ctrl.q", "redaction_checked": false, "token_defaults": {"{X}": "..."}}, ...}"""
    props = load_json(proposals_path); dec = load_json(decisions_path); folder = folder or HERE
    core_p, pack_ps = find_libraries(folder); libs = {'core': load_json(core_p)} if core_p else {}
    for p in pack_ps: libs['pack:' + os.path.basename(p)[len('skill-awp-lib-pack-'):-5]] = load_json(p)
    known = known_words(list(libs.values())); tokens = props.get('tokens_used') or {}; deny = props.get('denylist') or []
    words = sorted(set(props.get('engagement_words') or []) | set(engagement_words(tokens, list(libs.values()), deny)), key=len, reverse=True)
    sheet_ids = {s['id'] for lib in libs.values() for s in lib.get('sheets', [])}
    changed = set(); log = []
    for p in props['proposals']:
        d = dec.get(p['id'])
        if not d or not d.get('accept'): continue
        lib = libs.get(d.get('library') or ('core' if p.get('matched_library') == 'core' else next((k for k in libs if k.startswith('pack:')), 'core')))
        if lib is None: raise SystemExit(f"{p['id']}: library {d.get('library')} not found (use \"core\" or \"pack:<name>\")")
        h = dict(p['harvested']); h.update(d.get('edits') or {})
        text = json.dumps(h, ensure_ascii=False, default=str)
        side = json.dumps({k: v for k, v in d.items() if k not in ('accept', 'reason', 'redaction_checked', 'edits', 'replace', 'library')}, ensure_ascii=False)   # ids, titles, token_defaults, links
        terms = review_terms(text, known, words)                         # re-scan the text as it will be written, edits included
        hard = [w for w in words if _pattern(w, ci=True).search(text) or _pattern(w, ci=True).search(side)] + [m for m in PERSON.findall(text + ' ' + side)]
        if hard: raise SystemExit(f"{p['id']}: the engagement's own words or a person's name are still in the decision or its text {hard} - replace them with tokens or general wording")
        if '[REDACTED]' in text: raise SystemExit(f"{p['id']}: '[REDACTED]' is still in the text - rewrite that part in \"edits\"")
        side_terms = [w for w in review_terms(side, known, words) if not re.fullmatch(r'[a-z]+(\.[a-z0-9_]+)+', w)]
        ok = d.get('redaction_checked'); accepted = set(ok) if isinstance(ok, list) else (set(terms) | set(side_terms) if ok is True else set())
        left = [w for w in terms + side_terms if w not in accepted]
        if left: raise SystemExit(f"{p['id']}: terms still to review {left} - generalise them in \"edits\", or list the ones that are plainly generic in \"redaction_checked\": [..] (true accepts all of them)")
        mod = next((m for m in lib['modules'] if m['id'] == d.get('module')), None)
        if mod is None:
            if not d.get('module'): raise SystemExit(f"{p['id']}: decision needs \"module\" (an existing module id, or a new id with module_title, sheet and stage)")
            if not (d.get('module_title') and d.get('sheet') and d.get('stage')): raise SystemExit(f"{p['id']}: module {d['module']} does not exist - check the id (python skill-awp.py list), or create it by giving module_title, sheet and stage")
            if d['sheet'] not in sheet_ids: raise SystemExit(f"{p['id']}: sheet {d['sheet']} is not a sheet id in the libraries ({', '.join(sorted(sheet_ids))})")
            mod = {'id': d['module'], 'title': d['module_title'], 'stage': d['stage'], 'order': int(d.get('order', 50)), 'scope_areas': [], 'walkthrough': {'sheet': d['sheet'], 'section_title': d['module_title'], 'activities': [], 'questions': []}, 'risks': [], 'tests': [], 'analytics': []}
            lib['modules'].append(mod)
        if d.get('token_defaults'): lib.setdefault('token_defaults', {}).update(d['token_defaults'])
        stamp = {'learned_on': dt.date.today().isoformat(), 'source_hash': p.get('source_hash')}        # never the file name: it carries the client
        def put_item(coll, key, new):
            if d.get('replace'):
                for i, x in enumerate(coll):
                    if (x.get(key) if key else (x.get('control') or {}).get('id')) == d['replace']: coll[i] = new; return
                raise SystemExit(f"{p['id']}: nothing with id {d['replace']} to replace in {mod['id']}")
            coll.append(new)
        def link_risk(field, item_id):
            if d.get('risk'):
                rk = next((r for m in lib['modules'] for r in m.get('risks', []) if r['id'] == d['risk']), None)
                if rk is None: raise SystemExit(f"{p['id']}: risk {d['risk']} not found")
                if item_id not in rk.setdefault(field, []): rk[field].append(item_id)
        if p['kind'] == 'control':
            cid = d.get('replace') or d.get('id') or f"ctrl.{mod['id'].split('.', 1)[-1]}_{len(mod['walkthrough']['activities'])+1}"
            new = {'activity': h.get('activity'), 'performed_by': h.get('performed_by'), 'frequency': h.get('frequency'), 'document': h.get('document'), 'control': {'id': cid, 'text': h.get('control'), 'owner': h.get('owner'), 'classification': h.get('classification'), 'type': h.get('type'), **stamp}}
            put_item(mod['walkthrough']['activities'], None, new); link_risk('controls', cid)
        elif p['kind'] == 'question':
            qid = d.get('replace') or d.get('id') or f"q.{mod['id'].split('.', 1)[-1]}_{len(mod['walkthrough']['questions'])+1}"
            new = {'id': qid, 'question': h.get('question'), 'why': h.get('why'), 'documents': h.get('documents'), 'interviewee': h.get('interviewee'), 'control': d.get('control'), **stamp}
            put_item(mod['walkthrough']['questions'], 'id', new)
        elif p['kind'] == 'test':
            steps = h.get('steps')
            if isinstance(steps, str): steps = [re.sub(r'^\s*\d+\.\s*', '', l) for l in steps.split('\n') if re.match(r'^\s*\d+\.', l)]
            tid = d.get('replace') or d.get('id') or f"test.{mod['id'].split('.', 1)[-1]}_{len(mod['tests'])+1}"
            new = {'id': tid, 'title': h.get('title'), 'objective': h.get('objective'), 'basis': h.get('basis'), 'population': h.get('population'), 'sample_size': h.get('sample_size'), 'source': h.get('source'), 'steps': steps, 'detail_title': h.get('detail_title'), 'detail_columns': h.get('detail_columns'), 'detail_formulas': h.get('detail_formulas'), 'n_rows': h.get('n_rows'), **stamp}
            put_item(mod['tests'], 'id', new); link_risk('tests', tid)
        elif p['kind'] == 'risk':
            rid = d.get('replace') or d.get('id') or f"risk.{mod['id'].split('.', 1)[-1]}_{len(mod['risks'])+1}"
            new = {'id': rid, 'sub_process': h.get('sub_process'), 'risk': h.get('risk'), 'controls': d.get('controls', []), 'tests': d.get('tests', []), **stamp}
            put_item(mod['risks'], 'id', new)
        elif p['kind'] == 'analytic':
            aid = d.get('replace') or d.get('id') or f"da.{mod['id'].split('.', 1)[-1]}_{len(mod['analytics'])+1}"
            new = {'id': aid, 'analytic': h.get('analytic'), 'source': h.get('source'), 'identifies': h.get('identifies'), 'tests': d.get('tests', []), **stamp}
            put_item(mod['analytics'], 'id', new)
        key = next(k for k, v in libs.items() if v is lib); changed.add(key)
        log.append(f"{p['id']} {p['kind']} -> {key}/{mod['id']}" + (f" (replaced {d['replace']})" if d.get('replace') else ' (added)') + (f" linked to {d['risk']}" if d.get('risk') else '') + (f" - {d['reason']}" if d.get('reason') else ''))
    for key in changed:
        lib = libs[key]; lib['version'] = int(lib.get('version', 1)) + 1; lib['updated'] = dt.date.today().isoformat()
        lib.setdefault('changelog', []).append({'version': lib['version'], 'date': lib['updated'], 'source': '/awp-learn', 'changes': [l for l in log if f"-> {key}/" in l]})
        issues = [i for i in validate(lib, 'library') if i[0] == 'error']
        if issues: raise SystemExit(f'{key}: validation errors after apply, nothing saved: {issues}')
        path = core_p if key == 'core' else os.path.join(folder, f"skill-awp-lib-pack-{key.split(':', 1)[1]}.json")
        save_json(lib, path); print('updated in place:', path, '-> version', lib['version'], '(the previous version is in the release zip / GitHub)')
    print('\n'.join(log) or 'nothing accepted')

# =============================================================================== diff and selftest
def diff_workbooks(a, b):
    wa, wb_ = load_workbook(a), load_workbook(b); out = []
    for name in sorted(set(wa.sheetnames) | set(wb_.sheetnames)):
        if name not in wa.sheetnames: out.append((name, 'only in B', 0)); continue
        if name not in wb_.sheetnames: out.append((name, 'only in A', 0)); continue
        sa, sb = wa[name], wb_[name]; n = 0; sample = []
        for r in range(1, max(sa.max_row, sb.max_row) + 1):
            for c in range(1, max(sa.max_column, sb.max_column) + 1):
                va, vb = sa.cell(r, c).value, sb.cell(r, c).value
                if (va or None) != (vb or None):
                    n += 1
                    if len(sample) < 5: sample.append(f'{get_column_letter(c)}{r}: {str(va)[:30]!r} -> {str(vb)[:30]!r}')
        if n: out.append((name, f'{n} cells differ; e.g. ' + '; '.join(sample), n))
    total = sum(x[2] for x in out)
    print(f'DIFF {os.path.basename(a)} vs {os.path.basename(b)}: {total} differing cells in {len(out)} sheets' + (' - identical' if not out else ''))
    for name, msg, _ in out: print(f'  {name}: {msg}')
    return out

def selftest(folder=None):
    import tempfile, shutil
    folder = folder or HERE; out = tempfile.mkdtemp()
    build_awp(template_spec(n_tests=12), os.path.join(out, 'template.xlsx'))
    core_p, packs = find_libraries(folder)
    msg = [f'skill-awp v{VERSION}']
    if core_p:
        core = load_json(core_p); msg.append(f"core v{core.get('version')} ({len(core.get('modules', []))} modules)")
        sel = {'client': 'Selftest Co', 'process': 'Selftest process', 'period': 'FY', 'packs': [os.path.basename(packs[0])[len('skill-awp-lib-pack-'):-5]] if packs else [], 'tokens': {'{CLIENT}': 'Selftest Co', '{CLIENT_SHORT}': 'STC', '{SYSTEM}': 'System X', '{ERP}': 'ERP Y', '{DEPARTMENT}': 'Operations', '{HEAD_OF_DEPARTMENT}': 'Head of Operations', '{MANAGER}': 'Operations Manager', '{FINANCE}': 'Finance', '{HEAD_OF_FINANCE}': 'Head of Finance', '{IT}': 'IT', '{PROCESS}': 'selftest process', '{PERIOD}': 'the period'}}
        spec = compose(sel, folder); path = build_awp(spec, os.path.join(out, 'composed.xlsx'))
        back = export_spec(path); errs = [i for i in validate(spec) if i[0] == 'error']
        for p in packs: lib = load_json(p); msg.append(f"pack {lib.get('process')} v{lib.get('version')} ({len(lib.get('modules', []))} modules)")
        ok = len(back['tests']) == len(spec['tests']) and not errs
        msg.append(f"composed {len(load_workbook(path).sheetnames)} sheets, {len(spec['tests'])} tests, validate errors {len(errs)}")
    else: ok = True; msg.append('no libraries found next to the script (template only)')
    shutil.rmtree(out, ignore_errors=True)
    print('SELFTEST', 'PASSED' if ok else 'FAILED', '-', ' · '.join(msg)); return ok

def verify(path):
    """Step-7 checks on a built workbook: references, step columns, placeholders, formulas. Prints PASS/FAIL lines; returns True when all pass."""
    wb = load_workbook(path); names = set(wb.sheetnames); rep = []; spec = export_spec(path, rep); P = spec['prefix']; results = []
    def res(name, bad, detail=''): ok = not bad; results.append(ok); print(('PASS ' if ok else 'FAIL ') + name + (f' - {detail}' if detail else '') + ('' if ok else ' - ' + ', '.join(str(b) for b in bad[:8]) + (' …' if len(bad) > 8 else '')))
    ctrls = {a.get('control_ref') for w in spec['walkthroughs'] for s in w.get('sections', []) for a in s.get('activities', []) if a.get('control_ref')}
    tests = {t['ref'] for t in spec['tests']}
    bad = [f"RCM {r.get('sn')}: {c}" for sec in spec['rcm']['sections'] for r in sec['rows'] for c in str(r.get('control_no') or '').split('\n') if c and c != '-' and c not in ctrls]
    bad += [f"{w['ref']} Q{q.get('no')}: {q['control_ref']}" for w in spec['walkthroughs'] for q in w.get('questions', []) if q.get('control_ref') and q['control_ref'] not in ctrls and q['control_ref'] != '-']
    res('control references resolve', bad, f'{len(ctrls)} controls')
    bad = [f"RCM {r.get('sn')}: {x}" for sec in spec['rcm']['sections'] for r in sec['rows'] for x in str(r.get('wp_ref') or '').split('\n') if x and x not in tests]
    bad += [f"DA {a.get('sn')}: {x}" for a in spec['da']['analytics'] for x in str(a.get('feeds') or '').split('\n') if x and x not in tests]
    bad += [f"{w['ref']}: {m}" for w in spec['walkthroughs'] for s in w.get('sections', []) for a in s.get('activities', []) for m in re.findall(r'<([A-Z]+\d+)>', str(a.get('refer_to') or '')) if m not in names]
    res('WP refs, DA feeds and <refs> point at existing sheets', bad, f'{len(tests)} tests')
    nsteps = lambda t: len(re.findall(r'^\s*\d+[.]', str(t.get('steps') or ''), re.M))
    bad = ["%s: %s columns vs %d steps" % (t['ref'], t.get('n_steps'), nsteps(t)) for t in spec['tests'] if t.get('n_steps') != nsteps(t)]
    res('step columns equal numbered audit steps', bad)
    bad = []
    for ws in wb:
        if ws.title == '_build': continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and (re.search(r'\{[A-Za-z_0-9]+\}', c.value) or re.match(r'^\[.{3,}\]$', c.value.strip())): bad.append(f'{ws.title}!{c.coordinate}')
    res('no [placeholder] or {TOKEN} cells', bad)
    bad = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    for sh, cell in re.findall(r"'([^']+)'!([A-Z]+\d+)", c.value):
                        if sh not in names: bad.append(f'{ws.title}!{c.coordinate} -> {sh}')
                        elif wb[sh][cell].value in (None, '') and not re.fullmatch(r'[A-Z]+[1-9]\d?', cell): bad.append(f'{ws.title}!{c.coordinate} -> {sh}!{cell} is empty')
    res('formulas point at existing sheets and cells', bad)
    bad = [k.get('key_process') for k in spec['scope'] if not any(sec['title'].split('. ', 1)[-1].strip().lower() == str(k.get('key_process') or '').strip().lower() for sec in spec['rcm']['sections'])]
    res('every scope key process has an RCM section', bad, f"{len(spec['scope'])} key processes, {sum(len(s['rows']) for s in spec['rcm']['sections'])} risks")
    bad = [f"{w['ref']}: {sum(1 for s in w.get('sections', []) for a in s.get('activities', []) if a.get('control') and not str(a['control']).startswith('['))} controls / {len(w.get('questions', []))} questions" for w in spec['walkthroughs'] if sum(1 for s in w.get('sections', []) for a in s.get('activities', []) if a.get('control') and not str(a['control']).startswith('[')) < 2 or len(w.get('questions', [])) < 5]
    res('each walkthrough sheet has >= 2 controls and >= 5 questions', bad)
    res('index sheet carries the build stamp', [] if any(isinstance(c.value, str) and c.value.startswith('Built with skill-awp') for row in wb[P].iter_rows() for c in row) else ['no stamp'])
    for r in rep: print('note:', r)
    ok = all(results); print('VERIFY', 'PASSED' if ok else 'FAILED', f'- {sum(results)}/{len(results)} checks'); return ok

def list_libraries(folder=None, packs=None):
    """Prints the ids a selection can use: sheets, modules, risks, tests, controls, questions, analytics and tokens with their defaults."""
    core_p, pack_ps = find_libraries(folder)
    for p in ([core_p] if core_p else []) + [x for x in pack_ps if not packs or os.path.basename(x)[len('skill-awp-lib-pack-'):-5] in packs]:
        lib = load_json(p); print(f"\n== {os.path.basename(p)}  ({lib.get('library')}{' ' + str(lib.get('process')) if lib.get('process') else ''} v{lib.get('version')})")
        print('  sheets: ' + ', '.join(f"{s['id']} ({s['title']})" for s in lib.get('sheets', [])))
        if lib.get('sheet_map'): print('  sheet_map: ' + ', '.join(f'{k} -> {v}' for k, v in lib['sheet_map'].items()))
        for m in lib.get('modules', []):
            print(f"  module {m['id']} - {m.get('title')} [{m.get('stage')} {m.get('order')}, sheet {m.get('walkthrough', {}).get('sheet')}]")
            for rk in m.get('risks', []): print(f"      risk {rk['id']} - {rk.get('sub_process')}")
            for t in m.get('tests', []): print(f"      test {t['id']} - {t.get('title')}")
            print('      controls: ' + ', '.join(a['control']['id'] for a in m.get('walkthrough', {}).get('activities', []) if isinstance(a.get('control'), dict)))
            print('      questions: ' + ', '.join(str(q.get('id')) for q in m.get('walkthrough', {}).get('questions', [])))
            if m.get('analytics'): print('      analytics: ' + ', '.join(a['id'] for a in m['analytics']))
        if lib.get('token_defaults'): print('  token defaults: ' + ', '.join(f'{k}={v}' for k, v in lib['token_defaults'].items()))
    print('\nTokens derived from the selection: {CLIENT}, {CLIENT_SHORT}, {PROCESS}, {DEPARTMENT}, {PERIOD}, {HEAD_OF_DEPARTMENT}, {MANAGER}, {STAFF}. Everything else comes from "tokens" in the selection or the defaults above.')

def check(folder=None, deny_file=None):
    """Release acceptance in one command: selftest, library validation, reference build and round-trip, learn idempotence, leak scan."""
    import tempfile, shutil
    folder = folder or HERE; out = tempfile.mkdtemp(); results = []
    def res(name, ok, detail=''): results.append((name, ok, detail)); print(('PASS ' if ok else 'FAIL ') + name + (f' - {detail}' if detail else ''))
    res('selftest', selftest(folder))
    core_p, packs = find_libraries(folder)
    for p in ([core_p] if core_p else []) + packs:
        iss = validate(load_json(p)); errs = [m for s, m in iss if s == 'error']; warns = [m for s, m in iss if s == 'warn']
        res(f'validate {os.path.basename(p)}', not errs, f'{len(errs)} errors, {len(warns)} warnings' + (': ' + '; '.join(errs[:3]) if errs else ''))
    sel_p = os.path.join(folder, 'skill-awp-ref-selection-example.json')
    if os.path.exists(sel_p):
        spec = compose(load_json(sel_p), folder); notes = [n for n in spec.pop('compose_notes', []) if not n.startswith('info:')]; errs = [m for s, m in validate(spec) if s == 'error']
        res('reference compose', not errs and not notes, f'{len(errs)} errors, {len(notes)} notes')
        path = build_awp(spec, os.path.join(out, 'reference.xlsx')); back = export_spec(path)
        same = (len(back['tests']), len(back['walkthroughs']), sum(len(s['rows']) for s in back['rcm']['sections']), sum(len(w['questions']) for w in back['walkthroughs'])) == (len(spec['tests']), len(spec['walkthroughs']), sum(len(s['rows']) for s in spec['rcm']['sections']), sum(len(w['questions']) for w in spec['walkthroughs']))
        res('reference build and export round-trip', same, f"{len(load_workbook(path).sheetnames)} sheets")
        lp = learn(path, os.path.join(out, 'learn.json'), None, folder); res('learn idempotence on the reference build', lp['summary']['new'] + lp['summary']['refinement'] == 0, str(lp['summary']))
        wb = load_workbook(path, data_only=False); bad = []
        for ws in wb:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and re.search(r'\{[A-Z_0-9]+\}|\[[^\]]{3,}\]', c.value) and ws.title != '_build': bad.append(f'{ws.title}!{c.coordinate}')
        res('no unresolved tokens or placeholders in the reference build', not bad, ', '.join(bad[:5]))
    else: res('reference selection present', False, 'skill-awp-ref-selection-example.json missing')
    # leak scan of the libraries
    deny = [w.strip() for w in open(deny_file, encoding='utf-8')] if deny_file and os.path.exists(deny_file) else []
    for p in ([core_p] if core_p else []) + packs:
        lib = load_json(p); txt = json.dumps({k: v for k, v in lib.items() if k != 'changelog'}, ensure_ascii=False)
        hits = set(re.findall(r'[A-Z]{0,3}\$\s?\d[\d,]*|\bObservation\s+\d|\bFY\s?\d{2}', txt)) | {w for w in deny if w and re.search(r'(?<![A-Za-z0-9])' + re.escape(w) + r'(?![A-Za-z0-9])', txt, re.I)}
        res(f'leak scan {os.path.basename(p)}', not hits, ', '.join(sorted(hits)[:6]))
    shutil.rmtree(out, ignore_errors=True)
    ok = all(r[1] for r in results); print('CHECK', 'PASSED' if ok else 'FAILED', f'- {sum(1 for r in results if r[1])}/{len(results)} checks'); return ok

# =============================================================================== CLI
if __name__ == '__main__':
    a = sys.argv[1:]; cmd = a[0] if a else ''
    def opt(name, default=None):
        return a[a.index(name) + 1] if name in a else default
    if cmd == 'build': build_awp(load_json(a[1]), a[2]); print('built', a[2])
    elif cmd == 'template': build_awp(template_spec(n_walkthroughs=int(opt('--walkthroughs', 2)), n_tests=int(opt('--tests', 3))), a[1]); print('template', a[1])
    elif cmd == 'compose':
        sel = load_json(a[1]); spec = compose(sel, opt('--lib')); notes = spec.pop('compose_notes', []); save_json(spec, a[2]); iss = validate(spec)
        print('composed', a[2], f"- {len(spec['walkthroughs'])} walkthrough sheets, {len(spec['tests'])} tests, {sum(len(s['rows']) for s in spec['rcm']['sections'])} risks, {len(spec['da']['analytics'])} analytics")
        for n in notes: print(f'  note: {n}')
        for s, m in iss: print(f'  {s}: {m}')
    elif cmd == 'validate':
        for p in [x for x in a[1:] if not x.startswith('--') and x != opt('--lib')]:
            iss = validate(load_json(p)); print(f"{p}: {sum(1 for s,_ in iss if s=='error')} errors, {sum(1 for s,_ in iss if s=='warn')} warnings")
            for s, m in iss: print(f'  {s}: {m}')
    elif cmd == 'export':
        rep = []; save_json(export_spec(a[1], rep), a[2]); print('exported', a[2]); [print('  note:', r) for r in rep]
    elif cmd == 'learn':
        tk = load_json(opt('--tokens')) if opt('--tokens') else {}; out = learn(a[1], a[2], tk, opt('--lib')); print('proposals', a[2], out['summary'], f"| already learned: {out['already_learned']} | export notes:", out['export_report'] or 'none')
    elif cmd == 'apply': apply_proposals(a[1], a[2], opt('--lib'))
    elif cmd == 'diff': diff_workbooks(a[1], a[2])
    elif cmd == 'selftest': sys.exit(0 if selftest(opt('--lib')) else 1)
    elif cmd == 'check': sys.exit(0 if check(opt('--lib'), opt('--deny')) else 1)
    elif cmd == 'list': list_libraries(opt('--lib'), [x for x in a[1:] if not x.startswith('--') and x != opt('--lib')])
    elif cmd == 'verify': sys.exit(0 if verify(a[1]) else 1)
    else: print(__doc__)
