"""
update.py  v1.1 - update the RCM (100) and walkthrough (200) sheets of an AWP after the walkthrough (skill awp-update).

Run from the folder that contains SKILL.md:

    python scripts/update.py selftest                                 proves the environment and prints the versions (run first)
    python scripts/update.py show   AWP.xlsx                          overview: sheets, counts, what is still open
    python scripts/update.py show   AWP.xlsx A202 [A203 ...]          one sheet as text: background, activities, questions with the staff's responses
    python scripts/update.py show   AWP.xlsx rcm [full]               the risk and control matrix as text (full: with every audit step)
    python scripts/update.py check  AWP.xlsx UPDATE.json [...]        validates update files; writes nothing
    python scripts/update.py apply  AWP.xlsx OUT.xlsx UPDATE.json [...]   applies them: OUT.xlsx + "OUT - update log.xlsx", then verifies OUT

    --force  apply although something on a changed sheet cannot be carried over (the script lists exactly what would be lost).

The model writes the update files (what changes, from which source, on what basis); this script does everything else:
it rewrites only the sheets that change, in the house format, keeps every other sheet as it is, never overwrites a staff
response unless told to, renumbers, keeps the cross-references consistent and logs every change with its before and after.
Only dependency: openpyxl. Shares its engine with awp-generate: scripts/awp.py next to this file (release zip) or
../../awp-generate/scripts/awp.py (repository). The format lives there - never restyle a workbook by hand.
"""
import copy, json, os, re, sys, zipfile, datetime as dt, importlib.util
from collections import Counter

VERSION = '1.1'
HERE = os.path.dirname(os.path.abspath(__file__)); SKILL = os.path.dirname(HERE)
try: sys.stdout.reconfigure(errors='replace')
except Exception: pass
_awp_path = next((p for p in (os.path.join(HERE, 'awp.py'), os.path.join(os.path.dirname(SKILL), 'awp-generate', 'scripts', 'awp.py')) if os.path.exists(p)), None)
if not _awp_path: raise SystemExit('awp.py not found: keep it next to update.py (release zip) or keep awp-generate beside awp-update (repository)')
_spec = importlib.util.spec_from_file_location('awp', _awp_path); awp = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(awp)
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

A_FIELDS = ('activity', 'performed_by', 'frequency', 'control', 'owner', 'classification', 'type', 'de_finding', 'document')
A_REQUIRED = ('activity', 'performed_by', 'frequency', 'control', 'owner', 'classification', 'type', 'document')
Q_FIELDS = ('question', 'why', 'documents', 'interviewee', 'response', 'control_ref')
Q_REQUIRED = ('question', 'why', 'documents', 'interviewee')
R_FIELDS = ('sub_process', 'risk', 'control_no', 'de_finding', 'audit_strategy', 'wp_ref')
R_REQUIRED = ('sub_process', 'risk', 'control_no', 'de_finding', 'audit_strategy', 'wp_ref')
T_FIELDS = ('objective', 'basis', 'population', 'sample_size', 'source', 'steps', 'detail_title', 'detail_columns', 'detail_formulas', 'n_rows')
OPS = ('background', 'section', 'activity', 'add_activity', 'remove_activity', 'question', 'add_question', 'risk', 'add_risk', 'test')
PENDING = 'to be concluded after the walkthrough'
PLANNED = re.compile(r'(?:^|\n|[(:\[]\s*)\s*expected\b|\bexpected\s+(?:process|control|activit|procedure|practice|flow|step)', re.I)   # planning labels: "Expected: ...", "Expected process", "(expected)" - what we assumed before the walkthrough
SHOULD = re.compile(r'\bshould\b|\b(?:is|are) expected to\b|\bis to be\b|\bare to be\b|\bwill be\b|\bis required to\b|\bare required to\b', re.I)   # what ought to happen, not what is done
def planned(*texts): return any(PLANNED.search(str(t)) for t in texts if not blank(t))
T_DEFAULTS = {'notes': 'Nil', 'findings': 'Nil', 'conclusion': 'To be concluded on completion of fieldwork.'}
CLASSIFICATIONS = ('preventive', 'detective', 'preventive and detective'); TYPES = ('manual', 'it', 'manual and it')

def key(v):
    """Row numbers as typed in the workbook or in an update file: 3, 3.0 and "3" are the same row."""
    if isinstance(v, bool) or v is None: return ''
    if isinstance(v, (int, float)) and float(v).is_integer(): return str(int(v))
    s = str(v).strip(); return s[:-2] if re.fullmatch(r'\d+\.0', s) else s
def blank(v): return v is None or str(v).strip() == ''
def refs_of(v): return [x.strip() for x in re.split(r'[\n,;]+', str(v or '')) if x.strip() and x.strip() != '-']
def norm(v):
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return re.sub(r'\s+', '', str(v))

# =============================================================================== read: workbook -> spec with the original row keys
def read(path):
    if not os.path.exists(path): raise SystemExit(f'not found: {path}')
    if not path.lower().endswith('.xlsx'): raise SystemExit('the workbook must be an .xlsx file (save a copy as .xlsx first)')
    rep = []; spec = awp.export_spec(path, rep); wbf = load_workbook(path); P = spec['prefix']
    if f'{P}100' not in wbf.sheetnames or not spec['walkthroughs']: raise SystemExit(f'this is not an AWP in the house format: sheet {P}100 or the {P}2xx walkthrough sheets are missing')
    # the RCM "Testing Result" is a formula to the test's Conclusion cell; a workbook saved by Excel exports its cached text - keep it a formula
    ws = wbf[f'{P}100']; hr, _ = awp._find(ws, 'Risk', exact=True); rows = [r for s in spec['rcm']['sections'] for r in s['rows']]
    h = awp._hdr_map(ws, hr) if hr else {}
    if h.get('sub-process') and h.get('risk'):
        cB, cC, cG = h['sub-process'], h['risk'], h.get('testing result'); k = 0
        for r in range(hr + 1, ws.max_row + 1):
            if ws.cell(r, 1).value is None or (ws.cell(r, cC).value is None and ws.cell(r, cB).value is None): continue
            if k < len(rows) and cG:
                v = ws.cell(r, cG).value
                if blank(v) or (isinstance(v, str) and v.startswith('=')): rows[k]['testing_result'] = ''
            k += 1
    for r in rows: r['_k'] = key(r.get('sn'))
    for w in spec['walkthroughs']:
        for s in w['sections']:
            for a in s['activities']:
                a['_k'] = key(a.get('no'))
                if 'refer_to' not in a: a['control_ref'] = a.get('control_ref') or ''      # '' stops the builder from numbering a control the workbook left unnumbered
        for q in w['questions']: q['_k'] = key(q.get('no'))
    return wbf, spec, rep

def controls(spec): return [a['control_ref'] for w in spec['walkthroughs'] for s in w['sections'] for a in s['activities'] if a.get('control_ref')]
def remaining(spec):
    out = []
    for w in spec['walkthroughs']:
        q = [key(x.get('no')) for x in w['questions'] if blank(x.get('response'))]
        c = [a['control_ref'] for s in w['sections'] for a in s['activities'] if a.get('control_ref') and blank(a.get('de_finding'))]
        p = (['background'] if planned(w.get('background')) else []) + [f"section \"{s.get('title')}\"" for s in w['sections'] if planned(s.get('title'))] + [f"activity {key(a.get('no'))}" for s in w['sections'] for a in s['activities'] if planned(a.get('activity'), a.get('control'))]
        if p: out.append((w['ref'], 'still in planning wording ("Expected ..."), not yet rewritten as the actual process', ', '.join(p)))
        if q: out.append((w['ref'], 'questions without a response', ', '.join(q)))
        if c: out.append((w['ref'], 'controls without a design-effectiveness finding', ', '.join(c)))
    r = [key(x.get('sn')) for s in spec['rcm']['sections'] for x in s['rows'] if blank(x.get('de_finding')) or str(x['de_finding']).strip().lower().startswith(PENDING)]
    if r: out.append((f"{spec['prefix']}100", 'risks without a design-effectiveness conclusion', ', '.join(r)))
    return out

# =============================================================================== show
def show(path, sheets):
    _, spec, rep = read(path); P = spec['prefix']; L = []; v = lambda x: '(blank)' if blank(x) else str(x).strip()
    full = 'full' in [x.lower() for x in sheets]; sheets = [x for x in sheets if x.lower() != 'full']
    def strategy(r):
        t = v(r.get('audit_strategy')); lines = t.split('\n')
        return t if full or len(lines) < 3 else f'{lines[0]} … (+{len(lines) - 1} lines; add "full" after rcm to print them)'
    if not sheets:
        L.append(f"{os.path.basename(path)} - {spec.get('client')} / {spec.get('process')} / {spec.get('period')}")
        L.append(f"{P}100  risk and control matrix: {sum(len(s['rows']) for s in spec['rcm']['sections'])} risks in {len(spec['rcm']['sections'])} sections")
        for w in spec['walkthroughs']:
            acts = [a for s in w['sections'] for a in s['activities']]
            L.append(f"{w['ref']}  {w.get('title')}: {len(acts)} activity rows, {sum(1 for a in acts if a.get('control_ref'))} controls, {len(w['questions'])} questions ({sum(1 for q in w['questions'] if not blank(q.get('response')))} with a response)")
        L.append(f"{len(spec['tests'])} test sheets: " + ', '.join(f"{t['ref']} {t.get('title')}" for t in spec['tests']))
        L.append('Still open:'); L += [f'  {s}: {what}: {items}' for s, what, items in remaining(spec)] or ['  nothing']
        L.append(f'Next: python scripts/update.py show "{path}" {spec["walkthroughs"][0]["ref"]}   (one sheet at a time; "rcm" for {P}100)')
    for name in sheets:
        if name.lower() in ('rcm', f'{P}100'.lower()):
            L.append(f'=== {P}100  Risk and control matrix ===')
            for s in spec['rcm']['sections']:
                L.append(f"[section] {s.get('title')}")
                for r in s['rows']:
                    L.append(f"risk {r['_k']} | controls: {', '.join(refs_of(r.get('control_no'))) or '-'} | tests: {', '.join(refs_of(r.get('wp_ref'))) or '-'}")
                    L += [f'    {f}: {v(r.get(f))}' for f in ('sub_process', 'risk', 'de_finding')] + [f'    audit_strategy: {strategy(r)}']
            continue
        w = next((x for x in spec['walkthroughs'] if x['ref'].lower() == name.lower()), None)
        t = next((x for x in spec['tests'] if x['ref'].lower() == name.lower()), None)
        if t:
            L.append(f"=== {t['ref']}  Test {t.get('no')}: {t.get('title')} ===")
            L += [f'    {f}: {v(t.get(f))}' for f in ('objective', 'basis', 'population', 'sample_size', 'source', 'steps', 'detail_title', 'detail_columns')]
            L.append('    work recorded: ' + ('yes - this sheet can no longer be changed by an update' if test_started(t) else 'no'))
            continue
        if not w: L.append(f'{name}: no such sheet (walkthrough sheets: ' + ', '.join(x['ref'] for x in spec['walkthroughs']) + ')'); continue
        L.append(f"=== {w['ref']}  {w.get('title')} ==="); L.append('BACKGROUND' + (' | PLANNING WORDING - rewrite as the actual process' if planned(w.get('background')) else '')); L.append('    ' + v(w.get('background')).replace('\n', '\n    '))
        L.append('ACTIVITIES')
        for s in w['sections']:
            L.append(f"[section] {s.get('title')}")
            for a in s['activities']:
                if 'refer_to' in a: L.append(f"activity {a['_k']} | cross-reference only: {a['refer_to']}"); continue
                L.append(f"activity {a['_k']} | control {a.get('control_ref') or '(none)'}" + (' | PLANNING WORDING - rewrite as the actual process' if planned(a.get('activity'), a.get('control')) else ''))
                L += [f'    {f}: {v(a.get(f))}' for f in A_FIELDS]
        L.append('QUESTIONS')
        for q in w['questions']:
            L.append(f"question {q['_k']} | control {q.get('control_ref') or '(none)'} | interviewee: {v(q.get('interviewee'))}")
            L += [f'    {f}: {v(q.get(f))}' for f in ('question', 'why', 'documents', 'response')]
        rows = [r for s in spec['rcm']['sections'] for r in s['rows'] if set(refs_of(r.get('control_no'))) & {a.get('control_ref') for s2 in w['sections'] for a in s2['activities']}]
        L.append(f'RISKS IN {P}100 THAT RELY ON THESE CONTROLS' + ('' if rows else ': none'))
        for r in rows:
            L.append(f"risk {r['_k']} | controls: {', '.join(refs_of(r.get('control_no')))} | tests: {', '.join(refs_of(r.get('wp_ref'))) or '-'}")
            L += [f'    {f}: {v(r.get(f))}' for f in ('sub_process', 'risk', 'de_finding')] + [f'    audit_strategy: {strategy(r)}']
    L += [f'note: {r}' for r in rep]
    print('\n'.join(L)); return L

# =============================================================================== plan: update files -> new spec + change log
def test_started(t): return bool(t.get('rows')) or any(not blank(t.get(k)) and str(t.get(k)).strip() != d for k, d in T_DEFAULTS.items())

def plan(spec0, updates):
    """Applies the update files to a copy of the exported spec. Returns (spec, log, errors, warnings, touched sheets, sources, open points)."""
    spec = copy.deepcopy(spec0); P = spec['prefix']; rcm = f'{P}100'; log = []; err = []; warn = []; touched = set(); sources = {}; open_points = []; seen = set()
    wts = {w['ref']: w for w in spec['walkthroughs']}; tests = {t['ref']: t for t in spec['tests']}; renumber = set(); resteps = []
    allrows = lambda: [r for s in spec['rcm']['sections'] for r in s['rows']]
    def find(items, k, what, where):
        hit = [x for x in items if x.get('_k') == key(k) and not x.get('_removed')]
        if len(hit) == 1: return hit[0]
        err.append(f'{where}: {what} {k} ' + ('not found' if not hit else 'appears more than once on the sheet - renumber the rows in Excel first')); return None
    def entry(sheet, item, field, before, after, ch, obj=None):
        log.append({'sheet': sheet, 'item': item, 'field': field, 'before': before, 'after': after, 'source': ', '.join(ch.get('source') or []), 'basis': ch.get('basis') or '', 'obj': obj}); touched.add(sheet)
    def setf(obj, fields, ch, sheet, item, where, allowed):
        for f, new in fields.items():
            if f not in allowed: err.append(f'{where}: "{f}" is not a field of this item (allowed: {", ".join(allowed)})'); continue
            if isinstance(new, list) and f not in ('detail_columns',): new = '\n'.join(str(x) for x in new)
            if (id(obj), f) in seen: err.append(f'{where}: {item} field "{f}" is changed twice - merge the two changes'); continue
            seen.add((id(obj), f)); old = obj.get(f)
            if (old or '') == (new or '') or (isinstance(old, str) and isinstance(new, str) and old.strip() == new.strip()): warn.append(f'{where}: {item} field "{f}" already has this text - nothing to change'); continue
            if f == 'de_finding' and not blank(old) and re.search(r'prior[- ]year|observation \d', str(old), re.I) and not re.search(r'prior[- ]year|observation \d', str(new or ''), re.I):
                warn.append(f'{where}: {item} "de_finding" drops the prior-year observation that was in the cell - keep it and say whether the walkthrough shows it remediated')
            obj[f] = new; entry(sheet, item, f, old, new, ch, obj)
    for fname, upd in updates:
        if not isinstance(upd, dict) or not isinstance(upd.get('changes'), list): err.append(f'{fname}: the file must be a JSON object with a "changes" list'); continue
        src = upd.get('sources') or {}
        for k, d in src.items():
            if k in sources and sources[k] != d: err.append(f'{fname}: source id {k} is described differently in another update file - use one id per document')
            sources[k] = d
        open_points += [(fname, str(p)) for p in (upd.get('open_points') or [])]
        for n, ch in enumerate(upd['changes'], 1):
            where = f'{fname} change {n}'
            if not isinstance(ch, dict): err.append(f'{where}: not an object'); continue
            op, sheet = ch.get('op'), str(ch.get('sheet') or '')
            if op not in OPS: err.append(f'{where}: "op" must be one of {", ".join(OPS)}'); continue
            if sheet.lower() == 'rcm': sheet = rcm
            s = ch.get('source'); ch['source'] = [s] if isinstance(s, str) else (s or [])
            if not ch['source'] or any(x not in src for x in ch['source']): err.append(f'{where}: "source" must list ids from this file\'s "sources" ({", ".join(src) or "none declared"})')
            if blank(ch.get('basis')): err.append(f'{where}: "basis" is missing - say what in the source supports the change (who said it, or the SOP section)')
            fields = ch.get('set') or {}
            if not isinstance(fields, dict): err.append(f'{where}: "set" must be an object of field: text'); continue
            if op in ('risk', 'add_risk'):
                if sheet != rcm: err.append(f'{where}: risks are on sheet {rcm}'); continue
                if op == 'risk':
                    r = find(allrows(), ch.get('sn'), 'risk', where)
                    if r: setf(r, fields, ch, rcm, f"risk {r['_k']}", where, R_FIELDS)
                else:
                    miss = [f for f in R_REQUIRED if blank(fields.get(f))]
                    if miss: err.append(f'{where}: a new risk needs {", ".join(miss)}'); continue
                    if 'after' in ch:
                        a = find(allrows(), ch['after'], 'risk', where); sec = next((s for s in spec['rcm']['sections'] if a in s['rows']), None) if a else None
                        pos = sec['rows'].index(a) + 1 if sec else 0
                        while sec and pos < len(sec['rows']) and sec['rows'][pos].get('_new'): pos += 1
                    else:
                        sec = next((s for s in spec['rcm']['sections'] if str(s.get('title') or '').lower().startswith(str(ch.get('section') or '\0').lower())), None); pos = len(sec['rows']) if sec else 0
                        if not sec: err.append(f'{where}: give "after" (the risk number the new risk follows) or "section" (the start of an existing section title)')
                    if sec:
                        r = {'_new': True, 'testing_result': ''}; sec['rows'].insert(pos, r); renumber.add(rcm)
                        setf(r, fields, ch, rcm, 'new risk', where, R_FIELDS)
                continue
            if op == 'test':
                t = tests.get(sheet)
                if not t: err.append(f'{where}: {sheet} is not a test sheet ({", ".join(tests)})'); continue
                if test_started(t): err.append(f'{where}: work is already recorded on {sheet}; change that test in Excel, not through an update'); continue
                old_steps = t.get('steps'); setf(t, fields, ch, sheet, f"test {t.get('no')}", where, T_FIELDS)
                if 'steps' in fields: t.pop('n_steps', None); resteps.append((t, old_steps, ch))
                if 'detail_columns' in fields and not isinstance(fields['detail_columns'], list): err.append(f'{where}: "detail_columns" must be a list of column names')
                if 'detail_columns' in fields and t.get('detail_formulas') and 'detail_formulas' not in fields: warn.append(f'{where}: {sheet} has computed columns ({", ".join(t["detail_formulas"])}); their formulas refer to column letters - give "detail_formulas" again if the columns moved')
                continue
            w = wts.get(sheet)
            if not w: err.append(f'{where}: {sheet or "(no sheet)"} is not a walkthrough sheet ({", ".join(wts)})'); continue
            acts = [a for s in w['sections'] for a in s['activities']]
            if op == 'background':
                new = ch.get('text')
                if blank(new): err.append(f'{where}: "text" is missing'); continue
                if ch.get('mode') == 'append': new = (str(w.get('background') or '').rstrip() + '\n\n' + str(new).strip()).strip()
                elif ch.get('mode') not in (None, 'replace'): err.append(f'{where}: "mode" is replace (default) or append'); continue
                setf(w, {'background': new}, ch, sheet, 'background', where, ('background',))
            elif op == 'section':
                title = str(ch.get('section') or '').strip(); hit = [s for s in w['sections'] if title and str(s.get('title') or '').lower().startswith(title.lower())]
                if len(hit) != 1: err.append(f'{where}: "section" must be the start of exactly one existing section title on {sheet} ({"; ".join(str(s.get("title")) for s in w["sections"])})'); continue
                setf(hit[0], fields, ch, sheet, f'section "{hit[0].get("title")}"', where, ('title',))
            elif op == 'activity':
                a = find(acts, ch.get('no'), 'activity', where)
                if not a: continue
                item = f"activity {a['_k']}"
                if 'refer_to' in a:
                    if set(fields) == {'refer_to'}: setf(a, fields, ch, sheet, item, where, ('refer_to',)); continue
                    miss = [f for f in A_REQUIRED if blank(fields.get(f))]
                    if miss: err.append(f'{where}: {item} is a cross-reference row; to turn it into an activity give {", ".join(miss)}'); continue
                    entry(sheet, item, 'refer_to', a.pop('refer_to'), None, ch, a); a['control_ref'] = ''
                if 'control_ref' in fields:
                    if a.get('control_ref'): err.append(f'{where}: {item} already carries control {a["control_ref"]}; control numbers never change'); continue
                    a['_want_ref'] = fields.pop('control_ref')
                setf(a, fields, ch, sheet, item, where, A_FIELDS)
                if not blank(a.get('control')) and not a.get('control_ref'): a['_needs_ref'] = True
            elif op == 'add_activity':
                miss = [f for f in A_REQUIRED if blank(fields.get(f))]
                if miss: err.append(f'{where}: a new activity row needs {", ".join(miss)} (when there is no control, say so in "control" and record the gap in "de_finding")'); continue
                if 'after' in ch:
                    anchor = find(acts, ch['after'], 'activity', where); sec = next((s for s in w['sections'] if anchor in s['activities']), None) if anchor else None
                    pos = sec['activities'].index(anchor) + 1 if sec else 0
                    while sec and pos < len(sec['activities']) and sec['activities'][pos].get('_new'): pos += 1
                else:
                    title = str(ch.get('section') or '').strip(); sec = next((s for s in w['sections'] if title and str(s.get('title') or '').lower().startswith(title.lower())), None)
                    if not sec and title and ch.get('new_section'): sec = {'title': title, 'activities': []}; w['sections'].append(sec); entry(sheet, 'new section', 'title', None, title, ch)
                    if not sec: err.append(f'{where}: give "after" (the activity number the new row follows) or "section" (the start of an existing section title; add "new_section": true to create it)'); continue
                    pos = len(sec['activities'])
                a = {'_new': True, '_needs_ref': True, 'control_ref': ''}
                if 'control_ref' in fields: a['_want_ref'] = fields.pop('control_ref')
                sec['activities'].insert(pos, a); renumber.add(sheet); setf(a, fields, ch, sheet, 'new activity', where, A_FIELDS)
            elif op == 'remove_activity':
                a = find(acts, ch.get('no'), 'activity', where)
                if not a: continue
                a['_removed'] = True; renumber.add(sheet); entry(sheet, f"activity {a['_k']}", 'row', a.get('refer_to') or f"{a.get('activity')} | control {a.get('control_ref') or '-'}: {a.get('control')}", '(removed)', ch)
                cr = a.get('control_ref')
                if cr:
                    for q in [q for x in spec['walkthroughs'] for q in x['questions'] if q.get('control_ref') == cr]: entry(sheet, f"question {q['_k']}", 'control_ref', cr, None, ch, q); q['control_ref'] = None
                    for r in [r for r in allrows() if cr in refs_of(r.get('control_no'))]:
                        left = [x for x in refs_of(r['control_no']) if x != cr]; entry(rcm, f"risk {r.get('_k', 'new')}", 'control_no', r['control_no'], '\n'.join(left) or '-', ch, r); r['control_no'] = '\n'.join(left) or '-'
                        if not left: warn.append(f'{where}: risk {r.get("_k")} is left without a control - state the design gap in its "de_finding"')
            elif op == 'question':
                q = find(w['questions'], ch.get('no'), 'question', where)
                if not q: continue
                if 'response' in fields and not blank(q.get('response')):
                    mode = ch.get('mode') or 'fill'
                    if mode == 'append': fields['response'] = str(q['response']).rstrip() + '\n\n' + str(fields['response']).strip()
                    elif mode != 'replace': err.append(f'{where}: question {q["_k"]} already has the staff\'s response; add "mode": "append" to add to it or "mode": "replace" to rewrite it (the old text goes to the log)'); continue
                setf(q, fields, ch, sheet, f"question {q['_k']}", where, Q_FIELDS)
            elif op == 'add_question':
                miss = [f for f in Q_REQUIRED if blank(fields.get(f))]
                if miss: err.append(f'{where}: a new question needs {", ".join(miss)}'); continue
                q = {'_new': True}; pos = len(w['questions'])
                if 'after' in ch:
                    anchor = find(w['questions'], ch['after'], 'question', where)
                    if anchor:
                        pos = w['questions'].index(anchor) + 1
                        while pos < len(w['questions']) and w['questions'][pos].get('_new'): pos += 1
                w['questions'].insert(pos, q); renumber.add(sheet); setf(q, fields, ch, sheet, 'new question', where, Q_FIELDS)
    # control numbers for new controls: the number asked for if it is free, otherwise the next free number on the sheet; existing numbers never move
    for w in spec['walkthroughs']:
        for s in w['sections']: s['activities'] = [a for a in s['activities'] if not a.get('_removed')]
        w['sections'] = [s for s in w['sections'] if s['activities'] or not any(e['sheet'] == w['ref'] and e['after'] == '(removed)' for e in log)]
        acts = [a for s in w['sections'] for a in s['activities']]; used = {a['control_ref'] for a in acts if a.get('control_ref')} | {a['control_ref'] for x in spec0['walkthroughs'] if x['ref'] == w['ref'] for s in x['sections'] for a in s['activities'] if a.get('control_ref')}
        nxt = max([int(m.group(1)) for c in used for m in [re.fullmatch(re.escape(w['ref']) + r'-(\d+)', c)] if m] + [0]) + 1
        for a in acts:
            if not a.get('_needs_ref') or a.get('control_ref'): continue
            want = a.get('_want_ref')
            if want and (want in used or not re.fullmatch(re.escape(w['ref']) + r'-\d+', str(want))): err.append(f"{w['ref']}: control number {want} is taken or is not of the form {w['ref']}-n; leave \"control_ref\" out and the next free number ({w['ref']}-{nxt}) is used"); continue
            if not want:
                while f"{w['ref']}-{nxt}" in used: nxt += 1
                want = f"{w['ref']}-{nxt}"
            a['control_ref'] = want; used.add(want); log.append({'sheet': w['ref'], 'item': 'new activity' if a.get('_new') else f"activity {a['_k']}", 'field': 'control_ref', 'before': None, 'after': want, 'source': '', 'basis': 'control number assigned by the script', 'obj': a})
        if w['ref'] in renumber:
            for i, a in enumerate(acts, 1): a['no'] = i
            for i, q in enumerate(w['questions'], 1): q['no'] = i
    for t, old_steps, ch in resteps:                       # the RCM row that quotes a test's full steps follows the test
        for r in allrows():
            if not blank(old_steps) and str(r.get('audit_strategy') or '').strip() == str(old_steps).strip() and (id(r), 'audit_strategy') not in seen:
                r['audit_strategy'] = t['steps']; entry(rcm, f"risk {r.get('_k', 'new')}", 'audit_strategy', old_steps, t['steps'], dict(ch, basis=f"follows the audit steps of {t['ref']}: " + str(ch.get('basis') or '')), r)
    if rcm in renumber:
        for i, r in enumerate(allrows(), 1): r['sn'] = i
    # consistency of the result
    ctl = controls(spec); dup = [c for c, n in Counter(ctl).items() if n > 1]
    if dup: err.append('control numbers used twice: ' + ', '.join(dup))
    for r in allrows():
        bad = [c for c in refs_of(r.get('control_no')) if c not in ctl]
        if bad: err.append(f"{rcm} risk {r.get('sn')}: control {', '.join(bad)} is not on any walkthrough sheet (controls: use the numbers `show` prints; a new control gets the next free number on its sheet)")
        bad = [x for x in refs_of(r.get('wp_ref')) if x not in tests]
        if bad: err.append(f"{rcm} risk {r.get('sn')}: {', '.join(bad)} is not a test sheet")
    for w in spec['walkthroughs']:
        for q in w['questions']:
            if q.get('control_ref') and q['control_ref'] != '-' and q['control_ref'] not in ctl: err.append(f"{w['ref']} question {q.get('no')}: control {q['control_ref']} does not exist")
        for a in [a for s in w['sections'] for a in s['activities'] if 'refer_to' not in a]:
            if not blank(a.get('classification')) and str(a['classification']).strip().lower() not in CLASSIFICATIONS and any(e['obj'] is a and e['field'] == 'classification' for e in log): warn.append(f"{w['ref']} activity {a.get('no')}: classification \"{a['classification']}\" - the house values are Preventive / Detective / Preventive and detective")
            if not blank(a.get('type')) and str(a['type']).strip().lower() not in TYPES and any(e['obj'] is a and e['field'] == 'type' for e in log): warn.append(f"{w['ref']} activity {a.get('no')}: type \"{a['type']}\" - the house values are Manual / IT / Manual and IT")
    for e in log:
        txt = json.dumps(e['after'], ensure_ascii=False) if isinstance(e['after'], (list, dict)) else str(e['after'] or '')
        if e['field'] in ('activity', 'control', 'background', 'title') and e['sheet'] in wts:
            if PLANNED.search(txt): err.append(f"{e['sheet']} {e['item']} \"{e['field']}\": the new text still says \"{PLANNED.search(txt).group(0)}\" - write what the walkthrough established as fact (who does what, when, with what evidence); if the sources do not establish it, leave the row unchanged")
            elif e['field'] in ('activity', 'control') and SHOULD.search(txt): warn.append(f"{e['sheet']} {e['item']} \"{e['field']}\": \"{SHOULD.search(txt).group(0)}\" describes what ought to happen - write what is done, in the present tense (a requirement the client does not follow belongs in the design-effectiveness finding)")
        if re.search(r'\{[A-Za-z_0-9]+\}', txt) or re.match(r'^\[.{3,}\]$', txt.strip()): err.append(f"{e['sheet']} {e['item']} \"{e['field']}\": the new text still holds a {{TOKEN}} or a [placeholder] - write the engagement's actual words")
    before = {m for s, m in awp.validate(_clean(spec0))}
    warn += [m for s, m in awp.validate(_clean(spec)) if m not in before]
    if any(e['sheet'] in tests for e in log) or rcm in renumber: touched.add(rcm)
    for w in spec['walkthroughs']:                         # a design conclusion on a row that still describes the assumed process
        for a in [a for s in w['sections'] for a in s['activities'] if planned(a.get('activity'), a.get('control'))]:
            if any(e.get('obj') is a and e['field'] == 'de_finding' for e in log): warn.append(f"{w['ref']} activity {a.get('_k', 'new')}: a design-effectiveness finding is given but the activity or control still reads \"Expected ...\" - rewrite both as the actual process in the same change")
    for e in log:                                          # final numbers in the log (rows are renumbered after additions and removals)
        o = e.pop('obj', None)
        if o is None or o.get('_removed'): continue
        now = o.get('no', o.get('sn'))
        if now is not None and (o.get('_new') or (e['item'].split(' ')[0] in ('activity', 'question', 'risk') and key(now) != o.get('_k'))): e['item'] += f" (now no. {now})"
    return spec, log, err, warn, touched, sources, open_points

def _clean(spec):
    s = copy.deepcopy(spec)
    def strip(o):
        if isinstance(o, dict):
            for k in [k for k in o if k.startswith('_')]: del o[k]
            for v in o.values(): strip(v)
        elif isinstance(o, list): [strip(v) for v in o]
    strip(s); return s

# =============================================================================== apply: regenerate the changed sheets inside the workbook
def conclusion_rows(wbf, spec, rebuilt):
    out = {}
    for t in spec['tests']:
        if t['ref'] in rebuilt: out[t['ref']] = awp.test_layout(t)[7]
        elif t['ref'] in wbf.sheetnames:
            r, _ = awp._find(wbf[t['ref']], 'Conclusion', col='A')
            if r: out[t['ref']] = r
    return out

def render(wb, spec, name, wbf_for_rows, rebuilt=()):
    """Builds sheet `name` of `spec` into workbook `wb` (appended) with the engine's own builders."""
    P = spec['prefix']; s = _clean(spec)
    if name == f'{P}100': awp.build_rcm(wb, s, P, conclusion_rows(wbf_for_rows, s, rebuilt))
    elif name in [w['ref'] for w in s['walkthroughs']]: awp.build_walkthrough(wb, s, P, next(w for w in s['walkthroughs'] if w['ref'] == name))
    else: awp.build_test(wb, s, P, next(t for t in s['tests'] if t['ref'] == name))
    return wb[name]

def not_carried(path, wbf, spec0, touched):
    """What a changed sheet holds that the house layout cannot carry: free text outside the tables, pictures, charts, comments, embedded files."""
    hard = []; soft = []
    with zipfile.ZipFile(path) as z: names = z.namelist()
    if any(n.startswith('xl/embeddings/') for n in names): hard.append('the workbook has embedded files (documents inserted as objects); saving it from a script drops them from every sheet')
    if any(n.startswith('xl/threadedComments/') for n in names): hard.append('the workbook has Excel 365 threaded comments; saving it from a script turns them into plain notes or drops them, on every sheet')
    if any(n.startswith('xl/media/') for n in names):
        try: import PIL  # noqa: F401
        except ImportError: hard.append('the workbook has pictures and this environment cannot keep them (the Pillow package is missing); they would be dropped from every sheet')
    tmp = Workbook()
    for name in sorted(touched):
        ws = wbf[name]
        if getattr(ws, '_images', None): soft.append(f'{name}: {len(ws._images)} picture(s) on the sheet')
        if getattr(ws, '_charts', None): soft.append(f'{name}: {len(ws._charts)} chart(s) on the sheet')
        new = render(tmp, spec0, name, wbf); have = {norm(c.value) for row in new.iter_rows() for c in row if not blank(c.value)}
        for row in ws.iter_rows():
            for c in row:
                if c.comment: soft.append(f'{name}!{c.coordinate}: cell comment "{str(c.comment.text)[:120]}"')
                if not blank(c.value) and norm(c.value) not in have and not (isinstance(c.value, str) and c.value.startswith('=')): soft.append(f'{name}!{c.coordinate}: text outside the house layout "{str(c.value)[:120]}"')
    return hard, soft

def write_log(path, src_name, out_name, log, sources, open_points, left, warn):
    wb = Workbook(); ws = wb.active; ws.title = 'Changes'; F, put = awp.F, awp.put
    put(ws, 'A1', f'Update log - {out_name} (from {src_name}) - awp-update v{VERSION} - {dt.date.today().isoformat()}', F(True))
    hdr = ['No.', 'Sheet', 'Item', 'Field', 'Before', 'After', 'Source', 'Basis in the source']
    for j, h in enumerate(hdr): put(ws, f'{get_column_letter(j+1)}3', h, F(True), awp.F_HDR, awp.CTRW, awp.B_THIN)
    show_v = lambda v: '' if v is None else (json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else str(v))
    for i, e in enumerate(log, 1):
        vals = [i, e['sheet'], e['item'], e['field'], show_v(e['before']), show_v(e['after']), e['source'], e['basis']]
        for j, v in enumerate(vals): put(ws, f'{get_column_letter(j+1)}{3+i}', v, F(), align=awp.CTR if j == 0 else awp.LEFTW, border=awp.B_THIN)
        awp.fit_height(ws, 3 + i, [(vals[4], 55), (vals[5], 55), (vals[7], 45)], min_h=13.5)
    awp.widths(ws, {'A': 6, 'B': 8, 'C': 22, 'D': 15, 'E': 55, 'F': 55, 'G': 14, 'H': 45}); ws.freeze_panes = 'A4'; awp.finish(ws)
    for title, head, rows, w in [('Sources', ['Id', 'Document'], sorted(sources.items()), [10, 110]), ('Open points', ['From', 'Point for the team to resolve'], open_points, [24, 110]),
                                 ('Still open', ['Sheet', 'What', 'Items'], left, [10, 50, 70]), ('Warnings', ['Warning from the script'], [(w_,) for w_ in warn], [140])]:
        s2 = wb.create_sheet(title)
        for j, h in enumerate(head): put(s2, f'{get_column_letter(j+1)}1', h, F(True), awp.F_HDR, awp.CTRW, awp.B_THIN)
        for i, row in enumerate(rows, 2):
            for j, v in enumerate(row): put(s2, f'{get_column_letter(j+1)}{i}', v, F(), align=awp.LEFTW, border=awp.B_THIN)
            awp.fit_height(s2, i, [(row[-1], w[-1])], min_h=13.5)
        awp.widths(s2, {get_column_letter(j+1): x for j, x in enumerate(w)}); awp.finish(s2)
    wb.save(path)

def load_updates(paths):
    out = []
    for p in paths:
        try: out.append((os.path.basename(p), awp.load_json(p)))
        except Exception as e: raise SystemExit(f'{p}: cannot be read as JSON ({e})')
    return out

def run(path, out, files, force=False, write=True):
    if write and os.path.abspath(out) == os.path.abspath(path): raise SystemExit('OUT must be a new file: the workbook you were given is never overwritten')
    wbf, spec0, rep = read(path); P = spec0['prefix']
    spec, log, err, warn, touched, sources, open_points = plan(spec0, load_updates(files))
    for w in warn: print('warn:', w)
    for e in err: print('error:', e)
    if err: print(f'UPDATE NOT APPLIED - {len(err)} error(s); correct the update file(s) and run again'); return False
    if not log: print('UPDATE NOT APPLIED - the update files change nothing'); return False
    hard, soft = not_carried(path, wbf, spec0, touched)
    for x in hard + soft: print('would be lost:', x)
    for e in log:
        if e['field'] == 'control_ref' and e['before'] is None and not e['source']: print(f"new control number: {e['after']} ({e['sheet']} {e['item']})")
    by_sheet = Counter(e['sheet'] for e in log)
    summary = f"{len(log)} cells changed: " + ', '.join(f'{s} {n}' for s, n in sorted(by_sheet.items()))
    if not write:
        print('CHECK PASSED -', summary + (f' - but {len(hard + soft)} item(s) on the changed sheets cannot be carried over (listed above): move them into the table cells in Excel first, or apply with --force to drop them' if hard + soft else '')); return True
    if (hard or soft) and not force:
        print(f'UPDATE NOT APPLIED - {len(hard + soft)} item(s) cannot be carried over (listed above). Ask the user: move them into the table cells in Excel and run again, or run with --force to drop them. Nothing was written.'); return False
    tests = {t['ref'] for t in spec['tests']}; rebuilt = touched & tests; order = wbf.sheetnames[:]
    for name in sorted(touched, key=lambda n: (n == f'{P}100', n)):          # tests and walkthroughs first, the RCM last (it links to the tests' Conclusion cells)
        del wbf[name]; ws = render(wbf, spec, name, wbf, rebuilt); wbf.move_sheet(ws, offset=order.index(name) - wbf.sheetnames.index(name))
    wbf.active = wbf.sheetnames.index(P) if P in wbf.sheetnames else 0
    for ws in wbf: ws.sheet_view.tabSelected = (ws.title == wbf.active.title)
    if '_build' in wbf.sheetnames:
        b = wbf['_build']; r = b.max_row + 1; b[f'A{r}'] = 'updated'; b[f'B{r}'] = f"awp-update v{VERSION} · {dt.date.today().isoformat()} · {summary} · sources: {'; '.join(sources.values())}"[:2000]
        b[f'A{r}'].font = awp.F(); b[f'B{r}'].font = awp.F()
    wbf.save(out); left = remaining(spec)
    logp = os.path.splitext(out)[0] + ' - update log.xlsx'; write_log(logp, os.path.basename(path), os.path.basename(out), log, sources, open_points, left, warn + [f'dropped with --force: {x}' for x in hard + soft])
    ok = awp.verify(out)
    print(f"UPDATE APPLIED - {summary} - sheets rewritten: {', '.join(sorted(touched))} - every other sheet is unchanged")
    print(f'workbook: {out}\nupdate log: {logp}')
    for s, what, items in left: print(f'still open - {s}: {what}: {items}')
    for f, p in open_points: print(f'open point ({f}): {p}')
    if not ok: print('The updated workbook did not pass verify (lines above): correct the update file(s) and apply again from the original workbook.')
    return ok

# =============================================================================== selftest
def selftest():
    import tempfile, shutil
    d = tempfile.mkdtemp(); ok = False; msg = ''
    try:
        act = lambda n, ref: {'no': n, 'activity': f'Activity {n} is performed by the Manager in the system.', 'performed_by': 'Manager', 'frequency': 'Monthly', 'control_ref': ref, 'control': f'Control {n} is evidenced by a signed form.', 'owner': 'Head of Department', 'classification': 'Preventive', 'type': 'Manual', 'de_finding': '', 'document': 'Signed form'}
        q = lambda n, ref: {'no': n, 'question': f'Could you walk us through step {n}?', 'why': 'To understand the control.', 'documents': 'The form', 'interviewee': 'Manager', 'response': '', 'control_ref': ref}
        steps = 'Test 1: Approvals\nCheck that:\n1. The form was approved.\n2. The form was dated.\n3. The form was filed.'
        spec = {'prefix': 'A', 'client': 'Selftest Co', 'process': 'Selftest process', 'period': 'FY',
                'scope': [{'no': 1, 'key_process': 'Approvals', 'areas': ['Check that forms are approved.'], 'covered': 'Yes'}],
                'rcm': {'sections': [{'title': '1. Approvals', 'rows': [{'sn': 1, 'sub_process': 'Approvals', 'risk': 'Forms may not be approved. This may result in unauthorised transactions.', 'control_no': 'A201-1\nA201-2', 'de_finding': 'To be concluded after the walkthrough.', 'audit_strategy': steps, 'wp_ref': 'A301'}]}]},
                'walkthroughs': [{'ref': 'A201', 'title': 'Approvals', 'index_text': 'Approvals', 'background': 'The department approves forms.', 'sections': [{'title': '1.1 Approvals', 'activities': [act(1, 'A201-1'), act(2, 'A201-2')]}], 'questions': [q(i, 'A201-1') for i in range(1, 6)]}],
                'tests': [{'ref': 'A301', 'no': 1, 'title': 'Approvals', 'objective': 'To check that forms are approved.', 'basis': 'Haphazard.', 'population': 'Forms in the period', 'sample_size': 'Ten', 'source': 'Form listing', 'steps': steps, 'detail_title': 'Form details', 'detail_columns': ['Form', 'Date'], 'n_rows': 3}],
                'da': {'analytics': [{'sn': 1, 'analytic': 'Forms without approval date.', 'source': 'Form listing', 'identifies': 'Unapproved forms.', 'feeds': 'A301', 'result': '', 'status': 'Planned'}]}}
        src = os.path.join(d, 'awp.xlsx'); awp.build_awp(spec, src)
        wb = load_workbook(src); ws = wb['A201']; r, _ = awp._find(ws, 'Question', col='B', exact=True); ws[f'J{r+1}'] = 'Staff note: the Manager approves in the system.'; wb.save(src)
        upd = {'sources': {'M1': 'Walkthrough minutes (selftest)', 'N': 'Staff notes in the workbook'}, 'open_points': ['Obtain the approval matrix.'], 'changes': [
            {'sheet': 'A201', 'op': 'background', 'mode': 'append', 'text': 'The walkthrough confirmed that approvals moved into the system in the period.', 'source': ['M1'], 'basis': 'Manager described the change.'},
            {'sheet': 'A201', 'op': 'activity', 'no': 1, 'set': {'performed_by': 'Senior Manager', 'de_finding': 'No design gap noted.'}, 'source': ['M1'], 'basis': 'Manager: the Senior Manager approves.'},
            {'sheet': 'A201', 'op': 'add_activity', 'after': 1, 'set': {'activity': 'The Finance Manager reviews the approved forms every month.', 'performed_by': 'Finance Manager', 'frequency': 'Monthly', 'control': 'The monthly review is signed off on the form listing.', 'owner': 'Finance Manager', 'classification': 'Detective', 'type': 'Manual', 'de_finding': 'No design gap noted.', 'document': 'Signed form listing'}, 'source': ['M1'], 'basis': 'Finance Manager described the review.'},
            {'sheet': 'A201', 'op': 'question', 'no': 1, 'mode': 'append', 'set': {'response': 'Minutes: approval is by the Senior Manager.'}, 'source': ['M1'], 'basis': 'Minutes.'},
            {'sheet': 'A201', 'op': 'question', 'no': 2, 'set': {'response': 'Forms are filed monthly.'}, 'source': ['M1'], 'basis': 'Minutes.'},
            {'sheet': 'A201', 'op': 'add_question', 'set': {'question': 'Who reviews the listing?', 'why': 'A new detective control.', 'documents': 'The signed listing', 'interviewee': 'Finance Manager', 'response': 'The Finance Manager.', 'control_ref': 'A201-3'}, 'source': ['M1'], 'basis': 'Minutes.'},
            {'sheet': 'A100', 'op': 'risk', 'sn': 1, 'set': {'control_no': 'A201-1\nA201-2\nA201-3', 'de_finding': 'No design gap noted. The monthly review by the Finance Manager (A201-3) was identified in the walkthrough.'}, 'source': ['M1', 'N'], 'basis': 'Minutes and staff notes.'}]}
        up = os.path.join(d, 'u.json'); awp.save_json(upd, up); out = os.path.join(d, 'out.xlsx')
        import io, contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf): good = run(src, out, [up])
        a, b = load_workbook(src), load_workbook(out); s2 = awp.export_spec(out); w = s2['walkthroughs'][0]
        same = all(a[n].cell(r, c).value == b[n].cell(r, c).value for n in ('Scope', 'A', 'A200', 'A300', 'A301', 'A400') for r in range(1, a[n].max_row + 1) for c in range(1, a[n].max_column + 1))
        checks = [good, same, a.sheetnames == b.sheetnames, w['questions'][0]['response'].startswith('Staff note: the Manager approves in the system.') and 'Minutes:' in w['questions'][0]['response'],
                  [x.get('control_ref') for x in w['sections'][0]['activities']] == ['A201-1', 'A201-3', 'A201-2'], [x['no'] for x in w['sections'][0]['activities']] == [1, 2, 3], len(w['questions']) == 6,
                  s2['rcm']['sections'][0]['rows'][0]['de_finding'].startswith('No design gap'), str(b['A100']['G9'].value).startswith('='), os.path.exists(os.path.splitext(out)[0] + ' - update log.xlsx')]
        bad = dict(upd, changes=[{'sheet': 'A201', 'op': 'question', 'no': 1, 'set': {'response': 'overwrite'}, 'source': ['M1'], 'basis': 'x'}]); awp.save_json(bad, up)
        with contextlib.redirect_stdout(buf): checks.append(run(src, os.path.join(d, 'o2.xlsx'), [up]) is False)       # a staff response is never overwritten silently
        ok = all(checks); msg = '' if ok else f' - failed checks {[i for i, c in enumerate(checks) if not c]}\n' + buf.getvalue()
    finally: shutil.rmtree(d, ignore_errors=True)
    print('SELFTEST', 'PASSED' if ok else 'FAILED', f'- awp-update v{VERSION} · engine awp-generate v{awp.VERSION}' + msg); return ok

# =============================================================================== CLI
if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if x != '--force']; force = '--force' in sys.argv; cmd = a[0] if a else ''
    if cmd == 'selftest': sys.exit(0 if selftest() else 1)
    elif cmd == 'show' and len(a) >= 2: show(a[1], a[2:])
    elif cmd == 'check' and len(a) >= 3: sys.exit(0 if run(a[1], None, a[2:], write=False) else 1)
    elif cmd == 'apply' and len(a) >= 4: sys.exit(0 if run(a[1], a[2], a[3:], force=force) else 1)
    else: print(__doc__)
