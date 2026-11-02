"""
learn.py  v1.0 - harvest a finished AWP into library proposals and apply the reviewed decisions (skill awp-learn).

Run from the folder that contains SKILL.md:

    python scripts/learn.py learn   AWP.xlsx PROPOSALS.json [--tokens T.json]   anonymised proposals vs the libraries (+ PROPOSALS.md)
    python scripts/learn.py apply   PROPOSALS.json DECISIONS.json              merge accepted proposals into the libraries (version bump, changelog)
    python scripts/learn.py selftest                                          proves the environment and prints the versions

    --lib FOLDER points at another library folder. Libraries are found in lib/ next to this skill (release zip) or in ../awp-lib/ (repository).

Shares its engine with awp-generate: scripts/awp.py next to this file (release zip) or ../../awp-generate/scripts/awp.py (repository).
"""
import json, os, re, sys, difflib, datetime as dt, importlib.util

VERSION = '1.0'
HERE = os.path.dirname(os.path.abspath(__file__)); SKILL = os.path.dirname(HERE)
try: sys.stdout.reconfigure(errors='replace')
except Exception: pass
_awp_path = next((p for p in (os.path.join(HERE, 'awp.py'), os.path.join(os.path.dirname(SKILL), 'awp-generate', 'scripts', 'awp.py')) if os.path.exists(p)), None)
if not _awp_path: raise SystemExit('awp.py not found: keep it next to learn.py (release zip) or keep awp-generate beside awp-learn (repository)')
_spec = importlib.util.spec_from_file_location('awp', _awp_path); awp = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(awp)
from_awp = ('find_libraries', 'lib_dir', 'load_json', 'save_json', 'subst', 'validate', 'export_spec', 'TOKENS')
globals().update({k: getattr(awp, k) for k in from_awp})
LIB_DIRS = awp.LIB_DIRS

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
    props = load_json(proposals_path); dec = load_json(decisions_path); folder = folder or lib_dir()
    core_p, pack_ps = find_libraries(folder); libs = {'core': load_json(core_p)} if core_p else {}
    for p in pack_ps: libs['pack:' + os.path.basename(p)[len('pack-'):-5]] = load_json(p)
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
            if not (d.get('module_title') and d.get('sheet') and d.get('stage')): raise SystemExit(f"{p['id']}: module {d['module']} does not exist - check the id (python scripts/awp.py list (in awp-generate)), or create it by giving module_title, sheet and stage")
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
        path = core_p if key == 'core' else next(p for p in pack_ps if os.path.basename(p) == f"pack-{key.split(':', 1)[1]}.json")
        save_json(lib, path); print('updated in place:', path, '-> version', lib['version'], '(the previous version is in the release zip / the repository)')
    print('\n'.join(log) or 'nothing accepted')


def selftest(folder=None):
    core_p, packs = find_libraries(folder); msg = [f'awp-learn v{VERSION} (engine awp-generate v{awp.VERSION})']
    for p in ([core_p] if core_p else []) + packs:
        lib = load_json(p); msg.append(f"{'core' if lib.get('library') == 'core' else 'pack ' + str(lib.get('process'))} v{lib.get('version')}")
    ok = bool(core_p) and os.path.exists(os.path.join(SKILL, 'ref', 'learn.md'))
    print('SELFTEST', 'PASSED' if ok else 'FAILED', '-', ' · '.join(msg) + ('' if ok else ' - core.json or ref/learn.md missing')); return ok

if __name__ == '__main__':
    a = sys.argv[1:]; cmd = a[0] if a else ''
    def opt(name, default=None): return a[a.index(name) + 1] if name in a else default
    if cmd == 'learn':
        tk = load_json(opt('--tokens')) if opt('--tokens') else {}; out = learn(a[1], a[2], tk, opt('--lib')); print('proposals', a[2], out['summary'], f"| already learned: {out['already_learned']} | export notes:", out['export_report'] or 'none')
    elif cmd == 'apply': apply_proposals(a[1], a[2], opt('--lib'))
    elif cmd == 'selftest': sys.exit(0 if selftest(opt('--lib')) else 1)
    else: print(__doc__)
