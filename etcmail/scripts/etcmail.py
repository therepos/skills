"""
etcmail.py - build the "[For Approval] ETC Adjustment" email as an Outlook-ready draft (.eml).

    python scripts/etcmail.py selftest                      check the engine (no files left behind)
    python scripts/etcmail.py files [DIR ...]                list the image files the sandbox really has (size, shape, fingerprint)
    python scripts/etcmail.py build <input.json> [--out DIR] [--settings FILE]
                                                            write the draft and print the summary for the chat

The model reads the ETC Projection screenshots and writes <input.json> (see SKILL.md). The user's own details
(partner, CC) are not part of the skill: they come from the "settings" block in the input, which the
model copies from the agent's instructions (or from --settings FILE). Everything about the format - subject,
greeting, table, number formats and colours, screenshots - is here. No signature: the user adds
their own in Outlook (Insert > Signature).
Standard library only (Python 3.6+); cutting a full-page screenshot into the two house crops also uses Pillow and
numpy when the sandbox has them (Copilot's does) - without them the full screenshot is embedded instead. The .eml carries "X-Unsent: 1", so Outlook opens it as a new draft.
"""
import sys, os, re, json, html, struct, zlib, tempfile, shutil, mimetypes, hashlib, time
from email.message import EmailMessage
from email.utils import make_msgid, getaddresses, formataddr
from email import policy
from email.parser import BytesParser

VERSION = '1.4'
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
GREEN, RED = '#00B050', 'red'
MAX_IMG_W = 1147  # display width of each screenshot in the email (as in the house emails)
FONT = 'font-size:11.0pt;font-family:"Arial",sans-serif'  # Arial 11 pt on every run of text, so Outlook never falls back to its default size


# ------------------------------------------------------------------ settings
def load_settings(inp, path=None, base='.'):
    """The user's details: --settings FILE first, then the input's "settings" block (copied from the agent's
    instructions). Nothing personal ships with the skill."""
    s, src = {}, None
    if isinstance(inp.get('settings'), dict) and inp['settings']: s, src = dict(inp['settings']), 'input'
    if path:
        with open(path, encoding='utf-8') as fh: s, src = json.load(fh), path
    s['_source'] = src
    s['_base'] = os.path.dirname(os.path.abspath(path)) if path else base
    return s


# ------------------------------------------------------------------ formatting rules
def fmt_amount(v):
    """-499.33 -> '(499.33)', 1234.5 -> '1,234.50', 0 -> '0.00'."""
    v = round(float(v), 2)
    if v == 0: return '0.00'
    s = '{:,.2f}'.format(abs(v))
    return '(%s)' % s if v < 0 else s

def amount_colour(v):
    v = round(float(v), 2)
    return RED if v < 0 else GREEN if v > 0 else None

def fmt_margin(v):
    return '{:.1f}%'.format(float(v))

def norm_code(c):
    """'E- 12345678', 'e12345678', '12345678' -> 'E-12345678'."""
    d = re.sub(r'\D', '', str(c or ''))
    return 'E-' + d if d else ''

LEGAL = r'\b(PTE\.?|PRIVATE|LTD\.?|LIMITED|CORPORATION|CORP\.?|INC\.?|HOLDINGS|GROUP|COMPANY|CO\.?|BERHAD|BHD\.?|PLC|LLC|LLP)\b'
STOP = {'OF', 'FOR', 'AND', 'THE', '&'}

def client_short(name, overrides=None):
    """Short client name for the subject: map in settings first; 'MINISTRY OF HEALTH' -> 'MOH';
    'ACME CORPORATION LIMITED' -> 'Acme'; otherwise the name without legal suffixes, in title case."""
    name = (name or '').strip()
    for k, v in (overrides or {}).items():
        if k.strip().upper() == name.upper(): return v
    up = name.upper()
    if up.startswith('MINISTRY OF '):  # Singapore ministries keep the "O": MOH, MOE, MOF, MOM
        return 'MO' + ''.join(w[0] for w in re.findall(r"[A-Z]+", up[12:]) if w not in STOP)
    if up.startswith(('MINISTRY ', 'AGENCY ', 'AUTHORITY ', 'BOARD ', 'COUNCIL ')) or ' AUTHORITY' in up or ' BOARD' in up:
        return ''.join(w[0] for w in re.findall(r"[A-Z]+", up) if w not in STOP)
    base = re.sub(r'\s+', ' ', re.sub(LEGAL, ' ', up)).strip(' ,.-')
    words = base.split()
    if not words: return name
    return ' '.join(w.capitalize() if len(w) > 3 else w for w in words[:2]) if len(words) > 1 and len(words[0]) <= 3 else words[0].capitalize()

def first_name(full):
    return (full or '').strip().split(' ')[0] if full else ''

def join_names(names):
    names = [n for n in names if n]
    return names[0] if len(names) == 1 else ', '.join(names[:-1]) + ' and ' + names[-1] if names else ''


# ------------------------------------------------------------------ images (no third-party libraries)
def sniff(data):
    if data[:8] == b'\x89PNG\r\n\x1a\n': return 'png'
    if data[:3] == b'\xff\xd8\xff': return 'jpeg'
    if data[:6] in (b'GIF87a', b'GIF89a'): return 'gif'
    if data[:4] == b'RIFF' and data[8:12] == b'WEBP': return 'webp'
    if data[:2] == b'BM': return 'bmp'
    return None

def dims(data):
    k = sniff(data)
    try:
        if k == 'png': return struct.unpack('>II', data[16:24])
        if k == 'gif': return struct.unpack('<HH', data[6:10])
        if k == 'bmp': w, h = struct.unpack('<ii', data[18:26]); return w, abs(h)
        if k == 'jpeg':
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF: i += 1; continue
                m = data[i + 1]
                if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h, w = struct.unpack('>HH', data[i + 5:i + 9]); return w, h
                if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7: i += 2; continue
                i += 2 + struct.unpack('>H', data[i + 2:i + 4])[0]
    except Exception: pass
    return None

def _bands(rows, gap):
    out, start, last = [], None, None
    for y, v in enumerate(rows):
        if v:
            if start is None: start = y
            elif y - last > gap: out.append((start, last + 1)); start = y
            last = y
    if start is not None: out.append((start, last + 1))
    return out

def split_full(data):
    """A full-page ETC Projection screenshot -> [header crop, Change in Revenue Recognition strip] as PNG bytes,
    cut the way the house emails are: full width; the header from the top down to just above "Expand to view
    metrics" (title, client line, the four tiles with their vs Budget line); the strip a band around the red-bordered
    revenue-recognition box. Returns (crops, None) or (None, reason)."""
    try:
        import io
        import numpy as np
        from PIL import Image
    except ImportError:
        return None, 'Pillow/numpy not available'
    import io
    im = Image.open(io.BytesIO(data)).convert('RGB')
    a = np.asarray(im).astype(int)
    h, w, _ = a.shape
    sc = w / 1912.0  # the house screenshots are 1912 px wide
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    red = (r - g > 28) & (r - b > 20) & (r > 120)
    need = int(0.06 * w)

    def run(v, n):
        c = 0
        for x in v:
            c = c + 1 if x else 0
            if c >= n: return True
        return False
    rows = [y for y in np.nonzero(red.sum(axis=1) >= need)[0] if run(red[y], need)]
    if not rows: return None, 'the red-bordered Change in Revenue Recognition box was not found'
    top, bot = int(rows[0]), int(rows[-1])
    if not 20 * sc < bot - top < 200 * sc: return None, 'the Change in Revenue Recognition box has an unexpected size'
    ink = (a[:top, int(0.22 * w):].mean(axis=2) < 170).sum(axis=1) > 0  # text rows, right of the Options column
    bs = _bands(ink, int(round(3 * sc)))
    # bands: title, client line, tile labels, tile numbers (largest font), vs Budget, [Expand to view metrics ...]
    if len(bs) < 5: return None, 'the four metric tiles were not found above the box'
    hgt = [e - s_ for s_, e in bs]
    if not (hgt[3] > hgt[2] and hgt[3] > hgt[4]): return None, 'the four metric tiles were not where expected'
    end = bs[4][1]
    cut = bs[5][0] - int(round(4 * sc)) if len(bs) > 5 and bs[5][0] - end < 60 * sc else end + int(round(20 * sc))
    cut = max(end + 1, min(cut, top))
    pad = max(2, int(round((bot - top) * 0.07)))
    out = []
    for box in ((0, 0, w, cut), (0, max(cut, top - pad), w, min(h, bot + pad + 1))):
        buf = io.BytesIO(); im.crop(box).save(buf, 'PNG', optimize=True); out.append(buf.getvalue())
    return out, None

def png_bytes(w, h, rgb=(240, 240, 240)):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')


# ------------------------------------------------------------------ the email
def validate(inp, base):
    errs, warns, engs, seen = [], [], [], {}
    rows = inp.get('engagements') or []
    if not rows: errs.append('no engagements in the input')
    for n, e in enumerate(rows, 1):
        tag = 'engagement %d' % n
        name = (e.get('engagement') or '').strip()
        name = re.sub(r'^\s*ETC Projection\s*\|\s*', '', name)
        if not name: errs.append(tag + ': "engagement" (the title after "ETC Projection |") is empty')
        code = norm_code(e.get('code'))
        if not re.fullmatch(r'E-\d{8}', code or ''): (warns if code else errs).append('%s: engagement code %r is not E- plus 8 digits; check it' % (tag, e.get('code')))
        try: rc = float(e.get('revenue_change'))
        except (TypeError, ValueError): errs.append(tag + ': "revenue_change" must be a number (negative when the arrow is red/down)'); rc = 0.0
        try: mg = float(e.get('margin'))
        except (TypeError, ValueError): errs.append(tag + ': "margin" must be a number (Projected Margin, in %)'); mg = 0.0
        if not -100 <= mg <= 100: warns.append('%s: margin %s%% looks wrong' % (tag, mg))
        if abs(rc) >= 1e6: warns.append('%s: change in revenue recognition %s looks too large; check it' % (tag, rc))
        shots, missing = [], []
        for p in e.get('screenshots') or []:
            q = p if os.path.isabs(p) else os.path.join(base, p)
            if not os.path.exists(q): missing.append(p); continue
            with open(q, 'rb') as fh: data = fh.read()
            k = sniff(data)
            if not k: missing.append(p + ' (not an image)'); continue
            fp = hashlib.sha1(data).hexdigest()
            if fp in seen:  # same picture twice: pasted images often share one name and overwrite each other
                warns.append('%s: %s is the same picture as %s - only one copy kept; a screenshot is missing (pasted images can overwrite each other in the sandbox)' % (tag, p, seen[fp])); continue
            seen[fp] = p
            size = dims(data)
            if size and size[1] > 0.45 * size[0]:  # a full-page screenshot: cut it into the two house crops
                crops, why = split_full(data)
                if crops:
                    for cdata in crops: shots.append({'path': q, 'data': cdata, 'kind': 'png', 'size': dims(cdata)})
                    continue
                warns.append('%s: %s could not be cut into header and revenue-recognition crops (%s) - the full screenshot is embedded' % (tag, p, why))
            shots.append({'path': q, 'data': data, 'kind': k, 'size': size})
        if missing: warns.append('%s: screenshot(s) not found or unreadable: %s' % (tag, ', '.join(missing)))
        if not shots: warns.append('%s: no screenshots - the Supporting Documents cell says to paste them' % tag)
        elif all(sh['size'] and sh['size'][0] > 12 * sh['size'][1] for sh in shots):
            warns.append('%s: only the Change in Revenue Recognition strip is embedded - the ETC Projection header screenshot is missing' % tag)
        # header (tall) before the Change in Revenue Recognition box (flat), whatever order they came in
        shots.sort(key=lambda s: (s['size'][0] / float(s['size'][1])) if s['size'] and s['size'][1] else 0)
        engs.append({'name': name, 'code': code, 'client': (e.get('client') or '').strip(), 'short': (e.get('client_short') or '').strip(), 'rc': rc, 'margin': mg,
                     'partner': (e.get('partner') or '').strip(), 'note': (e.get('partner_note') or '').strip(), 'shots': shots})
    return engs, errs, warns

def compose(inp, settings, base='.'):
    engs, errs, warns = validate(inp, base)
    if errs: return None, errs, warns
    partners = settings.get('partners') or {}
    default = settings.get('default_partner') or ''
    for e in engs: e['partner'] = e['partner'] or default
    # recipients: the partner(s) in charge, in order of first appearance
    to, seen, greet = [], set(), []
    for e in engs:
        p = e['partner']
        if not p or p in seen: continue
        seen.add(p)
        match = next((k for k in partners if k.lower() == p.lower() or first_name(k).lower() == p.lower()), None)
        if match:
            to.append(formataddr((match, partners[match]))); greet.append(first_name(match))
            if p != match: e['partner'] = match
        else:
            warns.append('partner %r has no email address in settings - add it in Outlook before sending' % p); greet.append(first_name(p))
    if not engs[0]['partner']: warns.append('no partner given and no default_partner in settings')
    cc = list(settings.get('cc') or [])
    # subject: one engagement -> its name; several -> the clients' short names
    if inp.get('subject'): subject = inp['subject']
    else:
        prefix = settings.get('subject_prefix', '[For Approval] ETC Adjustment - ')
        if len(engs) == 1: tail = engs[0]['name']
        else:
            shorts = []
            for e in engs:
                s = e['short'] or (client_short(e['client'], settings.get('client_short')) if e['client'] else e['name'])
                if s not in shorts: shorts.append(s)
            tail = ', '.join(shorts)
        subject = prefix + tail
    return {'subject': subject, 'to': to, 'cc': cc, 'greeting': 'Hi %s,' % join_names(greet) if greet else 'Hi,',
            'engs': engs, 'settings': settings}, errs, warns

def td(content, width, head=False, first=False, center=True, pad=True):
    b = 'border:solid windowtext 1.0pt;' + ('' if first else 'border-left:none;') + ('' if head else 'border-top:none;')
    bg = 'background:#D9D9D9;' if head else ''
    p = 'padding:2.85pt 2.85pt 2.85pt 2.85pt' if pad else 'padding:0cm 0cm 0cm 0cm'
    al = " align=center style='text-align:center;margin:0'" if center else " style='margin:0'"
    paras = content if isinstance(content, list) else [content]
    inner = ''.join('<p class=MsoNormal%s>%s</p>' % (al, c) for c in paras)
    return "<td width=%d valign=top style='width:%dpx;%s%s%s'>%s</td>" % (width, width, b, bg, p, inner)

def span(text, bold=False, colour=None):
    st = FONT + (';color:%s' % colour if colour else '')
    s = "<span style='%s'>%s</span>" % (st, text)
    return '<b>%s</b>' % s if bold else s

def render(c):
    """Return (html, text, images) - images = [(cid, data, kind)]."""
    s = c['settings']
    images, rows, text_rows = [], [], []
    hdr = ['Engagement', 'Change in revenue recognition', 'Margin after adjustment', ['Partner', 'In-Charge'], 'Supporting Documents']
    widths = [274, 113, 92, 92, MAX_IMG_W + 1]
    head = ''.join(td([span(x, True, 'black') for x in (h if isinstance(h, list) else [h])], widths[i], head=True, first=(i == 0),
                      center=(i != 0), pad=(i < 3)) for i, h in enumerate(hdr))
    rows.append('<tr>%s</tr>' % head)
    for e in c['engs']:
        eng = html.escape(e['name']) + (' (%s)' % e['code'] if e['code'] else '')
        amt = span(fmt_amount(e['rc']), True, amount_colour(e['rc']))
        partner = [span(html.escape(e['partner'] or '[Partner]'))] + (['&nbsp;', span(html.escape(e['note']))] if e['note'] else [])
        pics = []
        for sh in e['shots']:
            cid = make_msgid(domain='etcmail')[1:-1]
            images.append((cid, sh['data'], sh['kind']))
            w, h = sh['size'] or (MAX_IMG_W, 0)
            dw = min(w, MAX_IMG_W); dh = int(round(h * dw / float(w))) if h else 0
            pics.append('<img width=%d%s style="width:%dpx%s" src="cid:%s" alt="ETC screenshot">' % (dw, ' height=%d' % dh if dh else '', dw, ';height:%dpx' % dh if dh else '', cid))
        if not pics: pics = [span('[paste the ETC Projection and Change in Revenue Recognition screenshots here]', colour='#7F7F7F')]
        cells = [td(span(eng), widths[0], first=True, center=False), td(amt, widths[1]), td(span(fmt_margin(e['margin'])), widths[2]),
                 td(partner, widths[3], pad=False), td(pics, widths[4], pad=False)]
        rows.append('<tr>%s</tr>' % ''.join(cells))
        text_rows.append('%s | %s | %s | %s%s' % (eng, fmt_amount(e['rc']), fmt_margin(e['margin']), e['partner'], (' ' + e['note']) if e['note'] else ''))
    table = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>%s</table>" % ''.join(rows)

    def p(x): return '<p class=MsoNormal style="margin:0">%s</p>' % x
    blank = p(span('&nbsp;'))
    intro = s.get('intro', 'Please find the proposed ETC adjustments for your approval:')
    closing = s.get('closing', 'Best regards,')
    body = ''.join([p(span(html.escape(c['greeting']))), blank, p(span(html.escape(intro))), blank, table, blank,
                    p(span(html.escape(closing)))])
    doc = ('<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style>p.MsoNormal{margin:0;font-size:11.0pt;%s}</style></head>'
           '<body lang=EN-GB style="word-wrap:break-word"><div class=WordSection1>%s</div></body></html>') % (FONT, body)
    text = '\n'.join([c['greeting'], '', intro, '', 'Engagement | Change in revenue recognition | Margin after adjustment | Partner In-Charge'] + text_rows +
                     ['', '(Supporting documents: see the screenshots in the HTML version.)', '', closing])
    return doc, text, images

def to_eml(c):
    doc, text, images = render(c)
    m = EmailMessage(policy=policy.SMTP)
    m['Subject'] = c['subject']
    if c['to']: m['To'] = ', '.join(c['to'])
    if c['cc']: m['Cc'] = ', '.join(c['cc'])
    m['X-Unsent'] = '1'  # Outlook opens the file as a new draft (editable, Send button)
    m.set_content(text)
    m.add_alternative(doc, subtype='html')
    htmlpart = m.get_payload()[1]
    for cid, data, kind in images:
        htmlpart.add_related(data, maintype='image', subtype=kind, cid='<%s>' % cid, disposition='inline',
                             filename='%s.%s' % (cid.split('@')[0][-8:], 'jpg' if kind == 'jpeg' else kind))
    return m.as_bytes(), doc

def safe_name(s):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', s).strip().rstrip('.')[:150] or 'ETC Adjustment'

IMG_EXT = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')

def list_files(dirs):
    """Every image file under the given folders (default: the usual upload folders and the current one), newest
    first, with what the build needs to tell them apart: shape, size and a fingerprint."""
    dirs = dirs or [d for d in ('/mnt/data', '/mnt/user-data/uploads', '/mnt/uploads', '/tmp/inputs', os.getcwd()) if os.path.isdir(d)]
    rows, fps = [], {}
    for d in dirs:
        for root, sub, files in os.walk(d):
            sub[:] = [x for x in sub if not x.startswith('.') and x not in ('scripts', 'ref', '__pycache__')]
            for f in files:
                q = os.path.join(root, f)
                try:
                    with open(q, 'rb') as fh: data = fh.read()
                except OSError: continue
                k = sniff(data)
                if not k: continue
                fp = hashlib.sha1(data).hexdigest()[:10]
                w, h = dims(data) or (0, 0)
                kind = 'revenue-recognition strip' if h and w > 12 * h else 'full page (cut automatically)' if h > 0.45 * w else 'header' if h else '?'
                rows.append((os.path.getmtime(q), q, w, h, len(data), fp, kind, fps.get(fp)))
                fps.setdefault(fp, q)
    rows.sort(reverse=True)
    for mt, q, w, h, n, fp, kind, dup in rows:
        print('%s  %5dx%-4d %8d B  %s  %-26s %s%s' % (time.strftime('%H:%M:%S', time.localtime(mt)), w, h, n, fp, kind, q, ('  = same picture as ' + dup) if dup else ''))
    print('IMAGES: %d file(s), %d distinct picture(s)' % (len(rows), len(fps)))

def summary(c, path, warns):
    out = ['DRAFT: %s' % path, 'Subject: %s' % c['subject'], 'To: %s' % (', '.join(c['to']) or '(none - add in Outlook)'), 'Cc: %s' % (', '.join(c['cc']) or '-'), '',
           '| Engagement | Change in revenue recognition | Margin after adjustment | Partner In-Charge | Screenshots |', '|---|---:|---:|---|---:|']
    for e in c['engs']:
        out.append('| %s (%s) | %s | %s | %s%s | %d |' % (e['name'], e['code'], fmt_amount(e['rc']), fmt_margin(e['margin']), e['partner'], (' ' + e['note']) if e['note'] else '', len(e['shots'])))
    if not c['settings'].get('_source'): warns = ['no "settings" in the input: To and Cc are empty - put your settings block in the agent\'s Instructions (ref/agent-builder.md)'] + warns
    out += ['', 'SCREENSHOTS EMBEDDED: %d' % sum(len(e['shots']) for e in c['engs'])] + ['WARN: ' + w for w in warns] + ['BUILD OK - etcmail v%s' % VERSION]
    return '\n'.join(out)

def build(inp_path, out_dir=None, settings_path=None):
    with open(inp_path, encoding='utf-8') as fh: inp = json.load(fh)
    base = os.path.dirname(os.path.abspath(inp_path))
    s = load_settings(inp, settings_path, base)
    c, errs, warns = compose(inp, s, base)
    if errs:
        print('\n'.join('ERROR: ' + e for e in errs + []) + ('\n' + '\n'.join('WARN: ' + w for w in warns) if warns else '')); sys.exit(2)
    out_dir = out_dir or os.path.dirname(os.path.abspath(inp_path))
    os.makedirs(out_dir, exist_ok=True)
    raw, doc = to_eml(c)
    path = os.path.join(out_dir, safe_name(c['subject']) + '.eml')
    with open(path, 'wb') as fh: fh.write(raw)
    print(summary(c, path, warns))
    return path


# ------------------------------------------------------------------ checks
def selftest():
    t = tempfile.mkdtemp(prefix='etcmail-')
    try:
        assert fmt_amount(-499.33) == '(499.33)' and fmt_amount(122.04) == '122.04' and fmt_amount(0) == '0.00' and fmt_amount(-0.001) == '0.00'
        assert fmt_amount(1234.5) == '1,234.50' and fmt_amount(-12345.678) == '(12,345.68)'
        assert amount_colour(-1) == RED and amount_colour(1) == GREEN and amount_colour(0) is None
        assert fmt_margin(32.6) == '32.6%' and fmt_margin(30) == '30.0%'
        assert norm_code('E- 12345678') == 'E-12345678' and norm_code('12345678') == 'E-12345678'
        assert client_short('MINISTRY OF HEALTH') == 'MOH' and client_short('MINISTRY OF EDUCATION') == 'MOE'
        assert client_short('MONETARY AUTHORITY OF SINGAPORE') == 'MAS' and client_short('HOUSING & DEVELOPMENT BOARD') == 'HDB'
        assert client_short('ACME CORPORATION LIMITED') == 'Acme' and client_short('GLOBEX PTE. LTD.') == 'Globex'
        assert client_short('ACME CORPORATION LIMITED', {'acme corporation limited': 'ACME Group'}) == 'ACME Group'
        assert join_names(['A']) == 'A' and join_names(['A', 'B']) == 'A and B' and join_names(['A', 'B', 'C']) == 'A, B and C'
        for nm, (w, h) in (('h1.png', (1910, 240)), ('r1.png', (1910, 80)), ('h2.png', (1910, 233)), ('r2.png', (1910, 77))):
            with open(os.path.join(t, nm), 'wb') as fh: fh.write(png_bytes(w, h))
        assert dims(png_bytes(7, 3)) == (7, 3)
        st = {'sender': {'name': 'Test Sender'}, 'partners': {'Pat Partner': 'pat.partner@example.com', 'Lee Lead': 'lee.lead@example.com'}, 'default_partner': 'Pat Partner',
              'cc': ['Cee Copy <cee@example.com>'], '_source': 'input', '_base': t}
        inp = {'engagements': [
            {'engagement': 'ETC Projection | Internal Audit FY26', 'code': 'E- 12345678', 'client': 'ACME CORPORATION LIMITED', 'revenue_change': -499.33,
             'margin': 30, 'partner_note': '(Previously Lee Lead)', 'screenshots': ['r1.png', 'h1.png']},
            {'engagement': 'Governance Review FY25-27', 'code': 'E-87654321', 'client': 'MINISTRY OF HEALTH', 'revenue_change': 122.04, 'margin': 26.0,
             'screenshots': ['h2.png', 'r2.png']}]}
        with open(os.path.join(t, 'in.json'), 'w') as fh: json.dump(inp, fh)
        c, errs, warns = compose(inp, st, t)
        assert not errs and not warns, (errs, warns)
        assert c['subject'] == '[For Approval] ETC Adjustment - Acme, MOH', c['subject']
        assert c['greeting'] == 'Hi Pat,' and c['to'] == ['Pat Partner <pat.partner@example.com>']
        assert c['engs'][0]['shots'][0]['size'][1] == 240, 'header screenshot must come first'
        raw, doc = to_eml(c)
        m = BytesParser(policy=policy.default).parsebytes(raw)
        assert m['X-Unsent'] == '1' and m['Subject'] == c['subject'] and 'cee@example.com' in m['Cc']
        h = m.get_body(('html',)).get_content()
        cids = re.findall(r'src="cid:([^"]+)"', h)
        parts = {p['Content-ID'].strip('<>') for p in m.walk() if p['Content-ID']}
        assert len(cids) == 4 and set(cids) <= parts, 'every screenshot must be embedded'
        for needle in ('Hi Pat,', '(499.33)', 'color:red', '122.04', 'color:#00B050', '30.0%', '26.0%', 'Internal Audit FY26 (E-12345678)',
                       '(Previously Lee Lead)', 'Pat Partner', 'Supporting Documents', 'In-Charge', 'Best regards,'):
            assert needle in h, needle
        assert h.count('font-size:11.0pt') >= 20 and 'font-size:9' not in h and 'font-size:12' not in h, 'all text Arial 11 pt'
        assert 'Test Sender' not in h, 'no signature: users add their own in Outlook'
        dup = {'engagements': [dict(inp['engagements'][0], screenshots=['h1.png', 'h1.png']), dict(inp['engagements'][1], screenshots=['r2.png'])]}
        cd, ed, wd = compose(dup, st, t)
        assert not ed and len(cd['engs'][0]['shots']) == 1 and any('same picture' in w for w in wd) and any('header screenshot is missing' in w for w in wd), wd
        try:  # full-page screenshot -> the two house crops (only where Pillow and numpy exist, as in Copilot)
            import io, numpy as np
            from PIL import Image
            pg = np.full((1300, 1912, 3), 248, dtype=np.uint8)
            for y0, y1 in ((16, 54), (65, 78), (156, 173), (195, 225), (244, 258), (298, 315), (380, 395)): pg[y0:y1, 600:1500] = 40
            pg[900:903, 1420:1668] = (200, 40, 40); pg[968:971, 1420:1668] = (200, 40, 40); pg[900:971, 1420:1423] = (200, 40, 40); pg[900:971, 1665:1668] = (200, 40, 40)
            buf = io.BytesIO(); Image.fromarray(pg).save(buf, 'JPEG', quality=90)
            crops, why = split_full(buf.getvalue())
            assert crops and len(crops) == 2, why
            (hw, hh), (sw, sh_) = dims(crops[0]), dims(crops[1])
            assert hw == sw == 1912 and 258 < hh < 298 and 70 < sh_ < 90, (hh, sh_)
            with open(os.path.join(t, 'full.jpg'), 'wb') as fh: fh.write(buf.getvalue())
            cf, ef, wf = compose({'engagements': [dict(inp['engagements'][1], screenshots=['full.jpg'])]}, st, t)
            assert not ef and not wf and len(cf['engs'][0]['shots']) == 2, wf
            assert split_full(png_bytes(1912, 1300))[0] is None
        except ImportError:
            pass
        one = {'engagements': [dict(inp['engagements'][1], revenue_change=0, screenshots=[])]}
        c1, e1, w1 = compose(one, st, t)
        assert c1['subject'] == '[For Approval] ETC Adjustment - Governance Review FY25-27' and any('no screenshots' in w for w in w1)
        h1 = to_eml(c1)[1]
        assert '<b><span style=\'font-size:11.0pt;font-family:"Arial",sans-serif\'>0.00</span></b>' in h1 and '[paste the ETC Projection' in h1
        bad = {'engagements': [{'engagement': '', 'code': 'x', 'revenue_change': 'n/a', 'margin': None}]}
        assert len(compose(bad, st, t)[1]) == 4
        stranger = {'engagements': [dict(inp['engagements'][1], partner='Sam Someone')]}
        c2, e2, w2 = compose(stranger, st, t)
        assert c2['to'] == [] and c2['greeting'] == 'Hi Sam,' and any('no email address' in w for w in w2)
        assert load_settings({'settings': {'cc': ['x@example.com']}})['_source'] == 'input' and load_settings({})['_source'] is None
        print('SELFTEST PASSED - etcmail v%s' % VERSION)
    finally:
        shutil.rmtree(t, ignore_errors=True)

if __name__ == '__main__':
    a = sys.argv[1:]
    if not a or a[0] in ('-h', '--help'): print(__doc__); sys.exit(0)
    if a[0] == 'selftest': selftest()
    elif a[0] == 'files': list_files(a[1:])
    elif a[0] == 'build' and len(a) >= 2:
        out = a[a.index('--out') + 1] if '--out' in a else None
        sp = a[a.index('--settings') + 1] if '--settings' in a else None
        build(a[1], out, sp)
    else: print(__doc__); sys.exit(1)
