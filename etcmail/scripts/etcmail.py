"""
etcmail.py - build the "[For Approval] ETC Adjustment" email as an Outlook-ready draft (.eml).

    python scripts/etcmail.py selftest                      check the engine (no files left behind)
    python scripts/etcmail.py settings                      show who the email goes to and the signature in use
    python scripts/etcmail.py build <input.json> [--out DIR]  write the draft and print the summary for the chat

The model reads the ETC Projection screenshots and writes <input.json> (see SKILL.md); everything about the
format - subject, recipients, greeting, table, number formats and colours, screenshots, signature - is here.
Standard library only (Python 3.6+). The .eml carries "X-Unsent: 1", so Outlook opens it as a new draft.
"""
import sys, os, re, json, html, struct, zlib, tempfile, shutil, mimetypes
from email.message import EmailMessage
from email.utils import make_msgid, getaddresses, formataddr
from email import policy
from email.parser import BytesParser

VERSION = '1.0'
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
LIB = os.path.join(SKILL, 'lib')
GREEN, RED = '#00B050', 'red'
MAX_IMG_W = 1147  # display width of each screenshot in the email (as in the house emails)
FONT = 'font-family:"Arial",sans-serif'


# ------------------------------------------------------------------ settings
def load_settings(path=None):
    """lib/settings.json (the user's own, never in the public repository), else the example with placeholders."""
    cands = [path] if path else [os.path.join(LIB, 'settings.json'), os.path.join(LIB, 'settings.example.json')]
    for p in cands:
        if p and os.path.exists(p):
            with open(p, encoding='utf-8') as fh: s = json.load(fh)
            s['_path'] = p
            s['_placeholder'] = os.path.basename(p) == 'settings.example.json'
            return s
    raise SystemExit('ERROR: no lib/settings.json or lib/settings.example.json next to the script')


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

def png_bytes(w, h, rgb=(240, 240, 240)):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')


# ------------------------------------------------------------------ the email
def validate(inp, base):
    errs, warns, engs = [], [], []
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
            shots.append({'path': q, 'data': data, 'kind': k, 'size': dims(data)})
        if missing: warns.append('%s: screenshot(s) not found or unreadable: %s' % (tag, ', '.join(missing)))
        if not shots: warns.append('%s: no screenshots - the Supporting Documents cell says to paste them' % tag)
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
    s = c['settings']; snd = s.get('sender') or {}
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
    sig_lines = []
    nm = html.escape(snd.get('name', ''))
    if nm: sig_lines.append("<p class=MsoNormal style='margin:0'><b><span style='font-size:9.0pt;%s'>%s | </span></b><span style='font-size:9.0pt;%s'>%s</span></p>" % (FONT, nm, FONT, html.escape(snd.get('title', ''))))
    sig_lines.append("<p class=MsoNormal style='margin:0'><span style='font-size:9.0pt;%s'>&nbsp;</span></p>" % FONT)
    for key in ('company', 'address'):
        if snd.get(key): sig_lines.append("<p class=MsoNormal style='margin:0'><span style='font-size:9.0pt;%s'>%s</span></p>" % (FONT, html.escape(snd[key])))
    contact = []
    if snd.get('mobile'): contact.append('Mobile: ' + html.escape(snd['mobile']))
    if snd.get('email'): contact.append('<a href="mailto:%s"><span style="color:#467886">%s</span></a>' % (html.escape(snd['email']), html.escape(snd['email'])))
    if contact: sig_lines.append("<p class=MsoNormal style='margin:0'><span style='font-size:9.0pt;%s'>%s</span></p>" % (FONT, ' | '.join(contact)))
    if snd.get('website'):
        w = html.escape(snd['website'])
        sig_lines.append("<p class=MsoNormal style='margin:0'><span style='font-size:9.0pt;%s'>Website: <a href=\"https://%s\"><span style='color:blue'>%s</span></a></span></p>" % (FONT, w.replace('https://', ''), w))
    logo_cell = ''
    lp = snd.get('logo') and os.path.join(os.path.dirname(s.get('_path', LIB)), snd['logo'])
    if lp and os.path.exists(lp):
        with open(lp, 'rb') as fh: ld = fh.read()
        k = sniff(ld)
        if k:
            cid = make_msgid(domain='etcmail')[1:-1]; images.append((cid, ld, k))
            lw, lh = dims(ld) or (94, 85)
            sw = min(lw, 94); sh_ = int(round(lh * sw / float(lw)))
            logo_cell = "<td valign=top style='padding:0cm 6pt 0cm 0cm'><img width=%d height=%d src=\"cid:%s\" alt=\"logo\"></td>" % (sw, sh_, cid)
    sig = "<table class=MsoNormalTable border=0 cellpadding=0><tr>%s<td valign=top style='padding:0cm'>%s</td></tr></table>" % (logo_cell, ''.join(sig_lines)) if (nm or logo_cell) else ''
    body = ''.join([p(span(html.escape(c['greeting']))), blank, p(span(html.escape(intro))), blank, table, blank,
                    p(span(html.escape(closing))), blank, sig])
    doc = ('<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style>p.MsoNormal{margin:0;font-size:11.0pt;%s}</style></head>'
           '<body lang=EN-GB style="word-wrap:break-word"><div class=WordSection1>%s</div></body></html>') % (FONT, body)
    text = '\n'.join([c['greeting'], '', intro, '', 'Engagement | Change in revenue recognition | Margin after adjustment | Partner In-Charge'] + text_rows +
                     ['', '(Supporting documents: see the screenshots in the HTML version.)', '', closing, '', snd.get('name', ''), snd.get('title', '')])
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

def summary(c, path, warns):
    out = ['DRAFT: %s' % path, 'Subject: %s' % c['subject'], 'To: %s' % (', '.join(c['to']) or '(none - add in Outlook)'), 'Cc: %s' % (', '.join(c['cc']) or '-'), '',
           '| Engagement | Change in revenue recognition | Margin after adjustment | Partner In-Charge | Screenshots |', '|---|---:|---:|---|---:|']
    for e in c['engs']:
        out.append('| %s (%s) | %s | %s | %s%s | %d |' % (e['name'], e['code'], fmt_amount(e['rc']), fmt_margin(e['margin']), e['partner'], (' ' + e['note']) if e['note'] else '', len(e['shots'])))
    if c['settings'].get('_placeholder'): warns = ['settings are the PLACEHOLDER example (lib/settings.example.json): recipients and signature are not real - the owner must add lib/settings.json'] + warns
    out += [''] + ['WARN: ' + w for w in warns] + ['BUILD OK - etcmail v%s' % VERSION]
    return '\n'.join(out)

def build(inp_path, out_dir=None, settings_path=None):
    with open(inp_path, encoding='utf-8') as fh: inp = json.load(fh)
    s = load_settings(settings_path)
    c, errs, warns = compose(inp, s, os.path.dirname(os.path.abspath(inp_path)))
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
        st = {'sender': {'name': 'Test Sender', 'title': 'Risk Consulting', 'company': 'Example Co', 'email': 'sender@example.com'},
              'partners': {'Pat Partner': 'pat.partner@example.com', 'Lee Lead': 'lee.lead@example.com'}, 'default_partner': 'Pat Partner',
              'cc': ['Cee Copy <cee@example.com>'], '_path': os.path.join(t, 'settings.json')}
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
                       '(Previously Lee Lead)', 'Pat Partner', 'Supporting Documents', 'In-Charge', 'Best regards,', 'Test Sender'):
            assert needle in h, needle
        one = {'engagements': [dict(inp['engagements'][1], revenue_change=0, screenshots=[])]}
        c1, e1, w1 = compose(one, st, t)
        assert c1['subject'] == '[For Approval] ETC Adjustment - Governance Review FY25-27' and any('no screenshots' in w for w in w1)
        h1 = to_eml(c1)[1]
        assert '<b><span style=\'font-family:"Arial",sans-serif\'>0.00</span></b>' in h1 and '[paste the ETC Projection' in h1
        bad = {'engagements': [{'engagement': '', 'code': 'x', 'revenue_change': 'n/a', 'margin': None}]}
        assert len(compose(bad, st, t)[1]) == 4
        stranger = {'engagements': [dict(inp['engagements'][1], partner='Sam Someone')]}
        c2, e2, w2 = compose(stranger, st, t)
        assert c2['to'] == [] and c2['greeting'] == 'Hi Sam,' and any('no email address' in w for w in w2)
        s = load_settings()
        print('SELFTEST PASSED - etcmail v%s · settings: %s' % (VERSION, 'PLACEHOLDER example (add lib/settings.json)' if s['_placeholder'] else
              '%s, default partner %s' % (os.path.basename(s['_path']), s.get('default_partner'))))
    finally:
        shutil.rmtree(t, ignore_errors=True)

def show_settings():
    s = load_settings(); snd = s.get('sender') or {}
    print('settings file: %s%s' % (s['_path'], '  (PLACEHOLDER - not the real recipients)' if s['_placeholder'] else ''))
    print('default partner: %s' % s.get('default_partner'))
    for k, v in (s.get('partners') or {}).items(): print('partner: %s <%s>' % (k, v))
    print('cc: %s' % ', '.join(s.get('cc') or []))
    print('signature: %s | %s | %s' % (snd.get('name'), snd.get('title'), snd.get('email')))
    print('logo: %s' % ('yes' if snd.get('logo') and os.path.exists(os.path.join(os.path.dirname(s['_path']), snd['logo'])) else 'none'))

if __name__ == '__main__':
    a = sys.argv[1:]
    if not a or a[0] in ('-h', '--help'): print(__doc__); sys.exit(0)
    if a[0] == 'selftest': selftest()
    elif a[0] == 'settings': show_settings()
    elif a[0] == 'build' and len(a) >= 2:
        out = a[a.index('--out') + 1] if '--out' in a else None
        sp = a[a.index('--settings') + 1] if '--settings' in a else None
        build(a[1], out, sp)
    else: print(__doc__); sys.exit(1)
