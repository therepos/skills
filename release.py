"""
release.py - build a self-contained release zip for one skill in this repository.

    python release.py awp-generate 2.1          -> dist/skill-awp-generate-2.1.zip
    python release.py awp-learn 1.0             -> dist/skill-awp-learn-1.0.zip
    python release.py zipper 1.3                -> dist/skill-zipper-1.3.zip

What goes into the zip: the skill folder as it is in the repository, plus whatever it needs to run on its own
(users only ever have the zip, never the repository):
  - any skill whose name starts with "awp-"  gets a copy of awp-lib/ as <skill>/lib/
  - awp-learn                                  gets a copy of awp-generate/scripts/awp.py (its engine)
The zip unpacks to a single folder named after the skill; SKILL.md is its entry point.
Also writes an extension-less 7z copy named YYYYMMDD<letter> when py7zr is installed (for chats that strip extensions).
"""
import os, sys, shutil, zipfile, datetime as dt

ROOT = os.path.dirname(os.path.abspath(__file__))

def build(skill, version):
    src = os.path.join(ROOT, skill)
    if not os.path.isdir(src) or not os.path.exists(os.path.join(src, 'SKILL.md')): raise SystemExit(f'{skill}: no such skill folder with a SKILL.md')
    stage = os.path.join(ROOT, 'dist', '_stage'); shutil.rmtree(stage, ignore_errors=True)
    dst = os.path.join(stage, skill); shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '.DS_Store'))
    if skill.startswith('awp-'):
        shutil.copytree(os.path.join(ROOT, 'awp-lib'), os.path.join(dst, 'lib'), ignore=shutil.ignore_patterns('README.md'))
    if skill == 'awp-learn':
        shutil.copy(os.path.join(ROOT, 'awp-generate', 'scripts', 'awp.py'), os.path.join(dst, 'scripts', 'awp.py'))
    os.makedirs(os.path.join(ROOT, 'dist'), exist_ok=True)
    out = os.path.join(ROOT, 'dist', f'skill-{skill}-{version}.zip')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for folder, _, files in os.walk(dst):
            for f in sorted(files):
                p = os.path.join(folder, f); z.write(p, os.path.relpath(p, stage))
    seven = None
    try:
        import py7zr
        seven = os.path.join(ROOT, 'dist', dt.date.today().strftime('%Y%m%d') + {'awp-generate': 'a', 'awp-learn': 'b'}.get(skill, 'z'))
        with py7zr.SevenZipFile(seven, 'w') as z: z.writeall(dst, skill)
    except ImportError: pass
    shutil.rmtree(stage, ignore_errors=True)
    n = len(zipfile.ZipFile(out).namelist())
    print(f'{out}  ({n} files)' + (f'\n{seven}  (same content, 7z, no extension)' if seven else ''))
    return out

if __name__ == '__main__':
    if len(sys.argv) != 3: print(__doc__); sys.exit(1)
    build(sys.argv[1], sys.argv[2])
