"""Split the senior's single-file analysis suite into one static page per module.

The original index.html keeps five self-contained tools as JSON ("tool-src") and mounts them
into iframes with srcdoc. Here each tool becomes static/suite/<id>.html, loading bridge.js
first (shared storage + browser download stand-in). Module code itself is copied verbatim.

    python tools/build_suite.py "<path to original index.html>"
    python tools/build_suite.py --module tcd "<path to one revised module exported on its own>"

A module he revises later arrives as a whole HTML document, not as a tool-src entry. --module
keeps that file in tools/src/<id>.html; from then on it wins over the tool-src copy, so rebuilding
from an older index.html cannot bring the old module back.
"""
import json
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'static', 'suite')
SRC = os.path.join(HERE, 'src')

# Everything else is copied verbatim. The one exception: the reactor tool fetches a sample CSV
# (data/sample_6hr_Red.b64.txt) that lived beside the original artifact and is not part of the
# single-file export, so on open it showed an "HTTP 404" error. Say so plainly instead.
PATCHES = {
    'reactor': [(
        "$('fileinfo').textContent='範例載入失敗，請直接拖入 CSV';loadMsg(String(err&&err.message||err),true)",
        "$('fileinfo').textContent='尚未載入檔案';loadMsg('本站沒有附範例檔，請拖入或選擇記錄器 CSV。')",
    )],
    # ASF classified rings before unsaturation, so cycloalkenes (cyclopentene, cyclohexene, ...)
    # landed in "cyclic". Same rule the FID classifier now uses: hetero -> aromatic -> unsaturated
    # -> ring -> alkane; "cyclic" means fully saturated rings only.
    'asf': [(
        "    else if (st.ring > 0) cls = 'cyclic';\n    else if (st.yne > 0) cls = 'alkyne';\n    else if (st.ene > 0) cls = 'alkene';\n",
        "    else if (st.yne > 0) cls = 'alkyne';\n    else if (st.ene > 0) cls = 'alkene';\n    else if (st.ring > 0) cls = 'cyclic';\n",
    ), (
        # Shimadzu exports name their only sheet "Sheet1"; two such files then collide as 「Sheet1」=「Sheet1」.
        # Label by file name unless the sheet name actually says something.
        "label: t.sheet || file.name.replace(/\\.[^.]+$/, ''),",
        "label: (t.sheet && !/^sheet\\d*$/i.test(t.sheet)) ? t.sheet : file.name.replace(/\\.[^.]+$/, ''),",
    )],
}


# Libraries the modules pull from a CDN. They are served from static/vendor instead (byte-identical,
# hashes in static/vendor/VERSIONS.txt) so the tools do not depend on a third party being reachable.
VENDOR = {
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js': '/static/vendor/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js': '/static/vendor/exceljs.min.js',
    'https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.27.0/plotly.min.js': '/static/vendor/plotly.min.js',
}


def standalone_part(path):
    """A module exported on its own is a whole document; the suite keeps its head and body apart.

    Inside index.html the tools already read and write through window.__LS (the shell's shared
    store, so one backup covers every module). Exported alone they call localStorage directly,
    which would bypass that store: saved runs would not reach the backup or the overview card.
    """
    text = open(path, encoding='utf-8').read()
    m = re.match(r'\s*<!doctype html>\s*<html[^>]*>\s*<head>(.*?)</head>\s*<body>(.*)</body>\s*</html>\s*$',
                 text, re.S | re.I)
    if not m:
        sys.exit(f'{path}: expected one <head> and one <body>')
    body = re.sub(r'\blocalStorage\.(getItem|setItem|removeItem)\(', r'__LS.\1(', m.group(2).lstrip('\n'))
    if 'localStorage' in body:
        sys.exit(f'{path}: localStorage used in a way the __LS rewrite does not cover')
    return {'head': m.group(1), 'body': body}


def write_module(tid, part):
    for old, new in PATCHES.get(tid, []):
        if part['body'].count(old) != 1:
            sys.exit(f'{tid}: patch target not found exactly once')
        part['body'] = part['body'].replace(old, new)
    for cdn, local in VENDOR.items():
        part['body'] = part['body'].replace(cdn, local)
    doc = ('<!doctype html>\n<html lang="zh-Hant"><head>'
           '<meta charset="utf-8"><script src="bridge.js"></script>' + part['head'] +
           '</head><body>' + part['body'] + '</body></html>\n')
    path = os.path.join(OUT, tid + '.html')
    with open(path, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(doc)
    print(f'{tid}: {len(doc):,} chars -> {os.path.normpath(path)}')


def audit():
    # A rebuilt page must still be a zero-AI-token, no-new-third-party page; refuse otherwise.
    sys.path.insert(0, HERE)
    import audit_suite
    if audit_suite.main() != 0:
        sys.exit('audit failed: see messages above')


def main(src):
    html = open(src, encoding='utf-8').read()
    m = re.search(r'<script type="application/json" id="tool-src">(.*?)</script>', html, re.S)
    if not m:
        sys.exit('tool-src block not found')
    tools = json.loads(m.group(1))
    os.makedirs(OUT, exist_ok=True)
    for tid, part in tools.items():
        revised = os.path.join(SRC, tid + '.html')
        if os.path.exists(revised):
            print(f'{tid}: using the revised module in tools/src instead of index.html')
            part = standalone_part(revised)
        write_module(tid, part)
    audit()


def main_module(tid, src):
    if not re.fullmatch(r'[a-z0-9]+', tid):
        sys.exit(f'bad module id: {tid!r}')
    os.makedirs(SRC, exist_ok=True)
    os.makedirs(OUT, exist_ok=True)
    kept = os.path.join(SRC, tid + '.html')
    part = standalone_part(src)
    shutil.copyfile(src, kept)
    write_module(tid, part)
    audit()


if __name__ == '__main__':
    if len(sys.argv) == 4 and sys.argv[1] == '--module':
        main_module(sys.argv[2], sys.argv[3])
    elif len(sys.argv) == 2:
        main(sys.argv[1])
    else:
        sys.exit(__doc__)
