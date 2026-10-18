"""Guard for everything the site serves to browsers: no AI/Claude calls, no new third parties.

The analysis tools run entirely in the visitor's browser and the Flask app has no model API key,
so using the site costs zero Claude tokens. This script keeps it that way: it is run at the end of
tools/build_suite.py (and can be run alone) and fails if a page starts to

  * mention an AI API or a claude.* capability other than downloads/permissions,
  * reference a host that is not on the allowlist,
  * use a network primitive outside the two known same-origin fetch() calls.

    python tools/audit_suite.py
"""
import glob
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..'))

# Served to browsers by app.py. static/vendor is third-party minified code and is checked by hash instead.
TARGETS = (glob.glob(os.path.join(ROOT, 'static', 'suite', '*.html')) +
           glob.glob(os.path.join(ROOT, 'static', 'suite', '*.js')) +
           [os.path.join(ROOT, 'templates', 'index.html')])

# Fonts are cosmetic and fall back to system fonts; w3.org only appears as the SVG xmlns string.
HOST_ALLOW = {'fonts.googleapis.com', 'fonts.gstatic.com', 'www.w3.org'}

# claude.use('<name>') capabilities the bridge answers. Anything else is rejected by bridge.js.
CAPABILITY_ALLOW = {'downloads', 'permissions'}

FORBIDDEN = [
    (r'anthropic', 'mentions Anthropic'),
    (r'api\.openai|openai\.com|generativelanguage|gemini\.google', 'mentions another AI API'),
    (r'/v1/messages|chat/completions', 'AI chat endpoint path'),
    (r'x-api-key|api[_-]?key|authorization\s*:\s*bearer', 'API key / bearer header'),
    (r'\bsendPrompt\s*\(', 'sendPrompt() sends a chat message'),
    # claude.ai is only ever user-facing text ("open this in claude.ai"), not a call
    (r'(?:window\.)?claude\.(?!use\b|shim\b|ai\b)[A-Za-z_]+', 'claude.<capability> other than use()/shim'),
]

# (pattern, files allowed to use it). fetch() is allowed only for the reactor sample (own server)
# and the classifier's POST /classify (own server).
NETWORK = [
    (r'XMLHttpRequest|sendBeacon|WebSocket|EventSource|importScripts', set()),
    (r'\bfetch\(', {'reactor.html', 'index.html'}),
]


def main():
    problems = []
    for path in sorted(TARGETS):
        name = os.path.basename(path)
        text = open(path, encoding='utf-8').read()
        for pat, why in FORBIDDEN:
            for m in re.finditer(pat, text, re.IGNORECASE):
                problems.append(f'{name}: {why}: {m.group(0)!r}')
        for host in set(re.findall(r'https?://([A-Za-z0-9.-]+)', text)):
            if host not in HOST_ALLOW:
                problems.append(f'{name}: host not on allowlist: {host}')
        for cap in re.findall(r'\.use\(\s*[\'"]([A-Za-z_]+)[\'"]\s*\)', text):
            if cap not in CAPABILITY_ALLOW:
                problems.append(f'{name}: unexpected capability use({cap!r})')
        for pat, allowed in NETWORK:
            if re.search(pat, text) and name not in allowed:
                problems.append(f'{name}: network primitive {pat!r} not allowed here')
    if problems:
        print('AUDIT FAILED')
        for p in problems:
            print('  -', p)
        return 1
    print(f'audit ok: {len(TARGETS)} files, no AI/Claude calls, hosts limited to {sorted(HOST_ALLOW)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
