"""One-off: restyle app.html with IBM Carbon. Run: python _apply_carbon.py"""
import re, sys

src = open('app.html', encoding='utf-8', newline='').read()
css = open('_carbon.css', encoding='utf-8').read()

# ── 1. Replace the <style> block + add IBM Plex ───────────────────────────────
FONT = ('<link rel="preconnect" href="https://fonts.googleapis.com">\n'
        '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">\n')
css += '  .cds-ic{width:1em;height:1em;vertical-align:-.125em;flex-shrink:0}\n'
m = re.search(r'<style>.*?</style>', src, re.S)
assert m, 'style block not found'
out = src[:m.start()] + FONT + '<style>\n' + css + '</style>' + src[m.end():]

# ── 2. Hard-coded colours (Tailwind -> Carbon) outside the stylesheet ─────────
COLORS = {
 '6b7280':'525252','9ca3af':'6f6f6f','e5e7eb':'e0e0e0','d1d5db':'c6c6c6','f3f4f6':'f4f4f4','f9fafb':'f4f4f4',
 'fafafa':'f4f4f4','f7f8fa':'f4f4f4','374151':'161616','1f2937':'161616','1f2328':'161616','57606a':'525252',
 'dc2626':'da1e28','b91c1c':'a2191f','16a34a':'198038','15803d':'0e6027','1d4ed8':'0f62fe','1e40af':'0043ce',
 '2563eb':'0f62fe','3b82f6':'4589ff','eff6ff':'edf5ff','dbeafe':'d0e2ff','bfdbfe':'a6c8ff','93c5fd':'78a9ff',
 '86efac':'6fdc8c','f0fdf4':'defbe6','dcfce7':'defbe6','fee2e2':'ffd7d9','fca5a5':'ff8389','fecaca':'ffb3b8',
 'fff1f2':'fff1f1','fef2f2':'fff1f1','fef3c7':'fcf4d6','fffbeb':'fcf4d6','fef9c3':'fcf4d6','fcd34d':'f1c21b',
 'fde68a':'fddc69','92400e':'684e00','713f12':'684e00','b45309':'8e6a00','d97706':'b28600','f59e0b':'f1c21b',
 'f97316':'ff832b','ea580c':'ba4e00','c2410c':'8a3800','7c3aed':'8a3ffc','8b5cf6':'a56eff','6d28d9':'6929c4',
 'ede9fe':'e8daff','c4b5fd':'d4bbff','6366f1':'0f62fe','0891b2':'0072c3','e0f2fe':'e5f6ff','10b981':'24a148',
 '0d9488':'007d79','ec4899':'d02670','14532d':'0e6027','e2e8f0':'e0e0e0','f1f5f9':'f4f4f4','f0f4ff':'e8e8e8',
 'ffedd5':'ffd9be',
}
head, sep, tail = out.partition('</style>')
n_col = 0
def col(mo):
    global n_col
    k = mo.group(1).lower()
    if k in COLORS:
        n_col += 1
        return '#' + COLORS[k]
    return mo.group(0)
tail = re.sub(r'#([0-9a-fA-F]{6})\b', col, tail)

# ── 3. Corners: Carbon is square; pills stay round ────────────────────────────
tail = re.sub(r'border-radius:\s*10px;padding:16px 20px', 'border-radius:0;padding:16px 20px', tail)
tail = re.sub(r'border-radius:\s*(2|4|6|8)px', 'border-radius:0', tail)
tail = re.sub(r'border-radius:\s*(10|12)px', 'border-radius:16px', tail)
tail = re.sub(r'font-family:\s*monospace', 'font-family:IBM Plex Mono,monospace', tail)

# ── 4. Emoji -> Carbon icons ──────────────────────────────────────────────────
P = {
 'check':'M16,2A14,14,0,1,0,30,16,14,14,0,0,0,16,2ZM14,21.5908l-5-5L10.5906,15,14,18.4092,21.41,11l1.5957,1.5859Z',
 'error':'M16,2A13.914,13.914,0,0,0,2,16,13.914,13.914,0,0,0,16,30,13.914,13.914,0,0,0,30,16,13.914,13.914,0,0,0,16,2Zm5.4449,21L9,10.5557,10.5557,9,23,21.4448Z',
 'warning':'M16,2C8.3,2,2,8.3,2,16s6.3,14,14,14s14-6.3,14-14C30,8.3,23.7,2,16,2z M14.9,8h2.2v11h-2.2V8z M16,25 c-0.8,0-1.5-0.7-1.5-1.5S15.2,22,16,22c0.8,0,1.5,0.7,1.5,1.5S16.8,25,16,25z',
 'time':'M16,30A14,14,0,1,1,30,16,14,14,0,0,1,16,30ZM16,4A12,12,0,1,0,28,16,12,12,0,0,0,16,4Z M20.59 22L15 16.41 15 7 17 7 17 15.58 22 20.59 20.59 22z',
 'settings':'M27,16.76c0-.25,0-.5,0-.76s0-.51,0-.77l1.92-1.68A2,2,0,0,0,29.3,11L26.94,7a2,2,0,0,0-1.73-1,2,2,0,0,0-.64.1l-2.43.82a11.35,11.35,0,0,0-1.31-.75l-.51-2.52a2,2,0,0,0-2-1.61H13.64a2,2,0,0,0-2,1.61l-.51,2.52a11.48,11.48,0,0,0-1.32.75L7.43,6.06A2,2,0,0,0,6.79,6,2,2,0,0,0,5.06,7L2.7,11a2,2,0,0,0,.41,2.51L5,15.24c0,.25,0,.5,0,.76s0,.51,0,.77L3.11,18.45A2,2,0,0,0,2.7,21L5.06,25a2,2,0,0,0,1.73,1,2,2,0,0,0,.64-.1l2.43-.82a11.35,11.35,0,0,0,1.31.75l.51,2.52a2,2,0,0,0,2,1.61h4.72a2,2,0,0,0,2-1.61l.51-2.52a11.48,11.48,0,0,0,1.32-.75l2.42.82a2,2,0,0,0,.64.1,2,2,0,0,0,1.73-1L29.3,21a2,2,0,0,0-.41-2.51ZM25.21,24l-3.43-1.16a8.86,8.86,0,0,1-2.71,1.57L18.36,28H13.64l-.71-3.55a9.36,9.36,0,0,1-2.7-1.57L6.79,24,4.43,20l2.72-2.4a8.9,8.9,0,0,1,0-3.13L4.43,12,6.79,8l3.43,1.16a8.86,8.86,0,0,1,2.71-1.57L13.64,4h4.72l.71,3.55a9.36,9.36,0,0,1,2.7,1.57L25.21,8,27.57,12l-2.72,2.4a8.9,8.9,0,0,1,0,3.13L27.57,20Z M16,22a6,6,0,1,1,6-6A5.94,5.94,0,0,1,16,22Zm0-10a3.91,3.91,0,0,0-4,4,3.91,3.91,0,0,0,4,4,3.91,3.91,0,0,0,4-4A3.91,3.91,0,0,0,16,12Z',
 'user':'M16 4a5 5 0 11-5 5 5 5 0 015-5m0-2a7 7 0 107 7A7 7 0 0016 2zM26 30H24V25a5 5 0 00-5-5H13a5 5 0 00-5 5v5H6V25a7 7 0 017-7h6a7 7 0 017 7z',
 'chart':'M4,2H2V28a2,2,0,0,0,2,2H30V28H4V25H26V17H4V13H18V5H4ZM24,19v4H4V19ZM16,7v4H4V7Z',
 'email':'M28,6H4A2,2,0,0,0,2,8V24a2,2,0,0,0,2,2H28a2,2,0,0,0,2-2V8A2,2,0,0,0,28,6ZM25.8,8,16,14.78,6.2,8ZM4,24V8.91l11.43,7.91a1,1,0,0,0,1.14,0L28,8.91V24Z',
 'renew':'M12 10H6.78A11 11 0 0127 16h2A13 13 0 006 7.68V4H4v8h8zM20 22h5.22A11 11 0 015 16H3a13 13 0 0023 8.32V28h2V20H20z',
 'export':'M26 24v4H6V24H4v4l.0076-.0049A1.9977 1.9977 0 006 30H26a2 2 0 002-2h0V24zM6 12L7.411 13.405 15 5.825 15 24 17 24 17 5.825 24.591 13.405 26 12 16 2 6 12z',
 'locked':'M24,14H22V8A6,6,0,0,0,10,8v6H8a2,2,0,0,0-2,2V28a2,2,0,0,0,2,2H24a2,2,0,0,0,2-2V16A2,2,0,0,0,24,14ZM12,8a4,4,0,0,1,8,0v6H12ZM24,28H8V16H24Z',
 'document':'M25.7,9.3l-7-7C18.5,2.1,18.3,2,18,2H8C6.9,2,6,2.9,6,4v24c0,1.1,0.9,2,2,2h16c1.1,0,2-0.9,2-2V10C26,9.7,25.9,9.5,25.7,9.3 z M18,4.4l5.6,5.6H18V4.4z M24,28H8V4h8v6c0,1.1,0.9,2,2,2h6V28z M10 22H22V24H10z M10 16H22V18H10z',
 'download':'M26 24v4H6V24H4v4H4a2 2 0 002 2H26a2 2 0 002-2h0V24zM26 14L24.59 12.59 17 20.17 17 2 15 2 15 20.17 7.41 12.59 6 14 16 24 26 14z',
 'dash':'M7.7 4.7a14.7 14.7 0 00-3 3.1L6.3 9A13.26 13.26 0 018.9 6.3zM4.6 12.3l-1.9-.6A12.51 12.51 0 002 16H4A11.48 11.48 0 014.6 12.3zM2.7 20.4a14.4 14.4 0 002 3.9l1.6-1.2a12.89 12.89 0 01-1.7-3.3zM7.8 27.3a14.4 14.4 0 003.9 2l.6-1.9A12.89 12.89 0 019 25.7zM11.7 2.7l.6 1.9A11.48 11.48 0 0116 4V2A12.51 12.51 0 0011.7 2.7zM24.2 27.3a15.18 15.18 0 003.1-3.1L25.7 23A11.53 11.53 0 0123 25.7zM27.4 19.7l1.9.6A15.47 15.47 0 0030 16H28A11.48 11.48 0 0127.4 19.7zM29.2 11.6a14.4 14.4 0 00-2-3.9L25.6 8.9a12.89 12.89 0 011.7 3.3zM24.1 4.6a14.4 14.4 0 00-3.9-2l-.6 1.9a12.89 12.89 0 013.3 1.7zM20.3 29.3l-.6-1.9A11.48 11.48 0 0116 28v2A21.42 21.42 0 0020.3 29.3z',
}
FILL = {'check':'#198038', 'error':'#da1e28', 'warning':'#b28600'}
CP = {0x2705:'check', 0x274C:'error', 0x26A0:'warning', 0x23F3:'time', 0x23F0:'time', 0x23F1:'time',
      0x2699:'settings', 0x26AA:'dash', 0x2B07:'download', 0x1F464:'user', 0x1F4CA:'chart', 0x1F4E7:'email',
      0x1F4CB:'document', 0x1F4C4:'document', 0x1F504:'renew', 0x1F4E4:'export', 0x1F512:'locked'}
def svg(name):
    return ('<svg class="cds-ic" viewBox="0 0 32 32" fill="%s" aria-hidden="true"><path d="%s"/></svg>'
            % (FILL.get(name, 'currentColor'), P[name]))

EMO = re.compile('[\U0001F300-\U0001FAFF\u2705\u274C\u23F3\u2699\u26A0\u26AA\u23F1\u23F0\u2B07]\uFE0F?')
unmapped = {}
def emoji(mo):
    ch = mo.group(0)[0]
    name = CP.get(ord(ch))
    if not name:
        unmapped[hex(ord(ch))] = unmapped.get(hex(ord(ch)), 0) + 1
        return ''
    return svg(name)

lines = tail.split('\n')
strip_ctx = re.compile(r'\.textContent|alert\(|confirm\(')
n_emo = n_strip = 0
for i, ln in enumerate(lines):
    if not EMO.search(ln):
        continue
    if strip_ctx.search(ln):                      # plain-text contexts cannot hold an <svg>
        lines[i], k = re.subn(EMO.pattern + r'\s?', '', ln); n_strip += k
    else:
        lines[i], k = EMO.subn(emoji, ln); n_emo += k
tail = '\n'.join(lines)

out = head + sep + tail
open('app.html', 'w', encoding='utf-8', newline='').write(out)
print('colours recoloured:', n_col, '| emoji->icon:', n_emo, '| stripped (text contexts):', n_strip)
print('unmapped emoji dropped:', unmapped)
print('svg inside attribute (must be 0):', len(re.findall(r'="[^"\n]*<svg', out)))
print('old Tailwind hex left:', sorted({h for h in re.findall(r'#([0-9a-fA-F]{6})\b', tail) if h.lower() in COLORS}))
