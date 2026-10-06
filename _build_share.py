#!/usr/bin/env python3
"""
Build share.html from app.html.
Strips out the Renewal Modal HTML block and all renewal-related JS functions.
Run: python _build_share.py
"""
import re

with open('app.html', 'r', encoding='utf-8') as f:
    src = f.read()

# 1. Update title
out = src.replace(
    '<title>TRAM Subcontractor Renewal &amp; Offboarding — Prototype</title>',
    '<title>Subcontractor Dashboard — Offboarding &amp; Analytics</title>'
)

# 2. Remove the Renewal Modal HTML block (<!-- RENEWAL MODAL --> ... </div>)
out = re.sub(
    r'<!-- ═+\s*RENEWAL MODAL\s*═+\s*-->\s*<div class="modal-overlay" id="renewalModal">.*?</div>\s*\n\s*</div>\s*\n',
    '',
    out,
    flags=re.DOTALL
)

# 3. Remove renewal-related JS blocks:
#    - openRenewalModal / closeRenewalModal
#    - _rnwRender, _rnwRenderStep1, _rnwRenderStep2, _rnwRenderStep3
#    - _rnwChecklistChanged, _rnwStep1Next, _rnwStep2Next, _rnwBack, _rnwToggleLaptop
#    - submitRenewal, _rnwShowPolling, _rnwAnimatePadSteps, _rnwPoll, showRenewalSuccess
#    - _rnwStepperHTML, _rnwContractorBar
#    - let renewalData, let _rnwStep

# Remove the large renewal section between the RENEWAL MODAL comment and OFFBOARDING MODAL comment in JS
out = re.sub(
    r'// ── RENEWAL MODAL.*?// ── OFFBOARDING MODAL',
    '// ── OFFBOARDING MODAL',
    out,
    flags=re.DOTALL
)

# Remove any remaining renewal modal open/close calls
# (already stripped from HTML; openRenewalModal won't be called from UI)

print(f'app.html: {len(src):,} chars -> share.html: {len(out):,} chars')
print(f'Reduction: {len(src)-len(out):,} chars removed')

with open('share.html', 'w', encoding='utf-8') as f:
    f.write(out)

print('OK share.html written.')
