"""
Export DODO_OnePager_DM.html → clean PNG (element-only, 2x) + single-page A4 PDF.
Run: python export_onepager.py
"""

import os
from pathlib import Path
from playwright.sync_api import sync_playwright

HTML = Path(r"c:\Users\C0378\Desktop\旺來PPT製作\嘟嘟房_提案專案\01_OnePager\DODO_OnePager_DM.html")
OUT  = Path(r"c:\Users\C0378\Desktop\旺來PPT製作\嘟嘟房_提案專案\01_OnePager\export")
OUT.mkdir(exist_ok=True)

PNG_PATH = OUT / "DODO_OnePager_DM.png"
PDF_PATH = OUT / "DODO_OnePager_DM.pdf"
URL      = HTML.as_uri()

with sync_playwright() as pw:
    browser = pw.chromium.launch()

    # ── PNG ──────────────────────────────────────────────────────────────────
    # device_scale_factor=2 → 1588×2246 px output (retina quality)
    # element.screenshot() crops to .page div exactly; no gray body background
    ctx = browser.new_context(viewport={"width": 794, "height": 1123},
                              device_scale_factor=2)
    pg = ctx.new_page()
    pg.goto(URL, wait_until="networkidle")
    pg.wait_for_timeout(800)
    pg.locator(".page").screenshot(path=str(PNG_PATH))
    ctx.close()
    print(f"PNG  {PNG_PATH.name}  {os.path.getsize(PNG_PATH):,} bytes")

    # ── PDF ──────────────────────────────────────────────────────────────────
    # viewport 794×1123 matches A4@96dpi → page.pdf with A4 format stays 1 page
    ctx2 = browser.new_context(viewport={"width": 794, "height": 1123})
    pg2  = ctx2.new_page()
    pg2.goto(URL, wait_until="networkidle")
    pg2.wait_for_timeout(800)
    pg2.pdf(
        path=str(PDF_PATH),
        format="A4",
        print_background=True,
        margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
        scale=1,
    )
    ctx2.close()
    print(f"PDF  {PDF_PATH.name}  {os.path.getsize(PDF_PATH):,} bytes")

    browser.close()

print("Done.")
