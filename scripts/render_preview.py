#!/usr/bin/env python3
# Copyright 2024-2026 Kaku Li (https://github.com/likaku)
# Licensed under the Apache License, Version 2.0 — see LICENSE and NOTICE.
"""Render a .pptx to one PNG per slide (for README screenshots and the gallery).

    python scripts/render_preview.py mck_demo.pptx docs/gallery

Needs LibreOffice (`soffice`) and poppler (`pdftoppm`):
    macOS:   brew install --cask libreoffice && brew install poppler
    Ubuntu:  sudo apt install libreoffice poppler-utils
Fonts: install KaiTi / Arial for pixel-faithful output.
"""
import os
import shutil
import subprocess
import sys
import tempfile


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    src, out = sys.argv[1], sys.argv[2]
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    ppm = shutil.which('pdftoppm')
    if not soffice or not ppm:
        sys.exit('Need soffice and pdftoppm on PATH. See the docstring.')
    os.makedirs(out, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp, src], check=True)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + '.pdf')
        subprocess.run([ppm, '-png', '-r', '110', pdf, os.path.join(out, 'slide')], check=True)
    print('wrote', sorted(os.listdir(out)))


if __name__ == '__main__':
    main()
