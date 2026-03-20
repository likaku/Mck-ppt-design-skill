#!/usr/bin/env python3
"""Generate white-on-transparent PNG icons for Staircase Evolution slides.

All icons: 200×200 px, transparent background, white strokes (~6px).
Store under assets/icons/ for use with eng.pyramid() icon parameter.

Icon library:
  Business icons:
    - icon_person_bust.png   — Single person bust (B2B decision maker)
    - icon_shield_check.png  — Shield with checkmark (quality certification)
    - icon_people_group.png  — Group of people (all consumers)
  Civilization icons:
    - icon_factory_gear.png  — Gear/cog wheel (Industrial Age)
    - icon_circuit_chip.png  — Microchip with pins (Information Age)
    - icon_ai_brain.png      — Brain with neural nodes (Intelligence Age)
"""
import math
import os

from PIL import Image, ImageDraw

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       'assets', 'icons')
os.makedirs(OUT_DIR, exist_ok=True)

SIZE = 200
WHITE = (255, 255, 255, 255)
TRANSPARENT = (0, 0, 0, 0)
STROKE = 6  # line thickness


# ═══════════════════════════════════════════
# Business Icons
# ═══════════════════════════════════════════

def icon_person_bust():
    """Single person bust — represents B2B procurement decision maker."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx = SIZE // 2

    head_r = 28
    head_cy = 58
    d.ellipse(
        [cx - head_r, head_cy - head_r, cx + head_r, head_cy + head_r],
        outline=WHITE, width=STROKE
    )

    body_top = head_cy + head_r + 12
    body_w = 72
    d.arc(
        [cx - body_w, body_top - 10, cx + body_w, body_top + body_w + 40],
        start=180, end=360, fill=WHITE, width=STROKE
    )

    cut_y = body_top + 42
    d.line(
        [(cx - body_w + 4, cut_y), (cx + body_w - 4, cut_y)],
        fill=WHITE, width=STROKE
    )
    return img


def icon_shield_check():
    """Shield with checkmark — represents quality certification."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx = SIZE // 2

    shield_pts = [
        (cx - 52, 35), (cx + 52, 35), (cx + 52, 100),
        (cx, 165), (cx - 52, 100),
    ]
    d.line(shield_pts + [shield_pts[0]], fill=WHITE, width=STROKE, joint='curve')

    check_pts = [(cx - 24, 98), (cx - 6, 120), (cx + 30, 72)]
    d.line(check_pts, fill=WHITE, width=STROKE + 1, joint='curve')
    return img


def icon_people_group():
    """Three people — represents all consumers / broad audience."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx = SIZE // 2

    # Center person
    d.ellipse([cx - 18, 38, cx + 18, 74], outline=WHITE, width=STROKE)
    d.arc([cx - 38, 78, cx + 38, 140], start=180, end=360, fill=WHITE, width=STROKE)

    # Left person
    lx = cx - 50
    d.ellipse([lx - 15, 52, lx + 15, 82], outline=WHITE, width=STROKE - 1)
    d.arc([lx - 32, 86, lx + 32, 140], start=180, end=360, fill=WHITE, width=STROKE - 1)

    # Right person
    rx = cx + 50
    d.ellipse([rx - 15, 52, rx + 15, 82], outline=WHITE, width=STROKE - 1)
    d.arc([rx - 32, 86, rx + 32, 140], start=180, end=360, fill=WHITE, width=STROKE - 1)

    # Smile arc
    d.arc([cx - 30, 142, cx + 30, 168], start=0, end=180, fill=WHITE, width=STROKE - 2)
    return img


# ═══════════════════════════════════════════
# Civilization Icons
# ═══════════════════════════════════════════

def icon_factory_gear():
    """Gear/cog wheel — represents Industrial Age."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx, cy = SIZE // 2, SIZE // 2

    outer_r, inner_r, teeth = 72, 52, 8
    tooth_half_angle = math.pi / (teeth * 2.5)

    pts = []
    for i in range(teeth):
        angle = 2 * math.pi * i / teeth
        for offset in [-tooth_half_angle, tooth_half_angle]:
            a = angle + offset
            pts.append((cx + outer_r * math.cos(a), cy + outer_r * math.sin(a)))
        valley_angle = angle + math.pi / teeth
        for offset in [-tooth_half_angle * 0.6, tooth_half_angle * 0.6]:
            a = valley_angle + offset
            pts.append((cx + inner_r * math.cos(a), cy + inner_r * math.sin(a)))

    d.line(pts + [pts[0]], fill=WHITE, width=STROKE, joint='curve')

    hole_r = 20
    d.ellipse(
        [cx - hole_r, cy - hole_r, cx + hole_r, cy + hole_r],
        outline=WHITE, width=STROKE
    )
    return img


def icon_circuit_chip():
    """Microchip with pins — represents Information Age."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx, cy = SIZE // 2, SIZE // 2

    chip_half = 42
    d.rounded_rectangle(
        [cx - chip_half, cy - chip_half, cx + chip_half, cy + chip_half],
        radius=6, outline=WHITE, width=STROKE
    )

    die_half = 22
    d.rectangle(
        [cx - die_half, cy - die_half, cx + die_half, cy + die_half],
        outline=WHITE, width=STROKE - 1
    )

    pin_count, pin_len = 4, 20
    pin_gap = (chip_half * 2) / (pin_count + 1)
    for i in range(1, pin_count + 1):
        offset = -chip_half + pin_gap * i
        d.line([(cx + offset, cy - chip_half), (cx + offset, cy - chip_half - pin_len)],
               fill=WHITE, width=STROKE - 2)
        d.line([(cx + offset, cy + chip_half), (cx + offset, cy + chip_half + pin_len)],
               fill=WHITE, width=STROKE - 2)
        d.line([(cx - chip_half, cy + offset), (cx - chip_half - pin_len, cy + offset)],
               fill=WHITE, width=STROKE - 2)
        d.line([(cx + chip_half, cy + offset), (cx + chip_half + pin_len, cy + offset)],
               fill=WHITE, width=STROKE - 2)
    return img


def icon_ai_brain():
    """Brain with neural nodes — represents Intelligence Age."""
    img = Image.new('RGBA', (SIZE, SIZE), TRANSPARENT)
    d = ImageDraw.Draw(img)
    cx, cy = SIZE // 2, SIZE // 2

    # Two hemispheres
    d.arc([cx - 68, cy - 52, cx + 2, cy + 52], start=90, end=270, fill=WHITE, width=STROKE)
    d.arc([cx - 2, cy - 52, cx + 68, cy + 52], start=270, end=450, fill=WHITE, width=STROKE)

    # Brain wrinkles
    d.arc([cx - 40, cy - 55, cx + 40, cy - 10], start=0, end=180, fill=WHITE, width=STROKE - 2)
    d.arc([cx - 35, cy - 18, cx + 35, cy + 18], start=180, end=360, fill=WHITE, width=STROKE - 2)

    # Center line
    d.line([(cx, cy - 52), (cx, cy + 52)], fill=WHITE, width=STROKE - 2)

    # Neural nodes
    node_r = 5
    nodes = [(cx - 30, cy - 20), (cx + 30, cy - 20),
             (cx - 20, cy + 18), (cx + 20, cy + 18), (cx, cy)]
    for nx, ny in nodes:
        d.ellipse([nx - node_r, ny - node_r, nx + node_r, ny + node_r], fill=WHITE)

    # Connections
    for a, b in [(0, 4), (1, 4), (2, 4), (3, 4), (0, 2), (1, 3)]:
        d.line([nodes[a], nodes[b]], fill=WHITE, width=2)
    return img


# ═══════════════════════════════════════════
# Generate and save all icons
# ═══════════════════════════════════════════

ALL_ICONS = [
    ('icon_person_bust.png', icon_person_bust),
    ('icon_shield_check.png', icon_shield_check),
    ('icon_people_group.png', icon_people_group),
    ('icon_factory_gear.png', icon_factory_gear),
    ('icon_circuit_chip.png', icon_circuit_chip),
    ('icon_ai_brain.png', icon_ai_brain),
]

if __name__ == '__main__':
    for name, fn in ALL_ICONS:
        path = os.path.join(OUT_DIR, name)
        fn().save(path, 'PNG')
        print(f'Saved: {path}')
    print(f'\nDone! {len(ALL_ICONS)} icons generated in {OUT_DIR}')
