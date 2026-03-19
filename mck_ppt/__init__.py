"""McKinsey PPT Design Framework — High-level Layout Function Library.

Usage:
    from mck_ppt import MckEngine
    eng = MckEngine(total_slides=30)
    eng.cover(title='My Title', subtitle='Subtitle')
    eng.toc(items=[('1','Topic','Description'), ...])
    eng.save('output/my_deck.pptx')
"""
from .engine import MckEngine
from .constants import *

__version__ = '1.0.0'
