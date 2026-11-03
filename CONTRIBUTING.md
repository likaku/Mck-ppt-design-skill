# Contributing

The three contributions that help most, in order:

1. **Show a deck.** Open a "Show your deck" issue with screenshots. Real decks are the best documentation.
2. **Report a broken layout.** Use the bug template: layout name, 5-line repro, screenshot, viewer (PowerPoint / WPS / Keynote).
3. **Add a layout.** Open a layout request first. A new layout must have: a one-line docstring, an entry in `CATEGORIES` in `mck_ppt/__main__.py`, and a character budget in `references/layout-matrix.yaml`.

## Local setup

```bash
git clone https://github.com/likaku/Mck-ppt-design-skill && cd Mck-ppt-design-skill
pip install python-pptx lxml
python -m mck_ppt demo          # must print "QA passed"
python run_qa_tests.py          # per-layout QA, normal content
python run_qa_tests.py --stress # long content
```

## Rules the engine enforces (do not break)

No connectors, no shadows, no 3D, no gradients. Lines are thin rectangles. Every CJK run gets an East Asian font. See `docs/ARCHITECTURE.md`.
