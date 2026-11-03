# Copyright 2024-2026 Kaku Li (https://github.com/likaku)
# Licensed under the Apache License, Version 2.0 — see LICENSE and NOTICE.
# Part of "Mck-ppt-design-skill" (McKinsey PPT Design Framework).
# NOTICE: This file must be retained in all copies or substantial portions.
#
"""Command line entry:  python -m mck_ppt <command>

    demo [-o FILE]      build a 13-slide sample deck (no API key, no network)
    layouts [--json]    list all layout methods with their one-line summary
    version             print the version
"""
import argparse
import inspect
import json
import sys

from . import __version__

# category -> layout method names (kept in sync with README)
CATEGORIES = {
    'Structure': ['cover', 'section_divider', 'toc', 'agenda', 'appendix_title', 'closing'],
    'Key numbers': ['big_number', 'two_stat', 'three_stat', 'metric_cards', 'scorecard',
                    'kpi_tracker', 'rag_status', 'data_table', 'harvey_ball_table'],
    'Frameworks': ['table_insight', 'matrix_2x2', 'pyramid', 'process_chevron', 'temple',
                   'swot', 'value_chain', 'risk_matrix', 'stakeholder_map', 'decision_tree'],
    'Comparison': ['side_by_side', 'before_after', 'pros_cons', 'metric_comparison'],
    'Narrative': ['executive_summary', 'key_takeaway', 'four_column', 'two_column_text',
                  'case_study', 'action_items', 'checklist', 'quote', 'meet_the_team',
                  'numbered_list_panel', 'icon_grid'],
    'Timeline': ['timeline', 'vertical_steps'],
    'Charts': ['donut', 'waterfall', 'line_chart', 'pareto', 'bubble', 'grouped_bar',
               'stacked_bar', 'horizontal_bar', 'stacked_area', 'multi_bar_panel',
               'dashboard_kpi_chart', 'dashboard_table_chart'],
    'Image layouts': ['content_right_image', 'three_images', 'image_four_points',
                      'full_width_image', 'case_study_image', 'quote_bg_image',
                      'goals_illustration', 'two_col_image_grid'],
}


def _layouts():
    from .engine import MckEngine
    rows = []
    for cat, names in CATEGORIES.items():
        for n in names:
            fn = getattr(MckEngine, n, None)
            if fn is None:
                continue
            doc = (inspect.getdoc(fn) or '').splitlines()
            rows.append({'category': cat, 'name': n, 'summary': doc[0] if doc else '',
                         'signature': f'{n}{inspect.signature(fn)}'.replace('(self, ', '(').replace('(self)', '()')})
    return rows


def main(argv=None):
    p = argparse.ArgumentParser(prog='mck_ppt', description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest='cmd')
    d = sub.add_parser('demo', help='build the sample deck')
    d.add_argument('-o', '--output', default='mck_demo.pptx')
    l = sub.add_parser('layouts', help='list layouts')
    l.add_argument('--json', action='store_true')
    sub.add_parser('version')
    a = p.parse_args(argv)

    if a.cmd == 'demo':
        from .demo import build_demo
        build_demo(a.output)
        print(f'\nOpen it: {a.output}')
    elif a.cmd == 'layouts':
        rows = _layouts()
        if a.json:
            print(json.dumps(rows, ensure_ascii=False, indent=2))
        else:
            cur = None
            for r in rows:
                if r['category'] != cur:
                    cur = r['category']
                    print(f'\n[{cur}]')
                print(f"  {r['name']:<22} {r['summary']}")
    elif a.cmd == 'version':
        print(__version__)
    else:
        p.print_help()
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
