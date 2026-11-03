# Copyright 2024-2026 Kaku Li (https://github.com/likaku)
# Licensed under the Apache License, Version 2.0 — see LICENSE and NOTICE.
# Part of "Mck-ppt-design-skill" (McKinsey PPT Design Framework).
# NOTICE: This file must be retained in all copies or substantial portions.
#
"""Zero-config demo deck — proves the engine works in one command.

    python -m mck_ppt demo            # writes ./mck_demo.pptx
    python -m mck_ppt demo -o out.pptx
"""
from .constants import *
from .deck_builder import DeckBuilder


def demo_storyline():
    """A 13-slide mini consulting deck: every slide uses a different layout."""
    return [
        {'type': 'cover', 'data': dict(
            title='Retail is shifting from stores to relationships',
            subtitle='Board strategy review', author='Strategy Office', date='2026')},
        {'type': 'toc', 'data': dict(title='Agenda', items=[
            ('1', 'Where we are', 'Performance and market position'),
            ('2', 'What changed', 'Three structural shifts'),
            ('3', 'What we do next', 'Roadmap and decisions needed')])},
        {'type': 'executive_summary', 'data': dict(
            title='Executive summary',
            headline='Growth is intact, but 70% of it now depends on one channel',
            items=[('1', 'Revenue', 'Grew 23% YoY to $4.2B, 5pp above guidance'),
                   ('2', 'Concentration', 'Digital channels deliver 70% of net new revenue'),
                   ('3', 'Decision', 'Approve a $120M channel-diversification program')],
            source='CFO report, Q1 2026')},
        {'type': 'big_number', 'data': dict(
            title='Revenue is ahead of plan', number='$4.2B', unit='Revenue',
            description='Year-over-year growth of 23%, exceeding target by 5 percentage points',
            source='Q1 2026 financial report')},
        {'type': 'table_insight', 'data': dict(
            title='Three shifts are restructuring the market',
            headers=['Dimension', 'Before', 'After'],
            rows=[['Distribution', 'Offline-first', 'Digital-first'],
                  ['Pricing', 'Cost-plus', 'Value-based'],
                  ['Loyalty', 'Points', 'Membership']],
            insights=['Digital channels now drive 60% of acquisition',
                      'Value pricing unlocks about 3x margin',
                      'Membership doubles repeat purchase'],
            source='Market study, 2026')},
        {'type': 'donut', 'data': dict(
            title='Revenue mix is balanced across four lines',
            segments=[(0.35, NAVY, 'Product A'), (0.25, ACCENT_BLUE, 'Product B'),
                      (0.20, ACCENT_GREEN, 'Product C'), (0.20, BG_GRAY, 'Other')],
            center_label='$4.2B', center_sub='Total revenue', source='Finance')},
        {'type': 'waterfall', 'data': dict(
            title='Price and volume drove the quarter',
            items=[('Q1 base', 100, 'base'), ('Price', 15, 'up'), ('Volume', 25, 'up'),
                   ('FX', -8, 'down'), ('Costs', -12, 'down'), ('Q1 total', 120, 'base')],
            summary='Price and volume added $40M; FX and costs gave back $20M',
            source='FP&A')},
        {'type': 'matrix_2x2', 'data': dict(
            title='Quick wins come first',
            quadrants=[('High impact / low effort', LIGHT_BLUE, 'Process automation'),
                       ('High impact / high effort', LIGHT_GREEN, 'Market expansion'),
                       ('Low impact / low effort', LIGHT_ORANGE, 'Minor fixes'),
                       ('Low impact / high effort', LIGHT_RED, 'Deprioritize')],
            source='Strategy office')},
        {'type': 'process_chevron', 'data': dict(
            title='Delivery runs in five steps',
            steps=[('01', 'Discover', 'Research and insight'), ('02', 'Design', 'Solution architecture'),
                   ('03', 'Build', 'Development sprint'), ('04', 'Test', 'Validation'),
                   ('05', 'Launch', 'Go to market')], source='PMO')},
        {'type': 'swot', 'data': dict(
            title='Position is strong; execution risk is internal',
            quadrants=[('Strengths', ACCENT_BLUE, LIGHT_BLUE, ['Market leader', 'Strong brand']),
                       ('Weaknesses', ACCENT_ORANGE, LIGHT_ORANGE, ['Legacy systems', 'High costs']),
                       ('Opportunities', ACCENT_GREEN, LIGHT_GREEN, ['AI adoption', 'Partnerships']),
                       ('Threats', ACCENT_RED, LIGHT_RED, ['New entrants', 'Regulation'])],
            source='Strategy office')},
        {'type': 'timeline', 'data': dict(
            title='The roadmap lands in four quarters',
            milestones=[('Q1', 'Foundation'), ('Q2', 'Pilot'), ('Q3', 'Scale'), ('Q4', 'Review')],
            source='PMO')},
        {'type': 'action_items', 'data': dict(
            title='Four decisions are needed this month',
            actions=[('Approve budget', 'Week 1', 'Release the first tranche', 'CFO'),
                     ('Name program lead', 'Week 1', 'Single accountable owner', 'CEO'),
                     ('Select vendors', 'Week 2-3', 'Shortlist and negotiate', 'Procurement'),
                     ('Kick off pilot', 'Week 4', 'Two regions first', 'COO')],
            source='PMO')},
        {'type': 'closing', 'data': dict(
            title='Thank you', message='Built with mck-ppt-design',
            source_text='github.com/likaku/Mck-ppt-design-skill')},
    ]


def build_demo(output_path='mck_demo.pptx'):
    return DeckBuilder.build(demo_storyline(), output_path)
