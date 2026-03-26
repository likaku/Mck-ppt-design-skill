#!/usr/bin/env python3
# Copyright 2024-2026 Kaku Li (https://github.com/likaku)
# Licensed under the Apache License, Version 2.0 — see LICENSE and NOTICE.
# Part of "Mck-ppt-design-skill" (McKinsey PPT Design Framework).
# NOTICE: This file must be retained in all copies or substantial portions.
#
"""
PPT QA Test Runner — generates all 55 layout methods and runs QA analysis.

Usage:
    python run_qa_tests.py                   # Run all, print summary
    python run_qa_tests.py --json report.json # Also save JSON report
    python run_qa_tests.py --only cover,toc   # Run specific methods only
    python run_qa_tests.py --stress           # Stress test with long content
"""

import os
import sys
import json
import time
import argparse
from datetime import datetime

# Add parent dir to path so we can import mck_ppt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx.util import Inches
from mck_ppt import MckEngine
from mck_ppt.constants import *
from mck_ppt.qa import PptQA, QAReport, Severity

# ── Test Data Fixtures ─────────────────────────────────────────────────

# Short / normal content
NORMAL = {
    "title": "Analysis of Market Dynamics",
    "subtitle": "Q1 2026 Strategic Review",
    "author": "Strategy Team",
    "long_title": "Comprehensive Analysis of Global Market Dynamics and Competitive Landscape Assessment for Strategic Decision Making",
    "bullet_3": ["Revenue grew 23% YoY", "Market share expanded to 34%", "Customer satisfaction at 92%"],
    "bullet_5": ["Revenue grew 23% YoY", "Market share expanded to 34%", "Customer satisfaction at 92%", "Employee engagement improved 15%", "Operating margin expanded 200bps"],
    "bullet_7": ["Revenue grew 23% YoY driven by new product launches", "Market share expanded to 34% across all segments", "Customer satisfaction at 92% reflecting service improvements", "Employee engagement improved 15% with new culture initiatives", "Operating margin expanded 200bps through cost optimization", "Digital transformation achieved 85% adoption rate", "Sustainability targets met across all business units"],
    "stats_2": [("$4.2B", "Revenue", "+23% YoY"), ("34%", "Market Share", "+5pp vs prior year")],
    "stats_3": [("$4.2B", "Revenue", "+23% YoY"), ("34%", "Market Share", "+5pp"), ("92%", "CSAT Score", "Industry benchmark: 78%")],
    "headers": ["Metric", "Q1", "Q2", "Q3", "Q4"],
    "rows_4": [
        ["Revenue ($M)", "245", "312", "298", "356"],
        ["Gross Margin", "62%", "64%", "63%", "65%"],
        ["Net Income ($M)", "42", "55", "48", "67"],
        ["Headcount", "1,234", "1,312", "1,298", "1,356"],
    ],
    "rows_8": [
        ["Revenue ($M)", "245", "312", "298", "356"],
        ["Cost of Goods", "93", "112", "110", "125"],
        ["Gross Margin", "62%", "64%", "63%", "65%"],
        ["OpEx ($M)", "89", "95", "92", "98"],
        ["EBITDA ($M)", "63", "105", "96", "133"],
        ["Net Income ($M)", "42", "55", "48", "67"],
        ["Headcount", "1,234", "1,312", "1,298", "1,356"],
        ["NPS Score", "72", "75", "74", "78"],
    ],
    "cards_3": [
        ("Revenue", "$4.2B", "+23%"),
        ("Market Share", "34%", "+5pp"),
        ("CSAT", "92%", "↑ 8pts"),
    ],
    "cards_5": [
        ("Revenue", "$4.2B", "+23%"),
        ("Market Share", "34%", "+5pp"),
        ("CSAT", "92%", "↑ 8pts"),
        ("Retention", "96%", "+2pp"),
        ("NPS", "78", "Industry avg: 62"),
    ],
}

# Stress test content (long text, many items)
STRESS = {
    "long_text": "This is a comprehensive analysis spanning multiple domains including market dynamics, competitive positioning, customer segmentation, product portfolio optimization, and operational efficiency improvements. " * 3,
    "bullet_12": [f"Critical initiative #{i+1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor" for i in range(12)],
    "rows_15": [[f"Row {i+1}"] + [f"{(i+1)*10+j}" for j in range(4)] for i in range(15)],
}


# ── Layout Test Definitions ────────────────────────────────────────────
def build_all_tests(eng: MckEngine, stress: bool = False):
    """
    Generate one slide per layout method. Returns list of (method_name, success, error).
    """
    results = []

    def _run(name, fn, *args, **kwargs):
        try:
            fn(*args, **kwargs)
            results.append((name, True, None))
        except Exception as e:
            results.append((name, False, str(e)))

    # ── Structure ──
    _run("cover", eng.cover, title=NORMAL["title"], subtitle=NORMAL["subtitle"], author=NORMAL["author"])
    _run("cover_long", eng.cover, title=NORMAL["long_title"], subtitle=NORMAL["subtitle"], author=NORMAL["author"], date="March 2026")
    _run("section_divider", eng.section_divider, section_label="01", title="Market Overview", subtitle="Current landscape analysis")
    _run("toc", eng.toc, title="Agenda", items=[("1", "Market Overview", "Current landscape"), ("2", "Competitive Analysis", "Key players"), ("3", "Strategy", "Recommended actions"), ("4", "Financials", "Projections and targets")])
    _run("appendix_title", eng.appendix_title, title="Appendix", subtitle="Supporting Materials")
    _run("closing", eng.closing, title="Thank You", message="Questions and discussion", source_text="Contact: strategy@company.com")

    # ── Data ──
    _run("big_number", eng.big_number, title="Revenue Growth", number="$4.2B", unit="Revenue", description="Year-over-year growth of 23%, exceeding target by 5 percentage points", source="Q1 2026 Financial Report")
    _run("two_stat", eng.two_stat, title="Key Metrics", stats=NORMAL["stats_2"], source="Internal data")
    _run("three_stat", eng.three_stat, title="Performance Dashboard", stats=NORMAL["stats_3"], source="Analytics team")
    _run("data_table_4", eng.data_table, title="Quarterly Performance", headers=NORMAL["headers"], rows=NORMAL["rows_4"], source="Finance")
    _run("data_table_8", eng.data_table, title="Detailed Financial Summary", headers=NORMAL["headers"], rows=NORMAL["rows_8"], source="Finance")
    _run("metric_cards_3", eng.metric_cards, title="Key Performance Indicators", cards=NORMAL["cards_3"], source="Dashboard")
    _run("metric_cards_5", eng.metric_cards, title="Extended KPI Dashboard", cards=NORMAL["cards_5"], source="Dashboard")
    # rag_status: rows = (name, status_color:RGBColor, *values, note)
    _run("rag_status", eng.rag_status, title="Project Status", headers=["Project", "Status", "Owner", "Due Date"], rows=[("Alpha", ACCENT_GREEN, "J. Smith", "Mar 2026"), ("Beta", ACCENT_ORANGE, "K. Lee", "Apr 2026"), ("Gamma", ACCENT_RED, "M. Wang", "May 2026"), ("Delta", ACCENT_GREEN, "A. Chen", "Jun 2026")], source="PMO")
    # scorecard: items = (name, score_str, pct_float)
    _run("scorecard", eng.scorecard, title="Team Scorecard", items=[("Revenue Target", "92%", 0.92), ("Customer NPS", "78", 0.78), ("Employee Engagement", "85%", 0.85), ("Cost Efficiency", "65%", 0.65), ("Innovation Index", "55%", 0.55)], source="HR")

    # ── Framework ──
    # matrix_2x2: quadrants = (label, bg_color:RGBColor, description)
    _run("matrix_2x2", eng.matrix_2x2, title="Priority Matrix", quadrants=[("High Impact / Low Effort", LIGHT_BLUE, "Quick wins, process automation"), ("High Impact / High Effort", LIGHT_GREEN, "Market expansion, M&A targets"), ("Low Impact / Low Effort", LIGHT_ORANGE, "Minor fixes, low priority"), ("Low Impact / High Effort", LIGHT_RED, "Deprioritize, avoid")], source="Strategy")
    # #14 three_pillar retired (v2.0.4) → use #71 table_insight
    _run("table_insight_strategic", eng.table_insight, title="Strategic Pillars", headers=["Pillar", "Key Initiatives", "Expected Impact"], rows=[["Growth", "Market expansion\nNew products\nPartnerships", "Revenue +20%"], ["Efficiency", "Cost optimization\nProcess automation\nShared services", "Margin +5pp"], ["Innovation", "R&D investment\nDigital transformation\nAI/ML adoption", "New revenue streams"]], insights=["Growth is the top priority for FY2026", "Efficiency gains fund innovation investments", "AI/ML adoption accelerates all three pillars"], source="CEO Office")
    # pyramid: levels = (label, description, width_inches:float)
    _run("pyramid", eng.pyramid, title="Capability Pyramid", levels=[("Vision", "Long-term direction", 4.0), ("Strategy", "How we compete", 6.0), ("Execution", "Day-to-day delivery", 8.0), ("Foundation", "People and culture", 10.0)], source="Org Design")
    # process_chevron: steps = (label, step_title, description)
    _run("process_chevron", eng.process_chevron, title="Delivery Process", steps=[("01", "Discovery", "Research & insights"), ("02", "Design", "Solution architecture"), ("03", "Build", "Development sprint"), ("04", "Test", "QA & validation"), ("05", "Launch", "Go-to-market")], source="PMO")
    # venn: circles = (label, points:list[str], x, y, w, h)
    _run("venn", eng.venn, title="Strategic Overlap", circles=[("Customer\nNeeds", ["Usability", "Reliability", "Value"], Inches(1.0), Inches(1.5), Inches(4.0), Inches(3.5)), ("Our\nCapabilities", ["Technology", "Talent", "IP"], Inches(4.0), Inches(1.5), Inches(4.0), Inches(3.5)), ("Market\nOpportunity", ["Growth areas", "Unmet needs"], Inches(2.5), Inches(3.0), Inches(4.0), Inches(3.0))], overlap_label="Sweet Spot", source="Strategy")
    _run("temple", eng.temple, title="Operating Model", roof_text="Customer Excellence", pillar_names=["People", "Process", "Technology"], foundation_text="Culture & Values", source="COO")
    # cycle: phases = (label, x_inches, y_inches)
    _run("cycle", eng.cycle, title="Innovation Cycle", phases=[("Ideate", 0.0, 1.5), ("Validate", 3.5, 1.5), ("Build", 3.5, 4.0), ("Scale", 0.0, 4.0)], source="Innovation Lab")
    # funnel: stages = (name, count_label, pct_float)
    _run("funnel", eng.funnel, title="Sales Funnel", stages=[("Awareness", "10,000", 1.0), ("Interest", "4,200", 0.42), ("Consideration", "1,800", 0.18), ("Decision", "720", 0.072), ("Purchase", "312", 0.031)], source="Sales Ops")

    # ── Comparison ──
    _run("side_by_side", eng.side_by_side, title="Option Comparison", options=[("Option A: Build", ["Full control", "Higher initial cost", "6-month timeline", "Custom to needs"]), ("Option B: Buy", ["Faster deployment", "Lower initial cost", "Vendor dependency", "Standard features"])], source="Tech team")
    _run("before_after", eng.before_after, title="Process Transformation", before_title="Current State", before_points=["Manual data entry", "5-day processing time", "12% error rate"], after_title="Future State", after_points=["Automated ingestion", "Same-day processing", "<1% error rate"], source="Ops")
    # pros_cons: conclusion = (label, text) or None
    _run("pros_cons", eng.pros_cons, title="Partnership Assessment", pros_title="Advantages", pros=["Market access", "Brand synergy", "Shared R&D costs"], cons_title="Risks", cons=["Cultural mismatch", "IP concerns", "Integration complexity"], conclusion=("Recommendation", "Proceed with phased approach"), source="BD Team")
    # swot: quadrants = (label, accent_color, light_bg, points:list[str])
    _run("swot", eng.swot, title="SWOT Analysis", quadrants=[("Strengths", ACCENT_BLUE, LIGHT_BLUE, ["Market leader", "Strong brand", "Talent pool"]), ("Weaknesses", ACCENT_ORANGE, LIGHT_ORANGE, ["Legacy systems", "High costs", "Slow innovation"]), ("Opportunities", ACCENT_GREEN, LIGHT_GREEN, ["Emerging markets", "AI adoption", "Partnerships"]), ("Threats", ACCENT_RED, LIGHT_RED, ["New entrants", "Regulation", "Recession risk"])], source="Strategy")

    # ── Narrative ──
    # executive_summary: items = (num, item_title, description)
    _run("executive_summary", eng.executive_summary, title="Executive Summary", headline="Strong Q1 performance driven by product innovation and market expansion", items=[("1", "Revenue", "Grew 23% YoY to $4.2B, exceeding guidance by 5pp"), ("2", "Market Share", "Expanded to 34%, up 5pp from prior year"), ("3", "Outlook", "Raising FY guidance by 8% based on strong pipeline")], source="CFO Report")
    _run("key_takeaway", eng.key_takeaway, title="Key Insights", left_text="Our analysis of 500+ enterprise customers reveals three critical success factors for digital transformation programs.", takeaways=["Executive sponsorship is the #1 predictor of success", "Cross-functional teams outperform siloed approaches by 3x", "Iterative delivery reduces time-to-value by 60%"], source="Research")
    _run("quote", eng.quote, quote_text="The best way to predict the future is to create it.", attribution="Peter Drucker")
    # two_column_text: columns = (letter, col_title, points:list[str])
    _run("two_column_text", eng.two_column_text, title="Market Analysis", columns=[("A", "Domestic Market", ["Revenue grew 23% YoY", "Market share expanded to 34%", "Customer satisfaction at 92%"]), ("B", "International Market", ["APAC grew 45% driven by China and India", "EMEA stable at 12% growth", "LATAM emerging as new growth driver"])], source="Market Intelligence")
    # four_column: items = (num, col_title, description)
    _run("four_column", eng.four_column, title="Capability Assessment", items=[("1", "People", "World-class talent with deep domain expertise across 30+ industries"), ("2", "Process", "Proven methodologies refined over thousands of engagements"), ("3", "Technology", "Cutting-edge tools and platforms enabling data-driven insights"), ("4", "Culture", "Collaborative environment fostering innovation and excellence")], source="HR")
    _run("meet_the_team", eng.meet_the_team, title="Leadership Team", members=[("Alice Chen", "CEO", "20 years in tech"), ("Bob Kim", "CFO", "Ex-Goldman Sachs"), ("Carol Liu", "CTO", "PhD CS, MIT"), ("David Park", "COO", "Supply chain expert")], source="HR")
    # case_study: sections = (letter, section_title, description)
    _run("case_study", eng.case_study, title="Client Success Story", sections=[("S", "Challenge", "Legacy systems causing 40% operational overhead"), ("A", "Approach", "Phased digital transformation over 18 months"), ("R", "Solution", "Cloud migration + process automation + AI analytics")], result_box=("Result", "60% cost reduction, 3x faster time-to-market"), source="Client reference (anonymized)")
    # action_items: actions = (action_title, timeline, description, owner)
    _run("action_items", eng.action_items, title="Next Steps", actions=[("Complete market research", "Week 1-2", "Conduct competitive analysis and customer surveys", "Strategy Team"), ("Draft partnership proposal", "Week 2-3", "Prepare terms and business case", "BD Team"), ("Build financial model", "Week 3-4", "Revenue projections and ROI analysis", "Finance"), ("Present to board", "Week 5", "Final recommendation and approval", "CEO")], source="PMO")

    # ── Timeline ──
    _run("timeline", eng.timeline, title="Project Roadmap", milestones=[("Q1", "Discovery & Planning"), ("Q2", "Design & Build MVP"), ("Q3", "Testing & Iteration"), ("Q4", "Launch & Scale")], source="PMO")
    _run("vertical_steps", eng.vertical_steps, title="Implementation Steps", steps=[("Step 1", "Assess current state", "Conduct diagnostic"), ("Step 2", "Design future state", "Define target architecture"), ("Step 3", "Build roadmap", "Prioritize initiatives"), ("Step 4", "Execute pilot", "Validate approach"), ("Step 5", "Scale rollout", "Full deployment")], source="Delivery")

    # ── Charts ──
    # grouped_bar: data[cat_idx][series_idx] — each inner list has len(series) elements
    _run("grouped_bar", eng.grouped_bar, title="Revenue by Region", categories=["NA", "EMEA", "APAC", "LATAM"], series=[("2024", NAVY), ("2025", ACCENT_BLUE)], data=[[120, 145], [85, 92], [95, 130], [45, 62]], source="Finance")
    # stacked_bar: data[period_idx][series_idx] — each inner list has len(series) elements
    _run("stacked_bar", eng.stacked_bar, title="Revenue Composition", periods=["Q1", "Q2", "Q3", "Q4"], series=[("Product", NAVY), ("Services", ACCENT_BLUE), ("License", ACCENT_GREEN)], data=[[50, 30, 20], [48, 30, 22], [45, 30, 25], [44, 28, 28]], source="Finance")
    # horizontal_bar: items = (name, pct_int, bar_color)
    _run("horizontal_bar", eng.horizontal_bar, title="Customer Satisfaction by Segment", items=[("Enterprise", 92, NAVY), ("Mid-Market", 85, ACCENT_BLUE), ("SMB", 78, ACCENT_GREEN), ("Consumer", 71, ACCENT_ORANGE)], source="CX Team")

    # ── Image layouts ──
    _run("content_right_image", eng.content_right_image, title="Product Overview", subtitle="Next-Generation Platform", bullets=["AI-powered analytics", "Real-time dashboards", "Cross-platform support"], takeaway="Launch scheduled for Q3 2026", source="Product")
    # three_images: items = (caption_title, description, image_label)
    _run("three_images", eng.three_images, title="Office Locations", items=[("New York HQ", "500+ employees serving East Coast clients", "NYC Office"), ("London", "200+ employees covering EMEA region", "London Office"), ("Singapore", "150+ employees for APAC expansion", "SG Office")], source="Facilities")
    _run("image_four_points", eng.image_four_points, title="Innovation Lab", image_label="Lab Photo", points=[("AI/ML", "Advanced machine learning capabilities"), ("Cloud", "Hybrid cloud infrastructure"), ("Data", "Petabyte-scale analytics"), ("Security", "Zero-trust architecture")], source="CTO")
    _run("full_width_image", eng.full_width_image, title="Campus Overview", image_label="Aerial View", overlay_text="Our state-of-the-art facility", source="Facilities")
    # case_study_image: sections = (label, text, accent_color)
    _run("case_study_image", eng.case_study_image, title="Digital Transformation", sections=[("Challenge", "Outdated infrastructure with 99.5% uptime target missed", ACCENT_RED), ("Solution", "Cloud-native platform with microservices architecture", ACCENT_GREEN)], image_label="Architecture Diagram", kpis=[("Uptime", "99.99%"), ("Latency", "-80%")], source="CTO")
    _run("quote_bg_image", eng.quote_bg_image, image_label="Background", quote_text="Innovation distinguishes between a leader and a follower.", attribution="Steve Jobs")
    # goals_illustration: goals = (goal_title, description, accent_color)
    _run("goals_illustration", eng.goals_illustration, title="2026 Strategic Goals", goals=[("Revenue $5B+", "23% growth target through organic and inorganic channels", ACCENT_BLUE), ("Market Leader", "Top 3 in all segments globally", ACCENT_GREEN), ("Best Employer", "Top 10 workplace globally with 90%+ engagement", ACCENT_ORANGE)], image_label="Vision Image", source="CEO")

    # ── Advanced Charts ──
    # donut: segments = (pct_float, color, label)
    _run("donut", eng.donut, title="Revenue Mix", segments=[(0.35, NAVY, "Product A"), (0.25, ACCENT_BLUE, "Product B"), (0.20, ACCENT_GREEN, "Product C"), (0.20, BG_GRAY, "Other")], center_label="$4.2B", center_sub="Total Revenue", source="Finance")
    _run("waterfall", eng.waterfall, title="Bridge Analysis", items=[("Q1 Base", 100, NAVY), ("Price", 15, ACCENT_GREEN), ("Volume", 25, ACCENT_GREEN), ("FX", -8, ACCENT_RED), ("Costs", -12, ACCENT_RED), ("Q1 Total", 120, NAVY)], source="FP&A")
    _run("line_chart", eng.line_chart, title="Monthly Trend", x_labels=["Jan", "Feb", "Mar", "Apr", "May", "Jun"], y_labels=["0", "25", "50", "75", "100"], values=[20, 35, 45, 60, 72, 88], legend_label="Revenue ($M)", source="Finance")
    _run("pareto", eng.pareto, title="Issue Analysis", items=[("System Errors", 45), ("User Errors", 30), ("Network Issues", 15), ("Hardware", 7), ("Other", 3)], source="Ops")
    # kpi_tracker: kpis = (name, pct_float, detail, status_key)
    _run("kpi_tracker", eng.kpi_tracker, title="KPI Dashboard", kpis=[("Revenue", 0.92, "$4.2B of $4.6B target", "on"), ("Margin", 0.85, "65% vs 70% target", "risk"), ("NPS", 0.78, "78 vs 85 target", "risk"), ("Retention", 0.96, "96% exceeding 95% target", "on")], source="Dashboard")
    # bubble: bubbles = (x_pct, y_pct, size_inches, label, color)
    _run("bubble", eng.bubble, title="Portfolio Analysis", bubbles=[(0.8, 0.9, 1.2, "Product A", NAVY), (0.4, 0.7, 0.8, "Product B", ACCENT_BLUE), (0.6, 0.4, 1.0, "Product C", ACCENT_GREEN), (0.2, 0.6, 0.5, "Product D", ACCENT_ORANGE)], x_label="Market Size", y_label="Growth Rate", source="Strategy")
    # risk_matrix: grid_lights must be RGBColor (not None), risks = (row, col, name)
    _run("risk_matrix", eng.risk_matrix, title="Risk Assessment", grid_colors=[[ACCENT_GREEN, ACCENT_GREEN, ACCENT_ORANGE], [ACCENT_GREEN, ACCENT_ORANGE, ACCENT_RED], [ACCENT_ORANGE, ACCENT_RED, ACCENT_RED]], grid_lights=[[LIGHT_GREEN, LIGHT_GREEN, LIGHT_ORANGE], [LIGHT_GREEN, LIGHT_ORANGE, LIGHT_RED], [LIGHT_ORANGE, LIGHT_RED, LIGHT_RED]], risks=[(2, 1, "Supply Chain"), (1, 2, "Cyber"), (0, 0, "Regulatory")], source="Risk Team")
    # gauge: benchmarks = (label, value_str, color)
    _run("gauge", eng.gauge, title="Performance Score", score=78, benchmarks=[("Industry Avg", "62", ACCENT_ORANGE), ("Target", "85", ACCENT_GREEN)], source="Analytics")
    _run("harvey_ball_table", eng.harvey_ball_table, title="Vendor Assessment", criteria=["Price", "Quality", "Support", "Innovation"], options=["Vendor A", "Vendor B", "Vendor C"], scores=[[4,3,2],[3,4,3],[2,2,4],[3,3,3]], source="Procurement")
    # pie: segments = (pct_float, color, label, sub_label)
    _run("pie", eng.pie, title="Market Share", segments=[(0.34, NAVY, "Us", "$1.4B"), (0.28, ACCENT_BLUE, "Competitor A", "$1.2B"), (0.22, ACCENT_GREEN, "Competitor B", "$0.9B"), (0.16, BG_GRAY, "Others", "$0.7B")], source="Market Research")
    _run("stacked_area", eng.stacked_area, title="Revenue Trend", years=["2022", "2023", "2024", "2025"], series_data=[("Product", [60, 75, 90, 110], NAVY), ("Services", [30, 38, 45, 55], ACCENT_BLUE), ("License", [20, 25, 30, 38], ACCENT_GREEN)], source="Finance")

    # ── Dashboard & Special ──
    _run("dashboard_kpi_chart", eng.dashboard_kpi_chart, title="Executive Dashboard", kpi_cards=[("Revenue", "$4.2B", "+23%", NAVY), ("Margin", "65%", "+2pp", ACCENT_BLUE), ("NPS", "78", "+8", ACCENT_GREEN)], source="Analytics")
    # stakeholder_map: quadrants = (label_cn, label_en, bg_color, members:list[str])
    _run("stakeholder_map", eng.stakeholder_map, title="Stakeholder Map", quadrants=[("高权力/高兴趣", "High Power / High Interest", LIGHT_BLUE, ["Board of Directors", "CEO", "Key Investors"]), ("高权力/低兴趣", "High Power / Low Interest", LIGHT_GREEN, ["Regulators", "Government Agencies"]), ("低权力/高兴趣", "Low Power / High Interest", LIGHT_ORANGE, ["Customers", "Media", "Employees"]), ("低权力/低兴趣", "Low Power / Low Interest", BG_GRAY, ["General Public", "Industry Associations"])], source="Comms")
    # decision_tree: root = (label,); branches = (L1_title, L1_metric, L1_color, children:list[(name, metric)])
    _run("decision_tree", eng.decision_tree, title="Decision Framework", root=("Strategic Investment?",), branches=[("Build In-House", "$2M / 6mo", ACCENT_BLUE, [("Team Available", "Start Q2"), ("Need Hiring", "Start Q3")]), ("Partner / Buy", "$1M / 3mo", ACCENT_GREEN, [("Existing Partner", "Negotiate terms"), ("New Vendor", "RFP process")])], source="Strategy")
    _run("checklist", eng.checklist, title="Project Readiness", columns=["Workstream", "Status", "Owner", "Due"], col_widths=[0.35, 0.2, 0.25, 0.2], rows=[("Data Migration", "✅ Complete", "J. Smith", "Done"), ("API Integration", "🔄 In Progress", "K. Lee", "Mar 30"), ("UAT Testing", "⏳ Pending", "M. Wang", "Apr 15"), ("Go-Live Prep", "⏳ Pending", "A. Chen", "Apr 30")], source="PMO")
    _run("metric_comparison", eng.metric_comparison, title="Year-over-Year", metrics=[("Revenue", "$3.4B", "$4.2B", "+23%"), ("Margin", "63%", "65%", "+2pp"), ("Headcount", "1,100", "1,356", "+23%")], source="Finance")
    # icon_grid: items = (item_title, description, accent_color)
    _run("icon_grid", eng.icon_grid, title="Capability Map", items=[("Analytics", "Data-driven insights platform", ACCENT_BLUE), ("AI/ML", "Predictive models and automation", ACCENT_GREEN), ("Cloud", "Scalable hybrid infrastructure", ACCENT_ORANGE), ("Security", "Zero-trust architecture", ACCENT_RED), ("Mobile", "Cross-platform native apps", NAVY), ("Global", "Multi-region deployment", ACCENT_BLUE)], source="CTO")
    # agenda: headers = (label, width), items = (*values, item_type)
    _run("agenda", eng.agenda, title="Workshop Agenda", headers=[("Time", 0.15), ("Topic", 0.55), ("Presenter", 0.30)], items=[("9:00", "Welcome & Objectives", "CEO", "key"), ("9:30", "Market Overview", "Strategy Lead", "normal"), ("10:30", "Break", "—", "break"), ("10:45", "Workshop Session", "All Participants", "key"), ("12:00", "Wrap-up & Next Steps", "COO", "normal")], source="EA")
    # value_chain: stages = (stage_title, description, accent_color)
    _run("value_chain", eng.value_chain, title="Value Chain", stages=[("Inbound\nLogistics", "Sourcing & receiving", ACCENT_BLUE), ("Operations", "Manufacturing & assembly", ACCENT_GREEN), ("Outbound\nLogistics", "Distribution & delivery", ACCENT_ORANGE), ("Marketing\n& Sales", "Go-to-market execution", ACCENT_RED), ("Service", "Post-sale support", NAVY)], source="COO")
    _run("numbered_list_panel", eng.numbered_list_panel, title="Key Recommendations", items=[("Accelerate digital transformation program by 6 months", "Leverage existing vendor partnerships to reduce timeline"), ("Invest $50M in AI capabilities", "Focus on customer-facing applications first"), ("Restructure operations for efficiency", "Target 15% cost reduction over 18 months")])
    # two_col_image_grid: items = (card_title, description, accent_color, image_label)
    _run("two_col_image_grid", eng.two_col_image_grid, title="Project Showcase", items=[("Phase 1: Discovery", "Research and stakeholder interviews", ACCENT_BLUE, "Research Photo"), ("Phase 2: Design", "Solution architecture and prototyping", ACCENT_GREEN, "Design Photo"), ("Phase 3: Build", "Agile development sprints", ACCENT_ORANGE, "Build Photo"), ("Phase 4: Launch", "Deployment and training", NAVY, "Launch Photo")], source="Delivery")

    # ── Stress tests (if enabled) ──
    if stress:
        _run("stress_data_table_15", eng.data_table, title="Extended Data (15 rows)", headers=NORMAL["headers"], rows=STRESS["rows_15"], source="Stress test")
        _run("stress_long_title", eng.executive_summary, title=STRESS["long_text"][:200], headline=STRESS["long_text"][:300], items=[("Key Finding", STRESS["long_text"])], source="Stress test")

    return results


# ── Main ───────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="PPT QA Test Runner")
    parser.add_argument("--json", help="Output JSON report path")
    parser.add_argument("--only", help="Comma-separated method names to test")
    parser.add_argument("--stress", action="store_true", help="Include stress tests")
    parser.add_argument("--outdir", default="qa_output", help="Output directory")
    args = parser.parse_args()

    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pptx_path = os.path.join(outdir, f"qa_test_{timestamp}.pptx")

    print(f"\n{'='*70}")
    print(f"  McKinsey PPT Design — QA Test Runner")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*70}\n")

    # Step 1: Generate test deck
    print("Step 1: Generating test deck with all layout methods...")
    t0 = time.time()
    eng = MckEngine(total_slides=80)
    results = build_all_tests(eng, stress=args.stress)

    # Filter if --only specified
    if args.only:
        allowed = set(args.only.split(","))
        results = [(n, s, e) for n, s, e in results if n in allowed]

    # Save
    eng.save(pptx_path)
    gen_time = time.time() - t0
    print(f"  → Generated {len(results)} layouts in {gen_time:.1f}s → {pptx_path}")

    # Report generation failures
    gen_failures = [(n, e) for n, s, e in results if not s]
    gen_success = [(n, e) for n, s, e in results if s]
    if gen_failures:
        print(f"\n  ⚠️  {len(gen_failures)} generation failure(s):")
        for name, err in gen_failures:
            print(f"     ❌ {name}: {err}")
    print(f"  ✅ {len(gen_success)} layouts generated successfully\n")

    # Step 2: Run QA analysis
    print("Step 2: Running QA analysis...")
    t1 = time.time()
    qa = PptQA(pptx_path)
    report = qa.run()
    qa_time = time.time() - t1
    print(f"  → Analyzed {report.total_slides} slides in {qa_time:.1f}s\n")

    # Step 3: Print results
    print("Step 3: Results\n")
    report.print_summary()

    # Annotated summary: map slide numbers to layout method names
    print(f"\n{'─'*70}")
    print(f"  Layout Method → Slide Score Mapping")
    print(f"{'─'*70}")
    slide_num = 0
    for name, success, _ in results:
        if success:
            slide_num += 1
            score = report.slide_scores.get(slide_num, 100)
            status = "✅" if score >= 90 else "⚠️ " if score >= 70 else "❌"
            issues_for_slide = [i for i in report.issues if i.slide_num == slide_num]
            issue_summary = ""
            if issues_for_slide:
                cats = set(i.category for i in issues_for_slide)
                issue_summary = f"  [{', '.join(cats)}]"
            print(f"  {status} Slide {slide_num:2d}: {name:30s} → {score:3d}/100{issue_summary}")

    # Step 4: Save JSON
    json_path = args.json or os.path.join(outdir, f"qa_report_{timestamp}.json")
    full_report = {
        "meta": {
            "timestamp": timestamp,
            "pptx_path": pptx_path,
            "generation_time_sec": round(gen_time, 1),
            "qa_analysis_time_sec": round(qa_time, 1),
            "total_layouts_tested": len(results),
            "generation_failures": len(gen_failures),
        },
        "layout_results": [
            {"method": n, "generated": s, "error": e, "slide_num": i+1 if s else None}
            for i, (n, s, e) in enumerate(results)
        ],
        "qa_report": json.loads(report.to_json()),
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(full_report, f, indent=2, ensure_ascii=False)
    print(f"\n  📄 JSON report: {json_path}")

    # Final verdict
    print(f"\n{'='*70}")
    if report.passed and not gen_failures:
        print(f"  ✅ ALL CHECKS PASSED — Score: {report.overall_score}/100")
    elif report.passed:
        print(f"  ⚠️  QA passed (Score: {report.overall_score}/100) but {len(gen_failures)} generation failure(s)")
    else:
        print(f"  ❌ {len(report.errors)} ERROR(s) found — Score: {report.overall_score}/100")
    print(f"{'='*70}\n")

    return 0 if report.passed and not gen_failures else 1


if __name__ == "__main__":
    sys.exit(main())