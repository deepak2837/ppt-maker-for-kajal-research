#!/usr/bin/env python3
"""
SRE Take-Home Assignment — PowerPoint Generator (Redesigned)
Condensed 17-slide deck with infographic visuals & premium fonts.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

# ── Paths ──────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
BG_TITLE   = os.path.join(BASE, "2024 Master Slide Deck (PPT) (6)-1.jpg")
BG_CONTENT = os.path.join(BASE, "2024 Master Slide Deck (PPT) (6)-3.jpg")
BG_SECTION = os.path.join(BASE, "2024 Master Slide Deck (PPT) (6)-2.jpg")
OUTPUT     = os.path.join(BASE, "SRE_Take_Home_Assignment.pptx")

# Infographic image paths
IMG = {}
for i in range(2, 16):
    IMG[i] = os.path.join(BASE, f"clean_Engineering_Reliability_Blueprint-{i}.jpg")

# ── Colours (Florence brand) ──────────────────────────────────────────
NAVY        = RGBColor(0x1B, 0x1F, 0x3B)
CORAL       = RGBColor(0xE8, 0x6B, 0x5A)
DARK_GREY   = RGBColor(0x33, 0x33, 0x33)
MID_GREY    = RGBColor(0x55, 0x55, 0x55)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_CORAL = RGBColor(0xF0, 0x9E, 0x90)
LIGHT_BG    = RGBColor(0xF8, 0xF8, 0xFA)

# ── Fonts (Premium) ───────────────────────────────────────────────────
FONT_TITLE  = "Century Gothic"
FONT_BODY   = "Calibri Light"
FONT_ACCENT = "Century Gothic"

# ── Slide dimensions (16:9) ───────────────────────────────────────────
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
BLANK_LAYOUT = prs.slide_layouts[6]

# ═══════════════════════════════════════════════════════════════════════
# Helper functions
# ═══════════════════════════════════════════════════════════════════════

def add_bg(slide, img_path):
    slide.shapes.add_picture(img_path, 0, 0, SLIDE_W, SLIDE_H)

def add_textbox(slide, left, top, width, height):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.word_wrap = True
    return txBox.text_frame

def set_para(tf, text, size=16, bold=False, color=DARK_GREY,
             align=PP_ALIGN.LEFT, after=Pt(8), before=Pt(0), font=None):
    if font is None:
        font = FONT_BODY
    p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font
    p.alignment = align
    p.space_after = after
    p.space_before = before
    return p

def accent_bar(slide, left, top, w, h, color=CORAL):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

def slide_num(slide, n):
    tf = add_textbox(slide, Inches(12.4), Inches(7.0), Inches(0.7), Inches(0.4))
    tf.paragraphs[0].text = str(n)
    tf.paragraphs[0].font.size = Pt(10)
    tf.paragraphs[0].font.color.rgb = MID_GREY
    tf.paragraphs[0].font.name = FONT_BODY
    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

def _title_bar(slide, title, size=26):
    """Shared title bar with coral accent for content slides."""
    accent_bar(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.06), CORAL)
    accent_bar(slide, Inches(0.5), Inches(0.32), Inches(0.08), Inches(0.50), CORAL)
    tf = add_textbox(slide, Inches(0.8), Inches(0.25), Inches(11.7), Inches(0.7))
    tf.paragraphs[0].text = title
    tf.paragraphs[0].font.size = Pt(size)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY
    tf.paragraphs[0].font.name = FONT_TITLE

# ═══════════════════════════════════════════════════════════════════════
# Slide layout builders
# ═══════════════════════════════════════════════════════════════════════

def title_slide(title, subtitle, extra=""):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_TITLE)
    tf = add_textbox(slide, Inches(1.5), Inches(2.0), Inches(10.3), Inches(1.5))
    tf.paragraphs[0].text = title
    tf.paragraphs[0].font.size = Pt(40)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY
    tf.paragraphs[0].font.name = FONT_TITLE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    set_para(tf, subtitle, size=22, color=CORAL, font=FONT_ACCENT,
             align=PP_ALIGN.CENTER, before=Pt(12))
    if extra:
        set_para(tf, extra, size=16, color=MID_GREY,
                 align=PP_ALIGN.CENTER, before=Pt(8))

def section_slide(title, subtitle=""):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_SECTION)
    tf = add_textbox(slide, Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.0))
    tf.paragraphs[0].text = title
    tf.paragraphs[0].font.size = Pt(38)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY
    tf.paragraphs[0].font.name = FONT_TITLE
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    if subtitle:
        set_para(tf, subtitle, size=20, color=NAVY, align=PP_ALIGN.LEFT,
                 before=Pt(12), font=FONT_ACCENT)

def content_slide(title, bullets, n=None):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_CONTENT)
    _title_bar(slide, title)
    tf = add_textbox(slide, Inches(0.8), Inches(1.2), Inches(11.7), Inches(5.5))
    for i, b in enumerate(bullets):
        if i == 0:
            tf.paragraphs[0].text = f"▸  {b}"
            tf.paragraphs[0].font.size = Pt(16)
            tf.paragraphs[0].font.color.rgb = DARK_GREY
            tf.paragraphs[0].font.name = FONT_BODY
            tf.paragraphs[0].space_after = Pt(10)
        else:
            set_para(tf, f"▸  {b}", size=16, after=Pt(10))
    if n: slide_num(slide, n)

def two_col_slide(title, l_title, l_bullets, r_title, r_bullets, n=None):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_CONTENT)
    _title_bar(slide, title)
    # Divider
    accent_bar(slide, Inches(6.55), Inches(1.2), Inches(0.03), Inches(5.5), LIGHT_CORAL)
    # Left
    tf_lh = add_textbox(slide, Inches(0.8), Inches(1.2), Inches(5.5), Inches(0.5))
    tf_lh.paragraphs[0].text = l_title
    tf_lh.paragraphs[0].font.size = Pt(18)
    tf_lh.paragraphs[0].font.bold = True
    tf_lh.paragraphs[0].font.color.rgb = CORAL
    tf_lh.paragraphs[0].font.name = FONT_ACCENT
    tf_l = add_textbox(slide, Inches(0.8), Inches(1.75), Inches(5.5), Inches(5.0))
    for i, b in enumerate(l_bullets):
        if i == 0:
            tf_l.paragraphs[0].text = f"▸  {b}"
            tf_l.paragraphs[0].font.size = Pt(14)
            tf_l.paragraphs[0].font.color.rgb = DARK_GREY
            tf_l.paragraphs[0].font.name = FONT_BODY
            tf_l.paragraphs[0].space_after = Pt(7)
        else:
            set_para(tf_l, f"▸  {b}", size=14, after=Pt(7))
    # Right
    tf_rh = add_textbox(slide, Inches(7.0), Inches(1.2), Inches(5.5), Inches(0.5))
    tf_rh.paragraphs[0].text = r_title
    tf_rh.paragraphs[0].font.size = Pt(18)
    tf_rh.paragraphs[0].font.bold = True
    tf_rh.paragraphs[0].font.color.rgb = CORAL
    tf_rh.paragraphs[0].font.name = FONT_ACCENT
    tf_r = add_textbox(slide, Inches(7.0), Inches(1.75), Inches(5.5), Inches(5.0))
    for i, b in enumerate(r_bullets):
        if i == 0:
            tf_r.paragraphs[0].text = f"▸  {b}"
            tf_r.paragraphs[0].font.size = Pt(14)
            tf_r.paragraphs[0].font.color.rgb = DARK_GREY
            tf_r.paragraphs[0].font.name = FONT_BODY
            tf_r.paragraphs[0].space_after = Pt(7)
        else:
            set_para(tf_r, f"▸  {b}", size=14, after=Pt(7))
    if n: slide_num(slide, n)

def infographic_slide(title, img_path, caption="", n=None):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_CONTENT)
    _title_bar(slide, title)
    slide.shapes.add_picture(img_path, Inches(1.2), Inches(1.1), Inches(10.9), Inches(5.7))
    if caption:
        tf = add_textbox(slide, Inches(1.0), Inches(6.85), Inches(11.3), Inches(0.4))
        tf.paragraphs[0].text = caption
        tf.paragraphs[0].font.size = Pt(12)
        tf.paragraphs[0].font.italic = True
        tf.paragraphs[0].font.color.rgb = MID_GREY
        tf.paragraphs[0].font.name = FONT_ACCENT
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    if n: slide_num(slide, n)

def img_bullets_slide(title, img_path, bullets, side="left", n=None):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_CONTENT)
    _title_bar(slide, title)
    if side == "left":
        img_x, txt_x = Inches(0.4), Inches(6.7)
    else:
        img_x, txt_x = Inches(6.7), Inches(0.4)
    slide.shapes.add_picture(img_path, img_x, Inches(1.15), Inches(6.1), Inches(5.9))
    tf = add_textbox(slide, txt_x, Inches(1.2), Inches(6.1), Inches(5.6))
    for i, b in enumerate(bullets):
        if i == 0:
            tf.paragraphs[0].text = f"▸  {b}"
            tf.paragraphs[0].font.size = Pt(14)
            tf.paragraphs[0].font.color.rgb = DARK_GREY
            tf.paragraphs[0].font.name = FONT_BODY
            tf.paragraphs[0].space_after = Pt(9)
        else:
            set_para(tf, f"▸  {b}", size=14, after=Pt(9))
    if n: slide_num(slide, n)

def table_slide(title, headers, rows, n=None):
    slide = prs.slides.add_slide(BLANK_LAYOUT)
    add_bg(slide, BG_CONTENT)
    _title_bar(slide, title)
    nr = len(rows) + 1
    nc = len(headers)
    table = slide.shapes.add_table(nr, nc, Inches(0.8), Inches(1.3),
                                   Inches(11.7), Inches(0.45) * nr).table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.font.name = FONT_ACCENT
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        sf = tcPr.makeelement(qn('a:solidFill'), {})
        sf.append(sf.makeelement(qn('a:srgbClr'), {'val': '1B1F3B'}))
        tcPr.append(sf)
    for r, rd in enumerate(rows):
        for c, v in enumerate(rd):
            cell = table.cell(r+1, c)
            cell.text = v
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.color.rgb = DARK_GREY
                p.font.name = FONT_BODY
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            sf = tcPr.makeelement(qn('a:solidFill'), {})
            fc = 'F5F5F5' if r % 2 == 0 else 'FFFFFF'
            sf.append(sf.makeelement(qn('a:srgbClr'), {'val': fc}))
            tcPr.append(sf)
    if n: slide_num(slide, n)


# ═══════════════════════════════════════════════════════════════════════
#  SLIDE CONTENT — 17 Slides, Condensed
# ═══════════════════════════════════════════════════════════════════════

# ── 1. Title ──────────────────────────────────────────────────────────
title_slide(
    "SRE Take-Home Assignment",
    "Observability  •  Performance  •  Incidents  •  Database  •  Philosophy",
    "Senior-Level Response  —  March 2026"
)

# ── 2. Section: Monitoring & Observability ────────────────────────────
section_slide("Part 1", "Monitoring & Observability Strategy")

# ── 3. 30-60-90 Day Roadmap ──────────────────────────────────────────
img_bullets_slide(
    "30-60-90 Day Observability Roadmap",
    IMG[3],
    [
        "As the first SRE hire: you cannot improve what you cannot see",
        "Days 1–30 — Instrument: deploy OTel collectors, export metrics to CloudWatch / Datadog",
        "Days 31–60 — Correlate: connect metrics + logs + traces; golden-signal dashboards",
        "Days 61–90 — Govern: define SLOs, wire error budgets, automate alerting",
        "Follow the See → Know → Act sequence",
        "Shift the org from reactive firefighting to data-driven governance",
    ],
    side="left", n=3
)

# ── 4. Metrics + SLIs/SLOs ───────────────────────────────────────────
two_col_slide(
    "Metrics Hierarchy & SLI/SLO Framework",
    "Key Metrics (Infra → App → Biz)",
    [
        "CPU/memory per ECS task, ALB 5xx, latency p50/p95/p99",
        "RDS connections, IOPS, replication lag",
        "Request/error rate, downstream dep latency",
        "Queue depth & consumer lag (SQS / async)",
        "Checkout success rate, user-facing latency",
        "Error budget burn rate — SLO margin consumption",
    ],
    "SLIs, SLOs & Error Budget",
    [
        "SLIs = quantitative service measures (latency, availability)",
        "SLOs = targets agreed by engineering + product jointly",
        "Availability ≥ 99.9% (30d), Latency p95 ≤ 300ms (7d)",
        "Error budget = tolerable unreliability (0.1% / month)",
        "Budget surplus → ship faster; exhausted → freeze releases",
        "Start conservative; tighten as measurement trust grows",
    ],
    n=4
)

# ── 5. Log Filtering + Tooling ───────────────────────────────────────
img_bullets_slide(
    "Signal Over Noise — Logs & Tooling Stack",
    IMG[6],
    [
        "Structured JSON logs: timestamp, level, service, trace_id, span_id",
        "Log discipline: DEBUG off in prod; ERROR = unhandled failures only",
        "Sample 1–5% healthy requests; always emit WARN/ERROR; head + tail sampling",
        "Inject trace_id at API gateway → every log ties to a single user request",
        "Stack: OpenTelemetry (vendor-agnostic) → Datadog / CloudWatch backend",
        "AWS FireLens for ECS log routing; X-Ray for distributed tracing",
    ],
    side="right", n=5
)

# ── 6. Section: Performance ──────────────────────────────────────────
section_slide("Part 2", "Performance Improvement Strategy")

# ── 7. Baseline, Bottlenecks & Golden Signals ────────────────────────
img_bullets_slide(
    "Baseline → Bottlenecks → Golden Signals",
    IMG[7],
    [
        "First 2 weeks: measure, don't fix — build a 'current state' dashboard",
        "Load test at 1×, 2×, 5× traffic (k6 / Locust); profile top 10 slowest endpoints",
        "Four bottleneck types: Compute, I/O, Concurrency, Network",
        "Four Golden Signals: Latency, Traffic, Errors, Saturation",
        "Document baseline as versioned artefact — the regression gate",
        "Frame perf in customer terms: 'p95=800ms costs 12% error budget'",
    ],
    side="left", n=7
)

# ── 8. Deployment Guardrails & Ownership ─────────────────────────────
img_bullets_slide(
    "Deployment Guardrails & Performance Ownership",
    IMG[14],
    [
        "Pre-deploy: k6 smoke tests on every PR; fail build if p95 exceeds baseline",
        "Deploy: canary at 5% traffic; compare golden signals vs control",
        "Post-deploy: auto-rollback on SLO breach; bake period before full rollout",
        "Assign rotating 'performance DRI' per service team",
        "Latency SLOs in definition of done — feature not done if p95 regresses >10%",
        "Monthly reliability newsletter + weekly 15-min async perf review with leads",
    ],
    side="right", n=8
)

# ── 9. Section: Incident Management ──────────────────────────────────
section_slide("Part 3", "L3 Support & Incident Management")

# ── 10. Severity & Triage ────────────────────────────────────────────
img_bullets_slide(
    "Incident Severity & Triage Workflow",
    IMG[15],
    [
        "SEV-1: Complete outage / data loss >50% users → 15 min response, 30 min comms",
        "SEV-2: Degraded, workaround exists → 30 min response",
        "SEV-3: Single-tenant / non-critical → 4 hr; SEV-4: Cosmetic → next sprint",
        "Replace ad-hoc Slack pings: dedicated Support Liaison (rotating role)",
        "Liaison validates L3 ticket: repro path, user IDs, traceIDs — incomplete → back to L2",
        "Every ticket tagged with root cause category; monthly Pareto → reliability OKRs",
    ],
    side="left", n=10
)

# ── 11. On-Call, Feedback & Recurring Issues ─────────────────────────
two_col_slide(
    "On-Call, Feedback Loops & Recurring Issues",
    "On-Call & Postmortems",
    [
        "Follow-the-Sun preferred; else week-on/week-off, 6–8 people",
        "Primary → initial alerts; Secondary → fallback & deep investigation",
        "Handoff document at end of each rotation",
        "Blameless postmortems for SEV-1/2 within 48 hrs",
        "Root cause → actionable follow-ups with owners + deadlines",
    ],
    "Reducing Recurrence",
    [
        "Same SEV-3 × 3 in a month → auto-escalate to 'Problem' ticket",
        "If it's happened twice → automate the fix, update runbook",
        "The 50% Rule: ≤50% time on toil; ≥50% on engineering",
        "Error budget >50% consumed → pause features for 1 sprint",
        "Introduce chaos engineering once SLOs & alerting are stable",
    ],
    n=11
)

# ── 12. Feedback Loop Infographic ────────────────────────────────────
infographic_slide(
    "The Reliability Loop — Detect → Respond → Learn → Prevent",
    IMG[13],
    caption="Every incident feeds back into prevention — continuous improvement cycle",
    n=12
)

# ── 13. Section: Database & SRE Philosophy ───────────────────────────
section_slide("Part 4", "Database Performance & SRE Philosophy")

# ── 14. Database Health & Query Strategy ─────────────────────────────
img_bullets_slide(
    "Database Health — SQL & NoSQL Monitoring",
    IMG[12],
    [
        "Postgres: pg_stat_activity, pg_stat_statements, replication lag, lock waits",
        "MongoDB: serverStatus(), currentOp(), Atlas Profiler (threshold: 100ms)",
        "Slow query threshold: 100ms; EXPLAIN ANALYZE for Postgres, .explain() for Mongo",
        "Key alerts: query p95 >200ms, connections >80% max, replication lag >5s",
        "Index strategy: highest-cardinality first; detect unused via idx_scan=0",
        "Connection pooling: PgBouncer (txn mode), 10–25 connections per instance",
    ],
    side="left", n=14
)

# ── 15. DB-as-Code + SRE Principles ─────────────────────────────────
two_col_slide(
    "DB-as-Code & SRE Core Principles",
    "Database Engineering",
    [
        "All schema changes in migration files (Flyway/Liquibase) → VCS",
        "No manual ALTER TABLE in production — ever",
        "PRs with DB changes must include EXPLAIN plan; SRE reviews",
        "Zero-downtime migrations via gh-ost (large MySQL tables)",
        "RDS Proxy for connection multiplexing",
    ],
    "SRE Philosophy",
    [
        "SLOs = common language between business, product & engineering",
        "Toil = manual, repetitive work lacking enduring value — target ≤50%",
        "Don't hire 'SRE' and treat as SysAdmin — empower code changes",
        "Automation ensures team scales without linear headcount",
        "Reliability is a feature that must be actively maintained",
    ],
    n=15
)

# ── 16. Reliability Roadmap Infographic ──────────────────────────────
infographic_slide(
    "Reliability Engineering Roadmap — Build In → Measure → Automate",
    IMG[10],
    caption="Build Reliability In → Measure & Improve → Automate Everything",
    n=16
)

# ── 17. Strategic Summary ────────────────────────────────────────────
content_slide(
    "Strategic Summary — Key Takeaways",
    [
        "Production-grade SaaS on AWS ECS rests on three pillars: metrics, logs, and traces",
        "30-60-90 day roadmap shifts the org from reactive to governed reliability",
        "Performance is a baseline we actively defend — load testing, tracing, canary deploys",
        "L3 support transforms from burnout source to engineering roadmap feedback loop",
        "SQL + NoSQL managed with DBRE rigor — the data layer is the hardest to scale",
        "Strategy aims for managed risk, not perfection — enabling faster, safer innovation",
    ],
    n=17
)

# ── 18. Thank You ────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_bg(slide, BG_TITLE)
tf = add_textbox(slide, Inches(1.5), Inches(2.5), Inches(10.3), Inches(2.0))
tf.paragraphs[0].text = "Thank You"
tf.paragraphs[0].font.size = Pt(44)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = NAVY
tf.paragraphs[0].font.name = FONT_TITLE
tf.paragraphs[0].alignment = PP_ALIGN.CENTER
set_para(tf, "Questions & Discussion", size=24, color=CORAL, font=FONT_ACCENT,
         align=PP_ALIGN.CENTER, before=Pt(20))

# ── Save ──────────────────────────────────────────────────────────────
prs.save(OUTPUT)
print(f"\n✅ Presentation saved to: {OUTPUT}")
print(f"   Total slides: {len(prs.slides)}")
