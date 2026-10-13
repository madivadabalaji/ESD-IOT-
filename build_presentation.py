import os
import math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# -----------------------------------------------------------------------------
# PRESENTATION CONFIGURATION & COLOR PALETTE
# -----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(13.333)  # 16:9 Widescreen
prs.slide_height = Inches(7.5)

# Color Scheme: Navy Blue, White, Cyan, Slate Gray
COLOR_NAVY_DARK   = RGBColor(11, 27, 61)     # #0B1B3D
COLOR_NAVY_MED    = RGBColor(16, 42, 67)     # #102A43
COLOR_NAVY_LIGHT  = RGBColor(31, 58, 90)     # #1F3A5A
COLOR_CYAN_ACCENT = RGBColor(0, 180, 216)    # #00B4D8
COLOR_CYAN_BRIGHT = RGBColor(0, 229, 255)    # #00E5FF
COLOR_WHITE       = RGBColor(255, 255, 255)  # #FFFFFF
COLOR_SLATE_DARK  = RGBColor(36, 59, 83)     # #243B53
COLOR_SLATE_MUTED = RGBColor(98, 125, 152)   # #627D98
COLOR_CARD_BG     = RGBColor(240, 244, 248)  # #F0F4F8
COLOR_CARD_TINT   = RGBColor(235, 248, 255)  # #EBF8FF
COLOR_CARD_BORDER = RGBColor(203, 213, 225)  # #CBD5E1
COLOR_ALERT_RED   = RGBColor(197, 48, 48)    # #C53030
COLOR_SUCCESS_GRN = RGBColor(40, 167, 69)    # #28A745
COLOR_CODE_BG     = RGBColor(15, 23, 42)     # #0F172A

FONT_HEADING = "Segoe UI"
FONT_BODY    = "Segoe UI"
FONT_MONO    = "Consolas"

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS FOR SLIDE BUILDING
# -----------------------------------------------------------------------------
def create_base_slide():
    """Returns a blank slide with standard white canvas."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_WHITE
    bg.line.fill.background()
    return slide

def add_header(slide, title_text, subtitle_text, exp_tag="EXP 11 | ESD & IoT"):
    """Adds a standard navy header banner with cyan accent line and badge."""
    # Top banner background
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = COLOR_NAVY_DARK
    header_box.line.fill.background()

    # Cyan divider line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.15), Inches(13.333), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    line.line.fill.background()

    # Title & Subtitle text frame
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.12), Inches(8.8), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_title = tf.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(21)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_WHITE

    p_sub = tf.add_paragraph()
    p_sub.text = subtitle_text
    p_sub.font.name = FONT_BODY
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = COLOR_CYAN_BRIGHT
    p_sub.space_before = Pt(3)

    # Category Pill (Top Right)
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.6), Inches(0.2), Inches(3.2), Inches(0.38))
    pill.fill.solid()
    pill.fill.fore_color.rgb = COLOR_NAVY_MED
    pill.line.color.rgb = COLOR_CYAN_ACCENT
    pill.line.width = Pt(1.2)
    p_pill = pill.text_frame.paragraphs[0]
    p_pill.text = exp_tag
    p_pill.font.name = FONT_HEADING
    p_pill.font.size = Pt(9.5)
    p_pill.font.bold = True
    p_pill.font.color.rgb = COLOR_CYAN_BRIGHT
    p_pill.alignment = PP_ALIGN.CENTER

    # Simulation Badge Pill
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.6), Inches(0.65), Inches(3.2), Inches(0.36))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(28, 47, 74)
    badge.line.color.rgb = COLOR_SLATE_MUTED
    badge.line.width = Pt(1)
    p_badge = badge.text_frame.paragraphs[0]
    p_badge.text = "[SOFTWARE SIMULATION ONLY]"
    p_badge.font.name = FONT_HEADING
    p_badge.font.size = Pt(8.5)
    p_badge.font.bold = True
    p_badge.font.color.rgb = RGBColor(224, 252, 255)
    p_badge.alignment = PP_ALIGN.CENTER

    # Footer banner
    footer_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.15), Inches(13.333), Inches(0.35))
    footer_box.fill.solid()
    footer_box.fill.fore_color.rgb = COLOR_NAVY_DARK
    footer_box.line.fill.background()

    footer_tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.18), Inches(12.133), Inches(0.3))
    ftf = footer_tb.text_frame
    ftf.margin_left = ftf.margin_top = ftf.margin_right = ftf.margin_bottom = 0
    p_foot = ftf.paragraphs[0]
    p_foot.text = "ESD & IoT | Exp 11: Digital Tachometer with Velocity Estimation  •  Software Simulation Model  •  Engineering Review"
    p_foot.font.name = FONT_BODY
    p_foot.font.size = Pt(8.5)
    p_foot.font.color.rgb = RGBColor(176, 196, 222)

def add_card(slide, left, top, width, height, title="", title_bg=None, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
    """Draws a rounded card container with optional header band."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.4)

    if title:
        header_h = Inches(0.46)
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, header_h)
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = title_bg if title_bg else COLOR_NAVY_DARK
        hdr.line.fill.background()

        hdr_tb = slide.shapes.add_textbox(left + Inches(0.15), top + Inches(0.04), width - Inches(0.3), Inches(0.38))
        htf = hdr_tb.text_frame
        htf.margin_left = htf.margin_top = htf.margin_right = htf.margin_bottom = 0
        p = htf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.LEFT
    return card

# =============================================================================
# SLIDE 1: TITLE SLIDE
# =============================================================================
def build_slide_1():
    slide = create_base_slide()

    # Dark high-tech full background for title
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_NAVY_DARK
    bg.line.fill.background()

    # Glowing top cyan bar
    top_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.1))
    top_line.fill.solid()
    top_line.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    top_line.line.fill.background()

    # Title Badge
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.55), Inches(5.8), Inches(0.38))
    pill.fill.solid()
    pill.fill.fore_color.rgb = COLOR_NAVY_MED
    pill.line.color.rgb = COLOR_CYAN_ACCENT
    pill.line.width = Pt(1.2)
    p_pill = pill.text_frame.paragraphs[0]
    p_pill.text = "EMBEDDED SYSTEMS DESIGN & IOT  |  LABORATORY EXPERIMENT 11"
    p_pill.font.name = FONT_HEADING
    p_pill.font.size = Pt(9.5)
    p_pill.font.bold = True
    p_pill.font.color.rgb = COLOR_CYAN_BRIGHT
    p_pill.alignment = PP_ALIGN.CENTER

    # Project Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.05), Inches(11.7), Inches(1.3))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p1 = tf_title.paragraphs[0]
    p1.text = "Digital Tachometer with Velocity Estimation"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_WHITE

    p2 = tf_title.add_paragraph()
    p2.text = "Discrete Pulse Frequency Counting & Real-Time Angular Velocity Estimation Model"
    p2.font.name = FONT_BODY
    p2.font.size = Pt(14)
    p2.font.color.rgb = COLOR_CYAN_BRIGHT
    p2.space_before = Pt(4)

    # Simulation Notice Banner
    sim_banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.45), Inches(11.733), Inches(0.42))
    sim_banner.fill.solid()
    sim_banner.fill.fore_color.rgb = RGBColor(16, 42, 67)
    sim_banner.line.color.rgb = RGBColor(229, 62, 62)
    sim_banner.line.width = Pt(1.5)
    p_sim = sim_banner.text_frame.paragraphs[0]
    p_sim.text = "STATUS: SOFTWARE SIMULATION ONLY  —  Mathematical Algorithm & STM32 Controller Logic Validated in Simulation Environment"
    p_sim.font.name = FONT_HEADING
    p_sim.font.size = Pt(9.5)
    p_sim.font.bold = True
    p_sim.font.color.rgb = RGBColor(254, 178, 178)
    p_sim.alignment = PP_ALIGN.CENTER

    # Insert Concept Graphic (assets/slide1_concept_art.png)
    if os.path.exists("assets/slide1_concept_art.png"):
        slide.shapes.add_picture("assets/slide1_concept_art.png", Inches(0.8), Inches(3.05), Inches(11.733), Inches(2.35))

    # Details Cards (Student Info, Faculty Info, College Info)
    card_l = Inches(0.8)
    card_w = Inches(3.7)
    card_h = Inches(1.35)
    card_y = Inches(5.6)

    # Card 1: Student Information
    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_l, card_y, card_w, card_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = COLOR_NAVY_MED
    c1.line.color.rgb = COLOR_SLATE_MUTED
    tf1 = c1.text_frame
    tf1.margin_left = Inches(0.15)
    tf1.margin_top = Inches(0.1)
    p = tf1.paragraphs[0]
    p.text = "STUDENT CREDENTIALS"
    p.font.name = FONT_HEADING
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_BRIGHT
    p_sub1 = tf1.add_paragraph()
    p_sub1.text = "• Student Name: Subhashree\n• Roll Number: [Roll No. / Reg. No.]\n• Degree: B.Tech / B.E. (ESD & IoT)"
    p_sub1.font.name = FONT_BODY
    p_sub1.font.size = Pt(9)
    p_sub1.font.color.rgb = COLOR_WHITE
    p_sub1.space_before = Pt(3)

    # Card 2: Faculty & Evaluation
    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_l + Inches(4.0), card_y, card_w, card_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = COLOR_NAVY_MED
    c2.line.color.rgb = COLOR_SLATE_MUTED
    tf2 = c2.text_frame
    tf2.margin_left = Inches(0.15)
    tf2.margin_top = Inches(0.1)
    p = tf2.paragraphs[0]
    p.text = "FACULTY & EVALUATION"
    p.font.name = FONT_HEADING
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_BRIGHT
    p_sub2 = tf2.add_paragraph()
    p_sub2.text = "• Evaluator: Course Faculty\n• Lab: Embedded Systems Design & IoT Lab\n• Academic Term: 2026 Evaluation"
    p_sub2.font.name = FONT_BODY
    p_sub2.font.size = Pt(9)
    p_sub2.font.color.rgb = COLOR_WHITE
    p_sub2.space_before = Pt(3)

    # Card 3: Department & Institution
    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_l + Inches(8.0), card_y, card_w, card_h)
    c3.fill.solid()
    c3.fill.fore_color.rgb = COLOR_NAVY_MED
    c3.line.color.rgb = COLOR_SLATE_MUTED
    tf3 = c3.text_frame
    tf3.margin_left = Inches(0.15)
    tf3.margin_top = Inches(0.1)
    p = tf3.paragraphs[0]
    p.text = "INSTITUTION DETAILS"
    p.font.name = FONT_HEADING
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_BRIGHT
    p_sub3 = tf3.add_paragraph()
    p_sub3.text = "• Dept: Electronics & Communication Engg.\n• Institution: Engineering College\n• Experiment Category: STM32 Systems"
    p_sub3.font.name = FONT_BODY
    p_sub3.font.size = Pt(9)
    p_sub3.font.color.rgb = COLOR_WHITE
    p_sub3.space_before = Pt(3)

    slide.notes_slide.notes_text_frame.text = (
        "Good morning respected faculty and project evaluators. Today I am presenting Experiment 11 for "
        "the Embedded Systems Design and IoT course, titled 'Digital Tachometer with Velocity Estimation'. "
        "In this project, we explore the automated measurement of rotational speed using rotary encoder pulse trains "
        "and an STM32 microcontroller model. Please note from the outset that this implementation has been "
        "conducted entirely in software simulation. No physical STM32 board, rotary encoder, or DC motor was "
        "assembled or physically tested. All pulse waveforms, graphs, and performance parameters presented in this review "
        "are simulated and mathematically verified results to validate the algorithm before hardware deployment."
    )

# =============================================================================
# SLIDE 2: ABSTRACT
# =============================================================================
def build_slide_2():
    slide = create_base_slide()
    add_header(slide, "Abstract & Executive Summary", "High-Precision Rotational Speed & Velocity Estimation via STM32 Pulse Counting Model")

    card_w = Inches(3.8)
    card_h = Inches(4.5)
    card_y = Inches(1.45)

    # Card 1: Core Sensing Principle
    add_card(slide, Inches(0.6), card_y, card_w, card_h, "1. Rotational Pulse Transduction", COLOR_NAVY_DARK)
    tb1 = slide.shapes.add_textbox(Inches(0.75), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "Digital Transduction Concept:"
    p.font.bold = True
    p.font.size = Pt(10.5)
    p.font.color.rgb = COLOR_NAVY_DARK

    bullets1 = [
        "Rotary encoders convert mechanical shaft angle into a discrete train of electrical square pulses.",
        "Unlike analog tachogenerators, digital pulse counting eliminates brush friction, drift, and calibration decay.",
        "The frequency of generated pulses directly corresponds to rotational velocity: f = (RPM × PPR) / 60.",
        "Provides precise digital resolution dependent on encoder pulses per revolution (PPR)."
    ]
    for b in bullets1:
        p_b = tf1.add_paragraph()
        p_b.text = "• " + b
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(5)

    # Card 2: STM32 Embedded Model
    add_card(slide, Inches(4.75), card_y, card_w, card_h, "2. STM32 Controller Architecture", COLOR_NAVY_LIGHT)
    tb2 = slide.shapes.add_textbox(Inches(4.9), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Timer & Processing Architecture:"
    p.font.bold = True
    p.font.size = Pt(10.5)
    p.font.color.rgb = COLOR_NAVY_DARK

    bullets2 = [
        "Models STM32 hardware Timer (TIMx) running in Input Capture / External Counter mode.",
        "Decoupled timekeeping: A dedicated 1.0s periodic timer interrupt dictates the sampling interval (Δt = 1.0 s).",
        "Edge accumulation: Counts incoming pulses (N) non-blockingly during the interval.",
        "Firmware engine computes instantaneous RPM and derives SI Angular Velocity (ω in rad/s).",
        "Division-by-zero protection ensures robust handling of stopped motor conditions."
    ]
    for b in bullets2:
        p_b = tf2.add_paragraph()
        p_b.text = "• " + b
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(5)

    # Card 3: Simulation Validation
    add_card(slide, Inches(8.9), card_y, card_w, card_h, "3. Software Simulation Scope", COLOR_NAVY_MED)
    tb3 = slide.shapes.add_textbox(Inches(9.05), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "Verification & Laboratory Scope:"
    p.font.bold = True
    p.font.size = Pt(10.5)
    p.font.color.rgb = COLOR_NAVY_DARK

    bullets3 = [
        "Implementation Status: Strictly software simulation using discrete modeling & Python validation.",
        "Simulates multi-speed motor regimes: 0 RPM (halt), 300 RPM, 600 RPM, and 900 RPM.",
        "Verifies exact mathematical concordance between generated pulses and calculated velocity.",
        "Provides pre-silicon verification of algorithms, quantization boundaries, and timing constraints prior to PCB fabrication."
    ]
    for b in bullets3:
        p_b = tf3.add_paragraph()
        p_b.text = "• " + b
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(5)

    # Highlight Callout Banner at bottom
    callout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.12), Inches(12.1), Inches(0.85))
    callout.fill.solid()
    callout.fill.fore_color.rgb = COLOR_CARD_TINT
    callout.line.color.rgb = COLOR_CYAN_ACCENT
    callout.line.width = Pt(1.5)
    ctb = callout.text_frame
    ctb.margin_left = Inches(0.15)
    ctb.margin_top = Inches(0.08)
    p_c1 = ctb.paragraphs[0]
    p_c1.text = "KEY TAKEAWAY: Discrete Edge Counting Yields Drift-Free Velocity Estimation"
    p_c1.font.bold = True
    p_c1.font.size = Pt(10)
    p_c1.font.color.rgb = COLOR_NAVY_DARK
    p_c2 = ctb.add_paragraph()
    p_c2.text = "By employing an STM32 timer model to accumulate encoder pulses over a deterministic 1-second sampling window, the system calculates both rotational speed (RPM) and angular velocity (rad/s) with zero theoretical error under discrete simulation."
    p_c2.font.size = Pt(8.5)
    p_c2.font.color.rgb = COLOR_SLATE_DARK
    p_c2.space_before = Pt(2)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 2 outlines the abstract of Experiment 11. Our digital tachometer operates on the principle of converting "
        "continuous mechanical rotational movement into digital pulses using an incremental rotary encoder model. "
        "The core advantage of digital pulse counting over analog tachogenerators is that it is immune to brush wear, "
        "thermal drift, and calibration decay. The STM32 microcontroller is modeled using two hardware timer functions: "
        "Timer 2 acts as an asynchronous pulse accumulator, while Timer 3 triggers a deterministic 1.0-second interrupt. "
        "During each interrupt, the pulse count is latched, and the firmware calculates RPM and angular velocity in radians per second. "
        "Crucially, the entire system has been modeled and evaluated within software simulation to establish algorithmic "
        "soundness before physical hardware construction."
    )

# =============================================================================
# SLIDE 3: PROBLEM STATEMENT AND MOTIVATION
# =============================================================================
def build_slide_3():
    slide = create_base_slide()
    add_header(slide, "Problem Statement & Engineering Motivation", "The Critical Need for Real-Time Automated Speed Measurement in Motor Monitoring & Automation")

    col_w = Inches(5.9)
    col_h = Inches(4.5)
    top_y = Inches(1.45)

    # Left Column: Limitations
    add_card(slide, Inches(0.6), top_y, col_w, col_h, "Limitations of Conventional Rotational Measurement", COLOR_NAVY_DARK)
    tb_l = slide.shapes.add_textbox(Inches(0.8), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    pts_l = [
        ("Manual & Stroboscopic Inconvenience", "Handheld optical or contact tachometers require human intervention, direct line-of-sight, and manual trigger operation, making them unfeasible for continuous 24/7 automation."),
        ("Mechanical & Wear Vulnerabilities", "Analog DC tachogenerators rely on mechanical brushes and slip-rings that cause electrical sparking, mechanical drag, contact friction, and calibration drift over operating life."),
        ("Analog Noise & Conversion Latency", "Analog voltage signals require low-pass filtering, signal amplification, and ADC conversion, introducing signal latency and vulnerability to electromagnetic interference (EMI)."),
        ("Lack of Integrated SI Velocity Estimation", "Most rudimentary tachometers output only crude RPM indicators without deriving instantaneous angular velocity (rad/s) necessary for modern robotic and kinematic controls.")
    ]
    for idx, (head, body) in enumerate(pts_l):
        p_h = tf_l.paragraphs[0] if idx == 0 else tf_l.add_paragraph()
        p_h.text = f"• {head}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9.5)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(5) if idx > 0 else Pt(0)

        p_b = tf_l.add_paragraph()
        p_b.text = f"  {body}"
        p_b.font.size = Pt(8.5)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(1)

    # Right Column: Advantages
    add_card(slide, Inches(6.8), top_y, col_w, col_h, "Engineering Motivation & Simulation Value", COLOR_NAVY_LIGHT)
    tb_r = slide.shapes.add_textbox(Inches(7.0), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    pts_r = [
        ("Automated Continuous Monitoring", "Industrial motors (BLDC, induction, stepper) in factory lines, CNC spindles, and EVs demand autonomous real-time speed monitoring for closed-loop PID control and safety interlocks."),
        ("High-Fidelity Discrete Transduction", "Optical rotary encoders generate noise-immune TTL digital pulses, enabling direct microcontroller timer hardware capture with nanosecond edge precision."),
        ("Dual Metric Estimation (RPM & ω)", "The embedded algorithm calculates both user-intuitive mechanical speed (RPM) and rigorous physical dynamics metric (angular velocity ω in rad/s)."),
        ("Value of Pre-Implementation Simulation", "Software simulation allows engineers to verify edge-counting logic, evaluate interrupt timing jitter, assess division-by-zero risks, and test speed steps before investing in hardware.")
    ]
    for idx, (head, body) in enumerate(pts_r):
        p_h = tf_r.paragraphs[0] if idx == 0 else tf_r.add_paragraph()
        p_h.text = f"✔ {head}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9.5)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(5) if idx > 0 else Pt(0)

        p_b = tf_r.add_paragraph()
        p_b.text = f"  {body}"
        p_b.font.size = Pt(8.5)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(1)

    # Bottom Context Banner
    ctx = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.12), Inches(12.1), Inches(0.85))
    ctx.fill.solid()
    ctx.fill.fore_color.rgb = COLOR_NAVY_DARK
    ctx.line.color.rgb = COLOR_CYAN_ACCENT
    ctx.line.width = Pt(1)
    tf_ctx = ctx.text_frame
    tf_ctx.margin_top = Inches(0.15)
    p_ctx = tf_ctx.paragraphs[0]
    p_ctx.text = "INDUSTRIAL APPLICATION DOMAIN: Motor Speed Monitoring  •  Robotics Joint Actuators  •  Automated Conveyors  •  Pump Telemetry"
    p_ctx.font.name = FONT_HEADING
    p_ctx.font.size = Pt(9.5)
    p_ctx.font.bold = True
    p_ctx.font.color.rgb = COLOR_CYAN_BRIGHT
    p_ctx.alignment = PP_ALIGN.CENTER

    slide.notes_slide.notes_text_frame.text = (
        "Slide 3 covers the problem statement and motivation. In modern industrial automation, continuous rotational speed "
        "measurement is essential. Traditional manual stroboscopes or contact tachometers cannot be integrated into "
        "automated feedback loops. Analog DC tachogenerators suffer from brush wear, mechanical friction, and thermal drift. "
        "A digital tachometer solves this by capturing optical encoder pulses directly through microcontroller timers. "
        "Furthermore, by simulating this system first, we avoid physical risks like motor runaway or counter overflow, "
        "enabling thorough algorithmic validation before committing to physical prototyping."
    )

# =============================================================================
# SLIDE 4: AIM AND OBJECTIVES
# =============================================================================
def build_slide_4():
    slide = create_base_slide()
    add_header(slide, "Aim & Technical Objectives", "Formal Scope of Laboratory Experiment 11 — Digital Tachometer Simulation")

    # Aim Banner Box
    aim_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.35), Inches(12.1), Inches(1.15))
    aim_box.fill.solid()
    aim_box.fill.fore_color.rgb = COLOR_NAVY_DARK
    aim_box.line.color.rgb = COLOR_CYAN_ACCENT
    aim_box.line.width = Pt(1.5)
    tf_aim = aim_box.text_frame
    tf_aim.margin_left = Inches(0.2)
    tf_aim.margin_top = Inches(0.12)
    p_aim_lbl = tf_aim.paragraphs[0]
    p_aim_lbl.text = "PRIMARY PROJECT AIM:"
    p_aim_lbl.font.name = FONT_HEADING
    p_aim_lbl.font.size = Pt(11)
    p_aim_lbl.font.bold = True
    p_aim_lbl.font.color.rgb = COLOR_CYAN_BRIGHT

    p_aim_txt = tf_aim.add_paragraph()
    p_aim_txt.text = (
        "To design, model, and simulate a digital tachometer system capable of estimating motor rotational speed (RPM) "
        "and physical angular velocity (ω in rad/s) from incremental rotary encoder pulse data using an STM32 "
        "microcontroller-based architecture model under discrete-time sampling."
    )
    p_aim_txt.font.name = FONT_BODY
    p_aim_txt.font.size = Pt(10)
    p_aim_txt.font.color.rgb = COLOR_WHITE
    p_aim_txt.space_before = Pt(3)

    # 6 Numbered Objectives Grid
    objs = [
        ("Objective 1: Pulse Generation Model",
         "Generate and model synthetic digital pulse trains corresponding to rotational shaft motion across varying motor velocities with a defined 20 PPR resolution."),
        ("Objective 2: Fixed-Interval Pulse Counting",
         "Implement an STM32 hardware timer/counter model that deterministically counts rising pulse edges over a specified sampling window (Δt = 1.0 second)."),
        ("Objective 3: RPM Calculation Algorithm",
         "Develop and verify the discrete mathematical conversion algorithm to calculate Revolutions Per Minute (RPM) from accumulated pulse count N and PPR."),
        ("Objective 4: Angular Velocity Estimation",
         "Derive instantaneous physical angular velocity (ω) in standard SI units (radians per second, rad/s) using exact conversion formulations (2π·RPM/60)."),
        ("Objective 5: Virtual Telemetry Interface",
         "Format and display simulated measurement outputs via a serial-terminal-style dashboard log and time-series telemetry plots for real-time monitoring."),
        ("Objective 6: Multi-Speed Model Benchmarking",
         "Evaluate and validate the algorithmic calculation pipeline across diverse simulated motor operating points: 0 RPM (stopped), 300 RPM, 600 RPM, and 900 RPM.")
    ]

    card_w = Inches(3.85)
    card_h = Inches(2.05)
    xs = [Inches(0.6), Inches(4.75), Inches(8.9)]
    ys = [Inches(2.65), Inches(4.85)]

    for idx, (otitle, odesc) in enumerate(objs):
        row = idx // 3
        col = idx % 3
        cx = xs[col]
        cy = ys[row]

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1.2)

        # Header pill inside card
        hpill = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, cy, card_w, Inches(0.42))
        hpill.fill.solid()
        hpill.fill.fore_color.rgb = COLOR_NAVY_LIGHT if row == 0 else COLOR_NAVY_DARK
        hpill.line.fill.background()
        tf_h = hpill.text_frame
        tf_h.margin_top = Inches(0.04)
        p_hp = tf_h.paragraphs[0]
        p_hp.text = otitle
        p_hp.font.name = FONT_HEADING
        p_hp.font.size = Pt(9.5)
        p_hp.font.bold = True
        p_hp.font.color.rgb = COLOR_WHITE
        p_hp.alignment = PP_ALIGN.LEFT

        # Text inside card
        tb_o = slide.shapes.add_textbox(cx + Inches(0.12), cy + Inches(0.48), card_w - Inches(0.24), card_h - Inches(0.55))
        tf_o = tb_o.text_frame
        tf_o.word_wrap = True
        tf_o.margin_left = tf_o.margin_top = tf_o.margin_right = tf_o.margin_bottom = 0
        p_desc = tf_o.paragraphs[0]
        p_desc.text = odesc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(8.5)
        p_desc.font.color.rgb = COLOR_SLATE_DARK

    slide.notes_slide.notes_text_frame.text = (
        "Slide 4 specifies our formal project aim and six measurable technical objectives. Our overarching aim is to "
        "design and simulate an STM32-based digital tachometer that estimates rotational speed and angular velocity from "
        "encoder pulse data. The six objectives systematically address: generating encoder pulses, setting up deterministic "
        "sampling intervals, computing RPM, calculating angular velocity in rad/s, presenting telemetry via a virtual "
        "interface, and finally benchmarking the algorithm across representative test speeds including 0, 300, 600, and 900 RPM."
    )

# =============================================================================
# SLIDE 5: SYSTEM ARCHITECTURE
# =============================================================================
def build_slide_5():
    slide = create_base_slide()
    add_header(slide, "System Architecture & Signal Flow", "Modular Software Simulation Pipeline: Transduction, Timing, Calculation, & Telemetry")

    if os.path.exists("assets/slide5_architecture_diagram.png"):
        slide.shapes.add_picture("assets/slide5_architecture_diagram.png", Inches(0.6), Inches(1.35), Inches(12.1), Inches(2.9))

    card_w = Inches(3.85)
    card_h = Inches(2.55)
    card_y = Inches(4.45)
    xs = [Inches(0.6), Inches(4.75), Inches(8.9)]

    # Stage 1
    add_card(slide, xs[0], card_y, card_w, card_h, "Stage A: Motion & Transduction", COLOR_NAVY_DARK)
    tb1 = slide.shapes.add_textbox(xs[0] + Inches(0.15), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    b1 = [
        "Simulated Motor: Software speed generator producing angular rotation (RPM).",
        "Rotary Encoder Model: Optical disc modeled with 20 pulses per revolution (PPR = 20).",
        "Pulse Frequency Generation: Pulse train frequency is dynamically synthesized as f = (RPM × PPR) / 60 Hz."
    ]
    for idx, b in enumerate(b1):
        p = tf1.paragraphs[0] if idx == 0 else tf1.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_SLATE_DARK
        p.space_before = Pt(4) if idx > 0 else Pt(0)

    # Stage 2
    add_card(slide, xs[1], card_y, card_w, card_h, "Stage B: STM32 Timer / Counter Core", COLOR_NAVY_LIGHT)
    tb2 = slide.shapes.add_textbox(xs[1] + Inches(0.15), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    b2 = [
        "TIM2 Asynchronous Counter: Accumulates incoming pulse edges without CPU polling overhead.",
        "TIM3 Periodic Sampling Timer: Triggers an interrupt service routine (ISR) precisely every Δt = 1.0 second.",
        "Atomic Latching: In the ISR, pulse register count N is latched and the hardware counter is reset to zero."
    ]
    for idx, b in enumerate(b2):
        p = tf2.paragraphs[0] if idx == 0 else tf2.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_SLATE_DARK
        p.space_before = Pt(4) if idx > 0 else Pt(0)

    # Stage 3
    add_card(slide, xs[2], card_y, card_w, card_h, "Stage C: Velocity Engine & Display", COLOR_NAVY_MED)
    tb3 = slide.shapes.add_textbox(xs[2] + Inches(0.15), card_y + Inches(0.55), card_w - Inches(0.3), card_h - Inches(0.65))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    b3 = [
        "RPM Calculation: Computes RPM = (N / (PPR × Δt)) × 60 with division-by-zero protection.",
        "Angular Velocity: Computes SI velocity ω = 2π × RPM / 60 [rad/s].",
        "Simulated Telemetry: Formats output packets for virtual serial terminal and Matplotlib dashboard visualization."
    ]
    for idx, b in enumerate(b3):
        p = tf3.paragraphs[0] if idx == 0 else tf3.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_SLATE_DARK
        p.space_before = Pt(4) if idx > 0 else Pt(0)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 5 presents the complete system architecture and signal flow. The simulation pipeline consists of three "
        "interconnected stages. In Stage A, the simulated motor rotation drives the incremental rotary encoder model, "
        "producing square pulses at a frequency directly proportional to rotational speed. In Stage B, the STM32 controller "
        "leverages dual hardware timers: Timer 2 counts external rising edges, while Timer 3 generates a deterministic "
        "periodic interrupt every 1.0 second. This hardware timer decoupling prevents CPU blocking. In Stage C, "
        "the firmware calculates rotational speed in RPM and angular velocity in radians per second, forwarding data "
        "to the simulated serial terminal and dashboard."
    )

# =============================================================================
# SLIDE 6: HARDWARE AND SOFTWARE REQUIREMENTS
# =============================================================================
def build_slide_6():
    slide = create_base_slide()
    add_header(slide, "Hardware & Software Requirements", "Clear Separation Between Proposed Physical Hardware and Actual Simulation Tools")

    col_w = Inches(5.9)
    col_h = Inches(5.45)
    top_y = Inches(1.45)

    # Column 1: Proposed Hardware
    add_card(slide, Inches(0.6), top_y, col_w, col_h, "PROPOSED HARDWARE (FUTURE IMPLEMENTATION)", COLOR_NAVY_DARK)
    tb_h = slide.shapes.add_textbox(Inches(0.8), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True

    p_warn = tf_h.paragraphs[0]
    p_warn.text = "⚠️ STATUS: PROPOSED DESIGN ONLY (NOT PHYSICALLY ASSEMBLED)"
    p_warn.font.bold = True
    p_warn.font.size = Pt(9.5)
    p_warn.font.color.rgb = COLOR_ALERT_RED

    hw_items = [
        ("STM32 Development Board", "Target: STM32F401RE Nucleo-64 (ARM Cortex-M4 @ 84MHz, 3.3V logic, multiple 16/32-bit hardware timers with input capture capabilities)."),
        ("Incremental Rotary Encoder", "Target: Optical slotted disc encoder (20 PPR disc with LM393 comparator) or industrial quadrature optical encoder (360/600 CPR)."),
        ("DC Motor & Rotating Shaft", "Target: 12V DC brushed/brushless geared motor with metal output shaft, mounting bracket, and mechanical coupling."),
        ("Motor Driver & Power Supply", "Target: L298N dual H-bridge motor driver module with external 12V / 2A regulated DC power adapter."),
        ("Display & Signal Interconnects", "Target: 16×2 character LCD (I2C interface) or 0.96-inch SSD1306 OLED display, breadboard, optocoupler isolation board, and jumper wires.")
    ]
    for idx, (title, desc) in enumerate(hw_items):
        p_t = tf_h.add_paragraph()
        p_t.text = f"• {title}:"
        p_t.font.bold = True
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = COLOR_NAVY_DARK
        p_t.space_before = Pt(5)

        p_d = tf_h.add_paragraph()
        p_d.text = f"  {desc}"
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_SLATE_DARK
        p_d.space_before = Pt(1)

    # Column 2: Simulation Tools
    add_card(slide, Inches(6.8), top_y, col_w, col_h, "SIMULATION TOOLS & ENVIRONMENT (ACTUALLY USED)", COLOR_NAVY_LIGHT)
    tb_s = slide.shapes.add_textbox(Inches(7.0), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_s = tb_s.text_frame
    tf_s.word_wrap = True

    p_succ = tf_s.paragraphs[0]
    p_succ.text = "✔ STATUS: SOFTWARE ENVIRONMENT USED FOR MODELING & EVALUATION"
    p_succ.font.bold = True
    p_succ.font.size = Pt(9.5)
    p_succ.font.color.rgb = COLOR_SUCCESS_GRN

    sw_items = [
        ("STM32CubeIDE / STM32CubeMX", "Used for architectural peripheral planning: configuring TIM2 external clock input, TIM3 1Hz periodic interrupt prescalers, and HAL C firmware structuring."),
        ("Software Simulation Environment", "Executed discrete-time pulse counting and state-machine simulation modeling STM32 timer register capture and update events."),
        ("Python 3.14 & NumPy", "Used for numerical pulse train synthesis, discrete sampling emulation, and verifying mathematical calculations without hardware noise."),
        ("Matplotlib 3.10 Visualization", "Used for plotting digital pulse waveforms, dynamic step response curves (RPM vs. time), and angular velocity tracking graphs."),
        ("Simulated Serial Dashboard", "Virtual UART serial-monitor-style telemetry logging capturing timestamps, pulse counts (N), RPM, and angular velocity (ω).")
    ]
    for idx, (title, desc) in enumerate(sw_items):
        p_t = tf_s.add_paragraph()
        p_t.text = f"✔ {title}:"
        p_t.font.bold = True
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = COLOR_NAVY_DARK
        p_t.space_before = Pt(5)

        p_d = tf_s.add_paragraph()
        p_d.text = f"  {desc}"
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_SLATE_DARK
        p_d.space_before = Pt(1)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 6 clearly distinguishes our proposed physical hardware from our software simulation environment. "
        "As highlighted in the project requirements, no physical hardware was assembled or tested. The left column lists "
        "the proposed physical components for future lab implementation, including an STM32F401RE Nucleo board, an optical encoder, "
        "a DC motor, an L298N driver, and an LCD. The right column details the software simulation tools actually utilized: "
        "STM32CubeIDE architectural models for timer configuration, Python 3.14 and NumPy for pulse modeling and discrete time "
        "simulation, Matplotlib for graphing simulated waveforms, and a virtual serial terminal for telemetry logging."
    )

# =============================================================================
# SLIDE 7: WORKING PRINCIPLE AND MATHEMATICAL EQUATIONS
# =============================================================================
def build_slide_7():
    slide = create_base_slide()
    add_header(slide, "Working Principle & Mathematical Formulations", "Mathematical Modeling of Discrete Pulse Counting, Sampling Window, & SI Velocity Conversion")

    col_w = Inches(5.9)
    col_h = Inches(5.45)
    top_y = Inches(1.45)

    # Left Column: Mathematical Equations & Variables
    add_card(slide, Inches(0.6), top_y, col_w, col_h, "Mathematical Formulation & Parameter Definitions", COLOR_NAVY_DARK)
    
    # Text box for parameters
    tb_m = slide.shapes.add_textbox(Inches(0.8), top_y + Inches(0.55), col_w - Inches(0.4), Inches(1.6))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    p_v = tf_m.paragraphs[0]
    p_v.text = "Mathematical Variables & Nomenclature:"
    p_v.font.bold = True
    p_v.font.size = Pt(10)
    p_v.font.color.rgb = COLOR_NAVY_DARK

    vars_list = [
        ("N", "Number of discrete pulses counted during sampling interval Δt [pulses]"),
        ("PPR", "Pulses Per Revolution (selected counting convention) [pulses/rev]"),
        ("Δt", "Sampling interval / time gating window [seconds] (Δt = 1.0 s)"),
        ("RPM", "Rotational Speed [Revolutions Per Minute]"),
        ("ω", "Angular Velocity [radians per second, rad/s] (SI unit)")
    ]
    for sym, desc in vars_list:
        p = tf_m.add_paragraph()
        p.text = f"• {sym} : {desc}"
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_SLATE_DARK
        p.space_before = Pt(2)

    # Equation Box 1: RPM
    eq1_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top_y + Inches(2.25), col_w - Inches(0.4), Inches(1.15))
    eq1_box.fill.solid()
    eq1_box.fill.fore_color.rgb = COLOR_CARD_TINT
    eq1_box.line.color.rgb = COLOR_CYAN_ACCENT
    eq1_box.line.width = Pt(1.5)
    tf_eq1 = eq1_box.text_frame
    tf_eq1.margin_left = Inches(0.12)
    tf_eq1.margin_top = Inches(0.08)
    p_e1_lbl = tf_eq1.paragraphs[0]
    p_e1_lbl.text = "EQUATION 1: ROTATIONAL SPEED (RPM)"
    p_e1_lbl.font.bold = True
    p_e1_lbl.font.size = Pt(9)
    p_e1_lbl.font.color.rgb = COLOR_NAVY_DARK
    p_e1_val = tf_eq1.add_paragraph()
    p_e1_val.text = "RPM = ( N / (PPR × Δt) ) × 60"
    p_e1_val.font.name = FONT_MONO
    p_e1_val.font.bold = True
    p_e1_val.font.size = Pt(12)
    p_e1_val.font.color.rgb = COLOR_NAVY_DARK
    p_e1_val.space_before = Pt(2)
    p_e1_sub = tf_eq1.add_paragraph()
    p_e1_sub.text = "Converts discrete pulse frequency into revolutions per minute."
    p_e1_sub.font.size = Pt(8)
    p_e1_sub.font.color.rgb = COLOR_SLATE_DARK

    # Equation Box 2: Angular Velocity
    eq2_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top_y + Inches(3.55), col_w - Inches(0.4), Inches(1.25))
    eq2_box.fill.solid()
    eq2_box.fill.fore_color.rgb = COLOR_CARD_TINT
    eq2_box.line.color.rgb = COLOR_CYAN_ACCENT
    eq2_box.line.width = Pt(1.5)
    tf_eq2 = eq2_box.text_frame
    tf_eq2.margin_left = Inches(0.12)
    tf_eq2.margin_top = Inches(0.08)
    p_e2_lbl = tf_eq2.paragraphs[0]
    p_e2_lbl.text = "EQUATION 2: ANGULAR VELOCITY (ω in rad/s)"
    p_e2_lbl.font.bold = True
    p_e2_lbl.font.size = Pt(9)
    p_e2_lbl.font.color.rgb = COLOR_NAVY_DARK
    p_e2_val = tf_eq2.add_paragraph()
    p_e2_val.text = "ω = (2π × RPM) / 60  =  (2π × N) / (PPR × Δt)\n   ≈  0.1047198 × RPM   [rad/s]"
    p_e2_val.font.name = FONT_MONO
    p_e2_val.font.bold = True
    p_e2_val.font.size = Pt(10)
    p_e2_val.font.color.rgb = COLOR_NAVY_DARK
    p_e2_val.space_before = Pt(2)
    p_e2_sub = tf_eq2.add_paragraph()
    p_e2_sub.text = "Standard SI metric directly applicable for kinematic motion analysis."
    p_e2_sub.font.size = Pt(8)
    p_e2_sub.font.color.rgb = COLOR_SLATE_DARK

    # Unit Conversion Pill
    uc_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top_y + Inches(4.9), col_w - Inches(0.4), Inches(0.45))
    uc_box.fill.solid()
    uc_box.fill.fore_color.rgb = COLOR_CARD_BG
    uc_box.line.color.rgb = COLOR_SLATE_MUTED
    tf_uc = uc_box.text_frame
    tf_uc.margin_top = Inches(0.08)
    p_uc = tf_uc.paragraphs[0]
    p_uc.text = "SI Unit Reference: 1 revolution = 2π radians (≈ 6.283185 rad), 1 min = 60 sec"
    p_uc.font.size = Pt(8)
    p_uc.font.bold = True
    p_uc.font.color.rgb = COLOR_NAVY_DARK
    p_uc.alignment = PP_ALIGN.CENTER

    # Right Column: Quadrature Decoding Conventions
    add_card(slide, Inches(6.8), top_y, col_w, col_h, "Quadrature Decoding & Counting Conventions", COLOR_NAVY_LIGHT)
    tb_q = slide.shapes.add_textbox(Inches(7.0), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_q = tb_q.text_frame
    tf_q.word_wrap = True

    p_qc = tf_q.paragraphs[0]
    p_qc.text = "Distinction: Encoder Cycles vs Counted Edges:"
    p_qc.font.bold = True
    p_qc.font.size = Pt(10)
    p_qc.font.color.rgb = COLOR_NAVY_DARK

    quad_modes = [
        ("1x Counting Mode (Single Channel A, Rising Edge)",
         "Counted edges per revolution = 1 × CPR. Only rising edges on channel A are accumulated. PPR in calculation equals optical physical line count (PPR = 20). Used in our primary simulation model."),
        ("2x Counting Mode (Channel A, Rising + Falling)",
         "Counted edges per revolution = 2 × CPR. Both rising and falling transitions of channel A are captured. The effective PPR in the formula must be doubled (2 × 20 = 40) to prevent 2× speed overestimation."),
        ("4x Counting Mode (Quadrature Mode: Channels A & B)",
         "Counted edges per revolution = 4 × CPR. All edges of both Channel A and Channel B are captured via STM32 Timer Encoder Mode. Effective PPR = 4 × CPR (80 edges). Divisor must match counted edges."),
        ("Double-Counting Prevention Rule",
         "CRITICAL: The PPR parameter in the denominator must always match the exact counted edge convention configured in the microcontroller timer to ensure mathematical consistency.")
    ]
    for idx, (title, desc) in enumerate(quad_modes):
        p_t = tf_q.add_paragraph()
        p_t.text = f"• {title}:"
        p_t.font.bold = True
        p_t.font.size = Pt(9)
        p_t.font.color.rgb = COLOR_NAVY_DARK
        p_t.space_before = Pt(5)

        p_d = tf_q.add_paragraph()
        p_d.text = f"  {desc}"
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_SLATE_DARK
        p_d.space_before = Pt(1)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 7 details the mathematical working principle and equations. In Equation 1, rotational speed in RPM is calculated "
        "as N divided by (PPR times delta t), multiplied by 60. In Equation 2, angular velocity omega is converted directly to "
        "radians per second by multiplying RPM by 2 pi over 60, which simplifies to approximately 0.10472 times RPM. "
        "On the right, we address a critical engineering detail: the difference between encoder cycles per revolution (CPR) "
        "and counted edges per revolution (PPR). In 1x mode, only rising edges of channel A are counted. In 2x mode, both edges "
        "are counted. In 4x quadrature mode, all transitions on both channels are counted. To avoid double-counting or speed "
        "inflation, the denominator PPR must precisely reflect the active edge-counting mode."
    )

# =============================================================================
# SLIDE 8: ALGORITHM AND FLOWCHART
# =============================================================================
def build_slide_8():
    slide = create_base_slide()
    add_header(slide, "Algorithm & Flowchart Logic", "Periodic Sampling Architecture, Pulse Edge Accumulation, & Zero-Speed Handling")

    if os.path.exists("assets/slide8_flowchart.png"):
        slide.shapes.add_picture("assets/slide8_flowchart.png", Inches(0.6), Inches(1.35), Inches(4.5), Inches(5.6))

    r_w = Inches(7.4)
    r_h = Inches(5.6)
    r_x = Inches(5.3)
    r_y = Inches(1.35)

    add_card(slide, r_x, r_y, r_w, r_h, "STM32 Firmware Execution Sequence & Exception Handling", COLOR_NAVY_DARK)
    tb_alg = slide.shapes.add_textbox(r_x + Inches(0.2), r_y + Inches(0.55), r_w - Inches(0.4), r_h - Inches(0.65))
    tf_alg = tb_alg.text_frame
    tf_alg.word_wrap = True

    steps = [
        ("Step 1: System Initialization & Peripheral Setup",
         "Initialize global registers: PPR = 20, sampling interval Δt = 1.0 s, pulse accumulator N = 0. Configure TIM2 in external clock mode (ETR/Input Capture) and TIM3 as periodic 1.0s interrupt timer."),
        ("Step 2: Non-Blocking Pulse Accumulation",
         "Encoder pulses trigger hardware timer TIM2 counter register directly without CPU intervention. This avoids interrupt latency at high rotational speeds."),
        ("Step 3: Deterministic Periodic Timer ISR (1.0s Window)",
         "TIM3 generates an update event interrupt precisely at t = 1.0 s. Inside the ISR: atomic read of TIM2 counter (N), reset TIM2 counter to 0, and raise computation flag."),
        ("Step 4: Zero-Speed / Motor Stopped Boundary Check",
         "If pulse count N == 0: Explicitly set RPM = 0.0 and ω = 0.00 rad/s. Prevents uninitialized states or division anomalies, accurately indicating motor halt."),
        ("Step 5: Mathematical Velocity Conversion",
         "If N > 0: Calculate RPM = (N × 60.0) / (PPR × Δt). Compute physical angular velocity ω = (2π × RPM) / 60.0 rad/s."),
        ("Step 6: Telemetry Transmission & Continuous Loop",
         "Format data payload into serial telemetry string: '[DATA] RPM: %0.2f | Omega: %0.2f rad/s | Pulses: %lu'. Transmit over UART to virtual dashboard. Re-arm TIM3 and repeat.")
    ]
    for idx, (title, desc) in enumerate(steps):
        p_t = tf_alg.paragraphs[0] if idx == 0 else tf_alg.add_paragraph()
        p_t.text = title
        p_t.font.bold = True
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = COLOR_NAVY_DARK
        p_t.space_before = Pt(5) if idx > 0 else Pt(0)

        p_d = tf_alg.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = COLOR_SLATE_DARK
        p_d.space_before = Pt(1)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 8 illustrates the complete algorithm and flowchart. The system uses a clean, non-blocking two-timer design. "
        "Timer 2 acts as a hardware pulse accumulator, while Timer 3 triggers an interrupt every 1.0 second. "
        "In Step 4, we explicitly handle the stopped-motor boundary condition: if pulse count N is zero, the algorithm "
        "assigns 0 RPM and 0 rad/s immediately without numerical instability. In Step 5, for positive pulse counts, "
        "the equations are evaluated and telemetry is dispatched via virtual UART in Step 6. This decoupled structure "
        "ensures deterministic timing and high measurement accuracy."
    )

# =============================================================================
# SLIDE 9: SIMULATION IMPLEMENTATION
# =============================================================================
def build_slide_9():
    slide = create_base_slide()
    add_header(slide, "Simulation Implementation & Code Architecture", "Simulation Workflow, Target STM32 C Firmware Model, and Python Verification Engine")

    # Top Workflow Banner
    wf_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.35), Inches(12.1), Inches(0.75))
    wf_box.fill.solid()
    wf_box.fill.fore_color.rgb = COLOR_NAVY_DARK
    wf_box.line.color.rgb = COLOR_CYAN_ACCENT
    wf_box.line.width = Pt(1)
    tf_wf = wf_box.text_frame
    tf_wf.margin_top = Inches(0.1)
    p_wfh = tf_wf.paragraphs[0]
    p_wfh.text = "7-STEP SIMULATION PIPELINE WORKFLOW:"
    p_wfh.font.bold = True
    p_wfh.font.size = Pt(9.5)
    p_wfh.font.color.rgb = COLOR_CYAN_BRIGHT
    p_wfs = tf_wf.add_paragraph()
    p_wfs.text = "1. Set PPR (20) & Δt (1.0s) → 2. Configure Motor Speed → 3. Synthesize Pulse Train → 4. Accumulate Pulses (N) → 5. Compute RPM & ω → 6. Compare vs Target → 7. Plot Graphs"
    p_wfs.font.size = Pt(9)
    p_wfs.font.color.rgb = COLOR_WHITE
    p_wfs.space_before = Pt(2)

    col_w = Inches(5.9)
    col_h = Inches(4.75)
    top_y = Inches(2.25)

    # Left: STM32 C Logic Model
    add_card(slide, Inches(0.6), top_y, col_w, col_h, "STM32 Target C Firmware Model (HAL-Based)", COLOR_NAVY_LIGHT)
    tb_c = slide.shapes.add_textbox(Inches(0.75), top_y + Inches(0.52), col_w - Inches(0.3), col_h - Inches(0.65))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True

    c_code = (
        "// STM32 Periodic Sampling Interrupt Handler\n"
        "#define ENCODER_PPR  20.0f\n"
        "#define SAMPLING_DT  1.0f\n"
        "#define PI           3.14159265f\n\n"
        "void HAL_TIM_PeriodElapsedCallback(TIM_HandleTypeDef *htim) {\n"
        "    if (htim->Instance == TIM3) { // 1.0s Periodic Event\n"
        "        uint32_t N = __HAL_TIM_GET_COUNTER(&htim2);\n"
        "        __HAL_TIM_SET_COUNTER(&htim2, 0);\n\n"
        "        float rpm = 0.0f, omega = 0.0f;\n"
        "        if (N > 0) {\n"
        "            rpm   = ((float)N * 60.0f) / (ENCODER_PPR * SAMPLING_DT);\n"
        "            omega = (2.0f * PI * rpm) / 60.0f;\n"
        "        }\n"
        "        // Send telemetry packet over virtual UART\n"
        "        printf(\"[SIM] N=%lu | RPM=%.2f | Omega=%.2f rad/s\\r\\n\",\n"
        "               N, rpm, omega);\n"
        "    }\n"
        "}"
    )
    p_code_c = tf_c.paragraphs[0]
    p_code_c.text = c_code
    p_code_c.font.name = FONT_MONO
    p_code_c.font.size = Pt(8.5)
    p_code_c.font.color.rgb = COLOR_NAVY_DARK

    # Right: Python Simulation Script
    add_card(slide, Inches(6.8), top_y, col_w, col_h, "Python Simulation Engine & Verification Model", COLOR_NAVY_DARK)
    tb_py = slide.shapes.add_textbox(Inches(6.95), top_y + Inches(0.52), col_w - Inches(0.3), col_h - Inches(0.65))
    tf_py = tb_py.text_frame
    tf_py.word_wrap = True

    py_code = (
        "# Discrete Tachometer Simulation Engine\n"
        "import numpy as np\n\n"
        "def simulate_tachometer(target_rpm, ppr=20, dt=1.0):\n"
        "    if target_rpm <= 0:\n"
        "        return 0, 0.0, 0.0\n"
        "    # Theoretical pulse frequency (Hz)\n"
        "    f_pulse = (target_rpm * ppr) / 60.0\n"
        "    # Counted integer pulses in window dt\n"
        "    pulse_count = int(round(f_pulse * dt))\n\n"
        "    # Reconstruct speed metrics\n"
        "    est_rpm = (pulse_count * 60.0) / (ppr * dt)\n"
        "    est_omega = (2.0 * np.pi * est_rpm) / 60.0\n"
        "    return pulse_count, est_rpm, est_omega\n\n"
        "# Verification across laboratory test conditions\n"
        "for speed in [0, 300, 600, 900]:\n"
        "    n, r, w = simulate_tachometer(speed)\n"
        "    print(f\"{speed} RPM -> N={n}, Calc_RPM={r}, Omega={w:.2f}\")"
    )
    p_code_py = tf_py.paragraphs[0]
    p_code_py.text = py_code
    p_code_py.font.name = FONT_MONO
    p_code_py.font.size = Pt(8.5)
    p_code_py.font.color.rgb = COLOR_NAVY_DARK

    slide.notes_slide.notes_text_frame.text = (
        "Slide 9 explains our simulation implementation. Across the top banner is the 7-step simulation workflow: "
        "we initialize parameters, apply simulated motor speeds, synthesize pulse trains, count pulses, calculate speed "
        "metrics, compare against configured speeds, and plot graphs. On the left is the target STM32 C code modeled "
        "in STM32CubeIDE. It resides in the HAL timer period elapsed callback. On the right is the Python verification "
        "script that emulates this discrete sampling logic. It proves that discrete pulse counts yield identical mathematical "
        "estimates without numerical divergence."
    )

# =============================================================================
# SLIDE 10: SIMULATED OUTPUT AND GRAPHS
# =============================================================================
def build_slide_10():
    slide = create_base_slide()
    add_header(slide, "Simulated Output & Performance Graphs", "Multi-Panel Graphical Telemetry & Illustrative Laboratory Verification Dataset")

    if os.path.exists("assets/slide10_simulation_plots.png"):
        slide.shapes.add_picture("assets/slide10_simulation_plots.png", Inches(0.6), Inches(1.35), Inches(6.8), Inches(5.6))

    r_w = Inches(5.1)
    r_h = Inches(5.6)
    r_x = Inches(7.6)
    r_y = Inches(1.35)

    add_card(slide, r_x, r_y, r_w, r_h, "Laboratory Test Benchmarking (Illustrative)", COLOR_NAVY_DARK)
    tb_tb = slide.shapes.add_textbox(r_x + Inches(0.15), r_y + Inches(0.52), r_w - Inches(0.3), Inches(1.2))
    tf_tb = tb_tb.text_frame
    tf_tb.word_wrap = True

    p_cfg_lbl = tf_tb.paragraphs[0]
    p_cfg_lbl.text = "Configured Simulation Test Parameters:"
    p_cfg_lbl.font.bold = True
    p_cfg_lbl.font.size = Pt(9.5)
    p_cfg_lbl.font.color.rgb = COLOR_NAVY_DARK

    cfgs = [
        "Encoder Resolution: PPR = 20 counted pulses/rev",
        "Sampling Interval: Δt = 1.0 second (fixed window)",
        "Pulse Counting: 1x rising-edge accumulation mode",
        "Test Speeds: 0 RPM, 300 RPM, 600 RPM, 900 RPM"
    ]
    for c in cfgs:
        p = tf_tb.add_paragraph()
        p.text = "• " + c
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_SLATE_DARK
        p.space_before = Pt(2)

    # Insert Native Table
    table_shape = slide.shapes.add_table(5, 5, r_x + Inches(0.1), r_y + Inches(1.85), r_w - Inches(0.2), Inches(2.2))
    table = table_shape.table
    table.columns[0].width = Inches(0.95)
    table.columns[1].width = Inches(0.95)
    table.columns[2].width = Inches(0.95)
    table.columns[3].width = Inches(1.0)
    table.columns[4].width = Inches(1.05)

    headers = ["Condition", "Input Speed", "Pulses (N)", "Calc. RPM", "Omega (ω)"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.margin_left = cell.margin_right = cell.margin_top = cell.margin_bottom = Pt(2)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = FONT_HEADING
        p.font.bold = True
        p.font.size = Pt(8)
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    data = [
        ["Halt", "0 RPM", "0 / s", "0.00 RPM", "0.00 rad/s"],
        ["Test 1", "300 RPM", "100 / s", "300.00 RPM", "31.42 rad/s"],
        ["Test 2", "600 RPM", "200 / s", "600.00 RPM", "62.83 rad/s"],
        ["Test 3", "900 RPM", "300 / s", "900.00 RPM", "94.25 rad/s"]
    ]
    for row_idx, row in enumerate(data):
        for col_idx, val in enumerate(row):
            cell = table.cell(row_idx + 1, col_idx)
            cell.margin_left = cell.margin_right = cell.margin_top = cell.margin_bottom = Pt(2)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_CARD_TINT if row_idx % 2 == 1 else COLOR_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_BODY
            p.font.size = Pt(8)
            p.font.color.rgb = COLOR_NAVY_DARK
            p.alignment = PP_ALIGN.CENTER

    # Mathematical Verification Callout Box below table
    callout_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, r_x + Inches(0.1), r_y + Inches(4.2), r_w - Inches(0.2), Inches(1.25))
    callout_box.fill.solid()
    callout_box.fill.fore_color.rgb = RGBColor(254, 243, 199)
    callout_box.line.color.rgb = RGBColor(217, 119, 6)
    callout_box.line.width = Pt(1.2)
    tf_co = callout_box.text_frame
    tf_co.margin_left = Inches(0.1)
    tf_co.margin_top = Inches(0.06)
    p_co_h = tf_co.paragraphs[0]
    p_co_h.text = "⚠️ SIMULATION DATA INTEGRITY NOTICE"
    p_co_h.font.bold = True
    p_co_h.font.size = Pt(8.5)
    p_co_h.font.color.rgb = RGBColor(180, 83, 9)

    p_co_t = tf_co.add_paragraph()
    p_co_t.text = (
        "All values presented above and in adjacent waveforms are mathematically calculated illustrative "
        "simulation outputs. No physical hardware testbed was operated. They verify that the firmware algorithm "
        "satisfies linear speed estimation: N = 100 → 300 RPM (31.42 rad/s), N = 200 → 600 RPM (62.83 rad/s), "
        "and N = 300 → 900 RPM (94.25 rad/s)."
    )
    p_co_t.font.size = Pt(7.5)
    p_co_t.font.color.rgb = RGBColor(120, 53, 15)
    p_co_t.space_before = Pt(2)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 10 presents the simulated outputs and waveforms. The four graphs on the left display: the simulated "
        "pulse trains at 100 Hz, 200 Hz, and 300 Hz; the dynamic RPM step response over time; the corresponding angular velocity "
        "in radians per second; and the pulse count linearity curve showing an exact R-squared of 1.0. "
        "The table on the right summarizes our standard test conditions: at 300 RPM, the system counts 100 pulses per second, "
        "yielding 300 RPM and 31.42 rad/s. At 600 RPM, it counts 200 pulses, yielding 62.83 rad/s. At 900 RPM, it counts 300 pulses, "
        "yielding 94.25 rad/s. As noted in the disclaimer banner, these results represent illustrative simulation data "
        "confirming algorithmic accuracy."
    )

# =============================================================================
# SLIDE 11: TESTING, RESULTS, APPLICATIONS, AND LIMITATIONS
# =============================================================================
def build_slide_11():
    slide = create_base_slide()
    add_header(slide, "Test Evaluation, Applications & Limitations", "Algorithmic Verification Benchmarks, Industrial Applications, & Engineering Constraints")

    col_w = Inches(3.85)
    col_h = Inches(5.45)
    top_y = Inches(1.45)
    xs = [Inches(0.6), Inches(4.75), Inches(8.9)]

    # Card 1: Test-Case Verification
    add_card(slide, xs[0], top_y, col_w, col_h, "1. Test-Case Verification", COLOR_NAVY_DARK)
    tb1 = slide.shapes.add_textbox(xs[0] + Inches(0.15), top_y + Inches(0.55), col_w - Inches(0.3), col_h - Inches(0.65))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "Reference Benchmark Cases:"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_NAVY_DARK

    tcases = [
        ("Case 0 (Motor Stopped):", "Input: 0 RPM  |  Pulses: 0\nOutput: 0.00 RPM, 0.00 rad/s\nStatus: Zero-pulse guard verified."),
        ("Case 1 (Low-Speed Test):", "Input: 300 RPM  |  Pulses: 100\nOutput: 300.00 RPM, 31.42 rad/s\nStatus: Exact 100% agreement."),
        ("Case 2 (Medium-Speed Test):", "Input: 600 RPM  |  Pulses: 200\nOutput: 600.00 RPM, 62.83 rad/s\nStatus: Exact 100% agreement."),
        ("Case 3 (High-Speed Test):", "Input: 900 RPM  |  Pulses: 300\nOutput: 900.00 RPM, 94.25 rad/s\nStatus: Exact 100% agreement."),
        ("Algorithm Linearity:", "Linearity coefficient R² = 1.000 across all discrete test intervals.")
    ]
    for h, b in tcases:
        p_h = tf1.add_paragraph()
        p_h.text = f"• {h}"
        p_h.font.bold = True
        p_h.font.size = Pt(9)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(4)

        p_b = tf1.add_paragraph()
        p_b.text = f"  {b}"
        p_b.font.size = Pt(8)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(1)

    # Card 2: Industrial Applications
    add_card(slide, xs[1], top_y, col_w, col_h, "2. Real-World Applications", COLOR_NAVY_LIGHT)
    tb2 = slide.shapes.add_textbox(xs[1] + Inches(0.15), top_y + Inches(0.55), col_w - Inches(0.3), col_h - Inches(0.65))
    tf2 = tb2.text_frame
    tf2.word_wrap = True

    p = tf2.paragraphs[0]
    p.text = "Target Engineering Domains:"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_NAVY_DARK

    apps = [
        ("Industrial Motor Speed Monitoring", "Closed-loop feedback for AC induction motors, BLDC drives, and CNC machine spindles to maintain steady velocity under load."),
        ("Robotics & Motion Control", "Wheel odometry for Autonomous Mobile Robots (AMRs) and joint angular velocity tracking for robotic manipulators."),
        ("Automated Conveyor Systems", "Ensuring uniform material transport speed in packaging and assembly lines; detecting belt jams or motor slippage."),
        ("Pumps, Turbines & Fans", "Condition monitoring and predictive maintenance in HVAC blowers, centrifugal pumps, and wind turbine shafts."),
        ("IoT Fleet Telemetry", "Streaming continuous rotational telemetry to cloud dashboards for anomaly detection and runtime logs.")
    ]
    for h, b in apps:
        p_h = tf2.add_paragraph()
        p_h.text = f"✔ {h}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(4)

        p_b = tf2.add_paragraph()
        p_b.text = f"  {b}"
        p_b.font.size = Pt(8)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(1)

    # Card 3: Limitations of Simulation
    add_card(slide, xs[2], top_y, col_w, col_h, "3. Technical Limitations", COLOR_NAVY_MED)
    tb3 = slide.shapes.add_textbox(xs[2] + Inches(0.15), top_y + Inches(0.55), col_w - Inches(0.3), col_h - Inches(0.65))
    tf3 = tb3.text_frame
    tf3.word_wrap = True

    p = tf3.paragraphs[0]
    p.text = "Critical Engineering Constraints:"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_NAVY_DARK

    lims = [
        ("No Physical Hardware Validation", "Simulation assumes ideal square pulses without optical sensor misalignment, mechanical shaft vibration, or back-EMF noise."),
        ("Measurement Granularity (Quantization)", "With PPR = 20 and Δt = 1.0s, the minimum detectable speed change is ΔRPM = 60/(20×1) = 3.0 RPM. Speeds can only be resolved in 3 RPM steps."),
        ("Sampling Window Trade-off", "A 1.0-second sampling window introduces latency during rapid motor acceleration or deceleration transients."),
        ("Pulse Loss Vulnerability", "Electrical EMI noise or optical dust in physical environments can cause missed pulses or false extra edge triggers.")
    ]
    for h, b in lims:
        p_h = tf3.add_paragraph()
        p_h.text = f"⚠️ {h}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(4)

        p_b = tf3.add_paragraph()
        p_b.text = f"  {b}"
        p_b.font.size = Pt(8)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(1)

    slide.notes_slide.notes_text_frame.text = (
        "Slide 11 evaluates our testing verification, target industrial applications, and practical engineering limitations. "
        "The test benchmarks on the left confirm that the algorithm reproduces zero error across ideal conditions, including "
        "the 0 RPM zero-guard condition. The middle card highlights key industrial applications, from industrial motor drives "
        "and robotics odometry to conveyor lines and HVAC pumps. On the right, we objectively address engineering limitations: "
        "simulation does not account for optical dust, mechanical vibration, or electrical noise. Furthermore, our 20-PPR encoder "
        "with a 1-second sampling window imposes a 3-RPM quantization step, and the 1-second gating time introduces latency "
        "during aggressive motor transients."
    )

# =============================================================================
# SLIDE 12: CONCLUSION AND FUTURE SCOPE
# =============================================================================
def build_slide_12():
    slide = create_base_slide()
    add_header(slide, "Conclusion & Future Roadmap", "Project Summary, Pre-Silicon Validation Outcomes, and Roadmap to Physical IoT Prototype")

    col_w = Inches(5.9)
    col_h = Inches(4.35)
    top_y = Inches(1.45)

    # Column 1: Conclusion Summary
    add_card(slide, Inches(0.6), top_y, col_w, col_h, "Project Synthesis & Conclusions", COLOR_NAVY_DARK)
    tb_c = slide.shapes.add_textbox(Inches(0.8), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True

    conclusions = [
        ("Successful Algorithmic Modeling", "Designed and validated a software simulation model of a digital tachometer with velocity estimation for ESD & IoT Experiment 11."),
        ("Decoupled Hardware Timer Concept", "Demonstrated that using STM32 Timer 2 for asynchronous pulse accumulation and Timer 3 for periodic 1.0s interrupt timing provides non-blocking, deterministic performance."),
        ("Accurate RPM & SI Velocity Derivation", "Successfully calculated rotational speed (RPM) and angular velocity (rad/s) across 0, 300, 600, and 900 RPM test regimes with zero mathematical error."),
        ("Value of Pre-Implementation Simulation", "Software simulation successfully proved firmware logic, division-by-zero protection, and edge counting conventions before physical PCB assembly.")
    ]
    for idx, (head, body) in enumerate(conclusions):
        p_h = tf_c.paragraphs[0] if idx == 0 else tf_c.add_paragraph()
        p_h.text = f"✔ {head}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9.5)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(5) if idx > 0 else Pt(0)

        p_b = tf_c.add_paragraph()
        p_b.text = f"  {body}"
        p_b.font.size = Pt(8.5)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(2)

    # Column 2: Future Scope
    add_card(slide, Inches(6.8), top_y, col_w, col_h, "Future Scope & IoT Prototype Roadmap", COLOR_NAVY_LIGHT)
    tb_f = slide.shapes.add_textbox(Inches(7.0), top_y + Inches(0.55), col_w - Inches(0.4), col_h - Inches(0.65))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True

    future_pts = [
        ("Phase 1: Physical Hardware Assembly", "Assemble physical testbed with STM32F401RE Nucleo board, slotted optical sensor with LM393 comparator, and 12V DC motor test rig."),
        ("Phase 2: Signal Conditioning & Digital Filtering", "Integrate hardware RC low-pass filtering / Schmitt triggers for contact debouncing, and implement moving average / IIR filtering for low speeds."),
        ("Phase 3: High-Resolution Quadrature Decoding", "Migrate to industrial 360/600 CPR optical encoder using STM32 hardware 4x encoder mode for bidirectional rotation sensing and sub-RPM precision."),
        ("Phase 4: IoT Cloud Telemetry & Remote Dashboard", "Interface ESP32 / Wi-Fi module over UART to stream real-time speed, duty cycle, and vibration data via MQTT to a ThingSpeak / AWS IoT cloud dashboard.")
    ]
    for idx, (head, body) in enumerate(future_pts):
        p_h = tf_f.paragraphs[0] if idx == 0 else tf_f.add_paragraph()
        p_h.text = f"➔ {head}:"
        p_h.font.bold = True
        p_h.font.size = Pt(9.5)
        p_h.font.color.rgb = COLOR_NAVY_DARK
        p_h.space_before = Pt(5) if idx > 0 else Pt(0)

        p_b = tf_f.add_paragraph()
        p_b.text = f"  {body}"
        p_b.font.size = Pt(8.5)
        p_b.font.color.rgb = COLOR_SLATE_DARK
        p_b.space_before = Pt(2)

    # Closing Banner with "Thank You" and "Questions?"
    close_banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(5.95), Inches(12.1), Inches(1.0))
    close_banner.fill.solid()
    close_banner.fill.fore_color.rgb = COLOR_NAVY_DARK
    close_banner.line.color.rgb = COLOR_CYAN_ACCENT
    close_banner.line.width = Pt(1.5)
    tf_cb = close_banner.text_frame
    p_cb1 = tf_cb.paragraphs[0]
    p_cb1.text = "THANK YOU!"
    p_cb1.font.name = FONT_HEADING
    p_cb1.font.bold = True
    p_cb1.font.size = Pt(15)
    p_cb1.font.color.rgb = COLOR_CYAN_BRIGHT
    p_cb1.alignment = PP_ALIGN.CENTER

    p_cb2 = tf_cb.add_paragraph()
    p_cb2.text = "Open for Questions, Technical Discussion, and Laboratory Evaluation"
    p_cb2.font.name = FONT_BODY
    p_cb2.font.size = Pt(10)
    p_cb2.font.color.rgb = COLOR_WHITE
    p_cb2.alignment = PP_ALIGN.CENTER
    p_cb2.space_before = Pt(2)

    slide.notes_slide.notes_text_frame.text = (
        "In conclusion, Experiment 11 successfully demonstrates a digital tachometer with velocity estimation "
        "using an STM32 microcontroller model in software simulation. The mathematical algorithms, timer interrupt workflows, "
        "and SI velocity conversions have been thoroughly validated across multiple speed regimes. As future scope, "
        "we plan to assemble the physical hardware using an STM32 Nucleo board, apply digital filtering for noisy environments, "
        "upgrade to bidirectional 4x quadrature decoding, and integrate an ESP32 co-processor for IoT cloud telemetry. "
        "Thank you for your attention, and I am now open to questions from the evaluation committee."
    )

# =============================================================================
# MAIN PRESENTATION BUILDER
# =============================================================================
def main():
    print("Building polished 12-slide engineering presentation...")
    build_slide_1()
    build_slide_2()
    build_slide_3()
    build_slide_4()
    build_slide_5()
    build_slide_6()
    build_slide_7()
    build_slide_8()
    build_slide_9()
    build_slide_10()
    build_slide_11()
    build_slide_12()

    output_path = "Digital_Tachometer_Velocity_Estimation_Exp11.pptx"
    prs.save(output_path)
    print(f"Presentation successfully created and saved to: {output_path}")

if __name__ == "__main__":
    main()
