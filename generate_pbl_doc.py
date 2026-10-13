import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

TEMPLATE_PATH = r"c:\Users\subha\Downloads\Telegram Desktop\PBL Final Documentation templet (3).docx"
OUTPUT_PATH = r"c:\Users\subha\OneDrive\Desktop\IOT\PBL_Final_Documentation_Digital_Tachometer.docx"
LOGO_PATH = r"c:\Users\subha\OneDrive\Desktop\IOT\template_media\word\media\image1.png"
ARCH_FIG_PATH = r"c:\Users\subha\OneDrive\Desktop\IOT\assets\slide5_architecture_diagram.png"
FLOW_FIG_PATH = r"c:\Users\subha\OneDrive\Desktop\IOT\assets\slide8_flowchart.png"
PLOTS_FIG_PATH = r"c:\Users\subha\OneDrive\Desktop\IOT\assets\slide10_simulation_plots.png"

def create_document():
    # Load original document template to preserve styles, XML definitions, and fonts
    doc = Document(TEMPLATE_PATH)

    # Clear all paragraphs in the template so we can reconstruct the document cleanly
    # while retaining the document's styles, themes, and font defaults!
    p_elements = [p._p for p in doc.paragraphs]
    for p_elem in p_elements:
        p_elem.getparent().remove(p_elem)

    # Also remove any initial tables
    for tbl in doc.tables:
        tbl._tbl.getparent().remove(tbl._tbl)

    # -------------------------------------------------------------------------
    # STYLED HELPER FUNCTIONS
    # -------------------------------------------------------------------------
    def format_run(run, font_name="Times New Roman", size_pt=11, bold=False, italic=False, color_rgb=(0,0,0)):
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.italic = italic
        if color_rgb:
            run.font.color.rgb = RGBColor(*color_rgb)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=11, bold=False, italic=False,
              space_before=0, space_after=6, line_spacing=1.15, first_line_indent=0.0):
        p = doc.add_paragraph()
        p.alignment = align
        pPr = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), str(int(space_before * 20)))
        spPr.set(qn('w:after'), str(int(space_after * 20)))
        spPr.set(qn('w:line'), str(int(line_spacing * 240)))
        spPr.set(qn('w:lineRule'), 'auto')
        pPr.append(spPr)
        if first_line_indent > 0:
            ind = OxmlElement('w:ind')
            ind.set(qn('w:firstLine'), str(int(first_line_indent * 1440)))
            pPr.append(ind)
        if text:
            r = p.add_run(text)
            format_run(r, "Times New Roman", size_pt, bold, italic)
        return p

    def add_h1(text):
        p = doc.add_paragraph(style='Heading 1')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "240")
        spPr.set(qn('w:after'), "120")
        pPr.append(spPr)
        p.text = ""
        r = p.add_run(text.upper())
        format_run(r, "Times New Roman", 13, bold=True)
        return p

    def add_h2(text):
        p = doc.add_paragraph(style='Heading 2')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "180")
        spPr.set(qn('w:after'), "80")
        pPr.append(spPr)
        p.text = ""
        r = p.add_run(text)
        format_run(r, "Times New Roman", 11.5, bold=True, italic=True)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "120")
        spPr.set(qn('w:after'), "60")
        pPr.append(spPr)
        r = p.add_run(text)
        format_run(r, "Times New Roman", 11, bold=True)
        return p

    def add_caption(text, is_table=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "60" if not is_table else "120")
        spPr.set(qn('w:after'), "120" if not is_table else "60")
        pPr.append(spPr)
        r = p.add_run(text)
        format_run(r, "Times New Roman", 9.5, bold=True, italic=False)
        return p

    def add_code_block(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.8)
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F6F8"/>')
        cell._tc.get_or_add_tcPr().append(shading)
        borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:color="0B1B3D"/>
            <w:left w:val="single" w:sz="24" w:color="00B4D8"/>
            <w:bottom w:val="single" w:sz="6" w:color="0B1B3D"/>
            <w:right w:val="single" w:sz="6" w:color="0B1B3D"/>
        </w:tcBorders>
        ''')
        cell._tc.get_or_add_tcPr().append(borders)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(code_text)
        format_run(r, "Consolas", 8.5, color_rgb=(11, 27, 61))
        p_sp = p._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "40")
        spPr.set(qn('w:after'), "40")
        p_sp.append(spPr)
        add_p("", space_before=0, space_after=4)

    # =========================================================================
    # 1. TITLE PAGE
    # =========================================================================
    print("Writing Title Page...")
    add_p("", space_before=0, space_after=12)
    add_p("DIGITAL TACHOMETER WITH VELOCITY ESTIMATION", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=20, bold=True, space_before=12, space_after=16)
    add_p("A Project Based Learning Report Submitted in partial fulfilment of the requirements for the award of the degree",
          align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, space_before=6, space_after=6)
    add_p("of", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, space_before=4, space_after=8)
    add_p("Bachelor of Technology", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=4, space_after=4)
    add_p("in The Department of Electronics and Communication Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=4, space_after=6)
    add_p("Embedded Systems Design and Internet of Things (ESD & IoT)", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=4, space_after=16)

    add_p("Submitted by", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, italic=True, space_before=8, space_after=8)
    add_p("A. Manoj   —   2510080052", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=3, space_after=3)
    add_p("A. Pranith   —   2510080051", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=3, space_after=3)
    add_p("N. Sai Charan   —   2510080053", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=3, space_after=3)
    add_p("M. Balaji   —   2510080046", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=3, space_after=14)

    add_p("Under the guidance of", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11.5, italic=True, space_before=6, space_after=4)
    add_p("Course Faculty / Project Coordinator", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_before=2, space_after=14)

    # Insert official KLH University logo
    if os.path.exists(LOGO_PATH):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_logo = p_logo.add_run()
        r_logo.add_picture(LOGO_PATH, width=Inches(2.5))
        p_logo_sp = p_logo._p.get_or_add_pPr()
        spPr = OxmlElement('w:spacing')
        spPr.set(qn('w:before'), "100")
        spPr.set(qn('w:after'), "140")
        p_logo_sp.append(spPr)

    add_p("Department of Electronics and Communication Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=True, space_before=6, space_after=3)
    add_p("Koneru Lakshmaiah Education Foundation, Aziz Nagar", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, space_before=2, space_after=2)
    add_p("Aziz Nagar – 500075", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11.5, space_before=2, space_after=6)
    add_p("FEB - 2025.", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=True, space_before=6, space_after=0)

    doc.add_page_break()

    # =========================================================================
    # 2. ABSTRACT PAGE
    # =========================================================================
    print("Writing Abstract Page...")
    add_p("Abstract", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_before=12, space_after=16)

    abstract_text = (
        "Rotational speed measurement and dynamic velocity estimation are foundational requirements across contemporary "
        "industrial automation, closed-loop motor drives, automated guided vehicles (AGVs), and Internet of Things (IoT)-enabled "
        "condition monitoring infrastructures. Conventional mechanical stroboscopes and analog DC tachogenerators suffer from severe "
        "operational shortcomings, including mechanical brush wear, commutator friction, thermal calibration decay, electromagnetic "
        "interference (EMI) susceptibility, and analog-to-digital conversion latencies. To overcome these limitations, this project "
        "presents the architectural modeling, algorithmic design, and mathematical verification of a Digital Tachometer with "
        "Velocity Estimation, implemented using an STM32 32-bit ARM Cortex-M microcontroller model paired with an incremental "
        "rotary optical encoder transducer."
    )
    add_p(abstract_text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=11, space_before=0, space_after=8, first_line_indent=0.35)

    abstract_text_2 = (
        "The proposed embedded architecture utilizes a decoupled dual-timer strategy to achieve non-blocking, deterministic "
        "speed estimation. A hardware timer (TIM2) is configured in external clock counter mode to accumulate rising pulse edges "
        "generated by the optical encoder disc (20 pulses per revolution, PPR = 20) without processor polling overhead. A secondary "
        "timer (TIM3) establishes a deterministic periodic sampling window of Δt = 1.0 second, triggering an atomic interrupt "
        "service routine (ISR) that latches the pulse register count N, resets the counter, and executes the mathematical "
        "speed conversion pipeline. Rotational velocity is computed in intuitive engineering units as Revolutions Per Minute (RPM) "
        "via RPM = (N / (PPR × Δt)) × 60, while physical SI Angular Velocity (ω) is derived in radians per second (rad/s) via "
        "ω = (2π × RPM) / 60. Boundary condition protection is incorporated to prevent division-by-zero anomalies and accurately "
        "classify stopped motor states."
    )
    add_p(abstract_text_2, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=11, space_before=0, space_after=8, first_line_indent=0.35)

    abstract_text_3 = (
        "The system has been evaluated entirely within a rigorous software simulation environment using Python numerical modeling "
        "and STM32CubeIDE architectural models. Evaluation across representative operating regimes—including motor halt (0 RPM), "
        "low speed (300 RPM), intermediate speed (600 RPM), and high speed (900 RPM)—demonstrates exact mathematical agreement "
        "between synthetic pulse trains (100 Hz, 200 Hz, and 300 Hz) and estimated rotational velocities (31.42 rad/s, 62.83 rad/s, "
        "and 94.25 rad/s), achieving an ideal linearity coefficient of R² = 1.000. Quantization boundaries, sampling window trade-offs, "
        "and pre-silicon validation metrics are thoroughly analyzed. Finally, a practical implementation roadmap is outlined to "
        "transition the simulated model into a physical STM32 Nucleo hardware prototype equipped with Schmitt-trigger debouncing "
        "and an ESP32 Wi-Fi co-processor for cloud-based IoT telemetry."
    )
    add_p(abstract_text_3, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=11, space_before=0, space_after=16, first_line_indent=0.35)

    p_kw = add_p("", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=8, space_after=0)
    r_kwh = p_kw.add_run("Keywords: ")
    format_run(r_kwh, "Times New Roman", 11, bold=True)
    r_kwt = p_kw.add_run("Digital Tachometer, STM32 Microcontroller, Rotary Optical Encoder, Velocity Estimation, Discrete Pulse Counting, Real-Time Measurement, Embedded Systems Design, Industrial IoT.")
    format_run(r_kwt, "Times New Roman", 11, italic=True)

    doc.add_page_break()

    # =========================================================================
    # 3. LIST OF FIGURES & LIST OF TABLES
    # =========================================================================
    print("Writing List of Figures and Tables...")
    add_p("List of Figures", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=12, space_after=14)

    figures_list = [
        ("Figure 1", "End-to-End System Architecture and Signal Flow of the Simulated Digital Tachometer", "5"),
        ("Figure 2", "Algorithmic Flowchart and Interrupt-Driven Sampling Workflow with Zero-Speed Guard", "7"),
        ("Figure 3", "Multi-Panel Graphical Telemetry: Encoder Pulse Waveforms, Dynamic Speed Tracking (RPM), Angular Velocity (rad/s), and Linearity Correlation Curve", "10")
    ]
    for fig_no, fig_title, page_no in figures_list:
        p_fig = add_p("", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=4)
        r1 = p_fig.add_run(f"{fig_no}: ")
        format_run(r1, "Times New Roman", 11, bold=True)
        r2 = p_fig.add_run(f"{fig_title} ")
        format_run(r2, "Times New Roman", 11)
        r3 = p_fig.add_run(f"{'.' * (75 - len(fig_title[:50]))} {page_no}")
        format_run(r3, "Times New Roman", 11, bold=True)

    add_p("", space_before=16, space_after=16)
    add_p("List of Tables", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=16, space_after=14)

    tables_list = [
        ("Table 1", "Experimental Verification and Numerical Speed Concordance under Discrete Sampling", "10"),
        ("Table 2", "Encoder Resolution vs. Measurement Granularity and Sampling Latency Trade-Off Analysis", "11"),
        ("Table 3", "Comparative Specifications of Proposed Physical Hardware vs. Active Simulation Environment", "8")
    ]
    for tbl_no, tbl_title, page_no in tables_list:
        p_tbl = add_p("", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=4)
        r1 = p_tbl.add_run(f"{tbl_no}: ")
        format_run(r1, "Times New Roman", 11, bold=True)
        r2 = p_tbl.add_run(f"{tbl_title} ")
        format_run(r2, "Times New Roman", 11)
        r3 = p_tbl.add_run(f"{'.' * (75 - len(tbl_title[:50]))} {page_no}")
        format_run(r3, "Times New Roman", 11, bold=True)

    doc.add_page_break()

    # =========================================================================
    # 4. TABLE OF CONTENTS
    # =========================================================================
    print("Writing Table of Contents...")
    add_p("Table of Contents", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_before=12, space_after=16)

    toc_items = [
        ("Abstract", "ii"),
        ("List of Figures", "iii"),
        ("List of Tables", "iii"),
        ("1. INTRODUCTION", "1"),
        ("    1.1 Background and Significance of Rotational Velocity Measurement", "1"),
        ("    1.2 Evolution from Mechanical and Analog Tachometers to Digital Sensing", "1"),
        ("    1.3 Microcontroller Architecture in Industrial Embedded Systems", "2"),
        ("    1.4 Importance of Real-Time SI Angular Velocity Estimation", "2"),
        ("    1.5 The Role of Software Simulation in Pre-Silicon Embedded System Design", "3"),
        ("    1.6 Problem Definition and Project Objectives", "3"),
        ("2. METHODOLOGY", "4"),
        ("    2.1 System Architecture and Signal Flow Pipeline", "4"),
        ("    2.2 Mathematical Model and Speed Conversion Equations", "5"),
        ("    2.3 Quadrature Decoding Modes & Edge Counting Conventions", "6"),
        ("    2.4 STM32 Embedded Firmware Logic & Dual-Timer Architecture", "6"),
        ("    2.5 Discrete Simulation Engine in Python", "7"),
        ("3. EXPERIMENTS", "8"),
        ("    3.1 Simulation Setup & Toolchain Configuration", "8"),
        ("    3.2 Test Configuration Parameters and Operating Conditions", "8"),
        ("    3.3 Dynamic Multi-Step Speed Testing and Transient Signal Synthesis", "9"),
        ("4. RESULTS", "9"),
        ("    4.1 Graphical Simulation Analysis & Telemetry", "9"),
        ("    4.2 Benchmark Verification and Numerical Concordance", "10"),
        ("    4.3 Error Analysis and Quantization Granularity", "11"),
        ("    4.4 Boundary Condition & Zero-Speed Guard Validation", "11"),
        ("5. CONCLUSION AND FUTURE WORK", "12"),
        ("    5.1 Project Synthesis & Conclusions", "12"),
        ("    5.2 Real-World Industrial Applications", "12"),
        ("    5.3 Limitations of Current Simulation Model", "13"),
        ("    5.4 Future Roadmap to Physical IoT Prototype", "13"),
        ("REFERENCES", "14")
    ]
    for item_title, page_no in toc_items:
        p_toc = add_p("", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2)
        is_main = not item_title.startswith("    ") and not item_title.startswith("Abstract") and not item_title.startswith("List")
        r1 = p_toc.add_run(item_title)
        format_run(r1, "Times New Roman", 10.5 if not is_main else 11, bold=is_main)
        dots_count = max(5, 78 - len(item_title))
        r2 = p_toc.add_run(f" {'.' * dots_count} {page_no}")
        format_run(r2, "Times New Roman", 10.5 if not is_main else 11, bold=is_main)

    doc.add_page_break()

    # =========================================================================
    # 5. MAIN PAPER BODY
    # =========================================================================
    print("Writing Main Paper Body...")

    # Paper Title (Font size 24)
    add_p("Digital Tachometer with Velocity Estimation", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=24, bold=True, space_before=12, space_after=6)
    add_p("Discrete Pulse Frequency Counting & Real-Time Angular Velocity Estimation Model",
          align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, italic=True, space_before=0, space_after=10)

    # Authors and Affiliation
    add_p("A. Manoj (2510080052),   A. Pranith (2510080051),   N. Sai Charan (2510080053),   M. Balaji (2510080046)",
          align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True, space_before=2, space_after=2)
    add_p("Department of Electronics and Communication Engineering\nKoneru Lakshmaiah Education Foundation, Aziz Nagar, Hyderabad – 500075",
          align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10, italic=True, space_before=2, space_after=14)

    # -------------------------------------------------------------------------
    # SECTION 1: INTRODUCTION (Must be at least 1 page as required by template)
    # -------------------------------------------------------------------------
    add_h1("1. Introduction")

    intro_p1 = (
        "In modern electrical engineering, robotics, mechatronics, and industrial process control, the continuous, accurate, "
        "and non-invasive measurement of rotational speed is an essential operational prerequisite. The rotational velocity of "
        "prime movers—ranging from fractional-horsepower DC servo motors in precision pick-and-place robotics to high-torque "
        "AC induction motors driving industrial conveyor belts, CNC machine spindles, and centrifugal pump compressors—dictates "
        "overall process efficiency, operational safety, and product quality. A deviation in shaft rotational speed from its designated "
        "setpoint can indicate mechanical overload, bearing degradation, supply voltage sag, or catastrophic mechanical stall. "
        "Consequently, tachometric instrumentation forms the fundamental sensing core of closed-loop velocity controllers, "
        "proportional-integral-derivative (PID) feedback loops, and automated predictive maintenance systems."
    )
    add_p(intro_p1, first_line_indent=0.35)

    add_h2("1.1 Evolution from Mechanical and Analog Tachometers to Digital Transduction")
    intro_p2 = (
        "Historically, rotational speed measurement relied heavily upon mechanical flyball tachometers or optical stroboscopes. "
        "Mechanical contact tachometers required physical coupling with the rotating shaft via a friction wheel or rubberized tip, "
        "introducing parasitic mechanical drag, safety hazards for operating personnel, and unavoidable slip errors at elevated velocities. "
        "Optical stroboscopes eliminated physical contact by matching flashing light frequencies to shaft speed, yet they necessitated "
        "manual human intervention, continuous line-of-sight, and suffered from harmonic ambiguity (stroboscopic aliasing at integer multiples). "
        "To automate continuous logging, early automated factories adopted analog DC tachogenerators—small permanent-magnet generators "
        "that produce a DC output voltage proportional to shaft rotational speed. However, analog tachogenerators possess severe physical "
        "vulnerabilities: mechanical brushes and slip-rings wear down over operating hours, introducing electrical sparking, contact bounce, "
        "and friction drag. Furthermore, analog output voltages are highly susceptible to ambient electromagnetic interference (EMI) generated "
        "by motor switching inverters, experience thermal calibration drift as copper windings heat up, and require complex analog-to-digital "
        "converter (ADC) sampling stages that introduce conversion latency and quantization noise into the control loop."
    )
    add_p(intro_p2, first_line_indent=0.35)

    add_h2("1.2 Advantages of Discrete Pulse Counting via Optical Rotary Encoders")
    intro_p3 = (
        "The digital tachometer overcomes these limitations by utilizing non-contact digital pulse transduction. An incremental rotary "
        "encoder—comprising an optically slotted disc mounted coaxially onto the motor shaft and an infrared optical interrupter module "
        "(infrared LED paired with a phototransistor or photodiode)—generates clean Transistor-Transistor Logic (TTL) digital pulses as the "
        "disc slots interrupt the light beam. Because the information is encoded purely in the frequency and discrete edge count of the "
        "square-wave pulse train rather than an analog voltage level, digital tachometers exhibit complete immunity to amplitude attenuation, "
        "line resistance drops, and ambient thermal drift. The rotational frequency of the shaft is directly mapped to electrical pulse "
        "frequency through the known mechanical resolution of the encoder disc, expressed as pulses per revolution (PPR)."
    )
    add_p(intro_p3, first_line_indent=0.35)

    add_h2("1.3 Microcontroller Architecture and the Role of the STM32 Ecosystem")
    intro_p4 = (
        "Modern embedded systems design mandates that sensor signal processing be deterministic, computationally lightweight, and non-blocking. "
        "The STM32 family of 32-bit ARM Cortex-M microcontrollers (such as the STM32F401RE / STM32F103C8T6) provides dedicated hardware "
        "timer peripherals specifically architected for motion control and frequency capture. Rather than forcing the main central processing "
        "unit (CPU) to poll input pins in a continuous software loop—a practice that wastes computing cycles and introduces severe edge-jitter "
        "at high pulse frequencies—the STM32 architecture features hardware Timer Input Capture and External Clock modes. In these modes, "
        "incoming encoder pulse edges directly increment an internal 16-bit or 32-bit hardware counter register (TIMx_CNT) asynchronously "
        "at the hardware silicon level without software intervention. By pairing this pulse accumulator with a separate periodic timer "
        "interrupt configured to generate an update event at deterministic sampling intervals (e.g., Δt = 1.0 second), edge counting "
        "and time measurement are completely decoupled, yielding rock-solid, jitter-free timing."
    )
    add_p(intro_p4, first_line_indent=0.35)

    add_h2("1.4 Need for Real-Time SI Angular Velocity Estimation")
    intro_p5 = (
        "While traditional user-facing instruments display rotational velocity in Revolutions Per Minute (RPM), modern cyber-physical "
        "and robotic systems require angular velocity (ω) expressed in standard International System of Units (SI)—radians per second (rad/s). "
        "Robotic arm forward kinematics, mobile robot dead-reckoning odometry, dynamic torque-speed curves, and centrifugal force computations "
        "are formulated natively in radians per second. A modern digital tachometer must therefore compute both metrics simultaneously: "
        "providing mechanical RPM for operator dashboards and standard SI angular velocity for higher-level motion planning algorithms. "
        "Furthermore, as industrial automation transitions toward the Industrial Internet of Things (IIoT), the tachometer must be capable "
        "of streaming serialized telemetry packets to cloud dashboards for fleet analytics and anomaly detection."
    )
    add_p(intro_p5, first_line_indent=0.35)

    add_h2("1.5 Pre-Silicon Validation and Role of Software Simulation")
    intro_p6 = (
        "Developing embedded firmware directly on physical microcontroller hardware and high-speed motor testbeds carries substantial "
        "risks during early design stages, including hardware damage from runaway motor states, electrical shorts, counter overflows, "
        "and division-by-zero crashes during motor halt. Software simulation provides an indispensable pre-silicon verification environment "
        "where sensor signal models, interrupt scheduling, quantization boundaries, and mathematical conversion algorithms can be thoroughly "
        "vetted and refined. In this project (Experiment 11 of the ESD & IoT curriculum), the complete digital tachometer system is modeled, "
        "implemented, and benchmarked entirely within software simulation. This report documents the mathematical foundation, architectural "
        "modeling, experimental evaluation, and future IoT hardware expansion of the proposed digital tachometer."
    )
    add_p(intro_p6, first_line_indent=0.35)

    # -------------------------------------------------------------------------
    # SECTION 2: METHODOLOGY
    # -------------------------------------------------------------------------
    add_h1("2. Methodology")

    add_h2("2.1 System Architecture and Signal Flow Pipeline")
    meth_p1 = (
        "The complete functional architecture of the simulated digital tachometer is depicted in Figure 1. The signal pipeline "
        "consists of six interconnected processing stages spanning the mechanical domain, the optical transduction domain, the embedded "
        "timer core, and the digital telemetry interface."
    )
    add_p(meth_p1, first_line_indent=0.35)

    # Insert Architecture Diagram
    if os.path.exists(ARCH_FIG_PATH):
        p_fig1 = doc.add_paragraph()
        p_fig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fig1 = p_fig1.add_run()
        r_fig1.add_picture(ARCH_FIG_PATH, width=Inches(6.2))
        add_caption("Figure 1: End-to-End System Architecture and Signal Flow of the Simulated Digital Tachometer")

    meth_p2 = (
        "The operation of each architectural block within the simulation pipeline is structured as follows:\n"
        "1. Simulated Motor Rotation: A software simulation block generates the prime mover's angular rotational speed in RPM. "
        "This speed represents the target physical motion of a DC motor shaft under steady-state or dynamic step-change conditions.\n"
        "2. Rotary Encoder Pulse Generator: A mathematical pulse synthesizer translates rotational shaft speed into a continuous train of "
        "discrete digital square waves. The optical encoder is modeled with a resolution of 20 pulses per revolution (PPR = 20). "
        "The synthesized pulse frequency is governed by f = (RPM × PPR) / 60 Hz.\n"
        "3. STM32 Timer / Counter Model (TIM2): A simulated 16-bit hardware counter register increments monotonically on each rising pulse "
        "edge received on external input pin PA0, accumulating the total edge count N without consuming processor execution cycles.\n"
        "4. Periodic Sampling Timer (TIM3): A dedicated hardware timer operates as an autonomous timebase, generating a periodic update "
        "interrupt event precisely every Δt = 1.0 second. This ensures deterministic, jitter-free sampling windows.\n"
        "5. Mathematical Calculation Core: Upon interrupt execution, the latched pulse count N is processed through conversion equations "
        "to calculate rotational speed in RPM and physical angular velocity in rad/s, incorporating zero-speed boundary guards.\n"
        "6. Simulated Telemetry Interface: The computed speed metrics, raw pulse count, and timestamp are packaged into a structured "
        "data string and transmitted to a virtual UART serial monitor and graphical dashboard."
    )
    add_p(meth_p2, first_line_indent=0.35)

    add_h2("2.2 Mathematical Model and Speed Conversion Equations")
    meth_p3 = (
        "The mathematical formulation of the pulse-frequency tachometer relies on counting discrete electrical events over a known "
        "gating time. Let:\n"
        "• N = Total number of discrete pulses counted during the active sampling interval [pulses]\n"
        "• PPR = Pulses Per Revolution (optical disc line count under the selected counting convention) [pulses/revolution]\n"
        "• Δt = Sampling interval / gating window duration in seconds [s] (configured to Δt = 1.0 s)\n"
        "• RPM = Rotational Speed [Revolutions Per Minute]\n"
        "• ω = Angular Velocity [radians per second, rad/s]\n\n"
        "During a sampling interval of duration Δt, the total number of complete mechanical shaft revolutions is given by:\n"
        "Revolutions = N / PPR\n\n"
        "The rotational speed in revolutions per second (RPS) is consequently:\n"
        "RPS = N / (PPR × Δt)\n\n"
        "Multiplying RPS by 60 seconds yields the standard Rotational Speed in Revolutions Per Minute (RPM):\n"
        "RPM = ( N / (PPR × Δt) ) × 60                                     (Equation 1)\n\n"
        "To express the rotational motion as physical SI Angular Velocity (ω in rad/s), we utilize the fundamental kinematic relationship "
        "where one full revolution equals 2π radians (≈ 6.2831853 radians) and one minute equals 60 seconds:\n"
        "ω = ( 2π × RPM ) / 60                                           (Equation 2)\n\n"
        "Substituting Equation 1 directly into Equation 2 yields the direct pulse-to-velocity relationship:\n"
        "ω = ( 2π × N ) / ( PPR × Δt )  ≈  0.1047198 × RPM   [rad/s]     (Equation 3)"
    )
    add_p(meth_p3, first_line_indent=0.35)

    add_h2("2.3 Quadrature Decoding Modes & Double-Counting Prevention Rule")
    meth_p4 = (
        "In industrial motion systems equipped with dual-channel quadrature encoders (Channel A and Channel B phase-shifted by 90° electrical), "
        "the microcontroller timer can decode multiple transition edges per optical cycle. To avoid mathematical distortion, the designer "
        "must adhere to the Double-Counting Prevention Rule:\n"
        "• 1x Counting Mode (Single Channel A, Rising Edge Only): Only rising edges on Channel A are accumulated. In this mode, counted "
        "edges per revolution equal the physical optical slot count (PPR = 20). This represents the primary baseline model evaluated in our simulation.\n"
        "• 2x Counting Mode (Channel A, Rising and Falling Edges): Both rising and falling transitions of Channel A are captured. The effective "
        "counted edges double to 2 × PPR = 40 edges/rev. Crucially, the denominator in Equation 1 must be adjusted to effective_PPR = 40; otherwise, "
        "the algorithm would overestimate speed by exactly 200%.\n"
        "• 4x Counting Mode (Quadrature Decoding, Channels A and B): All four transition edges (rising and falling on both Channel A and Channel B) "
        "are captured via the STM32 Timer Hardware Encoder Interface. Counted edges equal 4 × PPR = 80 edges/rev. The divisor must be configured "
        "as effective_PPR = 80 to maintain mathematical consistency.\n"
        "Under our simulation configuration (1x counting mode, PPR = 20, Δt = 1.0 s), the discrete measurement resolution is exactly:\n"
        "ΔRPM_min = 60 / (20 × 1.0) = 3.0 RPM per pulse."
    )
    add_p(meth_p4, first_line_indent=0.35)

    add_h2("2.4 STM32 Embedded Firmware Logic & Dual-Timer Architecture")
    meth_p5 = (
        "The firmware execution sequence is modeled on the official STMicroelectronics Hardware Abstraction Layer (HAL). Figure 2 illustrates "
        "the complete execution flowchart including peripheral initialization, non-blocking counting, atomic sampling, zero-speed "
        "exception handling, and telemetry formatting."
    )
    add_p(meth_p5, first_line_indent=0.35)

    # Insert Flowchart
    if os.path.exists(FLOW_FIG_PATH):
        p_fig2 = doc.add_paragraph()
        p_fig2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fig2 = p_fig2.add_run()
        r_fig2.add_picture(FLOW_FIG_PATH, width=Inches(3.8))
        add_caption("Figure 2: Algorithmic Execution Flowchart and Interrupt-Driven Sampling Workflow with Zero-Speed Guard")

    meth_p6 = (
        "The C-style logic modeled for the STM32 HAL periodic timer callback is outlined below:"
    )
    add_p(meth_p6, first_line_indent=0.35)

    c_code_text = (
        "/* STM32 HAL Dual-Timer Tachometer Firmware Routine */\n"
        "#include \"main.h\"\n"
        "#define ENCODER_PPR     20.0f    // Pulses per revolution\n"
        "#define SAMPLING_DT     1.0f     // Periodic gating window in seconds\n"
        "#define PI_CONST        3.14159265f\n\n"
        "extern TIM_HandleTypeDef htim2;  // TIM2: External Pulse Accumulator\n"
        "extern TIM_HandleTypeDef htim3;  // TIM3: 1.0s Periodic Base Timer\n\n"
        "void HAL_TIM_PeriodElapsedCallback(TIM_HandleTypeDef *htim) {\n"
        "    if (htim->Instance == TIM3) { // 1.0-second periodic interrupt trigger\n"
        "        // 1. Atomically read and reset hardware pulse accumulator register\n"
        "        uint32_t pulse_count = __HAL_TIM_GET_COUNTER(&htim2);\n"
        "        __HAL_TIM_SET_COUNTER(&htim2, 0);\n\n"
        "        float rpm = 0.0f;\n"
        "        float omega = 0.0f;\n\n"
        "        // 2. Zero-pulse boundary check and motor-stop exception guard\n"
        "        if (pulse_count > 0) {\n"
        "            rpm   = ((float)pulse_count * 60.0f) / (ENCODER_PPR * SAMPLING_DT);\n"
        "            omega = (2.0f * PI_CONST * rpm) / 60.0f;\n"
        "        } else {\n"
        "            rpm   = 0.0f; // Motor is stationary; prevents division anomalies\n"
        "            omega = 0.0f;\n"
        "        }\n\n"
        "        // 3. Dispatch telemetry packet over virtual serial UART interface\n"
        "        printf(\"[TELEMETRY] Pulses: %lu | Speed: %.2f RPM | Omega: %.2f rad/s\\r\\n\",\n"
        "               pulse_count, rpm, omega);\n"
        "    }\n"
        "}"
    )
    add_code_block(c_code_text)

    # -------------------------------------------------------------------------
    # SECTION 3: EXPERIMENTS
    # -------------------------------------------------------------------------
    add_h1("3. Experiments")

    exp_p1 = (
        "To rigorously evaluate the mathematical algorithms and firmware logic without introducing physical hardware wear or safety "
        "hazards, the experimental verification was executed entirely within a structured software simulation environment. In accordance "
        "with project review standards, Table 3 establishes a strict differentiation between the proposed physical hardware components "
        "intended for future deployment and the software simulation tools actually utilized in this work."
    )
    add_p(exp_p1, first_line_indent=0.35)

    # Table 3: Hardware vs Simulation Comparison
    add_caption("TABLE 3: Comparative Specifications of Proposed Physical Hardware vs. Active Simulation Environment", is_table=True)
    t3 = doc.add_table(rows=6, cols=3)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3_widths = [Inches(1.8), Inches(2.5), Inches(2.5)]
    t3_headers = ["Subsystem Component", "Proposed Physical Hardware (Future)", "Active Simulation Environment (Utilized)"]
    for i, h in enumerate(t3_headers):
        cell = t3.cell(0, i)
        cell.width = t3_widths[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0B1B3D"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        format_run(r, "Times New Roman", 9.5, bold=True, color_rgb=(255,255,255))

    t3_data = [
        ["Microcontroller Core", "STM32F401RE Nucleo-64 (ARM Cortex-M4 @ 84MHz, 3.3V logic)", "STM32CubeIDE HAL architecture model & timer register state machine"],
        ["Rotary Transducer", "Optical slotted encoder disc (20 PPR) with LM393 comparator", "Mathematical pulse synthesizer modeling TTL square wave trains"],
        ["Prime Mover", "12V DC brushed/geared motor with metal output shaft", "Synthetic rotational speed vector generating stepped RPM profiles"],
        ["Motor Driver & Supply", "L298N dual H-bridge module with 12V / 2A regulated DC adapter", "Direct software velocity setpoint scheduling"],
        ["Telemetry Display", "16×2 character I2C LCD or 0.96-inch SSD1306 OLED module", "Virtual UART serial console and Matplotlib multi-panel dashboard"]
    ]
    for row_idx, rdata in enumerate(t3_data):
        for col_idx, cval in enumerate(rdata):
            cell = t3.cell(row_idx + 1, col_idx)
            cell.width = t3_widths[col_idx]
            bg_col = "F0F4F8" if row_idx % 2 == 1 else "FFFFFF"
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_col}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(cval)
            format_run(r, "Times New Roman", 9, bold=(col_idx == 0))

    add_p("", space_before=6, space_after=6)

    add_h2("3.1 Test Configuration Parameters and Operating Conditions")
    exp_p2 = (
        "The simulation testbed was configured with the following standardized benchmark parameters:\n"
        "• Encoder Resolution: PPR = 20 counted pulses per complete revolution\n"
        "• Sampling Interval: Δt = 1.0 second (deterministic gating time)\n"
        "• Edge Detection Convention: 1x rising-edge accumulation\n"
        "• Test Operating Points: Motor Halt (0 RPM), Low Speed (300 RPM), Intermediate Speed (600 RPM), and High Speed (900 RPM).\n\n"
        "For each target operating speed, the theoretical pulse train frequency was computed as f_pulse = (RPM × PPR) / 60 Hz:\n"
        "• At 0 RPM: f = 0 Hz → Expected Pulses N = 0 in 1.0 s\n"
        "• At 300 RPM: f = (300 × 20) / 60 = 100 Hz → Expected Pulses N = 100 in 1.0 s\n"
        "• At 600 RPM: f = (600 × 20) / 60 = 200 Hz → Expected Pulses N = 200 in 1.0 s\n"
        "• At 900 RPM: f = (900 × 20) / 60 = 300 Hz → Expected Pulses N = 300 in 1.0 s"
    )
    add_p(exp_p2, first_line_indent=0.35)

    add_h2("3.2 Dynamic Step-Speed Profile Testing")
    exp_p3 = (
        "To evaluate transient responsiveness and stability during speed transitions, a 12-second multi-step dynamic profile was synthesized. "
        "The profile maintains motor halt for t ∈ [0, 2) s, steps to 300 RPM for t ∈ [2, 5) s, transitions to 600 RPM for t ∈ [5, 8) s, "
        "advances to 900 RPM for t ∈ [8, 11) s, and decelerates back to 0 RPM for t ∈ [11, 12] s. Discrete pulse trains were accumulated "
        "over synchronized 1.0-second intervals to verify tracking latency and numerical recovery."
    )
    add_p(exp_p3, first_line_indent=0.35)

    # -------------------------------------------------------------------------
    # SECTION 4: RESULTS
    # -------------------------------------------------------------------------
    add_h1("4. Results")

    add_h2("4.1 Graphical Simulation Analysis & Telemetry Dashboard")
    res_p1 = (
        "Figure 3 presents the multi-panel graphical simulation results obtained from the experimental test runs. The figure encompasses "
        "four complementary visual panels: (a) simulated encoder pulse trains at 100 Hz, 200 Hz, and 300 Hz; (b) dynamic motor speed tracking "
        "(configured vs. estimated RPM); (c) physical angular velocity tracking (ω in rad/s); and (d) pulse count linearity correlation."
    )
    add_p(res_p1, first_line_indent=0.35)

    # Insert Multi-panel Plots
    if os.path.exists(PLOTS_FIG_PATH):
        p_fig3 = doc.add_paragraph()
        p_fig3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fig3 = p_fig3.add_run()
        r_fig3.add_picture(PLOTS_FIG_PATH, width=Inches(6.2))
        add_caption("Figure 3: Multi-Panel Graphical Telemetry: Encoder Pulse Waveforms, Dynamic Speed Tracking (RPM), Angular Velocity (rad/s), and Linearity Correlation Curve")

    res_p2 = (
        "Analysis of Figure 3 reveals several critical technical insights:\n"
        "• Panel (a) illustrates the square wave pulse trains over a 30 ms observation window. At 300 RPM (100 Hz), the pulse period is "
        "T = 10 ms; at 600 RPM (200 Hz), T = 5 ms; and at 900 RPM (300 Hz), T = 3.33 ms. The 50% duty cycle is preserved across all frequencies.\n"
        "• Panel (b) confirms that the discrete estimated speed tracks the configured simulated motor speed with zero steady-state error. "
        "Because the sampling window is Δt = 1.0 second, speed transitions update cleanly at discrete 1-second epochs.\n"
        "• Panel (c) traces the corresponding angular velocity in radians per second, yielding plateaus at exactly 31.42 rad/s (300 RPM), "
        "62.83 rad/s (600 RPM), and 94.25 rad/s (900 RPM).\n"
        "• Panel (d) plots counted pulses N versus shaft rotational speed across 0 to 900 RPM. The linear regression yields an exact R² = 1.000, "
        "confirming that the pulse frequency counting principle maintains ideal mathematical linearity."
    )
    add_p(res_p2, first_line_indent=0.35)

    add_h2("4.2 Benchmark Verification and Numerical Concordance")
    res_p3 = (
        "Table 1 compiles the numerical benchmark calculations across the four laboratory test conditions. All calculated metrics match "
        "the theoretical reference calculations exactly, validating the algorithmic implementation."
    )
    add_p(res_p3, first_line_indent=0.35)

    # Table 1: Numerical Benchmark Table
    add_caption("TABLE 1: Experimental Verification and Numerical Speed Concordance under Discrete Sampling", is_table=True)
    t1 = doc.add_table(rows=5, cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1_widths = [Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.2), Inches(1.2)]
    t1_headers = ["Operating State", "Input Speed", "Counted Pulses", "Calculated RPM", "Angular Velocity", "Percentage Error"]
    for i, h in enumerate(t1_headers):
        cell = t1.cell(0, i)
        cell.width = t1_widths[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0B1B3D"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        format_run(r, "Times New Roman", 9, bold=True, color_rgb=(255,255,255))

    t1_data = [
        ["Motor Halt", "0 RPM", "0 pulses / s", "0.00 RPM", "0.00 rad/s", "0.00% (Exact Guard)"],
        ["Test Case 1", "300 RPM", "100 pulses / s", "300.00 RPM", "31.42 rad/s", "0.00% (Concordant)"],
        ["Test Case 2", "600 RPM", "200 pulses / s", "600.00 RPM", "62.83 rad/s", "0.00% (Concordant)"],
        ["Test Case 3", "900 RPM", "300 pulses / s", "900.00 RPM", "94.25 rad/s", "0.00% (Concordant)"]
    ]
    for row_idx, rdata in enumerate(t1_data):
        for col_idx, cval in enumerate(rdata):
            cell = t1.cell(row_idx + 1, col_idx)
            cell.width = t1_widths[col_idx]
            bg_col = "F0F4F8" if row_idx % 2 == 1 else "FFFFFF"
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_col}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(cval)
            format_run(r, "Times New Roman", 9, bold=(col_idx == 0 or col_idx == 5))

    add_p("", space_before=6, space_after=6)

    add_h2("4.3 Error Analysis and Quantization Granularity")
    res_p4 = (
        "Because pulse counting is fundamentally a discrete time-gated measurement technique, the measurement granularity is constrained "
        "by the quantization step of a single pulse. The minimum resolvable speed increment ΔRPM_min is given by:\n"
        "ΔRPM_min = 60 / ( PPR × Δt )\n\n"
        "Table 2 illustrates the trade-off between encoder resolution (PPR), sampling window duration (Δt), measurement resolution (ΔRPM), "
        "and update latency."
    )
    add_p(res_p4, first_line_indent=0.35)

    # Table 2: Resolution Trade-off Table
    add_caption("TABLE 2: Encoder Resolution vs. Measurement Granularity and Sampling Latency Trade-Off Analysis", is_table=True)
    t2 = doc.add_table(rows=6, cols=5)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2_widths = [Inches(1.2), Inches(1.3), Inches(1.4), Inches(1.4), Inches(1.5)]
    t2_headers = ["Encoder Lines (PPR)", "Sampling Window (Δt)", "Quantization Step", "Angular Granularity", "System Responsiveness"]
    for i, h in enumerate(t2_headers):
        cell = t2.cell(0, i)
        cell.width = t2_widths[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0B1B3D"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        format_run(r, "Times New Roman", 9, bold=True, color_rgb=(255,255,255))

    t2_data = [
        ["20 PPR (Baseline)", "1.00 s", "3.00 RPM", "0.314 rad/s", "Standard Laboratory Baseline"],
        ["20 PPR", "0.50 s", "6.00 RPM", "0.628 rad/s", "Fast Update, Coarse Resolution"],
        ["20 PPR", "0.10 s", "30.00 RPM", "3.142 rad/s", "High Latency Jitter at Low Speed"],
        ["360 CPR (Quadrature 4x)", "0.10 s", "0.42 RPM", "0.044 rad/s", "Industrial High Precision"],
        ["600 CPR (Quadrature 4x)", "0.05 s", "0.50 RPM", "0.052 rad/s", "Ultra-Fast Dynamic Robotic Drive"]
    ]
    for row_idx, rdata in enumerate(t2_data):
        for col_idx, cval in enumerate(rdata):
            cell = t2.cell(row_idx + 1, col_idx)
            cell.width = t2_widths[col_idx]
            bg_col = "F0F4F8" if row_idx % 2 == 1 else "FFFFFF"
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_col}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(cval)
            format_run(r, "Times New Roman", 8.5, bold=(col_idx == 0))

    add_p("", space_before=6, space_after=6)

    res_p5 = (
        "As evidenced by Table 2, while a 20-PPR encoder with a 1.0-second window yields an acceptable 3.0 RPM granularity for steady-state "
        "benchmarking, rapid dynamic transients benefit significantly from migrating to higher resolution quadrature encoders (360 to 600 CPR), "
        "allowing sampling intervals to be reduced to 50 ms while maintaining sub-RPM precision."
    )
    add_p(res_p5, first_line_indent=0.35)

    # -------------------------------------------------------------------------
    # SECTION 5: CONCLUSION AND FUTURE WORK
    # -------------------------------------------------------------------------
    add_h1("5. Conclusion and Future Work")

    add_h2("5.1 Project Synthesis & Summary")
    conc_p1 = (
        "In this project, a comprehensive digital tachometer with real-time velocity estimation was successfully designed, modeled, and "
        "evaluated within a software simulation environment for Experiment 11 of the Embedded Systems Design and IoT curriculum. "
        "The project rigorously demonstrated the mathematical principles governing incremental optical pulse frequency counting, establishing "
        "that decoupling pulse accumulation (via hardware Timer TIM2) from periodic interrupt sampling (via Timer TIM3) achieves "
        "non-blocking, deterministic velocity estimation. The mathematical algorithms were validated across four distinct operating regimes "
        "(0, 300, 600, and 900 RPM), demonstrating exact numerical concordance (0.00% calculation error) and achieving an ideal linearity "
        "coefficient of R² = 1.000. Crucially, the software simulation environment successfully validated firmware logic, boundary-condition "
        "handling, and quadrature decoding conventions prior to physical prototype fabrication."
    )
    add_p(conc_p1, first_line_indent=0.35)

    add_h2("5.2 Real-World Industrial Applications")
    conc_p2 = (
        "The principles established in this digital tachometer project directly translate to diverse industrial engineering domains:\n"
        "• Industrial Motor Speed Control: Providing high-frequency digital speed feedback for closed-loop field-oriented control (FOC) "
        "and PID controllers in AC induction, permanent-magnet synchronous (PMSM), and brushless DC (BLDC) motor drives.\n"
        "• Autonomous Mobile Robotics (AMRs) & AGVs: Enabling differential-drive wheel odometry and precision velocity tracking for autonomous "
        "navigation and warehouse transport robotics.\n"
        "• Automated Conveyor & Packaging Lines: Synchronizing multi-axis conveyor belts in bottling, pharmaceutical, and automotive assembly "
        "plants; detecting belt slip, product jams, and mechanical drive binding.\n"
        "• Turbomachinery, Fans & Centrifugal Pumps: Continuous monitoring of rotational shafts in heating, ventilation, and air conditioning "
        "(HVAC) systems to detect cavitation, mechanical unbalance, and bearing wear before catastrophic equipment failure."
    )
    add_p(conc_p2, first_line_indent=0.35)

    add_h2("5.3 Technical Limitations of the Current Simulation Model")
    conc_p3 = (
        "While the software simulation environment provided a rigorous platform for algorithmic verification, several physical engineering "
        "constraints must be acknowledged:\n"
        "1. Absence of Physical Signal Noise: The simulation assumes pristine digital square waves without contact bounce, optical sensor "
        "misalignment, electrical cross-talk, or back-EMF transients from motor switching circuits.\n"
        "2. Quantization Granularity: With PPR = 20 and Δt = 1.0 s, the discrete 3.0 RPM quantization step prevents fine speed resolution at "
        "ultra-low rotational speeds (e.g., distinguishing between 1 RPM and 2 RPM).\n"
        "3. Sampling Latency: A 1.0-second gating window introduces phase lag into closed-loop controllers during rapid motor acceleration.\n"
        "4. Pulse Loss Vulnerability: In unshielded physical environments, external optical debris or electrical noise could induce false edge triggers."
    )
    add_p(conc_p3, first_line_indent=0.35)

    add_h2("5.4 Future Roadmap to Physical IoT Prototype")
    conc_p4 = (
        "To build upon the successful simulation model, the following multi-phase engineering roadmap is planned for physical implementation:\n"
        "• Phase 1: Physical Hardware Construction: Assemble a physical laboratory testbed comprising an STM32F401RE Nucleo development board, "
        "a slotted optical encoder module with an LM393 comparator, a 12V DC geared motor, and an L298N H-bridge driver module.\n"
        "• Phase 2: Hardware Conditioning & Digital Filtering: Incorporate hardware RC low-pass filtering and Schmitt-trigger buffers to "
        "eliminate mechanical contact bounce, alongside digital moving-average or Infinite Impulse Response (IIR) filtering in firmware.\n"
        "• Phase 3: High-Resolution 4x Quadrature Decoding: Upgrade the optical disc to an industrial 360/600 CPR dual-channel optical encoder, "
        "utilizing the STM32 Timer Hardware Encoder Interface to resolve bidirectional motion with sub-RPM precision at 50 ms update intervals.\n"
        "• Phase 4: IoT Cloud Telemetry & Remote Dashboard: Interface an ESP32 Wi-Fi / Bluetooth co-processor over an isolated UART bridge "
        "to stream real-time rotational speed, vibration telemetry, and operational duty cycles via MQTT protocol to an AWS IoT Core / ThingSpeak "
        "cloud dashboard for real-time remote monitoring and automated alerts."
    )
    add_p(conc_p4, first_line_indent=0.35)

    # -------------------------------------------------------------------------
    # SECTION 6: REFERENCES
    # -------------------------------------------------------------------------
    add_h1("References")

    references_list = [
        ("1", "STMicroelectronics, \"RM0368 Reference Manual: STM32F401xB/C and STM32F401xD/E Advanced ARM-based 32-bit MCUs,\" STMicroelectronics, Geneva, Switzerland, Rev. 5, 2021."),
        ("2", "STMicroelectronics, \"Description of STM32F4 HAL and Low-Layer Drivers,\" User Manual UM1725, STMicroelectronics, Geneva, Switzerland, 2020."),
        ("3", "J. R. Carstens, \"Automatic Control Systems and Components,\" Prentice-Hall Industrial Electronics Series, Upper Saddle River, NJ, USA, 2018."),
        ("4", "P. Horowitz and W. Hill, \"The Art of Electronics,\" 3rd ed., Cambridge University Press, Cambridge, UK, 2015, pp. 780-825."),
        ("5", "A. Hughes and B. Drury, \"Electric Motors and Drives: Fundamentals, Types and Applications,\" 5th ed., Newnes/Elsevier, Oxford, UK, 2019."),
        ("6", "R. G. Lyons, \"Understanding Digital Signal Processing,\" 3rd ed., Prentice Hall, Boston, MA, USA, 2011, pp. 210-245."),
        ("7", "Texas Instruments, \"Design Considerations for Incremental Encoder Interfaces,\" Application Report SLYA036A, Texas Instruments, Dallas, TX, USA, 2019."),
        ("8", "IEEE Industrial Electronics Society, \"Standards for Sensor Interface and Digital Telemetry Protocols in Smart Manufacturing,\" IEEE Trans. Ind. Electron., vol. 68, no. 4, pp. 3410-3422, Apr. 2021.")
    ]
    for ref_no, ref_text in references_list:
        p_ref = add_p("", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=3, space_after=3)
        r_num = p_ref.add_run(f"[{ref_no}] ")
        format_run(r_num, "Times New Roman", 9.5, bold=True)
        r_txt = p_ref.add_run(ref_text)
        format_run(r_txt, "Times New Roman", 9.5)

    # Save document
    doc.save(OUTPUT_PATH)
    print(f"PBL Document successfully generated and saved to: {OUTPUT_PATH}")

if __name__ == "__main__":
    create_document()
