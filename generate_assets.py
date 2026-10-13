import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from PIL import Image

os.makedirs("assets", exist_ok=True)

# Set consistent matplotlib styling
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#627D98'
plt.rcParams['axes.linewidth'] = 1.2

# -----------------------------------------------------------------------------
# 1. SLIDE 1 CONCEPT ART: Motor + Encoder + Pulses + STM32 Block + Dashboard
# -----------------------------------------------------------------------------
def create_slide1_concept():
    fig, ax = plt.subplots(figsize=(11, 4.2), dpi=300)
    fig.patch.set_facecolor('#0B1B3D')
    ax.set_facecolor('#0B1B3D')
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 4.2)
    ax.axis('off')

    # Card 1: Motor & Rotating Shaft
    rect_motor = patches.FancyBboxPatch((0.4, 0.9), 2.2, 2.4, boxstyle="round,pad=0.1",
                                        facecolor='#102A43', edgecolor='#00B4D8', linewidth=2)
    ax.add_patch(rect_motor)
    ax.text(1.5, 3.0, "MOTOR & SHAFT", color='#00E5FF', fontsize=11, fontweight='bold', ha='center')
    # Draw motor body icon
    motor_body = patches.Rectangle((0.8, 1.4), 1.0, 1.2, facecolor='#243B53', edgecolor='#9FB3C8', linewidth=1.5)
    ax.add_patch(motor_body)
    shaft = patches.Rectangle((1.8, 1.85), 0.6, 0.3, facecolor='#D9E2EC', edgecolor='#627D98', linewidth=1)
    ax.add_patch(shaft)
    ax.annotate("", xy=(1.5, 2.3), xytext=(1.1, 2.3),
                arrowprops=dict(arrowstyle="->", color="#00E5FF", lw=2, connectionstyle="arc3,rad=-0.5"))
    ax.text(1.5, 1.1, "Simulated Rotational\nMotion (ω, RPM)", color='#BCCCDC', fontsize=9, ha='center')

    # Arrow 1
    ax.annotate("", xy=(3.0, 2.0), xytext=(2.6, 2.0),
                arrowprops=dict(arrowstyle="->", color="#00E5FF", lw=2.5))

    # Card 2: Rotary Optical Encoder
    rect_enc = patches.FancyBboxPatch((3.0, 0.9), 2.3, 2.4, boxstyle="round,pad=0.1",
                                      facecolor='#102A43', edgecolor='#00B4D8', linewidth=2)
    ax.add_patch(rect_enc)
    ax.text(4.15, 3.0, "ROTARY ENCODER", color='#00E5FF', fontsize=11, fontweight='bold', ha='center')
    # Encoder disc
    disc = patches.Circle((4.15, 2.0), 0.65, facecolor='#1F3A5A', edgecolor='#00E5FF', linewidth=1.5)
    ax.add_patch(disc)
    for angle in np.linspace(0, 2*np.pi, 8, endpoint=False):
        x_s = 4.15 + 0.45 * np.cos(angle)
        y_s = 2.0 + 0.45 * np.sin(angle)
        ax.plot([4.15, x_s], [2.0, y_s], color='#627D98', lw=1.5)
    ax.text(4.15, 1.1, "Incremental Pulse Disc\n(PPR = 20 pulses/rev)", color='#BCCCDC', fontsize=9, ha='center')

    # Pulse train arrow
    ax.annotate("", xy=(5.7, 2.0), xytext=(5.3, 2.0),
                arrowprops=dict(arrowstyle="->", color="#00E5FF", lw=2.5))
    # Pulse graphic above arrow
    t_pulse = [5.32, 5.38, 5.38, 5.46, 5.46, 5.54, 5.54, 5.62, 5.62, 5.68]
    v_pulse = [2.4, 2.4, 2.7, 2.7, 2.4, 2.4, 2.7, 2.7, 2.4, 2.4]
    ax.plot(t_pulse, v_pulse, color='#00E5FF', lw=2)
    ax.text(5.5, 2.8, "Pulse Train", color='#00E5FF', fontsize=8, ha='center')

    # Card 3: STM32 Microcontroller Model
    rect_stm = patches.FancyBboxPatch((5.7, 0.9), 2.5, 2.4, boxstyle="round,pad=0.1",
                                      facecolor='#102A43', edgecolor='#00E5FF', linewidth=2.5)
    ax.add_patch(rect_stm)
    ax.text(6.95, 3.0, "STM32 MCU MODEL", color='#00E5FF', fontsize=11, fontweight='bold', ha='center')
    ax.text(6.95, 2.5, "• Timer Input Capture (TIMx)\n• 1.0s Periodic Update ISR\n• RPM & ω Calculation Engine",
            color='#E0FCFF', fontsize=8.5, ha='center')
    # IC chip drawing
    mcu_chip = patches.Rectangle((6.2, 1.3), 1.5, 0.9, facecolor='#0B1B3D', edgecolor='#00B4D8', linewidth=1.5)
    ax.add_patch(mcu_chip)
    ax.text(6.95, 1.75, "STM32 Core", color='#00E5FF', fontsize=9, fontweight='bold', ha='center')
    ax.text(6.95, 1.45, "Cortex-M @ 84MHz", color='#829AB1', fontsize=7.5, ha='center')

    # Arrow 3
    ax.annotate("", xy=(8.6, 2.0), xytext=(8.2, 2.0),
                arrowprops=dict(arrowstyle="->", color="#00E5FF", lw=2.5))

    # Card 4: Simulated Telemetry Dashboard
    rect_dash = patches.FancyBboxPatch((8.6, 0.9), 2.0, 2.4, boxstyle="round,pad=0.1",
                                       facecolor='#102A43', edgecolor='#00B4D8', linewidth=2)
    ax.add_patch(rect_dash)
    ax.text(9.6, 3.0, "VIRTUAL TELEMETRY", color='#00E5FF', fontsize=10.5, fontweight='bold', ha='center')
    
    screen = patches.Rectangle((8.8, 1.45), 1.6, 1.25, facecolor='#0B1B3D', edgecolor='#334E68', linewidth=1)
    ax.add_patch(screen)
    ax.text(9.6, 2.45, "STATUS: SIMULATED", color='#38BEC9', fontsize=7.5, fontweight='bold', ha='center')
    ax.text(9.6, 2.15, "Speed: 600 RPM", color='#F0F4F8', fontsize=8.5, fontweight='bold', ha='center')
    ax.text(9.6, 1.85, "ω: 62.83 rad/s", color='#00E5FF', fontsize=8.5, fontweight='bold', ha='center')
    ax.text(9.6, 1.58, "Pulses (N): 200 / s", color='#9FB3C8', fontsize=7.5, ha='center')
    ax.text(9.6, 1.1, "Simulated Virtual Output", color='#BCCCDC', fontsize=8.5, ha='center')

    # Bottom badge
    badge = patches.FancyBboxPatch((2.8, 0.15), 5.4, 0.45, boxstyle="round,pad=0.08",
                                   facecolor='#102A43', edgecolor='#486581', linewidth=1)
    ax.add_patch(badge)
    ax.text(5.5, 0.35, "EXPERIMENT 11: SYSTEM CONCEPTUAL MODEL (SOFTWARE SIMULATION)",
            color='#9FB3C8', fontsize=8.5, fontweight='bold', ha='center', va='center')

    plt.tight_layout()
    plt.savefig("assets/slide1_concept_art.png", dpi=300, facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()
    print("Generated assets/slide1_concept_art.png")

# -----------------------------------------------------------------------------
# 2. SLIDE 5: SYSTEM ARCHITECTURE BLOCK DIAGRAM
# -----------------------------------------------------------------------------
def create_slide5_architecture():
    fig, ax = plt.subplots(figsize=(12, 4.5), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 4.5)
    ax.axis('off')

    # Box styling helper
    def draw_box(x, y, w, h, title, subtext, header_bg='#0B1B3D', body_bg='#F0F4F8', border='#102A43'):
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08",
                                     facecolor=body_bg, edgecolor=border, linewidth=1.8)
        ax.add_patch(box)
        # Header banner inside
        header_h = 0.65
        hdr = patches.FancyBboxPatch((x, y + h - header_h), w, header_h, boxstyle="round,pad=0.05",
                                     facecolor=header_bg, edgecolor=header_bg, linewidth=1)
        ax.add_patch(hdr)
        ax.text(x + w/2, y + h - 0.33, title, color='#FFFFFF', fontsize=9.5, fontweight='bold', ha='center', va='center')
        ax.text(x + w/2, y + (h - header_h)/2, subtext, color='#102A43', fontsize=8, ha='center', va='center')

    # 1. Simulated Motor
    draw_box(0.3, 1.8, 1.6, 2.0, "Simulated Motor", "Rotational Source\nVelocity input: ω\nTarget Speed: N RPM\n(Software Model)",
             header_bg='#0B1B3D', body_bg='#F0F4F8', border='#0B1B3D')

    # Arrow 1
    ax.annotate("", xy=(2.2, 2.8), xytext=(1.9, 2.8),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2.5))
    ax.text(2.05, 3.0, "Shaft RPM", color='#0B1B3D', fontsize=7.5, ha='center')

    # 2. Rotary Encoder Pulse Generator
    draw_box(2.2, 1.8, 1.8, 2.0, "Encoder Generator", "Optical/Magnetic Disc\nResolution: PPR = 20\nGenerates Square Waves\nFreq: f = (RPM×PPR)/60",
             header_bg='#0F3E6D', body_bg='#EBF8FF', border='#0F3E6D')

    # Arrow 2
    ax.annotate("", xy=(4.3, 2.8), xytext=(4.0, 2.8),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2.5))
    ax.text(4.15, 3.0, "Pulses", color='#0B1B3D', fontsize=7.5, ha='center')

    # 3. STM32 Timer / Counter Model
    draw_box(4.3, 1.8, 2.0, 2.0, "STM32 Timer Model", "Timer Input Capture\nTIMx External Clock\nRising-edge Counting\nPulse Accumulator N",
             header_bg='#005B96', body_bg='#E6FFFA', border='#005B96')

    # Real-Time Clock / Periodic Interrupt Block (below)
    draw_box(4.3, 0.2, 2.0, 1.2, "Periodic Timer ISR", "TIMx / SysTick Base\nSampling Window: Δt = 1.0s\nGenerates Update Event",
             header_bg='#2D3748', body_bg='#EDF2F7', border='#4A5568')

    # Arrow from ISR to STM32 Timer / Counter Model
    ax.annotate("", xy=(5.3, 1.8), xytext=(5.3, 1.4),
                arrowprops=dict(arrowstyle="->", color="#E53E3E", lw=2.2, ls="--"))
    ax.text(5.95, 1.55, "1.0s Interrupt\nTrigger", color='#C53030', fontsize=7.5, ha='left')

    # Arrow 3
    ax.annotate("", xy=(6.6, 2.8), xytext=(6.3, 2.8),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2.5))
    ax.text(6.45, 3.0, "Pulse N, Δt", color='#0B1B3D', fontsize=7.5, ha='center')

    # 4. RPM Calculation
    draw_box(6.6, 1.8, 1.7, 2.0, "RPM Calculation", "Mathematical Core:\nRPM = N/(PPR×Δt)×60\nHandles Zero Count\nBounds Checking",
             header_bg='#0B1B3D', body_bg='#F0F4F8', border='#0B1B3D')

    # Arrow 4
    ax.annotate("", xy=(8.6, 2.8), xytext=(8.3, 2.8),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2.5))
    ax.text(8.45, 3.0, "RPM Data", color='#0B1B3D', fontsize=7.5, ha='center')

    # 5. Angular Velocity Estimation
    draw_box(8.6, 1.8, 1.6, 2.0, "Velocity Estimator", "Angular Speed:\nω = 2π × RPM / 60\nUnit: radians / second\nPrecision: 2 Decimals",
             header_bg='#0F3E6D', body_bg='#EBF8FF', border='#0F3E6D')

    # Arrow 5
    ax.annotate("", xy=(10.5, 2.8), xytext=(10.2, 2.8),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2.5))
    ax.text(10.35, 3.0, "RPM + ω", color='#0B1B3D', fontsize=7.5, ha='center')

    # 6. Simulated Display / Dashboard
    draw_box(10.5, 1.8, 1.3, 2.0, "Simulated UI", "Virtual Terminal\nSerial Monitor\nMatplotlib Plot\nData Logging",
             header_bg='#008080', body_bg='#E6FFFA', border='#008080')

    # Overview banner at the top
    top_band = patches.FancyBboxPatch((0.3, 4.05), 11.5, 0.35, boxstyle="round,pad=0.03",
                                      facecolor='#0B1B3D', edgecolor='#00B4D8', linewidth=1)
    ax.add_patch(top_band)
    ax.text(6.0, 4.22, "END-TO-END SIGNAL FLOW & SYSTEM ARCHITECTURE (SIMULATION PIPELINE)",
            color='#00E5FF', fontsize=8.5, fontweight='bold', ha='center', va='center')

    plt.tight_layout()
    plt.savefig("assets/slide5_architecture_diagram.png", dpi=300, facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()
    print("Generated assets/slide5_architecture_diagram.png")

# -----------------------------------------------------------------------------
# 3. SLIDE 8: ALGORITHM FLOWCHART
# -----------------------------------------------------------------------------
def create_slide8_flowchart():
    fig, ax = plt.subplots(figsize=(6.2, 7.5), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    ax.set_xlim(0, 6.2)
    ax.set_ylim(0, 8.5)
    ax.axis('off')

    def draw_oval(x, y, text, w=2.4, h=0.6, bg='#0B1B3D', text_col='#FFFFFF'):
        box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.25",
                                     facecolor=bg, edgecolor='#00B4D8', linewidth=2)
        ax.add_patch(box)
        ax.text(x, y, text, color=text_col, fontsize=9, fontweight='bold', ha='center', va='center')

    def draw_rect(x, y, text, sub="", w=3.4, h=0.75, bg='#F0F4F8', border='#0B1B3D'):
        box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.08",
                                     facecolor=bg, edgecolor=border, linewidth=1.5)
        ax.add_patch(box)
        if sub:
            ax.text(x, y + 0.14, text, color='#0B1B3D', fontsize=8.5, fontweight='bold', ha='center', va='center')
            ax.text(x, y - 0.16, sub, color='#486581', fontsize=7.5, ha='center', va='center')
        else:
            ax.text(x, y, text, color='#0B1B3D', fontsize=8.5, fontweight='bold', ha='center', va='center')

    def draw_diamond(x, y, text, w=2.6, h=0.9, bg='#EBF8FF', border='#00B4D8'):
        pts = [[x, y + h/2], [x + w/2, y], [x, y - h/2], [x - w/2, y]]
        poly = patches.Polygon(pts, closed=True, facecolor=bg, edgecolor=border, linewidth=1.8)
        ax.add_patch(poly)
        ax.text(x, y, text, color='#0B1B3D', fontsize=8.2, fontweight='bold', ha='center', va='center')

    # Nodes
    draw_oval(3.1, 8.1, "START")
    
    ax.annotate("", xy=(3.1, 7.55), xytext=(3.1, 7.8),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))
    
    draw_rect(3.1, 7.15, "Initialize Variables & Peripherals", "PPR = 20, Δt = 1.0s, N = 0, TIMx Ready")

    ax.annotate("", xy=(3.1, 6.4), xytext=(3.1, 6.75),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))

    draw_rect(3.1, 6.0, "Simulate / Receive Encoder Pulses", "Interrupt counts rising edges during window Δt")

    ax.annotate("", xy=(3.1, 5.25), xytext=(3.1, 5.6),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))

    draw_diamond(3.1, 4.75, "Is Pulse Count\nN > 0?", w=2.4, h=0.9)

    # Branch NO (Motor Stopped)
    ax.annotate("", xy=(0.8, 4.75), xytext=(1.9, 4.75),
                arrowprops=dict(arrowstyle="->", color="#C53030", lw=1.8))
    ax.text(1.35, 4.95, "No (0 pulses)", color='#C53030', fontsize=7.5, fontweight='bold', ha='center')
    
    draw_rect(0.8, 3.8, "Motor Stopped", "RPM = 0\nω = 0.00 rad/s", w=1.4, h=0.75, bg='#FFF5F5', border='#E53E3E')

    # Branch YES
    ax.annotate("", xy=(3.1, 3.9), xytext=(3.1, 4.3),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))
    ax.text(3.3, 4.1, "Yes", color='#0B1B3D', fontsize=7.5, fontweight='bold')

    draw_rect(3.1, 3.5, "Compute RPM", "RPM = (N / (PPR × Δt)) × 60", w=3.0, h=0.65)

    ax.annotate("", xy=(3.1, 2.8), xytext=(3.1, 3.15),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))

    draw_rect(3.1, 2.45, "Compute Angular Velocity (ω)", "ω = 2π × RPM / 60  [rad/s]", w=3.0, h=0.65)

    # Route from Stopped to Output
    ax.annotate("", xy=(3.1, 1.7), xytext=(3.1, 2.1),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))
    ax.plot([0.8, 0.8, 2.5], [3.4, 1.45, 1.45], color='#C53030', lw=1.5, ls="--")
    ax.annotate("", xy=(2.6, 1.45), xytext=(2.3, 1.45),
                arrowprops=dict(arrowstyle="->", color="#C53030", lw=1.5))

    draw_rect(3.1, 1.35, "Update Simulated Display / Serial Log", "Print RPM, ω, Pulses (N), Timestamp", w=3.4, h=0.65, bg='#E6FFFA', border='#00A3C4')

    ax.annotate("", xy=(3.1, 0.7), xytext=(3.1, 1.0),
                arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.8))

    draw_rect(3.1, 0.45, "Reset Counter N = 0 & Arm Next Δt", "Wait for next periodic timer event", w=3.4, h=0.55)

    # Loop back arrow
    ax.plot([4.85, 5.7, 5.7, 4.85], [0.45, 0.45, 6.0, 6.0], color='#00B4D8', lw=2)
    ax.annotate("", xy=(4.82, 6.0), xytext=(5.1, 6.0),
                arrowprops=dict(arrowstyle="->", color="#00B4D8", lw=2))
    ax.text(5.95, 3.2, "Next Interval\n(Continuous Loop)", color='#00B4D8', fontsize=7.5, fontweight='bold', ha='center', rotation=270)

    plt.tight_layout()
    plt.savefig("assets/slide8_flowchart.png", dpi=300, facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()
    print("Generated assets/slide8_flowchart.png")

# -----------------------------------------------------------------------------
# 4. SLIDE 10: SIMULATION PLOTS (Waveform, RPM vs Time, Angular Velocity, Linearity)
# -----------------------------------------------------------------------------
def create_slide10_plots():
    fig, axs = plt.subplots(2, 2, figsize=(10.5, 6.2), dpi=300)
    fig.patch.set_facecolor('#FFFFFF')
    
    # Subplot 1: Pulse Waveforms Comparison
    ax1 = axs[0, 0]
    ax1.set_facecolor('#F8FAFC')
    t_ms = np.linspace(0, 30, 1500) # 30 ms window
    # 300 RPM -> f = 100 Hz -> T = 10 ms
    # 600 RPM -> f = 200 Hz -> T = 5 ms
    # 900 RPM -> f = 300 Hz -> T = 3.33 ms
    sq_300 = 0.5 * (1 + np.sign(np.sin(2 * np.pi * 100 * t_ms / 1000.0)))
    sq_600 = 0.5 * (1 + np.sign(np.sin(2 * np.pi * 200 * t_ms / 1000.0)))
    sq_900 = 0.5 * (1 + np.sign(np.sin(2 * np.pi * 300 * t_ms / 1000.0)))

    ax1.plot(t_ms, sq_300 + 4.8, color='#0B1B3D', lw=1.8, label='300 RPM (f=100 Hz, N=100/s)')
    ax1.plot(t_ms, sq_600 + 2.4, color='#008080', lw=1.8, label='600 RPM (f=200 Hz, N=200/s)')
    ax1.plot(t_ms, sq_900, color='#00B4D8', lw=1.8, label='900 RPM (f=300 Hz, N=300/s)')
    ax1.set_title("Simulated Encoder Pulse Trains (PPR=20)", fontsize=10, fontweight='bold', color='#0B1B3D')
    ax1.set_xlabel("Time (milliseconds)", fontsize=8.5, color='#102A43')
    ax1.set_ylabel("Digital Logic Level", fontsize=8.5, color='#102A43')
    ax1.set_yticks([0.5, 2.9, 5.3])
    ax1.set_yticklabels(['900 RPM', '600 RPM', '300 RPM'], fontsize=8)
    ax1.legend(loc='upper right', fontsize=7.5, framealpha=0.9)
    ax1.grid(True, linestyle=':', alpha=0.6)

    # Subplot 2: RPM vs Time Profile
    ax2 = axs[0, 1]
    ax2.set_facecolor('#F8FAFC')
    time_pts = np.array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
    target_rpm = np.array([0, 0, 300, 300, 300, 600, 600, 600, 900, 900, 900, 0])
    calc_rpm = np.array([0, 0, 300, 300, 300, 600, 600, 600, 900, 900, 900, 0])
    
    # Continuous step curve for visual
    t_step = np.repeat(time_pts, 2)[1:]
    rpm_step = np.repeat(target_rpm, 2)[:-1]

    ax2.plot(t_step, rpm_step, color='#627D98', lw=2, linestyle='--', label='Configured Simulated RPM')
    ax2.step(time_pts, calc_rpm, where='mid', color='#0B1B3D', lw=2.2, label='Estimated Speed from Pulses')
    ax2.scatter(time_pts, calc_rpm, color='#00B4D8', s=40, zorder=5, label='Sample Interval (Δt = 1.0s)')
    ax2.set_title("Simulated Motor Speed vs Time", fontsize=10, fontweight='bold', color='#0B1B3D')
    ax2.set_xlabel("Time (seconds)", fontsize=8.5, color='#102A43')
    ax2.set_ylabel("Speed (RPM)", fontsize=8.5, color='#102A43')
    ax2.set_ylim(-50, 1050)
    ax2.legend(loc='upper left', fontsize=7.5, framealpha=0.9)
    ax2.grid(True, linestyle=':', alpha=0.6)

    # Subplot 3: Angular Velocity vs Time
    ax3 = axs[1, 0]
    ax3.set_facecolor('#F8FAFC')
    omega_vals = (2 * np.pi * target_rpm) / 60.0
    omega_step = (2 * np.pi * rpm_step) / 60.0

    ax3.plot(t_step, omega_step, color='#0F3E6D', lw=2.2, label='Estimated Angular Velocity ω')
    ax3.scatter(time_pts, omega_vals, color='#00B4D8', s=35, zorder=5)
    ax3.set_title("Simulated Angular Velocity vs Time", fontsize=10, fontweight='bold', color='#0B1B3D')
    ax3.set_xlabel("Time (seconds)", fontsize=8.5, color='#102A43')
    ax3.set_ylabel("Angular Velocity ω (rad/s)", fontsize=8.5, color='#102A43')
    ax3.set_ylim(-5, 110)
    # Annotate test speed points
    ax3.annotate("31.42 rad/s\n(300 RPM)", xy=(3.5, 31.42), xytext=(2.2, 55),
                 arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.2), fontsize=7.5, fontweight='bold', color='#0B1B3D')
    ax3.annotate("62.83 rad/s\n(600 RPM)", xy=(6.5, 62.83), xytext=(5.2, 85),
                 arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.2), fontsize=7.5, fontweight='bold', color='#0B1B3D')
    ax3.annotate("94.25 rad/s\n(900 RPM)", xy=(9.5, 94.25), xytext=(8.2, 105),
                 arrowprops=dict(arrowstyle="->", color="#0B1B3D", lw=1.2), fontsize=7.5, fontweight='bold', color='#0B1B3D')
    ax3.legend(loc='upper left', fontsize=7.5, framealpha=0.9)
    ax3.grid(True, linestyle=':', alpha=0.6)

    # Subplot 4: Linearity & Correlation Curve
    ax4 = axs[1, 1]
    ax4.set_facecolor('#F8FAFC')
    test_rpms = np.array([0, 150, 300, 450, 600, 750, 900])
    counted_n = (test_rpms * 20 * 1.0) / 60.0

    ax4.plot(test_rpms, counted_n, color='#008080', lw=2.2, marker='o', markersize=6, markerfacecolor='#00E5FF', label='Pulse Count N = (RPM×PPR×Δt)/60')
    ax4.set_title("Pulse Count Linearity & Speed Correlation", fontsize=10, fontweight='bold', color='#0B1B3D')
    ax4.set_xlabel("Rotational Speed (RPM)", fontsize=8.5, color='#102A43')
    ax4.set_ylabel("Counted Pulses per Second (N)", fontsize=8.5, color='#102A43')
    ax4.text(350, 80, "PPR = 20 pulses/rev\nΔt = 1.0 second\nLinearity: R² = 1.000 (Exact Model)",
             fontsize=8, color='#0B1B3D', bbox=dict(boxstyle='round,pad=0.5', facecolor='#EBF8FF', edgecolor='#00B4D8'))
    ax4.legend(loc='upper left', fontsize=7.5, framealpha=0.9)
    ax4.grid(True, linestyle=':', alpha=0.6)

    plt.tight_layout()
    plt.savefig("assets/slide10_simulation_plots.png", dpi=300, facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()
    print("Generated assets/slide10_simulation_plots.png")

if __name__ == "__main__":
    create_slide1_concept()
    create_slide5_architecture()
    create_slide8_flowchart()
    create_slide10_plots()
    print("All visual assets successfully generated!")
