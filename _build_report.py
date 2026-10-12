# -*- coding: utf-8 -*-
"""Build an organized lab report DOCX from the raw exercise file + its screenshots."""
import os
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = '_media'
PREP = '_print'
os.makedirs(PREP, exist_ok=True)

# flatten transparency onto white so screenshots print cleanly
for f in os.listdir(SRC):
    if not f.endswith('.png'):
        continue
    im = Image.open(os.path.join(SRC, f))
    if im.mode in ('RGBA', 'LA', 'P'):
        im = im.convert('RGBA')
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert('RGB')
    im.save(os.path.join(PREP, f))

doc = Document()

# ---------- page & base styles ----------
sec = doc.sections[0]
sec.left_margin = Inches(0.85)
sec.right_margin = Inches(0.85)
sec.top_margin = Inches(0.8)
sec.bottom_margin = Inches(0.8)

normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(11)
normal.paragraph_format.space_after = Pt(6)

for name, size in (('Heading 1', 15), ('Heading 2', 12.5)):
    st = doc.styles[name]
    st.font.name = 'Calibri'
    st.font.size = Pt(size)
    st.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    st.font.bold = True

CONTENT_W = sec.page_width - sec.left_margin - sec.right_margin

fig_no = [0]


def para(text='', style=None, size=None, bold=False, italic=False, align=None,
         space_before=None, space_after=None, color=None, indent=None):
    p = doc.add_paragraph(style=style)
    if text:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    if align is not None:
        p.alignment = align
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if indent is not None:
        p.paragraph_format.left_indent = Inches(indent)
    return p


def bullet(text):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    return p


def question(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run('Q. ')
    r.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    p.add_run(text)
    return p


def answer(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run('A. ')
    r.bold = True
    p.add_run(text)
    return p


def note(text):
    return para(text, size=9.5, italic=True,
                color=RGBColor(0x44, 0x44, 0x44), indent=0.25)


def figure(fname, caption, width_in=5.9):
    fig_no[0] += 1
    path = os.path.join(PREP, fname)
    # never upscale a screenshot: cap the printed width at its native pixels / 150 dpi
    px_w = Image.open(path).size[0]
    width_in = min(width_in, px_w / 150.0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Inches(width_in))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(12)
    r = c.add_run('Figure %d: ' % fig_no[0])
    r.bold = True
    r.font.size = Pt(9.5)
    r2 = c.add_run(caption)
    r2.font.size = Pt(9.5)
    r2.italic = True
    return fig_no[0]


def table(headers, rows, caption):
    t = doc.add_table(rows=1, cols=len(headers))
    try:
        t.style = 'Light Grid Accent 1'
    except KeyError:
        t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        r = hdr[i].paragraphs[0].add_run(h)
        r.bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(12)
    r = cap.add_run(caption)
    r.font.size = Pt(9.5)
    r.italic = True
    return t


# ---------- footer with page number ----------
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run('Lab 0 - Getting Familiar  |  Page ')
run.font.size = Pt(9)
run2 = footer_p.add_run()
run2.font.size = Pt(9)
f1 = OxmlElement('w:fldChar')
f1.set(qn('w:fldCharType'), 'begin')
it = OxmlElement('w:instrText')
it.set(qn('xml:space'), 'preserve')
it.text = 'PAGE'
f2 = OxmlElement('w:fldChar')
f2.set(qn('w:fldCharType'), 'end')
run2._r.append(f1)
run2._r.append(it)
run2._r.append(f2)

# ================= title block =================
para('Date: ___________________          Name: ___________________________________',
     size=10)
para()
t = para('LABORATORY EXERCISE 0', size=20, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=0,
         color=RGBColor(0x1F, 0x38, 0x64))
para('GETTING FAMILIAR', size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     space_after=2, color=RGBColor(0x1F, 0x38, 0x64))
para('Process Control Lab - Simulation software: PC-ControLAB 3.1',
     size=10.5, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(10)
r = p.add_run('OBJECTIVE:  ')
r.bold = True
p.add_run('To become familiar with the simulation software.')

# ================= 1 =================
doc.add_heading('1.  Running the Program', level=1)
bullet('Run PC-ControLAB.')
bullet('In the installation folder, locate and read the file "QuickStartGuide.pdf".')
bullet('Slow down the simulation speed via View | Scroll Rate and select a mid-value, '
       'to have a more practical experience.')

# ================= 2 =================
doc.add_heading('2.  The Feedback Loop', level=1)
para('Retrieve the strategy via Control | Retrieve Strategy, Model and Tuning and select '
     '"Feedbck1" (in C:\\Users\\Public\\PC-ControLAB 3.1).', space_after=2)
para('The simulation starts with an initial controller output (OUT) of 35%. You can press '
     'PAUSE at any time and use the cursor to read the values. To rerun the simulation, '
     'press RUN.', space_after=10)

# ---- 2.1 ----
doc.add_heading('2.1  Initial Values', level=2)
f = figure('image1.png',
           'PC-ControLAB main window at start-up (MAN): Feed Rate trend with PV and OUT, '
           'and the readout panel showing Time = 36.70 min, PV1 = 275.00, Output1 = 35.00.')
question('What is the process variable (PV) current value?')
answer('275')
question('What is its range?')
answer('From 0 to 500 DegF')
question('What is the controller output range?')
answer('From 0 to 100 %')
question('What is the setpoint (SP)?')
answer('275 DegF')
table(['Quantity', 'Value'],
      [['Process variable PV (current value)', '275'],
       ['PV range', 'From 0 to 500 DegF'],
       ['Controller output range', 'From 0 to 100 %'],
       ['Setpoint SP', '275 DegF']],
      'Table 1: Initial loop values (see Figure 1).')

# ---- 2.2 ----
doc.add_heading('2.2  Manual Mode: Step Change in OUT', level=2)
para('With the controller in MAN, set the controller output to 20% and observe the '
     'process behavior. The continuously scrolling display records the previous 60 '
     '(simulated) minutes of the response; the trend that scrolls past the left display '
     'border cannot be retrieved.')
question('Can the process be described as static or dynamic?')
answer('Dynamic')
question('At which value does PV settle to?')
answer('162.50')
question('After how long does PV appear to settle at this value?')
answer('24.7 minutes')
question('Calculate the time elapsed since the controller OUT was changed to the time the '
         'PV becomes reasonably close to its final value (the usual definition of '
         '"reasonably close" is when 98% of the change has occurred).')
answer('0.98 * (275 - 162.5) = 110.25 , measure till (275 - 110.25 = 164.75)')
figure('image2.png',
       'Manual mode (OUT = 20%): PV decays from 275 DegF and settles near 162 DegF '
       '(readout: Time = 56.55 min, PV1 = 162.52, Output1 = 20.00).')
figure('image3.png',
       'Manual mode, zoomed view used to time the 98% settling point '
       '(readout: Time = 22.95 min, PV1 = 164.81, Output1 = 20.00).')

# ---- 2.3 ----
doc.add_heading('2.3  Load Disturbance in Manual Mode', level=2)
para('Change the LOAD by pressing StepIncr or StepDecr twice in rapid succession.')
question('What do you observe?')
answer('The PV decreases when decreasing the load and increases when increasing the load; '
       'since we are in manual mode the OUT does not change.')
figure('image4.png',
       'Manual mode (OUT = 20%), response to a load change: PV falls further, to about 88 DegF '
       '(readout: Time = 59.40 min, PV1 = 87.72, Output1 = 20.00).')

# ---- 2.4 ----
doc.add_heading('2.4  Automatic Mode', level=2)
para('Switch the controller to AUTO.')
question('Does the controller bring the PV back to the setpoint?')
answer('Yes, but after a lot of time and there was hunting in more than 2 clock cycles.')
figure('image5.png',
       'AUTO mode: PV hunting around SP after the controller takes over from MAN '
       '(readout: Time = 32.75 min, PV1 = 326.10, Output1 = 18.57).')
figure('image6.png',
       'AUTO mode: PV brought back to the setpoint with only small residual oscillation '
       '(readout: Time = 43.90 min, PV1 = 275.82, Output1 = 44.16).')

# ---- 2.5 ----
doc.add_heading('2.5  Setpoint Change to 350 \u00b0F', level=2)
para('Increase SP to 350 degF and observe the response.')
question('Does it exhibit more time variability or less compared to the process response '
         'with no feedback (i.e., when the controller was in MAN)?')
answer('It exhibits more variability. In MAN, the PV smoothly decayed to a new value '
       'without oscillating. In AUTO, the controller fights the error, causing the PV to '
       'swing up and down repeatedly till it settles.')
question('Do the oscillations die out?  (If so, they are called DAMPED oscillations.)')
answer('Yes')
question('Does the PV reach the new SP?')
answer('Yes , PV = 350.')
question('In how many cycles?')
answer('4 cycles')
figure('image7.png',
       'SP increased to 350 degF: closed-loop response with the PV swinging above the new '
       'setpoint (readout: Time = 41.35 min, PV1 = 373.87, Output1 = 40.02).')
figure('image8.png',
       'SP = 350 degF: damped oscillation of the PV about the new setpoint, with the '
       'response parameters marked on the trend '
       '(readout: Time = 22.10 min, PV1 = 342.41, Output1 = 60.17).')
question('What is the PV overshoot?  (see the response-parameter figure at the back of the exercise)')
answer('(410 - 350) / (350 - 275) = 17.14 %')
question('What is the period of oscillation?')
answer('38 - 22 = 16 minutes')
question('Does the controller do a good job?')
answer('No, as we notice it takes more than 1 cycle to settle.')
note('For a damped response, the industry standard is that a well-tuned controller will '
     'bring the PV to the SP in at most one cycle and with an overshoot not exceeding 5%.')
table(['Parameter', 'Value'],
      [['a  (peak above final value)', '410 - 350 = 60 DegF'],
       ['b  (total change: final value - initial value)', '350 - 275 = 75 DegF'],
       ['Overshoot  OS = a / b * 100%', '17.14 %'],
       ['Period of oscillation', '16 minutes']],
      'Table 2: Response parameters (a and b read from the traces in Figures 7 and 8).')

# ---- 2.6 ----
doc.add_heading('2.6  Load Disturbance in Automatic Mode', level=2)
para('Change the LOAD by pressing StepIncr or StepDecr twice in rapid succession.')
question('Does the controller manage to bring the PV back to the setpoint?')
answer('Yes, it settles again fast.')
figure('image9.png',
       'AUTO mode: PV held at the setpoint after a load disturbance '
       '(readout: Time = 21.45 min, PV1 = 355.16, Output1 = 46.56).')

# ---- 2.7 ----
doc.add_heading('2.7  Fluctuating Load (AutoLoad)', level=2)
para('Switch the controller to MAN. Allow the LOAD to fluctuate like real life situations '
     '(AutoLoad). Try to keep the PV as close as possible to the SP (keep changing the '
     'controller output up and down for few minutes).')
question('How easy is this task?')
answer('Very hard, almost cannot get an ideal value.')
figure('image10.png',
       'Manual control under fluctuating load (AutoLoad): the noisy LOAD trace disturbs the '
       'PV while OUT is held constant (readout: Time = 30.70 min, PV1 = 343.86, '
       'Output1 = 44.93).')
figure('image11.png',
       'Manual control under fluctuating load (AutoLoad), second capture showing the PV '
       'drifting away from SP (readout: Time = 18.85 min, PV1 = 338.25, Output1 = 53.96).')
para('While the LOAD is still fluctuating, switch the controller to AUTO.')
question('Observe PV fluctuations around SP. Do they seem to be acceptable for operation, '
         'i.e., confined within a relatively small range (say 1% or even 2% of the PV range)?')
answer('Yes, it is acceptable.')
figure('image12.png',
       'Automatic control under fluctuating load (AutoLoad): the PV stays close to the '
       'setpoint while OUT moves to absorb the disturbance '
       '(readout: Time = 28.50 min, PV1 = 343.12, Output1 = 51.76).')

# ---- 2.8 ----
doc.add_heading('2.8  Undamped Oscillations', level=2)
para('Switch off AutoLoad. Keep the controller in AUTO. Change the controller tuning '
     'parameters such that the PV exhibits sustained (UNDAMPED) oscillations. Hint: click '
     'TUNE then increase the controller gain. Record this gain.')
question('Controller gain for sustained oscillations')
answer('Gain = 3')
figure('image13.png',
       'Sustained (undamped) oscillations obtained with the controller gain set to 3: the PV '
       'oscillates with constant amplitude and OUT swings over its full range.')
table(['Quantity', 'Value'],
      [['Controller gain for sustained oscillations', '3']],
      'Table 3: Gain at sustained oscillations.')

# ---- 2.9 ----
doc.add_heading('2.9  Low Gain', level=2)
para('Reduce the controller gain to 0.1. Change the setpoint again.')
question('What do you observe?')
answer('No oscillation; it settles to the new SP fast and perfectly.')
figure('image14.png',
       'Controller gain reduced to 0.1: the PV moves to the new setpoint without any '
       'oscillation (readout: Time = 10.75 min, PV1 = 267.47, Output1 = 49.93).')

# ================= 3 =================
doc.add_heading('3.  Conclusion', level=1)
para('[Draft summary of the observations recorded above - edit or expand as needed.]',
     size=9.5, italic=True, color=RGBColor(0x44, 0x44, 0x44))
para('The process is dynamic: after a step in the controller output in manual mode the PV '
     'did not change instantly but moved to a new steady value (275 DegF to 162.50 DegF) '
     'over roughly 25 minutes. In manual mode the controller output is fixed, so any load '
     'change moves the PV directly, and keeping the PV at the setpoint by hand under a '
     'fluctuating load was very hard. In automatic mode the controller brought the PV back '
     'to the setpoint and rejected the same load disturbance, although with the initial '
     'tuning it took more than one cycle and overshot by 17.14%. The tuning test showed why: '
     'at a gain of 3 the loop gave sustained (undamped) oscillations, while at a gain of 0.1 '
     'the PV reached the new setpoint without oscillating at all. The best behaviour is '
     'therefore obtained somewhere between these two gains, and the loop as tuned in this '
     'exercise does not meet the industry standard of settling in at most one cycle with an '
     'overshoot below 5%.')

# ================= reference figure =================
doc.add_page_break()
doc.add_heading('Reference Figure: Determining Response Parameters', level=1)
fig_no[0] += 1
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(os.path.join(PREP, 'image15.png'), width=Inches(3.6))
c = doc.add_paragraph()
c.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = c.add_run('Response-parameter figure supplied with the exercise (a, b, c, Period P, '
              'tp, ts).')
r.font.size = Pt(9.5)
r.italic = True
para()
pp = doc.add_paragraph()
pp.add_run('Overshoot:            OS  =  a / b %').bold = True
note('a = peak value above the final value, b = total change (final value - initial value).')

out = 'lab0_getting_familiar_organized.docx'
doc.save(out)
print('saved', out, os.path.getsize(out), 'bytes; figures =', fig_no[0])
