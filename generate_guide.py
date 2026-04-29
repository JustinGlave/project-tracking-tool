from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_RIGHT

BLUE       = colors.HexColor('#487cff')
LIGHT_BLUE = colors.HexColor('#f4f6ff')
AMBER_BG   = colors.HexColor('#fff8e1')
AMBER_BDR  = colors.HexColor('#ffc107')
GREEN_BG   = colors.HexColor('#e8f5e9')
GREEN_BDR  = colors.HexColor('#4caf50')
LIGHT_GREY = colors.HexColor('#e0e0e0')
STRIPE     = colors.HexColor('#f7f8ff')
GREY       = colors.HexColor('#666666')
FOOTNOTE   = colors.HexColor('#999999')

styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, parent=styles['Normal'], **kw)

title_s  = S('T',  fontSize=18, fontName='Helvetica-Bold')
sub_s    = S('Su', fontSize=9,  fontName='Helvetica', textColor=GREY, alignment=TA_RIGHT)
h2_s     = S('H2', fontSize=10, fontName='Helvetica-Bold', textColor=BLUE,
              spaceBefore=14, spaceAfter=2, leading=14)
h3_s     = S('H3', fontSize=9.5, fontName='Helvetica-Bold', textColor=BLUE, spaceAfter=3)
body_s   = S('B',  fontSize=9.5, leading=14)
bull_s   = S('Bu', fontSize=9.5, leading=14, leftIndent=10)
note_s   = S('N',  fontSize=8.5, leading=13, textColor=colors.HexColor('#7a5c00'))
tip_s    = S('Tp', fontSize=8.5, leading=13, textColor=colors.HexColor('#2e6e30'))
foot_s   = S('F',  fontSize=8,  textColor=FOOTNOTE)
foot_r   = S('FR', fontSize=8,  textColor=FOOTNOTE, alignment=TA_RIGHT)
th_s     = S('TH', fontSize=9,  fontName='Helvetica-Bold', textColor=colors.white)


def header_block(subtitle_text):
    data = [[
        Paragraph('Project Tracking Tool', title_s),
        Paragraph(subtitle_text, sub_s),
    ]]
    t = Table(data, colWidths=[4*inch, 2.8*inch])
    t.setStyle(TableStyle([
        ('VALIGN',      (0,0),(-1,-1),'BOTTOM'),
        ('LINEBELOW',   (0,0),(-1,-1), 3, BLUE),
        ('BOTTOMPADDING',(0,0),(-1,-1), 6),
    ]))
    return t


def h2(text):
    return [
        Spacer(1, 4),
        Paragraph(text, h2_s),
        HRFlowable(width='100%', thickness=1, color=LIGHT_GREY, spaceAfter=4),
    ]


def callout(text, bg, border, ts):
    t = Table([[Paragraph(text, ts)]], colWidths=[6.8*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0),(-1,-1), bg),
        ('LEFTPADDING',  (0,0),(-1,-1), 10),
        ('RIGHTPADDING', (0,0),(-1,-1), 10),
        ('TOPPADDING',   (0,0),(-1,-1), 7),
        ('BOTTOMPADDING',(0,0),(-1,-1), 7),
        ('LINEBEFORE',   (0,0),(0,-1),  4, border),
    ]))
    return t


def note(text):
    return callout(text, AMBER_BG, AMBER_BDR, note_s)


def tip(text):
    return callout(text, GREEN_BG, GREEN_BDR, tip_s)


def two_cards(l_title, l_items, r_title, r_items):
    def cell(title, items):
        return [Paragraph(title, h3_s)] + [Paragraph(f'• {i}', bull_s) for i in items]

    t = Table([[cell(l_title, l_items), cell(r_title, r_items)]],
              colWidths=[3.3*inch, 3.3*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0),(-1,-1), LIGHT_BLUE),
        ('LEFTPADDING',  (0,0),(-1,-1), 10),
        ('RIGHTPADDING', (0,0),(-1,-1), 10),
        ('TOPPADDING',   (0,0),(-1,-1), 8),
        ('BOTTOMPADDING',(0,0),(-1,-1), 8),
        ('LINEAFTER',    (0,0),(0,-1),  1, LIGHT_GREY),
        ('VALIGN',       (0,0),(-1,-1), 'TOP'),
        ('LINEBEFORE',   (0,0),(0,-1),  4, BLUE),
        ('LINEBEFORE',   (1,0),(1,-1),  4, BLUE),
    ]))
    return t


def data_table(rows):
    header = [Paragraph(h, th_s) for h in ('Action / Feature', 'How To')]
    data   = [header] + [[Paragraph(a, body_s), Paragraph(b, body_s)] for a, b in rows]
    t = Table(data, colWidths=[2.2*inch, 4.6*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0),  BLUE),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [colors.white, STRIPE]),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('GRID',          (0,0),(-1,-1), 0.5, LIGHT_GREY),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
    ]))
    return t


def footer(page_num):
    t = Table(
        [[Paragraph('ATS Inc. — Project Tracking Tool', foot_s),
          Paragraph(f'Page {page_num} of 2', foot_r)]],
        colWidths=[3.4*inch, 3.4*inch],
    )
    t.setStyle(TableStyle([
        ('LINEABOVE',    (0,0),(-1,-1), 1, LIGHT_GREY),
        ('TOPPADDING',   (0,0),(-1,-1), 5),
    ]))
    return t


# ─────────────────────────────────────────── STORY ──
story = []

# ══ PAGE 1 ══════════════════════════════════════════
story.append(header_block('ATS Inc. — Controls Division<br/>Installation &amp; Setup Guide'))
story.append(Spacer(1, 10))

story += h2('INSTALLATION')
story.append(Paragraph('1.  Download <b>ProjectTrackingToolSetup.exe</b> from the link provided by your manager.', body_s))
story.append(Paragraph('2.  Run the installer — no admin rights required. A desktop shortcut will be created automatically.', body_s))
story.append(Paragraph('3.  Launch the app from your desktop shortcut.', body_s))
story.append(Spacer(1, 6))
story.append(note('<b>Note:</b> Your project data is stored separately from the app and is never lost during updates or reinstalls.'))
story.append(Spacer(1, 6))

story += h2('FIRST-TIME SETUP')
story.append(two_cards(
    'Step 1 — Shared Data Location',
    [
        'Go to <b>File &gt; Data Location...</b>',
        'Click <b>Browse</b> and navigate to the shared OneDrive folder:<br/><i>OneDrive - ATS Inc › Phoenix Project Tracker</i>',
        'Click <b>OK</b>. When prompted, choose to copy your existing data if this is your first time.',
    ],
    'Step 2 — Financial Data File',
    [
        'Go to <b>File &gt; Financial Data File...</b>',
        'Click <b>Browse</b> and navigate to the same shared OneDrive folder.',
        'Select the <b>.xlsb</b> tracking workbook and click <b>Open</b>.',
        'Financial data will now appear on all your jobs.',
    ]
))
story.append(Spacer(1, 8))
story.append(tip('<b>Tip:</b> Both the shared database and the financial file live in the same OneDrive folder — <i>OneDrive - ATS Inc › Phoenix Project Tracker</i>. You only need to set these paths once.'))
story.append(Spacer(1, 6))

story += h2('AUTO-UPDATES')
story.append(Paragraph(
    'The app checks for updates automatically on startup. When a new version is available, '
    'a banner will appear at the top of the window. Click <b>Install &amp; Restart</b> to update — your data is never affected.',
    body_s))
story.append(Spacer(1, 6))

story += h2('SHARED DATA — HOW IT WORKS')
story.append(Paragraph(
    'All team members read from and write to the same project file stored in the shared OneDrive folder. '
    'OneDrive keeps the file in sync across all machines. Because multiple people may be editing at the same time, '
    'follow this simple rule:',
    body_s))
story.append(Spacer(1, 4))
story.append(note('<b>Best practice:</b> Avoid having two people edit the <i>same job</i> simultaneously. Editing different jobs at the same time is fine.'))
story.append(Spacer(1, 16))
story.append(footer(1))
story.append(PageBreak())

# ══ PAGE 2 ══════════════════════════════════════════
story.append(header_block('ATS Inc. — Controls Division<br/>User Guide'))
story.append(Spacer(1, 10))

story += h2('MANAGING PROJECTS')
story.append(data_table([
    ('<b>Create a new job</b>',    'Click <b>New</b> in the sidebar and fill in the job details'),
    ('<b>Edit a job</b>',          'Select the job, then click <b>Edit</b> in the sidebar'),
    ('<b>Delete a job</b>',        'Select the job, then click <b>Delete</b> in the sidebar'),
    ('<b>Import from email</b>',   'Click <b>Import Email</b> and select an ODIN assignment <i>.eml</i> file'),
    ('<b>Search jobs</b>',         'Type in the search bar at the top of the sidebar — searches job name, number, PM, and SE'),
    ('<b>Sort jobs</b>',           'Use the sort dropdown (Last Updated / Name / Job Number) and the ↑↓ toggle'),
]))
story.append(Spacer(1, 8))

story += h2('TASKS')
story.append(Paragraph('Each job has a task list organized by phase and matching the segmented progress bar.', body_s))
for item in [
    'Check the box to mark a task complete',
    'Click <b>Add Task</b> to add a custom task',
    'Use <b>Templates</b> to apply a Standard or Phoenix task set to any job',
    'Right-click a task to edit it, delete it, or open Notes History',
    'Use <b>View &gt; Compact Task View</b> to shrink rows and hide the Notes column',
]:
    story.append(Paragraph(f'• {item}', bull_s))
story.append(Spacer(1, 8))

story += h2('FINANCIAL DATA (ODIN)')
story.append(Paragraph('Financial data is pulled from the shared ODIN tracking Excel file and displayed per job.', body_s))
story.append(Spacer(1, 6))
story.append(two_cards(
    'Viewing Financials',
    [
        'Each job card in the sidebar shows: contract value, actual margin, and delta vs. booked margin',
        'Click <b>Financials</b> (next to Project Info) for the full breakdown',
        'The <b>Project Info</b> dialog also shows a financial summary at the bottom',
        'The <i>"ODIN data as of"</i> timestamp below the job list shows the last refresh time',
    ],
    'Reading the Margins',
    [
        '<font color="#4caf50"><b>Green ▲</b></font> — Actual margin 2%+ better than booked',
        '<font color="#ff9800"><b>Yellow</b></font> — Within ±2% of booked margin',
        '<font color="#f44336"><b>Red ▼</b></font> — Actual margin 2%+ worse than booked',
        'Data auto-refreshes every <b>5 minutes</b>, or click <b>Refresh</b> in the Financials dialog',
    ]
))
story.append(Spacer(1, 8))

story += h2('OTHER FEATURES')
story.append(data_table([
    ('<b>Notes</b>',         'Click <b>Notes</b> on any job — add open/closed notes per project'),
    ('<b>Change Orders</b>', 'Click <b>Change Orders</b> — track COs with amounts and approval status'),
    ('<b>Export to Excel</b>', '<b>File &gt; Export to Excel</b> — exports the selected job\'s tasks and details'),
    ('<b>Compact Task View</b>', '<b>View &gt; Compact Task View</b> — toggles a denser task table layout'),
    ('<b>Project Info</b>',  'Click <b>Project Info</b> — view all job details including GC, warranty period, and ODIN financials'),
]))
story.append(Spacer(1, 8))
story.append(tip('<b>Tip:</b> If financial data shows "not found" for a job, make sure the job number in the tool exactly matches the job number in the ODIN tracking file.'))
story.append(Spacer(1, 16))
story.append(footer(2))

doc = SimpleDocTemplate(
    'Project Tracking Tool - User Guide.pdf',
    pagesize=letter,
    leftMargin=0.75*inch, rightMargin=0.75*inch,
    topMargin=0.65*inch,  bottomMargin=0.65*inch,
)
doc.build(story)
print('Done.')
