#!/usr/bin/env python3
"""
Generate DG_Analysis_Report.docx — Word version of the DG Sample Code Check analysis report.

Uses python-docx for document generation and mermaid.ink for diagram images.
"""

import base64
import io
import os
import urllib.parse
from datetime import datetime

import requests
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_cell_shading(cell, color_hex):
    """Set background color of a table cell."""
    shading = cell._element.get_or_add_tcPr()
    shd = shading.makeelement(qn('w:shd'), {
        qn('w:val'): 'clear',
        qn('w:color'): 'auto',
        qn('w:fill'): color_hex,
    })
    shading.append(shd)


def add_styled_table(doc, headers, rows, col_widths=None):
    """Add a table with header row styling and alternating row colors."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in paragraph.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(9)
        set_cell_shading(cell, '1A5276')

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(cell_text)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
            if r_idx % 2 == 1:
                set_cell_shading(cell, 'F2F7FB')

    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(width)

    return table


def fetch_mermaid_image(mermaid_code):
    """Fetch a PNG image from mermaid.ink for the given Mermaid code."""
    encoded = base64.urlsafe_b64encode(mermaid_code.encode('utf-8')).decode('utf-8')
    url = f'https://mermaid.ink/img/{encoded}'
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code == 200:
            return io.BytesIO(resp.content)
    except Exception as e:
        print(f'  Warning: Could not fetch Mermaid diagram: {e}')
    return None


def add_mermaid_diagram(doc, mermaid_code, caption='', width=6.0):
    """Add a Mermaid diagram as an embedded PNG image, or fallback to text description."""
    img_data = fetch_mermaid_image(mermaid_code)
    if img_data:
        doc.add_picture(img_data, width=Inches(width))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p = doc.add_paragraph()
        p.add_run('[Diagram could not be rendered. Mermaid code:]').italic = True
        doc.add_paragraph(mermaid_code, style='No Spacing')
    if caption:
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.runs[0].italic = True
        cap.runs[0].font.size = Pt(9)


def add_code_block(doc, code_text):
    """Add a code block styled paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(code_text)
    run.font.name = 'Consolas'
    run.font.size = Pt(8)
    pf = p.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)


# ---------------------------------------------------------------------------
# Mermaid diagram definitions
# ---------------------------------------------------------------------------

MERMAID_DEPENDENCY = """graph TD
    subgraph "Primary Data Records"
        EE["EMPLOYEES"]
        ER["EMPLOYERS"]
        EO["EROPTIONS"]
        DEP["DEPENDENTS"]
    end
    subgraph "Billing Records"
        ERADJ["ERADJ"]
    end
    subgraph "Correspondence Records"
        CDR["CORRESP.DATA.REC (in-memory)"]
        CT["CORRESP.TRACKING"]
    end
    subgraph "Audit Records"
        SL["SAVEDLISTS"]
        SM["SYSMGMT"]
    end
    subgraph "FI-Only Records"
        CL["CERTS.LAYOUT"]
        CM["CERTS.MERGE"]
        CT2["CERTS.TEXT2"]
        LC["LETTER.COPIES"]
        OC["OUTPUT.CONTROL"]
        GI["GROUPINFO"]
    end
    EE -- "EEFILENO" --> ER
    EE -- "EEFILENO" --> EO
    EE -- "EEID.DEPNO" --> DEP
    ER -- "ERBILLEDTO" --> ERADJ
    EO -- "EOCOBRA.NOTICE / EOHIPAAP" --> ERADJ
    EE --> CDR
    ER --> CDR
    DEP --> CDR
    CDR --> CT
    ER -- "ERTRUST" --> GI
    GI -- "GINPP" --> OC
    ER -- "ERINS.CARRIER" --> CL
    CL --> CM
    CM --> CT2"""

MERMAID_COUPLING = """graph LR
    subgraph "COBRAN"
        C1["Eligibility Filtering"]
        C2["Employer/Group Loading"]
        C3["Letter Generation (EE)"]
        C4["Letter Generation (Dep)"]
        C5["Billing (ERADJ)"]
        C6["Audit Lists"]
    end
    subgraph "HIPAA.NPP"
        H1["Eligibility Filtering"]
        H2["Employer/Group Loading"]
        H3["SF/FI Routing"]
        H4["SmartComm Submission"]
        H5["Legacy PCL Generation"]
        H6["Billing (ERADJ)"]
        H7["FTP Transmission"]
        H8["Audit Lists"]
        H9["Letter Storage"]
    end
    subgraph "PRINT.DISABLED.DEP.LETTERS"
        P1["Eligibility Filtering"]
        P2["Dep Need Validation"]
        P3["Letter Generation"]
        P4["Dep Record Update"]
        P5["Error Reporting/Email"]
        P6["Security Check"]
    end"""

MERMAID_DECOUPLED = """graph TD
    COBRAN["COBRAN (orchestrator)"]
    HIPAA["HIPAA.NPP (orchestrator)"]
    PDDL["PRINT.DISABLED.DEP.LETTERS (orchestrator)"]
    EF["*CHECK.LETTER.ELIGIBILITY"]
    EC["*LOAD.EMPLOYER.CONTEXT"]
    CS["*SUBMIT.LETTER.REQUEST"]
    BA["*ADD.NOTICE.BILLING"]
    FI["*GENERATE.FI.HIPAA.NOTICE"]
    FTP["*SUBMIT.PRINT.FTP"]
    AL["*SAVE.PROCESS.AUDIT"]
    DU["*UPDATE.DEP.LETTER.STATUS"]
    COBRAN --> EF
    COBRAN --> EC
    COBRAN --> CS
    COBRAN --> BA
    COBRAN --> AL
    HIPAA --> EF
    HIPAA --> EC
    HIPAA --> CS
    HIPAA --> BA
    HIPAA --> FI
    HIPAA --> FTP
    HIPAA --> AL
    PDDL --> EF
    PDDL --> CS
    PDDL --> DU
    PDDL --> EC"""

MERMAID_COBRAN_FLOW = """graph TD
    START["COBRAN Start"] --> INIT["INIT.VARS"]
    INIT --> SELECT["SELECT EMPLOYEES with NEED='C'"]
    SELECT --> LOOP["MAIN.LOOP: READNEXT EEID"]
    LOOP --> READ_EE["Read EMPLOYEES record"]
    READ_EE --> CHECK_NEED["Check EENEED for 'X'"]
    CHECK_NEED --> LOCATE_EFF["LOCATE effective date index"]
    LOCATE_EFF --> LOAD_ER["Load EMPLOYERS + EROPTIONS (if FILENO changed)"]
    LOAD_ER --> STATUS_CHECK["Check status, termination, COBRAFEE"]
    STATUS_CHECK --> DEP_CHECK["Check dependent COBRA eligibility"]
    DEP_CHECK -->|"Spouse at diff addr"| SUBMIT_DP["SUBMIT.CORRESP.REQUEST (DP template)"]
    DEP_CHECK -->|"Employee notice"| SUBMIT_EE["SUBMIT.CORRESP.REQUEST (EE template)"]
    SUBMIT_DP --> BILLING["DO.BILLING (ERADJ with COBRACHG)"]
    SUBMIT_EE --> BILLING
    BILLING --> LOOP
    LOOP -->|"End of list"| AUDIT["END.REPORT.FILE (save lists)"]"""

MERMAID_HIPAA_FLOW = """graph TD
    START["HIPAA.NPP Start"] --> INIT["Initialize + HIPAA.DAILY.COUNTER"]
    INIT --> FILTER["FILTER.EEIDS: Build EEID.LIST"]
    FILTER --> SF_FI{"SF or FI?"}
    SF_FI -->|"Self-Funded"| SF_LOOP["SF Loop: For each EEID.DEPNO"]
    SF_FI -->|"Fully-Insured"| FI_LOOP["FI Loop: For each EEID.DEPNO"]
    SF_LOOP --> SC_SUBMIT["SUBMIT.CORRESP.REQUEST (SmartComm)"]
    SC_SUBMIT --> SF_BILLING["DO.BILLING (ERADJ with HIPAA)"]
    SF_BILLING --> SF_LOOP
    FI_LOOP --> GET_FORMS["GET.HIPAA.FORMS (CERTS.LAYOUT lookup)"]
    GET_FORMS --> GEN_FORMS["GENERATE.HIPAA.FORMS (PCL output)"]
    GEN_FORMS --> FI_BILLING["DO.BILLING (ERADJ with HIPAA)"]
    FI_BILLING --> FI_LOOP
    FI_LOOP -->|"Done"| FTP["FTP.FILE"]
    SF_LOOP -->|"Done"| SAVE["Save audit lists"]
    FTP --> SAVE"""

MERMAID_PDDL_FLOW = """graph TD
    START["PRINT.DISABLED.DEP.LETTERS Start"] --> INIT["INIT.VARS (6 letter type mappings)"]
    INIT --> SELECT["SELECT EMPLOYEES with STATUS='H'"]
    SELECT --> LOOP["MAIN.LOOP: READNEXT EEID"]
    LOOP --> VALIDATE["VALIDATE.DEP.DATA"]
    VALIDATE --> CHECK["CHECK.DEP.NEED (read DEPENDENTS)"]
    CHECK --> NEED_LOOP["For each CUR.NEED in DPDEP.NEED"]
    NEED_LOOP --> SET_TMPL["SET.SMARTCOMM.TEMPLATE (lookup in LTR.TYPES)"]
    SET_TMPL --> SUBMIT["SUBMIT.CORRESP.REQUEST"]
    SUBMIT -->|"Success"| UPDATE["UPDATE.DEP.NEED"]
    SUBMIT -->|"Error"| ERR["Add to ERROR.DEP.NEED"]
    UPDATE --> NEED_LOOP
    ERR --> NEED_LOOP
    NEED_LOOP -->|"Done"| WRITE["WRITE.DEP.REC"]
    WRITE --> LOOP
    LOOP -->|"End"| EMAIL["Email error report"]"""

# ---------------------------------------------------------------------------
# Document generation
# ---------------------------------------------------------------------------

def build_report():
    doc = Document()

    # -- Title page ---
    doc.add_heading('DG Sample Code Check', level=0)
    doc.add_heading('System Analysis Report', level=1)
    doc.add_paragraph(f'Generated on {datetime.now().strftime("%B %d, %Y")}')
    doc.add_paragraph('Healthcare Benefits Correspondence Automation System')
    doc.add_page_break()

    # ===================================================================
    # SECTION 1: Executive Summary
    # ===================================================================
    doc.add_heading('1. Executive Summary', level=1)
    doc.add_paragraph(
        'This is a healthcare benefits correspondence automation system built on the '
        'Pick/MultiValue (UniData) platform. It manages three key categories of regulatory '
        'and benefits notices:'
    )
    doc.add_paragraph('COBRA notices — Initial introduction letters and dependent notices for COBRA continuation coverage', style='List Bullet')
    doc.add_paragraph('HIPAA privacy notices — Notice of Privacy Practices for both self-funded (SF) and fully-insured (FI) groups', style='List Bullet')
    doc.add_paragraph('Disabled dependent letters — Review requests, recertifications, approvals, denials, and no-response letters for disabled dependents', style='List Bullet')

    doc.add_paragraph('The system is implemented through 3 main programs:')
    add_styled_table(doc,
        ['Program', 'Purpose', 'Output'],
        [
            ['COBRAN', 'Process COBRA initial notices', 'SmartComm correspondence for employees and dependents'],
            ['HIPAA.NPP', 'Generate HIPAA Notice of Privacy Practices', 'SmartComm (SF) or PCL print files via FTP (FI)'],
            ['PRINT.DISABLED.DEP.LETTERS', 'Generate disabled dependent letters (6 types)', 'SmartComm correspondence'],
        ])

    doc.add_paragraph(
        'These programs interact with shared data records including EMPLOYEES, EMPLOYERS, '
        'EROPTIONS, DEPENDENTS, and ERADJ, and submit correspondence through a centralized '
        '*CREATE.CORRESP.TRACKING subroutine that interfaces with the SmartComm document '
        'management system.'
    )
    doc.add_paragraph(
        'Key Finding: Significant functional coupling exists across the 3 programs, with 8+ '
        'shared concerns (eligibility filtering, employer loading, billing, audit) duplicated '
        'in each. This report provides a detailed decoupling strategy with 8 proposed '
        'extracted subroutines.'
    )

    # ===================================================================
    # SECTION 2: Record Types
    # ===================================================================
    doc.add_heading('2. Record Types — Full Analysis', level=1)

    # 2.1 EMPLOYEES
    doc.add_heading('2.1 EMPLOYEES Record', level=2)
    doc.add_paragraph('Description: The core employee/member record. Contains multi-valued status codes, effective dates, coverage fields, and demographic information.')
    doc.add_paragraph('Key Format: EEID (Employee ID, e.g., 12345)')
    doc.add_paragraph('Loaded via: $INCLUDE BP EMPLOYEES.INS → MAT EEREC')
    add_styled_table(doc,
        ['Field Name', 'Purpose'],
        [
            ['EESTATUS', 'Multi-valued status codes: A (active), P (pending), T (terminated), D (deceased), W (waived), R (retired), I (inactive)'],
            ['EENEED', 'Trigger flag: C (COBRA), N (HIPAA NPP), X (exclude/old record)'],
            ['EEFILENO', 'Foreign key to EMPLOYERS record (multi-valued, indexed by EEIDX)'],
            ['EEEFFDATE', 'Multi-valued effective dates, sorted descending'],
            ['EERELATIONSHIP', 'E (employee), S (spouse), C (child)'],
            ['EETERMDATE', 'Employer termination date'],
            ['EEMEDCOV / EEDENCOV / EEVISCOV', 'Medical, dental, and vision coverage codes'],
            ['EEHIPAA.SENT', 'Multi-valued dates of last HIPAA notice sent (indexed by DEPIDX)'],
            ['EEDEPNO', 'Multi-valued dependent numbers'],
            ['EELAST / EEFIRST', 'Employee last and first name'],
            ['EEADD1/EEADD2/EECITY/EESTATE/EEZIP', 'Employee address fields'],
        ])
    add_code_block(doc, "CALL *READREC('',MAT EEREC,EMPLOYEES,EEID,'')\nLOCATE TODAY IN EEEFFDATE<1,1> BY 'DR' SETTING EEIDX ELSE NULL")
    doc.add_paragraph('Used by: COBRAN, HIPAA.NPP, PRINT.DISABLED.DEP.LETTERS')

    # 2.2 EMPLOYERS
    doc.add_heading('2.2 EMPLOYERS Record', level=2)
    doc.add_paragraph('Description: Employer/group master record. Contains employer-level configuration for billing, termination, trust type, and insurance carrier information.')
    doc.add_paragraph('Key Format: FILENO (derived from EEFILENO<1,EEIDX>)')
    doc.add_paragraph('Loaded via: $INCLUDE BP EMPLOYERS.INS → MAT ERREC')
    add_styled_table(doc,
        ['Field Name', 'Purpose'],
        [
            ['EREFFDATE', 'Multi-valued employer effective dates'],
            ['ERCOBRAFEE', 'COBRA fee flag (Y/N/dollar amount) — determines if COBRA billing applies'],
            ['ERBILLEDTO', 'Date through which employer has been billed; used to determine next billing period'],
            ['ERACTUAL.STATE', 'Employer actual state (for jurisdiction logic)'],
            ['ERTRUST', 'Trust code — determines SF (self-funded) vs FI (fully-insured) routing'],
            ['ERTERMDATE', 'Employer termination date'],
            ['ERINS.CARRIER', 'Insurance carrier code (used for FI CERTS.LAYOUT lookup)'],
        ])
    add_code_block(doc, "IF EEFILENO<1,EEIDX> <> FILENO THEN\n   FILENO = EEFILENO<1,EEIDX>\n   CALL *READREC('',MAT ERREC,EMPLOYERS,FILENO,'')\nEND")
    doc.add_paragraph('Used by: COBRAN, HIPAA.NPP, PRINT.DISABLED.DEP.LETTERS')

    # 2.3 EROPTIONS
    doc.add_heading('2.3 EROPTIONS Record', level=2)
    doc.add_paragraph('Description: Employer options/configuration record. Contains flags and fee amounts for COBRA notices, HIPAA privacy practices, and disabled dependent validation.')
    doc.add_paragraph('Key Format: FILENO (same key as EMPLOYERS)')
    doc.add_paragraph('Loaded via: $INCLUDE BP EROPTIONS.INS → MAT EOREC')
    add_styled_table(doc,
        ['Field Name', 'Purpose'],
        [
            ['EOEFFDATE', 'Multi-valued effective dates for option changes'],
            ['EOCOBRA.NOTICE', 'Dollar amount for COBRA notice billing'],
            ['EOHIPAAP', 'HIPAA fee/flag — blank means skip HIPAA processing for SF'],
            ['EODDV', 'Disabled dependent validation flag (Y = enabled)'],
        ])
    doc.add_paragraph('Used by: COBRAN, HIPAA.NPP')

    # 2.4 DEPENDENTS
    doc.add_heading('2.4 DEPENDENTS Record', level=2)
    doc.add_paragraph('Description: Dependent member record. Contains disabled dependent letter tracking, address information for alternate-address mailing, and relationship/coverage dates.')
    doc.add_paragraph('Key Format: EEID.DEPNO (e.g., 12345.01)')
    doc.add_paragraph('Loaded via: $INCLUDE BP DEPENDENTS.INS → MAT DPREC')
    add_styled_table(doc,
        ['Field Name', 'Purpose'],
        [
            ['DPDEP.NEED', 'Comma-delimited codes: RR, RC, IA, LA, DL, NR (disabled dep letter types)'],
            ['DPDDL.SENT', 'Multi-valued dates when disabled dep letters were last sent'],
            ['DPDDL.LETTERS', 'Multi-valued comma-delimited letter codes sent on each date'],
            ['DPDISABILITY.EXPIRATION.DATE', 'Required for Limited Approval (LA) letters'],
            ['DPREL.TYPE.EFFDATE', 'Relationship type effective date (used in HIPAA alt-address logic)'],
            ['DPADD1-DPZIP / DPCOUNTRY', 'Dependent address fields for alternate-address detection'],
        ])
    add_code_block(doc, "CALL *READREC('',MAT DPREC,DEPENDENTS,EEID.DEPNO,'SKIP')\nCALL *WRITEREC('',MAT DPREC,DEPENDENTS,EEID.DEPNO,'TRANS.LOG,':PROGRAM.NAME)")
    doc.add_paragraph('Used by: HIPAA.NPP, PRINT.DISABLED.DEP.LETTERS (COBRAN indirectly)')

    # 2.5 ERADJ
    doc.add_heading('2.5 ERADJ (Employer Adjustments) Record', level=2)
    doc.add_paragraph('Description: Billing adjustment record for employer-level charges. Uses an increment-or-create pattern with suffix iteration.')
    doc.add_paragraph('Key Format: FILENO + BILLPERIOD + SUFFIX (e.g., 12345200401)')
    doc.add_paragraph('Loaded via: $INCLUDE BP ERADJ.INS → MAT EAREC')
    add_styled_table(doc,
        ['Field Name', 'Purpose'],
        [
            ['EADESC', 'Description (e.g., "5 COBRA NOTICES" or "3 HIPAA")'],
            ['EACOMMENTS', 'Multi-valued list of EEID or EEID.DEPNO that were billed'],
            ['EADISTFIELD', 'Distribution field code: COBRACHG or HIPAA'],
            ['EADISTAMOUNT', 'Distribution dollar amount per line'],
            ['EACOVPERIOD', 'Coverage period in YYYYMM format'],
            ['EAAMOUNT', 'Running total amount'],
        ])
    doc.add_paragraph('Used by: COBRAN (COBRACHG), HIPAA.NPP (HIPAA)')

    # 2.6 CORRESP.DATA.REC
    doc.add_heading('2.6 CORRESP.DATA.REC (In-Memory Dynamic Array)', level=2)
    doc.add_paragraph('Not a persistent file record — this is an in-memory dynamic array assembled before calling *CREATE.CORRESP.TRACKING. Structure defined by CORRESP.INPUT.DEF record BASIC.SMARTCOMM.DOC.1.')
    add_styled_table(doc,
        ['Attribute', 'Field Name', 'Purpose'],
        [
            ['<2>', 'GroupNum / FILENO', 'Employer group number for the correspondence'],
            ['<5>', 'AdditionalEEIDDepno', 'Secondary dependent EEID.DEPNO'],
            ['<6>', 'AdditionalDependentPurpose', 'RECIPIENT or SUBJECT'],
        ])
    doc.add_paragraph('Used by: All 3 Programs')

    # 2.7 CORRESP.TRACKING
    doc.add_heading('2.7 CORRESP.TRACKING Record', level=2)
    doc.add_paragraph('Auto-generated by the *CREATE.CORRESP.TRACKING cataloged subroutine. Central correspondence dispatch mechanism.')
    add_styled_table(doc,
        ['Parameter', 'Value'],
        [
            ['CORRESP.INPUT.DEF.ID', "'BASIC.SMARTCOMM.DOC.1'"],
            ['RECIPIENT.TYPE', "'DEPENDENT'"],
            ['METHOD.REQUESTED', "'LOOKUP'"],
            ['REQUEST.SOURCE', 'PROGRAM.NAME (varies per program)'],
            ['DOCUMENT.NAME', 'Template name (varies by letter type)'],
        ])
    doc.add_paragraph('Used by: All 3 Programs')

    # 2.8 SAVEDLISTS
    doc.add_heading('2.8 SAVEDLISTS / Audit Lists', level=2)
    doc.add_paragraph('Named lists stored in the &SAVEDLISTS& system file. Used for audit trails, process resumption, and tracking.')
    add_styled_table(doc,
        ['List Name Pattern', 'Program', 'Purpose'],
        [
            ['COBRAN.DONE{logname}', 'COBRAN', 'Input list of EEIDs to process'],
            ['COBRAN.DONE{logname}.{date}.{time}', 'COBRAN', 'Timestamped completion audit'],
            ['STORE.MERITAIN.COBRA.INTRO.LETTER', 'COBRAN', 'EEIDs that received employee notice'],
            ['STORE.MERITAIN.COBRA.DEPENDENT.NOTICE', 'COBRAN', 'EEID.DEPNOs that received dep notice'],
            ['HIPAA.P{logname}', 'HIPAA.NPP', 'Input list of EEIDs for HIPAA processing'],
            ['HIPAA.PDONE{logname}', 'HIPAA.NPP', 'Completed EEID.DEPNOs (all)'],
            ['HIPAA.PDONE.SC.{logname}', 'HIPAA.NPP', 'SF members sent to SmartComm'],
            ['EEID.LIST.{logname}', 'HIPAA.NPP', 'Filtered EEID.DEPNO list for processing'],
        ])
    doc.add_paragraph('Used by: COBRAN, HIPAA.NPP')

    # 2.9 Supplementary Records
    doc.add_heading('2.9 Supplementary Records', level=2)
    doc.add_paragraph('Records used primarily by HIPAA.NPP for the fully-insured (FI) processing path:')
    add_styled_table(doc,
        ['Record', 'Key/Description', 'Purpose'],
        [
            ['SYSMGMT', 'HIPAA.DAILY.COUNTER', 'Daily run counter to prevent duplicate filenames'],
            ['GROUPINFO', '$INCLUDE BP GROUPINFO.INS', 'Group-level info; GINPP field determines output control'],
            ['CERTS.LAYOUT', '$INCLUDE BP CERTS.LAYOUT.INS', 'Certificate layout definitions keyed by insurance carrier'],
            ['CERTS.MERGE', '$INCLUDE BP CERTS.MERGE.INS', 'Merge field definitions for certificate generation'],
            ['CERTS.TEXT2', '$INCLUDE BP CERTS.TEXT2.INS', 'Text templates for FI HIPAA forms'],
            ['LETTER.COPIES', '$INCLUDE BP LETTER.COPIES.INS', 'Copy/distribution configuration'],
            ['OUTPUT.CONTROL', '$INCLUDE BP OUTPUT.CONTROL.INS', 'Output routing; contains HIPAA.YEARS for FI cutoff'],
            ['FTP.LOG / FTP.REQUESTS', '$INCLUDE BP FTP.*.INS', 'FTP transmission tracking for FI print files'],
        ])

    # 2.10 SmartComm Configuration
    doc.add_heading('2.10 SmartComm Configuration', level=2)
    doc.add_paragraph('All 9 document types use the BASIC.SMARTCOMM.DOC.1 setup in CORRESP.INPUT.DEF:')
    add_styled_table(doc,
        ['#', 'Letter Description', 'Document Name', 'SmartComm ID', 'Returns Image'],
        [
            ['1', 'HIPAA Notice of Privacy Practices', 'Meritain.HIPAA.Notice', '690679345', 'N'],
            ['2', 'COBRA Introduction Letter', 'Meritain.COBRA.Intro.Letter', '690680007', 'N'],
            ['3', 'Dependent COBRA Introduction Letter', 'Meritain.COBRA.Dependent.Notice', '690680105', 'N'],
            ['4', 'Disabled Dependent Denial Notice', 'Disabled.Dependent.-.Denial', '690918681', 'Y'],
            ['5', 'Disabled Dependent Indefinite Approval', 'Disabled.Dependent.-.Indefinite.Approval', '690918738', 'Y'],
            ['6', 'Disabled Dependent Recertification', 'Disabled.Dependent.-.Recertification', '690918752', 'Y'],
            ['7', 'Disabled Dependent Review Request', 'Disabled.Dependent.-.Review.Request', '690918765', 'Y'],
            ['8', 'Disabled Dependent Limited Approval', 'Disabled.Dependent.-.Limited.Approval', '690918728', 'Y'],
            ['9', 'Disabled Dependent No Response Letter', 'Disabled.Dependent.-.No.Response', '690918705', 'Y'],
        ])

    # ===================================================================
    # SECTION 3: The 9 Letter Types
    # ===================================================================
    doc.add_heading('3. The 9 Letter Types — Identified from Code', level=1)
    doc.add_paragraph('Each letter type is discoverable from the source code alone, without requiring the CSV configuration file.')

    doc.add_heading('From COBRAN.md', level=3)
    doc.add_paragraph('1. Meritain.COBRA.Intro.Letter — assigned to TEMPLATE.NAME at line 178')
    doc.add_paragraph('2. Meritain.COBRA.Dependent.Notice — assigned to DEP.TEMPLATE.NAME at line 182')

    doc.add_heading('From HIPAA.NPP.md', level=3)
    doc.add_paragraph('3. Meritain.HIPAA.Notice — assigned to DOCUMENT.NAME at line 1287')

    doc.add_heading('From PRINT.DISABLED.DEP.LETTERS.md (lines 148-173)', level=3)
    doc.add_paragraph('4. Disabled.Dependent.-.Review.Request (code: RR)')
    doc.add_paragraph('5. Disabled.Dependent.-.Recertification (code: RC)')
    doc.add_paragraph('6. Disabled.Dependent.-.Indefinite.Approval (code: IA)')
    doc.add_paragraph('7. Disabled.Dependent.-.Limited.Approval (code: LA)')
    doc.add_paragraph('8. Disabled.Dependent.-.Denial (code: DL)')
    doc.add_paragraph('9. Disabled.Dependent.-.No.Response (code: NR)')

    doc.add_heading('Summary Table', level=3)
    add_styled_table(doc,
        ['#', 'Letter Name', 'Template/Document Name', 'Source Program', 'Trigger Condition', 'SmartComm ID', 'Returns Image'],
        [
            ['1', 'COBRA Intro Letter', 'Meritain.COBRA.Intro.Letter', 'COBRAN', "EENEED='C', ERCOBRAFEE!='N'", '690680007', 'N'],
            ['2', 'COBRA Dependent Notice', 'Meritain.COBRA.Dependent.Notice', 'COBRAN', 'Spouse at different address', '690680105', 'N'],
            ['3', 'HIPAA Notice', 'Meritain.HIPAA.Notice', 'HIPAA.NPP', "EENEED='N', SF group, EOHIPAAP not blank", '690679345', 'N'],
            ['4', 'Disabled Dep Review Request', 'Disabled.Dependent.-.Review.Request', 'PDDL', "DPDEP.NEED contains 'RR'", '690918765', 'Y'],
            ['5', 'Disabled Dep Recertification', 'Disabled.Dependent.-.Recertification', 'PDDL', "DPDEP.NEED contains 'RC'", '690918752', 'Y'],
            ['6', 'Disabled Dep Indefinite Approval', 'Disabled.Dependent.-.Indefinite.Approval', 'PDDL', "DPDEP.NEED contains 'IA'", '690918738', 'Y'],
            ['7', 'Disabled Dep Limited Approval', 'Disabled.Dependent.-.Limited.Approval', 'PDDL', "DPDEP.NEED contains 'LA'", '690918728', 'Y'],
            ['8', 'Disabled Dep Denial', 'Disabled.Dependent.-.Denial', 'PDDL', "DPDEP.NEED contains 'DL'", '690918681', 'Y'],
            ['9', 'Disabled Dep No Response', 'Disabled.Dependent.-.No.Response', 'PDDL', "DPDEP.NEED contains 'NR'", '690918705', 'Y'],
        ])

    # ===================================================================
    # SECTION 4: Dependency Graph
    # ===================================================================
    doc.add_heading('4. Dependency Graph', level=1)
    doc.add_paragraph('This diagram shows how the 10+ record types relate to each other through foreign keys and data flow:')
    print('Fetching Mermaid diagram: Dependency Graph...')
    add_mermaid_diagram(doc, MERMAID_DEPENDENCY, 'Figure 4.1: Record Type Dependency Graph')

    # ===================================================================
    # SECTION 5: Current Coupling Map
    # ===================================================================
    doc.add_heading('5. Current Coupling Map', level=1)
    doc.add_paragraph('This diagram shows the 8+ functional concerns tangled across the 3 programs:')
    print('Fetching Mermaid diagram: Coupling Map...')
    add_mermaid_diagram(doc, MERMAID_COUPLING, 'Figure 5.1: Current Coupling Map')

    # ===================================================================
    # SECTION 6: Coupling Points Table
    # ===================================================================
    doc.add_heading('6. Coupling Points Table', level=1)
    add_styled_table(doc,
        ['Concern', 'Where It\'s Tangled', 'Shared Records'],
        [
            ['Eligibility Filtering', 'COBRAN: MAIN.LOOP\nHIPAA.NPP: FILTER.EEIDS\nPDDL: VALIDATE.DEP.DATA', 'EMPLOYEES, EMPLOYERS, EROPTIONS, DEPENDENTS'],
            ['Employer/Group Loading', 'COBRAN: MAIN.LOOP\nHIPAA.NPP: SET.SKIP.FILENO\nPDDL: GENERATE.DD.LETTERS', 'EMPLOYERS, EROPTIONS'],
            ['Correspondence Submission', 'COBRAN: SUBMIT.CORRESP.REQUEST\nHIPAA.NPP: SUBMIT.CORRESP.REQUEST\nPDDL: SUBMIT.CORRESP.REQUEST', 'CORRESP.DATA.REC, CORRESP.TRACKING'],
            ['Billing (ERADJ)', 'COBRAN: DO.BILLING (COBRACHG)\nHIPAA.NPP: DO.BILLING (HIPAA)', 'ERADJ, EROPTIONS, EMPLOYERS'],
            ['FI HIPAA Generation', 'HIPAA.NPP only: GET.HIPAA.FORMS + GENERATE.HIPAA.FORMS', 'CERTS.LAYOUT, CERTS.MERGE, CERTS.TEXT2, GROUPINFO, OUTPUT.CONTROL'],
            ['FTP Transmission', 'HIPAA.NPP only: FTP.FILE subroutine', 'FTP.LOG, FTP.REQUESTS, REDCARD.PCL'],
            ['Audit List Management', 'COBRAN: END.REPORT.FILE\nHIPAA.NPP: SAVE.LIST commands', '&SAVEDLISTS&'],
            ['Dependent Record Update', 'PDDL only: UPDATE.DEP.NEED + WRITE.DEP.REC', 'DEPENDENTS'],
        ])

    # ===================================================================
    # SECTION 7: Decoupling Strategy
    # ===================================================================
    doc.add_heading('7. Decoupling Strategy', level=1)
    doc.add_paragraph('The following 8 extracted subroutines are proposed to decouple the shared concerns:')

    # 7.1
    doc.add_heading('7.1 *CHECK.LETTER.ELIGIBILITY', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *CHECK.LETTER.ELIGIBILITY(IS.ELIGIBLE, EEID, MAT EEREC, MAT ERREC, MAT EOREC, LETTER.TYPE, TODAY)")
    doc.add_paragraph('Extracts from:')
    doc.add_paragraph('COBRAN — MAIN.LOOP (lines 196-271): status checks, EENEED=\'X\' skip, LOCATE effective date', style='List Bullet')
    doc.add_paragraph('HIPAA.NPP — FILTER.EEIDS (lines 458-556): status exclusions, coverage checks', style='List Bullet')
    doc.add_paragraph('PDDL — VALIDATE.DEP.DATA + CHECK.DEP.NEED (lines 303-467): DEP.NEED validation', style='List Bullet')
    doc.add_paragraph('Current Duplicated Pattern:')
    add_code_block(doc, "IF INDEX(EENEED,'X',1) > 0 THEN CONTINUE\nLOCATE TODAY IN EEEFFDATE<1,1> BY 'DR' SETTING EEIDX ELSE NULL\nIF EEEFFDATE<1,EEIDX> = '' AND EEIDX > 1 THEN EEIDX -= 1")
    doc.add_paragraph('Extraction: Consolidates the common effective-date LOCATE, EENEED=\'X\' exclusion, and status filtering into a single subroutine. Each program passes its LETTER.TYPE to invoke the appropriate eligibility rules.')

    # 7.2
    doc.add_heading('7.2 *LOAD.EMPLOYER.CONTEXT', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *LOAD.EMPLOYER.CONTEXT(LOAD.OK, FILENO, MAT ERREC, MAT EOREC, LAST.FILENO)")
    doc.add_paragraph('Extracts from:')
    doc.add_paragraph('COBRAN — MAIN.LOOP (lines 232-260): conditional read when FILENO changes', style='List Bullet')
    doc.add_paragraph('HIPAA.NPP — SET.SKIP.FILENO (lines 743-775): reads EMPLOYERS + EROPTIONS', style='List Bullet')
    doc.add_paragraph('PDDL — GENERATE.DD.LETTERS (lines 691-705): reads EMPLOYERS per dependent', style='List Bullet')
    doc.add_paragraph('Extraction: Encapsulates the cached employer+eroptions read pattern. Eliminates 3 copies of the same caching logic.')

    # 7.3
    doc.add_heading('7.3 *SUBMIT.LETTER.REQUEST', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *SUBMIT.LETTER.REQUEST(ERROR.MESSAGES, FILENO, RECIPIENT.ID, DOCUMENT.NAME, PROGRAM.NAME, ADDITIONAL.DEPNO, ADDITIONAL.PURPOSE)")
    doc.add_paragraph('Extracts from all 3 programs\' SUBMIT.CORRESP.REQUEST subroutines.')
    doc.add_paragraph('Extraction: Wraps the CORRESP.DATA.REC assembly and *CREATE.CORRESP.TRACKING call. All 3 programs use identical parameter setup.')

    # 7.4
    doc.add_heading('7.4 *ADD.NOTICE.BILLING', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *ADD.NOTICE.BILLING(EEID, FILENO, MAT EEREC, MAT EAREC, MAT EOREC, MAT ERREC, DIST.TYPE, EEIDX, EOIDX)")
    doc.add_paragraph('Extracts from COBRAN DO.BILLING (COBRACHG) and HIPAA.NPP DO.BILLING (HIPAA).')
    doc.add_paragraph('Extraction: Parameterizes the distribution type and the fee source field. The increment-or-create ERADJ loop is identical in both programs.')

    # 7.5
    doc.add_heading('7.5 *GENERATE.FI.HIPAA.NOTICE', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *GENERATE.FI.HIPAA.NOTICE(PCL.OUTPUT, EEID, MAT EEREC, MAT ERREC, FILENO, FORMS.DATA)")
    doc.add_paragraph('Extracts from HIPAA.NPP GET.HIPAA.FORMS + GENERATE.HIPAA.FORMS. Isolates the FI-specific PCL generation into a reusable subroutine.')

    # 7.6
    doc.add_heading('7.6 *SUBMIT.PRINT.FTP', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *SUBMIT.PRINT.FTP(FTP.STATUS, FILENAME, FTP.ID, PROGRAM.NAME)")
    doc.add_paragraph('Extracts from HIPAA.NPP FTP.FILE subroutine. Parameterizes the filename, FTP.ID, and program name.')

    # 7.7
    doc.add_heading('7.7 *SAVE.PROCESS.AUDIT', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *SAVE.PROCESS.AUDIT(PROGRAM.NAME, UV.LOGNAME, STORE.LISTS, LIST.NAMES)")
    doc.add_paragraph('Extracts from COBRAN END.REPORT.FILE and HIPAA.NPP end-of-process code. Standardizes audit trail naming and list saving.')

    # 7.8
    doc.add_heading('7.8 *UPDATE.DEP.LETTER.STATUS', level=2)
    doc.add_paragraph('Proposed Signature:')
    add_code_block(doc, "SUBROUTINE *UPDATE.DEP.LETTER.STATUS(MAT DPREC, EEID.DEPNO, CUR.NEED, ERROR.DEP.NEED, TODAY, PROGRAM.NAME)")
    doc.add_paragraph('Extracts from PDDL WRITE.DEP.REC + UPDATE.DEP.NEED. Encapsulates the dependent letter tracking update logic.')

    # ===================================================================
    # SECTION 8: Decoupled Architecture
    # ===================================================================
    doc.add_heading('8. Decoupled Architecture', level=1)
    doc.add_paragraph('After extraction, the 3 programs become thin orchestrators that delegate to shared subroutines:')
    print('Fetching Mermaid diagram: Decoupled Architecture...')
    add_mermaid_diagram(doc, MERMAID_DECOUPLED, 'Figure 8.1: Decoupled Architecture')

    # ===================================================================
    # SECTION 9: Priority Order for Decoupling
    # ===================================================================
    doc.add_heading('9. Priority Order for Decoupling', level=1)
    add_styled_table(doc,
        ['Priority', 'Extraction', 'Impact', 'Reason'],
        [
            ['1', '*SUBMIT.LETTER.REQUEST', 'High', 'Used by all 3 programs; most duplicated code block; standardizes SmartComm interface'],
            ['2', '*CHECK.LETTER.ELIGIBILITY', 'High', 'Most complex duplicated logic; eligibility bugs affect all letter types'],
            ['3', '*LOAD.EMPLOYER.CONTEXT', 'High', '3-way duplication; caching logic is error-prone; simplifies all 3 programs'],
            ['4', '*ADD.NOTICE.BILLING', 'Medium', '2-way duplication; billing errors have financial impact'],
            ['5', '*SAVE.PROCESS.AUDIT', 'Medium', '2-way duplication; standardizes audit trail naming; low risk'],
            ['6', '*UPDATE.DEP.LETTER.STATUS', 'Medium', 'Single program but complex multi-step update; isolates dependent record mutation'],
            ['7', '*GENERATE.FI.HIPAA.NOTICE', 'Low', 'Single program; isolates legacy PCL path; enables future deprecation'],
            ['8', '*SUBMIT.PRINT.FTP', 'Low', 'Single program; isolates FTP concern; simplifies HIPAA.NPP main flow'],
        ])

    # ===================================================================
    # SECTION 10: Program Flow Diagrams
    # ===================================================================
    doc.add_heading('10. Program Flow Diagrams', level=1)

    doc.add_heading('10.1 COBRAN Flow', level=2)
    doc.add_paragraph('Shows the GOSUB call flow for the COBRAN program:')
    print('Fetching Mermaid diagram: COBRAN Flow...')
    add_mermaid_diagram(doc, MERMAID_COBRAN_FLOW, 'Figure 10.1: COBRAN Program Flow')

    doc.add_heading('10.2 HIPAA.NPP Flow', level=2)
    doc.add_paragraph('Shows the dual-path flow for self-funded (SmartComm) and fully-insured (PCL/FTP) HIPAA notice processing:')
    print('Fetching Mermaid diagram: HIPAA.NPP Flow...')
    add_mermaid_diagram(doc, MERMAID_HIPAA_FLOW, 'Figure 10.2: HIPAA.NPP Program Flow')

    doc.add_heading('10.3 PRINT.DISABLED.DEP.LETTERS Flow', level=2)
    doc.add_paragraph('Shows the disabled dependent letter generation flow:')
    print('Fetching Mermaid diagram: PDDL Flow...')
    add_mermaid_diagram(doc, MERMAID_PDDL_FLOW, 'Figure 10.3: PRINT.DISABLED.DEP.LETTERS Program Flow')

    # ===================================================================
    # Save
    # ===================================================================
    output_path = os.path.join(os.path.dirname(__file__), 'DG_Analysis_Report.docx')
    doc.save(output_path)
    print(f'\nReport saved to: {output_path}')
    return output_path


if __name__ == '__main__':
    build_report()
