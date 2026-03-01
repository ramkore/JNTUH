from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER
import os
import math

# Path to the header image
HEADER_IMAGE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "header.png")


def create_mid_pdf(df_calc, output_path, mid_name="MID", subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem"):
    """Generate Mid exam PDF output"""
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.25*inch, rightMargin=0.25*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=14,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=4
    )
    sig_style = ParagraphStyle(
        'SigStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER
    )
    
    # Add header image if exists
    if os.path.exists(HEADER_IMAGE_PATH):
        img = Image(HEADER_IMAGE_PATH, width=500, height=70)
        elements.append(img)
        elements.append(Spacer(1, 10))
    
    # Title section
    elements.append(Paragraph(f"DEPARTMENT OF {department}", title_style))
    
    # 2x2 grid for A.Y., Subject, Year & Sem, INTERNAL MARKS SHEET
    info_data = [
        [f"A.Y.: {year}", f"Subject: {subject}"],
        [f"Year & Sem: {year_sem}", "INTERNAL MARKS SHEET"]
    ]
    info_table = Table(info_data, colWidths=[250, 250])
    info_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
    ]))
    elements.append(info_table)
    
    elements.append(Paragraph(f"<b>{mid_name}</b>", title_style))
    elements.append(Spacer(1, 10))
    
    # Check if Student Name column exists
    has_names = 'Student Name' in df_calc.columns and df_calc['Student Name'].notna().any() and (df_calc['Student Name'] != '').any()
    
    # Table data
    if has_names:
        table_data = [['S.No.', 'H.T. Number', 'Student Name', 'SA\n(10)', 'Essay\n(20)', 'Assg\n(5)', 'TOTAL Marks\n(35)']]
    else:
        table_data = [['S.No.', 'H.T. Number', 'SA\n(10)', 'Essay\n(20)', 'Assg\n(5)', 'TOTAL Marks\n(35)']]
    
    for idx, row in df_calc.iterrows():
        is_absent = str(row.get('Absent', 'no')).lower() == 'yes'
        assn_is_ab = str(row.get('Assignment', '')) == 'AB'
        if has_names:
            table_data.append([
                idx + 1,
                row['Roll No'],
                row.get('Student Name', ''),
                'AB' if is_absent else int(row['SA']),
                'AB' if is_absent else int(row['Essay']),
                'AB' if assn_is_ab else int(row['Assignment']),
                'AB' if is_absent else int(row['Total'])
            ])
        else:
            table_data.append([
                idx + 1,
                row['Roll No'],
                'AB' if is_absent else int(row['SA']),
                'AB' if is_absent else int(row['Essay']),
                'AB' if assn_is_ab else int(row['Assignment']),
                'AB' if is_absent else int(row['Total'])
            ])
    
    # Create table
    if has_names:
        col_widths = [28, 85, 155, 45, 50, 38, 65]
    else:
        col_widths = [45, 130, 70, 75, 60, 100]
    row_heights = [30] + [24] * (len(table_data) - 1)  # Header row + data rows
    table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
    
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.85, 0.93, 0.95)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]
    if has_names:
        style_cmds.append(('ALIGN', (2, 1), (2, -1), 'LEFT'))
    table.setStyle(TableStyle(style_cmds))
    
    elements.append(table)
    
    # Signature section
    elements.append(Spacer(1, 40))
    
    sig_data = [
        ['Name of the Class Faculty/Sign', 'Name of the Class Incharge/Sign', 'HOD'],
        ['', '', ''],
        ['', '', ''],
        ['COE', '', 'Principal']
    ]
    
    sig_table = Table(sig_data, colWidths=[200, 200, 120])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ]))
    
    elements.append(sig_table)
    
    doc.build(elements)


def create_avg_pdf(df_avg, output_path, subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem"):
    """Generate Average Marks PDF output"""""
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.25*inch, rightMargin=0.25*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=3
    )
    
    # Add header image if exists
    if os.path.exists(HEADER_IMAGE_PATH):
        img = Image(HEADER_IMAGE_PATH, width=500, height=65)
        elements.append(img)
        elements.append(Spacer(1, 8))
    
    # Title section
    elements.append(Paragraph(f"DEPARTMENT OF {department}", title_style))
    
    # 2x2 grid for A.Y., Subject, Year & Sem, INTERNAL MARKS SHEET
    info_data = [
        [f"A.Y.: {year}", f"Subject: {subject}"],
        [f"Year & Sem: {year_sem}", "INTERNAL MARKS SHEET"]
    ]
    info_table = Table(info_data, colWidths=[240, 240])
    info_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 8))
    
    # Check if Student Name column exists
    has_names = 'Student Name' in df_avg.columns and df_avg['Student Name'].notna().any() and (df_avg['Student Name'] != '').any()
    
    # Table data
    if has_names:
        table_data = [['S.No.', 'H.T. Number', 'Student Name', 'MID-I\n(35)', 'MID-II\n(35)', 'MID\n(AVG)', 
                       'PPT\n(5)', 'Internal\n(40)']]
    else:
        table_data = [['S.No.', 'H.T. Number', 'MID-I\n(35)', 'MID-II\n(35)', 'MID\n(AVG)', 
                       'PPT\n(5)', 'Internal\n(40)']]
    
    for idx, row in df_avg.iterrows():
        mid1_absent = str(row.get('Mid1_Absent', 'no')).lower() == 'yes'
        mid2_absent = str(row.get('Mid2_Absent', 'no')).lower() == 'yes'
        both_absent = mid1_absent and mid2_absent
        pres_is_ab = str(row.get('Presentation', '')) == 'AB'
        avg_is_ab = str(row.get('Mid_Average', '')) == 'AB'
        
        if avg_is_ab:
            avg_display = 'AB'
            pres_display = 'AB'
            internal_marks = 'AB'
        elif pres_is_ab:
            avg_display = round(row['Mid_Average'], 2)
            pres_display = 'AB'
            internal_marks = 'AB'
        else:
            avg_display = round(row['Mid_Average'], 2)
            mid_total = row['Mid_Average'] + row['Presentation']
            internal_marks = math.ceil(mid_total)
            pres_display = int(row['Presentation'])
        
        if has_names:
            table_data.append([
                idx + 1,
                row['Roll No'],
                row.get('Student Name', ''),
                'AB' if both_absent else int(row['Mid1_Total']),
                'AB' if both_absent else int(row['Mid2_Total']),
                avg_display,
                pres_display,
                internal_marks
            ])
        else:
            table_data.append([
                idx + 1,
                row['Roll No'],
                'AB' if both_absent else int(row['Mid1_Total']),
                'AB' if both_absent else int(row['Mid2_Total']),
                avg_display,
                pres_display,
                internal_marks
            ])
    
    # Create table - adjusted column widths for portrait A4
    if has_names:
        col_widths = [25, 80, 145, 45, 45, 42, 35, 48]
    else:
        col_widths = [35, 120, 70, 70, 60, 50, 70]
    row_heights = [30] + [24] * (len(table_data) - 1)  # Header row + data rows
    table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
    
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.85, 0.93, 0.95)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]
    if has_names:
        style_cmds.append(('ALIGN', (2, 1), (2, -1), 'LEFT'))
    table.setStyle(TableStyle(style_cmds))
    
    elements.append(table)
    
    # Signature section
    elements.append(Spacer(1, 30))
    
    sig_data = [
        ['Name of the Class Faculty/Sign', 'Name of the Class Incharge/Sign', 'HOD'],
        ['', '', ''],
        ['', '', ''],
        ['COE', '', 'Principal']
    ]
    
    sig_table = Table(sig_data, colWidths=[180, 180, 115])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    
    elements.append(sig_table)
    
    doc.build(elements)


def generate_combined_pdf(df_mid1_calc, df_mid2_calc, df_avg, output_folder, subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem",
                          mid1_name=None, mid2_name=None, avg_name=None):
    """Generate separate PDF files for Mid-1, Mid-2, and Average Marks"""
    # Create output folder if not exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Use custom names if provided, otherwise fallback to defaults
    if mid1_name is None:
        mid1_name = f"{subject}_Mid1"
    if mid2_name is None:
        mid2_name = f"{subject}_Mid2"
    if avg_name is None:
        avg_name = f"{subject}_Internal_Avg"
    
    # Generate individual PDFs
    mid1_path = os.path.join(output_folder, f"{mid1_name}.pdf")
    create_mid_pdf(df_mid1_calc, mid1_path, "MID-1", subject, department, year, year_sem)
    
    mid2_path = os.path.join(output_folder, f"{mid2_name}.pdf")
    create_mid_pdf(df_mid2_calc, mid2_path, "MID-2", subject, department, year, year_sem)
    
    avg_path = os.path.join(output_folder, f"{avg_name}.pdf")
    create_avg_pdf(df_avg, avg_path, subject, department, year, year_sem)
    
    return mid1_path, mid2_path, avg_path
