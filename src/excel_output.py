from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import pandas as pd
import os
import math

# Path to the header image
HEADER_IMAGE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "header.png")


def add_header_image(ws, end_column='F'):
    """Add college header image to the worksheet"""
    if os.path.exists(HEADER_IMAGE_PATH):
        img = Image(HEADER_IMAGE_PATH)
        # Scale image to fit the header area
        img.width = 550
        img.height = 80
        img.anchor = 'A1'
        ws.add_image(img)
        # Adjust row height for image
        ws.row_dimensions[1].height = 65
        return 2  # Return the number of rows used by image
    return 0


def create_mid_sheet(ws, df_calc, mid_name, subject, department, year, year_sem="I-I Sem"):
    """Create a Mid exam sheet in the workbook"""
    # Styles
    header_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=14)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
    
    # Determine end column based on whether Student Name is present
    has_names = 'Student Name' in df_calc.columns and df_calc['Student Name'].notna().any() and (df_calc['Student Name'] != '').any()
    end_col = 'G' if has_names else 'F'
    mid_col = 'D' if not has_names else 'E'
    
    # Add header image
    image_rows = add_header_image(ws, end_col)
    start_row = 1 + image_rows  # Offset rows for image
    
    # Title rows
    ws.merge_cells(f'A{start_row}:{end_col}{start_row}')
    ws[f'A{start_row}'] = f"DEPARTMENT OF {department}"
    ws[f'A{start_row}'].font = title_font
    ws[f'A{start_row}'].alignment = center_align
    
    # Row 1: A.Y. and Subject (2 columns)
    mid_left_end = 'C' if not has_names else 'D'
    ws.merge_cells(f'A{start_row+1}:{mid_left_end}{start_row+1}')
    ws[f'A{start_row+1}'] = f"A.Y.: {year}"
    ws[f'A{start_row+1}'].alignment = center_align
    
    mid_right_start = 'D' if not has_names else 'E'
    ws.merge_cells(f'{mid_right_start}{start_row+1}:{end_col}{start_row+1}')
    ws[f'{mid_right_start}{start_row+1}'] = f"Subject: {subject}"
    ws[f'{mid_right_start}{start_row+1}'].alignment = center_align
    
    # Row 2: Year & Sem and INTERNAL MARKS SHEET (2 columns)
    ws.merge_cells(f'A{start_row+2}:{mid_left_end}{start_row+2}')
    ws[f'A{start_row+2}'] = f"Year & Sem: {year_sem}"
    ws[f'A{start_row+2}'].alignment = center_align
    
    ws.merge_cells(f'{mid_right_start}{start_row+2}:{end_col}{start_row+2}')
    ws[f'{mid_right_start}{start_row+2}'] = "INTERNAL MARKS SHEET"
    ws[f'{mid_right_start}{start_row+2}'].alignment = center_align
    
    ws.merge_cells(f'A{start_row+3}:{end_col}{start_row+3}')
    ws[f'A{start_row+3}'] = mid_name
    ws[f'A{start_row+3}'].font = title_font
    ws[f'A{start_row+3}'].alignment = center_align
    
    # Check if Student Name column exists
    has_names = 'Student Name' in df_calc.columns and df_calc['Student Name'].notna().any() and (df_calc['Student Name'] != '').any()
    
    # Headers
    if has_names:
        headers = ['S.No.', 'H.T. Number', 'Student Name', 'SA\n(10)', 'Essay\n(20)', 'Assg\n(5)', 'TOTAL Marks\n(35)']
    else:
        headers = ['S.No.', 'H.T. Number', 'SA\n(10)', 'Essay\n(20)', 'Assg\n(5)', 'TOTAL Marks\n(35)']
    header_row = start_row + 4
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = header_fill
    
    # Data rows
    for idx, row in df_calc.iterrows():
        row_num = header_row + 1 + idx
        is_absent = str(row.get('Absent', 'no')).lower() == 'yes'
        
        col = 1
        cell = ws.cell(row=row_num, column=col, value=idx + 1)
        cell.alignment = center_align
        cell.border = thin_border
        
        col += 1
        cell = ws.cell(row=row_num, column=col, value=row['Roll No'])
        cell.border = thin_border
        
        if has_names:
            col += 1
            cell = ws.cell(row=row_num, column=col, value=row.get('Student Name', ''))
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = thin_border
        
        col += 1
        cell = ws.cell(row=row_num, column=col, value='AB' if is_absent else int(row['SA']))
        cell.alignment = center_align
        cell.border = thin_border
        
        col += 1
        cell = ws.cell(row=row_num, column=col, value='AB' if is_absent else int(row['Essay']))
        cell.alignment = center_align
        cell.border = thin_border
        
        col += 1
        assn_is_ab = str(row.get('Assignment', '')) == 'AB'
        cell = ws.cell(row=row_num, column=col, value='AB' if assn_is_ab else int(row['Assignment']))
        cell.alignment = center_align
        cell.border = thin_border
        
        col += 1
        cell = ws.cell(row=row_num, column=col, value='AB' if is_absent else int(row['Total']))
        cell.alignment = center_align
        cell.border = thin_border
    
    # Column widths
    if has_names:
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 13
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 8
        ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 8
        ws.column_dimensions['G'].width = 12
    else:
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 15
    
    ws.row_dimensions[header_row].height = 35
    
    # Signature section at bottom
    last_data_row = header_row + 1 + len(df_calc)
    sig_row = last_data_row + 3  # Leave 2 blank rows
    
    # Signature labels
    sig_font = Font(bold=True, size=10)
    if has_names:
        signatures = [
            ('A', 'Name of the Class Faculty/Sign'),
            ('D', 'Name of the Class Incharge/Sign'),
            ('F', 'HOD')
        ]
    else:
        signatures = [
            ('A', 'Name of the Class Faculty/Sign'),
            ('C', 'Name of the Class Incharge/Sign'),
            ('E', 'HOD')
        ]
    
    for col, label in signatures:
        cell = ws[f'{col}{sig_row}']
        cell.value = label
        cell.font = sig_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Second row of signatures
    sig_row2 = sig_row + 3
    if has_names:
        signatures2 = [
            ('C', 'COE'),
            ('E', 'Principal')
        ]
    else:
        signatures2 = [
            ('B', 'COE'),
            ('D', 'Principal')
        ]
    
    for col, label in signatures2:
        cell = ws[f'{col}{sig_row2}']
        cell.value = label
        cell.font = sig_font
        cell.alignment = Alignment(horizontal='center', vertical='center')


def create_avg_sheet(ws, df_avg, subject, department, year, year_sem="I-I Sem"):
    """Create Average Marks sheet in the workbook"""""
    # Styles
    header_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=12)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
    
    # Determine if Student Name column is present
    has_names = 'Student Name' in df_avg.columns and df_avg['Student Name'].notna().any() and (df_avg['Student Name'] != '').any()
    end_col = 'H' if has_names else 'G'
    mid_left_end = 'D' if has_names else 'C'
    mid_right_start = 'E' if has_names else 'D'
    
    # Add header image
    image_rows = add_header_image(ws, end_col)
    start_row = 1 + image_rows  # Offset rows for image
    
    # Title rows
    ws.merge_cells(f'A{start_row}:{end_col}{start_row}')
    ws[f'A{start_row}'] = f"DEPARTMENT OF {department}"
    ws[f'A{start_row}'].font = title_font
    ws[f'A{start_row}'].alignment = center_align
    
    # Row 1: A.Y. and Subject (2 columns)
    ws.merge_cells(f'A{start_row+1}:{mid_left_end}{start_row+1}')
    ws[f'A{start_row+1}'] = f"A.Y.: {year}"
    ws[f'A{start_row+1}'].alignment = center_align
    
    ws.merge_cells(f'{mid_right_start}{start_row+1}:{end_col}{start_row+1}')
    ws[f'{mid_right_start}{start_row+1}'] = f"Subject: {subject}"
    ws[f'{mid_right_start}{start_row+1}'].alignment = center_align
    
    # Row 2: Year & Sem and INTERNAL MARKS SHEET (2 columns)
    ws.merge_cells(f'A{start_row+2}:{mid_left_end}{start_row+2}')
    ws[f'A{start_row+2}'] = f"Year & Sem: {year_sem}"
    ws[f'A{start_row+2}'].alignment = center_align
    
    ws.merge_cells(f'{mid_right_start}{start_row+2}:{end_col}{start_row+2}')
    ws[f'{mid_right_start}{start_row+2}'] = "INTERNAL MARKS SHEET"
    ws[f'{mid_right_start}{start_row+2}'].alignment = center_align
    
    # Headers
    header_row = start_row + 4
    if has_names:
        headers = ['S.No.', 'H.T. Number', 'Student Name', 'MID -I\n(35Marks)', 'MID -II\n(35Marks)', 'MID\n(AVG)', 
                   'PPT\n(5Marks)', 'Internal Marks\n(40)']
    else:
        headers = ['S.No.', 'H.T. Number', 'MID -I\n(35Marks)', 'MID -II\n(35Marks)', 'MID\n(AVG)', 
                   'PPT\n(5Marks)', 'Internal Marks\n(40)']
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = header_fill
    
    # Data rows
    for idx, row in df_avg.iterrows():
        row_num = header_row + 1 + idx
        
        mid1_absent = str(row.get('Mid1_Absent', 'no')).lower() == 'yes'
        mid2_absent = str(row.get('Mid2_Absent', 'no')).lower() == 'yes'
        pres_is_ab = str(row.get('Presentation', '')) == 'AB'
        avg_is_ab = str(row.get('Mid_Average', '')) == 'AB'
        
        both_absent = mid1_absent and mid2_absent
        
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
            values = [
                idx + 1,
                row['Roll No'],
                row.get('Student Name', ''),
                'AB' if both_absent else int(row['Mid1_Total']),
                'AB' if both_absent else int(row['Mid2_Total']),
                avg_display,
                pres_display,
                internal_marks
            ]
        else:
            values = [
                idx + 1,
                row['Roll No'],
                'AB' if both_absent else int(row['Mid1_Total']),
                'AB' if both_absent else int(row['Mid2_Total']),
                avg_display,
                pres_display,
                internal_marks
            ]
        
        for col_num, value in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            if has_names and col_num == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = center_align
            cell.border = thin_border
    
    # Column widths
    if has_names:
        widths = [5, 13, 30, 9, 9, 8, 7, 10]
    else:
        widths = [6, 14, 12, 12, 10, 10, 12]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    
    ws.row_dimensions[header_row].height = 40
    
    # Signature section at bottom
    last_data_row = header_row + 1 + len(df_avg)
    sig_row = last_data_row + 3  # Leave 2 blank rows
    
    # Signature labels
    sig_font = Font(bold=True, size=10)
    if has_names:
        signatures = [
            ('A', 'Name of the Class Faculty/Sign'),
            ('D', 'Name of the Class Incharge/Sign'),
            ('G', 'HOD')
        ]
    else:
        signatures = [
            ('A', 'Name of the Class Faculty/Sign'),
            ('C', 'Name of the Class Incharge/Sign'),
            ('F', 'HOD')
        ]
    
    for col, label in signatures:
        cell = ws[f'{col}{sig_row}']
        cell.value = label
        cell.font = sig_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Second row of signatures
    sig_row2 = sig_row + 3
    if has_names:
        signatures2 = [
            ('C', 'COE'),
            ('G', 'Principal')
        ]
    else:
        signatures2 = [
            ('C', 'COE'),
            ('F', 'Principal')
        ]
    
    for col, label in signatures2:
        cell = ws[f'{col}{sig_row2}']
        cell.value = label
        cell.font = sig_font
        cell.alignment = Alignment(horizontal='center', vertical='center')


def generate_combined_output(df_mid1_calc, df_mid2_calc, df_avg, output_path, subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem"):
    """Generate single Excel file with Mid-1, Mid-2, and Average Marks sheets"""
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create Mid-1 sheet
    ws_mid1 = wb.create_sheet("Mid-1")
    create_mid_sheet(ws_mid1, df_mid1_calc, "MID-1", subject, department, year, year_sem)
    
    # Create Mid-2 sheet
    ws_mid2 = wb.create_sheet("Mid-2")
    create_mid_sheet(ws_mid2, df_mid2_calc, "MID-2", subject, department, year, year_sem)
    
    # Create Average Marks sheet
    ws_avg = wb.create_sheet("Internal Avg Marks")
    create_avg_sheet(ws_avg, df_avg, subject, department, year, year_sem)
    
    wb.save(output_path)


def generate_mid_output(df_calc, output_path, mid_name="MID", subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem"):
    """Generate Mid exam output Excel file"""
    wb = Workbook()
    ws = wb.active
    ws.title = f"{mid_name} Marks"
    create_mid_sheet(ws, df_calc, mid_name, subject, department, year, year_sem)
    wb.save(output_path)


def generate_avg_output(df_avg, output_path, subject="PPS", department="CIVIL ENGINEERING", year="2025-2026", year_sem="I-I Sem"):
    """Generate Average Marks output Excel file"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Average Marks"
    create_avg_sheet(ws, df_avg, subject, department, year, year_sem)
    wb.save(output_path)