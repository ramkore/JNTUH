import pandas as pd
from constants import *

def calculate_mid_marks(df, df_assignment=None):
    """
    Calculate marks for a single Mid exam using MAX of question pairs
    Formula: =SUM(Q1, MAX(Q2,Q3), MAX(Q4,Q5), MAX(Q6,Q7), MAX(Q8,Q9))
    
    SA (Q1): max 10 marks
    Essay: MAX of pairs (2,3), (4,5), (6,7), (8,9) - 4 pairs × 5 marks = 20 max
    Assignment: from input file (Roll No, Assn columns)
    Total: 35 marks
    
    For 30 marks calculation (without assignment): Q1 + MAX pairs = 10 + 20 = 30
    """
    result = df[['Roll No']].copy()
    
    # Carry Student Name if present
    if 'Student Name' in df.columns:
        result['Student Name'] = df['Student Name'].fillna('')
    else:
        result['Student Name'] = ''
    
    # Handle Absent
    if 'Absent' in df.columns:
        result['Absent'] = df['Absent'].fillna('no')
    else:
        result['Absent'] = 'no'
    
    # SA (Q1) - direct marks
    result['SA'] = df[1].fillna(0) if 1 in df.columns else 0
    
    # Essay: MAX of question pairs (2,3), (4,5), (6,7), (8,9)
    # This takes the best answer from each pair of questions
    essay_pairs = [(2, 3), (4, 5), (6, 7), (8, 9)]
    essay_total = pd.Series(0, index=df.index, dtype=float)
    
    for q1, q2 in essay_pairs:
        q1_marks = df[q1].fillna(0) if q1 in df.columns else 0
        q2_marks = df[q2].fillna(0) if q2 in df.columns else 0
        # Take MAX of each pair
        pair_max = pd.concat([pd.Series(q1_marks), pd.Series(q2_marks)], axis=1).max(axis=1)
        essay_total = essay_total + pair_max
    
    result['Essay'] = essay_total.clip(upper=ESSAY_MAX)
    
    # 30 Marks Total (without assignment) = SA + Essay
    result['Total_30'] = result['SA'] + result['Essay']
    
    # Assignment from input file (merge on Roll No)
    if df_assignment is not None and 'Assn' in df_assignment.columns:
        assn_merged = result[['Roll No']].merge(df_assignment[['Roll No', 'Assn']], on='Roll No', how='left')
        result['Assignment'] = assn_merged['Assn'].fillna(0)
    else:
        result['Assignment'] = 0
    
    # Set marks to 0 for absent students (keep Assignment - it's independent of exam attendance)
    absent_mask = result['Absent'].str.lower() == 'yes'
    result.loc[absent_mask, ['SA', 'Essay', 'Total_30']] = 0
    
    # Total (35 marks with assignment)
    # For absent students, Total = 0 (SA and Essay are already 0, don't add Assignment)
    result['Total'] = result['SA'] + result['Essay']
    result.loc[~absent_mask, 'Total'] = result.loc[~absent_mask, 'Total'] + result.loc[~absent_mask, 'Assignment']
    
    return result

def calculate_avg_marks(df_mid1_calc, df_mid2_calc, df_presentation=None, presentation_default=5):
    """
    Calculate average marks from Mid-1 and Mid-2
    Mid Average = (Mid1_Total + Mid2_Total) / 2
    Final Internal = Mid Average + Presentation
    """
    result = df_mid1_calc[['Roll No']].copy()
    
    # Carry Student Name if present
    if 'Student Name' in df_mid1_calc.columns:
        result['Student Name'] = df_mid1_calc['Student Name'].fillna('')
    else:
        result['Student Name'] = ''
    
    # Track absent status
    result['Mid1_Absent'] = df_mid1_calc['Absent']
    result['Mid2_Absent'] = df_mid2_calc['Absent']
    
    # Determine absent masks per mid
    mid1_absent = df_mid1_calc['Absent'].str.lower() == 'yes'
    mid2_absent = df_mid2_calc['Absent'].str.lower() == 'yes'
    both_absent = mid1_absent & mid2_absent
    
    # Mid-1 details
    result['Mid1_SA'] = df_mid1_calc['SA']
    result['Mid1_Essay'] = df_mid1_calc['Essay']
    result['Mid1_Assignment'] = df_mid1_calc['Assignment'].astype(object)
    result['Mid1_Total'] = df_mid1_calc['Total'].copy()
    
    # Mid-2 details
    result['Mid2_SA'] = df_mid2_calc['SA']
    result['Mid2_Essay'] = df_mid2_calc['Essay']
    result['Mid2_Assignment'] = df_mid2_calc['Assignment'].astype(object)
    result['Mid2_Total'] = df_mid2_calc['Total'].copy()
    
    # If absent in only one mid (not both), include assignment marks in that mid's total for avg calculation
    # e.g., absent in Mid1 with Assn=1: Mid1_Total = 0 + 1 = 1 (SA/Essay are 0, but assignment counts)
    only_mid1_absent = mid1_absent & ~mid2_absent
    only_mid2_absent = mid2_absent & ~mid1_absent
    result.loc[only_mid1_absent, 'Mid1_Total'] = result.loc[only_mid1_absent, 'Mid1_Assignment'].astype(float)
    result.loc[only_mid2_absent, 'Mid2_Total'] = result.loc[only_mid2_absent, 'Mid2_Assignment'].astype(float)
    
    # Only if absent in BOTH mids, set Assignment to AB for both and totals stay 0
    result.loc[both_absent, 'Mid1_Assignment'] = 'AB'
    result.loc[both_absent, 'Mid1_Total'] = 0
    result.loc[both_absent, 'Mid2_Assignment'] = 'AB'
    result.loc[both_absent, 'Mid2_Total'] = 0
    
    # Also update the original mid calc dataframes so individual mid sheets reflect this
    df_mid1_calc['Assignment'] = result['Mid1_Assignment']
    df_mid1_calc.loc[both_absent, 'Total'] = result.loc[both_absent, 'Mid1_Total']
    df_mid2_calc['Assignment'] = result['Mid2_Assignment']
    df_mid2_calc.loc[both_absent, 'Total'] = result.loc[both_absent, 'Mid2_Total']
    
    # Mid Average
    both_absent = mid1_absent & mid2_absent
    result['Mid_Average'] = ((result['Mid1_Total'] + result['Mid2_Total']) / 2).astype(object)
    result.loc[both_absent, 'Mid_Average'] = 'AB'
    
    # Presentation
    if df_presentation is not None and 'Presentation' in df_presentation.columns:
        pres_map = df_presentation.set_index('Roll No')['Presentation'].to_dict()
        result['Presentation'] = result['Roll No'].map(pres_map).fillna(presentation_default)
    else:
        result['Presentation'] = presentation_default
    
    # Only if student is absent in BOTH Mid-1 and Mid-2, set Presentation to 'AB'
    # If attended at least one mid, they get presentation marks
    result['Presentation'] = result['Presentation'].astype(object)
    result.loc[both_absent, 'Presentation'] = 'AB'
    
    # Final Internal = Mid Average + Presentation
    # If both absent: everything is AB
    result['Final_Internal'] = result.apply(
        lambda r: 'AB' if r['Mid_Average'] == 'AB' else (
            r['Mid_Average'] if r['Presentation'] == 'AB' else r['Mid_Average'] + r['Presentation']
        ), axis=1
    )
    
    return result