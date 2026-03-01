"""
Comparison module for roll number-wise MID-I vs MID-II marks analysis.
Excludes Presentation marks. Detects mismatches and missing roll numbers.
"""
import pandas as pd


def compare_mid_marks(df_mid1_calc, df_mid2_calc):
    """Compare MID-I and MID-II marks roll number-wise.

    Returns a DataFrame with columns:
        Roll No, Student Name,
        Mid1_SA, Mid1_Essay, Mid1_Assignment, Mid1_Total,
        Mid2_SA, Mid2_Essay, Mid2_Assignment, Mid2_Total,
        Difference (Mid2 - Mid1), Status
    """
    # Get all unique roll numbers from both mids
    all_rolls = pd.Index(
        pd.concat([df_mid1_calc['Roll No'], df_mid2_calc['Roll No']]).unique()
    ).sort_values()

    rows = []
    mid1_rolls = set(df_mid1_calc['Roll No'].tolist())
    mid2_rolls = set(df_mid2_calc['Roll No'].tolist())

    for roll in all_rolls:
        row = {'Roll No': roll}

        in_mid1 = roll in mid1_rolls
        in_mid2 = roll in mid2_rolls

        # Student Name (from whichever mid has it)
        name = ''
        if in_mid1:
            m1 = df_mid1_calc[df_mid1_calc['Roll No'] == roll].iloc[0]
            raw_name = m1.get('Student Name', '')
            name = str(raw_name) if pd.notna(raw_name) else ''
            row['Mid1_SA'] = m1['SA']
            row['Mid1_Essay'] = m1['Essay']
            row['Mid1_Assignment'] = m1['Assignment']
            row['Mid1_Total'] = m1['Total']
        else:
            row['Mid1_SA'] = '-'
            row['Mid1_Essay'] = '-'
            row['Mid1_Assignment'] = '-'
            row['Mid1_Total'] = '-'

        if in_mid2:
            m2 = df_mid2_calc[df_mid2_calc['Roll No'] == roll].iloc[0]
            if not name:
                raw_name = m2.get('Student Name', '')
                name = str(raw_name) if pd.notna(raw_name) else ''
            row['Mid2_SA'] = m2['SA']
            row['Mid2_Essay'] = m2['Essay']
            row['Mid2_Assignment'] = m2['Assignment']
            row['Mid2_Total'] = m2['Total']
        else:
            row['Mid2_SA'] = '-'
            row['Mid2_Essay'] = '-'
            row['Mid2_Assignment'] = '-'
            row['Mid2_Total'] = '-'

        row['Student Name'] = name

        # Determine status & difference
        if not in_mid1:
            row['Difference'] = '-'
            row['Status'] = 'Missing in MID-I'
        elif not in_mid2:
            row['Difference'] = '-'
            row['Status'] = 'Missing in MID-II'
        else:
            t1 = float(row['Mid1_Total']) if str(row['Mid1_Total']) not in ('AB', '-') else 0
            t2 = float(row['Mid2_Total']) if str(row['Mid2_Total']) not in ('AB', '-') else 0
            diff = t2 - t1
            row['Difference'] = diff
            if diff > 0:
                row['Status'] = 'Improved'
            elif diff < 0:
                row['Status'] = 'Declined'
            else:
                row['Status'] = 'Same'

        rows.append(row)

    cols = [
        'Roll No', 'Student Name',
        'Mid1_SA', 'Mid1_Essay', 'Mid1_Assignment', 'Mid1_Total',
        'Mid2_SA', 'Mid2_Essay', 'Mid2_Assignment', 'Mid2_Total',
        'Difference', 'Status'
    ]
    df = pd.DataFrame(rows, columns=cols)
    df.reset_index(drop=True, inplace=True)
    return df


def get_missing_rolls(df_mid1_calc, df_mid2_calc):
    """Return roll numbers present in one mid but not the other."""
    mid1_rolls = set(df_mid1_calc['Roll No'].tolist())
    mid2_rolls = set(df_mid2_calc['Roll No'].tolist())
    return {
        'missing_in_mid1': sorted(mid2_rolls - mid1_rolls),
        'missing_in_mid2': sorted(mid1_rolls - mid2_rolls),
    }


def get_comparison_summary(df_comparison):
    """Return counts of each status category."""
    status_counts = df_comparison['Status'].value_counts().to_dict()
    return {
        'total': len(df_comparison),
        'improved': status_counts.get('Improved', 0),
        'declined': status_counts.get('Declined', 0),
        'same': status_counts.get('Same', 0),
        'missing_mid1': status_counts.get('Missing in MID-I', 0),
        'missing_mid2': status_counts.get('Missing in MID-II', 0),
    }
