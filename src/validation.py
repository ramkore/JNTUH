import pandas as pd
from constants import *

def validate_mid_data(df, mid_name):
    """Validate Mid exam data"""
    errors = []
    
    # Check SA (Q1, max 10)
    if 1 in df.columns:
        sa_values = df[1].fillna(0)
        if (sa_values < 0).any() or (sa_values > SA_MAX).any():
            errors.append(f"{mid_name}: SA (Q1) must be 0-{SA_MAX}")
    
    # Check Essay questions (Q2-Q9, max 5 each)
    for q in range(2, 10):
        if q in df.columns:
            q_values = df[q].fillna(0)
            if (q_values < 0).any() or (q_values > ESSAY_QUESTION_MAX).any():
                errors.append(f"{mid_name}: Q{q} must be 0-{ESSAY_QUESTION_MAX}")
    
    # Note: Essay total is capped at 20 during calculation, no need to validate here
    
    return errors

def validate_presentation_data(df):
    """Validate Presentation data"""
    errors = []
    
    if 'Presentation' in df.columns:
        pres_values = df['Presentation'].fillna(0)
        if (pres_values < 0).any() or (pres_values > PRESENTATION_MAX).any():
            errors.append(f"Presentation must be 0-{PRESENTATION_MAX}")
    
    return errors

def validate_all(df_mid1, df_mid2, df_presentation=None):
    """Validate all input data"""
    errors = []
    errors.extend(validate_mid_data(df_mid1, "Mid-1"))
    errors.extend(validate_mid_data(df_mid2, "Mid-2"))
    if df_presentation is not None:
        errors.extend(validate_presentation_data(df_presentation))
    return errors