# Constants for marks calculation

# Mid exam breakdown
SA_MAX = 10          # Short Answer (Q1) max marks
ESSAY_MAX = 20       # Essay: MAX of pairs (2,3), (4,5), (6,7), (8,9) = 4 pairs × 5 marks each
ASSIGNMENT_MAX = 5   # Assignment max marks
MID_TOTAL_30 = 30    # SA + Essay = 10 + 20 = 30 (without assignment)
MID_TOTAL = 35       # SA + Essay + Assignment = 10 + 20 + 5 = 35

PRESENTATION_MAX = 5
FINAL_MAX = 40       # Mid Average (35) + Presentation (5) = 40

# Question structure - using MAX of pairs formula
# Formula: =SUM(Q1, MAX(Q2,Q3), MAX(Q4,Q5), MAX(Q6,Q7), MAX(Q8,Q9))
ESSAY_PAIRS = [(2, 3), (4, 5), (6, 7), (8, 9)]  # Question pairs for MAX calculation
NUM_ESSAY_PAIRS = 4        # Number of question pairs
ESSAY_QUESTION_MAX = 5     # Each essay question/pair max 5 marks