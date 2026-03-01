import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import os
import math
import json
import subprocess
import sys
import threading
from datetime import datetime
from validation import validate_all
from calculations import calculate_mid_marks, calculate_avg_marks
from excel_output import generate_mid_output, generate_avg_output, generate_combined_output
from pdf_output import create_mid_pdf, create_avg_pdf, generate_combined_pdf
from comparison import compare_mid_marks, get_comparison_summary
from comparison_output import generate_comparison_excel, generate_comparison_pdf


class MarksApp:
    def __init__(self, root):
        self.root = root
        self.root.title("B.Tech Internal Marks Automation System")
        self.root.geometry("1000x800")
        self.root.configure(bg='#f0f0f0')
        self.root.minsize(950, 750)
        
        self.mid1_file = None
        self.mid2_file = None
        self.mid1_assn_file = None
        self.mid2_assn_file = None
        self.df_mid1 = None
        self.df_mid2 = None
        self.df_mid1_assn = None
        self.df_mid2_assn = None
        self.df_mid1_calc = None
        self.df_mid2_calc = None
        self.df_avg = None
        self.df_comparison = None
        self.comparison_summary = None
        
        # Default values
        self.presentation_marks = tk.IntVar(value=5)
        self.program_name = tk.StringVar(value="B.Tech")
        self.subject_name = tk.StringVar(value="PPS")
        self.department_name = tk.StringVar(value="CIVIL ENGINEERING")
        self.department_code = tk.StringVar(value="CE")
        self.academic_year = tk.StringVar(value="2025-2026")
        self.year_sem = tk.StringVar(value="I-I Sem")
        
        # Auto-update department code when department name changes
        self.department_name.trace_add('write', self._auto_update_dept_code)
        
        # Configure styles
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure Treeview table style with grid lines
        style.configure('Table.Treeview',
                        background='white',
                        foreground='black',
                        rowheight=25,
                        fieldbackground='white',
                        borderwidth=1,
                        relief='solid')
        style.configure('Table.Treeview.Heading',
                        background='#DAEEF3',
                        foreground='black',
                        font=('Helvetica', 10, 'bold'),
                        borderwidth=1,
                        relief='solid')
        style.map('Table.Treeview',
                  background=[('selected', '#0078D7')],
                  foreground=[('selected', 'white')])
        
        # Configure red button style for Process All
        style.configure('Red.TButton',
                        background='#DC3545',
                        foreground='white',
                        font=('Helvetica', 10, 'bold'))
        style.map('Red.TButton',
                  background=[('active', '#C82333'), ('disabled', '#6c757d')],
                  foreground=[('disabled', 'white')])
        
        # Title Bar
        title_frame = tk.Frame(root, bg='#2c3e50', height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(title_frame, text="B.Tech Internal Marks Automation", 
                               font=('Helvetica', 18, 'bold'), bg='#2c3e50', fg='white')
        title_label.pack(pady=15)
        
        # Main Container
        main_frame = ttk.Frame(root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left Panel - Input & Settings with Scrollbar
        left_container = ttk.Frame(main_frame, width=300)
        left_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_container.pack_propagate(False)
        
        # Create canvas and scrollbar for left panel
        left_canvas = tk.Canvas(left_container, highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=left_canvas.yview)
        left_panel = ttk.Frame(left_canvas)
        
        left_panel.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left_panel, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Enable mouse wheel scrolling only for left panel when mouse is over it
        def _on_mousewheel_left(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_mousewheel_left(event):
            left_canvas.bind_all("<MouseWheel>", _on_mousewheel_left)
        
        def _unbind_mousewheel_left(event):
            left_canvas.unbind_all("<MouseWheel>")
        
        left_canvas.bind("<Enter>", _bind_mousewheel_left)
        left_canvas.bind("<Leave>", _unbind_mousewheel_left)
        
        # Dependencies Section
        dep_frame = ttk.LabelFrame(left_panel, text="🔧 Dependencies", padding=10)
        dep_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.install_req_btn = ttk.Button(dep_frame, text="Install Requirements", command=self.install_requirements, width=25)
        self.install_req_btn.pack(pady=3)
        
        self.dep_status_label = ttk.Label(dep_frame, text="", foreground='gray')
        self.dep_status_label.pack()
        
        # File Selection
        file_frame = ttk.LabelFrame(left_panel, text="Input Files", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(file_frame, text="Select Mid-1 File", command=self.select_mid1, width=25).pack(pady=3)
        self.mid1_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.mid1_label.pack()
        
        ttk.Button(file_frame, text="Select Mid-2 File", command=self.select_mid2, width=25).pack(pady=3)
        self.mid2_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.mid2_label.pack()
        
        ttk.Button(file_frame, text="Select Mid-1 Assignment File", command=self.select_mid1_assn, width=25).pack(pady=3)
        self.mid1_assn_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.mid1_assn_label.pack()
        
        ttk.Button(file_frame, text="Select Mid-2 Assignment File", command=self.select_mid2_assn, width=25).pack(pady=3)
        self.mid2_assn_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.mid2_assn_label.pack()
        
        # Settings
        settings_frame = ttk.LabelFrame(left_panel, text="Settings", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(settings_frame, text="Program:").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.program_name, width=25).pack(pady=2)
        
        ttk.Label(settings_frame, text="Subject:").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.subject_name, width=25).pack(pady=2)
        
        ttk.Label(settings_frame, text="Department:").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.department_name, width=25).pack(pady=2)

        ttk.Label(settings_frame, text="Department Code (for filenames):").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.department_code, width=25).pack(pady=2)
        
        ttk.Label(settings_frame, text="Academic Year:").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.academic_year, width=25).pack(pady=2)
        
        ttk.Label(settings_frame, text="Year & Sem:").pack(anchor=tk.W)
        ttk.Entry(settings_frame, textvariable=self.year_sem, width=25).pack(pady=2)
        
        ttk.Label(settings_frame, text="Presentation Marks (default 5):").pack(anchor=tk.W)
        ttk.Spinbox(settings_frame, from_=0, to=5, textvariable=self.presentation_marks, width=10).pack(anchor=tk.W, pady=2)
        
        # Actions - Process
        action_frame = ttk.LabelFrame(left_panel, text="Process", padding=10)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.process_btn = ttk.Button(action_frame, text="Process All", command=self.process_all, width=25, style='Red.TButton')
        self.process_btn.pack(pady=5)
        self.process_btn.config(state=tk.DISABLED)
        
        self.compare_btn = ttk.Button(action_frame, text="Compare Mid-I vs Mid-II", command=self.run_comparison, width=25, state=tk.DISABLED)
        self.compare_btn.pack(pady=3)
        
        # Excel Export Section
        excel_frame = ttk.LabelFrame(left_panel, text="📊 Excel Export", padding=10)
        excel_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.save_mid1_btn = ttk.Button(excel_frame, text="Save Mid-1 Excel", command=self.save_mid1, width=25, state=tk.DISABLED)
        self.save_mid1_btn.pack(pady=3)
        
        self.save_mid2_btn = ttk.Button(excel_frame, text="Save Mid-2 Excel", command=self.save_mid2, width=25, state=tk.DISABLED)
        self.save_mid2_btn.pack(pady=3)
        
        self.save_avg_btn = ttk.Button(excel_frame, text="Save Average Marks Excel", command=self.save_avg, width=25, state=tk.DISABLED)
        self.save_avg_btn.pack(pady=3)
        
        ttk.Separator(excel_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        
        self.save_combined_btn = ttk.Button(excel_frame, text="Save Combined Excel", command=self.save_combined, width=25, state=tk.DISABLED)
        self.save_combined_btn.pack(pady=3)
        
        self.save_all_btn = ttk.Button(excel_frame, text="Save All Excel to Folder", command=self.save_all, width=25, state=tk.DISABLED)
        self.save_all_btn.pack(pady=3)
        
        self.save_comp_excel_btn = ttk.Button(excel_frame, text="Save Comparison Excel", command=self.save_comparison_excel, width=25, state=tk.DISABLED)
        self.save_comp_excel_btn.pack(pady=3)
        
        # PDF Export Section
        pdf_frame = ttk.LabelFrame(left_panel, text="📄 PDF Export", padding=10)
        pdf_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.save_pdf_mid1_btn = ttk.Button(pdf_frame, text="Save Mid-1 PDF", command=self.save_mid1_pdf, width=25, state=tk.DISABLED)
        self.save_pdf_mid1_btn.pack(pady=3)
        
        self.save_pdf_mid2_btn = ttk.Button(pdf_frame, text="Save Mid-2 PDF", command=self.save_mid2_pdf, width=25, state=tk.DISABLED)
        self.save_pdf_mid2_btn.pack(pady=3)
        
        self.save_pdf_avg_btn = ttk.Button(pdf_frame, text="Save Average Marks PDF", command=self.save_avg_pdf, width=25, state=tk.DISABLED)
        self.save_pdf_avg_btn.pack(pady=3)
        
        ttk.Separator(pdf_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        
        self.save_all_pdf_btn = ttk.Button(pdf_frame, text="Save All PDFs to Folder", command=self.save_all_pdf, width=25, state=tk.DISABLED)
        self.save_all_pdf_btn.pack(pady=3)
        
        self.save_comp_pdf_btn = ttk.Button(pdf_frame, text="Save Comparison PDF", command=self.save_comparison_pdf, width=25, state=tk.DISABLED)
        self.save_comp_pdf_btn.pack(pady=3)
        
        # Combined Export Section
        combined_frame = ttk.LabelFrame(left_panel, text="📦 Export All (Excel + PDF)", padding=10)
        combined_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.save_all_both_btn = ttk.Button(combined_frame, text="Save All Excel + PDF", command=self.save_all_excel_and_pdf, width=25, state=tk.DISABLED, style='Red.TButton')
        self.save_all_both_btn.pack(pady=3)
        
        # Right Panel - Preview & Status
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Status
        status_frame = ttk.LabelFrame(right_panel, text="Status", padding=10)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = ttk.Label(status_frame, text="Ready. Select Mid-1 and Mid-2 Excel files to begin.")
        self.status_label.pack(anchor=tk.W)
        
        # Preview Notebook
        preview_frame = ttk.LabelFrame(right_panel, text="Data Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.notebook = ttk.Notebook(preview_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Mid-1 Tab
        self.mid1_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.mid1_tab, text="Mid-1")
        self.create_mid_tree(self.mid1_tab, 'mid1')
        
        # Mid-2 Tab
        self.mid2_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.mid2_tab, text="Mid-2")
        self.create_mid_tree(self.mid2_tab, 'mid2')
        
        # Average Tab
        self.avg_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.avg_tab, text="Average Marks")
        self.create_avg_tree(self.avg_tab)
        
        # Comparison Tab
        self.comp_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.comp_tab, text="Comparison")
        self.create_comparison_tree(self.comp_tab)
        
        # History Tab
        self.history_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.history_tab, text="History")
        self.create_history_tab(self.history_tab)
        self.load_history()
        
        # Footer
        footer = tk.Frame(root, bg='#34495e', height=30)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        
        footer_label = tk.Label(footer, text="SA(10) + Essay(20) + Assignment(5) = 35 | Mid Average + Presentation(5) = 40", 
                                font=('Helvetica', 9), bg='#34495e', fg='white')
        footer_label.pack(pady=5)
    
    def create_mid_tree(self, parent, name):
        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        columns = ('SNo', 'HT', 'SA', 'Essay', 'Assg', 'Total')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=15, style='Table.Treeview')
        
        tree.heading('SNo', text='S.No.')
        tree.heading('HT', text='H.T. Number')
        tree.heading('SA', text='SA (10)')
        tree.heading('Essay', text='Essay (20)')
        tree.heading('Assg', text='Assg (5)')
        tree.heading('Total', text='Total (35)')
        
        tree.column('SNo', width=50, anchor='center')
        tree.column('HT', width=120, anchor='center')
        tree.column('SA', width=70, anchor='center')
        tree.column('Essay', width=80, anchor='center')
        tree.column('Assg', width=70, anchor='center')
        tree.column('Total', width=80, anchor='center')
        
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Tag for alternating row colors
        tree.tag_configure('oddrow', background='#F5F5F5')
        tree.tag_configure('evenrow', background='white')
        
        setattr(self, f'tree_{name}', tree)
    
    def create_avg_tree(self, parent):
        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        columns = ('SNo', 'HT', 'Mid1', 'Mid2', 'Avg', 'PPT', 'Internal')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=15, style='Table.Treeview')
        
        headers = [('S.No.', 50), ('H.T. Number', 110), ('MID-I (35)', 80), ('MID-II (35)', 80), 
                   ('MID (AVG)', 80), ('PPT (5)', 60), ('Internal (40)', 90)]
        
        for col, (text, width) in zip(columns, headers):
            tree.heading(col, text=text)
            tree.column(col, width=width, anchor='center')
        
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Tag for alternating row colors
        tree.tag_configure('oddrow', background='#F5F5F5')
        tree.tag_configure('evenrow', background='white')
        
        self.tree_avg = tree
    
    def create_comparison_tree(self, parent):
        """Create the Comparison tab treeview."""
        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        columns = ('SNo', 'HT', 'Name', 'M1SA', 'M1Esy', 'M1Asg', 'M1Tot',
                   'M2SA', 'M2Esy', 'M2Asg', 'M2Tot', 'Diff', 'Status')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=15, style='Table.Treeview')
        
        headers = [
            ('S.No.', 40), ('Roll No', 110), ('Name', 120),
            ('M1 SA', 50), ('M1 Esy', 55), ('M1 Asg', 55), ('M1 Tot', 55),
            ('M2 SA', 50), ('M2 Esy', 55), ('M2 Asg', 55), ('M2 Tot', 55),
            ('Diff', 50), ('Status', 90)
        ]
        for col, (text, width) in zip(columns, headers):
            tree.heading(col, text=text)
            tree.column(col, width=width, anchor='center')
        
        vsb = ttk.Scrollbar(tree_container, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        tree.tag_configure('oddrow', background='#F5F5F5')
        tree.tag_configure('evenrow', background='white')
        tree.tag_configure('improved', background='#C6EFCE')
        tree.tag_configure('declined', background='#FFC7CE')
        tree.tag_configure('missing', background='#FFEB9C')
        
        self.tree_comp = tree
    
    def _get_history_path(self):
        """Return path to history JSON file."""
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        data_dir = os.path.join(base_dir, 'data')
        os.makedirs(data_dir, exist_ok=True)
        return os.path.join(data_dir, 'history.json')

    def create_history_tab(self, parent):
        """Create the History tab with a treeview and clear button."""
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, pady=(0, 5))

        ttk.Button(toolbar, text="Clear History", command=self.clear_history).pack(side=tk.RIGHT)

        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ('No', 'Timestamp', 'Subject', 'Dept', 'YearSem', 'Students', 'Mid1File', 'Mid2File')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=15, style='Table.Treeview')

        headers = [
            ('No', '#', 35), ('Timestamp', 'Date & Time', 140), ('Subject', 'Subject', 70),
            ('Dept', 'Department', 100), ('YearSem', 'Year-Sem', 70), ('Students', 'Students', 60),
            ('Mid1File', 'Mid-1 File', 150), ('Mid2File', 'Mid-2 File', 150)
        ]
        for col, text, width in headers:
            tree.heading(col, text=text)
            tree.column(col, width=width, anchor='center')

        vsb = ttk.Scrollbar(tree_container, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        tree.tag_configure('oddrow', background='#F5F5F5')
        tree.tag_configure('evenrow', background='white')

        self.tree_history = tree

    def load_history(self):
        """Load history from JSON file and populate treeview."""
        path = self._get_history_path()
        self.history_data = []
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    self.history_data = json.load(f)
            except (json.JSONDecodeError, IOError):
                self.history_data = []
        self._refresh_history_tree()

    def _save_history(self):
        """Persist history to JSON file."""
        path = self._get_history_path()
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.history_data, f, indent=2, ensure_ascii=False)

    def _refresh_history_tree(self):
        """Refresh the history treeview from in-memory data."""
        self.tree_history.delete(*self.tree_history.get_children())
        for idx, entry in enumerate(reversed(self.history_data)):  # newest first
            tag = 'oddrow' if idx % 2 == 0 else 'evenrow'
            self.tree_history.insert('', tk.END, values=(
                len(self.history_data) - idx,
                entry.get('timestamp', ''),
                entry.get('subject', ''),
                entry.get('department', ''),
                entry.get('year_sem', ''),
                entry.get('students', ''),
                entry.get('mid1_file', ''),
                entry.get('mid2_file', '')
            ), tags=(tag,))

    def _log_processing(self, student_count):
        """Append a processing record to history."""
        entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'subject': self.subject_name.get(),
            'department': self.department_name.get(),
            'year_sem': self.year_sem.get(),
            'students': student_count,
            'mid1_file': os.path.basename(self.mid1_file) if self.mid1_file else '',
            'mid2_file': os.path.basename(self.mid2_file) if self.mid2_file else ''
        }
        self.history_data.append(entry)
        self._save_history()
        self._refresh_history_tree()

    def clear_history(self):
        """Clear all history records."""
        if messagebox.askyesno("Clear History", "Are you sure you want to clear all processing history?"):
            self.history_data = []
            self._save_history()
            self._refresh_history_tree()

    def select_mid1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.mid1_file = file_path
            filename = os.path.basename(file_path)
            self.mid1_label.config(text=filename, foreground='green')
            self.check_ready()
    
    def select_mid2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.mid2_file = file_path
            filename = os.path.basename(file_path)
            self.mid2_label.config(text=filename, foreground='green')
            self.check_ready()
    
    def select_mid1_assn(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.mid1_assn_file = file_path
            filename = os.path.basename(file_path)
            self.mid1_assn_label.config(text=filename, foreground='green')
            self.check_ready()
    
    def select_mid2_assn(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.mid2_assn_file = file_path
            filename = os.path.basename(file_path)
            self.mid2_assn_label.config(text=filename, foreground='green')
            self.check_ready()
    
    def check_ready(self):
        if self.mid1_file and self.mid2_file and self.mid1_assn_file and self.mid2_assn_file:
            self.process_btn.config(state=tk.NORMAL)
            self.status_label.config(text="All files selected. Click 'Process All' to calculate.")
    
    def _derive_dept_code(self, dept_name):
        """Derive department code from full name by taking first letter of each significant word.
        Handles parenthetical specializations:
          'CIVIL ENGINEERING' -> 'CE'
          'ELECTRONICS AND COMMUNICATION ENGINEERING' -> 'ECE'
          'Electrical and Electronics Engineering' -> 'EEE'
          'COMPUTER SCIENCE AND ENGINEERING(Data Science)' -> 'CSE(DS)'
          'COMPUTER SCIENCE AND ENGINEERING(Cyber Security)' -> 'CSE(CS)'
          'COMPUTER SCIENCE AND ENGINEERING(Artificial Intelligence and Machine Learning)' -> 'CSE(AIML)'
        """
        import re
        skip_words = {'AND', 'OF', 'THE', 'IN', 'FOR'}
        
        # Split into base and specialization (if parenthesized)
        match = re.match(r'^([^(]+?)(?:\((.+)\))?$', dept_name.strip())
        if not match:
            return dept_name.upper()
        
        base_part = match.group(1).strip()
        spec_part = match.group(2)  # May be None
        
        # Derive code from base
        base_words = base_part.upper().split()
        base_code = ''.join(w[0] for w in base_words if w not in skip_words and len(w) > 0)
        
        if spec_part:
            # Derive code from specialization
            spec_words = spec_part.strip().upper().split()
            spec_code = ''.join(w[0] for w in spec_words if w not in skip_words and len(w) > 0)
            return f"{base_code}({spec_code})"
        
        return base_code
    
    def _auto_update_dept_code(self, *args):
        """Callback to auto-update department code when department name changes."""
        dept_name = self.department_name.get().strip()
        if dept_name:
            self.department_code.set(self._derive_dept_code(dept_name))
    
    def _build_filename(self, file_type):
        """Build output filename in format: <YEAR>_<SEM>_B.Tech_<SUBJECT>_<DEPT>_<TYPE>.
        Examples:
          file_type='MID_I'       -> I_I_B.Tech_PPS_ECE_MID_I
          file_type='MID_II'      -> I_I_B.Tech_PPS_ECE_MID_II
          file_type='Internal_Avg'-> I_I_B.Tech_PPS_ECE_Internal_Avg
        """
        # Convert year_sem: "I-I Sem" -> "I_I"
        ys = self.year_sem.get().replace(" Sem", "").replace("-", "_").strip()
        program = self.program_name.get().strip()  # e.g. "B.Tech"
        subject = self.subject_name.get().strip()
        dept_code = self.department_code.get().strip()
        return f"{ys}_{program}_{subject}_{dept_code}_{file_type}"

    def process_all(self):
        try:
            self.status_label.config(text="Processing...")
            self.root.update()
            
            # Read input files
            self.df_mid1 = pd.read_excel(self.mid1_file)
            self.df_mid2 = pd.read_excel(self.mid2_file)
            self.df_mid1_assn = pd.read_excel(self.mid1_assn_file)
            self.df_mid2_assn = pd.read_excel(self.mid2_assn_file)
            
            # Validate
            errors = validate_all(self.df_mid1, self.df_mid2)
            if errors:
                messagebox.showerror("Validation Errors", "\n".join(errors))
                self.status_label.config(text="Validation failed.")
                return
            
            # Calculate
            presentation = self.presentation_marks.get()
            
            self.df_mid1_calc = calculate_mid_marks(self.df_mid1, self.df_mid1_assn)
            self.df_mid2_calc = calculate_mid_marks(self.df_mid2, self.df_mid2_assn)
            
            # Build a master Student Name mapping from ALL input files
            # If any one of the 4 files has Student Name, use it for all outputs
            name_map = {}
            for df_src in [self.df_mid1, self.df_mid2, self.df_mid1_assn, self.df_mid2_assn]:
                if 'Student Name' in df_src.columns:
                    for _, r in df_src.iterrows():
                        roll = r['Roll No']
                        name = r.get('Student Name', '')
                        if pd.notna(name) and str(name).strip() and roll not in name_map:
                            name_map[roll] = str(name).strip()
            
            # Apply master names to calculated dataframes
            if name_map:
                self.df_mid1_calc['Student Name'] = self.df_mid1_calc['Roll No'].map(name_map).fillna('')
                self.df_mid2_calc['Student Name'] = self.df_mid2_calc['Roll No'].map(name_map).fillna('')
            
            self.df_avg = calculate_avg_marks(self.df_mid1_calc, self.df_mid2_calc, presentation_default=presentation)
            
            # Update previews
            self.update_mid_tree(self.tree_mid1, self.df_mid1_calc)
            self.update_mid_tree(self.tree_mid2, self.df_mid2_calc)
            self.update_avg_tree()
            
            # Enable save buttons
            self.save_mid1_btn.config(state=tk.NORMAL)
            self.save_mid2_btn.config(state=tk.NORMAL)
            self.save_avg_btn.config(state=tk.NORMAL)
            self.save_combined_btn.config(state=tk.NORMAL)
            self.save_all_btn.config(state=tk.NORMAL)
            
            # Enable PDF buttons
            self.save_pdf_mid1_btn.config(state=tk.NORMAL)
            self.save_pdf_mid2_btn.config(state=tk.NORMAL)
            self.save_pdf_avg_btn.config(state=tk.NORMAL)
            self.save_all_pdf_btn.config(state=tk.NORMAL)
            
            # Enable combined export button
            self.save_all_both_btn.config(state=tk.NORMAL)
            
            # Enable compare button
            self.compare_btn.config(state=tk.NORMAL)
            
            # Log to history
            self._log_processing(len(self.df_avg))
            
            self.status_label.config(text=f"Processing complete! {len(self.df_avg)} students processed.")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text=f"Error: {str(e)[:50]}")
    
    def update_mid_tree(self, tree, df):
        tree.delete(*tree.get_children())
        for idx, row in df.iterrows():
            is_absent = str(row.get('Absent', 'no')).lower() == 'yes'
            assn_is_ab = str(row.get('Assignment', '')) == 'AB'
            tag = 'oddrow' if idx % 2 == 0 else 'evenrow'
            tree.insert('', tk.END, values=(
                idx + 1,
                row['Roll No'],
                'AB' if is_absent else int(row['SA']),
                'AB' if is_absent else int(row['Essay']),
                'AB' if assn_is_ab else int(row['Assignment']),
                'AB' if is_absent else int(row['Total'])
            ), tags=(tag,))
    
    def update_avg_tree(self):
        self.tree_avg.delete(*self.tree_avg.get_children())
        for idx, row in self.df_avg.iterrows():
            pres_is_ab = str(row.get('Presentation', '')) == 'AB'
            mid1_absent = str(row.get('Mid1_Absent', 'no')).lower() == 'yes'
            mid2_absent = str(row.get('Mid2_Absent', 'no')).lower() == 'yes'
            
            both_absent = mid1_absent and mid2_absent
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
            
            tag = 'oddrow' if idx % 2 == 0 else 'evenrow'
            self.tree_avg.insert('', tk.END, values=(
                idx + 1,
                row['Roll No'],
                'AB' if both_absent else int(row['Mid1_Total']),
                'AB' if both_absent else int(row['Mid2_Total']),
                avg_display,
                pres_display,
                internal_marks
            ), tags=(tag,))
    
    def save_mid1(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                   initialfile=self._build_filename('MID_I') + ".xlsx")
        if file_path:
            generate_mid_output(self.df_mid1_calc, file_path, "MID-1", 
                               self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Mid-1 Excel saved!\n{file_path}")
    
    def save_mid2(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                   initialfile=self._build_filename('MID_II') + ".xlsx")
        if file_path:
            generate_mid_output(self.df_mid2_calc, file_path, "MID-2",
                               self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Mid-2 Excel saved!\n{file_path}")
    
    def save_avg(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                   initialfile=self._build_filename('Internal_Avg') + ".xlsx")
        if file_path:
            generate_avg_output(self.df_avg, file_path,
                               self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Average Marks Excel saved!\n{file_path}")
    
    def save_combined(self):
        """Save Mid-1, Mid-2, and Average Marks in a single Excel file with 3 sheets"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                   initialfile=self._build_filename('Internal_Marks') + ".xlsx")
        if file_path:
            generate_combined_output(self.df_mid1_calc, self.df_mid2_calc, self.df_avg, file_path,
                                    self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Combined Excel with 3 sheets saved!\n{file_path}\n\nSheets: Mid-1, Mid-2, Internal Avg Marks")
            self.status_label.config(text=f"Combined Excel saved: {os.path.basename(file_path)}")
    
    def save_all(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            subject = self.subject_name.get()
            dept = self.department_name.get()
            year = self.academic_year.get()
            year_sem = self.year_sem.get()
            
            # Save all three files using the canonical naming convention
            generate_mid_output(self.df_mid1_calc, os.path.join(folder, self._build_filename('MID_I') + '.xlsx'),
                               "MID-1", subject, dept, year, year_sem)
            generate_mid_output(self.df_mid2_calc, os.path.join(folder, self._build_filename('MID_II') + '.xlsx'),
                               "MID-2", subject, dept, year, year_sem)
            generate_avg_output(self.df_avg, os.path.join(folder, self._build_filename('Internal_Avg') + '.xlsx'),
                               subject, dept, year, year_sem)
            
            messagebox.showinfo("Success", f"All files saved to:\n{folder}")
            self.status_label.config(text=f"All files saved to {folder}")
    
    def save_mid1_pdf(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                   initialfile=self._build_filename('MID_I') + ".pdf")
        if file_path:
            create_mid_pdf(self.df_mid1_calc, file_path, "MID-1", 
                          self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Mid-1 PDF saved!\n{file_path}")
    
    def save_mid2_pdf(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                   initialfile=self._build_filename('MID_II') + ".pdf")
        if file_path:
            create_mid_pdf(self.df_mid2_calc, file_path, "MID-2",
                          self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Mid-2 PDF saved!\n{file_path}")
    
    def save_avg_pdf(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                   initialfile=self._build_filename('Internal_Avg') + ".pdf")
        if file_path:
            create_avg_pdf(self.df_avg, file_path,
                          self.subject_name.get(), self.department_name.get(), self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Average Marks PDF saved!\n{file_path}")
    
    def save_all_pdf(self):
        folder = filedialog.askdirectory(title="Select Output Folder for PDFs")
        if folder:
            subject = self.subject_name.get()
            dept = self.department_name.get()
            year = self.academic_year.get()
            year_sem = self.year_sem.get()
            
            mid1_name = self._build_filename('MID_I')
            mid2_name = self._build_filename('MID_II')
            avg_name = self._build_filename('Internal_Avg')
            
            mid1_path, mid2_path, avg_path = generate_combined_pdf(
                self.df_mid1_calc, self.df_mid2_calc, self.df_avg, folder, subject, dept, year, year_sem,
                mid1_name=mid1_name, mid2_name=mid2_name, avg_name=avg_name)
            
            messagebox.showinfo("Success", f"All PDF files saved to:\n{folder}")
            self.status_label.config(text=f"All PDFs saved to {folder}")
    
    def save_all_excel_and_pdf(self):
        """Save all Excel and PDF files to a single folder.
        Generates 6 files:
          - <YEAR>_<SEM>_B.Tech_<SUBJECT>_<DEPT>_MID_I.xlsx / .pdf
          - <YEAR>_<SEM>_B.Tech_<SUBJECT>_<DEPT>_MID_II.xlsx / .pdf
          - <YEAR>_<SEM>_B.Tech_<SUBJECT>_<DEPT>_Internal_Avg.xlsx / .pdf
        """
        folder = filedialog.askdirectory(title="Select Output Folder for Excel + PDF")
        if folder:
            subject = self.subject_name.get()
            dept = self.department_name.get()
            year = self.academic_year.get()
            year_sem = self.year_sem.get()
            
            mid1_name = self._build_filename('MID_I')
            mid2_name = self._build_filename('MID_II')
            avg_name = self._build_filename('Internal_Avg')
            
            # Generate Excel files
            generate_mid_output(self.df_mid1_calc, os.path.join(folder, mid1_name + '.xlsx'),
                               "MID-1", subject, dept, year, year_sem)
            generate_mid_output(self.df_mid2_calc, os.path.join(folder, mid2_name + '.xlsx'),
                               "MID-2", subject, dept, year, year_sem)
            generate_avg_output(self.df_avg, os.path.join(folder, avg_name + '.xlsx'),
                               subject, dept, year, year_sem)
            
            # Generate PDF files
            generate_combined_pdf(
                self.df_mid1_calc, self.df_mid2_calc, self.df_avg, folder, subject, dept, year, year_sem,
                mid1_name=mid1_name, mid2_name=mid2_name, avg_name=avg_name)
            
            file_list = f"\n".join([
                f"  📊 {mid1_name}.xlsx",
                f"  📊 {mid2_name}.xlsx",
                f"  📊 {avg_name}.xlsx",
                f"  📄 {mid1_name}.pdf",
                f"  📄 {mid2_name}.pdf",
                f"  📄 {avg_name}.pdf",
            ])
            messagebox.showinfo("Success", f"All 6 files saved to:\n{folder}\n\n{file_list}")
            self.status_label.config(text=f"All Excel + PDF files saved to {folder}")
    
    # ---- Comparison methods ----
    
    def run_comparison(self):
        """Run MID-I vs MID-II comparison (excludes Presentation)."""
        if self.df_mid1_calc is None or self.df_mid2_calc is None:
            messagebox.showwarning("Warning", "Please process data first before comparing.")
            return
        try:
            self.status_label.config(text="Running comparison...")
            self.root.update()
            
            self.df_comparison = compare_mid_marks(self.df_mid1_calc, self.df_mid2_calc)
            self.comparison_summary = get_comparison_summary(self.df_comparison)
            
            self.update_comparison_tree()
            
            # Enable comparison save buttons
            self.save_comp_excel_btn.config(state=tk.NORMAL)
            self.save_comp_pdf_btn.config(state=tk.NORMAL)
            
            # Switch to comparison tab
            self.notebook.select(self.comp_tab)
            
            s = self.comparison_summary
            self.status_label.config(
                text=f"Comparison done! {s['total']} students: "
                     f"{s['improved']} improved, {s['declined']} declined, {s['same']} same, "
                     f"{s['missing_mid1']+s['missing_mid2']} missing")
        except Exception as e:
            messagebox.showerror("Comparison Error", str(e))
            self.status_label.config(text=f"Comparison error: {str(e)[:50]}")
    
    def update_comparison_tree(self):
        """Populate the Comparison treeview."""
        self.tree_comp.delete(*self.tree_comp.get_children())
        for idx, row in self.df_comparison.iterrows():
            status = str(row['Status'])
            if status == 'Improved':
                tag = 'improved'
            elif status == 'Declined':
                tag = 'declined'
            elif 'Missing' in status:
                tag = 'missing'
            else:
                tag = 'oddrow' if idx % 2 == 0 else 'evenrow'
            
            diff_display = row['Difference']
            if isinstance(diff_display, float):
                diff_display = round(diff_display, 1)
            
            self.tree_comp.insert('', tk.END, values=(
                idx + 1, row['Roll No'],
                str(row['Student Name'])[:20],
                row['Mid1_SA'], row['Mid1_Essay'], row['Mid1_Assignment'], row['Mid1_Total'],
                row['Mid2_SA'], row['Mid2_Essay'], row['Mid2_Assignment'], row['Mid2_Total'],
                diff_display, status
            ), tags=(tag,))
    
    def save_comparison_excel(self):
        if self.df_comparison is None:
            messagebox.showwarning("Warning", "Run comparison first.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
            initialfile=self._build_filename('Comparison') + ".xlsx")
        if file_path:
            generate_comparison_excel(self.df_comparison, self.comparison_summary, file_path,
                                     self.subject_name.get(), self.department_name.get(),
                                     self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Comparison Excel saved!\n{file_path}")
    
    def save_comparison_pdf(self):
        if self.df_comparison is None:
            messagebox.showwarning("Warning", "Run comparison first.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            initialfile=self._build_filename('Comparison') + ".pdf")
        if file_path:
            generate_comparison_pdf(self.df_comparison, self.comparison_summary, file_path,
                                   self.subject_name.get(), self.department_name.get(),
                                   self.academic_year.get(), self.year_sem.get())
            messagebox.showinfo("Success", f"Comparison PDF saved!\n{file_path}")
    
    # ---- Dependency management ----
    
    def install_requirements(self):
        """Check and install missing Python packages from requirements.txt."""
        self.install_req_btn.config(state=tk.DISABLED)
        self.dep_status_label.config(text="Checking...", foreground='blue')
        self.root.update()
        
        # Run in a thread to avoid blocking the GUI
        thread = threading.Thread(target=self._install_requirements_thread, daemon=True)
        thread.start()
    
    def _install_requirements_thread(self):
        """Background thread for dependency installation."""
        req_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'requirements.txt')
        
        if not os.path.exists(req_path):
            self.root.after(0, lambda: self._dep_done("requirements.txt not found!", 'red'))
            return
        
        with open(req_path, 'r') as f:
            packages = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        
        # Check which packages are missing
        missing = []
        for pkg in packages:
            pkg_name = pkg.split('>=')[0].split('==')[0].split('<')[0].strip()
            try:
                __import__(pkg_name)
            except ImportError:
                # Some packages have different import names
                import_map = {'openpyxl': 'openpyxl', 'reportlab': 'reportlab',
                              'pandas': 'pandas', 'numpy': 'numpy'}
                actual_import = import_map.get(pkg_name, pkg_name)
                try:
                    __import__(actual_import)
                except ImportError:
                    missing.append(pkg)
        
        if not missing:
            self.root.after(0, lambda: self._dep_done(
                "All required packages are already installed.", 'green'))
            return
        
        # Install missing packages
        # Capture 'missing' now using default argument to avoid late-binding issues
        self.root.after(0, lambda n=len(missing): self.dep_status_label.config(
            text=f"Installing {n} package(s)...", foreground='blue'))
        
        try:
            result = subprocess.run(
                [sys.executable, '-m', 'pip', 'install'] + missing,
                capture_output=True, text=True, timeout=120)
            
            if result.returncode == 0:
                msg = f"Installed: {', '.join(missing)}"
                self.root.after(0, lambda m=msg: self._dep_done(m, 'green'))
            else:
                err = result.stderr[:100] if result.stderr else 'Unknown error'
                msg = f"Install failed: {err}"
                self.root.after(0, lambda m=msg: self._dep_done(m, 'red'))
        except Exception as e:
            msg = f"Error: {str(e)[:80]}"
            self.root.after(0, lambda m=msg: self._dep_done(m, 'red'))
    
    def _dep_done(self, message, color):
        """Update dependency status label from main thread."""
        self.dep_status_label.config(text=message, foreground=color)
        self.install_req_btn.config(state=tk.NORMAL)

