import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
from tkcalendar import DateEntry
from datetime import datetime
import sqlite3
import os
import pandas as pd
from outlook_example import import_emails_as_engagements


class EngagementTracker:
    def __init__(self):
        self.root = ThemedTk(theme="arc")  # Modern looking theme
        self.root.title("EngageCRM")
        self.root.geometry("1200x800")
        
        # Register datetime adapters
        sqlite3.register_adapter(datetime, self.adapt_datetime)
        sqlite3.register_converter("datetime", self.convert_datetime)
        
        # Initialize database
        self.init_database()
        
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Create tabs
        self.units_frame = ttk.Frame(self.notebook)
        self.researchers_frame = ttk.Frame(self.notebook)
        self.projects_frame = ttk.Frame(self.notebook)
        self.engagements_frame = ttk.Frame(self.notebook)
        self.reviews_frame = ttk.Frame(self.notebook)
        self.admin_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.units_frame, text='Organizations')
        self.notebook.add(self.researchers_frame, text='Personnel')
        self.notebook.add(self.projects_frame, text='Projects')
        self.notebook.add(self.engagements_frame, text='Engagements')
        self.notebook.add(self.reviews_frame, text='Reviews')
        self.notebook.add(self.admin_frame, text='Admin')
        
        # Initialize all tabs
        self.init_units_tab()
        self.init_researchers_tab()
        self.init_projects_tab()
        self.init_engagements_tab()
        self.init_reviews_tab()
        self.init_admin_tab()

    def adapt_datetime(self, val):
        """Convert datetime to SQLite TEXT format."""
        return val.isoformat()

    def convert_datetime(self, val):
        """Convert SQLite TEXT to datetime object."""
        try:
            return datetime.fromisoformat(val)
        except ValueError:
            return datetime.strptime(val, '%Y-%m-%d')

    def init_database(self):
        """Initialize SQLite database and create tables if they don't exist"""
        db_path = os.path.abspath('engagement_tracker.db')
        
        # Connect to database
        self.conn = sqlite3.connect(
            db_path,
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES
        )
        self.cursor = self.conn.cursor()
        
        # Create Units table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS units (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                type TEXT NOT NULL,
                location TEXT,
                commander TEXT,
                poc TEXT,
                notes TEXT
            )
        ''')
        
        # Create Researchers table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS researchers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                department TEXT,
                expertise TEXT,
                email TEXT,
                phone TEXT,
                notes TEXT
            )
        ''')
        
        # Create Projects table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                status TEXT,
                start_date DATE,
                end_date DATE,
                description TEXT,
                notes TEXT
            )
        ''')
        
        # Create Engagements table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS engagements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_time DATE NOT NULL,
                type TEXT NOT NULL,
                unit_id INTEGER,
                project_id INTEGER,
                summary TEXT,
                status TEXT,
                action_items TEXT,
                FOREIGN KEY (unit_id) REFERENCES units(id),
                FOREIGN KEY (project_id) REFERENCES projects(id)
            )
        ''')
        
        # Create Weekly Reviews table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS weekly_reviews (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_start DATE NOT NULL,
                summary TEXT,
                highlights TEXT,
                challenges TEXT,
                next_steps TEXT
            )
        ''')
        
        # Create Engagement_Participants table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS engagement_participants (
                engagement_id INTEGER,
                researcher_id INTEGER,
                PRIMARY KEY (engagement_id, researcher_id),
                FOREIGN KEY (engagement_id) REFERENCES engagements(id),
                FOREIGN KEY (researcher_id) REFERENCES researchers(id)
            )
        ''')
        
        self.conn.commit()

    def init_units_tab(self):
        """Initialize the Units tab"""
        # Search frame
        search_frame = ttk.Frame(self.units_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.unit_search = ttk.Entry(search_frame)
        self.unit_search.pack(side='left', fill='x', expand=True, padx=5)
        
        # Buttons frame
        btn_frame = ttk.Frame(self.units_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_btn = ttk.Button(
            btn_frame,
            text="Add Unit",
            command=self.add_unit_dialog
        )
        add_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(
            btn_frame,
            text="Edit Unit",
            command=self.edit_unit
        )
        edit_btn.pack(side='left', padx=5)
        
        del_btn = ttk.Button(
            btn_frame,
            text="Delete Unit",
            command=self.delete_unit
        )
        del_btn.pack(side='left', padx=5)
        
        # Treeview for units
        cols = ('ID', 'Name', 'Type', 'Location', 'Commander', 'POC')
        self.units_tree = ttk.Treeview(
            self.units_frame,
            columns=cols,
            show='headings'
        )
        
        # Configure columns
        for col in cols:
            self.units_tree.heading(col, text=col)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            self.units_frame,
            orient='vertical',
            command=self.units_tree.yview
        )
        self.units_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.units_tree.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')

    def init_researchers_tab(self):
        """Initialize the Researchers tab"""
        # Search frame
        search_frame = ttk.Frame(self.researchers_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.researcher_search = ttk.Entry(search_frame)
        self.researcher_search.pack(side='left', fill='x', expand=True, padx=5)
        
        # Buttons
        btn_frame = ttk.Frame(self.researchers_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_btn = ttk.Button(
            btn_frame,
            text="Add Researcher",
            command=self.add_researcher_dialog
        )
        add_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(
            btn_frame,
            text="Edit Researcher",
            command=self.edit_researcher
        )
        edit_btn.pack(side='left', padx=5)

        # Treeview for researchers
        cols = (
            'ID', 'Name', 'Department', 'Expertise',
            'Email', 'Phone'
        )
        self.researchers_tree = ttk.Treeview(
            self.researchers_frame,
            columns=cols,
            show='headings'
        )
        
        # Configure columns
        for col in cols:
            self.researchers_tree.heading(col, text=col)
            if col in ('ID', 'Phone'):
                self.researchers_tree.column(col, width=100)
            elif col in ('Name', 'Department', 'Email'):
                self.researchers_tree.column(col, width=150)
            else:
                self.researchers_tree.column(col, width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            self.researchers_frame,
            orient='vertical',
            command=self.researchers_tree.yview
        )
        self.researchers_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.researchers_tree.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')
        
        # Initial load
        self.refresh_researchers()

    def init_projects_tab(self):
        """Initialize the Projects tab"""
        # Search frame
        search_frame = ttk.Frame(self.projects_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.project_search = ttk.Entry(search_frame)
        self.project_search.pack(side='left', fill='x', expand=True, padx=5)
        
        # Buttons
        btn_frame = ttk.Frame(self.projects_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_btn = ttk.Button(
            btn_frame,
            text="Add Project",
            command=self.add_project_dialog
        )
        add_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(
            btn_frame,
            text="Edit Project",
            command=self.edit_project
        )
        edit_btn.pack(side='left', padx=5)

        # Treeview for projects
        cols = (
            'ID', 'Name', 'Status', 'Start Date',
            'End Date', 'Description', 'Notes'
        )
        self.projects_tree = ttk.Treeview(
            self.projects_frame,
            columns=cols,
            show='headings'
        )
        
        # Configure columns
        for col in cols:
            self.projects_tree.heading(col, text=col)
            if col == 'ID':
                self.projects_tree.column(col, width=50)
            elif col in ('Name', 'Status'):
                self.projects_tree.column(col, width=150)
            elif col in ('Start Date', 'End Date'):
                self.projects_tree.column(col, width=100)
            else:
                self.projects_tree.column(col, width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            self.projects_frame,
            orient='vertical',
            command=self.projects_tree.yview
        )
        self.projects_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.projects_tree.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')
        
        # Initial load
        self.refresh_projects()

    def init_engagements_tab(self):
        """Initialize the Engagements tab"""
        # Search frame
        search_frame = ttk.Frame(self.engagements_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.engagement_search = ttk.Entry(search_frame)
        self.engagement_search.pack(side='left', fill='x', expand=True, padx=5)
        
        # Buttons
        btn_frame = ttk.Frame(self.engagements_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_btn = ttk.Button(
            btn_frame,
            text="Add Engagement",
            command=self.add_engagement_dialog
        )
        add_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(
            btn_frame,
            text="Edit Engagement",
            command=self.edit_engagement
        )
        edit_btn.pack(side='left', padx=5)

        import_btn = ttk.Button(
            btn_frame,
            text="Import from Outlook",
            command=self.import_outlook_emails
        )
        import_btn.pack(side='left', padx=5)

        # Treeview for engagements
        cols = (
            'ID', 'Date', 'Type', 'Unit', 'Project',
            'Summary', 'Status', 'Participants'
        )
        self.engagements_tree = ttk.Treeview(
            self.engagements_frame,
            columns=cols,
            show='headings'
        )
        
        # Configure columns
        for col in cols:
            self.engagements_tree.heading(col, text=col)
            if col in ('ID', 'Date', 'Type', 'Status'):
                self.engagements_tree.column(col, width=100)
            elif col in ('Unit', 'Project'):
                self.engagements_tree.column(col, width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            self.engagements_frame,
            orient='vertical',
            command=self.engagements_tree.yview
        )
        self.engagements_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.engagements_tree.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')
        
        # Initial load
        self.refresh_engagements()

    def init_reviews_tab(self):
        """Initialize the Reviews tab"""
        # Search frame
        search_frame = ttk.Frame(self.reviews_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.review_search = ttk.Entry(search_frame)
        self.review_search.pack(side='left', fill='x', expand=True, padx=5)
        
        # Buttons
        btn_frame = ttk.Frame(self.reviews_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_btn = ttk.Button(
            btn_frame,
            text="Add Review",
            command=self.add_review_dialog
        )
        add_btn.pack(side='left', padx=5)
        
        edit_btn = ttk.Button(
            btn_frame,
            text="Edit Review",
            command=self.edit_review
        )
        edit_btn.pack(side='left', padx=5)

        # Treeview for reviews
        cols = (
            'ID', 'Week Start', 'Summary', 'Highlights',
            'Challenges', 'Next Steps'
        )
        self.reviews_tree = ttk.Treeview(
            self.reviews_frame,
            columns=cols,
            show='headings'
        )
        
        # Configure columns
        for col in cols:
            self.reviews_tree.heading(col, text=col)
            if col == 'ID':
                self.reviews_tree.column(col, width=50)
            elif col == 'Week Start':
                self.reviews_tree.column(col, width=100)
            else:
                self.reviews_tree.column(col, width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            self.reviews_frame,
            orient='vertical',
            command=self.reviews_tree.yview
        )
        self.reviews_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.reviews_tree.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')
        
        # Initial load
        self.refresh_reviews()

    def init_admin_tab(self):
        """Initialize the Admin tab"""
        # Report options frame
        options_frame = ttk.LabelFrame(
            self.admin_frame,
            text="Reports"
        )
        options_frame.pack(fill='x', padx=5, pady=5)
        
        # Date range selection
        date_frame = ttk.Frame(options_frame)
        date_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(date_frame, text="From:").pack(side='left', padx=5)
        self.start_date = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        self.start_date.pack(side='left', padx=5)
        
        ttk.Label(date_frame, text="To:").pack(side='left', padx=5)
        self.end_date = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        self.end_date.pack(side='left', padx=5)
        
        # Report type selection
        report_frame = ttk.Frame(options_frame)
        report_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(report_frame, text="Report Type:").pack(
            side='left',
            padx=5
        )
        
        report_values = [
            'Unit Engagement Summary',
            'Researcher Activity',
            'Project Status',
            'Weekly Review Summary'
        ]
        self.report_type = ttk.Combobox(
            report_frame,
            values=report_values
        )
        self.report_type.pack(side='left', padx=5)
        self.report_type.set(report_values[0])
        
        # Generate and export buttons
        btn_frame = ttk.Frame(self.admin_frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        gen_btn = ttk.Button(
            btn_frame,
            text="Generate Report",
            command=self.generate_report
        )
        gen_btn.pack(side='left', padx=5)
        
        export_btn = ttk.Button(
            btn_frame,
            text="Export to Excel",
            command=self.export_to_excel
        )
        export_btn.pack(side='left', padx=5)
        
        # Report preview area
        preview_frame = ttk.LabelFrame(
            self.admin_frame,
            text="Preview"
        )
        preview_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.report_text = tk.Text(
            preview_frame,
            wrap='word',
            height=20
        )
        self.report_text.pack(fill='both', expand=True, padx=5, pady=5)

    def add_contact_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Contact")
        dialog.geometry("400x500")
        
        # Contact details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Role:").pack(padx=5, pady=5)
        role_combo = ttk.Combobox(dialog, values=['Commander', 'Researcher'])
        role_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Unit/Department:").pack(padx=5, pady=5)
        unit_entry = ttk.Entry(dialog)
        unit_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Email:").pack(padx=5, pady=5)
        email_entry = ttk.Entry(dialog)
        email_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Phone:").pack(padx=5, pady=5)
        phone_entry = ttk.Entry(dialog)
        phone_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.pack(fill='x', padx=5)
        
        def save_contact():
            self.cursor.execute('''
                INSERT INTO contacts (name, role, unit_dept, email, phone, notes)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (name_entry.get(), role_combo.get(), unit_entry.get(),
                  email_entry.get(), phone_entry.get(), notes_text.get("1.0", "end-1c")))
            self.conn.commit()
            self.refresh_contacts()
            dialog.destroy()
        
        ttk.Button(dialog, text="Save", command=save_contact).pack(pady=20)

    def edit_contact(self):
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a contact to edit.")
            return
        
        contact_id = self.contacts_tree.item(selected[0])['values'][0]
        
        # Fetch contact details
        self.cursor.execute("SELECT * FROM contacts WHERE id=?", (contact_id,))
        contact = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Contact")
        dialog.geometry("400x500")
        
        # Contact details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, contact[1])
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Role:").pack(padx=5, pady=5)
        role_combo = ttk.Combobox(dialog, values=['Commander', 'Researcher'])
        role_combo.set(contact[2])
        role_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Unit/Department:").pack(padx=5, pady=5)
        unit_entry = ttk.Entry(dialog)
        unit_entry.insert(0, contact[3] or "")
        unit_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Email:").pack(padx=5, pady=5)
        email_entry = ttk.Entry(dialog)
        email_entry.insert(0, contact[4] or "")
        email_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Phone:").pack(padx=5, pady=5)
        phone_entry = ttk.Entry(dialog)
        phone_entry.insert(0, contact[5] or "")
        phone_entry.pack(fill='x', padx=5)

        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.insert("1.0", contact[6] or "")
        notes_text.pack(fill='x', padx=5)

        def update_contact():
            self.cursor.execute('''
                UPDATE contacts
                SET name=?, role=?, unit_dept=?, email=?, phone=?, notes=?
                WHERE id=?
            ''', (name_entry.get(), role_combo.get(), unit_entry.get(),
                  email_entry.get(), phone_entry.get(), notes_text.get("1.0", "end-1c"),
                  contact_id))
            self.conn.commit()
            self.refresh_contacts()
            dialog.destroy()

        ttk.Button(dialog, text="Update", command=update_contact).pack(pady=20)

    def delete_contact(self):
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a contact to delete.")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this contact?"):
            contact_id = self.contacts_tree.item(selected[0])['values'][0]
            self.cursor.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
            self.conn.commit()
            self.refresh_contacts()

    def refresh_contacts(self):
        """Refresh the contacts treeview"""
        for item in self.contacts_tree.get_children():
            self.contacts_tree.delete(item)
        
        self.cursor.execute("SELECT * FROM contacts")
        for contact in self.cursor.fetchall():
            self.contacts_tree.insert('', 'end', values=contact)

    def generate_report(self):
        """Generate a report based on selected type and date range"""
        report_type = self.report_type.get()
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()
        
        self.report_text.delete(1.0, tk.END)
        report = []
        
        if report_type == 'Unit Engagement Summary':
            self.cursor.execute('''
                SELECT
                    u.name as unit_name,
                    COUNT(e.id) as engagement_count,
                    GROUP_CONCAT(DISTINCT p.name) as projects,
                    GROUP_CONCAT(DISTINCT r.name) as researchers
                FROM units u
                LEFT JOIN engagements e ON u.id = e.unit_id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN projects p ON e.project_id = p.id
                LEFT JOIN engagement_participants ep ON e.id = ep.engagement_id
                LEFT JOIN researchers r ON ep.researcher_id = r.id
                GROUP BY u.id
                ORDER BY engagement_count DESC
            ''', (start_date, end_date))
            
            report.append("Unit Engagement Summary Report")
            report.append(f"Period: {start_date} to {end_date}\n")
            
            for row in self.cursor.fetchall():
                report.append(f"Unit: {row[0]}")
                report.append(f"Total Engagements: {row[1]}")
                if row[2]:
                    report.append("Projects Involved:")
                    for proj in row[2].split(','):
                        report.append(f"  - {proj.strip()}")
                if row[3]:
                    report.append("Researchers Involved:")
                    for res in row[3].split(','):
                        report.append(f"  - {res.strip()}")
                report.append("-" * 50)
        
        elif report_type == 'Researcher Activity':
            self.cursor.execute('''
                SELECT
                    r.name as researcher_name,
                    COUNT(e.id) as engagement_count,
                    GROUP_CONCAT(DISTINCT u.name) as units,
                    GROUP_CONCAT(DISTINCT p.name) as projects
                FROM researchers r
                LEFT JOIN engagement_participants ep ON r.id = ep.researcher_id
                LEFT JOIN engagements e ON ep.engagement_id = e.id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN units u ON e.unit_id = u.id
                LEFT JOIN projects p ON e.project_id = p.id
                GROUP BY r.id
                ORDER BY engagement_count DESC
            ''', (start_date, end_date))
            
            report.append("Researcher Activity Report")
            report.append(f"Period: {start_date} to {end_date}\n")
            
            for row in self.cursor.fetchall():
                report.append(f"Researcher: {row[0]}")
                report.append(f"Total Engagements: {row[1]}")
                if row[2]:
                    report.append("Units Engaged:")
                    for unit in row[2].split(','):
                        report.append(f"  - {unit.strip()}")
                if row[3]:
                    report.append("Projects Involved:")
                    for proj in row[3].split(','):
                        report.append(f"  - {proj.strip()}")
                report.append("-" * 50)
        
        elif report_type == 'Project Status':
            self.cursor.execute('''
                SELECT
                    p.name as project_name,
                    p.status,
                    COUNT(e.id) as engagement_count,
                    GROUP_CONCAT(DISTINCT u.name) as units,
                    GROUP_CONCAT(DISTINCT r.name) as researchers
                FROM projects p
                LEFT JOIN engagements e ON p.id = e.project_id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN units u ON e.unit_id = u.id
                LEFT JOIN engagement_participants ep ON e.id = ep.engagement_id
                LEFT JOIN researchers r ON ep.researcher_id = r.id
                GROUP BY p.id
                ORDER BY engagement_count DESC
            ''', (start_date, end_date))
            
            report.append("Project Status Report")
            report.append(f"Period: {start_date} to {end_date}\n")
            
            for row in self.cursor.fetchall():
                report.append(f"Project: {row[0]}")
                report.append(f"Status: {row[1] or 'Not Started'}")
                report.append(f"Total Engagements: {row[2]}")
                if row[3]:
                    report.append("Units Involved:")
                    for unit in row[3].split(','):
                        report.append(f"  - {unit.strip()}")
                if row[4]:
                    report.append("Researchers Involved:")
                    for res in row[4].split(','):
                        report.append(f"  - {res.strip()}")
                report.append("-" * 50)
        
        elif report_type == 'Weekly Review Summary':
            self.cursor.execute('''
                SELECT
                    week_start,
                    summary,
                    highlights,
                    challenges,
                    next_steps
                FROM weekly_reviews
                WHERE date(week_start) BETWEEN date(?) AND date(?)
                ORDER BY week_start DESC
            ''', (start_date, end_date))
            
            report.append("Weekly Review Summary Report")
            report.append(f"Period: {start_date} to {end_date}\n")
            
            for row in self.cursor.fetchall():
                report.append(f"Week Starting: {row[0]}")
                if row[1]:
                    report.append("Summary:")
                    report.append(row[1])
                if row[2]:
                    report.append("\nHighlights:")
                    report.append(row[2])
                if row[3]:
                    report.append("\nChallenges:")
                    report.append(row[3])
                if row[4]:
                    report.append("\nNext Steps:")
                    report.append(row[4])
                report.append("-" * 50)
        
        # Display report
        self.report_text.insert(tk.END, "\n".join(report))

    def export_to_excel(self):
        """Export the current report to Excel"""
        report_type = self.report_type.get()
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()
        
        if report_type == 'Unit Engagement Summary':
            self.cursor.execute('''
                SELECT
                    u.name as Unit,
                    COUNT(e.id) as "Total Engagements",
                    GROUP_CONCAT(DISTINCT p.name) as "Projects",
                    GROUP_CONCAT(DISTINCT r.name) as "Researchers"
                FROM units u
                LEFT JOIN engagements e ON u.id = e.unit_id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN projects p ON e.project_id = p.id
                LEFT JOIN engagement_participants ep ON e.id = ep.engagement_id
                LEFT JOIN researchers r ON ep.researcher_id = r.id
                GROUP BY u.id
                ORDER BY "Total Engagements" DESC
            ''', (start_date, end_date))
            
            df = pd.DataFrame(
                self.cursor.fetchall(),
                columns=['Unit', 'Total Engagements', 'Projects', 'Researchers']
            )
            
            filename = (
                f"unit_engagement_report_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
        
        elif report_type == 'Researcher Activity':
            self.cursor.execute('''
                SELECT
                    r.name as "Researcher",
                    COUNT(e.id) as "Total Engagements",
                    GROUP_CONCAT(DISTINCT u.name) as "Units",
                    GROUP_CONCAT(DISTINCT p.name) as "Projects"
                FROM researchers r
                LEFT JOIN engagement_participants ep ON r.id = ep.researcher_id
                LEFT JOIN engagements e ON ep.engagement_id = e.id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN units u ON e.unit_id = u.id
                LEFT JOIN projects p ON e.project_id = p.id
                GROUP BY r.id
                ORDER BY "Total Engagements" DESC
            ''', (start_date, end_date))
            
            df = pd.DataFrame(
                self.cursor.fetchall(),
                columns=[
                    'Researcher',
                    'Total Engagements',
                    'Units',
                    'Projects'
                ]
            )
            
            filename = (
                f"researcher_activity_report_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
        
        elif report_type == 'Project Status':
            self.cursor.execute('''
                SELECT
                    p.name as "Project",
                    p.status as "Status",
                    COUNT(e.id) as "Total Engagements",
                    GROUP_CONCAT(DISTINCT u.name) as "Units",
                    GROUP_CONCAT(DISTINCT r.name) as "Researchers"
                FROM projects p
                LEFT JOIN engagements e ON p.id = e.project_id
                    AND date(e.date_time) BETWEEN date(?) AND date(?)
                LEFT JOIN units u ON e.unit_id = u.id
                LEFT JOIN engagement_participants ep ON e.id = ep.engagement_id
                LEFT JOIN researchers r ON ep.researcher_id = r.id
                GROUP BY p.id
                ORDER BY "Total Engagements" DESC
            ''', (start_date, end_date))
            
            df = pd.DataFrame(
                self.cursor.fetchall(),
                columns=[
                    'Project',
                    'Status',
                    'Total Engagements',
                    'Units',
                    'Researchers'
                ]
            )
            
            filename = (
                f"project_status_report_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
        
        elif report_type == 'Weekly Review Summary':
            self.cursor.execute('''
                SELECT
                    week_start as "Week Starting",
                    summary as "Summary",
                    highlights as "Highlights",
                    challenges as "Challenges",
                    next_steps as "Next Steps"
                FROM weekly_reviews
                WHERE date(week_start) BETWEEN date(?) AND date(?)
                ORDER BY week_start DESC
            ''', (start_date, end_date))
            
            df = pd.DataFrame(
                self.cursor.fetchall(),
                columns=[
                    'Week Starting',
                    'Summary',
                    'Highlights',
                    'Challenges',
                    'Next Steps'
                ]
            )
            
            filename = (
                f"weekly_review_report_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
        
        # Export to Excel
        df.to_excel(filename, index=False)
        messagebox.showinfo(
            "Export Complete",
            f"Report exported to {filename}"
        )

    def add_engagement_dialog(self):
        """Dialog for adding a new engagement"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Engagement")
        dialog.geometry("600x800")
        
        # Engagement details
        ttk.Label(dialog, text="Date:").pack(padx=5, pady=5)
        date_entry = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        date_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Type:").pack(padx=5, pady=5)
        type_combo = ttk.Combobox(
            dialog,
            values=[
                'Initial Meeting',
                'Follow-up Meeting',
                'Training Session',
                'Field Test',
                'Demonstration',
                'Other'
            ]
        )
        type_combo.pack(fill='x', padx=5)
        
        # Unit selection
        ttk.Label(dialog, text="Unit:").pack(padx=5, pady=5)
        self.cursor.execute("SELECT id, name FROM units ORDER BY name")
        units = self.cursor.fetchall()
        unit_combo = ttk.Combobox(
            dialog,
            values=[unit[1] for unit in units]
        )
        unit_combo.pack(fill='x', padx=5)
        
        # Project selection
        ttk.Label(dialog, text="Project:").pack(padx=5, pady=5)
        self.cursor.execute("SELECT id, name FROM projects ORDER BY name")
        projects = self.cursor.fetchall()
        project_combo = ttk.Combobox(
            dialog,
            values=[project[1] for project in projects]
        )
        project_combo.pack(fill='x', padx=5)
        
        # Researcher selection
        ttk.Label(dialog, text="Participants:").pack(padx=5, pady=5)
        self.cursor.execute(
            "SELECT id, name FROM researchers ORDER BY name"
        )
        researchers = self.cursor.fetchall()
        researcher_frame = ttk.Frame(dialog)
        researcher_frame.pack(fill='x', padx=5)
        
        researcher_vars = []
        for researcher in researchers:
            var = tk.BooleanVar()
            researcher_vars.append((researcher[0], var))
            ttk.Checkbutton(
                researcher_frame,
                text=researcher[1],
                variable=var
            ).pack(anchor='w')
        
        ttk.Label(dialog, text="Summary:").pack(padx=5, pady=5)
        summary_text = tk.Text(dialog, height=4)
        summary_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Action Items:").pack(padx=5, pady=5)
        action_items_text = tk.Text(dialog, height=4)
        action_items_text.pack(fill='x', padx=5)
        
        def save_engagement():
            # Get unit ID
            unit_name = unit_combo.get()
            unit_id = next(
                (unit[0] for unit in units if unit[1] == unit_name),
                None
            )
            
            # Get project ID
            project_name = project_combo.get()
            project_id = next(
                (proj[0] for proj in projects if proj[1] == project_name),
                None
            )
            
            # Insert engagement
            self.cursor.execute('''
                INSERT INTO engagements (
                    date_time, type, unit_id, project_id,
                    summary, action_items
                )
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                date_entry.get_date().strftime('%Y-%m-%d'),
                type_combo.get(),
                unit_id,
                project_id,
                summary_text.get("1.0", "end-1c"),
                action_items_text.get("1.0", "end-1c")
            ))
            self.conn.commit()
            
            # Get the ID of the inserted engagement
            engagement_id = self.cursor.lastrowid
            
            # Insert participants
            for researcher_id, var in researcher_vars:
                if var.get():
                    self.cursor.execute('''
                        INSERT INTO engagement_participants (
                            engagement_id, researcher_id
                        )
                        VALUES (?, ?)
                    ''', (engagement_id, researcher_id))
            
            self.conn.commit()
            self.refresh_engagements()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_engagement
        ).pack(pady=20)

    def edit_engagement(self):
        """Edit an existing engagement"""
        selected = self.engagements_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select an engagement to edit."
            )
            return
        
        engagement_id = self.engagements_tree.item(selected[0])['values'][0]
        
        # Fetch engagement details
        self.cursor.execute('''
            SELECT
                e.*,
                GROUP_CONCAT(ep.researcher_id) as participant_ids
            FROM engagements e
            LEFT JOIN engagement_participants ep
                ON e.id = ep.engagement_id
            WHERE e.id = ?
            GROUP BY e.id
        ''', (engagement_id,))
        engagement = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Engagement")
        dialog.geometry("600x800")
        
        # Date selection
        date_frame = ttk.Frame(dialog)
        date_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(date_frame, text="Date:").pack(side='left', padx=5)
        date_entry = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        try:
            if isinstance(engagement[1], str):
                # Try parsing with time first
                date_obj = datetime.strptime(engagement[1], '%Y-%m-%d %H:%M:%S')
            else:
                # Already a datetime object
                date_obj = engagement[1]
        except ValueError:
            try:
                if isinstance(engagement[1], str):
                    # If that fails, try parsing just the date
                    date_obj = datetime.strptime(engagement[1], '%Y-%m-%d')
                else:
                    date_obj = engagement[1]
            except ValueError:
                messagebox.showerror("Error", "Invalid date format")
                return
        
        date_entry.set_date(date_obj)
        
        # Type selection
        type_frame = ttk.Frame(dialog)
        type_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(type_frame, text="Type:").pack(side='left', padx=5)
        type_values = [
            'Initial Meeting',
            'Follow-up',
            'Project Review',
            'Demo/Presentation',
            'Training',
            'Other'
        ]
        type_combo = ttk.Combobox(
            type_frame,
            values=type_values
        )
        type_combo.set(engagement[2])
        type_combo.pack(side='left', padx=5)
        
        # Unit selection
        unit_frame = ttk.Frame(dialog)
        unit_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(unit_frame, text="Unit:").pack(side='left', padx=5)
        self.cursor.execute("SELECT id, name FROM units ORDER BY name")
        units = self.cursor.fetchall()
        unit_combo = ttk.Combobox(
            unit_frame,
            values=[unit[1] for unit in units]
        )
        if engagement[3]:  # unit_id
            unit_name = next(
                unit[1] for unit in units if unit[0] == engagement[3]
            )
            unit_combo.set(unit_name)
        unit_combo.pack(side='left', padx=5)
        
        # Project selection
        project_frame = ttk.Frame(dialog)
        project_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(project_frame, text="Project:").pack(side='left', padx=5)
        self.cursor.execute("SELECT id, name FROM projects ORDER BY name")
        projects = self.cursor.fetchall()
        project_combo = ttk.Combobox(
            project_frame,
            values=[proj[1] for proj in projects]
        )
        if engagement[4]:  # project_id
            project_name = next(
                proj[1] for proj in projects if proj[0] == engagement[4]
            )
            project_combo.set(project_name)
        project_combo.pack(side='left', padx=5)
        
        # Participants selection
        participants_frame = ttk.LabelFrame(dialog, text="Participants")
        participants_frame.pack(fill='x', padx=5, pady=5)
        
        self.cursor.execute(
            "SELECT id, name FROM researchers ORDER BY name"
        )
        researchers = self.cursor.fetchall()
        researcher_vars = []
        
        participant_ids = []
        if engagement[-1]:  # participant_ids from GROUP_CONCAT
            participant_ids = [
                int(pid) for pid in engagement[-1].split(',')
            ]
        
        for researcher in researchers:
            var = tk.BooleanVar()
            var.set(researcher[0] in participant_ids)
            researcher_vars.append((researcher[0], var))
            ttk.Checkbutton(
                participants_frame,
                text=researcher[1],
                variable=var
            ).pack(anchor='w', padx=5)
        
        # Summary
        summary_frame = ttk.LabelFrame(dialog, text="Summary")
        summary_frame.pack(fill='x', padx=5, pady=5)
        
        summary_text = tk.Text(summary_frame, height=4)
        summary_text.insert("1.0", engagement[5] or "")  # summary
        summary_text.pack(fill='x', padx=5, pady=5)
        
        # Action Items
        action_frame = ttk.LabelFrame(dialog, text="Action Items")
        action_frame.pack(fill='x', padx=5, pady=5)
        
        action_text = tk.Text(action_frame, height=4)
        action_text.insert("1.0", engagement[7] or "")  # action_items
        action_text.pack(fill='x', padx=5, pady=5)
        
        # Status selection
        status_frame = ttk.Frame(dialog)
        status_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(status_frame, text="Status:").pack(side='left', padx=5)
        status_combo = ttk.Combobox(
            status_frame,
            values=['Open', 'In Progress', 'Completed', 'Cancelled']
        )
        status_combo.set(engagement[6] or 'Open')  # status
        status_combo.pack(side='left', padx=5)
        
        def update_engagement():
            # Get unit and project IDs
            unit_name = unit_combo.get()
            unit_id = next(
                (unit[0] for unit in units if unit[1] == unit_name),
                None
            )
            
            project_name = project_combo.get()
            project_id = next(
                (proj[0] for proj in projects if proj[1] == project_name),
                None
            )
            
            # Update engagement
            self.cursor.execute('''
                UPDATE engagements SET
                    date_time = ?,
                    type = ?,
                    unit_id = ?,
                    project_id = ?,
                    summary = ?,
                    status = ?,
                    action_items = ?
                WHERE id = ?
            ''', (
                date_entry.get_date().strftime('%Y-%m-%d'),
                type_combo.get(),
                unit_id,
                project_id,
                summary_text.get("1.0", "end-1c"),
                status_combo.get(),
                action_text.get("1.0", "end-1c"),
                engagement_id
            ))
            
            # Update participants
            self.cursor.execute(
                "DELETE FROM engagement_participants WHERE engagement_id = ?",
                (engagement_id,)
            )
            
            for researcher_id, var in researcher_vars:
                if var.get():
                    self.cursor.execute('''
                        INSERT INTO engagement_participants (
                            engagement_id,
                            researcher_id
                        )
                        VALUES (?, ?)
                    ''', (engagement_id, researcher_id))
            
            self.conn.commit()
            self.refresh_engagements()
            dialog.destroy()
        
        # Update button
        ttk.Button(
            dialog,
            text="Update Engagement",
            command=update_engagement
        ).pack(pady=20)

    def refresh_engagements(self):
        """Refresh the engagements treeview"""
        # Clear existing items
        for item in self.engagements_tree.get_children():
            self.engagements_tree.delete(item)
        
        # First get all engagements with basic info
        self.cursor.execute('''
            SELECT DISTINCT
                e.id,
                e.date_time,
                e.type,
                u.name AS unit_name,
                p.name AS project_name,
                e.summary,
                e.status
            FROM engagements e
            LEFT JOIN units u ON e.unit_id = u.id
            LEFT JOIN projects p ON e.project_id = p.id
            ORDER BY e.date_time DESC
        ''')
        
        engagements = self.cursor.fetchall()
        
        # Then for each engagement, get its participants
        for engagement in engagements:
            # Convert to list to allow modification
            values = list(engagement)
            
            # Get participants for this engagement
            self.cursor.execute('''
                SELECT GROUP_CONCAT(r.name)
                FROM engagement_participants ep
                JOIN researchers r ON ep.researcher_id = r.id
                WHERE ep.engagement_id = ?
                GROUP BY ep.engagement_id
            ''', (values[0],))  # values[0] is engagement ID
            
            # Get participant names
            participants = self.cursor.fetchone()
            participant_names = participants[0] if participants else ""
            values.append(participant_names)
            
            # Insert into treeview
            self.engagements_tree.insert('', 'end', values=values)

    def add_unit_dialog(self):
        """Dialog for adding a new unit"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Unit")
        dialog.geometry("400x500")
        
        # Unit details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Type:").pack(padx=5, pady=5)
        type_combo = ttk.Combobox(
            dialog,
            values=[
                'Combat Unit',
                'Support Unit',
                'Training Unit',
                'Research Unit',
                'Other'
            ]
        )
        type_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Location:").pack(padx=5, pady=5)
        location_entry = ttk.Entry(dialog)
        location_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Commander:").pack(padx=5, pady=5)
        commander_entry = ttk.Entry(dialog)
        commander_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="POC:").pack(padx=5, pady=5)
        poc_entry = ttk.Entry(dialog)
        poc_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.pack(fill='x', padx=5)
        
        def save_unit():
            self.cursor.execute('''
                INSERT INTO units (
                    name, type, location, commander, poc, notes
                )
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                name_entry.get(),
                type_combo.get(),
                location_entry.get(),
                commander_entry.get(),
                poc_entry.get(),
                notes_text.get("1.0", "end-1c")
            ))
            self.conn.commit()
            self.refresh_units()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_unit
        ).pack(pady=20)

    def edit_unit(self):
        """Edit an existing unit"""
        selected = self.units_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select a unit to edit."
            )
            return
        
        unit_id = self.units_tree.item(selected[0])['values'][0]
        
        # Fetch unit details
        self.cursor.execute(
            "SELECT * FROM units WHERE id=?",
            (unit_id,)
        )
        unit = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Unit")
        dialog.geometry("400x500")
        
        # Unit details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, unit[1])  # name
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Type:").pack(padx=5, pady=5)
        type_combo = ttk.Combobox(
            dialog,
            values=[
                'Combat Unit',
                'Support Unit',
                'Training Unit',
                'Research Unit',
                'Other'
            ]
        )
        type_combo.set(unit[2])  # type
        type_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Location:").pack(padx=5, pady=5)
        location_entry = ttk.Entry(dialog)
        location_entry.insert(0, unit[3] or "")  # location
        location_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Commander:").pack(padx=5, pady=5)
        commander_entry = ttk.Entry(dialog)
        commander_entry.insert(0, unit[4] or "")  # commander
        commander_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="POC:").pack(padx=5, pady=5)
        poc_entry = ttk.Entry(dialog)
        poc_entry.insert(0, unit[5] or "")  # poc
        poc_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.insert("1.0", unit[6] or "")  # notes
        notes_text.pack(fill='x', padx=5)
        
        def update_unit():
            self.cursor.execute('''
                UPDATE units SET
                    name = ?,
                    type = ?,
                    location = ?,
                    commander = ?,
                    poc = ?,
                    notes = ?
                WHERE id = ?
            ''', (
                name_entry.get(),
                type_combo.get(),
                location_entry.get(),
                commander_entry.get(),
                poc_entry.get(),
                notes_text.get("1.0", "end-1c"),
                unit_id
            ))
            self.conn.commit()
            self.refresh_units()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Update",
            command=update_unit
        ).pack(pady=20)

    def delete_unit(self):
        """Delete an existing unit"""
        selected = self.units_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select a unit to delete."
            )
            return
        
        if messagebox.askyesno(
            "Confirm Delete",
            "Are you sure you want to delete this unit?"
        ):
            unit_id = self.units_tree.item(selected[0])['values'][0]
            self.cursor.execute(
                "DELETE FROM units WHERE id=?",
                (unit_id,)
            )
            self.conn.commit()
            self.refresh_units()

    def refresh_units(self):
        """Refresh the units treeview"""
        for item in self.units_tree.get_children():
            self.units_tree.delete(item)
        
        self.cursor.execute(
            "SELECT * FROM units ORDER BY name"
        )
        for unit in self.cursor.fetchall():
            self.units_tree.insert('', 'end', values=unit)

    def add_researcher_dialog(self):
        """Dialog for adding a new researcher"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Researcher")
        dialog.geometry("400x500")
        
        # Researcher details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Department:").pack(padx=5, pady=5)
        department_entry = ttk.Entry(dialog)
        department_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Expertise:").pack(padx=5, pady=5)
        expertise_entry = ttk.Entry(dialog)
        expertise_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Email:").pack(padx=5, pady=5)
        email_entry = ttk.Entry(dialog)
        email_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Phone:").pack(padx=5, pady=5)
        phone_entry = ttk.Entry(dialog)
        phone_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.pack(fill='x', padx=5)
        
        def save_researcher():
            self.cursor.execute('''
                INSERT INTO researchers (
                    name, department, expertise, email, phone, notes
                )
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                name_entry.get(),
                department_entry.get(),
                expertise_entry.get(),
                email_entry.get(),
                phone_entry.get(),
                notes_text.get("1.0", "end-1c")
            ))
            self.conn.commit()
            self.refresh_researchers()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_researcher
        ).pack(pady=20)

    def edit_researcher(self):
        """Edit an existing researcher"""
        selected = self.researchers_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select a researcher to edit."
            )
            return
        
        researcher_id = self.researchers_tree.item(selected[0])['values'][0]
        
        # Fetch researcher details
        self.cursor.execute(
            "SELECT * FROM researchers WHERE id=?",
            (researcher_id,)
        )
        researcher = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Researcher")
        dialog.geometry("400x500")
        
        # Researcher details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, researcher[1])  # name
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Department:").pack(padx=5, pady=5)
        department_entry = ttk.Entry(dialog)
        department_entry.insert(0, researcher[2] or "")  # department
        department_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Expertise:").pack(padx=5, pady=5)
        expertise_entry = ttk.Entry(dialog)
        expertise_entry.insert(0, researcher[3] or "")  # expertise
        expertise_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Email:").pack(padx=5, pady=5)
        email_entry = ttk.Entry(dialog)
        email_entry.insert(0, researcher[4] or "")  # email
        email_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Phone:").pack(padx=5, pady=5)
        phone_entry = ttk.Entry(dialog)
        phone_entry.insert(0, researcher[5] or "")  # phone
        phone_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.insert("1.0", researcher[6] or "")  # notes
        notes_text.pack(fill='x', padx=5)
        
        def update_researcher():
            self.cursor.execute('''
                UPDATE researchers SET
                    name = ?,
                    department = ?,
                    expertise = ?,
                    email = ?,
                    phone = ?,
                    notes = ?
                WHERE id = ?
            ''', (
                name_entry.get(),
                department_entry.get(),
                expertise_entry.get(),
                email_entry.get(),
                phone_entry.get(),
                notes_text.get("1.0", "end-1c"),
                researcher_id
            ))
            self.conn.commit()
            self.refresh_researchers()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Update",
            command=update_researcher
        ).pack(pady=20)

    def refresh_researchers(self):
        """Refresh the researchers treeview"""
        for item in self.researchers_tree.get_children():
            self.researchers_tree.delete(item)
        
        self.cursor.execute(
            "SELECT * FROM researchers ORDER BY name"
        )
        for researcher in self.cursor.fetchall():
            self.researchers_tree.insert('', 'end', values=researcher)

    def add_project_dialog(self):
        """Dialog for adding a new project"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Project")
        dialog.geometry("400x600")
        
        # Project details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Status:").pack(padx=5, pady=5)
        status_combo = ttk.Combobox(
            dialog,
            values=[
                'Planning',
                'In Progress',
                'On Hold',
                'Completed',
                'Cancelled'
            ]
        )
        status_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Start Date:").pack(padx=5, pady=5)
        start_date = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        start_date.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="End Date:").pack(padx=5, pady=5)
        end_date = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        end_date.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Description:").pack(padx=5, pady=5)
        description_text = tk.Text(dialog, height=4)
        description_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.pack(fill='x', padx=5)
        
        def save_project():
            self.cursor.execute('''
                INSERT INTO projects (
                    name, status, start_date, end_date,
                    description, notes
                )
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                name_entry.get(),
                status_combo.get(),
                start_date.get_date().strftime('%Y-%m-%d'),
                end_date.get_date().strftime('%Y-%m-%d'),
                description_text.get("1.0", "end-1c"),
                notes_text.get("1.0", "end-1c")
            ))
            self.conn.commit()
            self.refresh_projects()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_project
        ).pack(pady=20)

    def edit_project(self):
        """Edit an existing project"""
        selected = self.projects_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select a project to edit."
            )
            return
        
        project_id = self.projects_tree.item(selected[0])['values'][0]
        
        # Fetch project details
        self.cursor.execute(
            "SELECT * FROM projects WHERE id=?",
            (project_id,)
        )
        project = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Project")
        dialog.geometry("400x600")
        
        # Project details
        ttk.Label(dialog, text="Name:").pack(padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, project[1])  # name
        name_entry.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Status:").pack(padx=5, pady=5)
        status_combo = ttk.Combobox(
            dialog,
            values=[
                'Planning',
                'In Progress',
                'On Hold',
                'Completed',
                'Cancelled'
            ]
        )
        status_combo.set(project[2] or 'Planning')  # status
        status_combo.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Start Date:").pack(padx=5, pady=5)
        start_date = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        if project[3]:  # start_date
            if isinstance(project[3], str):
                start_date.set_date(datetime.strptime(project[3], '%Y-%m-%d'))
            else:
                start_date.set_date(project[3])  # Already a datetime object
        start_date.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="End Date:").pack(padx=5, pady=5)
        end_date = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        if project[4]:  # end_date
            if isinstance(project[4], str):
                end_date.set_date(datetime.strptime(project[4], '%Y-%m-%d'))
            else:
                end_date.set_date(project[4])  # Already a datetime object
        end_date.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Description:").pack(padx=5, pady=5)
        description_text = tk.Text(dialog, height=4)
        description_text.insert("1.0", project[5] or "")  # description
        description_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Notes:").pack(padx=5, pady=5)
        notes_text = tk.Text(dialog, height=4)
        notes_text.insert("1.0", project[6] or "")  # notes
        notes_text.pack(fill='x', padx=5)
        
        def update_project():
            self.cursor.execute('''
                UPDATE projects SET
                    name = ?,
                    status = ?,
                    start_date = ?,
                    end_date = ?,
                    description = ?,
                    notes = ?
                WHERE id = ?
            ''', (
                name_entry.get(),
                status_combo.get(),
                start_date.get_date().strftime('%Y-%m-%d'),
                end_date.get_date().strftime('%Y-%m-%d'),
                description_text.get("1.0", "end-1c"),
                notes_text.get("1.0", "end-1c"),
                project_id
            ))
            self.conn.commit()
            self.refresh_projects()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Update",
            command=update_project
        ).pack(pady=20)

    def refresh_projects(self):
        """Refresh the projects treeview"""
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        
        self.cursor.execute(
            "SELECT * FROM projects ORDER BY name"
        )
        for project in self.cursor.fetchall():
            self.projects_tree.insert('', 'end', values=project)

    def add_review_dialog(self):
        """Dialog for adding a new weekly review"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Weekly Review")
        dialog.geometry("400x600")
        
        # Review details
        ttk.Label(dialog, text="Week Start Date:").pack(padx=5, pady=5)
        week_start = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        week_start.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Summary:").pack(padx=5, pady=5)
        summary_text = tk.Text(dialog, height=4)
        summary_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Highlights:").pack(padx=5, pady=5)
        highlights_text = tk.Text(dialog, height=4)
        highlights_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Challenges:").pack(padx=5, pady=5)
        challenges_text = tk.Text(dialog, height=4)
        challenges_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Next Steps:").pack(padx=5, pady=5)
        next_steps_text = tk.Text(dialog, height=4)
        next_steps_text.pack(fill='x', padx=5)
        
        def save_review():
            self.cursor.execute('''
                INSERT INTO weekly_reviews (
                    week_start, summary, highlights,
                    challenges, next_steps
                )
                VALUES (?, ?, ?, ?, ?)
            ''', (
                week_start.get_date().strftime('%Y-%m-%d'),
                summary_text.get("1.0", "end-1c"),
                highlights_text.get("1.0", "end-1c"),
                challenges_text.get("1.0", "end-1c"),
                next_steps_text.get("1.0", "end-1c")
            ))
            self.conn.commit()
            self.refresh_reviews()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_review
        ).pack(pady=20)

    def edit_review(self):
        """Edit an existing weekly review"""
        selected = self.reviews_tree.selection()
        if not selected:
            messagebox.showwarning(
                "No Selection",
                "Please select a review to edit."
            )
            return
        
        review_id = self.reviews_tree.item(selected[0])['values'][0]
        
        # Fetch review details
        self.cursor.execute(
            "SELECT * FROM weekly_reviews WHERE id=?",
            (review_id,)
        )
        review = self.cursor.fetchone()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Weekly Review")
        dialog.geometry("400x600")
        
        # Review details
        ttk.Label(dialog, text="Week Start Date:").pack(padx=5, pady=5)
        week_start = DateEntry(
            dialog,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2
        )
        if review[1]:  # week_start
            if isinstance(review[1], str):
                week_start.set_date(datetime.strptime(review[1], '%Y-%m-%d'))
            else:
                week_start.set_date(review[1])  # Already a datetime object
        week_start.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Summary:").pack(padx=5, pady=5)
        summary_text = tk.Text(dialog, height=4)
        summary_text.insert("1.0", review[2] or "")  # summary
        summary_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Highlights:").pack(padx=5, pady=5)
        highlights_text = tk.Text(dialog, height=4)
        highlights_text.insert("1.0", review[3] or "")  # highlights
        highlights_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Challenges:").pack(padx=5, pady=5)
        challenges_text = tk.Text(dialog, height=4)
        challenges_text.insert("1.0", review[4] or "")  # challenges
        challenges_text.pack(fill='x', padx=5)
        
        ttk.Label(dialog, text="Next Steps:").pack(padx=5, pady=5)
        next_steps_text = tk.Text(dialog, height=4)
        next_steps_text.insert("1.0", review[5] or "")  # next_steps
        next_steps_text.pack(fill='x', padx=5)
        
        def update_review():
            self.cursor.execute('''
                UPDATE weekly_reviews SET
                    week_start = ?,
                    summary = ?,
                    highlights = ?,
                    challenges = ?,
                    next_steps = ?
                WHERE id = ?
            ''', (
                week_start.get_date().strftime('%Y-%m-%d'),
                summary_text.get("1.0", "end-1c"),
                highlights_text.get("1.0", "end-1c"),
                challenges_text.get("1.0", "end-1c"),
                next_steps_text.get("1.0", "end-1c"),
                review_id
            ))
            self.conn.commit()
            self.refresh_reviews()
            dialog.destroy()
        
        ttk.Button(
            dialog,
            text="Update",
            command=update_review
        ).pack(pady=20)

    def refresh_reviews(self):
        """Refresh the reviews treeview"""
        for item in self.reviews_tree.get_children():
            self.reviews_tree.delete(item)
        
        self.cursor.execute('''
            SELECT
                id,
                week_start,
                summary,
                highlights,
                challenges,
                next_steps
            FROM weekly_reviews
            ORDER BY week_start DESC
        ''')
        
        for review in self.cursor.fetchall():
            self.reviews_tree.insert('', 'end', values=review)

    def import_outlook_emails(self):
        """Import emails from Outlook folder as engagements"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Import Outlook Emails")
        dialog.geometry("800x600")
        
        # Folder selection
        folder_frame = ttk.Frame(dialog)
        folder_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(
            folder_frame,
            text="Outlook folder name:"
        ).pack(side='left', padx=5)
        
        folder_entry = ttk.Entry(folder_frame)
        folder_entry.insert(0, "Python Emails")
        folder_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        # Preview area
        preview_frame = ttk.LabelFrame(dialog, text="Preview")
        preview_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        preview_text = tk.Text(preview_frame, wrap='word', height=20)
        preview_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        def preview_import():
            folder_name = folder_entry.get()
            try:
                engagements = import_emails_as_engagements(folder_name)
                
                preview = ["Email Import Preview:", ""]
                for i, eng in enumerate(engagements, 1):
                    preview.extend([
                        f"Email {i}:",
                        f"Date: {eng['date_time']}",
                        f"Type: {eng['type']}",
                        f"Unit: {eng['unit'] or 'Unknown'}",
                        f"Participants: {', '.join(eng['participants'])}",
                        f"Attachments: {len(eng['attachments'])}",
                        "Summary: " + eng['summary'][:100] + "...",
                        "-" * 50,
                        ""
                    ])
                
                preview_text.delete(1.0, tk.END)
                preview_text.insert(1.0, "\n".join(preview))
                
            except Exception as e:
                messagebox.showerror(
                    "Preview Error",
                    f"Error previewing emails: {str(e)}"
                )
        
        # Button frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(
            btn_frame,
            text="Preview",
            command=preview_import
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame,
            text="Import",
            command=lambda: self.do_import(folder_entry.get())
        ).pack(side='left', padx=5)

    def do_import(self, folder_name):
        """Import emails from the specified folder into engagements."""
        try:
            engagements = import_emails_as_engagements(folder_name)
            
            imported = 0
            for eng in engagements:
                # Get or create unit if specified
                unit_id = None
                if eng['unit']:
                    self.cursor.execute(
                        "SELECT id FROM units WHERE name=?",
                        (eng['unit'],)
                    )
                    result = self.cursor.fetchone()
                    if result:
                        unit_id = result[0]
                    else:
                        self.cursor.execute(
                            "INSERT INTO units (name, type) VALUES (?, ?)",
                            (eng['unit'], 'Unknown')
                        )
                        unit_id = self.cursor.lastrowid
                
                # Insert engagement
                self.cursor.execute('''
                    INSERT INTO engagements (
                        date_time, type, unit_id, summary,
                        status, action_items
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    eng['date_time'],
                    eng['type'],
                    unit_id,
                    eng['summary'],
                    eng['status'],
                    eng['action_items']
                ))
                
                engagement_id = self.cursor.lastrowid
                
                # Add participants
                for participant in eng['participants']:
                    # Check if researcher exists
                    self.cursor.execute(
                        "SELECT id FROM researchers WHERE name=?",
                        (participant,)
                    )
                    result = self.cursor.fetchone()
                    if result:
                        researcher_id = result[0]
                    else:
                        # Create new researcher
                        self.cursor.execute(
                            "INSERT INTO researchers (name) VALUES (?)",
                            (participant,)
                        )
                        researcher_id = self.cursor.lastrowid
                    
                    # Link researcher to engagement
                    self.cursor.execute('''
                        INSERT INTO engagement_participants (
                            engagement_id, researcher_id
                        )
                        VALUES (?, ?)
                    ''', (engagement_id, researcher_id))
                
                imported += 1
            
            self.conn.commit()
            self.refresh_engagements()
            messagebox.showinfo(
                "Import Complete",
                f"Successfully imported {imported} engagements"
            )
            
        except Exception as e:
            messagebox.showerror(
                "Import Error",
                f"Error importing emails: {str(e)}"
            )

    def run(self):
        """Start the application"""
        # Load initial data
        self.refresh_units()
        self.refresh_researchers()
        self.refresh_engagements()
        self.refresh_reviews()
        self.root.mainloop()


if __name__ == "__main__":
    app = EngagementTracker()
    app.run()
