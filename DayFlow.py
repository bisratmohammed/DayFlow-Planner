import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import calendar
import json
import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import uuid

# ==============================================================================
# DAYFLOW DAILY PLANNER - PYTHON DESKTOP EDITION (TKINTER)
# ==============================================================================
# Features:
# 1. Time-Blocked Schedule (6:00 AM - 11:00 PM) with quick hour assignment
# 2. Task & Priority Manager (Categories, Priorities, Filtering, Search)
# 3. Energy/Focus Batching (Deep Work, Refinement, Admin)
# 4. Daily Review & Reflection Summary with automatic analytics
# 5. Local JSON Storage Persistence & Backup Export/Import
# 6. Light / Dark Mode Theme Toggle
# ==============================================================================

USER_DATA_ROOT = os.environ.get("LOCALAPPDATA") or os.path.join(
    os.path.expanduser("~"), ".local", "share"
)
DATA_DIR = os.path.join(USER_DATA_ROOT, "DayFlow")
DATA_FILE = os.path.join(DATA_DIR, "dayflow_data.json")


class DayFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DayFlow — Minimalist Daily Planner & Energy Batcher")
        self.root.geometry("1120x740")
        self.root.minsize(900, 620)

        # State variables
        self.dark_mode = False
        self.layout_mode = "Desktop"
        self.mobile_timeline_panel = "schedule"
        self.current_date = datetime.now().strftime("%Y-%m-%d")
        self.tasks = []
        self.reflections = {}
        self.tracker_ranges = {}

        # Load persisted data or initialize an empty user profile.
        self.load_data()

        # Color palettes for Light & Dark themes
        self.themes = {
            "light": {
                "bg": "#F8FAFC",
                "card": "#FFFFFF",
                "card_alt": "#F1F5F9",
                "text": "#0F172A",
                "subtext": "#64748B",
                "border": "#E2E8F0",
                "accent": "#4F46E5",
                "accent_hover": "#4338CA",
                "accent_fg": "#FFFFFF",
                "input_bg": "#FFFFFF",
            },
            "dark": {
                "bg": "#090D16",
                "card": "#111827",
                "card_alt": "#1E293B",
                "text": "#F8FAFC",
                "subtext": "#94A3B8",
                "border": "#1E293B",
                "accent": "#6366F1",
                "accent_hover": "#818CF8",
                "accent_fg": "#FFFFFF",
                "input_bg": "#1E293B",
            },
        }

        # Category badge colors
        self.category_colors = {
            "Work": {"bg": "#EEF2FF", "fg": "#4338CA", "border": "#6366F1"},
            "Personal": {"bg": "#FDF4FF", "fg": "#A21CAF", "border": "#D946EF"},
            "Health": {"bg": "#ECFDF5", "fg": "#047857", "border": "#10B981"},
            "Errands": {"bg": "#FFFBEB", "fg": "#B45309", "border": "#F59E0B"},
        }

        # Energy badge colors
        self.energy_colors = {
            "Deep Work": "#7C3AED",
            "Refinement": "#0891B2",
            "Admin": "#EA580C",
        }

        # Priority colors
        self.priority_colors = {
            "High": "#E11D48",
            "Med": "#D97706",
            "Low": "#0284C7",
        }

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.build_ui()
        self.apply_theme()
        self.apply_layout()
        self.refresh_all_views()

    # --------------------------------------------------------------------------
    # DATA PERSISTENCE
    # --------------------------------------------------------------------------
    def load_data(self):
        """Load tasks and reflections from local JSON storage."""
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.tasks = data.get("tasks", [])
                    self.reflections = data.get("reflections", {})
                    self.dark_mode = data.get("dark_mode", False)
                    layout_mode = data.get("layout_mode", "Desktop")
                    if layout_mode in ("Desktop", "Mobile"):
                        self.layout_mode = layout_mode
                    return
            except Exception as e:
                print(f"Error loading data: {e}")

        self.tasks = []
        self.reflections = {}
        self.save_data()

    def save_data(self):
        """Save current tasks, reflections, and preferences to JSON."""
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "tasks": self.tasks,
                        "reflections": self.reflections,
                        "dark_mode": self.dark_mode,
                        "layout_mode": self.layout_mode,
                    },
                    f,
                    indent=2,
                )
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save data: {e}")

    def export_json(self):
        """Export backup JSON file via file dialog."""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=f"dayflow_backup_{self.current_date}.json",
            filetypes=[("JSON Files", "*.json")],
        )
        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(
                        {"tasks": self.tasks, "reflections": self.reflections},
                        f,
                        indent=2,
                    )
                messagebox.showinfo("Export Complete", "Backup exported successfully!")
            except Exception as e:
                messagebox.showerror("Export Error", str(e))

    def export_excel(self):
        """Export task records to an Excel-compatible workbook."""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"dayflow_tasks_{self.current_date}.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not filepath:
            return

        try:
            self.write_excel_workbook(filepath)
            messagebox.showinfo("Export Complete", "Tasks exported to Excel successfully!")
        except Exception as e:
            messagebox.showerror("Excel Export Error", str(e))

    def write_excel_workbook(self, filepath):
        """Write task records as a minimal Office Open XML workbook."""
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        ET.register_namespace("", main_ns)
        ET.register_namespace("r", rel_ns)

        headers = (
            "Task ID",
            "Title",
            "Date",
            "Time Range",
            "Category",
            "Priority",
            "Energy",
            "Status",
        )
        worksheet = ET.Element(f"{{{main_ns}}}worksheet")
        sheet_data = ET.SubElement(worksheet, f"{{{main_ns}}}sheetData")

        for row_number, values in enumerate(
            [headers]
            + [
                (
                    task.get("id", ""),
                    task.get("title", ""),
                    task.get("date", ""),
                    self.format_task_time_slot(task.get("timeSlot", ""))
                    if task.get("timeSlot")
                    else "Unscheduled",
                    task.get("category", ""),
                    task.get("priority", ""),
                    task.get("energy", ""),
                    "Completed" if task.get("completed") else "Pending",
                )
                for task in self.tasks
            ],
            start=1,
        ):
            row = ET.SubElement(sheet_data, f"{{{main_ns}}}row", {"r": str(row_number)})
            for column_number, value in enumerate(values, start=1):
                column_name = chr(64 + column_number)
                cell = ET.SubElement(
                    row,
                    f"{{{main_ns}}}c",
                    {"r": f"{column_name}{row_number}", "t": "inlineStr"},
                )
                inline_string = ET.SubElement(cell, f"{{{main_ns}}}is")
                text = ET.SubElement(inline_string, f"{{{main_ns}}}t")
                text.text = str(value)

        workbook = ET.Element(f"{{{main_ns}}}workbook")
        sheets = ET.SubElement(workbook, f"{{{main_ns}}}sheets")
        ET.SubElement(
            sheets,
            f"{{{main_ns}}}sheet",
            {
                "name": "Tasks",
                "sheetId": "1",
                f"{{{rel_ns}}}id": "rId1",
            },
        )

        root_rels = ET.Element(f"{{{package_rel_ns}}}Relationships")
        ET.SubElement(
            root_rels,
            f"{{{package_rel_ns}}}Relationship",
            {
                "Id": "rId1",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "Target": "xl/workbook.xml",
            },
        )
        workbook_rels = ET.Element(f"{{{package_rel_ns}}}Relationships")
        ET.SubElement(
            workbook_rels,
            f"{{{package_rel_ns}}}Relationship",
            {
                "Id": "rId1",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": "worksheets/sheet1.xml",
            },
        )
        content_types = ET.Element(f"{{{content_ns}}}Types")
        ET.SubElement(
            content_types,
            f"{{{content_ns}}}Default",
            {
                "Extension": "rels",
                "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
            },
        )
        ET.SubElement(
            content_types,
            f"{{{content_ns}}}Default",
            {"Extension": "xml", "ContentType": "application/xml"},
        )
        ET.SubElement(
            content_types,
            f"{{{content_ns}}}Override",
            {
                "PartName": "/xl/workbook.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            },
        )
        ET.SubElement(
            content_types,
            f"{{{content_ns}}}Override",
            {
                "PartName": "/xl/worksheets/sheet1.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            },
        )

        with zipfile.ZipFile(filepath, "w", compression=zipfile.ZIP_DEFLATED) as workbook_file:
            workbook_file.writestr(
                "[Content_Types].xml",
                ET.tostring(content_types, encoding="utf-8", xml_declaration=True),
            )
            workbook_file.writestr(
                "_rels/.rels",
                ET.tostring(root_rels, encoding="utf-8", xml_declaration=True),
            )
            workbook_file.writestr(
                "xl/workbook.xml",
                ET.tostring(workbook, encoding="utf-8", xml_declaration=True),
            )
            workbook_file.writestr(
                "xl/_rels/workbook.xml.rels",
                ET.tostring(workbook_rels, encoding="utf-8", xml_declaration=True),
            )
            workbook_file.writestr(
                "xl/worksheets/sheet1.xml",
                ET.tostring(worksheet, encoding="utf-8", xml_declaration=True),
            )

    def import_json(self):
        """Import tasks and reflections from a JSON backup."""
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filepath:
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "tasks" in data:
                        self.tasks = data["tasks"]
                    if "reflections" in data:
                        self.reflections = data["reflections"]
                    self.save_data()
                    self.refresh_all_views()
                    messagebox.showinfo("Import Complete", "Data imported successfully!")
            except Exception as e:
                messagebox.showerror("Import Error", f"Invalid JSON file: {e}")

    # --------------------------------------------------------------------------
    # UI CONSTRUCTION
    # --------------------------------------------------------------------------
    def build_ui(self):
        """Build the main header, tabs, and content containers."""
        # Top Accent Strip
        self.accent_bar = tk.Frame(self.root, height=4, bg="#4F46E5")
        self.accent_bar.pack(fill="x")

        # Header Frame
        self.header_frame = tk.Frame(self.root, padx=20, pady=12)
        self.header_frame.pack(fill="x")

        # Logo & Title
        logo_box = tk.Frame(self.header_frame, bg="#4F46E5", width=36, height=36)
        logo_box.pack_propagate(False)
        logo_box.pack(side="left")
        self.logo_box = logo_box
        tk.Label(
            logo_box,
            text="DF",
            bg="#4F46E5",
            fg="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
        ).pack(expand=True)

        title_box = tk.Frame(self.header_frame)
        title_box.pack(side="left", padx=12)
        self.title_lbl = tk.Label(
            title_box, text="DayFlow", font=("Segoe UI", 15, "bold")
        )
        self.title_lbl.pack(anchor="w")
        self.subtitle_lbl = tk.Label(
            title_box,
            text="Minimalist Daily Planner & Energy Batcher",
            font=("Segoe UI", 9),
        )
        self.subtitle_lbl.pack(anchor="w")
        self.title_box = title_box

        # Right Controls in Header
        controls_frame = tk.Frame(self.header_frame)
        controls_frame.pack(side="right")
        self.controls_frame = controls_frame

        date_controls_frame = tk.Frame(controls_frame)
        date_controls_frame.pack(side="left")
        self.date_controls_frame = date_controls_frame
        view_controls_frame = tk.Frame(self.header_frame)
        view_controls_frame.pack(side="top", fill="x")
        self.view_controls_frame = view_controls_frame

        # Date Navigation
        self.date_var = tk.StringVar(value=self.current_date)
        self.date_entry = tk.Entry(
            date_controls_frame,
            textvariable=self.date_var,
            width=12,
            font=("Segoe UI", 10, "bold"),
            justify="center",
            relief="flat",
        )
        self.date_entry.pack(side="left", padx=4, ipady=4)
        self.date_entry.bind("<Return>", lambda e: self.change_date())

        self.calendar_btn = tk.Button(
            date_controls_frame,
            text="📅",
            command=self.open_calendar,
            font=("Segoe UI", 10),
            relief="flat",
            padx=6,
            pady=4,
            cursor="hand2",
        )
        self.calendar_btn.pack(side="left", padx=(0, 4))

        self.today_btn = tk.Button(
            date_controls_frame,
            text="Today",
            command=self.select_today,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=8,
            pady=4,
            cursor="hand2",
        )
        self.today_btn.pack(side="left", padx=(0, 10))

        # Progress Pill
        self.progress_lbl = tk.Label(
            date_controls_frame,
            text="Progress: 0/0 (0%)",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
        )
        self.progress_lbl.pack(side="left", padx=8)

        # Theme Toggle Button
        self.theme_btn = tk.Button(
            view_controls_frame,
            text="🌙 Dark",
            command=self.toggle_theme,
            font=("Segoe UI", 9, "bold"),
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=5,
            cursor="hand2",
        )
        self.theme_btn.pack(side="right", padx=(4, 0))

        self.layout_toggle_btn = tk.Button(
            view_controls_frame,
            text="Switch to Mobile",
            command=self.toggle_layout_mode,
            font=("Segoe UI", 8, "bold"),
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=4,
            cursor="hand2",
        )
        self.layout_toggle_btn.pack(side="right", padx=(4, 0))

        # Notebook (Tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(4, 16))

        # Tab 1: Timeline
        self.timeline_tab = tk.Frame(self.notebook)
        self.notebook.add(self.timeline_tab, text="  🕒 Daily Timeline  ")
        self.build_timeline_tab()

        # Tab 2: Task Manager
        self.manager_tab = tk.Frame(self.notebook)
        self.notebook.add(self.manager_tab, text="  📋 Task & Priority Manager  ")
        self.build_manager_tab()

        # Tab 3: Daily Review
        self.review_tab = tk.Frame(self.notebook)
        self.notebook.add(self.review_tab, text="  📊 Daily Review & Summary  ")
        self.build_review_tab()

        self.tracker_tab = tk.Frame(self.notebook)
        self.notebook.add(self.tracker_tab, text="  📊 Task Trackers  ")
        self.tracker_notebook = ttk.Notebook(self.tracker_tab)
        self.tracker_notebook.pack(fill="both", expand=True)

        self.weekly_tracker_tab = tk.Frame(self.tracker_notebook)
        self.tracker_notebook.add(self.weekly_tracker_tab, text="Weekly")
        self.build_task_tracker_tab(self.weekly_tracker_tab, "week")

        self.monthly_tracker_tab = tk.Frame(self.tracker_notebook)
        self.tracker_notebook.add(self.monthly_tracker_tab, text="Monthly")
        self.build_task_tracker_tab(self.monthly_tracker_tab, "month")

        self.yearly_tracker_tab = tk.Frame(self.tracker_notebook)
        self.tracker_notebook.add(self.yearly_tracker_tab, text="Yearly")
        self.build_task_tracker_tab(self.yearly_tracker_tab, "year")

    # --------------------------------------------------------------------------
    # TAB 1: TIME-BLOCKED TIMELINE & UNSCHEDULED INBOX
    # --------------------------------------------------------------------------
    def build_timeline_tab(self):
        """Construct the 6 AM - 11 PM scrollable timeline + unscheduled sidebar."""
        container = tk.Frame(self.timeline_tab)
        container.pack(fill="both", expand=True, pady=10)
        self.tl_container = container

        mobile_panel_switcher = tk.Frame(container)
        self.mobile_panel_switcher = mobile_panel_switcher
        self.mobile_schedule_btn = tk.Button(
            mobile_panel_switcher,
            text="Schedule",
            command=lambda: self.toggle_mobile_timeline_panel("schedule"),
            font=("Segoe UI", 9, "bold"),
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=6,
            cursor="hand2",
        )
        self.mobile_schedule_btn.pack(side="left", fill="x", expand=True, padx=(0, 3))
        self.mobile_inbox_btn = tk.Button(
            mobile_panel_switcher,
            text="Unscheduled / Quick Tasks",
            command=lambda: self.toggle_mobile_timeline_panel("inbox"),
            font=("Segoe UI", 9, "bold"),
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=6,
            cursor="hand2",
        )
        self.mobile_inbox_btn.pack(side="left", fill="x", expand=True, padx=(3, 0))

        # Left Pane: Timeline (70% width)
        left_pane = tk.Frame(container)
        left_pane.pack(side="left", fill="both", expand=True, padx=(0, 12))
        self.tl_left_pane = left_pane

        tl_header = tk.Label(
            left_pane,
            text="Time-Blocked Schedule (6:00 AM – 11:00 PM)",
            font=("Segoe UI", 12, "bold"),
        )
        tl_header.pack(anchor="w", pady=(0, 8))
        self.tl_header_lbl = tl_header

        # Scrollable Canvas for Hours
        canvas_frame = tk.Frame(left_pane)
        canvas_frame.pack(fill="both", expand=True)

        self.tl_canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            canvas_frame, orient="vertical", command=self.tl_canvas.yview
        )
        self.tl_scrollable = tk.Frame(self.tl_canvas)

        self.tl_scrollable.bind(
            "<Configure>",
            lambda e: self.tl_canvas.configure(scrollregion=self.tl_canvas.bbox("all")),
        )
        self.tl_canvas_window = self.tl_canvas.create_window(
            (0, 0), window=self.tl_scrollable, anchor="nw"
        )
        self.tl_canvas.bind(
            "<Configure>",
            lambda e: self.tl_canvas.itemconfig(self.tl_canvas_window, width=e.width),
        )
        self.tl_canvas.configure(yscrollcommand=scrollbar.set)

        self.tl_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse wheel scrolling
        self.tl_canvas.bind_all("<MouseWheel>", self.handle_mousewheel)

        # Right Pane: Unscheduled Inbox & Backup (30% width)
        right_pane = tk.Frame(container, width=320)
        right_pane.pack(side="right", fill="y")
        right_pane.pack_propagate(False)
        self.tl_right_pane = right_pane

        inbox_header = tk.Label(
            right_pane,
            text="Unscheduled / Quick Tasks",
            font=("Segoe UI", 11, "bold"),
        )
        inbox_header.pack(anchor="w", pady=(0, 6))
        self.inbox_header_lbl = inbox_header

        add_backlog_btn = tk.Button(
            right_pane,
            text="+ New Task",
            command=lambda: self.open_task_modal(initial_time=""),
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=6,
            cursor="hand2",
        )
        add_backlog_btn.pack(fill="x", pady=(0, 8))
        self.add_backlog_btn = add_backlog_btn

        inbox_canvas_frame = tk.Frame(right_pane)
        inbox_canvas_frame.pack(fill="both", expand=True)
        self.inbox_canvas_frame = inbox_canvas_frame
        self.inbox_canvas = tk.Canvas(inbox_canvas_frame, highlightthickness=0)
        inbox_scrollbar = ttk.Scrollbar(
            inbox_canvas_frame,
            orient="vertical",
            command=self.inbox_canvas.yview,
        )
        self.inbox_list_frame = tk.Frame(self.inbox_canvas)
        self.inbox_canvas_window = self.inbox_canvas.create_window(
            (0, 0), window=self.inbox_list_frame, anchor="nw"
        )
        self.inbox_list_frame.bind(
            "<Configure>",
            lambda event: self.inbox_canvas.configure(
                scrollregion=self.inbox_canvas.bbox("all")
            ),
        )
        self.inbox_canvas.bind(
            "<Configure>",
            lambda event: self.inbox_canvas.itemconfig(
                self.inbox_canvas_window, width=event.width
            ),
        )
        self.inbox_canvas.configure(yscrollcommand=inbox_scrollbar.set)
        self.inbox_canvas.pack(side="left", fill="both", expand=True)
        inbox_scrollbar.pack(side="right", fill="y")

        # Backup / Sync Controls
        backup_frame = tk.Frame(right_pane, pady=8)
        backup_frame.pack(fill="x", side="bottom")
        self.backup_frame = backup_frame

        tk.Label(
            backup_frame,
            text="On-device Storage & Export",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        btn_row = tk.Frame(backup_frame)
        btn_row.pack(fill="x")

        self.export_excel_btn = tk.Button(
            btn_row,
            text="⬇ Excel",
            command=self.export_excel,
            font=("Segoe UI", 9),
            relief="flat",
            pady=5,
            cursor="hand2",
        )
        self.export_excel_btn.pack(side="left", fill="x", expand=True, padx=(0, 3))

        self.export_btn = tk.Button(
            btn_row,
            text="⬇ JSON",
            command=self.export_json,
            font=("Segoe UI", 9),
            relief="flat",
            pady=5,
            cursor="hand2",
        )
        self.export_btn.pack(side="left", fill="x", expand=True, padx=3)

        self.import_btn = tk.Button(
            btn_row,
            text="⬆ JSON",
            command=self.import_json,
            font=("Segoe UI", 9),
            relief="flat",
            pady=5,
            cursor="hand2",
        )
        self.import_btn.pack(side="left", fill="x", expand=True, padx=(3, 0))

    def handle_mousewheel(self, event):
        """Scroll the timeline or inbox depending on which canvas owns the pointer."""
        widget = event.widget
        canvas = self.tl_canvas
        while widget is not None:
            if widget == self.inbox_canvas:
                canvas = self.inbox_canvas
                break
            if widget == self.mgr_canvas:
                canvas = self.mgr_canvas
                break
            if widget == self.tl_canvas:
                canvas = self.tl_canvas
                break
            widget = getattr(widget, "master", None)
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def refresh_timeline(self):
        """Render all hourly slots (06:00 to 23:00) and unscheduled tasks."""
        for widget in self.tl_scrollable.winfo_children():
            widget.destroy()
        for widget in self.inbox_list_frame.winfo_children():
            widget.destroy()

        theme = self.themes["dark" if self.dark_mode else "light"]
        todays_tasks = [t for t in self.tasks if t.get("date") == self.current_date]

        scheduled = []
        for task in todays_tasks:
            if not task.get("timeSlot"):
                continue
            start_text, end_text = self.parse_task_time_slot(task["timeSlot"])
            if not start_text or not end_text:
                continue
            start = self.time_to_minutes(start_text)
            end = self.time_to_minutes(end_text)
            if start >= end or end <= 6 * 60 or start >= 24 * 60:
                continue
            scheduled.append((start, end, task))

        lane_ends = []
        task_lanes = {}
        for start, end, task in sorted(scheduled, key=lambda item: (item[0], item[1])):
            lane = next(
                (index for index, lane_end in enumerate(lane_ends) if lane_end <= start),
                len(lane_ends),
            )
            if lane == len(lane_ends):
                lane_ends.append(end)
            else:
                lane_ends[lane] = end
            task_lanes[task["id"]] = lane
        lane_count = max(1, len(lane_ends))

        # Render 6 AM to 11 PM slots
        for hour in range(6, 24):
            time_str = f"{hour:02d}:00"
            label_str = (
                "12:00 PM"
                if hour == 12
                else f"{hour - 12}:00 PM"
                if hour > 12
                else f"{hour}:00 AM"
            )

            slot_row = tk.Frame(
                self.tl_scrollable,
                bg=theme["card"],
                highlightbackground=theme["border"],
                highlightthickness=1,
                padx=10,
                pady=8,
            )
            slot_row.pack(fill="x", pady=3, padx=2)

            # Time column
            time_lbl = tk.Label(
                slot_row,
                text=label_str,
                width=10,
                anchor="w",
                bg=theme["card"],
                fg=theme["subtext"],
                font=("Segoe UI", 9, "bold"),
            )
            time_lbl.pack(side="left", anchor="n", pady=2)

            # Keep overlapping tasks in separate lanes across all hourly rows.
            slot_start = hour * 60
            slot_end = (hour + 1) * 60
            slot_tasks = [
                (start, end, task)
                for start, end, task in scheduled
                if start < slot_end and end > slot_start
            ]

            if not slot_tasks:
                tasks_container = tk.Frame(slot_row, bg=theme["card"])
                tasks_container.pack(side="left", fill="x", expand=True)
                empty_lbl = tk.Label(
                    tasks_container,
                    text="Free slot — click + to schedule",
                    bg=theme["card"],
                    fg=theme["subtext"],
                    font=("Segoe UI", 9, "italic"),
                )
                empty_lbl.pack(side="left", pady=2)
            else:
                lanes_container = tk.Frame(slot_row, bg=theme["card"])
                lanes_container.pack(side="left", fill="x", expand=True)
                for lane in range(lane_count):
                    lanes_container.grid_columnconfigure(lane, weight=1, uniform="task_lane")

                for start, end, task in slot_tasks:
                    lane = task_lanes[task["id"]]
                    lane_frame = tk.Frame(lanes_container, bg=theme["card"])
                    lane_frame.grid(
                        row=0,
                        column=lane,
                        sticky="nsew",
                        padx=(0 if lane == 0 else 3, 3),
                    )
                    if start >= slot_start:
                        self.render_task_card(lane_frame, task, compact=True)
                    else:
                        self.render_task_continuation(lane_frame, task)

            # Quick "+ Add" button for this slot
            add_slot_btn = tk.Button(
                slot_row,
                text="+",
                command=lambda ts=time_str: self.open_task_modal(initial_time=ts),
                bg=theme["card_alt"],
                fg=theme["accent"],
                font=("Segoe UI", 10, "bold"),
                relief="flat",
                width=3,
                cursor="hand2",
            )
            add_slot_btn.pack(side="right", anchor="n")

        # Render Unscheduled Inbox tasks
        unscheduled = [t for t in todays_tasks if not t.get("timeSlot")]
        if not unscheduled:
            tk.Label(
                self.inbox_list_frame,
                text="All tasks for today are scheduled!",
                bg=theme["bg"],
                fg=theme["subtext"],
                font=("Segoe UI", 9, "italic"),
                pady=20,
            ).pack()
        else:
            for task in unscheduled:
                self.render_task_card(self.inbox_list_frame, task, compact=True)

    # --------------------------------------------------------------------------
    # TAB 2: TASK & PRIORITY MANAGER
    # --------------------------------------------------------------------------
    def build_manager_tab(self):
        """Construct the searchable and filterable Task Manager tab."""
        container = tk.Frame(self.manager_tab)
        container.pack(fill="both", expand=True, pady=10)
        self.mgr_container = container

        # Filter Bar
        filter_bar = tk.Frame(container)
        filter_bar.pack(fill="x")
        self.mgr_filter_bar = filter_bar

        search_group = tk.Frame(filter_bar)
        search_group.pack(side="left")
        self.mgr_search_group = search_group
        tk.Label(search_group, text="Search:", font=("Segoe UI", 9, "bold")).pack(
            side="left", padx=(0, 6)
        )
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.refresh_manager())
        self.search_entry = tk.Entry(
            search_group, textvariable=self.search_var, width=26, font=("Segoe UI", 10)
        )
        self.search_entry.pack(side="left", padx=(0, 14), ipady=3)

        cat_group = tk.Frame(filter_bar)
        cat_group.pack(side="left")
        self.mgr_cat_group = cat_group
        tk.Label(cat_group, text="Category:", font=("Segoe UI", 9, "bold")).pack(
            side="left", padx=(0, 6)
        )
        self.cat_filter_var = tk.StringVar(value="All")
        cat_combo = ttk.Combobox(
            cat_group,
            textvariable=self.cat_filter_var,
            values=["All", "Work", "Personal", "Health", "Errands"],
            state="readonly",
            width=12,
        )
        cat_combo.pack(side="left", padx=(0, 10))
        cat_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_manager())

        pri_group = tk.Frame(filter_bar)
        pri_group.pack(side="left")
        self.mgr_pri_group = pri_group
        tk.Label(pri_group, text="Priority:", font=("Segoe UI", 9, "bold")).pack(
            side="left", padx=(0, 6)
        )
        self.pri_filter_var = tk.StringVar(value="All")
        pri_combo = ttk.Combobox(
            pri_group,
            textvariable=self.pri_filter_var,
            values=["All", "High", "Med", "Low"],
            state="readonly",
            width=10,
        )
        pri_combo.pack(side="left", padx=(0, 10))
        pri_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_manager())

        eng_group = tk.Frame(filter_bar)
        eng_group.pack(side="left")
        self.mgr_eng_group = eng_group
        tk.Label(eng_group, text="Energy:", font=("Segoe UI", 9, "bold")).pack(
            side="left", padx=(0, 6)
        )
        self.eng_filter_var = tk.StringVar(value="All")
        eng_combo = ttk.Combobox(
            eng_group,
            textvariable=self.eng_filter_var,
            values=["All", "Deep Work", "Refinement", "Admin"],
            state="readonly",
            width=12,
        )
        eng_combo.pack(side="left")
        eng_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_manager())

        # Scrollable Task List Area
        mgr_canvas_frame = tk.Frame(container)
        mgr_canvas_frame.pack(fill="both", expand=True)
        self.mgr_canvas_frame = mgr_canvas_frame
        self.mgr_canvas = tk.Canvas(mgr_canvas_frame, highlightthickness=0)
        mgr_scrollbar = ttk.Scrollbar(
            mgr_canvas_frame,
            orient="vertical",
            command=self.mgr_canvas.yview,
        )
        self.mgr_list_frame = tk.Frame(self.mgr_canvas)
        self.mgr_canvas_window = self.mgr_canvas.create_window(
            (0, 0), window=self.mgr_list_frame, anchor="nw"
        )
        self.mgr_list_frame.bind(
            "<Configure>",
            lambda event: self.mgr_canvas.configure(
                scrollregion=self.mgr_canvas.bbox("all")
            ),
        )
        self.mgr_canvas.bind(
            "<Configure>",
            lambda event: self.mgr_canvas.itemconfig(
                self.mgr_canvas_window, width=event.width
            ),
        )
        self.mgr_canvas.configure(yscrollcommand=mgr_scrollbar.set)
        self.mgr_canvas.pack(side="left", fill="both", expand=True)
        mgr_scrollbar.pack(side="right", fill="y")

    def refresh_manager(self):
        """Render filtered tasks in the Task Manager view."""
        for widget in self.mgr_list_frame.winfo_children():
            widget.destroy()

        theme = self.themes["dark" if self.dark_mode else "light"]
        query = self.search_var.get().strip().lower()
        cat = self.cat_filter_var.get()
        pri = self.pri_filter_var.get()
        eng = self.eng_filter_var.get()

        filtered = [
            t
            for t in self.tasks
            if t.get("date") == self.current_date
            and (cat == "All" or t.get("category") == cat)
            and (pri == "All" or t.get("priority") == pri)
            and (eng == "All" or t.get("energy") == eng)
            and (not query or query in t.get("title", "").lower())
        ]

        if not filtered:
            tk.Label(
                self.mgr_list_frame,
                text="No tasks match your current filters.",
                bg=theme["bg"],
                fg=theme["subtext"],
                font=("Segoe UI", 10, "italic"),
                pady=30,
            ).pack()
        else:
            for task in filtered:
                self.render_task_card(self.mgr_list_frame, task, compact=False)

        self.mgr_canvas.after_idle(self.mgr_canvas.yview_moveto, 0)

    # --------------------------------------------------------------------------
    # TAB 3: DAILY REVIEW & ANALYTICS
    # --------------------------------------------------------------------------
    def build_review_tab(self):
        """Construct the End-of-Day summary and reflection logger."""
        container = tk.Frame(self.review_tab)
        container.pack(fill="both", expand=True, pady=12, padx=10)
        self.rev_container = container

        # Analytics Cards Row
        stats_row = tk.Frame(container)
        stats_row.pack(fill="x", pady=(0, 16))
        self.rev_stats_row = stats_row

        self.stat_completion_lbl = tk.Label(
            stats_row,
            text="Completion Rate: 0%",
            font=("Segoe UI", 12, "bold"),
            padx=18,
            pady=14,
        )
        self.stat_completion_lbl.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.stat_deepwork_lbl = tk.Label(
            stats_row,
            text="Deep Work Completed: 0",
            font=("Segoe UI", 12, "bold"),
            padx=18,
            pady=14,
        )
        self.stat_deepwork_lbl.pack(side="left", fill="x", expand=True, padx=8)

        self.stat_pending_lbl = tk.Label(
            stats_row,
            text="Pending Tasks: 0",
            font=("Segoe UI", 12, "bold"),
            padx=18,
            pady=14,
        )
        self.stat_pending_lbl.pack(side="left", fill="x", expand=True, padx=(8, 0))

        # Reflection Form
        tk.Label(
            container,
            text="🏆 Key Wins & Accomplishments",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(6, 4))
        self.wins_text = tk.Text(container, height=4, font=("Segoe UI", 10), wrap="word")
        self.wins_text.pack(fill="x", pady=(0, 12))

        tk.Label(
            container,
            text="🚧 Friction Points & Blockers",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(4, 4))
        self.blockers_text = tk.Text(
            container, height=4, font=("Segoe UI", 10), wrap="word"
        )
        self.blockers_text.pack(fill="x", pady=(0, 12))

        mood_row = tk.Frame(container)
        mood_row.pack(fill="x", pady=(4, 14))
        self.rev_mood_row = mood_row

        tk.Label(
            mood_row,
            text="⚡ Overall Energy / Mood:",
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left", padx=(0, 10))

        self.mood_var = tk.StringVar(value="Productive")
        mood_combo = ttk.Combobox(
            mood_row,
            textvariable=self.mood_var,
            values=["Productive", "Flow State", "Energized", "Tired", "Distracted"],
            state="readonly",
            width=18,
        )
        mood_combo.pack(side="left")

        save_rev_btn = tk.Button(
            container,
            text="Save Daily Review",
            command=self.save_reflection,
            bg="#10B981",
            fg="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            cursor="hand2",
        )
        save_rev_btn.pack(anchor="w")

    def default_tracker_range(self, period, anchor):
        """Return the default date range for a weekly, monthly, or yearly tracker."""
        if period == "week":
            start_date = anchor - timedelta(days=anchor.weekday())
            return start_date, start_date + timedelta(days=6)
        if period == "month":
            start_date = anchor.replace(day=1)
            if anchor.month == 12:
                next_month = anchor.replace(year=anchor.year + 1, month=1, day=1)
            else:
                next_month = anchor.replace(month=anchor.month + 1, day=1)
            return start_date, next_month - timedelta(days=1)
        return anchor.replace(month=1, day=1), anchor.replace(month=12, day=31)

    def build_task_tracker_tab(self, parent, period):
        """Build a review table with date controls for the selected period."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        container = tk.Frame(parent, padx=16, pady=14)
        container.pack(fill="both", expand=True)

        title = tk.Label(
            container,
            text="",
            bg=theme["bg"],
            fg=theme["text"],
            font=("Segoe UI", 14, "bold"),
        )
        title.pack(anchor="w")
        preset_frame = tk.Frame(container, bg=theme["bg"])

        if period in ("week", "month"):
            start_date, end_date = self.default_tracker_range(
                period, datetime.strptime(self.current_date, "%Y-%m-%d").date()
            )
            from_var = tk.StringVar(value=start_date.isoformat())
            to_var = tk.StringVar(value=end_date.isoformat())
            range_frame = tk.Frame(container, bg=theme["bg"])
            range_frame.pack(fill="x", pady=(8, 2))
            from_label = tk.Label(
                range_frame, text="From", bg=theme["bg"], fg=theme["text"]
            )
            from_label.pack(side="left", padx=(0, 4))
            from_entry = tk.Entry(
                range_frame, textvariable=from_var, width=10, justify="center"
            )
            from_entry.pack(side="left")
            from_calendar_btn = tk.Button(
                range_frame,
                text="📅",
                command=lambda var=from_var: self.open_date_picker(var, from_calendar_btn),
                relief="flat",
                cursor="hand2",
            )
            from_calendar_btn.pack(side="left", padx=(2, 10))
            to_label = tk.Label(
                range_frame, text="To", bg=theme["bg"], fg=theme["text"]
            )
            to_label.pack(side="left", padx=(0, 4))
            to_entry = tk.Entry(
                range_frame, textvariable=to_var, width=10, justify="center"
            )
            to_entry.pack(side="left")
            to_calendar_btn = tk.Button(
                range_frame,
                text="📅",
                command=lambda var=to_var: self.open_date_picker(var, to_calendar_btn),
                relief="flat",
                cursor="hand2",
            )
            to_calendar_btn.pack(side="left", padx=(2, 8))
            apply_range_btn = tk.Button(
                range_frame,
                text="Apply Range",
                command=lambda selected_period=period: self.apply_tracker_range(
                    selected_period
                ),
                relief="flat",
                cursor="hand2",
            )
            apply_range_btn.pack(side="left")
            self.tracker_ranges[period] = {
                "from_var": from_var,
                "to_var": to_var,
                "frame": range_frame,
                "from_label": from_label,
                "to_label": to_label,
                "from_entry": from_entry,
                "to_entry": to_entry,
                "from_calendar_btn": from_calendar_btn,
                "to_calendar_btn": to_calendar_btn,
                "apply_button": apply_range_btn,
                "start_date": start_date,
                "end_date": end_date,
            }
        else:
            year_frame = tk.Frame(container, bg=theme["bg"])
            year_frame.pack(anchor="w", pady=(8, 2))
            self.year_tracker_frame = year_frame
            year_label = tk.Label(
                year_frame, text="Year", bg=theme["bg"], fg=theme["text"]
            )
            year_label.pack(
                side="left", padx=(0, 6)
            )
            self.year_tracker_label = year_label
            self.year_tracker_var = tk.StringVar(
                value=str(datetime.strptime(self.current_date, "%Y-%m-%d").year)
            )
            self.year_tracker_menu = ttk.Combobox(
                year_frame,
                textvariable=self.year_tracker_var,
                values=tuple(
                    str(year)
                    for year in range(
                        max(1900, datetime.now().year - 100),
                        min(9999, datetime.now().year + 21),
                    )
                ),
                state="readonly",
                width=8,
            )
            self.year_tracker_menu.pack(side="left")
            self.year_tracker_menu.bind(
                "<<ComboboxSelected>>",
                lambda event: self.refresh_task_tracker("year"),
            )
            self.yearly_tracker_container = container

        preset_frame.pack(anchor="w", pady=(6, 2))
        preset_button = tk.Button(
            preset_frame,
            text={"week": "This Week", "month": "This Month", "year": "This Year"}[
                period
            ],
            command=lambda selected_period=period: self.select_current_period(
                selected_period
            ),
            bg=theme["accent"],
            fg=theme["accent_fg"],
            activebackground=theme["accent_hover"],
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
        )
        preset_button.pack(side="left")
        if period in ("week", "month"):
            self.tracker_ranges[period]["preset_frame"] = preset_frame
            self.tracker_ranges[period]["preset_button"] = preset_button
        else:
            self.year_preset_frame = preset_frame
            self.year_preset_button = preset_button

        summary = tk.Label(
            container,
            text="",
            bg=theme["bg"],
            fg=theme["subtext"],
            font=("Segoe UI", 10),
        )
        summary.pack(anchor="w", pady=(3, 10))

        table_frame = tk.Frame(container, bg=theme["bg"])
        table_frame.pack(fill="both", expand=True)
        columns = ("date", "status", "task", "time", "category", "priority")
        tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            style="Tracker.Treeview",
        )
        headings = {
            "date": ("Date", 105),
            "status": ("Status", 95),
            "task": ("Task", 300),
            "time": ("Time", 130),
            "category": ("Category", 100),
            "priority": ("Priority", 90),
        }
        for column, (heading, width) in headings.items():
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=65, anchor="w")

        vertical_scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=tree.yview
        )
        horizontal_scrollbar = ttk.Scrollbar(
            table_frame, orient="horizontal", command=tree.xview
        )
        tree.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")

        if period == "week":
            self.weekly_tracker_title = title
            self.weekly_tracker_summary = summary
            self.weekly_tracker_tree = tree
            self.weekly_tracker_container = container
        elif period == "month":
            self.monthly_tracker_title = title
            self.monthly_tracker_summary = summary
            self.monthly_tracker_tree = tree
            self.monthly_tracker_container = container
        else:
            self.yearly_tracker_title = title
            self.yearly_tracker_summary = summary
            self.yearly_tracker_tree = tree

    def apply_tracker_range(self, period):
        """Apply validated From/To dates to a weekly or monthly tracker."""
        range_state = self.tracker_ranges[period]
        try:
            start_date = datetime.strptime(
                range_state["from_var"].get().strip(), "%Y-%m-%d"
            ).date()
            end_date = datetime.strptime(
                range_state["to_var"].get().strip(), "%Y-%m-%d"
            ).date()
        except ValueError:
            messagebox.showwarning(
                "Invalid Date Range", "Use YYYY-MM-DD dates for both range fields."
            )
            return

        if start_date > end_date:
            messagebox.showwarning(
                "Invalid Date Range", "The From date must be on or before the To date."
            )
            return

        range_state["start_date"] = start_date
        range_state["end_date"] = end_date
        self.refresh_task_tracker(period)

    def select_current_period(self, period):
        """Show the current calendar week, month, or year in its tracker."""
        today = datetime.now().date()
        if period == "year":
            self.year_tracker_var.set(str(today.year))
            self.refresh_task_tracker("year")
            return

        start_date, end_date = self.default_tracker_range(period, today)
        range_state = self.tracker_ranges[period]
        range_state["from_var"].set(start_date.isoformat())
        range_state["to_var"].set(end_date.isoformat())
        range_state["start_date"] = start_date
        range_state["end_date"] = end_date
        self.refresh_task_tracker(period)

    def reset_tracker_ranges(self):
        """Reset weekly and monthly ranges to the selected planner date."""
        anchor = datetime.strptime(self.current_date, "%Y-%m-%d").date()
        for period in ("week", "month"):
            start_date, end_date = self.default_tracker_range(period, anchor)
            state = self.tracker_ranges[period]
            state["from_var"].set(start_date.isoformat())
            state["to_var"].set(end_date.isoformat())
            state["start_date"] = start_date
            state["end_date"] = end_date
        self.year_tracker_var.set(str(anchor.year))

    def refresh_task_tracker(self, period):
        """Refresh a weekly, monthly, or yearly task list."""
        if period in ("week", "month"):
            range_state = self.tracker_ranges[period]
            start_date = range_state["start_date"]
            end_date = range_state["end_date"]
            if period == "week":
                title_prefix = "Weekly Task Tracker"
                tree = self.weekly_tracker_tree
                title_label = self.weekly_tracker_title
                summary_label = self.weekly_tracker_summary
            else:
                title_prefix = "Monthly Task Tracker"
                tree = self.monthly_tracker_tree
                title_label = self.monthly_tracker_title
                summary_label = self.monthly_tracker_summary
        else:
            try:
                selected_year = int(self.year_tracker_var.get())
                start_date = datetime(selected_year, 1, 1).date()
                end_date = datetime(selected_year, 12, 31).date()
            except ValueError:
                return
            title_prefix = "Yearly Task Tracker"
            tree = self.yearly_tracker_tree
            title_label = self.yearly_tracker_title
            summary_label = self.yearly_tracker_summary

        period_label = f"{start_date:%b %d, %Y} – {end_date:%b %d, %Y}"
        if self.layout_mode == "Mobile":
            title = f"{title_prefix}\n{period_label}"
        elif period == "year":
            title = f"{title_prefix} · {selected_year}"
        else:
            title = f"{title_prefix} · {period_label}"

        period_tasks = []
        for task in self.tasks:
            try:
                task_date = datetime.strptime(task.get("date", ""), "%Y-%m-%d").date()
            except (TypeError, ValueError):
                continue
            if start_date <= task_date <= end_date:
                period_tasks.append((task_date, task))

        def sort_key(item):
            task_date, task = item
            start_text, _ = self.parse_task_time_slot(task.get("timeSlot", ""))
            start_minutes = self.time_to_minutes(start_text) if start_text else 24 * 60
            return task_date, start_minutes, task.get("title", "").lower()

        period_tasks.sort(key=sort_key)
        completed = sum(1 for _, task in period_tasks if task.get("completed"))
        title_label.config(
            text=title,
            font=("Segoe UI", 11 if self.layout_mode == "Mobile" else 14, "bold"),
        )
        summary_label.config(
            text=f"{completed} of {len(period_tasks)} tasks completed"
        )

        for row in tree.get_children():
            tree.delete(row)
        for task_date, task in period_tasks:
            tree.insert(
                "",
                "end",
                values=(
                    task_date.strftime("%a, %b %d"),
                    "Completed" if task.get("completed") else "Pending",
                    task.get("title", ""),
                    self.format_task_time_slot(task.get("timeSlot", ""))
                    if task.get("timeSlot")
                    else "Unscheduled",
                    task.get("category", ""),
                    task.get("priority", ""),
                ),
                tags=("completed",) if task.get("completed") else ("pending",),
            )

        if not period_tasks:
            tree.insert("", "end", values=("", "", "No tasks in this period", "", "", ""))

    def refresh_task_trackers(self):
        self.refresh_task_tracker("week")
        self.refresh_task_tracker("month")
        self.refresh_task_tracker("year")

    def refresh_review(self):
        """Populate analytics and saved reflection text for the selected date."""
        todays_tasks = [t for t in self.tasks if t.get("date") == self.current_date]
        total = len(todays_tasks)
        completed = sum(1 for t in todays_tasks if t.get("completed"))
        pending = total - completed
        rate = int((completed / total) * 100) if total > 0 else 0
        deep_work = sum(
            1
            for t in todays_tasks
            if t.get("completed") and t.get("energy") == "Deep Work"
        )

        self.stat_completion_lbl.config(
            text=f"Completion Rate: {rate}% ({completed}/{total})"
        )
        self.stat_deepwork_lbl.config(text=f"Deep Work Sessions Done: {deep_work}")
        self.stat_pending_lbl.config(text=f"Pending Tasks: {pending}")

        ref = self.reflections.get(
            self.current_date, {"wins": "", "blocker": "", "mood": "Productive"}
        )
        self.wins_text.delete("1.0", tk.END)
        self.wins_text.insert("1.0", ref.get("wins", ""))

        self.blockers_text.delete("1.0", tk.END)
        self.blockers_text.insert("1.0", ref.get("blocker", ""))

        self.mood_var.set(ref.get("mood", "Productive"))

    def save_reflection(self):
        """Persist daily review notes to local JSON storage."""
        self.reflections[self.current_date] = {
            "wins": self.wins_text.get("1.0", tk.END).strip(),
            "blocker": self.blockers_text.get("1.0", tk.END).strip(),
            "mood": self.mood_var.get(),
        }
        self.save_data()
        messagebox.showinfo("Saved", "Your daily review has been saved!")

    # --------------------------------------------------------------------------
    # REUSABLE TASK CARD WIDGET
    # --------------------------------------------------------------------------
    def render_task_continuation(self, parent, task):
        """Render the later-hour segment of a task in its assigned timeline lane."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        category = task.get("category", "Work")
        category_style = self.category_colors.get(
            category, {"bg": "#E2E8F0", "fg": "#334155", "border": "#64748B"}
        )
        segment = tk.Frame(
            parent,
            bg=theme["card_alt"],
            highlightbackground=category_style["border"],
            highlightthickness=1,
            padx=6,
            pady=6,
        )
        segment.pack(fill="x", pady=3)
        tk.Frame(segment, bg=category_style["border"], width=4).pack(
            side="left", fill="y", padx=(0, 6)
        )
        details = tk.Frame(segment, bg=theme["card_alt"])
        details.pack(side="left", fill="x", expand=True)
        tk.Label(
            details,
            text=f"CONTINUED BOOKING · {self.format_task_time_slot(task.get('timeSlot', ''))}",
            bg=theme["card_alt"],
            fg=category_style["fg"],
            font=("Segoe UI", 7, "bold"),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            details,
            text=task.get("title", ""),
            bg=theme["card_alt"],
            fg=theme["subtext"] if task.get("completed") else theme["text"],
            font=("Segoe UI", 9, "overstrike" if task.get("completed") else "bold"),
            anchor="w",
        ).pack(anchor="w")

    def render_task_card(self, parent, task, compact=False):
        """Render a single task item with checkbox, badges, quick hour move, edit, and delete."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        cat = task.get("category", "Work")
        pri = task.get("priority", "Med")
        eng = task.get("energy", "Deep Work")
        done = task.get("completed", False)

        cat_style = self.category_colors.get(
            cat, {"bg": "#E2E8F0", "fg": "#334155", "border": "#64748B"}
        )

        card = tk.Frame(
            parent,
            bg=theme["card_alt"] if compact else theme["card"],
            highlightbackground=cat_style["border"] if not done else theme["border"],
            highlightthickness=1,
            padx=10,
            pady=6,
        )
        card.pack(fill="x", pady=3)

        # Left color strip
        strip = tk.Frame(card, bg=cat_style["border"], width=4)
        strip.pack(side="left", fill="y", padx=(0, 8))

        # Completion checkbox
        checkbox = tk.Canvas(
            card,
            width=18,
            height=18,
            bg=card["bg"],
            highlightthickness=0,
            cursor="hand2",
        )
        checkbox.create_rectangle(
            2,
            2,
            16,
            16,
            outline=theme["accent"] if done else theme["border"],
            fill=theme["accent"] if done else card["bg"],
            width=2,
        )
        if done:
            checkbox.create_line(
                5,
                9,
                8,
                12,
                14,
                5,
                fill="#FFFFFF",
                width=2,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND,
            )
        checkbox.bind(
            "<Button-1>",
            lambda event, tid=task["id"]: self.toggle_task(tid),
        )
        checkbox.pack(side="left")

        # Title & Badges
        info_frame = tk.Frame(card, bg=card["bg"])
        info_frame.pack(side="left", fill="x", expand=True, padx=4)

        title_font = (
            ("Segoe UI", 10, "overstrike") if done else ("Segoe UI", 10, "bold")
        )
        title_lbl = tk.Label(
            info_frame,
            text=task.get("title", ""),
            bg=card["bg"],
            fg=theme["subtext"] if done else theme["text"],
            font=title_font,
            anchor="w",
        )
        title_lbl.pack(anchor="w")

        meta_frame = tk.Frame(info_frame, bg=card["bg"])
        meta_frame.pack(anchor="w", pady=(2, 0))

        # Category badge
        tk.Label(
            meta_frame,
            text=cat,
            bg=cat_style["bg"],
            fg=cat_style["fg"],
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
        ).pack(side="left", padx=(0, 6))

        # Priority badge
        tk.Label(
            meta_frame,
            text=f"● {pri}",
            bg=card["bg"],
            fg=self.priority_colors.get(pri, "#64748B"),
            font=("Segoe UI", 8, "bold"),
        ).pack(side="left", padx=(0, 8))

        # Energy badge
        tk.Label(
            meta_frame,
            text=f"⚡ {eng}",
            bg=card["bg"],
            fg=self.energy_colors.get(eng, "#6366F1"),
            font=("Segoe UI", 8, "bold"),
        ).pack(side="left", padx=(0, 8))

        # Scheduled time if present
        if task.get("timeSlot"):
            tk.Label(
                meta_frame,
                text=f"🕒 {self.format_task_time_slot(task['timeSlot'])}",
                bg=card["bg"],
                fg=theme["subtext"],
                font=("Segoe UI", 8),
            ).pack(side="left")

        # Action Buttons (Edit & Delete)
        actions = tk.Frame(card, bg=card["bg"])
        actions.pack(side="right")

        tk.Button(
            actions,
            text="✎",
            command=lambda t=task: self.open_task_modal(task=t),
            bg=card["bg"],
            fg=theme["subtext"],
            font=("Segoe UI", 10),
            relief="flat",
            cursor="hand2",
        ).pack(side="left", padx=2)

        tk.Button(
            actions,
            text="🗑",
            command=lambda tid=task["id"]: self.delete_task(tid),
            bg=card["bg"],
            fg="#E11D48",
            font=("Segoe UI", 10),
            relief="flat",
            cursor="hand2",
        ).pack(side="left", padx=2)

    # --------------------------------------------------------------------------
    # TASK CRUD & MODAL DIALOG
    # --------------------------------------------------------------------------
    def time_to_minutes(self, time_text):
        """Convert an HH:MM time string to minutes after midnight."""
        try:
            hour, minute = time_text.split(":")
            return int(hour) * 60 + int(minute)
        except (AttributeError, TypeError, ValueError):
            return 0

    def minutes_to_time(self, minutes):
        """Convert minutes after midnight to an HH:MM time string."""
        minutes = max(0, min(minutes, 23 * 60 + 59))
        return f"{minutes // 60:02d}:{minutes % 60:02d}"

    def parse_task_time_slot(self, time_slot):
        """Return the start and end of a task range; legacy values default to one hour."""
        if not time_slot:
            return "", ""
        if "-" in time_slot:
            start, end = time_slot.split("-", 1)
            return start.strip(), end.strip()
        start = time_slot.strip()
        end = self.minutes_to_time(min(self.time_to_minutes(start) + 60, 23 * 60 + 30))
        return start, end

    def format_task_time_slot(self, time_slot):
        """Format a task range for display, preserving the existing single-time format."""
        start, end = self.parse_task_time_slot(time_slot)
        return f"{start} - {end}" if start and end else ""

    def task_time_ranges_overlap(self, task, range_start, range_end, _same_date=True):
        """Check whether a task range overlaps an inclusive/exclusive time interval."""
        start_text, end_text = self.parse_task_time_slot(task.get("timeSlot", ""))
        if not start_text or not end_text:
            return False
        start = self.time_to_minutes(start_text)
        end = self.time_to_minutes(end_text)
        return start < range_end and end > range_start

    def find_time_conflict(self, start_text, end_text, exclude_task_id=None):
        """Find a task on the active date whose range overlaps the proposed range."""
        start = self.time_to_minutes(start_text)
        end = self.time_to_minutes(end_text)
        for other in self.tasks:
            if other.get("id") == exclude_task_id:
                continue
            if other.get("date") != self.current_date or not other.get("timeSlot"):
                continue
            other_start_text, other_end_text = self.parse_task_time_slot(other["timeSlot"])
            if not other_start_text or not other_end_text:
                continue
            other_start = self.time_to_minutes(other_start_text)
            other_end = self.time_to_minutes(other_end_text)
            if start < other_end and end > other_start:
                return other
        return None

    def open_task_modal(self, task=None, initial_time="09:00"):
        """Open a clean dialog to add a new task or edit an existing one."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        modal = tk.Toplevel(self.root)
        modal.title("Edit Task" if task else "New Task")
        modal.geometry("360x430" if self.layout_mode == "Mobile" else "420x390")
        modal.configure(bg=theme["card"])
        modal.transient(self.root)
        modal.grab_set()

        tk.Label(
            modal,
            text="Edit Task" if task else "Create New Task",
            bg=theme["card"],
            fg=theme["text"],
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w", padx=20, pady=(18, 12))

        form = tk.Frame(modal, bg=theme["card"], padx=20)
        form.pack(fill="both", expand=True)

        # Title
        tk.Label(
            form, text="Task Title", bg=theme["card"], fg=theme["subtext"], font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        title_var = tk.StringVar(value=task["title"] if task else "")
        title_entry = tk.Entry(form, textvariable=title_var, font=("Segoe UI", 10))
        title_entry.pack(fill="x", pady=(2, 10), ipady=4)
        title_entry.focus()

        # Category & Priority Row
        row1 = tk.Frame(form, bg=theme["card"])
        row1.pack(fill="x", pady=(0, 10))

        col1 = tk.Frame(row1, bg=theme["card"])
        col1.pack(side="left", fill="x", expand=True, padx=(0, 6))
        tk.Label(
            col1, text="Category", bg=theme["card"], fg=theme["subtext"], font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        cat_var = tk.StringVar(value=task["category"] if task else "Work")
        ttk.Combobox(
            col1,
            textvariable=cat_var,
            values=["Work", "Personal", "Health", "Errands"],
            state="readonly",
        ).pack(fill="x", pady=2)

        col2 = tk.Frame(row1, bg=theme["card"])
        col2.pack(side="right", fill="x", expand=True, padx=(6, 0))
        tk.Label(
            col2, text="Priority", bg=theme["card"], fg=theme["subtext"], font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        pri_var = tk.StringVar(value=task["priority"] if task else "Med")
        ttk.Combobox(
            col2,
            textvariable=pri_var,
            values=["High", "Med", "Low"],
            state="readonly",
        ).pack(fill="x", pady=2)

        # Energy & Time Slot Row
        row2 = tk.Frame(form, bg=theme["card"])
        row2.pack(fill="x", pady=(0, 16))

        col3 = tk.Frame(row2, bg=theme["card"])
        col3.pack(side="left", fill="x", expand=True, padx=(0, 6))
        tk.Label(
            col3, text="Energy / Focus", bg=theme["card"], fg=theme["subtext"], font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        eng_var = tk.StringVar(value=task["energy"] if task else "Deep Work")
        ttk.Combobox(
            col3,
            textvariable=eng_var,
            values=["Deep Work", "Refinement", "Admin"],
            state="readonly",
        ).pack(fill="x", pady=2)

        col4 = tk.Frame(row2, bg=theme["card"])
        col4.pack(side="right", fill="x", expand=True, padx=(6, 0))
        tk.Label(
            col4, text="Time Range", bg=theme["card"], fg=theme["subtext"], font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")

        time_options = []
        for minutes in range(6 * 60, 24 * 60, 30):
            time_options.append(f"{minutes // 60:02d}:{minutes % 60:02d}")

        existing_start = ""
        existing_end = ""
        if task and task.get("timeSlot"):
            existing_start, existing_end = self.parse_task_time_slot(task["timeSlot"])

        start_slot = existing_start or (initial_time if initial_time else "09:00")
        if existing_end:
            end_slot = existing_end
        else:
            start_minutes = self.time_to_minutes(start_slot)
            end_slot = self.minutes_to_time(min(start_minutes + 60, 23 * 60 + 30))

        start_var = tk.StringVar(value=start_slot)
        end_var = tk.StringVar(value=end_slot)
        range_frame = tk.Frame(col4, bg=theme["card"])
        range_frame.pack(fill="x", pady=2)
        ttk.Combobox(
            range_frame, textvariable=start_var, values=time_options, state="readonly", width=7
        ).pack(side="left", fill="x", expand=True)
        tk.Label(
            range_frame, text=" to ", bg=theme["card"], fg=theme["subtext"]
        ).pack(side="left")
        ttk.Combobox(
            range_frame, textvariable=end_var, values=time_options, state="readonly", width=7
        ).pack(side="left", fill="x", expand=True)

        unscheduled_var = tk.BooleanVar(
            value=bool(
                (task and not task.get("timeSlot"))
                or (not task and not initial_time)
            )
        )
        def toggle_unscheduled():
            disabled = unscheduled_var.get()
            start_combo.state(["disabled"] if disabled else ["!disabled"])
            end_combo.state(["disabled"] if disabled else ["!disabled"])

        unscheduled_check = tk.Checkbutton(
            col4,
            text="Unscheduled",
            variable=unscheduled_var,
            command=toggle_unscheduled,
            bg=theme["card"],
            fg=theme["subtext"],
            activebackground=theme["card"],
            selectcolor=theme["card"],
        )
        unscheduled_check.pack(anchor="w", pady=(4, 0))
        start_combo = range_frame.winfo_children()[0]
        end_combo = range_frame.winfo_children()[2]
        toggle_unscheduled()

        def on_save():
            clean_title = title_var.get().strip()
            if not clean_title:
                messagebox.showwarning("Validation", "Please enter a task title.")
                return

            chosen_slot = ""
            if not unscheduled_var.get():
                start_slot = start_var.get()
                end_slot = end_var.get()
                start_minutes = self.time_to_minutes(start_slot)
                end_minutes = self.time_to_minutes(end_slot)
                if end_minutes <= start_minutes:
                    messagebox.showwarning(
                        "Invalid Time Range",
                        "The end time must be later than the start time.",
                    )
                    return

                conflict = self.find_time_conflict(
                    start_slot,
                    end_slot,
                    exclude_task_id=task["id"] if task else None,
                )
                if conflict:
                    save_anyway = messagebox.askyesno(
                        "Time Conflict",
                        f'There is a conflict with "{conflict.get("title", "another task")}" '
                        f'({self.format_task_time_slot(conflict["timeSlot"])}).\n\n'
                        "Save anyway?",
                        icon="warning",
                    )
                    if not save_anyway:
                        return

                chosen_slot = f"{start_slot}-{end_slot}"

            if task:
                task["title"] = clean_title
                task["category"] = cat_var.get()
                task["priority"] = pri_var.get()
                task["energy"] = eng_var.get()
                task["timeSlot"] = chosen_slot
            else:
                self.tasks.append(
                    {
                        "id": str(uuid.uuid4()),
                        "title": clean_title,
                        "category": cat_var.get(),
                        "priority": pri_var.get(),
                        "energy": eng_var.get(),
                        "timeSlot": chosen_slot,
                        "completed": False,
                        "date": self.current_date,
                    }
                )

            self.save_data()
            self.refresh_all_views()
            modal.destroy()

        btn_frame = tk.Frame(form, bg=theme["card"])
        btn_frame.pack(fill="x", pady=10)

        tk.Button(
            btn_frame,
            text="Save Task",
            command=on_save,
            bg="#4F46E5",
            fg="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
        ).pack(side="right")

        tk.Button(
            btn_frame,
            text="Cancel",
            command=modal.destroy,
            bg=theme["card_alt"],
            fg=theme["text"],
            font=("Segoe UI", 10),
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
        ).pack(side="right", padx=8)

    def toggle_task(self, task_id):
        """Toggle completion status of a task."""
        for t in self.tasks:
            if t["id"] == task_id:
                t["completed"] = not t.get("completed", False)
                break
        self.save_data()
        self.refresh_all_views()

    def delete_task(self, task_id):
        """Remove a task after confirmation."""
        if messagebox.askyesno("Delete Task", "Remove this task?"):
            self.tasks = [t for t in self.tasks if t["id"] != task_id]
            self.save_data()
            self.refresh_all_views()

    def open_calendar(self):
        """Open the calendar for selecting the planner date."""
        self.open_date_picker(self.date_var, self.calendar_btn, self.change_date)

    def open_date_picker(self, date_var, anchor_widget, on_select=None):
        """Open a calendar that writes the selected day into a date field."""
        try:
            selected = datetime.strptime(date_var.get().strip(), "%Y-%m-%d")
        except ValueError:
            selected = datetime.strptime(self.current_date, "%Y-%m-%d")

        theme = self.themes["dark" if self.dark_mode else "light"]
        popup = tk.Toplevel(self.root)
        popup.title("Select Date")
        popup.configure(bg=theme["card"])
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.geometry(
            f"+{anchor_widget.winfo_rootx()}+"
            f"{anchor_widget.winfo_rooty() + anchor_widget.winfo_height()}"
        )

        month_state = {"year": selected.year, "month": selected.month}
        header = tk.Frame(popup, bg=theme["card"], padx=8, pady=8)
        header.pack(fill="x")
        days_frame = tk.Frame(popup, bg=theme["card"], padx=8, pady=0)
        days_frame.pack(pady=(0, 8))

        def render_month():
            for widget in header.winfo_children():
                widget.destroy()
            for widget in days_frame.winfo_children():
                widget.destroy()

            year = month_state["year"]
            month = month_state["month"]
            tk.Button(
                header,
                text="‹",
                command=lambda: change_month(-1),
                bg=theme["card_alt"],
                fg=theme["text"],
                relief="flat",
                width=3,
            ).pack(side="left")
            tk.Label(
                header,
                text=f"{calendar.month_name[month]} {year}",
                bg=theme["card"],
                fg=theme["text"],
                font=("Segoe UI", 10, "bold"),
                width=18,
            ).pack(side="left", expand=True)
            tk.Button(
                header,
                text="›",
                command=lambda: change_month(1),
                bg=theme["card_alt"],
                fg=theme["text"],
                relief="flat",
                width=3,
            ).pack(side="right")

            for column, weekday in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
                tk.Label(
                    days_frame,
                    text=weekday,
                    bg=theme["card"],
                    fg=theme["subtext"],
                    width=4,
                    font=("Segoe UI", 8, "bold"),
                ).grid(row=0, column=column, pady=(0, 4))

            for row, week in enumerate(calendar.monthcalendar(year, month), start=1):
                for column, day in enumerate(week):
                    if not day:
                        continue
                    is_selected = (
                        year == selected.year
                        and month == selected.month
                        and day == selected.day
                    )
                    tk.Button(
                        days_frame,
                        text=str(day),
                        command=lambda d=day: choose_date(d),
                        bg=theme["accent"] if is_selected else theme["card_alt"],
                        fg="#FFFFFF" if is_selected else theme["text"],
                        activebackground=theme["accent_hover"],
                        relief="flat",
                        width=3,
                        cursor="hand2",
                    ).grid(row=row, column=column, padx=1, pady=1)

        def change_month(offset):
            month_state["month"] += offset
            if month_state["month"] < 1:
                month_state["month"] = 12
                month_state["year"] -= 1
            elif month_state["month"] > 12:
                month_state["month"] = 1
                month_state["year"] += 1
            render_month()

        def choose_date(day):
            date_var.set(
                f"{month_state['year']:04d}-{month_state['month']:02d}-{day:02d}"
            )
            if on_select:
                on_select()
            popup.destroy()

        render_month()
        popup.focus_set()

    def change_date(self):
        """Switch active planner date."""
        new_date = self.date_var.get().strip()
        try:
            datetime.strptime(new_date, "%Y-%m-%d")
            date_changed = new_date != self.current_date
            self.current_date = new_date
            if date_changed:
                self.reset_tracker_ranges()
            self.refresh_all_views()
        except ValueError:
            messagebox.showerror(
                "Invalid Date", "Please use YYYY-MM-DD format (e.g., 2026-11-24)."
            )

    def select_today(self):
        """Set the planner date to today and refresh its views."""
        self.date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.change_date()

    # --------------------------------------------------------------------------
    # THEME & REFRESH ORCHESTRATION
    # --------------------------------------------------------------------------
    def set_layout_mode(self, mode):
        """Apply and persist the selected desktop or mobile layout."""
        if mode not in ("Desktop", "Mobile") or mode == self.layout_mode:
            return
        self.layout_mode = mode
        self.apply_layout()
        self.refresh_task_trackers()
        self.update_layout_toggle()
        self.save_data()

    def toggle_layout_mode(self):
        """Switch between mobile and desktop layouts."""
        target_mode = "Mobile" if self.layout_mode == "Desktop" else "Desktop"
        self.set_layout_mode(target_mode)

    def update_layout_toggle(self):
        """Show the layout that will be selected by the toggle."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        next_mode = "Mobile" if self.layout_mode == "Desktop" else "Desktop"
        self.layout_toggle_btn.configure(
            text=f"Switch to {next_mode}",
            bg=theme["card_alt"],
            fg=theme["text"],
            activebackground=theme["accent_hover"],
            activeforeground=theme["accent_fg"],
        )

    def toggle_mobile_timeline_panel(self, panel):
        """Expand one mobile timeline panel and collapse the other."""
        if self.layout_mode != "Mobile" or panel not in ("schedule", "inbox"):
            return
        if panel == self.mobile_timeline_panel:
            self.mobile_timeline_panel = (
                "inbox" if panel == "schedule" else "schedule"
            )
        else:
            self.mobile_timeline_panel = panel
        self.update_mobile_timeline_panel()

    def update_mobile_timeline_panel(self):
        """Show the selected mobile panel."""
        for pane in (self.tl_left_pane, self.tl_right_pane):
            pane.pack_forget()

        if self.mobile_timeline_panel == "schedule":
            self.tl_left_pane.pack(side="top", fill="both", expand=True)
        else:
            self.tl_right_pane.pack(side="top", fill="both", expand=True)

        self.update_mobile_switcher_style()

    def update_mobile_switcher_style(self):
        """Reflect which mobile panel is expanded in the switcher buttons."""
        schedule_open = self.mobile_timeline_panel == "schedule"
        inbox_open = self.mobile_timeline_panel == "inbox"
        self.mobile_schedule_btn.configure(
            text="Schedule ▼" if schedule_open else "Schedule ▸"
        )
        self.mobile_inbox_btn.configure(
            text=(
                "Unscheduled / Quick Tasks ▼"
                if inbox_open
                else "Unscheduled / Quick Tasks ▸"
            )
        )
        theme = self.themes["dark" if self.dark_mode else "light"]
        self.mobile_schedule_btn.configure(
            bg=theme["accent"] if schedule_open else theme["card_alt"],
            fg=theme["accent_fg"] if schedule_open else theme["text"],
            activebackground=theme["accent_hover"],
        )
        self.mobile_inbox_btn.configure(
            bg=theme["accent"] if inbox_open else theme["card_alt"],
            fg=theme["accent_fg"] if inbox_open else theme["text"],
            activebackground=theme["accent_hover"],
        )

    def apply_layout(self):
        """Reflow the main window for the selected desktop or mobile view."""
        mobile = self.layout_mode == "Mobile"
        self.root.geometry("390x850" if mobile else "1120x740")
        if mobile:
            self.root.minsize(360, 640)
        else:
            self.root.minsize(900, 620)
        self.header_frame.configure(
            padx=8 if mobile else 20,
            pady=8 if mobile else 12,
        )
        self.notebook.pack_configure(
            padx=5 if mobile else 20,
            pady=(3, 8) if mobile else (4, 16),
        )

        for widget in (
            self.logo_box,
            self.title_box,
            self.controls_frame,
            self.view_controls_frame,
        ):
            widget.pack_forget()
            widget.grid_forget()

        if mobile:
            self.header_frame.grid_columnconfigure(0, weight=0)
            self.header_frame.grid_columnconfigure(1, weight=1)
            self.view_controls_frame.grid(
                row=0, column=0, columnspan=2, sticky="e"
            )
            self.logo_box.grid(row=1, column=0, sticky="w")
            self.title_box.grid(row=1, column=1, sticky="w", padx=8)
            self.controls_frame.grid(
                row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0)
            )
            self.date_controls_frame.pack_forget()
            self.date_controls_frame.pack(side="top", fill="x")
            self.title_lbl.configure(font=("Segoe UI", 13, "bold"))
            self.subtitle_lbl.configure(font=("Segoe UI", 8))
            self.search_entry.configure(width=15)
            for group in (
                self.mgr_search_group,
                self.mgr_cat_group,
                self.mgr_pri_group,
                self.mgr_eng_group,
            ):
                group.pack_forget()
                group.grid_forget()
            self.mgr_filter_bar.grid_columnconfigure(0, weight=1)
            self.mgr_filter_bar.grid_columnconfigure(1, weight=1)
            self.mgr_search_group.grid(
                row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4)
            )
            self.mgr_cat_group.grid(row=1, column=0, sticky="w")
            self.mgr_pri_group.grid(row=1, column=1, sticky="w")
            self.mgr_eng_group.grid(
                row=2, column=0, columnspan=2, sticky="w", pady=(4, 0)
            )
            self.stat_completion_lbl.pack_forget()
            self.stat_deepwork_lbl.pack_forget()
            self.stat_pending_lbl.pack_forget()
            for label in (
                self.stat_completion_lbl,
                self.stat_deepwork_lbl,
                self.stat_pending_lbl,
            ):
                label.pack(side="top", fill="x", pady=2)
            self.stat_completion_lbl.configure(font=("Segoe UI", 10, "bold"))
            self.stat_deepwork_lbl.configure(font=("Segoe UI", 10, "bold"))
            self.stat_pending_lbl.configure(font=("Segoe UI", 10, "bold"))
            for tracker in (self.weekly_tracker_tree, self.monthly_tracker_tree):
                for column, width in (
                    ("date", 72),
                    ("status", 68),
                    ("task", 130),
                    ("time", 90),
                    ("category", 80),
                    ("priority", 65),
                ):
                    tracker.column(column, width=width)
            for column, width in (
                ("date", 72),
                ("status", 68),
                ("task", 130),
                ("time", 90),
                ("category", 80),
                ("priority", 65),
            ):
                self.yearly_tracker_tree.column(column, width=width)
            self.notebook.tab(self.timeline_tab, text="Timeline")
            self.notebook.tab(self.manager_tab, text="Tasks")
            self.notebook.tab(self.review_tab, text="Review")
            self.notebook.tab(self.tracker_tab, text="Tracker")
            self.tracker_notebook.tab(self.weekly_tracker_tab, text="Week")
            self.tracker_notebook.tab(self.monthly_tracker_tab, text="Month")
            self.tracker_notebook.tab(self.yearly_tracker_tab, text="Year")
        else:
            self.view_controls_frame.pack(side="top", fill="x")
            self.logo_box.pack(side="left")
            self.title_box.pack(side="left", padx=12)
            self.controls_frame.pack(side="right")
            self.date_controls_frame.pack_forget()
            self.date_controls_frame.pack(side="left")
            self.title_lbl.configure(font=("Segoe UI", 15, "bold"))
            self.subtitle_lbl.configure(font=("Segoe UI", 9))
            self.search_entry.configure(width=26)
            for group in (
                self.mgr_search_group,
                self.mgr_cat_group,
                self.mgr_pri_group,
                self.mgr_eng_group,
            ):
                group.grid_forget()
            for group in (
                self.mgr_search_group,
                self.mgr_cat_group,
                self.mgr_pri_group,
                self.mgr_eng_group,
            ):
                group.pack(side="left")
            self.stat_completion_lbl.pack_forget()
            self.stat_deepwork_lbl.pack_forget()
            self.stat_pending_lbl.pack_forget()
            self.stat_completion_lbl.pack(
                side="left", fill="x", expand=True, padx=(0, 8)
            )
            self.stat_deepwork_lbl.pack(
                side="left", fill="x", expand=True, padx=8
            )
            self.stat_pending_lbl.pack(
                side="left", fill="x", expand=True, padx=(8, 0)
            )
            self.stat_completion_lbl.configure(font=("Segoe UI", 12, "bold"))
            self.stat_deepwork_lbl.configure(font=("Segoe UI", 12, "bold"))
            self.stat_pending_lbl.configure(font=("Segoe UI", 12, "bold"))
            for tracker in (self.weekly_tracker_tree, self.monthly_tracker_tree):
                for column, width in (
                    ("date", 105),
                    ("status", 95),
                    ("task", 300),
                    ("time", 130),
                    ("category", 100),
                    ("priority", 90),
                ):
                    tracker.column(column, width=width)
            for column, width in (
                ("date", 105),
                ("status", 95),
                ("task", 300),
                ("time", 130),
                ("category", 100),
                ("priority", 90),
            ):
                self.yearly_tracker_tree.column(column, width=width)
            self.notebook.tab(self.timeline_tab, text="  🕒 Daily Timeline  ")
            self.notebook.tab(self.manager_tab, text="  📋 Task & Priority Manager  ")
            self.notebook.tab(self.review_tab, text="  📊 Daily Review & Summary  ")
            self.notebook.tab(self.tracker_tab, text="  📊 Task Trackers  ")
            self.tracker_notebook.tab(self.weekly_tracker_tab, text="Weekly")
            self.tracker_notebook.tab(self.monthly_tracker_tab, text="Monthly")
            self.tracker_notebook.tab(self.yearly_tracker_tab, text="Yearly")

        for widget in (self.tl_left_pane, self.tl_right_pane):
            widget.pack_forget()
        self.mobile_panel_switcher.pack_forget()
        if mobile:
            self.mobile_panel_switcher.pack(
                side="top", fill="x", pady=(0, 5)
            )
            self.update_mobile_timeline_panel()
        else:
            self.tl_left_pane.pack(
                side="left", fill="both", expand=True, padx=(0, 12)
            )
            self.tl_right_pane.pack(side="right", fill="y")

    def toggle_theme(self):
        """Switch between Light and Dark mode."""
        self.dark_mode = not self.dark_mode
        self.save_data()
        self.apply_theme()
        self.refresh_all_views()

    def apply_theme(self):
        """Apply current color palette to static widgets."""
        theme = self.themes["dark" if self.dark_mode else "light"]
        self.root.configure(bg=theme["bg"])
        self.header_frame.configure(bg=theme["card"])
        self.title_box.configure(bg=theme["card"])
        self.title_lbl.configure(bg=theme["card"], fg=theme["text"])
        self.subtitle_lbl.configure(bg=theme["card"], fg=theme["subtext"])
        self.controls_frame.configure(bg=theme["card"])
        self.date_controls_frame.configure(bg=theme["card"])
        self.view_controls_frame.configure(bg=theme["card"])

        self.date_entry.configure(bg=theme["card_alt"], fg=theme["text"])
        self.calendar_btn.configure(bg=theme["card_alt"], fg=theme["text"])
        self.today_btn.configure(bg=theme["card_alt"], fg=theme["text"])
        self.progress_lbl.configure(bg=theme["card_alt"], fg=theme["accent"])
        self.theme_btn.configure(
            text="☀️ Light" if self.dark_mode else "🌙 Dark",
            bg=theme["card_alt"],
            fg=theme["text"],
        )
        self.update_layout_toggle()

        # Tabs & Frames
        for frame in [
            self.timeline_tab,
            self.tl_container,
            self.mobile_panel_switcher,
            self.tl_left_pane,
            self.tl_right_pane,
            self.inbox_canvas_frame,
            self.inbox_list_frame,
            self.backup_frame,
            self.manager_tab,
            self.mgr_container,
            self.mgr_filter_bar,
            self.mgr_search_group,
            self.mgr_cat_group,
            self.mgr_pri_group,
            self.mgr_eng_group,
            self.mgr_canvas_frame,
            self.mgr_list_frame,
            self.review_tab,
            self.tracker_tab,
            self.rev_container,
            self.rev_stats_row,
            self.rev_mood_row,
            self.weekly_tracker_tab,
            self.weekly_tracker_container,
            self.monthly_tracker_tab,
            self.monthly_tracker_container,
            self.yearly_tracker_tab,
            self.yearly_tracker_container,
        ]:
            frame.configure(bg=theme["bg"])

        self.tl_canvas.configure(bg=theme["bg"])
        self.tl_scrollable.configure(bg=theme["bg"])
        self.inbox_canvas.configure(bg=theme["bg"])
        self.mgr_canvas.configure(bg=theme["bg"])
        self.mobile_panel_switcher.configure(bg=theme["bg"])
        self.update_mobile_switcher_style()
        self.tl_header_lbl.configure(bg=theme["bg"], fg=theme["text"])
        self.inbox_header_lbl.configure(bg=theme["bg"], fg=theme["text"])
        self.add_backlog_btn.configure(bg=theme["card_alt"], fg=theme["accent"])
        self.export_excel_btn.configure(bg=theme["card_alt"], fg=theme["text"])
        self.export_btn.configure(bg=theme["card_alt"], fg=theme["text"])
        self.import_btn.configure(bg=theme["card_alt"], fg=theme["text"])

        self.stat_completion_lbl.configure(bg=theme["card"], fg=theme["accent"])
        self.stat_deepwork_lbl.configure(bg=theme["card"], fg="#10B981")
        self.stat_pending_lbl.configure(bg=theme["card"], fg="#F59E0B")

        self.wins_text.configure(bg=theme["card"], fg=theme["text"], insertbackground=theme["text"])
        self.blockers_text.configure(bg=theme["card"], fg=theme["text"], insertbackground=theme["text"])

        self.style.configure("TNotebook", background=theme["bg"], borderwidth=0)
        self.style.configure(
            "Tracker.Treeview",
            background=theme["card"],
            fieldbackground=theme["card"],
            foreground=theme["text"],
            bordercolor=theme["border"],
            rowheight=28,
        )
        self.style.configure(
            "Tracker.Treeview.Heading",
            background=theme["card_alt"],
            foreground=theme["text"],
            font=("Segoe UI", 9, "bold"),
        )
        self.style.map(
            "Tracker.Treeview",
            background=[("selected", theme["accent"])],
            foreground=[("selected", theme["accent_fg"])],
        )
        for label in (
            self.weekly_tracker_title,
            self.monthly_tracker_title,
            self.yearly_tracker_title,
        ):
            label.configure(bg=theme["bg"], fg=theme["text"])
        for label in (
            self.weekly_tracker_summary,
            self.monthly_tracker_summary,
            self.yearly_tracker_summary,
        ):
            label.configure(bg=theme["bg"], fg=theme["subtext"])
        for tree in (
            self.weekly_tracker_tree,
            self.monthly_tracker_tree,
            self.yearly_tracker_tree,
        ):
            tree.tag_configure("completed", foreground=theme["subtext"])
            tree.tag_configure("pending", foreground=theme["text"])
        self.year_tracker_menu.configure(style="TCombobox")
        for range_state in self.tracker_ranges.values():
            range_state["frame"].configure(bg=theme["bg"])
            range_state["preset_frame"].configure(bg=theme["bg"])
            for label_key in ("from_label", "to_label"):
                range_state[label_key].configure(bg=theme["bg"], fg=theme["text"])
            range_state["from_entry"].configure(
                bg=theme["card_alt"], fg=theme["text"]
            )
            range_state["to_entry"].configure(bg=theme["card_alt"], fg=theme["text"])
            range_state["from_calendar_btn"].configure(
                bg=theme["card_alt"], fg=theme["text"]
            )
            range_state["to_calendar_btn"].configure(
                bg=theme["card_alt"], fg=theme["text"]
            )
            range_state["apply_button"].configure(
                bg=theme["accent"], fg=theme["accent_fg"]
            )
            range_state["preset_button"].configure(
                bg=theme["accent"],
                fg=theme["accent_fg"],
                activebackground=theme["accent_hover"],
            )
        self.year_tracker_frame.configure(bg=theme["bg"])
        self.year_tracker_label.configure(bg=theme["bg"], fg=theme["text"])
        self.year_preset_frame.configure(bg=theme["bg"])
        self.year_preset_button.configure(
            bg=theme["accent"],
            fg=theme["accent_fg"],
            activebackground=theme["accent_hover"],
        )
        self.style.configure(
            "TNotebook.Tab",
            background=theme["card_alt"],
            foreground=theme["text"],
            bordercolor=theme["border"],
            lightcolor=theme["border"],
            darkcolor=theme["border"],
            padding=[12, 6],
            font=("Segoe UI", 9, "bold"),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", theme["accent"])],
            foreground=[("selected", "#FFFFFF")],
        )

    def refresh_all_views(self):
        """Update the header progress badge and all planner views."""
        todays_tasks = [t for t in self.tasks if t.get("date") == self.current_date]
        total = len(todays_tasks)
        completed = sum(1 for t in todays_tasks if t.get("completed"))
        pct = int((completed / total) * 100) if total > 0 else 0
        self.progress_lbl.config(text=f"Progress: {completed}/{total} ({pct}%)")

        self.refresh_timeline()
        self.refresh_manager()
        self.refresh_review()
        self.refresh_task_trackers()


if __name__ == "__main__":
    root = tk.Tk()
    app = DayFlowApp(root)
    root.mainloop()