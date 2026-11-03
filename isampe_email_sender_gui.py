import os
import sys
import shlex
import queue
import threading
import subprocess

from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

SENDER_SCRIPT = BASE_DIR / "bulk_email_sender.py"
RECONCILE_SCRIPT = BASE_DIR / "reconcile_emails.py"

INPUT_DIR = BASE_DIR / "input"
LOG_DIR = BASE_DIR / "logs"

REPORT_FILE = BASE_DIR / "email_reconciliation.xlsx"

APP_TITLE = "ISAMPE EMAIL SENDER"

BG = "#101827"
SIDEBAR = "#0B1220"
CARD = "#1B293B"
ACCENT = "#16A6A1"
TEXT = "#F1F5F9"
MUTED = "#94A3B8"
GREEN = "#22C55E"
RED = "#EF4444"


# ============================================================
# APPLICATION
# ============================================================

class ISAMPEEmailApp:

    def __init__(self, root):

        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1150x760")
        self.root.minsize(950, 650)
        self.root.configure(bg=BG)

        self.process = None
        self.log_queue = queue.Queue()

        self.sender_log = None

        self.sender_args = tk.StringVar(value="--send")
        self.log_path = tk.StringVar(value="")

        self.status_text = tk.StringVar(
            value="Ready"
        )

        self.setup_styles()
        self.build_interface()

        self.refresh_dashboard()
        self.root.after(100, self.process_queue)

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

    # ========================================================
    # STYLES
    # ========================================================

    def setup_styles(self):

        style = ttk.Style()

        style.theme_use("clam")

        style.configure(
            ".",
            font=("Segoe UI", 10)
        )

        style.configure(
            "TNotebook",
            background=BG,
            borderwidth=0
        )

        style.configure(
            "TNotebook.Tab",
            background=CARD,
            foreground=TEXT,
            padding=(18, 12),
            borderwidth=0
        )

        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", ACCENT)
            ],
            foreground=[
                ("selected", "#FFFFFF")
            ]
        )

        style.configure(
            "Treeview",
            background=CARD,
            foreground=TEXT,
            fieldbackground=CARD,
            rowheight=28,
            borderwidth=0
        )

        style.configure(
            "Treeview.Heading",
            background=SIDEBAR,
            foreground=TEXT,
            relief="flat",
            padding=8
        )

        style.map(
            "Treeview",
            background=[
                ("selected", ACCENT)
            ]
        )

        style.configure(
            "TProgressbar",
            troughcolor=SIDEBAR,
            background=ACCENT,
            thickness=12
        )

    # ========================================================
    # UI HELPERS
    # ========================================================

    def label(
        self,
        parent,
        text,
        size=10,
        color=TEXT,
        bold=False
    ):

        return tk.Label(
            parent,
            text=text,
            bg=parent.cget("bg"),
            fg=color,
            font=(
                "Segoe UI",
                size,
                "bold" if bold else "normal"
            )
        )

    def button(
        self,
        parent,
        text,
        command,
        color=ACCENT
    ):

        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg="white",
            activebackground=color,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=10,
            cursor="hand2",
            font=("Segoe UI", 10, "bold")
        )

    def card(self, parent):

        return tk.Frame(
            parent,
            bg=CARD,
            padx=20,
            pady=18
        )

    # ========================================================
    # MAIN INTERFACE
    # ========================================================

    def build_interface(self):

        main = tk.Frame(
            self.root,
            bg=BG
        )

        main.pack(
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # SIDEBAR
        # ----------------------------------------------------

        sidebar = tk.Frame(
            main,
            bg=SIDEBAR,
            width=225
        )

        sidebar.pack(
            side="left",
            fill="y"
        )

        sidebar.pack_propagate(False)

        tk.Label(
            sidebar,
            text="ISAMPE",
            bg=SIDEBAR,
            fg=ACCENT,
            font=("Segoe UI", 23, "bold")
        ).pack(
            anchor="w",
            padx=22,
            pady=(30, 0)
        )

        tk.Label(
            sidebar,
            text="EMAIL MANAGEMENT",
            bg=SIDEBAR,
            fg=MUTED,
            font=("Segoe UI", 9, "bold")
        ).pack(
            anchor="w",
            padx=24,
            pady=(2, 35)
        )

        menu_items = [
            "Dashboard",
            "Send Emails",
            "Reconciliation",
            "Activity Log",
            "Reports"
        ]

        for item in menu_items:

            self.button(
                sidebar,
                item,
                lambda x=item: self.select_tab(x),
                color=SIDEBAR
            ).pack(
                fill="x",
                padx=12,
                pady=4
            )

        tk.Frame(
            sidebar,
            bg="#263449",
            height=1
        ).pack(
            fill="x",
            padx=20,
            pady=25
        )

        self.label(
            sidebar,
            "CAMPAIGN STATUS",
            9,
            MUTED,
            True
        ).pack(
            anchor="w",
            padx=22
        )

        self.label(
            sidebar,
            "●  Application ready",
            10,
            GREEN
        ).pack(
            anchor="w",
            padx=22,
            pady=10
        )

        # ----------------------------------------------------
        # CONTENT AREA
        # ----------------------------------------------------

        content = tk.Frame(
            main,
            bg=BG
        )

        content.pack(
            side="left",
            fill="both",
            expand=True
        )

        header = tk.Frame(
            content,
            bg=BG,
            height=85
        )

        header.pack(
            fill="x",
            padx=28,
            pady=(20, 5)
        )

        tk.Label(
            header,
            text=APP_TITLE,
            bg=BG,
            fg=TEXT,
            font=("Segoe UI", 22, "bold")
        ).pack(
            anchor="w"
        )

        tk.Label(
            header,
            text="Bulk email campaigns and recipient reconciliation",
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 10)
        ).pack(
            anchor="w"
        )

        self.notebook = ttk.Notebook(content)

        self.notebook.pack(
            fill="both",
            expand=True,
            padx=25,
            pady=15
        )

        self.tabs = {}

        for name in [
            "Dashboard",
            "Send Emails",
            "Reconciliation",
            "Activity Log",
            "Reports"
        ]:

            frame = tk.Frame(
                self.notebook,
                bg=BG
            )

            self.notebook.add(
                frame,
                text=name
            )

            self.tabs[name] = frame

        self.build_dashboard()
        self.build_sender_tab()
        self.build_reconciliation_tab()
        self.build_activity_tab()
        self.build_reports_tab()

        # ----------------------------------------------------
        # STATUS BAR
        # ----------------------------------------------------

        footer = tk.Frame(
            content,
            bg=SIDEBAR,
            height=35
        )

        footer.pack(
            fill="x",
            side="bottom"
        )

        tk.Label(
            footer,
            textvariable=self.status_text,
            bg=SIDEBAR,
            fg=MUTED
        ).pack(
            anchor="w",
            padx=15,
            pady=7
        )

    # ========================================================
    # TAB NAVIGATION
    # ========================================================

    def select_tab(self, name):

        for index in range(self.notebook.index("end")):

            if self.notebook.tab(index, "text") == name:

                self.notebook.select(index)
                break

    # ========================================================
    # DASHBOARD
    # ========================================================

    def build_dashboard(self):

        frame = self.tabs["Dashboard"]

        self.label(
            frame,
            "Campaign Overview",
            18,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25,
            pady=20
        )

        cards = tk.Frame(
            frame,
            bg=BG
        )

        cards.pack(
            fill="x",
            padx=15
        )

        self.dashboard_values = {}

        metrics = [
            "Excel Files",
            "Total Records",
            "Unique Recipients",
            "Last Report"
        ]

        for metric in metrics:

            card = self.card(cards)

            card.pack(
                side="left",
                fill="both",
                expand=True,
                padx=8
            )

            self.label(
                card,
                metric,
                10,
                MUTED
            ).pack(
                anchor="w"
            )

            value = self.label(
                card,
                "--",
                23,
                TEXT,
                True
            )

            value.pack(
                anchor="w",
                pady=(10, 0)
            )

            self.dashboard_values[metric] = value

        # Quick actions.

        actions = self.card(frame)

        actions.pack(
            fill="x",
            padx=23,
            pady=25
        )

        self.label(
            actions,
            "Quick Actions",
            14,
            TEXT,
            True
        ).pack(
            anchor="w",
            pady=(0, 15)
        )

        self.button(
            actions,
            "Open Send Emails",
            lambda: self.select_tab("Send Emails")
        ).pack(
            side="left",
            padx=(0, 10)
        )

        self.button(
            actions,
            "Open Reconciliation",
            lambda: self.select_tab("Reconciliation")
        ).pack(
            side="left"
        )

        self.label(
            frame,
            "Project Directory",
            11,
            MUTED,
            True
        ).pack(
            anchor="w",
            padx=25
        )

        self.label(
            frame,
            str(BASE_DIR),
            10,
            TEXT
        ).pack(
            anchor="w",
            padx=25,
            pady=8
        )

    def refresh_dashboard(self):

        files = list(INPUT_DIR.glob("*.xlsx")) if INPUT_DIR.exists() else []

        total_records = 0
        unique_emails = set()

        try:

            import pandas as pd

            for file in files:

                df = pd.read_excel(
                    file,
                    dtype=object
                )

                total_records += len(df)

                if "Email" in df.columns:

                    for value in df["Email"].dropna():

                        matches = re.findall(
                            r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
                            str(value)
                        )

                        unique_emails.update(
                            email.casefold()
                            for email in matches
                        )

        except Exception:
            pass

        self.dashboard_values["Excel Files"].config(
            text=str(len(files))
        )

        self.dashboard_values["Total Records"].config(
            text=str(total_records)
        )

        self.dashboard_values["Unique Recipients"].config(
            text=str(len(unique_emails))
        )

        self.dashboard_values["Last Report"].config(
            text="Available" if REPORT_FILE.exists() else "None"
        )

    # ========================================================
    # SEND EMAILS TAB
    # ========================================================

    def build_sender_tab(self):

        frame = self.tabs["Send Emails"]

        self.label(
            frame,
            "Bulk Email Campaign",
            18,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25,
            pady=20
        )

        info = self.card(frame)

        info.pack(
            fill="x",
            padx=25,
            pady=10
        )

        self.label(
            info,
            "Input Excel Files",
            12,
            TEXT,
            True
        ).pack(
            anchor="w"
        )

        self.file_list = tk.Listbox(
            info,
            bg=SIDEBAR,
            fg=TEXT,
            selectbackground=ACCENT,
            height=6,
            relief="flat",
            font=("Segoe UI", 10)
        )

        self.file_list.pack(
            fill="x",
            pady=12
        )

        self.refresh_file_list()

        self.button(
            info,
            "Refresh Files",
            self.refresh_file_list
        ).pack(
            anchor="w"
        )

        # Sending configuration.

        config = self.card(frame)

        config.pack(
            fill="x",
            padx=25,
            pady=10
        )

        self.label(
            config,
            "Sender Script Arguments",
            11,
            TEXT,
            True
        ).pack(
            anchor="w"
        )

        tk.Entry(
            config,
            textvariable=self.sender_args,
            bg=SIDEBAR,
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            font=("Segoe UI", 11)
        ).pack(
            fill="x",
            pady=12,
            ipady=8
        )

        self.label(
            config,
            "Default: --send",
            9,
            MUTED
        ).pack(
            anchor="w"
        )

        # Start button.

        self.send_button = self.button(
            frame,
            "START EMAIL CAMPAIGN",
            self.start_sending,
            color=GREEN
        )

        self.send_button.pack(
            anchor="w",
            padx=25,
            pady=20
        )

        self.label(
            frame,
            "Sending Output",
            11,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25
        )

        self.progress = ttk.Progressbar(
            frame,
            mode="indeterminate"
        )

        self.progress.pack(
            fill="x",
            padx=25,
            pady=12
        )

    def refresh_file_list(self):

        self.file_list.delete(0, tk.END)

        if not INPUT_DIR.exists():

            self.file_list.insert(
                tk.END,
                "Input directory not found"
            )

            return

        files = sorted(
            INPUT_DIR.glob("*.xlsx")
        )

        for file in files:

            self.file_list.insert(
                tk.END,
                f"  {file.name}"
            )

        self.refresh_dashboard()

    # ========================================================
    # RECONCILIATION TAB
    # ========================================================

    def build_reconciliation_tab(self):

        frame = self.tabs["Reconciliation"]

        self.label(
            frame,
            "Email Reconciliation",
            18,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25,
            pady=20
        )

        info = self.card(frame)

        info.pack(
            fill="x",
            padx=25,
            pady=10
        )

        self.label(
            info,
            "Select Terminal Log",
            12,
            TEXT,
            True
        ).pack(
            anchor="w"
        )

        row = tk.Frame(
            info,
            bg=CARD
        )

        row.pack(
            fill="x",
            pady=15
        )

        tk.Entry(
            row,
            textvariable=self.log_path,
            bg=SIDEBAR,
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat"
        ).pack(
            side="left",
            fill="x",
            expand=True,
            ipady=10
        )

        self.button(
            row,
            "Browse",
            self.select_log
        ).pack(
            side="left",
            padx=10
        )

        self.label(
            info,
            "The selected log will be compared against all Excel files in input/.",
            10,
            MUTED
        ).pack(
            anchor="w"
        )

        self.reconcile_button = self.button(
            frame,
            "GENERATE RECONCILIATION REPORT",
            self.start_reconciliation
        )

        self.reconcile_button.pack(
            anchor="w",
            padx=25,
            pady=20
        )

        self.label(
            frame,
            "Generated report: email_reconciliation.xlsx",
            10,
            MUTED
        ).pack(
            anchor="w",
            padx=25
        )

    def select_log(self):

        path = filedialog.askopenfilename(

            title="Select Terminal Log",

            initialdir=LOG_DIR if LOG_DIR.exists() else BASE_DIR,

            filetypes=[
                ("Text Logs", "*.txt *.log"),
                ("All Files", "*.*")
            ]

        )

        if path:

            self.log_path.set(path)

    # ========================================================
    # ACTIVITY LOG TAB
    # ========================================================

    def build_activity_tab(self):

        frame = self.tabs["Activity Log"]

        self.label(
            frame,
            "Live Application Output",
            18,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25,
            pady=20
        )

        container = tk.Frame(
            frame,
            bg=SIDEBAR
        )

        container.pack(
            fill="both",
            expand=True,
            padx=25,
            pady=10
        )

        self.output = tk.Text(
            container,
            bg=SIDEBAR,
            fg="#D1FAE5",
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 10),
            wrap="word"
        )

        self.output.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        self.output.insert(
            tk.END,
            "ISAMPE EMAIL SENDER\n"
            "Application initialized.\n"
            "Waiting for an operation...\n\n"
        )

        self.output.config(
            state="disabled"
        )

        self.button(
            frame,
            "Clear Output",
            self.clear_output
        ).pack(
            anchor="w",
            padx=25,
            pady=10
        )

    def write_output(self, message):

        self.output.config(
            state="normal"
        )

        self.output.insert(
            tk.END,
            message + "\n"
        )

        self.output.see(
            tk.END
        )

        self.output.config(
            state="disabled"
        )

    def clear_output(self):

        self.output.config(
            state="normal"
        )

        self.output.delete(
            "1.0",
            tk.END
        )

        self.output.config(
            state="disabled"
        )

    # ========================================================
    # REPORTS TAB
    # ========================================================

    def build_reports_tab(self):

        frame = self.tabs["Reports"]

        self.label(
            frame,
            "Generated Reports",
            18,
            TEXT,
            True
        ).pack(
            anchor="w",
            padx=25,
            pady=20
        )

        reports = [

            (
                "Email Reconciliation",
                REPORT_FILE
            ),

            (
                "Sender Summary",
                BASE_DIR / "email_summary.xlsx"
            ),

            (
                "Failure Report",
                BASE_DIR / "failed_emails.xlsx"
            ),

            (
                "Sender Progress",
                BASE_DIR / "email_progress.json"
            )

        ]

        for name, path in reports:

            card = self.card(frame)

            card.pack(
                fill="x",
                padx=25,
                pady=7
            )

            self.label(
                card,
                name,
                12,
                TEXT,
                True
            ).pack(
                anchor="w"
            )

            self.label(
                card,
                path.name,
                10,
                MUTED
            ).pack(
                anchor="w",
                pady=5
            )

            self.button(
                card,
                "Open",
                lambda p=path: self.open_file(p)
            ).pack(
                anchor="e"
            )

    def open_file(self, path):

        if not path.exists():

            messagebox.showwarning(
                "File Not Found",
                f"File does not exist:\n{path}"
            )

            return

        try:

            if sys.platform.startswith("linux"):

                subprocess.Popen(
                    ["xdg-open", str(path)]
                )

            elif sys.platform == "win32":

                os.startfile(str(path))

            else:

                subprocess.Popen(
                    ["open", str(path)]
                )

        except Exception as e:

            messagebox.showerror(
                "Unable to Open File",
                str(e)
            )

    # ========================================================
    # BACKGROUND PROCESS MANAGEMENT
    # ========================================================

    def run_process(self, command, operation, log_file=None):

        try:

            self.log_queue.put(
                f"\n[{operation}] Starting...\n"
            )

            with subprocess.Popen(

                command,

                cwd=BASE_DIR,

                stdout=subprocess.PIPE,

                stderr=subprocess.STDOUT,

                text=True,

                bufsize=1,

                env=os.environ.copy()

            ) as process:

                self.process = process

                for line in process.stdout:

                    line = line.rstrip()

                    self.log_queue.put(line)

                    if log_file:

                        with open(
                            log_file,
                            "a",
                            encoding="utf-8"
                        ) as f:

                            f.write(line + "\n")

                return_code = process.wait()

            if return_code == 0:

                self.log_queue.put(
                    f"\n[{operation}] Completed successfully."
                )

            else:

                self.log_queue.put(
                    f"\n[{operation}] Process exited with code {return_code}."
                )

        except Exception as e:

            self.log_queue.put(
                f"\nERROR: {e}"
            )

        finally:

            self.process = None

            self.log_queue.put(
                "__PROCESS_FINISHED__"
            )

    def launch_background(self, command, operation, log_file=None):

        if self.process is not None:

            messagebox.showwarning(
                "Process Running",
                "Another operation is already running."
            )

            return

        self.status_text.set(
            f"{operation} running..."
        )

        self.progress.start(12) if hasattr(self, "progress") else None

        thread = threading.Thread(

            target=self.run_process,

            args=(
                command,
                operation,
                log_file
            ),

            daemon=True

        )

        thread.start()

    # ========================================================
    # START SENDING
    # ========================================================

    def start_sending(self):

        if not SENDER_SCRIPT.exists():

            messagebox.showerror(
                "Sender Script Missing",
                f"Cannot find:\n{SENDER_SCRIPT}"
            )

            return

        if not INPUT_DIR.exists():

            messagebox.showerror(
                "Input Directory Missing",
                str(INPUT_DIR)
            )

            return

        excel_files = list(
            INPUT_DIR.glob("*.xlsx")
        )

        if not excel_files:

            messagebox.showwarning(
                "No Excel Files",
                "Place recipient Excel files inside input/."
            )

            return

        if not any(
            os.getenv(f"EMAIL_{i}")
            and os.getenv(f"PASSWORD_{i}")
            for i in range(1, 5)
        ):

            messagebox.showwarning(
                "Sender Credentials Not Found",
                "No sender credentials were found in the GUI environment.\n\n"
                "Launch the GUI from a terminal after loading your "
                "environment variables."
            )

            return

        confirmed = messagebox.askyesno(

            "Confirm Email Campaign",

            f"Start the email campaign?\n\n"
            f"Excel files: {len(excel_files)}\n"
            f"Sender script: {SENDER_SCRIPT.name}\n"
            f"Arguments: {self.sender_args.get()}\n\n"
            f"This may send real emails."

        )

        if not confirmed:
            return

        LOG_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        self.sender_log = (
            LOG_DIR /
            f"gui_sender_{timestamp}.log"
        )

        args = shlex.split(
            self.sender_args.get()
        )

        command = [
            sys.executable,
            str(SENDER_SCRIPT),
            *args
        ]

        self.select_tab("Activity Log")

        self.launch_background(
            command,
            "EMAIL SENDING",
            self.sender_log
        )

        self.status_text.set(
            "Email campaign running..."
        )

    # ========================================================
    # START RECONCILIATION
    # ========================================================

    def start_reconciliation(self):

        if not RECONCILE_SCRIPT.exists():

            messagebox.showerror(
                "Script Missing",
                f"Cannot find:\n{RECONCILE_SCRIPT}"
            )

            return

        log_file = Path(
            self.log_path.get()
        )

        if not log_file.is_file():

            messagebox.showwarning(
                "Select a Log",
                "Select a valid terminal log first."
            )

            return

        if not INPUT_DIR.exists():

            messagebox.showerror(
                "Input Directory Missing",
                str(INPUT_DIR)
            )

            return

        command = [

            sys.executable,

            str(RECONCILE_SCRIPT),

            "--log",

            str(log_file)

        ]

        confirmed = messagebox.askyesno(

            "Generate Report",

            "Run reconciliation using:\n\n"
            f"{log_file.name}\n\n"
            "All Excel files in input/ will be checked."

        )

        if not confirmed:
            return

        self.select_tab("Activity Log")

        self.launch_background(
            command,
            "EMAIL RECONCILIATION"
        )

    # ========================================================
    # PROCESS OUTPUT QUEUE
    # ========================================================

    def process_queue(self):

        try:

            while True:

                message = self.log_queue.get_nowait()

                if message == "__PROCESS_FINISHED__":

                    self.progress.stop() if hasattr(self, "progress") else None

                    self.status_text.set(
                        "Ready"
                    )

                    self.refresh_dashboard()

                    continue

                self.write_output(message)

        except queue.Empty:
            pass

        self.root.after(
            100,
            self.process_queue
        )

    # ========================================================
    # APPLICATION CLOSE
    # ========================================================

    def on_close(self):

        if self.process is not None:

            confirmed = messagebox.askyesno(

                "Process Running",

                "An operation is still running.\n\n"
                "Closing this window may interrupt it.\n"
                "Do you want to close anyway?"

            )

            if not confirmed:
                return

        self.root.destroy()


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    import re

    root = tk.Tk()

    app = ISAMPEEmailApp(root)

    root.mainloop()