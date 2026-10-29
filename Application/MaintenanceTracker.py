"""Vehicle Maintenance Tracker
240902-6 - Maintenance Tracker
Nathanael Wiebe
Version 0.0.2 - 10/29/2026

Updated for better UI and functionality

"Vibe" coded using Claude
tkinter + SQLite (standard library only).

Folder layout (all next to this script):
    forms/   pre-made PDFs, e.g. 02_6_Month_Maintenance_Interval_A.pdf
    scans/   created automatically; attached scans are copied here
    maintenance.db
"""
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "maintenance.db"
FORMS_DIR = BASE_DIR / "forms"
SCANS_DIR = BASE_DIR / "scans"

# A form is matched by a token in its file name (case-insensitive). If you
# keep several revisions, the last one alphabetically wins.
FORM_TOKENS = {"3 Month": "3 month", "6 Month": "6 month", "3 Year": "3 year",
               "6 Year": "6 year", "9 Year": "9 year", None: "custom"}
SERVICE_NAMES = {"3 Month": "3 Month Service", "6 Month": "6 Month Service",
                 "3 Year": "3 Year Service", "6 Year": "6 Year Service",
                 "9 Year": "9 Year Service", None: ""}
# Order used by the menus (the main window keeps buttons for the first four only).
MENU_INTERVALS = [(None, "Custom Maintenance"), ("3 Month", "3 Month Service"),
                  ("6 Month", "6 Month Service"), ("3 Year", "3 Year Service"),
                  ("6 Year", "6 Year Service"), ("9 Year", "9 Year Service")]

# Optional: path to SumatraPDF.exe. If set, Print Form opens the print dialog
# directly. Otherwise the PDF opens in your default viewer (then press Ctrl+P).
SUMATRA_PATH = None  # e.g. r"C:\Program Files\SumatraPDF\SumatraPDF.exe"

SCHEMA = """
CREATE TABLE IF NOT EXISTS vehicles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    vin TEXT NOT NULL UNIQUE, year INTEGER, make TEXT, model TEXT,
    trim TEXT, engine TEXT, notes TEXT
);
CREATE TABLE IF NOT EXISTS maintenance (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
    vehicle_desc TEXT, vin TEXT,
    service_date TEXT NOT NULL, mileage INTEGER, service TEXT NOT NULL,
    cost REAL, hours REAL, parts TEXT, notes TEXT, scan_path TEXT
);
"""
MIGRATE_COLS = {"vehicle_desc": "TEXT", "vin": "TEXT", "hours": "REAL",
                "parts": "TEXT", "scan_path": "TEXT"}


def vehicle_desc(v):
    return " ".join(str(v[k]) for k in ("year", "make", "model", "trim") if v[k]) or v["vin"]


def open_file(path):
    """Open a file in the OS default application."""
    if hasattr(os, "startfile"):
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def find_form(kind):
    token = FORM_TOKENS[kind]
    matches = sorted(p for p in FORMS_DIR.glob("*.pdf") if token in p.name.lower())
    return matches[-1] if matches else None


def safe_name(s):
    return re.sub(r"[^\w.-]+", "_", s).strip("_") or "x"


def store_scan(src, vin, service_date, service):
    """Copy a scan into scans/<VIN>/ and return its path relative to the app folder."""
    src = Path(src)
    folder = SCANS_DIR / safe_name(vin)
    folder.mkdir(parents=True, exist_ok=True)
    stem = f"{service_date}_{safe_name(service)}"
    dest, n = folder / f"{stem}{src.suffix.lower()}", 2
    while dest.exists():
        dest = folder / f"{stem}_{n}{src.suffix.lower()}"
        n += 1
    shutil.copy2(src, dest)
    return str(dest.relative_to(BASE_DIR))


class DB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        have = {r["name"] for r in self.conn.execute("PRAGMA table_info(maintenance)")}
        for col, typ in MIGRATE_COLS.items():  # upgrade a database from an older version
            if col not in have:
                self.conn.execute(f"ALTER TABLE maintenance ADD COLUMN {col} {typ}")
        self.conn.commit()

    def vehicles(self):
        return self.conn.execute("SELECT * FROM vehicles ORDER BY year DESC, make, model").fetchall()

    def add_vehicle(self, d):
        cur = self.conn.execute(
            "INSERT INTO vehicles (vin, year, make, model, trim, engine, notes) "
            "VALUES (:vin, :year, :make, :model, :trim, :engine, :notes)", d)
        self.conn.commit()
        return cur.lastrowid

    def update_vehicle(self, vid, d):
        self.conn.execute(
            "UPDATE vehicles SET vin=:vin, year=:year, make=:make, model=:model, "
            "trim=:trim, engine=:engine, notes=:notes WHERE id=:id", {**d, "id": vid})
        self.conn.commit()
        
    def delete_vehicle(self, vid):
        self.conn.execute(
            "DELETE FROM vehicles WHERE id=?", (vid,))  # log entries cascade
        self.conn.commit()

    def log(self, vid):
        return self.conn.execute("SELECT * FROM maintenance WHERE vehicle_id=? "
                                 "ORDER BY service_date DESC, id DESC", (vid,)).fetchall()

    def get_log(self, log_id):
        return self.conn.execute("SELECT * FROM maintenance WHERE id=?", (log_id,)).fetchone()

    def add_log(self, vid, d):
        self.conn.execute(
            "INSERT INTO maintenance (vehicle_id, vehicle_desc, vin, service_date, mileage, "
            "service, cost, hours, parts, notes, scan_path) VALUES "
            "(:vid, :vehicle_desc, :vin, :date, :mileage, :service, :cost, :hours, :parts, "
            ":notes, :scan_path)", {**d, "vid": vid})
        self.conn.commit()

    def update_log(self, log_id, d):
        self.conn.execute(
            "UPDATE maintenance SET service_date=:date, mileage=:mileage, service=:service, "
            "cost=:cost, hours=:hours, parts=:parts, notes=:notes, scan_path=:scan_path "
            "WHERE id=:id", {**d, "id": log_id})
        self.conn.commit()

    def set_scan(self, log_id, path):
        self.conn.execute("UPDATE maintenance SET scan_path=? WHERE id=?", (path, log_id))
        self.conn.commit()

    def delete_log(self, log_id):
        self.conn.execute("DELETE FROM maintenance WHERE id=?", (log_id,))
        self.conn.commit()


SCAN_TYPES = [("Scans", "*.pdf *.jpg *.jpeg *.png *.tif *.tiff"), ("All files", "*.*")]


class VehicleDialog(tk.Toplevel):
    FIELDS = ["VIN", "Year", "Make", "Model", "Trim", "Engine"]

    def __init__(self, parent, vehicle=None):
        super().__init__(parent)
        self.title("Edit Vehicle" if vehicle else "Add Vehicle")
        self.resizable(False, False)
        self.result = None
        self.entries = {}
        for r, name in enumerate(self.FIELDS):
            ttk.Label(self, text=name).grid(row=r, column=0, sticky="w", padx=10, pady=3)
            e = ttk.Entry(self, width=32)
            e.grid(row=r, column=1, padx=10, pady=3)
            if vehicle and vehicle[name.lower()] is not None:
                e.insert(0, vehicle[name.lower()])
            self.entries[name.lower()] = e
        ttk.Label(self, text="Notes").grid(row=6, column=0, sticky="nw", padx=10, pady=3)
        self.notes = tk.Text(self, width=32, height=5, wrap="word")
        self.notes.grid(row=6, column=1, padx=10, pady=3)
        if vehicle and vehicle["notes"]:
            self.notes.insert("1.0", vehicle["notes"])
        btns = ttk.Frame(self)
        btns.grid(row=7, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="Save", command=self.save).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=6)
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.grab_set()
        self.entries["vin"].focus()
        self.wait_window(self)

    def save(self):
        d = {k: e.get().strip() for k, e in self.entries.items()}
        d["notes"] = self.notes.get("1.0", "end").strip()
        if not d["vin"]:
            return messagebox.showerror("Missing VIN", "Please enter a VIN.", parent=self)
        if d["year"]:
            if not d["year"].isdigit():
                return messagebox.showerror("Invalid year", "Year must be a number.", parent=self)
            d["year"] = int(d["year"])
        else:
            d["year"] = None
        self.result = d
        self.destroy()


class MaintenanceDialog(tk.Toplevel):
    """Mirrors the paper form: Vehicle / VIN / Date / Mileage header, plus the
    notes the form asks for (parts & fluids, total cost, time spent)."""
    SERVICES = ["3 Month Service", "6 Month Service", "3 Year Service",
                "6 Year Service", "9 Year Service", "Custom Maintenance"]

    def __init__(self, parent, desc, vin, record=None, service=""):
        super().__init__(parent)
        self.title("Edit Maintenance" if record else "Add Maintenance")
        self.resizable(False, False)
        self.result = None
        self.scan = (record["scan_path"] or "") if record else ""
        fmt = lambda v, f="{}": f.format(v) if v is not None else ""
        self.vars = {
            "date": tk.StringVar(value=record["service_date"] if record else date.today().isoformat()),
            "mileage": tk.StringVar(value=fmt(record["mileage"], "{:,}") if record else ""),
            "service": tk.StringVar(value=record["service"] if record else service),
            "cost": tk.StringVar(value=fmt(record["cost"], "{:.2f}") if record else ""),
            "hours": tk.StringVar(value=fmt(record["hours"]) if record else ""),
        }
        ttk.Label(self, text="Vehicle").grid(row=0, column=0, sticky="w", padx=10, pady=3)
        ttk.Label(self, text=desc, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, sticky="w", padx=10)
        ttk.Label(self, text="VIN").grid(row=1, column=0, sticky="w", padx=10, pady=3)
        ttk.Label(self, text=vin).grid(row=1, column=1, sticky="w", padx=10)
        rows = [("Date (YYYY-MM-DD)", "date"), ("Mileage", "mileage"), ("Service", "service"),
                ("Total cost ($)", "cost"), ("Time spent (hours)", "hours")]
        for r, (text, key) in enumerate(rows, start=2):
            ttk.Label(self, text=text).grid(row=r, column=0, sticky="w", padx=10, pady=3)
            if key == "service":
                w = ttk.Combobox(self, width=34, textvariable=self.vars[key], values=self.SERVICES)
            else:
                w = ttk.Entry(self, width=37, textvariable=self.vars[key])
            w.grid(row=r, column=1, sticky="w", padx=10, pady=3)
        self.parts = self._text("Parts & fluids used\n(part numbers)", 7, 3, record["parts"] if record else "")
        self.notes = self._text("Notes", 8, 4, record["notes"] if record else "")
        ttk.Label(self, text="Scanned form").grid(row=9, column=0, sticky="w", padx=10, pady=3)
        sf = ttk.Frame(self)
        sf.grid(row=9, column=1, sticky="w", padx=10)
        self.scan_lbl = ttk.Label(sf, width=24)
        self.scan_lbl.pack(side="left")
        ttk.Button(sf, text="Browse...", command=self.browse).pack(side="left", padx=2)
        ttk.Button(sf, text="Clear", command=lambda: self.set_scan("")).pack(side="left")
        self.set_scan(self.scan)
        btns = ttk.Frame(self)
        btns.grid(row=10, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="Save", command=self.save).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=6)
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.grab_set()
        self.wait_window(self)

    def _text(self, label, row, height, value):
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="nw", padx=10, pady=3)
        t = tk.Text(self, width=37, height=height, wrap="word")
        t.grid(row=row, column=1, padx=10, pady=3)
        t.insert("1.0", value or "")
        return t

    def set_scan(self, path):
        self.scan = path
        self.scan_lbl.config(text=Path(path).name if path else "(none)")

    def browse(self):
        p = filedialog.askopenfilename(parent=self, title="Select scanned form", filetypes=SCAN_TYPES)
        if p:
            self.set_scan(p)

    def save(self):
        v = {k: var.get().strip() for k, var in self.vars.items()}
        try:
            datetime.strptime(v["date"], "%Y-%m-%d")
            mileage = int(v["mileage"].replace(",", "")) if v["mileage"] else None
            cost = float(v["cost"].replace("$", "").replace(",", "")) if v["cost"] else None
            hours = float(v["hours"]) if v["hours"] else None
        except ValueError:
            return messagebox.showerror(
                "Invalid entry", "Check the date (YYYY-MM-DD), mileage, cost, and hours.", parent=self)
        if not v["service"]:
            return messagebox.showerror("Missing service", "Please enter a service.", parent=self)
        self.result = {"date": v["date"], "mileage": mileage, "service": v["service"],
                       "cost": cost, "hours": hours, "scan_path": self.scan,
                       "parts": self.parts.get("1.0", "end").strip(),
                       "notes": self.notes.get("1.0", "end").strip()}
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Maintenance Tracker")
        self.geometry("820x720")
        self.minsize(680, 540)
        FORMS_DIR.mkdir(exist_ok=True)
        self.db = DB(DB_PATH)
        self.vehicle_rows = []
        self.build_ui()
        self.refresh_vehicles()

    # ---------- UI ----------
    def build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        g = ttk.LabelFrame(self, text="Vehicle")
        g.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        g.columnconfigure(0, weight=1)  # empty spacer column pushes everything to the right
        ttk.Button(g, text="Add Vehicle", command=self.add_vehicle).grid(row=0, column=1, rowspan=2, padx=8, pady=10)
        ttk.Button(g, text="Edit Vehicle", command=self.edit_vehicle).grid(row=0, column=2, rowspan=2, padx=8, pady=10)
        ttk.Label(g, text="Select Vehicle").grid(row=0, column=3, sticky="sw", padx=(16, 0))
        self.combo = ttk.Combobox(g, state="readonly", width=40)
        self.combo.grid(row=1, column=3, padx=(16, 8), pady=(0, 8))
        self.combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_log())

        BUTTONS = [("Custom Maintenance", None), ("3 Month Service", "3 Month"),
                   ("6 Month Service", "6 Month"), ("3 Year Service", "3 Year")]
        for row, (title, handler) in enumerate([("Print Form", self.print_form),
                                                ("Add Maintenance", self.add_maintenance)], start=1):
            g = ttk.LabelFrame(self, text=title)
            g.grid(row=row, column=0, sticky="ew", padx=12, pady=6)
            for col, (label, kind) in enumerate(BUTTONS):
                ttk.Button(g, text=label, width=18, command=lambda k=kind, h=handler: h(k)).grid(
                    row=0, column=col, padx=8, pady=10, ipady=4)

        g = ttk.LabelFrame(self, text="Maintenance Log")
        g.grid(row=3, column=0, sticky="nsew", padx=12, pady=(6, 12))
        g.columnconfigure(0, weight=1)
        g.rowconfigure(0, weight=1)
        cols = (("date", "Date", 100, "w"), ("mileage", "Mileage", 100, "e"),
                ("service", "Service", 220, "w"), ("scan", "Scan", 60, "center"))
        self.tree = ttk.Treeview(g, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for key, text, width, anchor in cols:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor=anchor)
        sb = ttk.Scrollbar(g, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 2))
        sb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=(8, 2))
        ttk.Label(g, text="Double-click to edit  |  Right-click for scan options",
                  foreground="gray").grid(row=1, column=0, sticky="w", padx=8, pady=(0, 6))

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Edit entry...", command=self.edit_entry)
        self.menu.add_command(label="Open scan", command=self.open_scan)
        self.menu.add_command(label="Attach scan...", command=self.attach_scan)
        self.menu.add_separator()
        self.menu.add_command(label="Delete entry", command=self.delete_entry)
        self.tree.bind("<Double-1>", lambda e: self.edit_entry() if self.tree.identify_row(e.y) else None)
        self.tree.bind("<Delete>", lambda e: self.delete_entry())
        self.tree.bind("<Button-3>", self.show_menu)
        self.tree.bind("<Button-2>", self.show_menu)  # macOS
        self.build_menu()

    def build_menu(self):
        bar = tk.Menu(self)

        file_m = tk.Menu(bar, tearoff=0)
        file_m.add_command(label="Open Forms Folder", command=lambda: self.open_folder(FORMS_DIR))
        file_m.add_command(label="Open Scans Folder", command=lambda: self.open_folder(SCANS_DIR))
        file_m.add_separator()
        file_m.add_command(label="Exit", command=self.destroy)
        bar.add_cascade(label="File", menu=file_m)

        veh_m = tk.Menu(bar, tearoff=0)
        veh_m.add_command(label="Add Vehicle...", command=self.add_vehicle)
        veh_m.add_command(label="Edit Vehicle...", command=self.edit_vehicle)
        veh_m.add_separator()
        veh_m.add_command(label="Remove Vehicle...", command=self.remove_vehicle)
        bar.add_cascade(label="Vehicle", menu=veh_m)

        add_m = tk.Menu(bar, tearoff=0)
        print_m = tk.Menu(bar, tearoff=0)
        for kind, label in MENU_INTERVALS:
            add_m.add_command(label=label, command=lambda k=kind: self.add_maintenance(k))
            print_m.add_command(label=label, command=lambda k=kind: self.print_form(k))
        bar.add_cascade(label="Add Maintenance", menu=add_m)
        bar.add_cascade(label="Print Form", menu=print_m)

        entry_m = tk.Menu(bar, tearoff=0)
        for label, fn in [("Edit Entry...", self.edit_entry), ("Open Scan", self.open_scan),
                          ("Attach Scan...", self.attach_scan), (None, None),
                          ("Delete Entry", self.delete_entry)]:
            if label is None:
                entry_m.add_separator()
            else:
                entry_m.add_command(label=label, command=lambda f=fn: self.with_entry(f))
        bar.add_cascade(label="Log Entry", menu=entry_m)

        help_m = tk.Menu(bar, tearoff=0)
        help_m.add_command(label="About", command=lambda: messagebox.showinfo(
            "About", "Maintenance Tracker\n\nForms: " + str(FORMS_DIR) + "\nScans: " + str(SCANS_DIR)))
        bar.add_cascade(label="Help", menu=help_m)
        self.config(menu=bar)

    def with_entry(self, fn):
        if self.selected_id():
            fn()
        else:
            messagebox.showinfo("No entry selected", "Select an entry in the Maintenance Log first.")

    def open_folder(self, folder):
        folder.mkdir(exist_ok=True)
        open_file(folder)

    def show_menu(self, e):
        row = self.tree.identify_row(e.y)
        if row:
            self.tree.selection_set(row)
            self.menu.tk_popup(e.x_root, e.y_root)
            self.menu.grab_release()

    # ---------- Vehicles ----------
    def current_vehicle(self):
        i = self.combo.current()
        return self.vehicle_rows[i] if 0 <= i < len(self.vehicle_rows) else None

    def require_vehicle(self):
        v = self.current_vehicle()
        if not v:
            messagebox.showinfo("No vehicle", "Add or select a vehicle first.")
        return v

    def refresh_vehicles(self, select_id=None):
        cur = self.current_vehicle()
        keep = select_id or (cur["id"] if cur else None)
        self.vehicle_rows = self.db.vehicles()
        labels = [vehicle_desc(r) for r in self.vehicle_rows]
        self.combo["values"] = labels
        idx = next((i for i, r in enumerate(self.vehicle_rows) if r["id"] == keep), 0 if labels else -1)
        if idx >= 0:
            self.combo.current(idx)
        else:
            self.combo.set("")
        self.refresh_log()

    def add_vehicle(self):
        dlg = VehicleDialog(self)
        if dlg.result:
            try:
                self.refresh_vehicles(self.db.add_vehicle(dlg.result))
            except sqlite3.IntegrityError:
                messagebox.showerror("Duplicate VIN", "A vehicle with that VIN already exists.")

    def edit_vehicle(self):
        v = self.require_vehicle()
        if not v:
            return
        dlg = VehicleDialog(self, v)
        if dlg.result:
            try:
                self.db.update_vehicle(v["id"], dlg.result)
                self.refresh_vehicles(v["id"])
            except sqlite3.IntegrityError:
                messagebox.showerror("Duplicate VIN", "A vehicle with that VIN already exists.")

    def remove_vehicle(self):
        v = self.require_vehicle()
        if not v:
            return
        n = len(self.db.log(v["id"]))
        msg = (f"Remove {vehicle_desc(v)} (VIN {v['vin']})?\n\n"
               f"This also deletes its {n} maintenance log entr{'y' if n == 1 else 'ies'}. "
               "Attached scan files stay in the scans folder.")
        if messagebox.askyesno("Remove Vehicle", msg, icon="warning", default="no"):
            self.db.delete_vehicle(v["id"])
            self.refresh_vehicles()

    # ---------- Log ----------
    def selected_id(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def refresh_log(self):
        self.tree.delete(*self.tree.get_children())
        v = self.current_vehicle()
        if not v:
            return
        one = lambda s: (s or "").replace("\n", "; ")
        for r in self.db.log(v["id"]):
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["service_date"],
                f"{r['mileage']:,}" if r["mileage"] is not None else "",
                r["service"],
                "Yes" if r["scan_path"] else ""))

    def finish_scan(self, path, vin, when, service):
        """Copy a newly chosen (absolute) scan into scans/; leave stored paths alone."""
        if path and Path(path).is_absolute():
            try:
                return store_scan(path, vin, when, service)
            except OSError as ex:
                messagebox.showerror("Could not copy scan", str(ex))
                return ""
        return path or ""

    def add_maintenance(self, kind):
        v = self.require_vehicle()
        if not v:
            return
        desc = vehicle_desc(v)
        dlg = MaintenanceDialog(self, desc, v["vin"], service=SERVICE_NAMES[kind])
        if dlg.result:
            d = dlg.result
            d["scan_path"] = self.finish_scan(d["scan_path"], v["vin"], d["date"], d["service"])
            d.update(vehicle_desc=desc, vin=v["vin"])
            self.db.add_log(v["id"], d)
            self.refresh_log()

    def edit_entry(self):
        log_id, v = self.selected_id(), self.current_vehicle()
        if not log_id or not v:
            return
        rec = self.db.get_log(log_id)
        dlg = MaintenanceDialog(self, rec["vehicle_desc"] or vehicle_desc(v), rec["vin"] or v["vin"], rec)
        if dlg.result:
            d = dlg.result
            d["scan_path"] = self.finish_scan(d["scan_path"], rec["vin"] or v["vin"], d["date"], d["service"])
            self.db.update_log(log_id, d)
            self.refresh_log()

    def open_scan(self):
        log_id = self.selected_id()
        rec = self.db.get_log(log_id) if log_id else None
        if not rec or not rec["scan_path"]:
            return messagebox.showinfo("No scan", "This entry has no scan attached yet.")
        path = BASE_DIR / rec["scan_path"]
        if not path.exists():
            return messagebox.showerror("Scan missing", f"Could not find:\n{path}")
        open_file(path)

    def attach_scan(self):
        log_id, v = self.selected_id(), self.current_vehicle()
        if not log_id or not v:
            return
        src = filedialog.askopenfilename(title="Select scanned form", filetypes=SCAN_TYPES)
        if src:
            rec = self.db.get_log(log_id)
            rel = self.finish_scan(src, rec["vin"] or v["vin"], rec["service_date"], rec["service"])
            if rel:
                self.db.set_scan(log_id, rel)
                self.refresh_log()

    def delete_entry(self):
        log_id = self.selected_id()
        if log_id and messagebox.askyesno("Delete entry", "Delete the selected log entry?\n"
                                          "(An attached scan file stays in the scans folder.)"):
            self.db.delete_log(log_id)
            self.refresh_log()

    # ---------- Printing ----------
    def print_form(self, kind):
        if not self.require_vehicle():
            return
        pdf = find_form(kind)
        if not pdf:
            return messagebox.showerror(
                "Form not found",
                f"No PDF with '{FORM_TOKENS[kind]}' in its file name was found in:\n{FORMS_DIR}")
        if SUMATRA_PATH and Path(SUMATRA_PATH).exists():
            subprocess.Popen([SUMATRA_PATH, "-print-dialog", str(pdf)])
        else:
            open_file(pdf)  # opens in your PDF viewer; press Ctrl+P


if __name__ == "__main__":
    App().mainloop()