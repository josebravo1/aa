import json
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from tkcalendar import Calendar
from excel_manager import ExcelManager


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
HEADERS = [
    "Device",
    "Estación",
    "Fecha",
    "Turno",
    "Hora",
    "Eventos",
    "Problema",
    "Solución",
    "Problemas de otros departamentos",
    "Tiempo muerto (Min)",
    "Piezas falladas",
    "NDF",
]

FILE_CONFIG = {
    "Semi 1": {
        "path": "",
        "headers": HEADERS,
        "stations": {"EOL": "115038", "R1": "115019", "R2": "115030"},
    },
    "Semi 2": {
        "path": "",
        "headers": HEADERS,
        "stations": {"EOL": "118027", "R1": "118016", "R2": "118022"},
    },
    "Manual": {
        "path": "",
        "headers": HEADERS,
        "stations": {"EOL": "111025", "R1": "111016", "R2": "111021"},
    },
    "22PCBA": {
        "path": "",
        "headers": HEADERS,
        "stations": {
            "FW": ["415008", "415009", "415010", "415011"],
            "FCT": ["415012", "415013"],
        },
    },
    "22TLA": {
        "path": "",
        "headers": HEADERS,
        "stations": {
            "FW": ["415008", "415009", "415010", "415011"],
            "FCT": ["415012", "415013"],
        },
    },
}


def load_configured_paths():
    """Read Excel file destinations from config.json beside the app/executable."""
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            config = json.load(file)
    except FileNotFoundError as exc:
        raise RuntimeError(
            f"No se encontró config.json:\n{CONFIG_PATH}\n\n"
            "Coloca config.json junto a main.exe (o junto a app.py al ejecutar desde Python)."
        ) from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"config.json no contiene JSON válido:\n{exc}") from exc

    files = config.get("files")
    if not isinstance(files, dict):
        raise RuntimeError('config.json debe tener un objeto "files".')

    for line_name in FILE_CONFIG:
        entry = files.get(line_name)
        if not isinstance(entry, dict) or not isinstance(entry.get("path"), str):
            raise RuntimeError(f'Falta "files.{line_name}.path" en config.json.')
        path = entry["path"].strip()
        if path:
            resolved_path = Path(path)
            if not resolved_path.is_absolute():
                resolved_path = APP_DIR / resolved_path
            FILE_CONFIG[line_name]["path"] = str(resolved_path)


class App:
    def __init__(self):
        self.root = tk.Tk()
        try:
            load_configured_paths()
        except RuntimeError as error:
            messagebox.showerror("Error de configuración", str(error), parent=self.root)
            self.root.destroy()
            return

        self.root.title("Registro de Downtime")
        self.root.geometry("320x430")
        self.root.configure(bg="#eaf2f8")
        self.root.minsize(220, 300)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TButton", padding=10, font=("Segoe UI", 10, "bold"))
        style.map("TButton", background=[("active", "#2563eb")], foreground=[("active", "white")])

        tk.Label(
            self.root,
            text="Selecciona una Linea",
            font=("Segoe UI", 18, "bold"),
            bg="#eaf2f8",
            fg="#1f2937",
        ).pack(pady=(24, 18))

        buttons_frame = tk.Frame(self.root, bg="#eaf2f8")
        buttons_frame.pack(fill="x", padx=28)
        for name in FILE_CONFIG:
            ttk.Button(
                buttons_frame,
                text=name,
                width=26,
                command=lambda selected=name: self.open_form(selected),
            ).pack(pady=5, fill="x")

    def open_form(self, line_name):
        config = FILE_CONFIG[line_name]
        if not config["path"]:
            messagebox.showwarning(
                "Falta configurar el Excel",
                f'Agrega la ruta del Excel de "{line_name}" en config.json.',
                parent=self.root,
            )
            return
        FormWindow(self.root, line_name, config["path"], config["headers"], config["stations"])

    def run(self):
        self.root.mainloop()


class FormWindow:
    def __init__(self, parent, filename, file_path, headers, stations):
        self.filename = filename
        self.headers = headers
        self.stations = stations
        self.manager = ExcelManager(file_path)

        self.win = tk.Toplevel(parent)
        self.win.title(f"{filename} - Captura")
        self.win.geometry("800x400")
        self.win.configure(bg="#eef4fb")
        self.win.minsize(600, 320)
        ttk.Style(self.win).configure("TButton", font=("Segoe UI", 10, "bold"))

        tk.Label(
            self.win,
            text=f"Formulario: {filename}",
            font=("Segoe UI", 16, "bold"),
            bg="#eef4fb",
            fg="#111827",
        ).pack(pady=(18, 12))

        form_frame = tk.Frame(self.win, bg="#eef4fb")
        form_frame.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        form_frame.columnconfigure(0, weight=1)
        form_frame.columnconfigure(1, weight=1)
        self.entries = {}

        self.hour_values = [
            "23:00 - 23:54" if hour == 23 else f"{hour:02d}:00 - {(hour + 1) % 24:02d}:00"
            for hour in range(24)
        ]

        for index, header in enumerate(headers):
            row, col = divmod(index, 2)
            cell = tk.Frame(form_frame, bg="#eef4fb")
            cell.grid(row=row, column=col, sticky="ew", padx=8, pady=5)
            cell.columnconfigure(1, weight=1)
            tk.Label(
                cell,
                text=header,
                width=22,
                anchor="w",
                bg="#eef4fb",
                fg="#1f2937",
                font=("Segoe UI", 10),
            ).grid(row=0, column=0, sticky="w", padx=(0, 8))

            if header == "Device":
                widget = ttk.Combobox(cell, width=30, state="readonly", justify="center")
                widget.grid(row=0, column=1, sticky="ew")
                self.entries[header] = widget
            elif header == "Estación":
                widget = ttk.Combobox(cell, values=list(stations), width=30, state="readonly")
                widget.grid(row=0, column=1, sticky="ew")
                widget.set(next(iter(stations)))
                widget.bind("<<ComboboxSelected>>", self.on_station_selected)
                self.entries[header] = widget
            elif header == "Fecha":
                date_frame = tk.Frame(cell, bg="#eef4fb")
                date_frame.grid(row=0, column=1, sticky="ew")
                date_var = tk.StringVar(value=datetime.now().strftime("%m-%d"))
                date_entry = tk.Entry(date_frame, textvariable=date_var, width=20, justify="center")
                date_entry.pack(side="left", fill="x", expand=True)
                ttk.Button(
                    date_frame,
                    text="📅",
                    width=3,
                    command=lambda value=date_var: self.open_calendar(value),
                ).pack(side="left", padx=(6, 0))
                self.entries[header] = date_entry
            elif header == "Turno":
                widget = ttk.Combobox(cell, values=["1ST", "2ND", "3RD"], width=30, state="readonly")
                widget.grid(row=0, column=1, sticky="ew")
                widget.set(self.get_turno_for_hour(datetime.now().strftime("%H:%M")))
                self.entries[header] = widget
            elif header == "Hora":
                widget = ttk.Combobox(cell, values=self.hour_values, width=30, state="readonly")
                widget.grid(row=0, column=1, sticky="ew")
                widget.set(self.get_current_hour_option())
                widget.bind("<<ComboboxSelected>>", self.on_hour_selected)
                self.entries[header] = widget
            elif header == "Eventos":
                widget = tk.Spinbox(cell, from_=1, to=999, width=15, increment=1, justify="center")
                widget.grid(row=0, column=1, sticky="ew")
                widget.delete(0, tk.END)
                widget.insert(0, "1")
                self.entries[header] = widget
            elif header == "Problema":
                widget = tk.Entry(cell, width=30)
                widget.grid(row=0, column=1, sticky="ew")
                widget.bind("<KeyRelease>", self.on_problem_change)
                self.entries[header] = widget
            elif header == "Problemas de otros departamentos":
                widget = tk.Entry(cell, width=30)
                widget.grid(row=0, column=1, sticky="ew")
                self.entries[header] = widget
            elif header == "Tiempo muerto (Min)":
                validate = (self.win.register(self.validate_numeric), "%P")
                widget = tk.Entry(
                    cell,
                    width=30,
                    state="disabled",
                    validate="key",
                    validatecommand=validate,
                    justify="center",
                )
                widget.grid(row=0, column=1, sticky="ew")
                self.entries[header] = widget
            else:
                widget = tk.Entry(cell, width=30)
                widget.grid(row=0, column=1, sticky="ew")
                self.entries[header] = widget

        self.update_device_from_station()
        self.on_problem_change()

        buttons = tk.Frame(self.win, bg="#eef4fb")
        buttons.pack(pady=(18, 22))
        ttk.Button(buttons, text="Guardar y nueva fila", command=self.save_data).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cerrar", command=self.win.destroy).pack(side="left", padx=8)

    @staticmethod
    def validate_numeric(value):
        return value == "" or value.isdigit()

    def get_current_hour_option(self):
        current_hour = datetime.now().strftime("%H:00")
        return next((value for value in self.hour_values if value.startswith(current_hour)), self.hour_values[0])

    @staticmethod
    def get_turno_for_hour(hour_value):
        try:
            hour, minute = map(int, hour_value.split(":")[:2])
        except (ValueError, AttributeError):
            hour, minute = datetime.now().hour, datetime.now().minute
        if (6, 0) <= (hour, minute) < (14, 1):
            return "1ST"
        if (14, 1) <= (hour, minute) <= (23, 59):
            return "2ND"
        return "3RD"

    def open_calendar(self, date_var):
        top = tk.Toplevel(self.win)
        top.title("Selecciona una fecha")
        top.geometry("300x280")
        calendar = Calendar(
            top,
            selectmode="day",
            date_pattern="mm/dd/yyyy",
            background="#eef4fb",
            foreground="#1f2937",
            headersbackground="#dfeaf5",
            weekendsbackground="#f8fafc",
            normalbackground="#ffffff",
            selectbackground="#2563eb",
            selectforeground="white",
        )
        calendar.pack(fill="both", expand=True, padx=10, pady=10)

        def set_date():
            selected = calendar.get_date()
            try:
                date_var.set(datetime.strptime(selected, "%m/%d/%Y").strftime("%m-%d"))
            except ValueError:
                date_var.set(selected)
            top.destroy()

        ttk.Button(top, text="Aceptar", command=set_date).pack(pady=(0, 10))

    def on_station_selected(self, event=None):
        self.update_device_from_station()

    def on_hour_selected(self, event=None):
        selected = self.entries["Hora"].get()
        if selected:
            self.entries["Turno"].set(self.get_turno_for_hour(selected.split(" - ")[0]))

    def update_device_from_station(self):
        values = self.stations.get(self.entries["Estación"].get(), [])
        options = [values] if isinstance(values, str) else list(values)
        device = self.entries["Device"]
        device["values"] = options
        device.set(options[0] if options else "")

    def on_problem_change(self, event=None):
        problem = self.entries["Problema"].get().strip()
        other = self.entries["Problemas de otros departamentos"]
        downtime = self.entries["Tiempo muerto (Min)"]
        if problem:
            other.config(state="disabled")
            other.delete(0, tk.END)
            downtime.config(state="normal")
        else:
            other.config(state="normal")
            downtime.config(state="disabled")
            downtime.delete(0, tk.END)

    def save_data(self):
        values = {header: widget.get() for header, widget in self.entries.items()}
        if not any(values.values()):
            messagebox.showwarning("Datos vacíos", "Debes llenar al menos un campo.", parent=self.win)
            return
        try:
            self.manager.append_row(self.headers, values)
            messagebox.showinfo("Guardado", f"Se guardó una nueva fila en {self.filename}", parent=self.win)
            self.clear_form()
        except Exception as error:
            messagebox.showerror("Error al guardar", str(error), parent=self.win)

    def clear_form(self):
        station = next(iter(self.stations))
        for header, widget in self.entries.items():
            if header == "Estación":
                widget.set(station)
            elif header == "Device":
                continue
            elif header == "Fecha":
                widget.delete(0, tk.END)
                widget.insert(0, datetime.now().strftime("%m-%d"))
            elif header == "Turno":
                widget.set(self.get_turno_for_hour(datetime.now().strftime("%H:%M")))
            elif header == "Hora":
                widget.set(self.get_current_hour_option())
            elif header == "Eventos":
                widget.delete(0, tk.END)
                widget.insert(0, "1")
            elif header == "Problemas de otros departamentos":
                widget.config(state="normal")
                widget.delete(0, tk.END)
            elif header == "Tiempo muerto (Min)":
                widget.config(state="disabled")
                widget.delete(0, tk.END)
            elif hasattr(widget, "delete"):
                widget.delete(0, tk.END)
        self.update_device_from_station()
        self.on_problem_change()
