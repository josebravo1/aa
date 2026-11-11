import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from tkcalendar import Calendar
from excel_manager import ExcelManager

FILE_CONFIG = {
    "Semi 1": {
        "path": r"\\10.234.37.10\T1-Test-Data\PE-Projects\A1-KPIS- PE\7-EVP005\3-OEE\EVP005 SEMI 1\11-NOV\Wk Down time EVP005 SEMI 1 NOV - Copy.xlsx",
        "headers": [
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
            "NDF"
        ],
        "stations": {
            "EOL": "115038",
            "R1": "115019",
            "R2": "115030"
        }
    },
    "Semi 2": {
        "path": r"C:\Users\Jose_Bravo\Desktop\New folder (2)\excel-gui-app\data\semi_2.xlsx",
        "headers": [
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
            "NDF"
        ],
        "stations": {
            "EOL": "118027",
            "R1": "118016",
            "R2": "118022"
        }
    },
    "Manual": {
        "path": r"C:\Users\Jose_Bravo\Desktop\New folder (2)\excel-gui-app\data\manual.xlsx",
        "headers": [
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
            "NDF"
        ],
        "stations": {
            "EOL": "111025",
            "R1": "111016",
            "R2": "111021"
        }
    }
}


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Registro de Downtime")
        self.root.geometry("300x300")
        self.root.configure(bg="#eaf2f8")
        self.root.minsize(160, 160)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TButton", padding=10, font=("Segoe UI", 10, "bold"))
        style.map(
            "TButton",
            background=[("active", "#2563eb")],
            foreground=[("active", "white")]
        )

        title = tk.Label(
            self.root,
            text="Selecciona una Linea",
            font=("Segoe UI", 18, "bold"),
            bg="#eaf2f8",
            fg="#1f2937"
        )
        title.pack(pady=(24, 18))

        buttons_frame = tk.Frame(self.root, bg="#eaf2f8")
        buttons_frame.pack(fill="x", padx=28)

        for button_name in FILE_CONFIG:
            btn = ttk.Button(
                buttons_frame,
                text=button_name,
                width=26,
                command=lambda name=button_name: self.open_form(name)
            )
            btn.pack(pady=8, fill="x")

    def open_form(self, filename):
        config = FILE_CONFIG[filename]
        FormWindow(self.root, filename, config["path"], config["headers"], config["stations"])

    def run(self):
        self.root.mainloop()


class FormWindow:
    def __init__(self, parent, filename, file_path, headers, stations):
        self.filename = filename
        self.file_path = file_path
        self.headers = headers
        self.stations = stations
        self.manager = ExcelManager(file_path)

        self.win = tk.Toplevel(parent)
        self.win.title(f"{filename} - Captura")
        self.win.geometry("800x400")
        self.win.configure(bg="#eef4fb")
        self.win.minsize(600, 320)

        style = ttk.Style(self.win)
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10, "bold"))

        title = tk.Label(
            self.win,
            text=f"Formulario: {filename}",
            font=("Segoe UI", 16, "bold"),
            bg="#eef4fb",
            fg="#111827"
        )
        title.pack(pady=(18, 12))

        form_frame = tk.Frame(self.win, bg="#eef4fb")
        form_frame.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        form_frame.columnconfigure(0, weight=1)
        form_frame.columnconfigure(1, weight=1)

        self.entries = {}

        hour_values = []
        for h in range(24):
            if h == 23:
                hour_values.append("23:00 - 23:54")
            else:
                end_h = (h + 1) % 24
                hour_values.append(f"{h:02d}:00 - {end_h:02d}:00")

        self.hour_values = hour_values

        for idx, header in enumerate(headers):
            row = idx // 2
            col = idx % 2

            cell = tk.Frame(form_frame, bg="#eef4fb")
            cell.grid(row=row, column=col, sticky="ew", padx=8, pady=5)
            cell.columnconfigure(0, weight=0)
            cell.columnconfigure(1, weight=1)

            label = tk.Label(
                cell,
                text=header,
                width=22,
                anchor="w",
                bg="#eef4fb",
                fg="#1f2937",
                font=("Segoe UI", 10)
            )
            label.grid(row=0, column=0, sticky="w", padx=(0, 8))

            if header == "Device":
                entry = tk.Entry(cell, width=30, state="readonly", justify="center")
                entry.grid(row=0, column=1, sticky="ew")
                self.entries[header] = entry

            elif header == "Estación":
                station_var = tk.StringVar()
                combo = ttk.Combobox(cell, textvariable=station_var, width=30, state="readonly")
                combo["values"] = ["EOL", "R1", "R2"]
                combo.grid(row=0, column=1, sticky="ew")
                combo.set("EOL")
                combo.bind("<<ComboboxSelected>>", self.on_station_selected)
                self.entries[header] = combo

            elif header == "Fecha":
                frame_date = tk.Frame(cell, bg="#eef4fb")
                frame_date.grid(row=0, column=1, sticky="ew")

                date_var = tk.StringVar()
                date_entry = tk.Entry(frame_date, textvariable=date_var, width=20, justify="center")
                date_entry.pack(side="left", fill="x", expand=True)
                date_entry.insert(0, datetime.now().strftime("%m-%d"))

                btn_calendar = ttk.Button(
                    frame_date,
                    text="📅",
                    width=3,
                    command=lambda v=date_var: self.open_calendar(v)
                )
                btn_calendar.pack(side="left", padx=(6, 0))

                self.entries[header] = date_entry

            elif header == "Turno":
                combo = ttk.Combobox(cell, values=["1ST", "2ND", "3RD"], width=30, state="readonly")
                combo.grid(row=0, column=1, sticky="ew")
                combo.set(self.get_turno_for_hour(datetime.now().strftime("%H:%M")))
                self.entries[header] = combo

            elif header == "Hora":
                combo = ttk.Combobox(cell, values=self.hour_values, width=30, state="readonly")
                combo.grid(row=0, column=1, sticky="ew")
                current_option = self.get_current_hour_option()
                combo.set(current_option)
                combo.bind("<<ComboboxSelected>>", self.on_hour_selected)
                self.entries[header] = combo

            elif header == "Eventos":
                spin = tk.Spinbox(cell, from_=1, to=999, width=15, increment=1, justify="center")
                spin.grid(row=0, column=1, sticky="ew")
                spin.delete(0, tk.END)
                spin.insert(0, "1")
                self.entries[header] = spin

            elif header == "Problema":
                entry = tk.Entry(cell, width=30)
                entry.grid(row=0, column=1, sticky="ew")
                entry.bind("<KeyRelease>", self.on_problem_change)
                self.entries[header] = entry

            elif header == "Problemas de otros departamentos":
                entry = tk.Entry(cell, width=30)
                entry.grid(row=0, column=1, sticky="ew")
                self.entries[header] = entry

            elif header == "Tiempo muerto (Min)":
                vcmd = (self.win.register(self.validate_numeric), "%P")
                entry = tk.Entry(
                    cell,
                    width=30,
                    state="disabled",
                    validate="key",
                    validatecommand=vcmd,
                    justify="center"
                )
                entry.grid(row=0, column=1, sticky="ew")
                self.entries[header] = entry

            else:
                entry = tk.Entry(cell, width=30)
                entry.grid(row=0, column=1, sticky="ew")
                self.entries[header] = entry

        self.update_device_from_station()
        self.on_problem_change()

        button_container = tk.Frame(self.win, bg="#eef4fb")
        button_container.pack(pady=(18, 22))

        save_btn = ttk.Button(
            button_container,
            text="Guardar y nueva fila",
            command=self.save_data
        )
        save_btn.pack(side="left", padx=8)

        close_btn = ttk.Button(
            button_container,
            text="Cerrar",
            command=self.win.destroy
        )
        close_btn.pack(side="left", padx=8)

    def validate_numeric(self, value):
        return value == "" or value.isdigit()

    def get_current_hour_option(self):
        current_hour = datetime.now().strftime("%H:00")
        for value in self.hour_values:
            if value.startswith(current_hour):
                return value
        return self.hour_values[0]

    def get_turno_for_hour(self, hour_value):
        try:
            current_time = datetime.strptime(hour_value, "%H:%M")
        except ValueError:
            current_time = datetime.strptime(datetime.now().strftime("%H:%M"), "%H:%M")

        hour = current_time.hour
        minute = current_time.minute

        if (6, 0) <= (hour, minute) < (14, 1):
            return "1ST"
        if (14, 1) <= (hour, minute) <= (23, 59):
            return "2ND"
        return "3RD"

    def open_calendar(self, date_var):
        top = tk.Toplevel(self.win)
        top.title("Selecciona una fecha")
        top.geometry("300x280")

        cal = Calendar(
            top,
            selectmode="day",
            date_pattern="mm/dd/yyyy",
            background="#eef4fb",
            foreground="#1f2937",
            headersbackground="#dfeaf5",
            weekendsbackground="#f8fafc",
            normalbackground="#ffffff",
            selectbackground="#2563eb",
            selectforeground="white"
        )
        cal.pack(fill="both", expand=True, padx=10, pady=10)

        def set_date():
            selected = cal.get_date()
            try:
                parsed = datetime.strptime(selected, "%m/%d/%Y")
                date_var.set(parsed.strftime("%m-%d"))
            except ValueError:
                date_var.set(selected)
            top.destroy()

        ttk.Button(top, text="Aceptar", command=set_date).pack(pady=(0, 10))

    def on_station_selected(self, event=None):
        self.update_device_from_station()

    def on_hour_selected(self, event=None):
        if "Turno" in self.entries:
            now = datetime.now()
            hour_value = now.strftime("%H:%M")
            self.entries["Turno"].set(self.get_turno_for_hour(hour_value))

    def update_device_from_station(self):
        station = self.entries["Estación"].get()
        device = self.stations.get(station, "")
        self.entries["Device"].config(state="normal")
        self.entries["Device"].delete(0, tk.END)
        self.entries["Device"].insert(0, device)
        self.entries["Device"].config(state="readonly")

    def on_problem_change(self, event=None):
        problem_widget = self.entries.get("Problema")
        other_widget = self.entries.get("Problemas de otros departamentos")
        time_widget = self.entries.get("Tiempo muerto (Min)")

        if problem_widget is not None:
            has_problem = bool(problem_widget.get().strip())

            if other_widget is not None:
                if has_problem:
                    other_widget.config(state="disabled")
                    other_widget.delete(0, tk.END)
                else:
                    other_widget.config(state="normal")

            if time_widget is not None:
                if has_problem:
                    time_widget.config(state="normal")
                else:
                    time_widget.config(state="disabled")
                    time_widget.delete(0, tk.END)

    def save_data(self):
        values = {}
        for header, widget in self.entries.items():
            if hasattr(widget, "get"):
                values[header] = widget.get()

        if not any(values.values()):
            messagebox.showwarning("Datos vacíos", "Debes llenar al menos un campo.")
            return

        try:
            self.manager.append_row(self.headers, values)
            messagebox.showinfo("Guardado", f"Se guardó una nueva fila en {self.filename}")
            self.clear_form()
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def clear_form(self):
        for header, widget in self.entries.items():
            if header == "Estación":
                widget.set("EOL")
            elif header == "Device":
                widget.config(state="normal")
                widget.delete(0, tk.END)
                widget.config(state="readonly")
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
            elif header == "Problema":
                widget.delete(0, tk.END)
            elif header == "Problemas de otros departamentos":
                widget.config(state="normal")
                widget.delete(0, tk.END)
            elif header == "Tiempo muerto (Min)":
                widget.config(state="disabled")
                widget.delete(0, tk.END)
            else:
                if hasattr(widget, "delete"):
                    widget.delete(0, tk.END)

        self.update_device_from_station()
        self.on_problem_change()