import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from ics import Calendar, Event
import logging
import pytz

# Logging konfigurieren
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filename="app.log",
    filemode="w",
)

# Zusätzlicher Log-Handler für die Konsole
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
console_handler.setFormatter(console_formatter)
logging.getLogger().addHandler(console_handler)

# Globale Variablen
df = None  # Daten aus Excel-Datei
table_canvas = None  # Canvas für die Tabelle
scrollable_frame = None  # Scrollbarer Frame für die Tabelle
scrollbar_y = None  # Vertikale Scrollbar
scrollbar_x = None  # Horizontale Scrollbar
displayed_rows = []  # Liste der aktuell angezeigten Zeilen
checkboxes = {}  # Checkboxen für Verantwortlichkeiten
selected_count_label = None  # Label für die Anzahl der ausgewählten Termine
file_type_var = None  # Variable für Dateityp-Auswahl
event_preview_index = 0  # Index für die Vorschau von Kalender-Events
event_preview_label = None  # Label für die Event-Vorschau

# Zeitzone für Berlin definieren
berlin_tz = pytz.timezone("Europe/Berlin")


# Funktion zum Importieren und Anzeigen der Tabelle
def import_and_display_table():
    global df, table_canvas, scrollable_frame, scrollbar_y, scrollbar_x, displayed_rows, checkboxes, selected_count_label

    # Datei auswählen
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not filepath:
        logging.warning("Keine Datei ausgewählt.")
        messagebox.showerror("Fehler", "Keine Datei ausgewählt.")
        return

    try:
        # Excel-Datei laden und irrelevante Spalte "Wochentag" (Spalte B) ignorieren
        raw_df = pd.read_excel(filepath, skiprows=9)  # Lese ab Zeile 10 (Index 9)
        logging.info("Excel-Datei erfolgreich geladen.")

        # Spaltennamen anpassen
        raw_df = raw_df.rename(
            columns=lambda x: x.strip().replace(" ", "_")
        )  # Entferne überflüssige Leerzeichen
        if "Dienstort" not in raw_df.columns:
            raw_df.insert(
                4, "Dienstort", raw_df.iloc[:, 7]
            )  # Spalte H explizit zu „Dienstort“ umbenennen

        # Relevante Spalten extrahieren
        columns_to_keep = [
            "Datum",
            "Uhrzeit",
            "Wer",
            "Aktivität",
            "Dienstort",
            "Planer_1",
            "Planer_2",
            "Planer_3",
        ]
        df = raw_df[columns_to_keep]

        # Zeilen mit ungültigen Datumswerten entfernen
        df = df[pd.to_datetime(df["Datum"], errors="coerce").notna()]
        df = df.reset_index(drop=True)

        # Konvertiere die Spalte "Uhrzeit" in ein korrektes Zeitformat
        def parse_time(value):
            try:
                datetime_obj = pd.to_datetime(value, format="%H:%M:%S", errors="coerce")
                if pd.notna(datetime_obj):
                    return datetime_obj.time()
                return None
            except Exception as e:
                logging.error(f"Fehler bei der Uhrzeit-Verarbeitung: {e}")
                return None

        df["Uhrzeit"] = df["Uhrzeit"].apply(parse_time)
        logging.info("Ungültige Datumswerte entfernt und Tabelle aufbereitet.")

        # Dynamisch Checkboxen basierend auf der Spalte "Wer" erstellen
        unique_values = df["Wer"].dropna().unique()

        for widget in checkbox_frame.winfo_children():
            widget.destroy()
        checkboxes.clear()

        for i, value in enumerate(unique_values):
            var = tk.IntVar()
            checkbox = tk.Checkbutton(
                checkbox_frame, text=value, variable=var, command=update_highlight
            )
            checkbox.grid(row=0, column=i, sticky="w", padx=5)
            checkboxes[value] = var
        logging.info("Checkboxen erfolgreich erstellt.")

        if table_canvas:
            table_canvas.destroy()
        if scrollbar_y:
            scrollbar_y.destroy()
        if scrollbar_x:
            scrollbar_x.destroy()

        # Canvas erstellen
        table_canvas = tk.Canvas(main_frame)
        table_canvas.grid(row=2, column=0, columnspan=4, sticky="nsew")

        # Scrollbars hinzufügen
        scrollbar_y = tk.Scrollbar(
            main_frame, orient=tk.VERTICAL, command=table_canvas.yview
        )
        scrollbar_x = tk.Scrollbar(
            main_frame, orient=tk.HORIZONTAL, command=table_canvas.xview
        )
        scrollbar_y.grid(row=2, column=4, sticky="ns")
        scrollbar_x.grid(row=3, column=0, columnspan=4, sticky="ew")
        table_canvas.configure(
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )

        scrollable_frame = tk.Frame(table_canvas)
        table_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def on_frame_configure(event):
            table_canvas.config(scrollregion=table_canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", on_frame_configure)

        def on_mouse_wheel(event):
            table_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        table_canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        displayed_rows.clear()
        for i, col in enumerate(df.columns):
            tk.Label(
                scrollable_frame,
                text=col,
                borderwidth=1,
                relief="solid",
                padx=5,
                pady=5,
                bg="lightgray",
            ).grid(row=0, column=i, sticky="nsew")
        for row_index, row in df.iterrows():
            displayed_row = []
            for col_index, value in enumerate(row):
                lbl = tk.Label(
                    scrollable_frame,
                    text=value,
                    borderwidth=1,
                    relief="solid",
                    padx=5,
                    pady=5,
                )
                lbl.grid(row=row_index + 1, column=col_index, sticky="nsew")
                displayed_row.append(lbl)
            displayed_rows.append(displayed_row)

        root.update_idletasks()

        if not selected_count_label:
            selected_count_label = tk.Label(main_frame, text="Ausgewählte Termine: 0")
            selected_count_label.grid(row=4, column=0, columnspan=4, pady=10)

        update_selected_count()

    except KeyError as e:
        logging.error(f"Fehlende Spalte: {e}")
        messagebox.showerror(
            "Fehler", f"Fehlende Spalte: {e}. Bitte prüfen Sie die Datei."
        )
    except Exception as e:
        logging.error(f"Ein Fehler ist aufgetreten: {e}")
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {str(e)}")


# Funktion zum Aktualisieren der Markierung
def update_highlight():
    global displayed_rows
    if df is None:
        logging.warning("Tabelle wurde nicht geladen. Highlighting nicht möglich.")
        return

    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    logging.info(f"Ausgewählte Verantwortlichkeiten: {selected_options}")

    for row_index, row in df.iterrows():
        if row["Wer"] in selected_options:
            for lbl in displayed_rows[row_index]:
                lbl.config(bg="lightyellow")
        else:
            for lbl in displayed_rows[row_index]:
                lbl.config(bg="white")

    update_selected_count()


# Funktion zum Aktualisieren der Anzahl der ausgewählten Termine
def update_selected_count():
    global selected_count_label, df

    if df is None or selected_count_label is None:
        logging.warning("Tabelle oder Label für die Zählung nicht verfügbar.")
        return

    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    selected_count = len(df[df["Wer"].isin(selected_options)])

    logging.info(f"Anzahl der ausgewählten Termine: {selected_count}")
    selected_count_label.config(text=f"Ausgewählte Termine: {selected_count}")


# Funktion zur Überprüfung des Event-Objekts
def validate_event(event):
    errors = []

    if not event.name:
        errors.append("Name des Events fehlt.")

    if not event.begin:
        errors.append("Startzeit fehlt.")

    if not event.location:
        errors.append("Ort des Events fehlt.")

    if errors:
        logging.error(f"Event-Validierungsfehler: {', '.join(errors)}")
    else:
        logging.info("Event erfolgreich validiert.")

    return errors


# Funktion zum Durchblättern der Kalender-Events
def preview_events():
    global event_preview_index, event_preview_label

    if df is None:
        messagebox.showerror("Fehler", "Bitte laden Sie zuerst eine Tabelle.")
        return

    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    filtered_df = df[df["Wer"].isin(selected_options)]

    if filtered_df.empty:
        messagebox.showinfo(
            "Info", "Keine Events für die ausgewählten Verantwortlichkeiten."
        )
        return

    events = []
    for _, row in filtered_df.iterrows():
        event = Event()
        event.name = row["Aktivität"]
        datum = pd.to_datetime(row["Datum"], errors="coerce")

        if pd.notna(datum):
            if pd.notna(row["Uhrzeit"]):
                uhrzeit = pd.to_datetime(
                    row["Uhrzeit"], format="%H:%M:%S", errors="coerce"
                ).time()
                start_datetime = pd.Timestamp.combine(datum.date(), uhrzeit).replace(
                    tzinfo=None
                )
                start_datetime = berlin_tz.localize(start_datetime, is_dst=None)
                event.begin = start_datetime
            else:
                all_day_date = pd.Timestamp(datum).replace(tzinfo=None)
                all_day_date = berlin_tz.localize(all_day_date, is_dst=None)
                event.begin = all_day_date
                event.make_all_day()

        event.location = row["Dienstort"]
        event.description = (
            f"Planer: {row['Planer_1']}, {row['Planer_2']}, {row['Planer_3']}"
        )

        errors = validate_event(event)
        if not errors:
            events.append(event)

    if not events:
        messagebox.showinfo("Info", "Keine gültigen Events gefunden.")
        return

    def update_preview():
        global event_preview_index
        if events:
            event = events[event_preview_index]
            event_details = (
                f"Name: {event.name}\n"
                f"Beginn: {event.begin}\n"
                f"Ort: {event.location}\n"
                f"Beschreibung: {event.description}"
            )
            event_preview_label.config(text=event_details)

    def next_event():
        global event_preview_index
        if events:
            event_preview_index = (event_preview_index + 1) % len(events)
            update_preview()

    def prev_event():
        global event_preview_index
        if events:
            event_preview_index = (event_preview_index - 1) % len(events)
            update_preview()

    # Vorschau-UI erstellen
    preview_window = tk.Toplevel(root)
    preview_window.title("Event-Vorschau")

    event_preview_label = tk.Label(
        preview_window, text="", justify="left", padx=10, pady=10, font=("Arial", 12)
    )
    event_preview_label.pack(fill=tk.BOTH, expand=True)

    btn_prev = tk.Button(preview_window, text="Vorheriges", command=prev_event)
    btn_prev.pack(side=tk.LEFT, padx=10, pady=10)

    btn_next = tk.Button(preview_window, text="Nächstes", command=next_event)
    btn_next.pack(side=tk.RIGHT, padx=10, pady=10)

    update_preview()

# Funktion zum Exportieren der gefilterten Daten
def export_filtered():
    global df

    if df is None:
        logging.warning("Export fehlgeschlagen: Keine Tabelle geladen.")
        messagebox.showerror("Fehler", "Bitte zuerst eine Tabelle importieren.")
        return

    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    if not selected_options:
        logging.warning("Keine Verantwortlichkeiten ausgewählt.")
        messagebox.showerror("Fehler", "Bitte wählen Sie mindestens eine Verantwortlichkeit aus.")
        return

    folder_path = filedialog.askdirectory()
    if not folder_path:
        logging.warning("Kein Speicherort ausgewählt.")
        messagebox.showerror("Fehler", "Bitte wählen Sie einen Speicherort aus.")
        return

    file_type = file_type_var.get()
    calendar = Calendar()

    for option in selected_options:
        filtered_df = df[df["Wer"] == option]
        if filtered_df.empty:
            logging.info(f"Keine Einträge für {option} gefunden.")
            continue

        for _, row in filtered_df.iterrows():
            try:
                event = Event()
                event.name = row["Aktivität"]

                datum = pd.to_datetime(row["Datum"], errors="coerce")
                if pd.isna(datum):
                    logging.warning(f"Ungültiges Datum in Zeile: {row}")
                    continue

                datum = datum.date()

                if pd.notna(row["Uhrzeit"]):
                    try:
                        uhrzeit = pd.to_datetime(row["Uhrzeit"], format="%H:%M:%S", errors="coerce").time()
                        if uhrzeit is None:
                            raise ValueError("Ungültige Uhrzeit")
                    except Exception as e:
                        logging.warning(f"Ungültige Uhrzeit in Zeile: {row}. Fehler: {e}")
                        continue

                    start_datetime = pd.Timestamp.combine(datum, uhrzeit).replace(tzinfo=None)
                    start_datetime = berlin_tz.localize(start_datetime, is_dst=None)
                    event.begin = start_datetime.strftime("%Y%m%dT%H%M%S%z")
                else:
                    event.begin = datum.strftime("%Y%m%d")  # Korrektes Ganztagesformat
                    event.make_all_day()

                event.location = row["Dienstort"] if pd.notna(row["Dienstort"]) else "Unbekannter Ort"
                event.description = f"Planer: {row['Planer_1']}, {row['Planer_2']}, {row['Planer_3']}"

                calendar.events.add(event)
                logging.info(f"Event hinzugefügt: {event.name}, Beginn: {event.begin}")

            except Exception as e:
                logging.error(f"Fehler beim Verarbeiten eines Termins: {e}")

    save_path = f"{folder_path}/Gesamter_Kalender.{file_type}"
    try:
        with open(save_path, 'w') as f:
            f.writelines(calendar)
        logging.info(f"Gemeinsame Kalenderdatei erfolgreich gespeichert unter: {save_path}")
    except Exception as e:
        logging.error(f"Fehler beim Speichern der Datei {save_path}: {e}")
        messagebox.showerror("Fehler", f"Fehler beim Speichern der Datei {save_path}: {str(e)}")

    messagebox.showinfo("Erfolg", "Kalenderdatei erfolgreich erstellt.")


# GUI erstellen
root = tk.Tk()
root.title("Kalender Export Tool")

main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.rowconfigure(2, weight=1)
main_frame.columnconfigure(0, weight=1)

import_button = tk.Button(
    main_frame, text="Tabelle importieren", command=import_and_display_table
)
import_button.grid(row=0, column=0, pady=5, sticky="w")

preview_button = tk.Button(main_frame, text="Vorschau", command=preview_events)
preview_button.grid(row=0, column=1, pady=5, sticky="w")

export_button = tk.Button(main_frame, text="Exportieren", command=export_filtered)
export_button.grid(row=0, column=2, pady=5, sticky="w")

file_type_var = tk.StringVar(value="ics")
file_type_label = tk.Label(main_frame, text="Dateityp:")
file_type_label.grid(row=0, column=3, padx=5, sticky="w")
file_type_menu = tk.OptionMenu(main_frame, file_type_var, "ics", "ical")
file_type_menu.grid(row=0, column=4, padx=5, sticky="w")

checkbox_frame = tk.Frame(main_frame)
checkbox_frame.grid(row=1, column=0, columnspan=5, sticky="w", pady=10)

root.geometry("800x600")
root.minsize(600, 400)
root.mainloop()
