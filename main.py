import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from ics import Calendar, Event
import logging
import pytz

# Logging konfigurieren
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', filename='app.log', filemode='w')

# Zusätzlicher Log-Handler für die Konsole
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
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

        # Spaltennamen anpassen und explizit die Spalte „Dienstort“ aus Spalte H benennen
        raw_df = raw_df.rename(columns=lambda x: x.strip().replace(" ", "_"))  # Entferne überflüssige Leerzeichen und ersetze durch _
        if "Dienstort" not in raw_df.columns:
            raw_df.insert(4, "Dienstort", raw_df.iloc[:, 7])  # Spalte H explizit zu „Dienstort“ umbenennen

        # Relevante Spalten extrahieren
        columns_to_keep = ["Datum", "Uhrzeit", "Wer", "Aktivität", "Dienstort", "Planer_1", "Planer_2", "Planer_3"]
        df = raw_df[columns_to_keep]

        # Zeilen mit ungültigen Datumswerten entfernen
        df = df[pd.to_datetime(df["Datum"], errors="coerce").notna()]
        df = df.reset_index(drop=True)  # Reset der Indizes

        # Konvertiere die Spalte "Uhrzeit" in ein korrektes Zeitformat mit Zeitzonenanpassung
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
        unique_values = df["Wer"].dropna().unique()  # Einzigartige Werte in der Spalte "Wer"

        # Vorherige Checkboxen löschen
        for widget in checkbox_frame.winfo_children():
            widget.destroy()
        checkboxes.clear()

        for i, value in enumerate(unique_values):
            var = tk.IntVar()
            checkbox = tk.Checkbutton(checkbox_frame, text=value, variable=var, command=update_highlight)
            checkbox.grid(row=0, column=i, sticky="w", padx=5)
            checkboxes[value] = var
        logging.info("Checkboxen erfolgreich erstellt.")

        # Vorherige Widgets löschen
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
        scrollbar_y = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=table_canvas.yview)
        scrollbar_x = tk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=table_canvas.xview)
        scrollbar_y.grid(row=2, column=4, sticky="ns")
        scrollbar_x.grid(row=3, column=0, columnspan=4, sticky="ew")

        # Canvas mit Scrollbars verbinden
        table_canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Frame für Tabelle erstellen
        scrollable_frame = tk.Frame(table_canvas)
        table_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Funktion für Scrolling
        def on_frame_configure(event):
            table_canvas.config(scrollregion=table_canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", on_frame_configure)

        # Maus-Scrollen aktivieren
        def on_mouse_wheel(event):
            table_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        table_canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        # Tabelle füllen
        displayed_rows.clear()
        for i, col in enumerate(df.columns):
            tk.Label(scrollable_frame, text=col, borderwidth=1, relief="solid", padx=5, pady=5, bg="lightgray").grid(row=0, column=i, sticky="nsew")
        for row_index, row in df.iterrows():
            displayed_row = []
            for col_index, value in enumerate(row):
                lbl = tk.Label(scrollable_frame, text=value, borderwidth=1, relief="solid", padx=5, pady=5)
                lbl.grid(row=row_index + 1, column=col_index, sticky="nsew")
                displayed_row.append(lbl)
            displayed_rows.append(displayed_row)

        # Dynamische Anpassung der Fenstergröße
        root.update_idletasks()
        table_width = scrollable_frame.winfo_reqwidth() + 50
        table_height = scrollable_frame.winfo_reqheight() + 150
        window_width = max(800, table_width)
        window_height = max(600, table_height)
        root.geometry(f"{window_width}x{window_height}")

        # Dynamische Anpassung des Hauptframes
        main_frame.rowconfigure(2, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Label für ausgewählte Termine aktualisieren
        if not selected_count_label:
            selected_count_label = tk.Label(main_frame, text="Ausgewählte Termine: 0")
            selected_count_label.grid(row=4, column=0, columnspan=4, pady=10)

        update_selected_count()

    except KeyError as e:
        logging.error(f"Fehlende Spalte: {e}")
        messagebox.showerror("Fehler", f"Fehlende Spalte: {e}. Bitte prüfen Sie die Datei.")
    except Exception as e:
        logging.error(f"Ein Fehler ist aufgetreten: {e}")
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {str(e)}")

# Funktion zum Aktualisieren der Markierung und der Anzahl basierend auf Checkboxen
def update_highlight():
    global displayed_rows

    if df is None:
        logging.warning("Tabelle wurde nicht geladen. Highlighting nicht möglich.")
        return

    # Verantwortlichkeiten aus Checkboxen abfragen
    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    logging.info(f"Ausgewählte Verantwortlichkeiten: {selected_options}")

    # Markieren der entsprechenden Zeilen
    for row_index, row in df.iterrows():
        if row["Wer"] in selected_options:
            for lbl in displayed_rows[row_index]:
                lbl.config(bg="lightyellow")
        else:
            for lbl in displayed_rows[row_index]:
                lbl.config(bg="white")

    # Aktualisieren der Anzahl der ausgewählten Termine
    update_selected_count()

# Funktion zum Aktualisieren des Labels für die Anzahl der ausgewählten Termine
def update_selected_count():
    global selected_count_label, df

    if df is None or selected_count_label is None:
        logging.warning("Tabelle oder Label für die Zählung nicht verfügbar.")
        return

    # Verantwortlichkeiten aus Checkboxen abfragen
    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    selected_count = len(df[df["Wer"].isin(selected_options)])

    logging.info(f"Anzahl der ausgewählten Termine: {selected_count}")
    selected_count_label.config(text=f"Ausgewählte Termine: {selected_count}")

# Funktion zum Exportieren der gefilterten Daten
def export_filtered():
    global df

    if df is None:
        logging.warning("Export fehlgeschlagen: Keine Tabelle geladen.")
        messagebox.showerror("Fehler", "Bitte zuerst eine Tabelle importieren.")
        return

    # Verantwortlichkeiten aus Checkboxen abfragen
    selected_options = [key for key, var in checkboxes.items() if var.get() == 1]
    if not selected_options:
        logging.warning("Keine Verantwortlichkeiten ausgewählt.")
        messagebox.showerror("Fehler", "Bitte wählen Sie mindestens eine Verantwortlichkeit aus.")
        return

    # Filterung der Daten
    filtered_df = df[df["Wer"].isin(selected_options)]

    if filtered_df.empty:
        logging.info("Keine Einträge für die ausgewählten Verantwortlichkeiten gefunden.")
        messagebox.showinfo("Info", "Keine Einträge für die ausgewählten Verantwortlichkeiten gefunden.")
        return

    # Kalender erstellen
    calendar = Calendar()
    for _, row in filtered_df.iterrows():
        try:
            event = Event()
            event.name = row["Aktivität"]
            
            # Überprüfung des Datums
            datum = pd.to_datetime(row["Datum"], errors="coerce")
            if pd.isna(datum):
                logging.warning(f"Ungültiges Datum in Zeile: {row}")
                continue  # Überspringe ungültige Zeile
            
            datum = datum.date()  # Konvertiere in ein Datum-Objekt
            
            # Überprüfung der Uhrzeit
            if pd.notna(row["Uhrzeit"]):
                try:
                    uhrzeit = pd.to_datetime(row["Uhrzeit"], format="%H:%M:%S", errors="coerce").time()
                    if uhrzeit is None:
                        raise ValueError("Ungültige Uhrzeit")
                except Exception as e:
                    logging.warning(f"Ungültige Uhrzeit in Zeile: {row}. Fehler: {e}")
                    continue  # Überspringe ungültige Zeile
                
                start_datetime = pd.Timestamp.combine(datum, uhrzeit).replace(tzinfo=None)
                start_datetime = berlin_tz.localize(start_datetime, is_dst=None)
                event.begin = start_datetime.strftime("%Y%m%dT%H%M%S%z")  # Ausgabe mit lokalem Offset
            else:
                # Ganztägige Termine
                all_day_date = pd.Timestamp(datum).replace(tzinfo=None)
                all_day_date = berlin_tz.localize(all_day_date, is_dst=None)
                event.begin = all_day_date.strftime("%Y%m%dT%H%M%S%z")
                event.make_all_day()

            # Event-Details
            event.location = row["Dienstort"] if pd.notna(row["Dienstort"]) else "Unbekannter Ort"
            event.description = f"Planer: {row['Planer_1']}, {row['Planer_2']}, {row['Planer_3']}"
            
            # Event zum Kalender hinzufügen
            calendar.events.add(event)
            logging.info(f"Event hinzugefügt: {event.name}, Beginn: {event.begin}")
        
        except Exception as e:
            logging.error(f"Fehler beim Verarbeiten eines Termins: {e}")

    # Speicherort abfragen
    save_path = filedialog.asksaveasfilename(defaultextension=".ics", filetypes=[("ICS Files", "*.ics")])
    if save_path:
        try:
            with open(save_path, 'w') as f:
                f.writelines(calendar)
            logging.info(f"Kalenderdatei erfolgreich gespeichert unter: {save_path}")
            messagebox.showinfo("Erfolg", f"Kalenderdatei erfolgreich gespeichert unter: {save_path}")
        except Exception as e:
            logging.error(f"Fehler beim Speichern der Datei: {e}")
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Datei: {str(e)}")

# GUI erstellen
root = tk.Tk()
root.title("Kalender Export Tool")

# Hauptframe
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.rowconfigure(2, weight=1)
main_frame.columnconfigure(0, weight=1)

# Buttons
import_button = tk.Button(main_frame, text="Tabelle importieren", command=import_and_display_table)
import_button.grid(row=0, column=0, pady=5, sticky="w")

export_button = tk.Button(main_frame, text="Exportieren", command=export_filtered)
export_button.grid(row=0, column=1, pady=5, sticky="w")

# Checkboxen für Verantwortlichkeiten
checkbox_frame = tk.Frame(main_frame)
checkbox_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=10)

root.geometry("800x600")
root.minsize(600, 400)
root.mainloop()
