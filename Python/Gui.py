import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime
import os
from ExpenseManager import ExpenseManager 
import shutil

class Gui:
    def __init__(self, root):
        self.root = root
        self.root.title("Finanzverwaltung")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)  # fürs Schließen-Event

        self.file_year = None  # Jahr der geöffneten/erstellten Datei
        self.manager = ExpenseManager()

        # --- Menü ---
        menubar = tk.Menu(root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Neue Datei erstellen", command=self.create_file_dialog)
        filemenu.add_command(label="Datei öffnen", command=self.open_file_dialog)
        menubar.add_cascade(label="Datei", menu=filemenu)
        root.config(menu=menubar)

        # --- Tabs ---
        tab_control = ttk.Notebook(root)
        self.tab_add = ttk.Frame(tab_control)
        self.tab_edit = ttk.Frame(tab_control)
        tab_control.add(self.tab_add, text="Hinzufügen")
        tab_control.add(self.tab_edit, text="Bearbeiten/Löschen")
        tab_control.pack(expand=1, fill="both")

        # --- Tab: Hinzufügen ---
        frame_input = ttk.LabelFrame(self.tab_add, text="Eintrag hinzufügen")
        frame_input.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_input, text="Datum (TT.MM.JJJJ):").grid(row=0, column=0, padx=5, pady=2)
        self.entry_date = ttk.Entry(frame_input)
        self.entry_date.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frame_input, text="Geschäft:").grid(row=1, column=0, padx=5, pady=2)
        self.entry_company = ttk.Entry(frame_input)
        self.entry_company.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(frame_input, text="Kategorie:").grid(row=2, column=0, padx=5, pady=2)
        self.entry_category = ttk.Entry(frame_input)
        self.entry_category.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(frame_input, text="Produkt:").grid(row=3, column=0, padx=5, pady=2)
        self.entry_product = ttk.Entry(frame_input)
        self.entry_product.grid(row=3, column=1, padx=5, pady=2)

        ttk.Label(frame_input, text="Betrag (€):").grid(row=4, column=0, padx=5, pady=2)
        self.entry_amount = ttk.Entry(frame_input)
        self.entry_amount.grid(row=4, column=1, padx=5, pady=2)

        # Buttons in einer Reihe
        self.btn_add = ttk.Button(frame_input, text="Hinzufügen", command=self.add_entry)
        self.btn_add.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

        self.btn_save = ttk.Button(frame_input, text="Speichern", command=self.save_entries)
        self.btn_save.grid(row=5, column=1, padx=5, pady=5, sticky="ew")


        # --- Tab: Bearbeiten/Löschen ---
        frame_edit = ttk.LabelFrame(self.tab_edit, text="Eintrag bearbeiten/löschen")
        frame_edit.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_edit, text="Datum (TT.MM.JJJJ):").grid(row=0, column=0, padx=5, pady=2)
        self.entry_edit_date = ttk.Entry(frame_edit)
        self.entry_edit_date.grid(row=0, column=1, padx=5, pady=2)

        # Treeview für mehrere Einträge
        self.tree = ttk.Treeview(frame_edit, columns=("Geschäft","Kategorie","Produkt","Betrag"), show="headings")
        self.tree.heading("Geschäft", text="Geschäft")
        self.tree.heading("Kategorie", text="Kategorie")
        self.tree.heading("Produkt", text="Produkt")
        self.tree.heading("Betrag", text="Betrag (€)")
        self.tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        # Klick-Event → ausgewählten Eintrag ins Formular laden
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Eingabefelder für Bearbeitung
        ttk.Label(frame_edit, text="Geschäft:").grid(row=2, column=0, padx=5, pady=2)
        self.entry_edit_company = ttk.Entry(frame_edit)
        self.entry_edit_company.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(frame_edit, text="Kategorie:").grid(row=3, column=0, padx=5, pady=2)
        self.entry_edit_category = ttk.Entry(frame_edit)
        self.entry_edit_category.grid(row=3, column=1, padx=5, pady=2)

        ttk.Label(frame_edit, text="Produkt:").grid(row=4, column=0, padx=5, pady=2)
        self.entry_edit_product = ttk.Entry(frame_edit)
        self.entry_edit_product.grid(row=4, column=1, padx=5, pady=2)

        ttk.Label(frame_edit, text="Betrag (€):").grid(row=5, column=0, padx=5, pady=2)
        self.entry_edit_amount = ttk.Entry(frame_edit)
        self.entry_edit_amount.grid(row=5, column=1, padx=5, pady=2)

        btn_load = ttk.Button(frame_edit, text="Einträge laden", command=self.load_entries)
        btn_load.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

        # --- Neue Buttons unten nebeneinander ---
        btn_frame = ttk.Frame(frame_edit)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=10)

        self.btn_edit = ttk.Button(btn_frame, text="Bearbeiten", command=self.edit_entry)
        self.btn_edit.pack(side="left", padx=5)

        self.btn_delete = ttk.Button(btn_frame, text="Löschen", command=self.delete_entry)
        self.btn_delete.pack(side="left", padx=5)

        self.btn_save_edit = ttk.Button(btn_frame, text="Speichern", command=self.save_entries)
        self.btn_save_edit.pack(side="left", padx=5)



    # --- Methoden ---

    def on_closing(self):
        if self.manager.filename != "":
            answer = messagebox.askyesnocancel(
                "Beenden", "Möchten Sie vor dem Beenden speichern?"
            )
            if answer is None:  # Abbrechen
                return
            elif answer:  # Ja → temp Datei in Original kopieren
                try:
                    self.manager.export_all_data(temp=True)  # Temp-Datei wird aktualisiert
                    shutil.copyfile(self.manager.temp_filename, "Daten/" + self.manager.filename)
                    if os.path.exists(self.manager.temp_filename):
                        os.remove(self.manager.temp_filename)
                    messagebox.showinfo("Info", "Alle Änderungen gespeichert.")
                except Exception as e:
                    messagebox.showerror("Fehler beim Speichern", str(e))
            else:  # Nein → temp Datei löschen
                if os.path.exists(self.manager.temp_filename):
                    os.remove(self.manager.temp_filename)
        self.root.destroy()


    def edit_entry(self):
        if self.manager.filename == "":
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei öffnen oder erstellen.")
            return

        date_str = self.entry_edit_date.get()
        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Ungültiges Datum", "Datum ungültig.")
            return

        # Prüfen nur auf Jahr
        if self.file_year and dt.year != self.file_year:
            messagebox.showerror("Falsches Jahr", f"Das Datum muss im Jahr {self.file_year} liegen!")
            return

        # Prüfen, ob ein Eintrag ausgewählt ist
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("Fehler", "Bitte zuerst einen Eintrag auswählen.")
            return

        try:
            import pandas as pd
            month = self.manager.months[dt.month - 1]

            # DataFrame aus Temp-Datei laden
            df = pd.read_excel(self.manager.temp_filename, sheet_name=month, parse_dates=["Datum"])

            # Werte aus Treeview holen
            item = self.tree.item(selected[0])
            values = item["values"]
            betrag = float(values[3])  # String → float

            # Zeile finden
            idx = df.index[
                (df["Datum"].dt.date == dt.date()) &
                (df["Geschäft"] == values[0]) &
                (df["Kategorie"] == values[1]) &
                (df["Produkt"] == values[2]) &
                (df["Betrag (€)"] == betrag)
            ]
            if idx.empty:
                messagebox.showerror("Fehler", "Eintrag nicht gefunden.")
                return

            row_idx = idx[0]

            # Temp-Datei bearbeiten
            df.at[row_idx, "Datum"] = dt
            df.at[row_idx, "Geschäft"] = self.entry_edit_company.get()
            df.at[row_idx, "Kategorie"] = self.entry_edit_category.get()
            df.at[row_idx, "Produkt"] = self.entry_edit_product.get()
            df.at[row_idx, "Betrag (€)"] = float(self.entry_edit_amount.get())

            # In Temp-Datei speichern
            with pd.ExcelWriter(self.manager.temp_filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'

            messagebox.showinfo("Info", f"Eintrag für {date_str} temporär bearbeitet.")
            self.load_entries()  # Treeview neu laden

        except Exception as e:
            messagebox.showerror("Fehler beim Bearbeiten", str(e))





    def create_file_dialog(self):
        year = simpledialog.askinteger("Jahr eingeben", "Bitte Jahr angeben (z.B. 2024):")
        if not year:
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Finanzen_{year}.xlsx",
            filetypes=[("Excel-Dateien", "*.xlsx")],
            initialdir="Daten"
        )

        if filename:
            base = os.path.basename(filename)
            if os.path.exists("Daten/" + base):
                messagebox.showwarning("Warnung", f"Datei {base} existiert bereits!")
                return

            self.manager.create_file(base, year)
            self.file_year = year  # <-- Jahr merken
            messagebox.showinfo("Info", f"Neue Datei erstellt: {base}")


    def open_file_dialog(self):
        filename = filedialog.askopenfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel-Dateien", "*.xlsx")],
                                            initialdir="Daten")
        if filename:
            base = os.path.basename(filename)
            self.manager.filename = base
            # Temp-Datei erstellen oder laden
            self.manager.load_file(base)  # <-- hier Temp-Datei richtig setzen und kopieren

            try:
                # Jahr aus Dateinamen extrahieren (z. B. "Finanzen_2024.xlsx")
                self.file_year = int(base.split("_")[1].split(".")[0])
            except:
                self.file_year = None

            messagebox.showinfo("Info", f"Datei geladen: {self.manager.filename}")



    def add_entry(self):
        if self.manager.filename == "": 
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei öffnen oder erstellen.")
            return
        
        # Prüfen, ob eine Datei für die Sitzung geladen/erstellt wurde
        if not self.manager.temp_filename:
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei öffnen oder erstellen.")
            return

        # --- Datum prüfen ---
        try:
            dt = datetime.strptime(self.entry_date.get(), "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Ungültiges Datum", "Bitte Datum im Format TT.MM.JJJJ eingeben (z.B. 12.03.2024).")
            return

        # --- Jahr prüfen ---
        if self.file_year and dt.year != self.file_year:
            messagebox.showerror("Falsches Jahr", f"Das Datum muss im Jahr {self.file_year} liegen!")
            return

        # --- Betrag prüfen ---
        try:
            amount = float(self.entry_amount.get())
        except ValueError:
            messagebox.showerror("Ungültiger Betrag", "Bitte einen gültigen Betrag eingeben (z.B. 12.50).")
            return

        # --- Eintrag hinzufügen ---
        try:
            self.manager.add(
                self.entry_date.get(),
                self.entry_company.get(),
                self.entry_category.get(),
                self.entry_product.get(),
                amount
            )
            messagebox.showinfo("Info", "Eintrag hinzugefügt.")
        except Exception as e:
            messagebox.showerror("Fehler beim Hinzufügen", str(e))

    def delete_entry(self):
        if self.manager.filename == "":
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei öffnen oder erstellen.")
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("Fehler", "Bitte zuerst einen Eintrag im Treeview auswählen.")
            return

        # Daten aus dem ausgewählten Treeview-Eintrag holen
        item = self.tree.item(selected[0])
        values = item["values"]
        date_str = self.entry_edit_date.get()  # Datum muss im Eingabefeld stehen
        if not date_str:
            messagebox.showerror("Fehler", "Datum fehlt im Eingabefeld.")
            return

        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Fehler", "Datum ungültig.")
            return

        # Jahr prüfen, falls gesetzt
        if self.file_year and dt.year != self.file_year:
            messagebox.showerror("Falsches Jahr", f"Das Datum muss im Jahr {self.file_year} liegen!")
            return

        try:
            import pandas as pd
            betrag = float(values[3])
            sheets_to_edit = [self.manager.months[dt.month - 1], "AlleDaten"]  # Monat + AlleDaten

            for sheet in sheets_to_edit:
                df = pd.read_excel(self.manager.temp_filename, sheet_name=sheet, parse_dates=["Datum"])

                # Zeile finden
                idx = df.index[
                    (df["Datum"].dt.date == dt.date()) &
                    (df["Geschäft"] == values[0]) &
                    (df["Kategorie"] == values[1]) &
                    (df["Produkt"] == values[2]) &
                    (df["Betrag (€)"] == betrag)
                ]

                if not idx.empty:
                    df.at[idx[0], "Geschäft"] = ""
                    df.at[idx[0], "Kategorie"] = ""
                    df.at[idx[0], "Produkt"] = ""
                    df.at[idx[0], "Betrag (€)"] = pd.NA

                    # In Temp-Datei speichern
                    with pd.ExcelWriter(self.manager.temp_filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                        df.to_excel(writer, sheet_name=sheet, index=False)
                        ws = writer.sheets[sheet]
                        for cell in ws['A']:
                            cell.number_format = 'DD.MM.YYYY'

            messagebox.showinfo("Info", f"Eintrag für {date_str} gelöscht.")

            # Treeview aktualisieren und Eingabefelder leeren
            self.load_entries()
            self.entry_edit_company.delete(0, tk.END)
            self.entry_edit_category.delete(0, tk.END)
            self.entry_edit_product.delete(0, tk.END)
            self.entry_edit_amount.delete(0, tk.END)

        except Exception as e:
            messagebox.showerror("Fehler", str(e))



    def save_entries(self):
        if self.manager.filename == "":
            messagebox.showerror("Fehler", "Keine Datei geöffnet.")
            return

        try:
            # Zuerst das "AlleDaten"-Sheet aus der Temp-Datei aktualisieren
            self.manager.export_all_data(temp=True)

            # Temp-Datei in die Originaldatei kopieren
            shutil.copyfile(self.manager.temp_filename, "Daten/" + self.manager.filename)

            messagebox.showinfo("Info", "Alle Änderungen gespeichert.")
        except Exception as e:
            messagebox.showerror("Fehler beim Speichern", str(e))

    def load_entries(self):
        """Lädt alle Einträge für das Datum aus der Temp-Datei ins Treeview, ignoriert leere Zeilen."""
        date_str = self.entry_edit_date.get()
        if not date_str:
            messagebox.showerror("Fehler", "Bitte Datum eingeben.")
            return

        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
            month = self.manager.months[dt.month - 1]

            # --- Temp-Datei einlesen ---
            import pandas as pd
            df = pd.read_excel(self.manager.temp_filename, sheet_name=month, parse_dates=["Datum"])

            # Treeview leeren
            for row in self.tree.get_children():
                self.tree.delete(row)

            # Alle Einträge mit passendem Datum einfügen, nur wenn Betrag nicht NaN oder leer
            idx = df.index[(df["Datum"].dt.date == dt.date()) & (df["Betrag (€)"].notna()) & (df["Betrag (€)"] != "")]
            if idx.empty:
                messagebox.showinfo("Info", "Keine Einträge für dieses Datum gefunden.")
                return

            for i in idx:
                row = df.loc[i]
                self.tree.insert("", "end", values=(row["Geschäft"], row["Kategorie"], row["Produkt"], row["Betrag (€)"]))

        except Exception as e:
            messagebox.showerror("Fehler", str(e))



    def on_tree_select(self, event):
        """Lädt die ausgewählte Zeile ins Formular."""
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")

        self.entry_edit_company.delete(0, tk.END)
        self.entry_edit_company.insert(0, values[0])

        self.entry_edit_category.delete(0, tk.END)
        self.entry_edit_category.insert(0, values[1])

        self.entry_edit_product.delete(0, tk.END)
        self.entry_edit_product.insert(0, values[2])

        self.entry_edit_amount.delete(0, tk.END)
        self.entry_edit_amount.insert(0, values[3])


    def show_month(self):
        if self.manager.filename == "":
            messagebox.showerror("Fehler", "Keine Datei geöffnet.")
            return
        month = self.month_var.get()
        try:
            import pandas as pd
            df = pd.read_excel("Daten/" + self.manager.filename, sheet_name=month)
            self.text_output.delete("1.0", tk.END)
            self.text_output.insert(tk.END, df.to_string(index=False))
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
    
    def save_all(self):
        if self.manager.filename == "":
            messagebox.showerror("Fehler", "Keine Datei geöffnet.")
            return
        try:
            self.manager.export_all_data()
            messagebox.showinfo("Info", "Alle Daten erfolgreich ins Sheet 'AlleDaten' exportiert.")
        except Exception as e:
            messagebox.showerror("Fehler", str(e))