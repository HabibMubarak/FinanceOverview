import pandas as pd
import calendar
from datetime import datetime
from numpy import nan
import shutil
import os

class ExpenseManager:
    def __init__(self):
        self.filename = ""
        self.temp_filename = ""
        self.months = ["Januar", "Februar", "März", "April", "Mai", "Juni", 
                          "Juli", "August", "September", "Oktober", "November", "Dezember"]
        self.data_cache = {}

    def load_file(self, filename):
        self.filename = filename
        base_path = os.path.join(os.getcwd(), "Daten")  # CWD/Daten
        full_path = os.path.join(base_path, filename)  # C:\...\Daten\Finanzen_2023.xlsx
        self.temp_filename = os.path.join(base_path, "temp_" + filename)  # temp-Datei ebenfalls in Daten/
        shutil.copyfile(full_path, self.temp_filename)

    def create_file(self, filename, year=None):
        """
        Erstellt eine Excel-Datei für ein Jahr mit 12 Monaten als einzelne Sheets.
        Datumsspalte wird als echtes Datum gespeichert, Anzeige TT.MM.YYYY.
        """
        if not filename.endswith(".xlsx"):
            print("Dateiformat ungültig.")
            return

        try:
            if year is None:
                year = datetime.now().year

            with pd.ExcelWriter("Daten/" + filename, engine="openpyxl") as writer:
                for i, m in enumerate(self.months, start=1):
                    days_in_month = calendar.monthrange(year,i)[1]
                    dates = pd.date_range(start=f"{year}-{i:02d}-01",
                                          end=f"{year}-{i:02d}-{days_in_month}",
                                          freq="D")
                    
                    df_neu = pd.DataFrame({
                        "Datum": dates,   # echtes Datum
                        "Geschäft": "",
                        "Kategorie": "",
                        "Produkt": "",
                        "Betrag (€)": ""
                    })

                    df_neu.to_excel(writer, sheet_name=m, index=False)

                    ws = writer.sheets[m]
                    for cell in ws['A']:
                        cell.number_format = 'DD.MM.YYYY'  # Anzeige TT.MM.YYYY

            self.filename = filename

            # Temp-Datei direkt anlegen
            base_path = os.path.join(os.getcwd(), "Daten")
            self.temp_filename = os.path.join(base_path, "temp_" + filename)
            shutil.copyfile(os.path.join(base_path, filename), self.temp_filename)

            print(f"Datei {filename} und Temp erstellt.")

        except Exception as e:
            print("Fehler:", e)

    def add(self, date, company, category, product, amount):
        """Fügt einen neuen Eintrag hinzu (session-safe, in temp-Datei)."""

        if not self.temp_filename:
            raise FileNotFoundError("Keine Datei für die Sitzung geöffnet. Bitte zuerst eine Datei erstellen oder laden.")

        try:
            dt_obj = datetime.strptime(date, "%d.%m.%Y")
            month = self.months[dt_obj.month - 1]

            # Temp-Datei einlesen
            df = pd.read_excel(self.temp_filename, sheet_name=month, parse_dates=["Datum"])

            # Spalten explizit als Objekt/String setzen
            for col in ["Geschäft", "Kategorie", "Produkt"]:
                if df[col].dtype != "object":
                    df[col] = df[col].astype(object)

            # Prüfen, ob es eine leere Zeile für dieses Datum gibt
            idx_empty = df.index[
                (df["Datum"].dt.date == dt_obj.date()) &
                (df["Geschäft"].isna() | (df["Geschäft"] == "")) &
                (df["Kategorie"].isna() | (df["Kategorie"] == "")) &
                (df["Produkt"].isna() | (df["Produkt"] == "")) &
                (df["Betrag (€)"].isna() | (df["Betrag (€)"] == ""))
            ]

            if len(idx_empty) > 0:
                # Erste leere Zeile füllen
                row_index = idx_empty[0]
                df.at[row_index, "Geschäft"] = company
                df.at[row_index, "Kategorie"] = category
                df.at[row_index, "Produkt"] = product
                df.at[row_index, "Betrag (€)"] = amount
            else:
                # Neue Zeile unter dem letzten Eintrag für das Datum einfügen
                new_entry = pd.DataFrame([{
                    "Datum": dt_obj,
                    "Geschäft": company,
                    "Kategorie": category,
                    "Produkt": product,
                    "Betrag (€)": amount
                }])
                idx_date = df.index[df["Datum"].dt.date == dt_obj.date()]
                if len(idx_date) > 0:
                    insert_pos = idx_date[-1] + 1
                    df = pd.concat([df.iloc[:insert_pos], new_entry, df.iloc[insert_pos:]]).reset_index(drop=True)
                else:
                    df = pd.concat([df, new_entry], ignore_index=True)

            # In temp-Datei speichern
            with pd.ExcelWriter(self.temp_filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'

            print(f"Eintrag für {date} hinzugefügt.")

        except Exception as e:
            print("Fehler beim Hinzufügen:", e)

    
    def save(self):
        """Speichert alle Monats-Sheets und aktualisiert das 'AlleDaten'-Sheet."""
        try:
            # Alle Monats-Sheets speichern
            with pd.ExcelWriter("Daten/" + self.filename, engine="openpyxl") as writer:
                for month, df in self.data_cache.items():
                    df.to_excel(writer, sheet_name=month, index=False)
            
            # Danach das aggregierte Sheet erstellen
            self.export_all_data()

            print("Alle Daten und 'AlleDaten'-Sheet gespeichert.")
        except Exception as e:
            print("Fehler beim Speichern:", e)


    def discard_changes(self):
        """Verwirft alle Änderungen in der aktuellen Sitzung"""
        self.load_file(self.filename)
        print("Alle Änderungen dieser Sitzung verworfen.")

    
    def export_all_data(self, temp=False):
        """Alle Monats-Sheets zusammenfassen und als Tabelle 'AlleDaten' speichern."""
        try:
            file_path = self.temp_filename if temp else "Daten/" + self.filename
            all_data = pd.DataFrame()

            for month in self.months:
                df = pd.read_excel(file_path, sheet_name=month, parse_dates=["Datum"])
                df["Monat"] = month  # Zusatzspalte, damit man weiß, aus welchem Sheet es kam
                all_data = pd.concat([all_data, df], ignore_index=True)

            with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                all_data.to_excel(writer, sheet_name="AlleDaten", index=False)

            print("Sheet 'AlleDaten' erstellt.")

        except Exception as e:
            print("Fehler beim Exportieren:", e)
    
    def edit(self, date, company=None, category=None, product=None, amount=None):
        """
        Bearbeitet einen bestehenden Eintrag für das angegebene Datum.
        Es werden nur Felder überschrieben, die nicht None sind.

        Args:
            date (str): Datum im Format "TT.MM.JJJJ"
            company (str, optional)
            category (str, optional)
            product (str, optional)
            amount (float, optional)
        """
        try:
            dt_obj = datetime.strptime(date, "%d.%m.%Y")
            month = self.months[dt_obj.month - 1]

            df = pd.read_excel("Daten/" + self.filename, sheet_name=month, parse_dates=["Datum"])
            idx = df.index[df["Datum"].dt.date == dt_obj.date()]

            if len(idx) == 0:
                print(f"Datum {date} nicht gefunden.")
                return

            row_index = idx[0]

            if company is not None:
                df.at[row_index, "Geschäft"] = company
            if category is not None:
                df.at[row_index, "Kategorie"] = category
            if product is not None:
                df.at[row_index, "Produkt"] = product
            if amount is not None:
                df.at[row_index, "Betrag (€)"] = amount

            with pd.ExcelWriter("Daten/" + self.filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'

            print(f"Eintrag für {date} bearbeitet.")

        except Exception as e:
            print("Fehler beim Bearbeiten:", e)


    # private Hilfsmethode
    def __append_cell_value(self, current, new_value):
        if pd.isna(current) or current == "":
            return str(new_value)
        else:
            return str(current) + ", " + str(new_value)


    def print_month(self, month):
        """
        Gibt die Finanzübersicht für einen bestimmten Monat aus.

        Args:
            month (str): Monatsname (z. B. "März"), der dem Tabellennamen 
                        in der Excel-Datei entspricht.

        Returns:
            None
        """
        if self.filename != "":
            df = pd.read_excel("Daten/" + self.filename, sheet_name=month)
            print(df)


    

    def delete(self, date):
        """
        Löscht einen Eintrag für das angegebene Datum in Monatsblatt und 'AlleDaten'.

        Args:
            date (str): Datum im Format "TT.MM.JJJJ", dessen Eintrag gelöscht werden soll.
        """
        # Datum parsen
        date = datetime.strptime(date, "%d.%m.%Y")

        # --- Monatsblatt löschen ---
        month = self.months[date.month - 1]
        df = pd.read_excel("Daten/" + self.filename, sheet_name=month, parse_dates=["Datum"])

        idx = df.index[df["Datum"] == date].tolist()
        print("idx:", idx)
        if idx:
            row_index = idx[0]
            df.at[row_index, "Geschäft"] = ""
            df.at[row_index, "Kategorie"] = ""
            df.at[row_index, "Produkt"] = ""
            df.at[row_index, "Betrag (€)"] = nan

            # Monatsblatt überschreiben
            with pd.ExcelWriter("Daten/" + self.filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'
            print(f"Eintrag für {date.strftime('%d.%m.%Y')} im Monatsblatt gelöscht.")
        else:
            print(f"Datum {date.strftime('%d.%m.%Y')} nicht im Monatsblatt gefunden.")

        # --- 'AlleDaten' anpassen ---
        try:
            df_all = pd.read_excel("Daten/" + self.filename, sheet_name="AlleDaten", parse_dates=["Datum"])
            # Zeilen mit passendem Datum löschen
            df_all = df_all[df_all["Datum"] != date]

            # Sheet überschreiben
            with pd.ExcelWriter("Daten/" + self.filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df_all.to_excel(writer, sheet_name="AlleDaten", index=False)
                ws = writer.sheets["AlleDaten"]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'
            print(f"Eintrag für {date.strftime('%d.%m.%Y')} auch aus 'AlleDaten' entfernt.")
        except Exception as e:
            print("Fehler beim Löschen in 'AlleDaten':", e)
