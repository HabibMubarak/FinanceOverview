import pandas as pd
import calendar
from datetime import datetime
from numpy import nan


class ExpenseManager:
    def __init__(self):
        self.filename = ""
        self.months = ["Januar", "Februar", "März", "April", "Mai", "Juni", 
                          "Juli", "August", "September", "Oktober", "November", "Dezember"]

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
            print(f"Datei {filename} erstellt.")

        except Exception as e:
            print("Fehler:", e)

    def add(self, date, company, category, product, amount):
        """
        Fügt einen neuen Eintrag hinzu. Datum intern als datetime, Excel zeigt TT.MM.YYYY.
        """
        try:
            dt_obj = datetime.strptime(date, "%d.%m.%Y")
            month = self.months[dt_obj.month - 1]

            df = pd.read_excel("Daten/" + self.filename, sheet_name=month, parse_dates=["Datum"])

            df["Datum"] = pd.to_datetime(df["Datum"], errors='coerce')

            # Spalten als string behandeln
            for col in ["Geschäft", "Kategorie", "Produkt"]:
                if col in df.columns:
                    df[col] = df[col].astype("string")

            idx = df.index[df["Datum"] == dt_obj].tolist()
            if not idx:
                print("Datum nicht gefunden:", date)
                return

            row_index = idx[0]
            df.at[row_index, "Geschäft"] = self.__append_cell_value(df.at[row_index, "Geschäft"], company)
            df.at[row_index, "Kategorie"] = self.__append_cell_value(df.at[row_index, "Kategorie"], category)
            df.at[row_index, "Produkt"] = self.__append_cell_value(df.at[row_index, "Produkt"], product)
            df["Betrag (€)"] = df["Betrag (€)"].astype("string")
            df.at[row_index, "Betrag (€)"] = self.__append_cell_value(df.at[row_index, "Betrag (€)"], amount)

            with pd.ExcelWriter("Daten/" + self.filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'

            print(f"Eintrag für {date} hinzugefügt/aktualisiert.")

        except Exception as e:
            print("Fehler beim Hinzufügen:", e)

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
        Löscht einen Eintrag für das angegebene Datum, indem alle Felder geleert werden.

        Args:
            date (str): Datum im Format "TT.MM.JJJJ", dessen Eintrag gelöscht werden soll.

        Returns:
            None
        """

        # Datum parsen
        date = datetime.strptime(date, "%d.%m.%Y")

        # Sheet-Namen bestimmen
        month = self.months[date.month - 1]
        df = pd.read_excel("Daten/" + self.filename, sheet_name=month, parse_dates=["Datum"])

        # Datum suchen (Vergleich mit String!)
        idx = df.index[df["Datum"] == date].tolist()
        if not idx:
            print("Datum nicht gefunden:", date)
            return

        row_index = idx[0]
        df.at[row_index, "Geschäft"] = ""
        df.at[row_index, "Kategorie"] = ""
        df.at[row_index, "Produkt"] = ""
        df.at[row_index, "Betrag (€)"] = nan

        # Sheet überschreiben
        with pd.ExcelWriter("Daten/" + self.filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=month, index=False)
                ws = writer.sheets[month]
                for cell in ws['A']:
                    cell.number_format = 'DD.MM.YYYY'

        print(f"Eintrag für {date} gelöscht.")


"""
# Test
e = ExpenseManager()
year = 2024
e.create_file(f"Finanzen_{year}.xlsx", year)


e.add("15.03.2024", "Rewe", "Lebensmittel", "Milch", 1.29)
e.add("17.03.2024", "Rewe", "Lebensmittel", "Milch", 1.45)
e.delete("17.03.2024")

"""