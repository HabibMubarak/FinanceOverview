import pandas as pd

class ExpenseManger:
    def __init__(self):
        self.filename = ""
        pass

    def load_file(self):
        pass

    def create_file(self, filename):
        if filename.endswith(".xlsx"):
            try:
                # Neue Daten für ein Blatt
                data = [
                    {"Datum": "", "Geschäft": "", "Kategorie": "", "Produkt": "", "Betrag (€)": 0},
                ]

                df_neu = pd.DataFrame(data)
                df_neu.to_excel("Daten/"+filename)
                self.filename = filename

            except FileNotFoundError:
                print("Datei nicht gefunden oder Pfad stimmt nicht.")
            except Exception as e:
                print("Fehler:", e)
        else:
            print("Dateiformat ungültig.")

    def create_new_month(self, month):
        # Neue Daten für ein Blatt
        data = [
            {"Datum": "", "Geschäft": "", "Kategorie": "", "Produkt": "", "Betrag (€)": 0},
        ]

        df_neu = pd.DataFrame(data)
        df_neu.to_excel(self.filename)
        
        
        with pd.ExcelWriter("Daten/"+self.filename, mode="a", engine="openpyxl") as writer:
            df_neu.to_excel(writer, sheet_name=month)

    def create_new_year(self):
        pass