import pandas as pd

file_path = r"C:\Users\user\Arbeitsplatz\PythonProjekte\FinanceOverview\Daten\Finanzen_2024.xlsx"
months = ["Januar", "Februar", "März", "April", "Mai", "Juni",
          "Juli", "August", "September", "Oktober", "November", "Dezember"]

dataset = pd.DataFrame()

for month in months:
    df = pd.read_excel(file_path, sheet_name=month, parse_dates=["Datum"])
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")  # Datum als datetime

    new_rows = []

    for _, row in df.iterrows():
        # Spalten außer Datum synchron splitten
        split_cols = {}
        for col in df.columns:
            if col == "Datum":
                split_cols[col] = [row[col]]
            else:
                split_cols[col] = [v.strip() for v in str(row[col]).split(",")]

        # max. Länge aller Spalten
        max_len = max(len(vals) for vals in split_cols.values())

        # Alle Spalten auf max_len auffüllen (letztes Element wiederholen)
        for col, vals in split_cols.items():
            if len(vals) < max_len:
                vals.extend([vals[-1]] * (max_len - len(vals)))

        # Zeilen synchron zusammenbauen
        for i in range(max_len):
            new_rows.append({col: split_cols[col][i] for col in split_cols})

    df_new = pd.DataFrame(new_rows)

    # Betrag wieder in Zahl umwandeln (falls vorhanden)
    if "Betrag (€)" in df_new.columns:
        df_new["Betrag (€)"] = pd.to_numeric(df_new["Betrag (€)"], errors="coerce")

    dataset = pd.concat([dataset, df_new], ignore_index=True)

print(dataset.dtypes)
print(dataset.head(20))
