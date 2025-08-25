import random
from datetime import datetime
from ExpenseManager import ExpenseManager
import calendar

if __name__ == "__main__":
    
    companies = ["Rewe", "Lidl", "Aldi", "DM", "Rossmann"]
    categories = ["Lebensmittel", "Drogerie", "Haushalt", "Bekleidung", "Sonstiges"]
    products = ["Milch", "Brot", "Zahnpasta", "Shampoo", "Socken"]

    # Instanz erstellen
    e = ExpenseManager()
    year = 2024
    filename = f"Finanzen_{year}.xlsx"

    # Excel-Datei für 2024 erstellen
    e.create_file(filename, year)

    # Zufällige Einträge für jeden Monat erzeugen
    for month_index, month_name in enumerate(e.months, start=1):
        # Zufällig 5–10 Einträge pro Monat
        for _ in range(random.randint(5, 7)):
            days = random.randint(1, calendar.monthrange(year, month_index)[1]) # random Bereich von 1 bis Anzahl Tage des Monats
            date_str = f"{days:02d}.{month_index:02d}.{year}"

            company = random.choice(companies)
            category = random.choice(categories)
            product = random.choice(products)
            amount = round(random.uniform(0.5, 50.0), 2)

            e.add(date_str, company, category, product, amount)
            e.add(date_str, company, category, product, amount)
            