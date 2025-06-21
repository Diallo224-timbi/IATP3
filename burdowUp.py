import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl.chart.label

# === Données combinant Dates + Itérations ===
df = pd.DataFrame({
    "Date": [
        "03/06/2024", "10/06/2024", "17/06/2024", "24/06/2024",
        "01/07/2024", "08/07/2024", "15/07/2024", "22/07/2024", "29/07/2024"
    ],
    "Iteration": [f"Iteration {i}" for i in range(9)],
    "Completed": [0, 6, 18, 31, 45, 57, 70, 84, 97],
    "Total Scope": [100, 100, 100, 100, 100, 100, 120, 120, 120]
})

# Colonne combinée : "Date - Iteration"
df["Date_Iteration"] = df["Date"] + " - " + df["Iteration"]

# === Créer le fichier Excel ===
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Burnup Chart"

# Réorganiser les colonnes pour l'affichage
df_export = df[["Date_Iteration", "Completed", "Total Scope"]]

# Ajouter les données à la feuille
for r in dataframe_to_rows(df_export, index=False, header=True):
    ws.append(r)

# === Créer le graphique Burnup ===
chart = LineChart()
chart.title = "Burnup Chart"
chart.style = 13
chart.y_axis.title = "Story Points"
chart.x_axis.title = "Date - Iteration"

# Ajouter les séries
data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=len(df) + 1)
cats = Reference(ws, min_col=1, min_row=2, max_row=len(df) + 1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Affichage des points
for s in chart.series:
    s.data_labels = openpyxl.chart.label.DataLabelList()
    s.data_labels.show_val = True

# Couleurs : Completed = rouge, Total Scope = noir
chart.series[0].graphicalProperties.line.solidFill = "FF0000"
chart.series[1].graphicalProperties.line.solidFill = "000000"

# Ajouter le graphique à la feuille
ws.add_chart(chart, "E2")

# Sauvegarder le fichier
wb.save("Burnup_Chart_Dates_Iterations.xlsx")
