import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl.chart.label

# ✅ Étape 1 : Créer le DataFrame (les données pour le burnup)
df = pd.DataFrame({
    "Iteration": [f"Iteration {i}" for i in range(9)],
    "Completed": [0, 6, 18, 31, 45, 57, 70, 84, 97],
    "Total Scope": [100, 100, 100, 100, 100, 100, 120, 120, 120]
})

# ✅ Étape 2 : Créer un classeur Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Burnup Data"

# ✅ Étape 3 : Ajouter les données
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

# ✅ Étape 4 : Créer le graphique
chart = LineChart()
chart.title = "Burnup Chart"
chart.style = 10
chart.y_axis.title = "Story Points"
chart.x_axis.title = "Iteration"

# ✅ Étape 5 : Ajouter les séries et les catégories
data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=10)  # Completed + Total Scope
cats = Reference(ws, min_col=1, min_row=2, max_row=10)  # Iteration labels
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Afficher les valeurs sur les points du graphique
for s in chart.series:
    s.data_labels = openpyxl.chart.label.DataLabelList()
    s.data_labels.show_val = True

# ✅ Étape 6 : Ajouter le graphique à la feuille
ws.add_chart(chart, "E2")

# ✅ Étape 7 : Sauvegarder le fichier
wb.save("Burnup_Chart_up_Graph.xlsx")
