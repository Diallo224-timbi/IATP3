import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl.chart.label

# Données
df = pd.DataFrame({
    "Iteration": [f"Iteration {i}" for i in range(9)],
    "Completed": [0, 6, 18, 31, 45, 57, 70, 84, 97],
    "Total Scope": [100, 100, 100, 100, 100, 100, 120, 120, 120]
})

# Calcul des points restants
df["Remaining"] = df["Total Scope"] - df["Completed"]

# Calcul de la courbe idéale (droite linéaire descendante)
initial_scope = df["Total Scope"].iloc[0]
ideal_burndown = [initial_scope - i * (initial_scope / (len(df) - 1)) for i in range(len(df))]
df["Ideal"] = ideal_burndown

# Créer le fichier Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Burndown Chart"

# Ajouter les données
for r in dataframe_to_rows(df[["Iteration", "Remaining", "Ideal"]], index=False, header=True):
    ws.append(r)

# Créer le graphique
chart = LineChart()
chart.title = "Burndown Chart (Réel vs Idéal)"
chart.style = 13
chart.y_axis.title = "Remaining Story Points"
chart.x_axis.title = "Iteration"

# Ajouter données
data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=len(df)+1)
cats = Reference(ws, min_col=1, min_row=2, max_row=len(df)+1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Afficher les valeurs sur les points
for s in chart.series:
    s.data_labels = openpyxl.chart.label.DataLabelList()
    s.data_labels.show_val = True

# Couleur rouge pour la courbe idéale
chart.series[1].graphicalProperties.line.solidFill = "FF0000"  # rouge
chart.series[0].graphicalProperties.line.solidFill = "000000"  # noir

# Ajouter le graphique
ws.add_chart(chart, "E2")

# Sauvegarder
wb.save("Burndown_With_Ideal_Line.xlsx")
