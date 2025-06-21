# Burnup Chart Generator avec Python

![Exemple de Burnup Chart](https://via.placeholder.com/800x400/CCCCCC/333333?text=Burnup+Chart+Exemple)

Générez automatiquement des graphiques Burnup pour le suivi de projet en utilisant Python, pandas et Openpyxl. Ce script crée un fichier Excel avec un graphique professionnel montrant l'avancement des travaux et l'évolution du scope.

## Fonctionnalités clés
- 📅 Combinaison automatique des dates et itérations
- 📈 Génération de graphiques Burnup avec mise en forme professionnelle
- 🔴🔵 Personnalisation des couleurs (travail complété vs scope total)
- 🔢 Calcul automatique des échelles et positions
- 💾 Export en un seul fichier Excel prêt à l'emploi

## Structure du code
```python
import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
# ... (code complet fourni)