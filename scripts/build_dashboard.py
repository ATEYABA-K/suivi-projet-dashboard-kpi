"""Construit un classeur Excel de suivi de projet (jalons, budget, avancement),
sur le projet fictif déjà cadré dans note-cadrage-scoring-ia-pme, suivi jusqu'à S14."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

jalons = [
    # nom, semaine_debut, semaine_fin_prevue, semaine_fin_reelle (None si pas fini), statut, budget_prevu, budget_reel
    ("Cadrage validé",        0, 0,  0,    "Terminé",   2000,  2000),
    ("Données consolidées",   1, 3,  3,    "Terminé",   8000,  9200),
    ("Premier modèle",        4, 6,  6,    "Terminé",   10000, 9800),
    ("Validation métier",     7, 9,  None, "En retard", 6000,  4500),
    ("Intégration CRM",       10, 12, None, "À venir",   12000, 0),
    ("Bilan à 3 mois",        24, 24, None, "À venir",   2000,  0),
]

SEMAINE_ACTUELLE = 14

wb = openpyxl.Workbook()

# --- Feuille 1 : Suivi ---
ws = wb.active
ws.title = "Suivi_Jalons"
headers = ["Jalon", "Semaine début (prévu)", "Semaine fin (prévu)", "Semaine fin (réel)",
           "Statut", "Budget prévu (€)", "Budget réel (€)", "Écart budget (€)"]
ws.append(headers)
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4C72B0")
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

statut_couleur = {
    "Terminé": "C6E0B4",
    "En retard": "F8CBAD",
    "À venir": "D9D9D9",
    "En cours": "FFE699",
}

for row in jalons:
    nom, sd, sf, sr, statut, bp, br = row
    ecart = br - bp if br else 0
    ws.append([nom, sd, sf, sr if sr is not None else "—", statut, bp, br, ecart])
    r = ws.max_row
    ws.cell(r, 5).fill = PatternFill("solid", fgColor=statut_couleur[statut])

for col in range(1, 9):
    ws.column_dimensions[get_column_letter(col)].width = 20 if col == 1 else 15

# ligne de total
total_row = ws.max_row + 2
ws.cell(total_row, 1, "TOTAL").font = Font(bold=True)
ws.cell(total_row, 6, f"=SUM(F2:F{ws.max_row - 1})").font = Font(bold=True)
ws.cell(total_row, 7, f"=SUM(G2:G{ws.max_row - 1})").font = Font(bold=True)
ws.cell(total_row, 8, f"=SUM(H2:H{ws.max_row - 1})").font = Font(bold=True)

# --- Feuille 2 : Dashboard ---
ws2 = wb.create_sheet("Dashboard")
ws2["A1"] = "Tableau de bord — suivi du projet (semaine 14)"
ws2["A1"].font = Font(bold=True, size=14)

jalons_termines = sum(1 for j in jalons if j[4] == "Terminé")
jalons_retard = sum(1 for j in jalons if j[4] == "En retard")
budget_prevu_engage = sum(j[5] for j in jalons if j[4] in ("Terminé", "En retard"))
budget_reel_engage = sum(j[6] for j in jalons if j[4] in ("Terminé", "En retard"))
ecart_pct = round((budget_reel_engage - budget_prevu_engage) / budget_prevu_engage * 100, 1)
avancement_pct = round(jalons_termines / len(jalons) * 100, 1)

kpi_rows = [
    ("Jalons terminés", f"{jalons_termines} / {len(jalons)}"),
    ("Jalons en retard", jalons_retard),
    ("Avancement (jalons terminés)", f"{avancement_pct} %"),
    ("Budget engagé (prévu)", f"{budget_prevu_engage} €"),
    ("Budget engagé (réel)", f"{budget_reel_engage} €"),
    ("Écart budgétaire", f"{ecart_pct:+.1f} %"),
]
ws2["A3"] = "Indicateurs clés"
ws2["A3"].font = Font(bold=True, size=12)
for i, (label, val) in enumerate(kpi_rows):
    ws2.cell(4 + i, 1, label).font = Font(bold=True)
    ws2.cell(4 + i, 2, val)

ws2.column_dimensions["A"].width = 32
ws2.column_dimensions["B"].width = 18

# Table pour graphique budget prévu vs réel (par jalon)
start_row = 12
ws2.cell(start_row, 1, "Jalon").font = Font(bold=True)
ws2.cell(start_row, 2, "Budget prévu").font = Font(bold=True)
ws2.cell(start_row, 3, "Budget réel").font = Font(bold=True)
for i, (nom, sd, sf, sr, statut, bp, br) in enumerate(jalons):
    ws2.cell(start_row + 1 + i, 1, nom)
    ws2.cell(start_row + 1 + i, 2, bp)
    ws2.cell(start_row + 1 + i, 3, br)

chart1 = BarChart()
chart1.type = "col"
chart1.title = "Budget prévu vs réel, par jalon"
chart1.y_axis.title = "Euros"
data = Reference(ws2, min_col=2, max_col=3, min_row=start_row, max_row=start_row + len(jalons))
cats = Reference(ws2, min_col=1, min_row=start_row + 1, max_row=start_row + len(jalons))
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.width = 18
chart1.height = 9
ws2.add_chart(chart1, "E12")

# Table pour graphique statuts (pie)
start_row2 = start_row + len(jalons) + 3
ws2.cell(start_row2, 1, "Statut").font = Font(bold=True)
ws2.cell(start_row2, 2, "Nombre de jalons").font = Font(bold=True)
statuts_count = {}
for j in jalons:
    statuts_count[j[4]] = statuts_count.get(j[4], 0) + 1
for i, (statut, count) in enumerate(statuts_count.items()):
    ws2.cell(start_row2 + 1 + i, 1, statut)
    ws2.cell(start_row2 + 1 + i, 2, count)

chart2 = PieChart()
chart2.title = "Répartition des jalons par statut"
data2 = Reference(ws2, min_col=2, min_row=start_row2, max_row=start_row2 + len(statuts_count))
cats2 = Reference(ws2, min_col=1, min_row=start_row2 + 1, max_row=start_row2 + len(statuts_count))
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.width = 12
chart2.height = 9
ws2.add_chart(chart2, "E30")

wb.save("suivi_projet_scoring_churn.xlsx")
print("Classeur généré : suivi_projet_scoring_churn.xlsx")
print("Avancement:", avancement_pct, "% | Jalons en retard:", jalons_retard, "| Écart budget:", ecart_pct, "%")
