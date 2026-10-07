# Suivre un projet après l'avoir cadré

## En bref
Ce projet fait suite à [note-cadrage-scoring-ia-pme](https://github.com/ATEYABA-K/note-cadrage-scoring-ia-pme) : après avoir écrit la note de cadrage (objectifs, jalons, risques), je me suis demandé comment je suivrais concrètement ce projet une fois lancé, pas seulement le planifier, mais mesurer où il en est vraiment. Toujours un cas fictif (aucune vraie entreprise), simulé jusqu'à la semaine 14 sur les 24 prévues.

## Le principe
Un tableau de suivi (`Suivi_Jalons`) reprend exactement les 6 jalons de la note de cadrage, avec pour chacun : la semaine prévue, la semaine réelle de fin (si terminé), un statut, et un budget prévu vs réel. Un onglet `Dashboard` calcule les indicateurs à partir de ce tableau et les met en graphique.

## Les résultats, à la semaine 14

- **Avancement** : 3 jalons terminés sur 6 (50 %)
- **1 jalon en retard** : la "Validation métier" (prévue S7-S9), toujours en cours à S14, un retard de plusieurs semaines
- **Écart budgétaire cumulé** : −1,9 % (léger mieux que prévu au global, mais ça cache un vrai dépassement sur "Données consolidées" : +15 %, cohérent avec le risque "données insuffisantes ou mal qualifiées" déjà identifié dans la note de cadrage)

Le dashboard visualise ça avec un graphique budget prévu/réel par jalon et une répartition des statuts.

## Ce que j'en retiens
Le retard sur "Validation métier" n'est pas une surprise si on relit la note de cadrage : c'est le jalon qui dépend le plus des retours humains (chargés de compte), donc le plus difficile à cadencer précisément. Faire ce suivi m'a fait comprendre une différence concrète entre cadrer et piloter : le cadrage anticipe les risques en théorie, le suivi révèle lesquels se sont vraiment matérialisés, et sur quel jalon précis, pas juste "le projet a du retard" en général.

## Reproduire
```bash
pip install -r requirements.txt
python3 scripts/build_dashboard.py
```
Génère `suivi_projet_scoring_churn.xlsx` (2 onglets : Suivi_Jalons, Dashboard).

Alvin Kouadio
