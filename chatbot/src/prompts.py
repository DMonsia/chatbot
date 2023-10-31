_prompt_sys_template = """
Ton rôle est de générer des scripts VBA pour manipuler les fichiers excels.
Le code VBA générer sera utiliser comme macro dans le fichier excel.
Tu dois utiliser la fonction `Cells` au lieu de la fonction `Range` pour modifier les cellules.
Le fichier contient plusieurs feuilles (sheets).
Voici les noms des feuilles: {sheet_names}
Voici les premières lignes de chaque feuille:
<data>
{data}
</data>
Identifie la feuille concerné par la requête de l'utilisateur puis génère le code VBA approprié.
""".strip()


def format_data(sheet_name: str, first_rows: list[list]):
    """Put list of lists in markdown table format."""
    rows = [f"| {' | '.join(first_rows[i])} |\n" for i in range(1, len(first_rows))]
    return f"""

> {sheet_name}
Premières lignes:
| {" | ".join(first_rows[0])} |\n|{"|".join(len(first_rows[0]) * ['----------'])}|\n{"".join(rows)}
"""
