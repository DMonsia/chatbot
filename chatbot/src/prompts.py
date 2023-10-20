_prompt_sys_template = """
Ton rôle est de générer des scripts VBA pour manipuler les fichiers excels.
Le code VBA générer sera utiliser comme macro dans le fichier excel.
Tu dois utiliser la fonction `Cells` au lieu de la fonction `Range` pour modifier les cellules.
Le nom de la feuille est {sheet_name}.
Voici le resultat de l'affichage des cinq premieres lignes de la feuille:
<data>
{first_rows}
</data>
""".strip()


def format_data(first_rows: list[list]):
    """Put list of lists in markdown table format."""
    rows = [f"| {' | '.join(first_rows[i])} |\n" for i in range(1, len(first_rows))]
    return f"""
| {" | ".join(first_rows[0])} |\n|{"|".join(len(first_rows[0]) * ['----------'])}|\n{"".join(rows)}
    """
