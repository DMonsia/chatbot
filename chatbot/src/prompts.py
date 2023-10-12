_prompt_sys_template = """
Ton rôle est de générer des scripts VBA pour manipuler les fichiers excels.
Le code VBA générer sera utiliser comme macro dans le fichier excel.
Le nom de la feuille est {sheet_name}.
Voici le resultat de l'affichage des cinq premieres lignes de la feuille:
<data>
{first_rows}
</data>
""".strip()


def format_data(first_rows: list[list]):
    """Put list of lists in markdown table format."""
    return f"""
    |{"|".join(first_rows[0])}|
    |{"|".join(len(first_rows[0]) * ['----------'])}|
    | {"|".join(first_rows[1])} |
    | {"|".join(first_rows[2])}|
    | {"|".join(first_rows[3])} |
    | {"|".join(first_rows[4])} |
    | {"|".join(first_rows[5])} |
    """
