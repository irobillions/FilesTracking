import csv
import pandas as pd
import os


def get_file_path(filename):
    file_path = os.path.abspath(filename)
    directory_path = os.path.dirname(file_path)

    return directory_path


def csv_to_excel(csv_file, excel_file):
    """
    Convertit un fichier CSV en tableau dans un fichier Excel.

    Args:
        csv_file (str): Chemin du fichier CSV d'entrée.
        excel_file (str): Chemin du fichier Excel de sortie.
    """
    try:
        df = pd.read_csv(csv_file)
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Formater le tableau comme un tableau Excel
            from openpyxl.worksheet.table import Table, TableStyleInfo

            # Définir la plage de cellules pour le tableau
            table_ref = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
            table = Table(displayName="Table1", ref=table_ref)

            style = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=True
            )
            table.tableStyleInfo = style

            # Ajouter le tableau à la feuille de calcul
            worksheet.add_table(table)

        print(f"Le fichier Excel a été créé avec succès : {excel_file}")
    except Exception as e:
        print(f"Erreur lors de la conversion du fichier CSV en Excel : {e}")


def write_results_to_csv(results, output_file):
    """
    Écrit les résultats de la recherche dans un fichier CSV.

    Args:
        results (list): Liste des objets FileInfo contenant les résultats de la recherche.
        output_file (str): Chemin du fichier CSV de sortie.
    """
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Nom du fichier', 'Chemin complet', 'Résumé']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for result in results:
                writer.writerow(
                    {'Nom du fichier': result.name, 'Chemin complet': result.path, 'Résumé': result.summary})
    except Exception as e:
        print(f"Erreur lors de l'écriture du fichier CSV: {output_file}, Erreur: {e}")
