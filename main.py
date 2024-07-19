from tracking.trackingFunc import *
from tracking.utils import *
import pandas as pd


def display_csv(file_path):
    """
        Affiche le contenu d'un fichier CSV.

        Args:
            file_path (str): Chemin du fichier CSV.
        """
    if os.path.exists(file_path):
        df = pd.read_csv(file_path)
        print(df)
    else:
        print(f"Le fichier {file_path} n'existe pas.")


def main():
    """
    Fonction principale pour exécuter la recherche et écrire les résultats.
    """
    directory = input("Entrez le chemin du répertoire: ")
    keyword = input("Entrez le mot-clé à rechercher: ")

    # # Recherche des fichiers par contenu
    # results_content = search_files_contain_keyword(directory, keyword)
    # if not results_content:
    #     print(f"Aucun fichier trouvé contenant le mot-clé '{keyword}' dans son contenu.")
    # else:
    #     output_file_content = 'resultats_contenu.csv'
    #     write_results_to_csv(results_content, output_file_content)
    #     print(f"Résultats des contenus écrits dans '{output_file_content}'.")

    # Recherche des fichiers par nom
    results_name = search_files_with_keyword(directory, keyword)
    if not results_name:
        print(f"Aucun fichier trouvé contenant le mot-clé '{keyword}' dans son nom.")
    else:
        output_file_name = 'resultats_nom_fichiers.csv'
        write_results_to_csv(results_name, output_file_name)
        print(f"Résultats des noms de fichiers écrits dans '{output_file_name}'.")

        # Exemple d'utilisation
        path_dir = get_file_path(output_file_name)

        csv_file = path_dir + '\\' + output_file_name  # Remplacez par le chemin de votre fichier CSV
        excel_file = path_dir + '\\' + 'results.xlsx'  # Remplacez par le chemin de votre fichier Excel de sortie
        csv_to_excel(csv_file, excel_file)


if __name__ == "__main__":
    main()
