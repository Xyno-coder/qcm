import pandas as pd
import os

# Définition des codes ANSI pour les couleurs
RED = '\033[91m'
GREEN = '\033[92m'
END = '\033[0m'  # Pour réinitialiser la couleur

def read_excel_files(file_paths):
    data = []
    for file_path in file_paths:
        if os.path.exists(file_path):
            excel_data = pd.read_excel(file_path, sheet_name=None)
            for sheet_name, df in excel_data.items():
                if df.shape[1] >= 2:  # Vérifie qu'il y a au moins deux colonnes
                    for row in df.itertuples(index=False, name=None):
                        if not pd.isna(row[1]):  # Vérifie si la réponse n'est pas NaN
                            data.append((str(row[0]).lower(), row[1]))  # Convertit la question en minuscules
    return data

def find_answer(data, query, printed_questions):
    results = [(index, question, answer) for index, (question, answer) in enumerate(data) if query.lower() in question.lower() and not pd.isna(answer) and question not in printed_questions]
    return results

if __name__ == "__main__":
    # Paths to the Excel files
    excel_files = [
        'QCM-FM-1.xlsx',
        'QCM-FM-2.xlsx',
        'QCM-FM-3.xlsx',
        'QCM-FM-4.xlsx'
    ]

    # Read data from all Excel files
    data = read_excel_files(excel_files)
    printed_questions = {}

    while True:
        i = 1
        query = input("Salut Beau gosse, tape le début de ta question: ")
        if query.lower() == "exit":
            break

        answers = find_answer(data, query.lower(), printed_questions)  # Convertit la requête en minuscules
        if answers:
            for index, question, answer in answers:
                if question not in printed_questions:
                    # Affichage de la question en rouge et de la réponse en vert
                    print(f"{RED}Question {i}:{END} {RED}{question}:{END} {GREEN}{answer}{END}\n")
                    printed_questions[question] = answer
                    i += 1
        else:
            print("Aucune nouvelle réponse trouvée dans la colonne B.")
