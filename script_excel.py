import pandas as pd
from openpyxl import load_workbook

def modifier_excel(file_path):
    # Charger le fichier Excel avec openpyxl
    workbook = load_workbook(file_path)

    # Récupérer la feuille active
    feuille_active = workbook.active

    # Vérifier la valeur de la cellule F5
    valeur_F5 = feuille_active['F5'].value
    
    if valeur_F5 == 'carbone bas':
        # Récupérer les valeurs des cellules H43 à H46
        valeurs_H43_H46 = [feuille_active[f'H{i}'].value for i in range(43, 47)]

        # Calculer le minimum des cellules H43:H46
        minimum_H43_H46 = min(valeurs_H43_H46)

        # Enregistrer le résultat dans la cellule H47
        feuille_active['H47'] = minimum_H43_H46

        # Sauvegarder les modifications dans le fichier Excel
        workbook.save(file_path)

        print("Modifications enregistrées avec succès.")
    else:
        print("La valeur de la cellule F5 n'est pas 'carbone bas'.")

# Chemin d'accès vers le fichier Excel
chemin_excel = r"C:\Users\douni\OneDrive\Bureau\Imane\outil.xlsx"

# Appeler la fonction pour effectuer les modifications
modifier_excel(chemin_excel)
