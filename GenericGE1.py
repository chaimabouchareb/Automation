import pandas as pd
from datetime import date, timedelta
import os
import tkinter as tk
from tkinter import simpledialog

# --- Configuration et Formatage (Les formats restent les mêmes) ---
TAB_COLORS = {
    0: '#4682B4', 1: '#3CB371', 2: '#FFD700', 3: '#FF8C00', 4: '#DC143C',
}
HEADERS = ['Jira ticket', 'status', 'type', 'project', 'comment', 'reporter', 'assignee']
STATUS_OPTIONS = [
    'Open', 'UAT', 'Done', 'Failed on test - Ready for dev', 'Rejected',
    'Ready for PRD', 'IAT', 'QAT', 'Implemented'
]
TYPE_OPTIONS = [
    'Sub-bug', 'US', 'Bug', 'Tech-Story', 'AC'
]

# ----------------------------------------------------------------------
## 1. Fonction pour la sélection de la date via Pop-up
# ----------------------------------------------------------------------

def get_month_and_year():
    """Affiche un pop-up pour demander le mois et l'année à l'utilisateur."""
    # Crée la fenêtre racine Tkinter (elle est masquée)
    ROOT = tk.Tk()
    ROOT.withdraw() 
    
    # Demande le mois (par exemple: 10 pour Octobre)
    month = simpledialog.askinteger(
        "Sélection du Mois et de l'Année", 
        "Entrez le numéro du mois (ex: 10 pour Octobre):", 
        parent=ROOT,
        minvalue=1, maxvalue=12
    )
    
    # Si l'utilisateur annule
    if month is None:
        return None, None

    # Demande l'année (par exemple: 2025)
    year = simpledialog.askinteger(
        "Sélection du Mois et de l'Année", 
        "Entrez l'année (ex: 2025):", 
        parent=ROOT,
        minvalue=2023, maxvalue=2050
    )
    
    return month, year

# ----------------------------------------------------------------------
## 2. Fonction de génération Excel (mise à jour pour utiliser les paramètres)
# ----------------------------------------------------------------------

def generate_excel_file(month, year):
    """Génère le fichier Excel formaté pour le mois et l'année donnés."""
    
    try:
        start_date = date(year, month, 1)
        # Calcule le dernier jour du mois (premier jour du mois suivant moins un jour)
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)
    except ValueError:
        print("Erreur: Les dates fournies sont invalides.")
        return

    month_name = start_date.strftime('%B').capitalize()
    
    all_sheets = []
    current_date = start_date
    weekend_counter = 0

    # 1. Déterminer les jours ouvrables et les weekends
    while current_date <= end_date:
        if 0 <= current_date.weekday() <= 4:
            sheet_name = current_date.strftime('%d-%m-%Y')
            all_sheets.append({'name': sheet_name, 'type': 'work', 'day_of_week': current_date.weekday()})
            
            if current_date.weekday() == 4 and current_date != end_date:
                weekend_counter += 1
                weekend_name = f'WEEKEND {weekend_counter}'
                all_sheets.append({'name': weekend_name, 'type': 'weekend', 'day_of_week': -1})
                
        current_date += timedelta(days=1)
        
    # 2. Créer le fichier Excel
    file_name = f'TO DO LIST - {month_name} {year}.xlsx'
    
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter') 

    work_data = pd.DataFrame(columns=HEADERS)
    weekend_data = pd.DataFrame() 

    workbook = writer.book
    
    # Définition des formats
    light_gray_format = workbook.add_format({'bg_color': '#D3D3D3'})
    header_format = workbook.add_format({
        'bold': True, 'fg_color': '#000099', 'font_color': '#FFFFFF',
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    # 3. Création et formatage des feuilles
    for sheet in all_sheets:
        worksheet = workbook.add_worksheet(sheet['name'])
        
        if sheet['type'] == 'work':
            day_color = TAB_COLORS[sheet['day_of_week']]
            body_format = workbook.add_format({'bg_color': day_color})
            
            worksheet.set_tab_color(day_color)
            
            # Écrire les EN-TÊTES
            for col_num, value in enumerate(HEADERS):
                worksheet.write(0, col_num, value, header_format)
            
            # Appliquer la couleur de fond
            worksheet.conditional_format('A2:G1000', {'type': 'no_blanks', 'format': body_format})
            worksheet.set_column('A:G', 20)
            
            # Listes Déroulantes (Status et Type)
            worksheet.data_validation('B2:B1000', {'validate': 'list', 'source': STATUS_OPTIONS})
            worksheet.data_validation('C2:C1000', {'validate': 'list', 'source': TYPE_OPTIONS})
            
        else:
            # Feuille WEEKEND
            worksheet.set_tab_color('#D3D3D3') 
            worksheet.conditional_format('A1:G1000', {'type': 'no_blanks', 'format': light_gray_format})
            worksheet.set_column('A:G', 20, light_gray_format) 

    # 4. Écrire les DataFrames vides
    for sheet in all_sheets:
        if sheet['type'] == 'work':
            work_data.to_excel(writer, sheet_name=sheet['name'], startrow=1, index=False, header=False)
        else:
            weekend_data.to_excel(writer, sheet_name=sheet['name'], index=False)
            
    writer.close()
    print(f"Fichier '{file_name}' créé avec succès pour {month_name} {year}.")

# ----------------------------------------------------------------------
## 3. Fonction principale pour exécuter le flux
# ----------------------------------------------------------------------

def run_excel_generator():
    """Exécute le pop-up, puis génère le fichier Excel si les entrées sont valides."""
    month, year = get_month_and_year()
    
    if month is not None and year is not None:
        generate_excel_file(month, year)
    else:
        print("Génération annulée par l'utilisateur.")

# Execute the main function
run_excel_generator()