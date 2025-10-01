import pandas as pd
from datetime import date, timedelta
import os
import tkinter as tk
from tkinter import simpledialog, messagebox

# --- Configuration et Formatage ---
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
## 1. Fonction de Génération Excel
# ----------------------------------------------------------------------

def generate_excel_file(month, year):
    """Génère le fichier Excel formaté pour le mois et l'année donnés."""
    
    try:
        start_date = date(year, month, 1)
        # Calcule le dernier jour du mois
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)
    except ValueError:
        messagebox.showerror("Erreur de Date", "Le mois ou l'année est invalide.")
        return

    # --- CHANGEMENT CLÉ POUR LE NOM DU FICHIER ---
    # '%m' -> numéro du mois (09, 10)
    # '%B' -> nom complet du mois (September, October)
    formatted_month = start_date.strftime('%m-%B') 
    
    file_name = f'TO DO LIST-{formatted_month} {year}.xlsx'
    # Exemple: TO DO LIST-10-October 2025 FINAL.xlsx
    # ---------------------------------------------
    
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
    try:
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
        messagebox.showinfo("Succès", f"Fichier '{file_name}' créé avec succès dans le même dossier que l'application.")
        
    except Exception as e:
        messagebox.showerror("Erreur de Fichier", f"Impossible de générer le fichier : {e}")

# ----------------------------------------------------------------------
## 2. Interface Utilisateur (GUI)
# ----------------------------------------------------------------------

class ExcelGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("Générateur de Liste TO DO")

        self.frame = tk.Frame(master, padx=20, pady=20)
        self.frame.pack()

        # Labels et Champs de saisie (Mois)
        self.label_month = tk.Label(self.frame, text="Mois (1-12):")
        self.label_month.grid(row=0, column=0, pady=5, sticky='w')

        self.entry_month = tk.Entry(self.frame)
        self.entry_month.grid(row=0, column=1, pady=5)
        self.entry_month.insert(0, str(date.today().month))

        # Labels et Champs de saisie (Année)
        self.label_year = tk.Label(self.frame, text="Année (ex: 2025):")
        self.label_year.grid(row=1, column=0, pady=5, sticky='w')

        self.entry_year = tk.Entry(self.frame)
        self.entry_year.grid(row=1, column=1, pady=5)
        self.entry_year.insert(0, str(date.today().year))

        # Bouton Générer
        self.generate_button = tk.Button(
            self.frame, 
            text="Générer Fichier Excel", 
            command=self.validate_and_generate,
            bg='#3CB371',
            fg='white',
            padx=10,
            pady=5
        )
        self.generate_button.grid(row=2, column=0, columnspan=2, pady=20)

    def validate_and_generate(self):
        """Valide les entrées et appelle la fonction de génération."""
        try:
            month = int(self.entry_month.get())
            year = int(self.entry_year.get())
            
            if 1 <= month <= 12 and 2020 <= year <= 2050:
                self.generate_button.config(state=tk.DISABLED, text="Génération en cours...")
                self.master.update()
                
                generate_excel_file(month, year)
                
                self.generate_button.config(state=tk.NORMAL, text="Générer Fichier Excel")
            else:
                messagebox.showerror("Erreur de Saisie", "Veuillez entrer un mois entre 1 et 12 et une année valide (2020-2050).")
        
        except ValueError:
            messagebox.showerror("Erreur de Saisie", "Veuillez entrer des nombres entiers pour le mois et l'année.")


# ----------------------------------------------------------------------
## 3. Fonction principale
# ----------------------------------------------------------------------

def run_excel_generator_app():
    """Crée et lance l'application GUI."""
    root = tk.Tk()
    app = ExcelGeneratorApp(root)
    root.mainloop()

# Execute the main function
if __name__ == '__main__':
    run_excel_generator_app()