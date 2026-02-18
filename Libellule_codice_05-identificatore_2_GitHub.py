import pandas as pd
import os
import re
import shutil
from jinja2 import Environment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

def select_file(title, filetypes):
    root = Tk()
    root.withdraw()
    return askopenfilename(title=title, filetypes=filetypes)

def select_folder(title):
    root = Tk()
    root.withdraw()
    return askdirectory(title=title)

def clean_filename(identificatore):
    identificatori = [name.strip() for name in identificatore.split(',')]  
    identificatori = [re.sub(r"[^a-zA-Z0-9 ]", '', name) for name in identificatori]  
    identificatori = [name.replace(' ', '_') for name in identificatori]  
    identificatori = list(dict.fromkeys(identificatori))  
    return "__".join(identificatori).lower().lstrip('/')  # 👈 aggiunta .lstrip('/')


def safe_copy(src, dest):
    try:
        # Confronto dei percorsi assoluti per evitare copia inutile
        if os.path.abspath(src) != os.path.abspath(dest):
            shutil.copy(src, dest)
            print(f"File copiato da {src} a {dest}")
        else:
            print(f"I file {src} e {dest} sono già identici, copia ignorata.")
    except Exception as e:
        print(f"Errore nella copia del file {src} a {dest}: {e}")

def create_html_pages_by_identifier(excel_file, output_folder, home_file, home_icon, left_logo, right_logo):
    try:
        # Leggi il file Excel
        df = pd.read_excel(excel_file, sheet_name='Occurrences', engine="openpyxl")
        print("Colonne trovate nel file Excel:", df.columns)
        print("Prime righe del file Excel:")
        print(df.head())
    except Exception as e:
        print(f"Errore nella lettura del file Excel: {e}")
        return
    
    df.fillna('', inplace=True)
    
    # Pulizia dei dati
    df['Identificatore'] = df['Identificatore'].astype(str).str.strip()
    df['NomeScientifico'] = df['NomeScientifico'].astype(str).str.strip()
    df['Continente'] = df['Continente'].astype(str).str.strip()
    df['Paese'] = df['Paese'].astype(str).str.strip()
    df['Località'] = df['Località'].astype(str).str.strip()
    
    # Raggruppamento dei dati
    df_grouped = df.groupby(
        ['Identificatore', 'NomeScientifico', 'Continente', 'Paese', 'Località'], as_index=False
    ).size().rename(columns={'size': 'NumeroIndividui'})
    
    print("Esempio di dati raggruppati:")
    print(df_grouped.head())
    
    # Template HTML
    template_html = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ identificatore }}</title>
        <style>
            .container { text-align: center; }
            .data-table { width: 80%; margin: 20px auto; border-collapse: collapse; }
            .data-table th, .data-table td { border: 1px solid #ddd; padding: 8px; }
            h1 { font-style: italic; }
            .home-icon { position: absolute; top: 10px; left: 10px; width: 40px; height: 40px; }
            .logos { display: flex; justify-content: space-between; margin-top: 20px; }
            .logos img { width: 100px; }
        </style>
    </head>
    <body>
        <a href="home.html">
            <img src="{{ home_icon }}" alt="Home" class="home-icon">
        </a>
        <div class="container">
            <h1>{{ identificatore }}</h1>
            <h2>Data</h2>
            <table class="data-table">
                <tr><th>Species</th><th>Continent</th><th>Country</th><th>Location</th><th>Individuals</th></tr>
                {% for row in data %}
                <tr>
                    <td><em>{{ row.NomeScientifico }}</em></td>
                    <td>{{ row.Continente }}</td>
                    <td>{{ row.Paese }}</td>
                    <td>{{ row.Località }}</td>
                    <td>{{ row.NumeroIndividui }}</td>
                </tr>
                {% endfor %}
            </table>
            <div class="logos">
                <img src="{{ left_logo }}" alt="Left Logo">
                <img src="{{ right_logo }}" alt="Right Logo">
            </div>
        </div>
    </body>
    </html>
    '''
    
    env = Environment()
    template = env.from_string(template_html)
    
    # Creazione della cartella di output se non esiste
    os.makedirs(output_folder, exist_ok=True)
    
    # Copia delle immagini nella cartella di output
    try:
        safe_copy(home_icon, os.path.join(output_folder, 'home.png'))
        safe_copy(left_logo, os.path.join(output_folder, 'logo1.png'))
        safe_copy(right_logo, os.path.join(output_folder, 'logo2.png'))
        print("Immagini copiate correttamente.")
    except Exception as e:
        print(f"Errore nel copiare le immagini: {e}")
        return

    # Percorsi relativi per le immagini copiate
    home_icon_rel = 'home.png'
    left_logo_rel = 'logo1.png'
    right_logo_rel = 'logo2.png'
    
    # Creazione dei file HTML
    for identificatore, group in df_grouped.groupby('Identificatore'):
        try:
            print(f"Identificatore: {identificatore}, Numero di righe trovate: {len(group)}")
            safe_filename = clean_filename(identificatore)
            output_file = os.path.join(output_folder, f"{safe_filename}.html")
            
            html_content = template.render(
                identificatore=identificatore,
                data=group.to_dict(orient='records'),
                home_page="home.html",
                home_icon=home_icon_rel,
                left_logo=left_logo_rel,
                right_logo=right_logo_rel
            )

            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"File HTML creato per {identificatore}: {output_file}")
        except Exception as e:
            print(f"Errore nella creazione del file HTML per {identificatore}: {e}")
    
    print(f'HTML pages created in {output_folder}')

# Selezione dei file e delle cartelle
excel_file = select_file("Seleziona il file Excel", [("Excel files", "*.xlsx")])
output_folder = select_folder("Seleziona la cartella di output per salvare i file HTML")
home_file = select_file("Seleziona il file Home", [("HTML files", "*.html")])
home_icon = select_file("Seleziona l'icona Home", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
left_logo = select_file("Seleziona il logo sinistro", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
right_logo = select_file("Seleziona il logo destro", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])

if all([excel_file, output_folder, home_file, home_icon, left_logo, right_logo]):
    create_html_pages_by_identifier(excel_file, output_folder, home_file, home_icon, left_logo, right_logo)
else:
    print("Operazione annullata: uno o più file/cartelle non sono stati selezionati.")

input("Premi Invio per chiudere...")
