import pandas as pd
import os
import shutil
import re
from jinja2 import Environment
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory

# Funzione per selezionare il file Excel
def select_excel_file():
    root = Tk()
    root.withdraw()
    return askopenfilename(title="Seleziona il file Excel", filetypes=[("Excel files", "*.xlsx")])

# Funzione per selezionare la cartella delle immagini
def select_image_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Seleziona la cartella contenente le immagini")

# Funzione per selezionare la cartella dei loghi
def select_logo_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Seleziona la cartella contenente i loghi")

# Funzione per selezionare la cartella di output per salvare i file HTML
def select_output_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Seleziona la cartella di output per salvare i file HTML")

# Funzione per selezionare l'icona home
def select_home_icon():
    root = Tk()
    root.withdraw()
    return askopenfilename(title="Seleziona l'icona home (PNG)", filetypes=[("PNG files", "*.png")])

# Funzione per copiare file con richiesta di sovrascrittura
def copy_file_with_prompt(src, dest):
    # Controlla se il file di origine esiste
    if not os.path.exists(src):
        print(f"❌ ERRORE: Il file '{src}' non esiste! Salto la copia.")
        return

    # Controlla se il file sorgente e il file di destinazione sono lo stesso
    if os.path.abspath(src) == os.path.abspath(dest):
        print(f"ℹ️ Il file '{src}' è già nella destinazione. Nessuna copia necessaria.")
        return

    print(f"🔄 Copiando da: {os.path.abspath(src)} a {os.path.abspath(dest)}")

    if os.path.exists(dest):
        response = messagebox.askyesno("File Esistente", f"Il file '{os.path.basename(dest)}' esiste già.\nVuoi sovrascriverlo?")
        if not response:
            print(f"✅ File mantenuto: {dest}")
            return

    shutil.copy(src, dest)
    print(f"✅ File copiato: {dest}")



# Funzione per creare le pagine HTML
def create_html_pages(excel_file, image_folder, logo_folder, output_folder, home_icon):
    df = pd.read_excel(excel_file, engine="openpyxl")

    template_html = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Image Page</title>
        <style>
            .container { width: 80%; margin: 0 auto; text-align: center; position: relative; }
            .image { width: 100%; max-width: 600px; height: auto; }
            .data-table { width: 100%; margin: 20px 0; border-collapse: collapse; }
            .data-table th, .data-table td { border: 1px solid #ddd; padding: 8px; }
            .data-table th { background-color: #f2f2f2; }
            .italic { font-style: italic; }
            .title { font-size: 1.5em; }
            .logo-container { width: 100%; margin-top: 20px; display: flex; justify-content: space-between; }
            .logo { width: 100px; height: auto; }
            .home-icon { width: 40px; height: auto; position: absolute; top: 20px; right: 20px; }
        </style>
    </head>
    <body>
        <div class="container">
            <a href="home.html">
                <img src="home.png" alt="Home" class="home-icon" class="home-icon">
            </a>
            <h1 class="italic title">{{ data['NomeScientifico'] }}</h1>
            <img src="{{ image }}" alt="Image" class="image">
            <h2>Data</h2>
            <table class="data-table">
                <tbody>
                    <tr><td>Kingdom:</td><td>{{ data['Regno'] }}</td></tr>
                    <tr><td>Phylum:</td><td>{{ data['Phylum'] }}</td></tr>
                    <tr><td>Class:</td><td>{{ data['Classe'] }}</td></tr>
                    <tr><td>Order:</td><td>{{ data['Ordine'] }}</td></tr>
                    <tr><td>Family:</td><td>{{ data['Famiglia'] }}</td></tr>
                    <tr><td>Genus:</td><td class="italic">{{ data['Genere'] }}</td></tr>
                    <tr><td>Species:</td><td class="italic">{{ data['Specie'] }}</td></tr>
                    <tr><td>Identifier:</td><td>{{ data['Identificatore'] }}</td></tr>
                    <tr><td>Continent:</td><td>{{ data['Continente'] }}</td></tr>
                    <tr><td>Country:</td><td>{{ data['Paese'] }}</td></tr>
                    <tr><td>Location:</td><td>{{ data['Località'] }}</td></tr>
                    <tr><td>Coordinates:</td><td>{{ data['Coordinate'] }}</td></tr>
                    <tr><td>Gbif:</td><td>{{ data['OccurrenceID'] }}</td></tr>
                </tbody>
            </table>
            <div class="logo-container">
                <img src="logo1.png" alt="Logo 1" class="logo" style="align-self: flex-start;">
                <img src="logo2.png" alt="Logo 2" class="logo" style="align-self: flex-end;">
            </div>
        </div>
    </body>
    </html>
    '''

    env = Environment()
    template = env.from_string(template_html)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    print(f"Cartella di output: {output_folder}")

    # Copia i loghi e l'icona home con richiesta di sovrascrittura
    copy_file_with_prompt(os.path.join(logo_folder, 'logo1.png'), os.path.join(output_folder, 'logo1.png'))
    copy_file_with_prompt(os.path.join(logo_folder, 'logo2.png'), os.path.join(output_folder, 'logo2.png'))
    home_dest = os.path.join(output_folder, 'home.png')

    # Evita di copiare il file se è già nella destinazione
    if os.path.abspath(home_icon) != os.path.abspath(home_dest):
        copy_file_with_prompt(home_icon, home_dest)
    else:
        print(f"ℹ️ Il file home.png è già nella destinazione. Nessuna copia necessaria.")

        logo1_path = "logo1.png"
        logo2_path = "logo2.png"
        home_icon = "home.png"


    # Itera attraverso le righe del DataFrame per creare i file HTML
    for index, row in df.iterrows():
        if pd.isna(row['Foto']):
            continue

        image_file = os.path.join(image_folder, row['Foto'])
        if not os.path.exists(image_file):
            print(f"❌ Immagine non trovata: {image_file}")
            continue

        # Percorso relativo per l'HTML (senza copia)
        image_path = row['Foto']

        data = row.to_dict()
        data['Coordinate'] = row['Coordinate'] if pd.notna(row['Coordinate']) else "Not available"

        html_content = template.render(
        image=image_path, 
        data=data, 
        logo1=logo1_path, 
        logo2=logo2_path, 
        home_icon=home_icon
)

        image_name = os.path.splitext(row['Foto'])[0]
        safe_image_name = re.sub(r'[^a-zA-Z0-9_-]', '_', image_name)
        output_file = os.path.join(output_folder, f"{safe_image_name}.html")

        print(f"✅ Creando file HTML: {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

    print(f'🎉 HTML pages created in {output_folder}')



# Selezione dei file e cartelle
excel_file = select_excel_file()
image_folder = select_image_folder()
logo_folder = select_logo_folder()
output_folder = select_output_folder()
home_icon = select_home_icon()

if excel_file and image_folder and logo_folder and output_folder and home_icon:
    create_html_pages(excel_file, image_folder, logo_folder, output_folder, home_icon)
else:
    print("Operazione annullata: uno o più file/cartelle non sono stati selezionati.")
