import pandas as pd
import os
import shutil
import re
from jinja2 import Environment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

# Funzioni per selezionare file e cartelle
def select_file(title, filetypes):
    root = Tk()
    root.withdraw()
    return askopenfilename(title=title, filetypes=filetypes)

def select_folder(title):
    root = Tk()
    root.withdraw()
    return askdirectory(title=title)

# Funzione per creare pagine HTML
def create_html_pages(excel_file, output_folder, home_icon, left_logo, right_logo, image_folder):
    try:
        df = pd.read_excel(excel_file, sheet_name='Occurrences', engine="openpyxl")
        print(f"Dati letti correttamente da {excel_file}")
    except Exception as e:
        print(f"Errore nella lettura del file Excel: {e}")
        return

    try:
        df_grouped = df.groupby(['NomeScientifico', 'Continente', 'Paese', 'Località'], as_index=False).size()
    except Exception as e:
        print(f"Errore nel raggruppamento dei dati: {e}")
        return

    env = Environment()

    # Template per la pagina della specie
    species_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ nome_scientifico }}</title>
        <style>
            .container { text-align: center; }
            .data-table { width: 80%; margin: 20px auto; border-collapse: collapse; }
            .data-table th, .data-table td { border: 1px solid #ddd; padding: 8px; }
            h1 { font-style: italic; }
            .home-icon { position: absolute; top: 10px; left: 10px; width: 40px; height: 40px; }
            .logos { display: flex; justify-content: space-between; margin-top: 20px; }
            .logos img { width: 100px; }
            .gallery-button { margin-top: 20px; display: inline-block; padding: 10px 20px; background: #007bff; color: white; text-decoration: none; border-radius: 5px; }
        </style>
    </head>
    <body>
        <a href="home.html">
            <img src="home.png" alt="Home" class="home-icon">
        </a>
        <div class="container">
            <h1>{{ nome_scientifico }}</h1>
            <h2>Data</h2>
            <table class="data-table">
                <tr><th>Continent</th><th>Country</th><th>Location</th><th>Individuals</th></tr>
                {% for row in data %}
                <tr>
                    <td>{{ row.Continente }}</td>
                    <td>{{ row.Paese }}</td>
                    <td>{{ row.Località }}</td>
                    <td>{{ row.size }}</td>
                </tr>
                {% endfor %}
            </table>

            <a href="{{ gallery_filename }}" class="gallery-button">Gallery</a>

            <div class="logos">
                <img src="logo1.png" alt="Left Logo">
                <img src="logo2.png" alt="Right Logo">
            </div>
        </div>
    </body>
    </html>
    ''')

    # Template per la galleria
    gallery_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Gallery - {{ nome_scientifico }}</title>
        <style>
            .gallery { display: flex; flex-wrap: wrap; justify-content: center; gap: 10px; }
            .gallery img { width: 150px; height: auto; border: 1px solid #ccc; }
        </style>
    </head>
    <body>
        <h1>Gallery - {{ nome_scientifico }}</h1>
        <div class="gallery">
            {% for image in images %}
            <a href="{{ image.page_filename }}"><img src="{{ image.src }}" alt=""></a>
            {% endfor %}
        </div>
    </body>
    </html>
    ''')

    # Template per immagine singola
    image_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ nome_immagine }}</title>
    </head>
    <body style="text-align:center">
        <h1>{{ nome_immagine }}</h1>
        <img src="{{ src }}" style="max-width: 90%; height: auto;">
        <br><br>
        <a href="{{ back_link }}">Back to gallery</a>
    </body>
    </html>
    ''')

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Copia delle immagini principali se necessario
    for icon, name in [(home_icon, 'home.png'), (left_logo, 'logo1.png'), (right_logo, 'logo2.png')]:
        dst = os.path.join(output_folder, name)
        if not os.path.exists(dst):
            shutil.copy(icon, dst)

    # Generazione dei file HTML
    for nome_scientifico, group in df_grouped.groupby('NomeScientifico'):
        nome_scientifico_sanitized = re.sub(r'[^a-zA-Z0-9_-]', '_', nome_scientifico).lower()
        html_filename = f"{nome_scientifico_sanitized}.html"
        gallery_filename = f"gallery_{nome_scientifico_sanitized}.html"

        species_rows = df[df['NomeScientifico'] == nome_scientifico]
        immagini = species_rows['Foto'].dropna().unique().tolist()

        image_entries = []
        for img in immagini:
            abs_path = os.path.join(image_folder, img)
            image_entries.append({
                'src': f"file:///{abs_path.replace(os.sep, '/')}",
                'page_filename': f"img_{nome_scientifico_sanitized}_{img}.html",
                'filename': img
            })

        for image in image_entries:
            image_html = image_template.render(
                nome_immagine=image['filename'],
                src=image['src'],
                back_link=gallery_filename
            )
            with open(os.path.join(output_folder, image['page_filename']), 'w', encoding='utf-8') as f:
                f.write(image_html)

        gallery_html = gallery_template.render(
            nome_scientifico=nome_scientifico,
            images=image_entries
        )
        with open(os.path.join(output_folder, gallery_filename), 'w', encoding='utf-8') as f:
            f.write(gallery_html)

        species_html = species_template.render(
            nome_scientifico=nome_scientifico,
            data=group.to_dict(orient='records'),
            gallery_filename=gallery_filename
        )
        with open(os.path.join(output_folder, html_filename), 'w', encoding='utf-8') as f:
            f.write(species_html)

    print(f'Tutti i file HTML sono stati creati in: {output_folder}')

# Selezione file e cartelle
excel_file = select_file("Seleziona il file Excel", [("Excel files", "*.xlsx")])
output_folder = select_folder("Seleziona la cartella di output per salvare i file HTML")
image_folder = select_folder("Seleziona la cartella contenente tutte le immagini delle specie")
home_icon = select_file("Seleziona l'icona Home", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
left_logo = select_file("Seleziona il logo sinistro", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
right_logo = select_file("Seleziona il logo destro", [("Image files", "*.png;*.jpg;*.jpeg;*.gif")])

# Controllo che tutto sia stato selezionato
if all([excel_file, output_folder, image_folder, home_icon, left_logo, right_logo]):
    try:
        create_html_pages(excel_file, output_folder, home_icon, left_logo, right_logo, image_folder)
    except Exception as e:
        print(f"\n❌ Errore durante la creazione delle pagine HTML: {e}\n")
else:
    print("\n⚠️ Operazione annullata: uno o più file/cartelle non sono stati selezionati.\n")

# Permette di vedere i messaggi prima che la finestra si chiuda
input("Premi un tasto per uscire...")
