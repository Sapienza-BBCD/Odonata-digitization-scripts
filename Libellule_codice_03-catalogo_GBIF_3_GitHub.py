import os
import json
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
from string import Template

Tk().withdraw()

file_path = askopenfilename(title="Seleziona il file Excel", filetypes=[("Excel files", "*.xlsx")])
if not file_path:
    print("Nessun file selezionato. Il programma verrà terminato.")
    exit()

html_folder_path = askdirectory(title="Seleziona la cartella contenente i file .html")
if not html_folder_path:
    print("Nessuna cartella selezionata. Il programma verrà terminato.")
    exit()

logo1_path = askopenfilename(title="Seleziona il file logo1", filetypes=[("Image files", "*.png")])
logo2_path = askopenfilename(title="Seleziona il file logo2", filetypes=[("Image files", "*.png")])
if not logo1_path or not logo2_path:
    print("Logo non selezionato. Il programma verrà terminato.")
    exit()

def select_home_icon():
    root = Tk()
    root.withdraw()
    return askopenfilename(title="Seleziona l'icona home (PNG)", filetypes=[("PNG files", "*.png")])

# NUOVO BLOCCO: selezione home icon
home_icon_path = select_home_icon()
if not home_icon_path:
    print("Icona home non selezionata. Il programma verrà terminato.")
    exit()

save_folder_path = askdirectory(title="Seleziona la cartella dove salvare il file HTML")
if not save_folder_path:
    print("Nessuna cartella di salvataggio selezionata. Il programma verrà terminato.")
    exit()

try:
    occurrences_df = pd.read_excel(file_path, sheet_name='Occurrences')
    occurrences_df.columns = occurrences_df.columns.str.strip()
    occurrences_df = occurrences_df.dropna(subset=['OccurrenceID', 'Foto'])
    occurrence_data = occurrences_df.set_index('OccurrenceID')['Foto'].to_dict()
except Exception as e:
    print(f"Errore durante la lettura del file Excel: {e}")
    exit()

catalog_template = Template("""
<!DOCTYPE html>
<html>
<head>
    <title>GBIF Catalog</title>
    <style>
        .container {
            text-align: center;
            padding: 20px;
            margin-bottom: 150px;
        }
        .box {
            display: inline-block;
            width: 40px;
            height: 40px;
            line-height: 40px;
            text-align: center;
            font-size: 24px;
            border: 2px solid #000;
            margin: 5px;
            background-color: #600;
            font-style: inherit;
            color: #FFF;
        }
        .dropdown {
            width: 60px;
            font-size: 20px;
        }
        button {
            font-size: 24px;
            background-color: #821a33;
            color: white;
            padding: 15px 25px;
            border: none;
            cursor: pointer;
            border-radius: 8px;
        }
        button:hover {
            background-color: #c54c00;
        }
        .logos {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 20px;
        }
        .logos img {
            width: 178.5px;
            height: 47px;
        }
        .home-icon {
            width: 40px;
            height: auto;
            position: absolute;
            top: 20px;
            left: 20px;
            cursor: pointer;
        }
        #result {
            font-size: 20px;
            color: red;
        }
        h2 {
            color: #821a33;
            font-size: 60px;
            text-align: center;
        }
    </style>
</head>
<body>
    <!-- ICONA HOME AGGIUNTA QUI -->
    <img src="$home_icon" alt="Home" class="home-icon" onclick="location.href='index.html'" />

    <h2>GBIF Catalog</h2>
    <div class="container">
        <div class="box">M</div>
        <div class="box">Z</div>
        <div class="box">U</div>
        <div class="box">R</div>
        <div class="box">O</div>
        <div class="box">D</div>
        <div class="box">O</div>
        <div class="box">B</div>
        <select id="num1" class="dropdown">$numbers</select>
        <select id="num2" class="dropdown">$numbers</select>
        <select id="num3" class="dropdown">$numbers</select>
        <select id="num4" class="dropdown">$numbers</select>
        <select id="num5" class="dropdown">$numbers</select>
        <br>
        <button onclick="checkCode()">Open Record</button>
        <p id="result"></p>
    </div>
    <div class="logos">
        <img src="$logo1" alt="logo 1"/>
        <img src="$logo2" alt="logo 2"/>
    </div>
    <script type="text/javascript">
        function checkCode() {
            var code = "MZURODOB" +
                       document.getElementById("num1").value +
                       document.getElementById("num2").value +
                       document.getElementById("num3").value +
                       document.getElementById("num4").value +
                       document.getElementById("num5").value;

            var occurrenceData = $occurrences;

            var resultElement = document.getElementById("result");
            if (code in occurrenceData) {
                var jpgFile = occurrenceData[code];
                var htmlFile = jpgFile.replace('.jpg', '.html');
                var relativePath = htmlFile;
                resultElement.innerHTML = "<a href='" + relativePath + "' target='_self'>Apri " + htmlFile + "</a>";
            } else {
                resultElement.innerText = "Codice mancante";
            }
        }
    </script>
</body>
</html>
""")

numbers_options = "\n".join([f"<option value='{i}'>{i}</option>" for i in range(10)])

html_content = catalog_template.substitute(
    numbers=numbers_options,
    occurrences=json.dumps(occurrence_data),
    html_folder=html_folder_path.replace("\\", "/"),
    logo1=os.path.basename(logo1_path),
    logo2=os.path.basename(logo2_path),
    home_icon=os.path.basename(home_icon_path)
)


catalogo_gbif_path = os.path.join(save_folder_path, 'catalogo_gbif.html')
with open(catalogo_gbif_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"✅ File HTML 'catalogo_gbif.html' creato con successo in: {catalogo_gbif_path}")
