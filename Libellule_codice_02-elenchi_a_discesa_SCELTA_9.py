import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
import json
import os

# Crea una finestra nascosta per la selezione del file
Tk().withdraw()

# Chiedi di selezionare il file Excel
file_path = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx")])

if not file_path:
    print("No file selected. The program will exit.")
else:
    html_folder_path = askdirectory(title="Select the folder containing .html files")

    if not html_folder_path:
        print("No folder selected. The program will exit.")
    else:
        save_folder_path = askdirectory(title="Select the folder to save the HTML file")

        if not save_folder_path:
            print("No save folder selected. The program will exit.")
        else:
            logo1_path = askopenfilename(title="Select logo1 file", filetypes=[("Image files", "*.png")])
            logo2_path = askopenfilename(title="Select logo2 file", filetypes=[("Image files", "*.png")])

            if not logo1_path or not logo2_path:
                print("Logo not selected. The program will exit.")
            else:
                try:
                    # Leggi i dati dalla cartella "Occurrences"
                    occurrences_df = pd.read_excel(file_path, sheet_name='Occurrences')
                    
                    # Converti la colonna "Data" in stringa
                    occurrences_df['Data'] = occurrences_df['Data'].astype(str)

                    # Prepara i dati per la generazione della pagina HTML
                    scientific_names = occurrences_df['NomeScientifico'].unique()
                    identifiers = {name: list(occurrences_df[occurrences_df['NomeScientifico'] == name]['Identificatore'].unique()) for name in scientific_names}
                    locations = {name: {identifier: list(occurrences_df[(occurrences_df['NomeScientifico'] == name) & (occurrences_df['Identificatore'] == identifier)]['Località'].unique()) for identifier in identifiers[name]} for name in scientific_names}
                    dates = {name: {identifier: {location: list(occurrences_df[(occurrences_df['NomeScientifico'] == name) & (occurrences_df['Identificatore'] == identifier) & (occurrences_df['Località'] == location)]['Data']) for location in locations[name][identifier]} for identifier in identifiers[name]} for name in scientific_names}
                    photos = {name: {identifier: {location: {date: list(occurrences_df[(occurrences_df['NomeScientifico'] == name) & (occurrences_df['Identificatore'] == identifier) & (occurrences_df['Località'] == location) & (occurrences_df['Data'] == date)]['Foto']) for date in dates[name][identifier][location]} for location in locations[name][identifier]} for identifier in identifiers[name]} for name in scientific_names}

                    # Converti i dizionari in JSON
                    identifiers_json = json.dumps(identifiers)
                    locations_json = json.dumps(locations)
                    dates_json = json.dumps(dates)
                    photos_json = json.dumps(photos)

                    # Generare le opzioni per il campo NomeScientifico
                    scientific_names_options = "".join([f"<option value='{name}'>{name}</option>" for name in sorted(scientific_names)])

                    # HTML Template con il menu a discesa "Mode"
                    html_template = """<!DOCTYPE html>
                    <html>
                    <head>
                        <title>Odonata Collection</title>
                        <script type="text/javascript">
                            var identifiers = $identifiers;
                            var locations = $locations;
                            var dates = $dates;
                            var photos = $photos;

                            function toggleFields() {
                                var mode = document.getElementById("mode").value;
                                var scientificNameField = document.getElementById("scientificNameField");
                                var identifierField = document.getElementById("identifierField");
                                var locationField = document.getElementById("locationField");
                                var dateField = document.getElementById("dateField");
                                var photoField = document.getElementById("photoField");

                                if (mode === "general") {
                                    scientificNameField.style.display = "block";
                                    identifierField.style.display = "block";
                                    locationField.style.display = "block";
                                    dateField.style.display = "block";
                                    photoField.style.display = "block";
                                } else if (mode === "scientificName") {
                                    scientificNameField.style.display = "block";
                                    identifierField.style.display = "none";
                                    locationField.style.display = "none";
                                    dateField.style.display = "none";
                                    photoField.style.display = "none";
                                } else if (mode === "identifier") {
                                    scientificNameField.style.display = "none";
                                    identifierField.style.display = "block";
                                    locationField.style.display = "none";
                                    dateField.style.display = "none";
                                    photoField.style.display = "none";
                                // Assicurati che il campo identifier venga popolato
                                populateIdentifiers();
                                }
                            }
function populateIdentifiers() {
    var identifierSelect = document.getElementById("identifier");
    identifierSelect.innerHTML = "<option value=''>Select Identifier</option>";
    
    // Usa un Set per raccogliere identificatori unici
    var uniqueIdentifiers = new Set();
    
    // Aggiungi gli identificatori unici al Set
    for (var scientificName in identifiers) {
        for (var i = 0; i < identifiers[scientificName].length; i++) {
            uniqueIdentifiers.add(identifiers[scientificName][i]);  // Usa add per aggiungere solo valori unici
        }
    }
    
    // Aggiungi ogni identificatore unico come opzione nell'elenco
    uniqueIdentifiers.forEach(function(identifier) {
        var option = document.createElement("option");
        option.value = identifier;
        option.text = identifier;
        identifierSelect.add(option);
    });
}


                            function updateOptions() {
                                var scientificName = document.getElementById("scientificName").value;

                                // Update Identifiers
                                var identifierSelect = document.getElementById("identifier");
                                identifierSelect.innerHTML = "<option value=''>Select Identifier</option>";
                                if (scientificName in identifiers) {
                                    for (var i = 0; i < identifiers[scientificName].length; i++) {
                                        var option = document.createElement("option");
                                        option.value = identifiers[scientificName][i];
                                        option.text = identifiers[scientificName][i];
                                        identifierSelect.add(option);
                                    }
                                }

                                // Clear other fields
                                document.getElementById("location").innerHTML = "<option value=''>Select Locality</option>";
                                document.getElementById("date").innerHTML = "<option value=''>Select Data</option>";
                                document.getElementById("photo").innerHTML = "<option value=''>Select Photo</option>";
                                document.getElementById("message").innerHTML = "";
                            }

                            function updateLocations() {
                                var scientificName = document.getElementById("scientificName").value;
                                var identifier = document.getElementById("identifier").value;

                                // Update Locations
                                var locationSelect = document.getElementById("location");
                                locationSelect.innerHTML = "<option value=''>Select Locality</option>";
                                if (scientificName in locations && identifier in locations[scientificName]) {
                                    for (var i = 0; i < locations[scientificName][identifier].length; i++) {
                                        var option = document.createElement("option");
                                        option.value = locations[scientificName][identifier][i];
                                        option.text = locations[scientificName][identifier][i];
                                        locationSelect.add(option);
                                    }
                                }
                            }

                            function updateDates() {
                                var scientificName = document.getElementById("scientificName").value;
                                var identifier = document.getElementById("identifier").value;
                                var location = document.getElementById("location").value;

                                // Update Dates
                                var dateSelect = document.getElementById("date");
                                dateSelect.innerHTML = "<option value=''>Select Data</option>";
                                if (scientificName in dates && identifier in dates[scientificName] && location in dates[scientificName][identifier]) {
                                    for (var i = 0; i < dates[scientificName][identifier][location].length; i++) {
                                        var dateValue = dates[scientificName][identifier][location][i];
                                        var formattedDate = dateValue.split(' ')[0];  // Remove time part
                                        var option = document.createElement("option");
                                        option.value = dateValue;
                                        option.text = formattedDate;
                                        dateSelect.add(option);
                                    }
                                }
                            }

                            function updatePhotos() {
                                var scientificName = document.getElementById("scientificName").value;
                                var identifier = document.getElementById("identifier").value;
                                var location = document.getElementById("location").value;
                                var date = document.getElementById("date").value;

                                // Update Photos
                                var photoSelect = document.getElementById("photo");
                                photoSelect.innerHTML = "<option value=''>Select Photo</option>";
                                document.getElementById("message").innerHTML = "";
                                if (scientificName in photos && identifier in photos[scientificName] && location in photos[scientificName][identifier] && date in photos[scientificName][identifier][location]) {
                                    for (var i = 0; i < photos[scientificName][identifier][location][date].length; i++) {
                                        var photoValue = photos[scientificName][identifier][location][date][i];
                                        var option = document.createElement("option");
                                        option.value = photoValue;
                                        option.text = photoValue;

                                        // Verifica se la foto è "N/D"
                                        if (photoValue === "N/D") {
                                            document.getElementById("message").innerHTML = "Scheda non disponibile";
                                            option.disabled = true;  // Disabilita l'opzione se è "N/D"
                                        }
                                        photoSelect.add(option);
                                    }
                                }
                            }

function openPhoto() {
    var mode = document.getElementById("mode").value;
    var basePath = "./";  // Questo forza il percorso relativo alla cartella in cui è home.html

    if (mode === "scientificName") {
        var scientificName = document.getElementById("scientificName").value;
        if (scientificName) {
            var htmlFile = basePath + scientificName.toLowerCase().replace(/ /g, "_") + ".html";  
            window.location.href = htmlFile;
        } else {
            alert("Seleziona un nome scientifico.");
        }

    } else if (mode === "identifier") {
        var identifier = document.getElementById("identifier").value;
        if (identifier) {
            var formattedIdentifier = identifier.toLowerCase()
                .replace(/\./g, "")  
                .replace(/, /g, "__")  
                .replace(/ /g, "_")  
                .replace(/'/g, "");  
            var htmlFile = basePath + formattedIdentifier + ".html";
            window.location.href = htmlFile;
        } else {
            alert("Seleziona un identificatore.");
        }

    } else if (mode === "general") {
        var photo = document.getElementById("photo").value;
        if (photo && photo !== "N/D") {
            var htmlFile = basePath + photo.replace('.jpg', '.html');
            window.location.href = htmlFile;
        } else {
            alert("Seleziona una foto valida.");
        }
    } else {
        alert("Seleziona una modalità valida.");
    }
}





                        </script>
                    <style>
                        /* Stili Generali */
                        h1 {
                            color: #821a33;
                            font-size: 60px;
                            text-align: center;
                        }

                        label {
                            color: darkgreen;
                            font-size: 32px;
                        }
			label1 {
                            color: #821a33;
                            font-size: 32px;
                        }

                        select#scientificName {
                            font-style: italic;
                        }
													
                        select#mode {
    			    background-color: #821a33;
    			    color: white;
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

                        .container {
                            text-align: center;
                            padding: 20px;
                        }

                        /* Allineamento Etichette e Selettori */
                        .form-group {
                            display: flex;
                            align-items: center; /* Centra verticalmente il contenuto */
                            justify-content: center; /* Centra orizzontalmente il contenuto */
                            margin-bottom: 20px; /* Spazio tra i gruppi di form */
                        }

                        .form-group label {
                            margin-right: 20px; /* Spazio tra etichetta e selettore */
                        }

                        .form-group select {
                            font-size: 24px; /* Dimensione del testo nei selettori */
                            padding: 8px; /* Aggiunge padding per allineare meglio */
                            border-radius: 4px; /* Arrotonda gli angoli */

                    </style>                        
                    </head>
                    <body>
                        <h1>Odonata Collection</h1>
                        <div class="container">
                            <form>
                                <div class="form-group">
                                    <label1 for="mode"><strong>Mode:</strong></label1>
                                    <select id="mode" name="mode" onchange="toggleFields()">
                                        <option value="general">General</option>
                                        <option value="scientificName">Scientific Name</option>
                                        <option value="identifier">Identifier</option>
                                    </select>
                                </div>

                                <div id="scientificNameField" class="form-group" style="display: block;">
                                    <label for="scientificName"><strong>Scientific Name</strong>:</label>
                                    <select id="scientificName" name="scientificName" onchange="updateOptions()">
                                        <option value="">Select Scientific Name</option>
                                        $scientific_names_options
                                    </select>
                                </div>

                                <div id="identifierField" class="form-group" style="display: block;">
                                    <label for="identifier"><strong>Identifier</strong>:</label>
                                    <select id="identifier" name="identifier" onchange="updateLocations()">
                                        <option value="">Select Identifier</option>
                                    </select>
                                </div>

                                <div id="locationField" class="form-group" style="display: block;">
                                    <label for="location"><strong>Locality</strong>:</label>
                                    <select id="location" name="location" onchange="updateDates()">
                                        <option value="">Select Locality</option>
                                    </select>
                                </div>

                                <div id="dateField" class="form-group" style="display: block;">
                                    <label for="date"><strong>Data</strong>:</label>
                                    <select id="date" name="date" onchange="updatePhotos()">
                                        <option value="">Select Data</option>
                                    </select>
                                </div>

                                <div id="photoField" class="form-group" style="display: block;">
                                    <label for="photo"><strong>Photo</strong>:</label>
                                    <select id="photo" name="photo">
                                        <option value="">Select Photo</option>
                                    </select>
                                </div>

                                <div id="message" style="color: red; font-size: 18px;"></div>
                                <button type="button" onclick="openPhoto()">Open Record</button>
                                <p>
                                <button type="button" onclick="window.location.href='catalogo_gbif.html'">Go to GBIF Catalog</button>
                            </p>
                            <div class="logos">
                                <img src="$logo1" alt="Logo 1"/>
                                <img src="$logo2" alt="Logo 2"/>
                            </div>   
                            </form>
                        </div>
                    </body>
                    </html>"""

                    # Scrivi il file HTML
                    output_file = os.path.join(save_folder_path, "home.html")
                    html_template = html_template.replace('$identifiers', identifiers_json) \
                        .replace('$locations', locations_json) \
                        .replace('$dates', dates_json) \
                        .replace('$photos', photos_json) \
                        .replace('$scientific_names_options', scientific_names_options) \
                        .replace('$html_folder', "./") \
                        .replace('$logo1', os.path.basename(logo1_path)) \
                        .replace('$logo2', os.path.basename(logo2_path))

                    with open(output_file, 'w', encoding="utf-8") as file:
                        file.write(html_template)
                    print(f"File HTML creato: {output_file}")

                except Exception as e:
                    print(f"Si è verificato un errore: {e}")
