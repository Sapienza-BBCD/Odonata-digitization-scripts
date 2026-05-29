import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
import json
import os

# Create a hidden window for file selection
Tk().withdraw()

# Prompt the user to select the Excel file
file_path = askopenfilename(
    title="Select the Excel file",
    filetypes=[("Excel files", "*.xlsx")]
)

if not file_path:
    print("No file selected. The program will exit.")

else:
    # Select the folder containing HTML files
    html_folder_path = askdirectory(
        title="Select the folder containing HTML files"
    )

    if not html_folder_path:
        print("No folder selected. The program will exit.")

    else:
        # Select the folder where the HTML file will be saved
        save_folder_path = askdirectory(
            title="Select the folder for saving the HTML file"
        )

        if not save_folder_path:
            print("No output folder selected. The program will exit.")

        else:
            # Select logo files
            logo1_path = askopenfilename(
                title="Select logo1 file",
                filetypes=[("Image files", "*.png")]
            )

            logo2_path = askopenfilename(
                title="Select logo2 file",
                filetypes=[("Image files", "*.png")]
            )

            if not logo1_path or not logo2_path:
                print("Logo files not selected. The program will exit.")

            else:
                try:
                    # Read data from the "Occurrences" worksheet
                    occurrences_df = pd.read_excel(
                        file_path,
                        sheet_name='Occurrences'
                    )

                    # Convert the "EventDate" column to string format
                    occurrences_df['EventDate'] = occurrences_df['EventDate'].astype(str)

                    # Prepare data for HTML page generation
                    scientific_names = occurrences_df['ScientificName'].unique()

                    identifiers = {
                        name: list(
                            occurrences_df[
                                occurrences_df['ScientificName'] == name
                            ]['Identifier'].unique()
                        )
                        for name in scientific_names
                    }

                    localities = {
                        name: {
                            identifier: list(
                                occurrences_df[
                                    (occurrences_df['ScientificName'] == name) &
                                    (occurrences_df['Identifier'] == identifier)
                                ]['Locality'].unique()
                            )
                            for identifier in identifiers[name]
                        }
                        for name in scientific_names
                    }

                    dates = {
                        name: {
                            identifier: {
                                locality: list(
                                    occurrences_df[
                                        (occurrences_df['ScientificName'] == name) &
                                        (occurrences_df['Identifier'] == identifier) &
                                        (occurrences_df['Locality'] == locality)
                                    ]['EventDate']
                                )
                                for locality in localities[name][identifier]
                            }
                            for identifier in identifiers[name]
                        }
                        for name in scientific_names
                    }

                    photos = {
                        name: {
                            identifier: {
                                locality: {
                                    date: list(
                                        occurrences_df[
                                            (occurrences_df['ScientificName'] == name) &
                                            (occurrences_df['Identifier'] == identifier) &
                                            (occurrences_df['Locality'] == locality) &
                                            (occurrences_df['EventDate'] == date)
                                        ]['Photo']
                                    )
                                    for date in dates[name][identifier][locality]
                                }
                                for locality in localities[name][identifier]
                            }
                            for identifier in identifiers[name]
                        }
                        for name in scientific_names
                    }

                    # Convert dictionaries to JSON format
                    identifiers_json = json.dumps(identifiers)
                    localities_json = json.dumps(localities)
                    dates_json = json.dumps(dates)
                    photos_json = json.dumps(photos)

                    # Generate options for the Scientific Name dropdown menu
                    scientific_names_options = "".join([
                        f"<option value='{name}'>{name}</option>"
                        for name in sorted(scientific_names)
                    ])

                    # HTML template
                    html_template = """<!DOCTYPE html>
<html>
<head>
    <title>Odonata Collection</title>

    <script type="text/javascript">

        var identifiers = $identifiers;
        var localities = $localities;
        var dates = $dates;
        var photos = $photos;

        function toggleFields() {

            var mode = document.getElementById("mode").value;

            var scientificNameField = document.getElementById("scientificNameField");
            var identifierField = document.getElementById("identifierField");
            var localityField = document.getElementById("localityField");
            var dateField = document.getElementById("dateField");
            var photoField = document.getElementById("photoField");

            if (mode === "general") {

                scientificNameField.style.display = "block";
                identifierField.style.display = "block";
                localityField.style.display = "block";
                dateField.style.display = "block";
                photoField.style.display = "block";

            } else if (mode === "scientificName") {

                scientificNameField.style.display = "block";
                identifierField.style.display = "none";
                localityField.style.display = "none";
                dateField.style.display = "none";
                photoField.style.display = "none";

            } else if (mode === "identifier") {

                scientificNameField.style.display = "none";
                identifierField.style.display = "block";
                localityField.style.display = "none";
                dateField.style.display = "none";
                photoField.style.display = "none";

                // Ensure that the identifier field is populated
                populateIdentifiers();
            }
        }

        function populateIdentifiers() {

            var identifierSelect = document.getElementById("identifier");

            identifierSelect.innerHTML =
                "<option value=''>Select Identifier</option>";

            // Use a Set to collect unique identifiers
            var uniqueIdentifiers = new Set();

            // Add unique identifiers to the Set
            for (var scientificName in identifiers) {

                for (var i = 0; i < identifiers[scientificName].length; i++) {

                    uniqueIdentifiers.add(
                        identifiers[scientificName][i]
                    );
                }
            }

            // Add each unique identifier as an option
            uniqueIdentifiers.forEach(function(identifier) {

                var option = document.createElement("option");

                option.value = identifier;
                option.text = identifier;

                identifierSelect.add(option);
            });
        }

        function updateOptions() {

            var scientificName =
                document.getElementById("scientificName").value;

            var identifierSelect =
                document.getElementById("identifier");

            identifierSelect.innerHTML =
                "<option value=''>Select Identifier</option>";

            if (scientificName in identifiers) {

                for (var i = 0;
                     i < identifiers[scientificName].length;
                     i++) {

                    var option = document.createElement("option");

                    option.value = identifiers[scientificName][i];
                    option.text = identifiers[scientificName][i];

                    identifierSelect.add(option);
                }
            }

            document.getElementById("locality").innerHTML =
                "<option value=''>Select Locality</option>";

            document.getElementById("date").innerHTML =
                "<option value=''>Select Event Date</option>";

            document.getElementById("photo").innerHTML =
                "<option value=''>Select Photo</option>";

            document.getElementById("message").innerHTML = "";
        }

    </script>

</head>

<body>

<h1>Odonata Collection</h1>

</body>
</html>
"""

                    # Write the HTML file
                    output_file = os.path.join(
                        save_folder_path,
                        "home.html"
                    )

                    html_template = html_template.replace(
                        '$identifiers',
                        identifiers_json
                    ).replace(
                        '$localities',
                        localities_json
                    ).replace(
                        '$dates',
                        dates_json
                    ).replace(
                        '$photos',
                        photos_json
                    ).replace(
                        '$scientific_names_options',
                        scientific_names_options
                    ).replace(
                        '$logo1',
                        os.path.basename(logo1_path)
                    ).replace(
                        '$logo2',
                        os.path.basename(logo2_path)
                    )

                    with open(output_file, 'w', encoding="utf-8") as file:
                        file.write(html_template)

                    print(f"HTML file created: {output_file}")

                except Exception as e:
                    print(f"An error occurred: {e}")
