"""
Generate the main collection index page from a Darwin Core dataset.

This version preserves the Darwin Core field mapping and restores
the complete dynamic search interface from the original working script.

Search modes:
- General: Scientific Name -> Identifier -> Locality -> Event Date -> Photo
- Scientific Name
- Identifier
"""

import html
import json
import os
import re
import shutil

import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory

# ============================================================
# DARWIN CORE CONFIGURATION
# ============================================================

SHEET_NAME = "Occurrences"

SCIENTIFIC_NAME_COLUMN = "scientificName"

# IMPORTANT:
# The old script used "Identificatore". In this adapted version,
# "identifiedBy" is retained because it was used in the referee-modified code.
# If your intended "Identifier" is instead catalogNumber or occurrenceID,
# change this constant accordingly.
IDENTIFIER_COLUMN = "identifiedBy"

LOCALITY_COLUMN = "locality"
EVENT_DATE_COLUMN = "eventDate"
PHOTO_COLUMN = "associatedMedia"
RECORD_PAGE_COLUMN = "_recordPage"

REQUIRED_COLUMNS = [
    SCIENTIFIC_NAME_COLUMN,
    IDENTIFIER_COLUMN,
    LOCALITY_COLUMN,
    EVENT_DATE_COLUMN,
    PHOTO_COLUMN,
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_text(value):
    """Convert missing values to an empty string and trim whitespace."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def validate_columns(dataframe):
    """Check that all required Darwin Core columns are present."""
    missing_columns = [
        column for column in REQUIRED_COLUMNS
        if column not in dataframe.columns
    ]
    if missing_columns:
        raise ValueError(
            "Missing Darwin Core columns:\n\n"
            + "\n".join(f"• {column}" for column in missing_columns)
        )


def unique_non_empty(values):
    """Return sorted, unique, non-empty values."""
    cleaned = {
        normalize_text(value)
        for value in values
        if normalize_text(value)
    }
    return sorted(cleaned, key=lambda value: value.casefold())


def copy_logo(source_path, destination_folder, output_name):
    """Copy a logo to the destination folder with overwrite confirmation."""
    if not source_path or not os.path.isfile(source_path):
        raise FileNotFoundError(f"Logo file not found: {source_path}")

    destination_path = os.path.join(destination_folder, output_name)

    if os.path.abspath(source_path) == os.path.abspath(destination_path):
        print(f"ℹ️ Logo already in destination: {destination_path}")
        return

    if os.path.exists(destination_path):
        overwrite = messagebox.askyesno(
            "Existing file",
            f"The file '{output_name}' already exists.\nDo you want to overwrite it?"
        )
        if not overwrite:
            print(f"ℹ️ Existing logo retained: {destination_path}")
            return

    shutil.copy2(source_path, destination_path)
    print(f"✅ Logo copied: {destination_path}")


def first_media_name(value):
    """
    Return the first associatedMedia item.
    Supports values separated by | or ; and strips URL/path components.
    """
    value = normalize_text(value)
    if not value:
        return ""

    first_item = re.split(r"[|;]", value)[0].strip()
    if not first_item:
        return ""

    # Strip query/fragment for URLs and retain basename.
    first_item = first_item.split("?", 1)[0].split("#", 1)[0]
    return os.path.basename(first_item.replace("\\", "/"))


def build_photo_record(media_value, page_filename):
    """
    Build the photo entry used by the General search interface.

    The visible label is derived from associatedMedia (for example,
    DSC_2527.jpg), while the option value points to the corresponding
    progressively numbered specimen page (for example, 54.html).

    RECORD_PAGE_COLUMN is an internal workflow field and is not exported
    as a Darwin Core term.
    """
    media_name = first_media_name(media_value)
    if not media_name:
        media_name = normalize_text(media_value)

    return {
        "label": media_name,
        "page": normalize_text(page_filename),
    }



def safe_identifier_filename(value):
    """
    Create the standardized filename used for identifier summary pages.

    Multiple identifiers separated by "|" are normalized individually
    and joined with a double underscore.
    """
    value = normalize_text(value)

    parts = [
        part.strip()
        for part in value.split("|")
        if part.strip()
    ]

    cleaned_parts = []

    for part in parts:
        part = part.replace(".", "")
        part = part.replace("'", "")
        part = re.sub(r"\s+", "_", part.strip())
        part = re.sub(r"_+", "_", part)
        part = part.strip("_").lower()

        if part:
            cleaned_parts.append(part)

    return "__".join(cleaned_parts)



def safe_scientific_name_filename(value):
    """
    Create the standardized filename used for taxon summary pages.

    The same normalization rule must be used by both the Python generators
    and the JavaScript navigation code.
    """
    value = normalize_text(value)
    value = re.sub(r"[^a-zA-Z0-9_-]", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_").lower()


def build_nested_data(dataframe):
    """
    Build dictionaries required by the dynamic HTML interface.
    """
    scientific_names = unique_non_empty(dataframe[SCIENTIFIC_NAME_COLUMN])

    identifiers = {}
    for scientific_name in scientific_names:
        subset = dataframe[
            dataframe[SCIENTIFIC_NAME_COLUMN] == scientific_name
        ]
        identifiers[scientific_name] = unique_non_empty(
            subset[IDENTIFIER_COLUMN]
        )

    localities = {}
    for scientific_name, identifier_values in identifiers.items():
        localities[scientific_name] = {}
        for identifier in identifier_values:
            subset = dataframe[
                (dataframe[SCIENTIFIC_NAME_COLUMN] == scientific_name)
                & (dataframe[IDENTIFIER_COLUMN] == identifier)
            ]
            localities[scientific_name][identifier] = unique_non_empty(
                subset[LOCALITY_COLUMN]
            )

    dates = {}
    for scientific_name, identifier_map in localities.items():
        dates[scientific_name] = {}
        for identifier, locality_values in identifier_map.items():
            dates[scientific_name][identifier] = {}
            for locality in locality_values:
                subset = dataframe[
                    (dataframe[SCIENTIFIC_NAME_COLUMN] == scientific_name)
                    & (dataframe[IDENTIFIER_COLUMN] == identifier)
                    & (dataframe[LOCALITY_COLUMN] == locality)
                ]
                dates[scientific_name][identifier][locality] = unique_non_empty(
                    subset[EVENT_DATE_COLUMN]
                )

    photos = {}
    for scientific_name, identifier_map in dates.items():
        photos[scientific_name] = {}
        for identifier, locality_map in identifier_map.items():
            photos[scientific_name][identifier] = {}
            for locality, date_values in locality_map.items():
                photos[scientific_name][identifier][locality] = {}
                for event_date in date_values:
                    subset = dataframe[
                        (dataframe[SCIENTIFIC_NAME_COLUMN] == scientific_name)
                        & (dataframe[IDENTIFIER_COLUMN] == identifier)
                        & (dataframe[LOCALITY_COLUMN] == locality)
                        & (dataframe[EVENT_DATE_COLUMN] == event_date)
                    ]

                    photo_records = []
                    seen_entries = set()

                    for _, row in subset.iterrows():
                        media_value = normalize_text(row[PHOTO_COLUMN])
                        page_filename = normalize_text(row[RECORD_PAGE_COLUMN])

                        if not media_value or not page_filename:
                            continue

                        photo_record = build_photo_record(
                            media_value,
                            page_filename,
                        )

                        entry_key = (
                            photo_record["label"],
                            photo_record["page"],
                        )

                        if entry_key not in seen_entries:
                            photo_records.append(photo_record)
                            seen_entries.add(entry_key)

                    photo_records.sort(
                        key=lambda item: item["label"].casefold()
                    )

                    photos[scientific_name][identifier][locality][event_date] = (
                        photo_records
                    )

    return scientific_names, identifiers, localities, dates, photos


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    root = Tk()
    root.withdraw()

    try:
        file_path = askopenfilename(
            title="Select the Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            print("No file selected. The program will exit.")
            return

        html_folder_path = askdirectory(
            title="Select the folder containing HTML record files"
        )
        if not html_folder_path:
            print("No HTML folder selected. The program will exit.")
            return

        save_folder_path = askdirectory(
            title="Select the folder for saving index.html"
        )
        if not save_folder_path:
            print("No output folder selected. The program will exit.")
            return

        logo1_path = askopenfilename(
            title="Select logo1 file",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.webp")]
        )
        logo2_path = askopenfilename(
            title="Select logo2 file",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.webp")]
        )
        if not logo1_path or not logo2_path:
            print("Logo files not selected. The program will exit.")
            return

        occurrences_df = pd.read_excel(
            file_path,
            sheet_name=SHEET_NAME,
            engine="openpyxl",
            dtype=object
        )

        occurrences_df.columns = [
            str(column).strip()
            for column in occurrences_df.columns
        ]

        validate_columns(occurrences_df)

        occurrences_df = occurrences_df[REQUIRED_COLUMNS].copy()

        for column in REQUIRED_COLUMNS:
            occurrences_df[column] = occurrences_df[column].apply(normalize_text)

        # Preserve the same progressive numbering used by the specimen-page
        # generator: the first spreadsheet record corresponds to 1.html,
        # the second to 2.html, and so on.
        occurrences_df[RECORD_PAGE_COLUMN] = [
            f"{record_number}.html"
            for record_number in range(1, len(occurrences_df) + 1)
        ]

        occurrences_df = occurrences_df[
            occurrences_df[SCIENTIFIC_NAME_COLUMN] != ""
        ].copy()

        if occurrences_df.empty:
            raise ValueError(
                "No valid records were found after cleaning the Excel file."
            )

        (
            scientific_names,
            identifiers,
            localities,
            dates,
            photos,
        ) = build_nested_data(occurrences_df)

        identifiers_json = json.dumps(identifiers, ensure_ascii=False)
        localities_json = json.dumps(localities, ensure_ascii=False)
        dates_json = json.dumps(dates, ensure_ascii=False)
        photos_json = json.dumps(photos, ensure_ascii=False)

        scientific_names_options = "".join(
            (
                f'<option value="{html.escape(name, quote=True)}">'
                f'{html.escape(name)}</option>'
            )
            for name in scientific_names
        )

        html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Odonata Collection</title>

    <script type="text/javascript">
        var identifiers = $identifiers;
        var localities = $localities;
        var dates = $dates;
        var photos = $photos;

        function resetSelect(id, placeholder) {
            document.getElementById(id).innerHTML =
                "<option value=''>" + placeholder + "</option>";
        }

        function toggleFields() {
            var mode = document.getElementById("mode").value;
            var scientificNameField = document.getElementById("scientificNameField");
            var identifierField = document.getElementById("identifierField");
            var localityField = document.getElementById("localityField");
            var dateField = document.getElementById("dateField");
            var photoField = document.getElementById("photoField");

            if (mode === "general") {
                scientificNameField.style.display = "flex";
                identifierField.style.display = "flex";
                localityField.style.display = "flex";
                dateField.style.display = "flex";
                photoField.style.display = "flex";
                resetSelect("identifier", "Select Identifier");
                resetSelect("locality", "Select Locality");
                resetSelect("date", "Select Event Date");
                resetSelect("photo", "Select Photo");

            } else if (mode === "scientificName") {
                scientificNameField.style.display = "flex";
                identifierField.style.display = "none";
                localityField.style.display = "none";
                dateField.style.display = "none";
                photoField.style.display = "none";

            } else if (mode === "identifier") {
                scientificNameField.style.display = "none";
                identifierField.style.display = "flex";
                localityField.style.display = "none";
                dateField.style.display = "none";
                photoField.style.display = "none";
                populateIdentifiers();
            }
        }

        function populateIdentifiers() {
            var identifierSelect = document.getElementById("identifier");
            identifierSelect.innerHTML = "<option value=''>Select Identifier</option>";

            var uniqueIdentifiers = new Set();

            for (var scientificName in identifiers) {
                for (var i = 0; i < identifiers[scientificName].length; i++) {
                    uniqueIdentifiers.add(identifiers[scientificName][i]);
                }
            }

            Array.from(uniqueIdentifiers)
                .sort(function(a, b) { return a.localeCompare(b); })
                .forEach(function(identifier) {
                    var option = document.createElement("option");
                    option.value = identifier;
                    option.text = identifier;
                    identifierSelect.add(option);
                });
        }

        function updateOptions() {
            var scientificName = document.getElementById("scientificName").value;
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

            resetSelect("locality", "Select Locality");
            resetSelect("date", "Select Event Date");
            resetSelect("photo", "Select Photo");
            document.getElementById("message").innerHTML = "";
        }

        function updateLocalities() {
            var scientificName = document.getElementById("scientificName").value;
            var identifier = document.getElementById("identifier").value;
            var localitySelect = document.getElementById("locality");

            localitySelect.innerHTML = "<option value=''>Select Locality</option>";

            if (
                scientificName in localities &&
                identifier in localities[scientificName]
            ) {
                var values = localities[scientificName][identifier];

                for (var i = 0; i < values.length; i++) {
                    var option = document.createElement("option");
                    option.value = values[i];
                    option.text = values[i];
                    localitySelect.add(option);
                }
            }

            resetSelect("date", "Select Event Date");
            resetSelect("photo", "Select Photo");
        }

        function updateDates() {
            var scientificName = document.getElementById("scientificName").value;
            var identifier = document.getElementById("identifier").value;
            var locality = document.getElementById("locality").value;
            var dateSelect = document.getElementById("date");

            dateSelect.innerHTML = "<option value=''>Select Event Date</option>";

            if (
                scientificName in dates &&
                identifier in dates[scientificName] &&
                locality in dates[scientificName][identifier]
            ) {
                var values = dates[scientificName][identifier][locality];

                for (var i = 0; i < values.length; i++) {
                    var dateValue = values[i];
                    var option = document.createElement("option");
                    option.value = dateValue;
                    option.text = dateValue.split(" ")[0];
                    dateSelect.add(option);
                }
            }

            resetSelect("photo", "Select Photo");
        }

        function updatePhotos() {
            var scientificName = document.getElementById("scientificName").value;
            var identifier = document.getElementById("identifier").value;
            var locality = document.getElementById("locality").value;
            var date = document.getElementById("date").value;
            var photoSelect = document.getElementById("photo");

            photoSelect.innerHTML = "<option value=''>Select Photo</option>";
            document.getElementById("message").innerHTML = "";

            if (
                scientificName in photos &&
                identifier in photos[scientificName] &&
                locality in photos[scientificName][identifier] &&
                date in photos[scientificName][identifier][locality]
            ) {
                var values = photos[scientificName][identifier][locality][date];

                for (var i = 0; i < values.length; i++) {
                    var photoRecord = values[i];
                    var option = document.createElement("option");

                    // Display the image filename, but open the progressively
                    // numbered specimen HTML page associated with that record.
                    option.value = photoRecord.page;
                    option.text = photoRecord.label;

                    if (
                        photoRecord.label === "N/D" ||
                        photoRecord.label.toLowerCase() === "not available"
                    ) {
                        document.getElementById("message").innerHTML =
                            "Record not available";
                        option.disabled = true;
                    }

                    photoSelect.add(option);
                }
            }
        }


        function sanitizeScientificNameFilename(value) {
            /*
             * Create exactly the same taxon-page filename convention used
             * by the Python species-page generator.
             *
             * Examples:
             * Aeshna mixta (Latreille, 1805)
             *     -> aeshna_mixta_latreille_1805.html
             *
             * Anax imperator Leach, 1815
             *     -> anax_imperator_leach_1815.html
             *
             * Cordulegaster
             *     -> cordulegaster.html
             */
            return value
                .trim()
                .replace(/[^a-zA-Z0-9_-]/g, "_")
                .replace(/_+/g, "_")
                .replace(/^_+|_+$/g, "")
                .toLowerCase();
        }


        function sanitizeIdentifierFilename(value) {
            /*
             * Use the same identifier-page filename convention as the
             * identifier-page generator.
             *
             * A. Baldi | E. Avellinese | G. Casini
             * -> a_baldi__e_avellinese__g_casini.html
             */
            return value
                .split("|")
                .map(function(part) {
                    return part
                        .trim()
                        .replace(/\./g, "")
                        .replace(/'/g, "")
                        .replace(/\s+/g, "_")
                        .replace(/_+/g, "_")
                        .replace(/^_+|_+$/g, "")
                        .toLowerCase();
                })
                .filter(function(part) {
                    return part.length > 0;
                })
                .join("__");
        }


        function openRecord() {
            var mode = document.getElementById("mode").value;
            var basePath = "./";

            saveInterfaceState();

            if (mode === "scientificName") {
                var scientificName =
                    document.getElementById("scientificName").value;

                if (scientificName) {
                    var htmlFile =
                        basePath +
                        sanitizeScientificNameFilename(scientificName) +
                        ".html";

                    window.location.href = htmlFile;
                } else {
                    alert("Select a scientific name.");
                }

            } else if (mode === "identifier") {
                var identifier =
                    document.getElementById("identifier").value;

                if (identifier) {
                    var formattedIdentifier =
                        sanitizeIdentifierFilename(identifier);

                    window.location.href =
                        basePath + formattedIdentifier + ".html";
                } else {
                    alert("Select an identifier.");
                }

            } else if (mode === "general") {
                var specimenPage = document.getElementById("photo").value;

                if (specimenPage) {
                    window.location.href = basePath + specimenPage;
                } else {
                    alert("Select a valid photo.");
                }

            } else {
                alert("Select a valid mode.");
            }
        }

        function saveInterfaceState() {
            var state = {
                mode: document.getElementById("mode").value,
                scientificName: document.getElementById("scientificName").value,
                identifier: document.getElementById("identifier").value,
                locality: document.getElementById("locality").value,
                date: document.getElementById("date").value,
                photo: document.getElementById("photo").value
            };

            sessionStorage.setItem(
                "odonataCollectionSearchState",
                JSON.stringify(state)
            );
        }

        function restoreInterfaceState() {
            var storedState = sessionStorage.getItem(
                "odonataCollectionSearchState"
            );

            if (!storedState) {
                toggleFields();
                return;
            }

            var state;

            try {
                state = JSON.parse(storedState);
            } catch (error) {
                toggleFields();
                return;
            }

            document.getElementById("mode").value =
                state.mode || "general";

            toggleFields();

            if (state.mode === "identifier") {
                populateIdentifiers();

                if (state.identifier) {
                    document.getElementById("identifier").value =
                        state.identifier;
                }

                return;
            }

            if (state.scientificName) {
                document.getElementById("scientificName").value =
                    state.scientificName;
            }

            if (state.mode === "scientificName") {
                return;
            }

            updateOptions();

            if (state.identifier) {
                document.getElementById("identifier").value =
                    state.identifier;
            }

            updateLocalities();

            if (state.locality) {
                document.getElementById("locality").value =
                    state.locality;
            }

            updateDates();

            if (state.date) {
                document.getElementById("date").value =
                    state.date;
            }

            updatePhotos();

            if (state.photo) {
                document.getElementById("photo").value =
                    state.photo;
            }
        }

        window.addEventListener("DOMContentLoaded", function() {
            toggleFields();

            [
                "mode",
                "scientificName",
                "identifier",
                "locality",
                "date",
                "photo"
            ].forEach(function(elementId) {
                document.getElementById(elementId).addEventListener(
                    "change",
                    saveInterfaceState
                );
            });
        });

        window.addEventListener("pageshow", function(event) {
            var navigationEntries =
                performance.getEntriesByType("navigation");

            var isBackForward =
                event.persisted ||
                (
                    navigationEntries.length > 0 &&
                    navigationEntries[0].type === "back_forward"
                );

            if (isBackForward) {
                restoreInterfaceState();
            }
        });
    </script>

    <style>
        html,
        body {
            margin: 0;
            padding: 0;
            height: 100%;
            font-family: Arial, Helvetica, sans-serif;
            background: white;
            color: #222;
            overflow: hidden;
        }

        /*
         * Fixed upper band containing the collection title and Mode selector.
         * Only the central search area is allowed to scroll.
         */
        .fixed-header {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
            background: white;
            border-bottom: 1px solid #ddd;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            padding: 14px 20px 12px;
        }

        h1 {
            color: #821a33;
            font-size: 54px;
            text-align: center;
            margin: 0 0 12px;
        }

        .header-mode {
            margin-bottom: 0;
        }

        /*
         * Scrollable central area between the fixed header and footer.
         */
        .scroll-area {
            position: fixed;
            top: 150px;
            bottom: 105px;
            left: 0;
            right: 0;
            overflow-y: auto;
            overflow-x: hidden;
            -webkit-overflow-scrolling: touch;
        }

        .container {
            width: min(1000px, 92%);
            margin: 0 auto;
            text-align: center;
            padding: 28px 20px 40px;
            box-sizing: border-box;
        }

        .form-group {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 20px;
            margin-bottom: 20px;
        }

        .form-group label,
        .mode-label {
            min-width: 240px;
            text-align: right;
            font-size: 30px;
            font-weight: bold;
        }

        .form-group label {
            color: darkgreen;
        }

        .mode-label {
            color: #821a33;
        }

        select {
            font-size: 22px;
            padding: 9px 12px;
            border-radius: 5px;
            min-width: 340px;
            max-width: 55vw;
        }

        select#scientificName {
            font-style: italic;
        }

        select#mode {
            background-color: #821a33;
            color: white;
        }

        button {
            font-size: 22px;
            background-color: #821a33;
            color: white;
            padding: 13px 22px;
            border: none;
            cursor: pointer;
            border-radius: 8px;
            margin: 8px;
        }

        button:hover {
            background-color: #c54c00;
        }

        /*
         * Fixed lower band containing the two institutional logos.
         */
        .fixed-footer {
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            z-index: 1000;
            background: white;
            border-top: 1px solid #ddd;
            box-shadow: 0 -2px 8px rgba(0, 0, 0, 0.06);
            padding: 10px 24px;
        }

        .logos {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 30px;
            width: 100%;
            max-width: 1200px;
            margin: 0 auto;
        }

        .logos img {
            width: 178.5px;
            max-height: 72px;
            object-fit: contain;
        }

        #message {
            color: red;
            font-size: 18px;
            min-height: 24px;
            margin: 10px 0;
        }

        @media (max-width: 700px) {
            .fixed-header {
                padding: 10px 12px;
            }

            h1 {
                font-size: 38px;
                margin-bottom: 8px;
            }

            .scroll-area {
                top: 135px;
                bottom: 90px;
            }

            .fixed-footer {
                padding: 8px 14px;
            }

            .logos img {
                width: 130px;
                max-height: 60px;
            }

            .form-group {
                flex-direction: column;
                gap: 8px;
            }

            .form-group label,
            .mode-label {
                min-width: 0;
                text-align: center;
                font-size: 24px;
            }

            select {
                min-width: min(90vw, 420px);
            }
        }
    </style>
</head>

<body>

    <!-- Fixed upper band: collection title and Mode selector -->
    <header class="fixed-header">
        <h1>Odonata Collection</h1>

        <div class="form-group header-mode">
            <label class="mode-label" for="mode">Mode:</label>
            <select id="mode" name="mode" onchange="toggleFields()">
                <option value="general">General</option>
                <option value="scientificName">Scientific Name</option>
                <option value="identifier">Identifier</option>
            </select>
        </div>
    </header>

    <!-- Scrollable central area -->
    <main class="scroll-area">
        <div class="container">
            <form onsubmit="return false;">

            <div id="scientificNameField" class="form-group">
                <label for="scientificName">Scientific Name:</label>
                <select
                    id="scientificName"
                    name="scientificName"
                    onchange="updateOptions()"
                >
                    <option value="">Select Scientific Name</option>
                    $scientific_names_options
                </select>
            </div>

            <div id="identifierField" class="form-group">
                <label for="identifier">Identifier:</label>
                <select
                    id="identifier"
                    name="identifier"
                    onchange="updateLocalities()"
                >
                    <option value="">Select Identifier</option>
                </select>
            </div>

            <div id="localityField" class="form-group">
                <label for="locality">Locality:</label>
                <select
                    id="locality"
                    name="locality"
                    onchange="updateDates()"
                >
                    <option value="">Select Locality</option>
                </select>
            </div>

            <div id="dateField" class="form-group">
                <label for="date">Event Date:</label>
                <select
                    id="date"
                    name="date"
                    onchange="updatePhotos()"
                >
                    <option value="">Select Event Date</option>
                </select>
            </div>

            <div id="photoField" class="form-group">
                <label for="photo">Photo:</label>
                <select id="photo" name="photo">
                    <option value="">Select Photo</option>
                </select>
            </div>

            <div id="message"></div>

            <button type="button" onclick="openRecord()">
                Open Record
            </button>

            <p>
                <button
                    type="button"
                    onclick="window.location.href='catalogo_gbif.html'"
                >
                    Search by Catalog Number
                </button>
            </p>

            </form>
        </div>
    </main>

    <!-- Fixed lower band: institutional logos -->
    <footer class="fixed-footer">
        <div class="logos">
            <img src="$logo1" alt="Logo 1">
            <img src="$logo2" alt="Logo 2">
        </div>
    </footer>

</body>
</html>
"""

        os.makedirs(save_folder_path, exist_ok=True)

        copy_logo(
            logo1_path,
            save_folder_path,
            os.path.basename(logo1_path)
        )
        copy_logo(
            logo2_path,
            save_folder_path,
            os.path.basename(logo2_path)
        )

        output_file = os.path.join(
            save_folder_path,
            "index.html"
        )

        html_template = (
            html_template
            .replace("$identifiers", identifiers_json)
            .replace("$localities", localities_json)
            .replace("$dates", dates_json)
            .replace("$photos", photos_json)
            .replace("$scientific_names_options", scientific_names_options)
            .replace("$logo1", os.path.basename(logo1_path))
            .replace("$logo2", os.path.basename(logo2_path))
        )

        with open(output_file, "w", encoding="utf-8") as file:
            file.write(html_template)

        print(f"✅ HTML file created: {output_file}")
        print(f"✅ Records processed: {len(occurrences_df)}")
        print(f"✅ Scientific names found: {len(scientific_names)}")
        print(f"ℹ️ HTML pages folder selected: {html_folder_path}")

        messagebox.showinfo(
            "Operation completed",
            (
                "index.html was created successfully.\n\n"
                f"Records processed: {len(occurrences_df)}\n"
                f"Scientific names: {len(scientific_names)}\n\n"
                f"Output file:\n{output_file}"
            )
        )

    except Exception as error:
        print(f"❌ An error occurred: {error}")
        messagebox.showerror("Error", str(error))

    finally:
        root.destroy()


if __name__ == "__main__":
    main()
