"""
Generate specimen HTML pages from a Darwin Core dataset.

The script reads a Darwin Core Excel file, lets the user choose
which metadata fields to display, uses associatedMedia directly for specimen images, copies only shared assets,
and creates one HTML page
for each occurrence record.
"""

import os
import re
import shutil
import html
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from jinja2 import Environment, select_autoescape

from tkinter import (
    Tk,
    Toplevel,
    Canvas,
    Frame,
    Checkbutton,
    Button,
    BooleanVar,
    Label,
    messagebox
)
from tkinter.filedialog import askopenfilename, askdirectory


# ============================================================
# DARWIN CORE CONFIGURATION
# ============================================================

# Darwin Core column containing the image filename or URL.
IMAGE_COLUMN = "associatedMedia"

# Fields preselected in the column selection dialog.
DEFAULT_SELECTED_COLUMNS = [
    "occurrenceID",
    "basisOfRecord",
    "eventDate",
    "scientificNameAuthorship",
    "scientificName",
    "vernacularName",
    "kingdom",
    "phylum",
    "class",
    "order",
    "family",
    "genus",
    "specificEpithet",
    "taxonRank",
    "identifiedBy",
    "decimalLatitude",
    "decimalLongitude",
    "coordinateUncertaintyInMeters",
    "continent",
    "country",
    "countryCode",
    "stateProvince",
    "locality",
    "institutionID",
    "collectionCode",
    "catalogNumber",
    "sex",
    "individualCount",
    "samplingProtocol",
    "preparations",
    "lifeStage",
]

# Fields displayed in italics in the HTML table.
ITALIC_COLUMNS = {
    "scientificName",
    "genus",
    "specificEpithet",
}

# Human-readable labels displayed in the HTML table.
FIELD_LABELS = {
    "occurrenceID": "Occurrence ID",
    "basisOfRecord": "Basis of record",
    "eventDate": "Event date",
    "startDayOfYear": "Start day of year",
    "endDayOfYear": "End day of year",
    "verbatimEventDate": "Original event date",
    "scientificNameAuthorship": "Scientific name authorship",
    "scientificName": "Scientific name",
    "vernacularName": "Vernacular name",
    "higherClassification": "Higher classification",
    "kingdom": "Kingdom",
    "phylum": "Phylum",
    "class": "Class",
    "order": "Order",
    "family": "Family",
    "genus": "Genus",
    "specificEpithet": "Specific epithet",
    "taxonRank": "Taxon rank",
    "identifiedBy": "Identified by",
    "nomenclaturalCode": "Nomenclatural code",
    "decimalLatitude": "Decimal latitude",
    "decimalLongitude": "Decimal longitude",
    "geodeticDatum": "Geodetic datum",
    "coordinateUncertaintyInMeters": "Coordinate uncertainty",
    "verbatimCoordinates": "Original coordinates",
    "verbatimCoordinateSystem": "Coordinate system",
    "georeferencedDate": "Georeferenced date",
    "georeferenceProtocol": "Georeference protocol",
    "georeferenceSources": "Georeference sources",
    "georeferenceVerificationStatus": "Georeference verification status",
    "higherGeography": "Higher geography",
    "continent": "Continent",
    "island": "Island",
    "country": "Country",
    "countryCode": "Country code",
    "stateProvince": "State / Province",
    "locality": "Locality",
    "institutionID": "Institution ID",
    "collectionCode": "Collection code",
    "catalogNumber": "Catalogue number",
    "sex": "Sex",
    "organismID": "Organism ID",
    "individualCount": "Individual count",
    "organismQuantity": "Organism quantity",
    "organismQuantityType": "Organism quantity type",
    "samplingProtocol": "Sampling protocol",
    "preparations": "Preparations",
    "lifeStage": "Life stage",
    "otherCatalogNumbers": "Other catalogue numbers",
    "associatedMedia": "Associated media",
}


# ============================================================
# FILE AND FOLDER SELECTION
# ============================================================

def select_excel_file():
    return askopenfilename(
        title="Select the Darwin Core Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )



def select_logo_folder():
    return askdirectory(
        title="Select the folder containing logo files"
    )


def select_output_folder():
    return askdirectory(
        title="Select the output folder for HTML files"
    )


def select_home_icon():
    return askopenfilename(
        title="Select the home icon",
        filetypes=[
            ("Image files", "*.png *.jpg *.jpeg *.gif *.webp"),
            ("PNG files", "*.png"),
        ]
    )


# ============================================================
# COLUMN SELECTION DIALOG
# ============================================================

def select_columns(columns):
    """
    Mostra una finestra scorrevole con una casella per ogni colonna.

    Restituisce:
        lista delle colonne selezionate;
        None se l'utente annulla.
    """

    selection_window = Toplevel()
    selection_window.title("Select fields to display")
    selection_window.geometry("570x700")
    selection_window.minsize(450, 400)

    result = {
        "columns": None
    }

    Label(
        selection_window,
        text="Select the Darwin Core fields to display in the HTML table:",
        font=("Arial", 11, "bold"),
        pady=10
    ).pack()

    Label(
        selection_window,
        text="The fields will appear in the same order as in the Excel file.",
        font=("Arial", 9),
        pady=3
    ).pack()

    outer_frame = Frame(selection_window)
    outer_frame.pack(fill="both", expand=True, padx=10, pady=10)

    canvas = Canvas(outer_frame)
    scrollbar = __import__("tkinter").Scrollbar(
        outer_frame,
        orient="vertical",
        command=canvas.yview
    )

    checkbox_frame = Frame(canvas)

    checkbox_frame.bind(
        "<Configure>",
        lambda event: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window(
        (0, 0),
        window=checkbox_frame,
        anchor="nw"
    )

    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    variables = {}

    for column in columns:
        selected_by_default = column in DEFAULT_SELECTED_COLUMNS
        variable = BooleanVar(value=selected_by_default)
        variables[column] = variable

        display_label = FIELD_LABELS.get(column, column)

        checkbox_text = f"{display_label}  [{column}]"

        Checkbutton(
            checkbox_frame,
            text=checkbox_text,
            variable=variable,
            anchor="w",
            justify="left",
            padx=5,
            pady=2
        ).pack(fill="x", anchor="w")

    def select_all():
        for variable in variables.values():
            variable.set(True)

    def select_none():
        for variable in variables.values():
            variable.set(False)

    def select_recommended():
        for column, variable in variables.items():
            variable.set(column in DEFAULT_SELECTED_COLUMNS)

    def confirm_selection():
        selected = [
            column
            for column in columns
            if variables[column].get()
        ]

        if not selected:
            messagebox.showwarning(
                "No fields selected",
                "Select at least one field."
            )
            return

        result["columns"] = selected
        selection_window.destroy()

    def cancel_selection():
        result["columns"] = None
        selection_window.destroy()

    buttons_frame = Frame(selection_window)
    buttons_frame.pack(fill="x", padx=10, pady=10)

    Button(
        buttons_frame,
        text="Select all",
        command=select_all
    ).pack(side="left", padx=4)

    Button(
        buttons_frame,
        text="Select none",
        command=select_none
    ).pack(side="left", padx=4)

    Button(
        buttons_frame,
        text="Recommended",
        command=select_recommended
    ).pack(side="left", padx=4)

    Button(
        buttons_frame,
        text="Cancel",
        command=cancel_selection
    ).pack(side="right", padx=4)

    Button(
        buttons_frame,
        text="Generate HTML",
        command=confirm_selection
    ).pack(side="right", padx=4)

    selection_window.protocol(
        "WM_DELETE_WINDOW",
        cancel_selection
    )

    selection_window.grab_set()
    selection_window.wait_window()

    return result["columns"]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def copy_file_with_prompt(src, dest):
    if not os.path.exists(src):
        print(f"❌ ERROR: File not found: {src}")
        return False

    if os.path.abspath(src) == os.path.abspath(dest):
        print(f"ℹ️ File already in destination: {src}")
        return True

    if os.path.exists(dest):
        response = messagebox.askyesno(
            "Existing file",
            (
                f"The file '{os.path.basename(dest)}' already exists.\n"
                "Do you want to overwrite it?"
            )
        )

        if not response:
            print(f"ℹ️ Existing file retained: {dest}")
            return True

    try:
        shutil.copy2(src, dest)
        print(f"✅ File copied: {dest}")
        return True
    except OSError as error:
        print(f"❌ Error copying {src}: {error}")
        return False


def safe_filename(value):
    """
    Crea un nome file sicuro a partire da catalogNumber,
    occurrenceID o nome dell'immagine.
    """

    value = str(value).strip()

    if not value:
        value = "record"

    value = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE)
    value = value.strip("._")

    return value or "record"


def is_url(value):
    try:
        parsed = urlparse(str(value))
        return parsed.scheme in {"http", "https"}
    except Exception:
        return False


def normalize_value(value):
    """
    Converte valori NaN o mancanti in una stringa vuota.
    """

    if pd.isna(value):
        return ""

    # Prevent integer values from being displayed as 2020.0.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def field_label(column):
    """
    Restituisce un'etichetta leggibile.
    Se non esiste nella mappa, separa automaticamente
    le parole camelCase.
    """

    if column in FIELD_LABELS:
        return FIELD_LABELS[column]

    label = re.sub(r"(?<!^)(?=[A-Z])", " ", column)
    return label[:1].upper() + label[1:]


def resolve_image_source(media_value):
    """
    Return the first image reference stored in associatedMedia.

    HTTP/HTTPS URLs are used directly. Relative filenames are also preserved
    as written. No specimen image is copied into the output repository.
    """

    media_value = normalize_value(media_value)

    if not media_value:
        return None

    first_media = re.split(r"[|;]", media_value)[0].strip()

    if not first_media:
        return None

    first_media = (
        first_media
        .replace("file:///", "")
        .replace("file://", "")
    )

    return first_media


def create_table_fields(row, selected_columns):
    """
    Prepara i campi per il template HTML.
    """

    fields = []

    for column in selected_columns:
        value = normalize_value(row.get(column, ""))

        fields.append({
            "column": column,
            "label": field_label(column),
            "value": value if value else "Not available",
            "italic": column in ITALIC_COLUMNS,
            "is_link": is_url(value),
        })

    return fields


# ============================================================
# HTML PAGE GENERATION
# ============================================================

def create_html_pages(
    df,
    selected_columns,
    logo_folder,
    output_folder,
    home_icon
):
    template_html = """
<!DOCTYPE html>
<html lang="en">
<head>

    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>{{ page_title }}</title>

    <style>

        html,
        body {
            margin: 0;
            padding: 0;
            height: 100%;
            overflow: hidden;
            font-family: Arial, Helvetica, sans-serif;
            color: #222;
            background-color: #ffffff;
        }

        /*
         * Fixed upper band containing the Home icon and record title.
         */
        .fixed-header {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
            min-height: 105px;
            background: white;
            border-bottom: 1px solid #ddd;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            padding: 18px 70px 12px;
            box-sizing: border-box;
        }

        .home-icon {
            width: 40px;
            height: 40px;
            object-fit: contain;
            position: absolute;
            top: 20px;
            left: 20px;
        }

        .title {
            margin: 8px 0 0;
            padding: 0 20px;
            text-align: center;
            font-size: 1.7rem;
            color: #821a33;
        }

        .italic {
            font-style: italic;
        }

        /*
         * Only the specimen image and Darwin Core table scroll.
         */
        .scroll-area {
            position: fixed;
            top: 105px;
            bottom: 92px;
            left: 0;
            right: 0;
            overflow-y: auto;
            overflow-x: hidden;
            -webkit-overflow-scrolling: touch;
        }

        .container {
            width: min(1000px, 92%);
            margin: 0 auto;
            padding: 28px 0 40px;
            text-align: center;
            box-sizing: border-box;
        }

        .specimen-image {
            display: block;
            width: auto;
            max-width: 100%;
            max-height: 650px;
            height: auto;
            margin: 20px auto;
            object-fit: contain;
        }

        .image-missing {
            width: 80%;
            margin: 25px auto;
            padding: 35px 15px;
            border: 1px dashed #aaa;
            color: #666;
        }

        .data-table {
            width: 100%;
            margin: 25px 0;
            border-collapse: collapse;
            table-layout: fixed;
            text-align: left;
        }

        .data-table th,
        .data-table td {
            border: 1px solid #d8d8d8;
            padding: 10px;
            overflow-wrap: anywhere;
            vertical-align: top;
        }

        .data-table th {
            width: 32%;
            background-color: #f2f2f2;
            font-weight: bold;
        }

        .data-table tr:nth-child(even) td {
            background-color: #fafafa;
        }

        /*
         * Fixed lower band containing the institutional logos.
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
            box-sizing: border-box;
        }

        .logo-container {
            width: min(1200px, 96%);
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 30px;
        }

        .logo {
            width: 110px;
            max-height: 70px;
            object-fit: contain;
        }

        a {
            color: #075da8;
        }

        @media (max-width: 600px) {

            .fixed-header {
                min-height: 90px;
                padding: 14px 55px 10px;
            }

            .home-icon {
                width: 34px;
                height: 34px;
                top: 16px;
                left: 14px;
            }

            .title {
                font-size: 1.35rem;
                padding: 0 10px;
            }

            .scroll-area {
                top: 90px;
                bottom: 80px;
            }

            .container {
                width: 95%;
            }

            .logo {
                width: 90px;
                max-height: 55px;
            }

            .data-table,
            .data-table tbody,
            .data-table tr,
            .data-table th,
            .data-table td {
                display: block;
                width: auto;
            }

            .data-table tr {
                margin-bottom: 12px;
                border: 1px solid #d8d8d8;
            }

            .data-table th,
            .data-table td {
                border: none;
            }
        }

    </style>
</head>

<body>

    <!-- Fixed upper band -->
    <header class="fixed-header">

        <a href="index.html" aria-label="Home">

            <img
                src="home.png"
                alt="Home"
                class="home-icon"
            >

        </a>

        <h1 class="italic title">
            {{ page_title }}
        </h1>

    </header>

    <!-- Scrollable specimen content -->
    <main class="scroll-area">

        <div class="container">

            {% if image %}

                <img
                    src="{{ image }}"
                    alt="{{ page_title }}"
                    class="specimen-image"
                >

            {% else %}

                <div class="image-missing">
                    Image not available
                </div>

            {% endif %}

            <h2>Data</h2>

            <table class="data-table">

                <tbody>

                    {% for field in fields %}

                    <tr>

                        <th scope="row">
                            {{ field.label }}
                        </th>

                        <td class="{% if field.italic %}italic{% endif %}">

                            {% if field.is_link %}

                                <a
                                    href="{{ field.value }}"
                                    target="_blank"
                                    rel="noopener noreferrer"
                                >
                                    {{ field.value }}
                                </a>

                            {% else %}

                                {{ field.value }}

                            {% endif %}

                        </td>

                    </tr>

                    {% endfor %}

                </tbody>

            </table>

        </div>

    </main>

    <!-- Fixed lower band -->
    <footer class="fixed-footer">

        <div class="logo-container">

            <img
                src="logo1.png"
                alt="Logo 1"
                class="logo"
            >

            <img
                src="logo2.png"
                alt="Logo 2"
                class="logo"
            >

        </div>

    </footer>

</body>
</html>
"""

    env = Environment(
        autoescape=select_autoescape(
            enabled_extensions=("html", "xml"),
            default_for_string=True
        )
    )

    template = env.from_string(template_html)

    os.makedirs(output_folder, exist_ok=True)


    # I percorsi vengono sempre definiti, indipendentemente
    # dal risultato della copia.
    logo1_source = os.path.join(logo_folder, "logo1.png")
    logo2_source = os.path.join(logo_folder, "logo2.png")

    logo1_destination = os.path.join(
        output_folder,
        "logo1.png"
    )

    logo2_destination = os.path.join(
        output_folder,
        "logo2.png"
    )

    home_destination = os.path.join(
        output_folder,
        "home.png"
    )

    copy_file_with_prompt(
        logo1_source,
        logo1_destination
    )

    copy_file_with_prompt(
        logo2_source,
        logo2_destination
    )

    copy_file_with_prompt(
        home_icon,
        home_destination
    )

    used_filenames = set()
    generated_pages = 0

    for index, row in df.iterrows():
        media_value = row.get(IMAGE_COLUMN, "")

        image_source = resolve_image_source(
            media_value=media_value
        )

        scientific_name = normalize_value(
            row.get("scientificName", "")
        )

        catalog_number = normalize_value(
            row.get("catalogNumber", "")
        )

        occurrence_id = normalize_value(
            row.get("occurrenceID", "")
        )

        page_title = (
            scientific_name
            or catalog_number
            or occurrence_id
            or f"Record {index + 1}"
        )

        base_filename = (
            catalog_number
            or occurrence_id
            or scientific_name
            or f"record_{index + 1}"
        )

        filename = safe_filename(base_filename)

        # Evita che record con lo stesso catalogNumber
        # sovrascrivano la stessa pagina.
        original_filename = filename
        counter = 2

        while filename.lower() in used_filenames:
            filename = f"{original_filename}_{counter}"
            counter += 1

        used_filenames.add(filename.lower())

        fields = create_table_fields(
            row=row,
            selected_columns=selected_columns
        )

        html_content = template.render(
            page_title=page_title,
            image=image_source,
            fields=fields
        )

        output_file = os.path.join(
            output_folder,
            f"{filename}.html"
        )

        try:
            with open(
                output_file,
                "w",
                encoding="utf-8"
            ) as file:
                file.write(html_content)

            generated_pages += 1
            print(f"✅ Generated: {output_file}")

        except OSError as error:
            print(
                f"❌ Error writing {output_file}: {error}"
            )

    messagebox.showinfo(
        "Operation completed",
        (
            f"{generated_pages} HTML pages created.\n\n"
            f"Output folder:\n{output_folder}"
        )
    )

    print(
        f"🎉 {generated_pages} HTML pages created "
        f"in {output_folder}"
    )


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    root = Tk()
    root.withdraw()

    excel_file = select_excel_file()

    if not excel_file:
        print("Operation canceled.")
        root.destroy()
        return

    try:
        df = pd.read_excel(
            excel_file,
            engine="openpyxl"
        )
    except Exception as error:
        messagebox.showerror(
            "Excel error",
            f"Unable to read the Excel file:\n\n{error}"
        )
        root.destroy()
        return

    # Normalize column names before processing the Darwin Core fields.
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    if df.empty:
        messagebox.showerror(
            "Empty file",
            "The selected Excel file contains no records."
        )
        root.destroy()
        return

    print("Columns found:")
    print(list(df.columns))

    # Non propone associatedMedia nella tabella, perché viene
    # già usata per l'immagine. Rimuovere questa esclusione
    # se si desidera poterla selezionare.
    selectable_columns = [
        column
        for column in df.columns
        if column != IMAGE_COLUMN
    ]

    selected_columns = select_columns(
        selectable_columns
    )

    if not selected_columns:
        print("Operation canceled: no columns selected.")
        root.destroy()
        return

    logo_folder = select_logo_folder()
    output_folder = select_output_folder()
    home_icon = select_home_icon()

    if not all([
        logo_folder,
        output_folder,
        home_icon
    ]):
        print(
            "Operation canceled: one or more "
            "files/folders were not selected."
        )
        root.destroy()
        return

    # Sostituisce NaN con stringhe vuote senza modificare
    # le intestazioni.
    df = df.fillna("")

    create_html_pages(
        df=df,
        selected_columns=selected_columns,
        logo_folder=logo_folder,
        output_folder=output_folder,
        home_icon=home_icon
    )

    root.destroy()


if __name__ == "__main__":
    main()

