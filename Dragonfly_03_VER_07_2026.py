"""
Generate occurrence-level specimen pages and a GBIF-style catalog interface
from a Darwin Core dataset.

The script creates:

1. One HTML page for every occurrenceID:
       MZURODOB00001.html
       MZURODOB00002.html
       ...

2. A catalog interface:
       catalogo_gbif.html

Each catalog code links directly to the corresponding occurrence page.

The occurrence pages display:
- the specimen image from associatedMedia;
- the occurrenceID;
- all non-empty Darwin Core metadata available in the spreadsheet.

The generated files are designed to be stored in the same website folder as
index.html so that all relative links remain compatible with GitHub Pages.
"""

import html
import json
import os
import re
import shutil
from string import Template
from urllib.parse import urlparse

import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory


# ============================================================
# DARWIN CORE CONFIGURATION
# ============================================================

SHEET_NAME = "Occurrences"

OCCURRENCE_ID_COLUMN = "occurrenceID"
MEDIA_COLUMN = "associatedMedia"
SCIENTIFIC_NAME_COLUMN = "scientificName"

REQUIRED_COLUMNS = [
    OCCURRENCE_ID_COLUMN,
]

# Human-readable labels displayed in occurrence-page metadata tables.
# The Darwin Core column names in the Excel file are not changed.
FIELD_LABELS = {
    "occurrenceID": "Occurrence ID",
    "basisOfRecord": "Basis of record",
    "eventDate": "Date",
    "startDayOfYear": "Day of the year",
    "endDayOfYear": "End day of the year",
    "verbatimEventDate": "Original date",
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
    "coordinateUncertaintyInMeters": "Coordinate uncertainty (m)",
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
    "institutionCode": "Institution code",
    "collectionCode": "Collection code",
    "catalogNumber": "Catalog number",
    "otherCatalogNumbers": "Other catalog numbers",
    "sex": "Sex",
    "organismID": "Organism ID",
    "individualCount": "Individual count",
    "organismQuantity": "Organism quantity",
    "organismQuantityType": "Organism quantity type",
    "samplingProtocol": "Sampling protocol",
    "preparations": "Preparations",
    "lifeStage": "Life stage",
    "recordedBy": "Recorded by",
    "fieldNumber": "Field number",
    "eventID": "Event ID",
    "locationID": "Location ID",
    "taxonID": "Taxon ID",
    "acceptedNameUsage": "Accepted name",
    "acceptedNameUsageID": "Accepted name ID",
    "parentNameUsage": "Parent taxon",
    "parentNameUsageID": "Parent taxon ID",
    "nameAccordingTo": "Name according to",
    "nameAccordingToID": "Name according to ID",
    "taxonomicStatus": "Taxonomic status",
    "identificationQualifier": "Identification qualifier",
    "dateIdentified": "Date identified",
    "identificationRemarks": "Identification remarks",
}


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
        column
        for column in REQUIRED_COLUMNS
        if column not in dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing Darwin Core columns:\n\n"
            + "\n".join(
                f"• {column}"
                for column in missing_columns
            )
        )


def first_media(value):
    """
    Return the first associatedMedia item.

    Multiple media references separated by "|" or ";" are supported.
    """
    value = normalize_text(value)

    if not value:
        return ""

    for separator in ("|", ";"):
        if separator in value:
            value = value.split(separator, 1)[0].strip()

    return value


def media_filename(media_value):
    """Extract the filename from a local path or web URL."""
    media_value = normalize_text(media_value)

    if not media_value:
        return ""

    parsed = urlparse(media_value)

    if parsed.scheme in {"http", "https"}:
        return os.path.basename(parsed.path)

    return os.path.basename(
        media_value.replace("\\", "/")
    )


def copy_asset(source_path, destination_folder, output_name):
    """Copy a shared visual asset to the website folder."""
    if not source_path or not os.path.isfile(source_path):
        raise FileNotFoundError(
            f"File not found: {source_path}"
        )

    destination_path = os.path.join(
        destination_folder,
        output_name,
    )

    if os.path.abspath(source_path) == os.path.abspath(destination_path):
        return

    if os.path.exists(destination_path):
        overwrite = messagebox.askyesno(
            "Existing file",
            (
                f"The file '{output_name}' already exists.\n"
                "Do you want to overwrite it?"
            ),
        )

        if not overwrite:
            return

    shutil.copy2(
        source_path,
        destination_path,
    )


def field_label(column):
    """
    Return a human-readable label for a Darwin Core field.

    Explicit labels in FIELD_LABELS are preferred. Unknown camelCase terms
    are converted automatically, for example:
        georeferenceSources -> Georeference sources
    """
    column = str(column)

    if column in FIELD_LABELS:
        return FIELD_LABELS[column]

    label = re.sub(r"(?<!^)(?=[A-Z])", " ", column)
    return label[:1].upper() + label[1:]


def build_metadata_rows(row):
    """
    Return all non-empty Darwin Core fields for display in the occurrence page.

    associatedMedia is omitted from the table because it is displayed
    separately as the specimen image.
    """
    metadata_rows = []

    for column, value in row.items():
        if column == MEDIA_COLUMN:
            continue

        clean_value = normalize_text(value)

        if not clean_value:
            continue

        metadata_rows.append(
            {
                "field": field_label(column),
                "value": clean_value,
            }
        )

    return metadata_rows


# ============================================================
# OCCURRENCE PAGE TEMPLATE
# ============================================================

occurrence_template = Template("""
<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="UTF-8">

    <meta name="viewport"
          content="width=device-width, initial-scale=1.0">

    <title>$occurrence_id</title>

    <style>

        html,
        body {
            margin: 0;
            padding: 0;
            height: 100%;
            font-family: Arial, Helvetica, sans-serif;
            color: #222;
            background-color: white;
            overflow: hidden;
        }

        /*
         * Fixed upper band containing the Home icon, occurrenceID,
         * and scientific name.
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
            min-height: 118px;
            padding: 14px 70px 12px;
            box-sizing: border-box;
            text-align: center;
        }

        .home-icon {
            width: 40px;
            height: auto;
            position: absolute;
            top: 20px;
            left: 20px;
            cursor: pointer;
        }

        h1 {
            color: #821a33;
            font-size: 40px;
            text-align: center;
            margin: 0 0 8px;
        }

        h2 {
            text-align: center;
            font-style: italic;
            color: darkgreen;
            margin: 0;
            font-size: 28px;
        }

        /*
         * Only the central occurrence content scrolls.
         */
        .scroll-area {
            position: fixed;
            top: 118px;
            bottom: 118px;
            left: 0;
            right: 0;
            overflow-y: auto;
            overflow-x: hidden;
            -webkit-overflow-scrolling: touch;
        }

        .container {
            width: min(1150px, 94%);
            margin: 0 auto;
            padding: 24px 20px 40px;
            text-align: center;
            box-sizing: border-box;
        }

        .specimen-image {
            max-width: 90%;
            max-height: 750px;
            height: auto;
            object-fit: contain;
            margin: 20px auto 30px;
            border: 1px solid #ddd;
        }

        .data-table {
            width: min(1000px, 96%);
            margin: 25px auto;
            border-collapse: collapse;
            text-align: left;
        }

        .data-table th,
        .data-table td {
            border: 1px solid #ddd;
            padding: 9px 12px;
            vertical-align: top;
        }

        .data-table th {
            width: 32%;
            background-color: #f3f3f3;
            color: #821a33;
        }

        /*
         * Fixed lower band containing navigation controls and institutional logos.
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
            padding: 8px 20px 10px;
            box-sizing: border-box;
        }

        .footer-content {
            width: min(1200px, 96%);
            margin: 0 auto;
        }

        .logos {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 30px;
        }

        .logos img {
            width: 178.5px;
            max-height: 58px;
            object-fit: contain;
        }

        .navigation {
            margin: 0 0 8px;
            text-align: center;
        }

        .navigation a {
            display: inline-block;
            margin: 6px;
            padding: 10px 18px;
            background-color: #821a33;
            color: white;
            text-decoration: none;
            border-radius: 6px;
        }

        .navigation a:hover {
            background-color: #c54c00;
        }

        .missing-image {
            margin: 30px auto;
            color: #821a33;
            font-size: 18px;
        }

        @media (max-width: 700px) {
            .fixed-header {
                min-height: 108px;
                padding: 12px 55px 10px;
            }

            .home-icon {
                width: 34px;
                top: 16px;
                left: 14px;
            }

            h1 {
                font-size: 30px;
            }

            h2 {
                font-size: 22px;
            }

            .scroll-area {
                top: 108px;
                bottom: 108px;
            }

            .logos img {
                width: 125px;
                max-height: 48px;
            }

            .navigation a {
                padding: 8px 12px;
                margin: 3px;
            }
        }

    </style>

</head>

<body>

    <!-- Fixed upper band -->
    <header class="fixed-header">

        <img
            src="$home_icon"
            alt="Home"
            class="home-icon"
            onclick="location.href='index.html'"
        />

        <h1>$occurrence_id</h1>

        $scientific_name_heading

    </header>

    <!-- Scrollable occurrence content -->
    <main class="scroll-area">

        <div class="container">

            $image_block

            <table class="data-table">

                $metadata_rows

            </table>

        </div>

    </main>

    <!-- Fixed lower band -->
    <footer class="fixed-footer">

        <div class="footer-content">

            <div class="navigation">

                <a href="catalogo_gbif.html">
                    Back to GBIF Catalog
                </a>

                <a href="index.html">
                    Home
                </a>

            </div>

            <div class="logos">

                <img src="$logo1" alt="Logo 1">

                <img src="$logo2" alt="Logo 2">

            </div>

        </div>

    </footer>

</body>

</html>
""")


# ============================================================
# GBIF CATALOG TEMPLATE
# ============================================================

catalog_template = Template("""
<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="UTF-8">

    <meta name="viewport"
          content="width=device-width, initial-scale=1.0">

    <title>Catalog based on GBIF dataset</title>

    <style>

        html,
        body {
            margin: 0;
            padding: 0;
            height: 100%;
            overflow: hidden;
            font-family: Arial, Helvetica, sans-serif;
            background: white;
            color: #222;
        }

        /*
         * Fixed upper band containing the Home icon and catalog title.
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
            min-height: 105px;
            box-sizing: border-box;
        }

        /*
         * Scrollable catalog controls between the fixed header and footer.
         */
        .scroll-area {
            position: fixed;
            top: 105px;
            bottom: 85px;
            left: 0;
            right: 0;
            overflow-y: auto;
            overflow-x: hidden;
            -webkit-overflow-scrolling: touch;
        }

        .container {
            text-align: center;
            padding: 35px 20px;
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

        .logos {
            display: flex;
            justify-content: space-between;
            align-items: center;
            width: min(1200px, 96%);
            margin: 0 auto;
        }

        .logos img {
            width: 178.5px;
            height: 47px;
            object-fit: contain;
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
            font-size: 54px;
            text-align: center;
            margin: 22px 70px 12px;
        }

        @media (max-width: 700px) {
            .fixed-header {
                min-height: 90px;
            }

            .scroll-area {
                top: 90px;
                bottom: 76px;
            }

            h2 {
                font-size: 38px;
                margin-top: 22px;
            }

            .home-icon {
                width: 34px;
                top: 16px;
                left: 14px;
            }

            .logos img {
                width: 130px;
                height: auto;
                max-height: 50px;
            }
        }

    </style>

</head>

<body>

    <!-- Fixed upper band -->
    <header class="fixed-header">

        <img
            src="$home_icon"
            alt="Home"
            class="home-icon"
            onclick="location.href='index.html'"
        />

        <h2>Catalog based on GBIF dataset</h2>

    </header>

    <!-- Scrollable catalog controls -->
    <main class="scroll-area">

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

            <br><br>

            <button onclick="checkCode()">
                Open Record
            </button>

            <p id="result"></p>

        </div>

    </main>

    <!-- Fixed lower band -->
    <footer class="fixed-footer">

        <div class="logos">

            <img src="$logo1" alt="Logo 1">

            <img src="$logo2" alt="Logo 2">

        </div>

    </footer>

    <script type="text/javascript">

        var occurrenceData = $occurrences;

        function checkCode() {

            var code =
                "MZURODOB" +
                document.getElementById("num1").value +
                document.getElementById("num2").value +
                document.getElementById("num3").value +
                document.getElementById("num4").value +
                document.getElementById("num5").value;

            var resultElement =
                document.getElementById("result");

            if (code in occurrenceData) {

                var htmlFile =
                    occurrenceData[code];

                resultElement.innerHTML =
                    "<a href='" +
                    htmlFile +
                    "' target='_self'>" +
                    "Open " +
                    code +
                    "</a>";

            } else {

                resultElement.innerText =
                    "Code not found";

            }
        }

    </script>

</body>

</html>
""")


# ============================================================
# PAGE GENERATION
# ============================================================

def generate_occurrence_pages(
    dataframe,
    website_folder,
    logo1_name,
    logo2_name,
    home_icon_name,
):
    """
    Generate one HTML page for each occurrenceID.

    Returns:
        dict:
            occurrenceID -> HTML filename
    """
    occurrence_data = {}
    generated_pages = 0
    skipped_rows = 0

    total_records = len(dataframe)

    for position, (_, row) in enumerate(
        dataframe.iterrows(),
        start=1,
    ):
        occurrence_id = normalize_text(
            row.get(OCCURRENCE_ID_COLUMN)
        )

        if not occurrence_id:
            skipped_rows += 1
            continue

        html_filename = (
            f"{occurrence_id}.html"
        )

        occurrence_data[
            occurrence_id
        ] = html_filename

        scientific_name = normalize_text(
            row.get(
                SCIENTIFIC_NAME_COLUMN,
                "",
            )
        )

        if scientific_name:
            scientific_name_heading = (
                "<h2>"
                + html.escape(
                    scientific_name
                )
                + "</h2>"
            )
        else:
            scientific_name_heading = ""

        media_value = first_media(
            row.get(
                MEDIA_COLUMN,
                "",
            )
        )

        if media_value:
            safe_media = html.escape(
                media_value,
                quote=True,
            )

            image_block = (
                '<img '
                f'src="{safe_media}" '
                'alt="Specimen image" '
                'class="specimen-image">'
            )
        else:
            image_block = (
                '<div class="missing-image">'
                'No associated image available'
                '</div>'
            )

        metadata_rows_data = (
            build_metadata_rows(
                row
            )
        )

        metadata_rows_html = "\n".join(
            (
                "<tr>"
                f"<th>{html.escape(item['field'])}</th>"
                f"<td>{html.escape(item['value'])}</td>"
                "</tr>"
            )
            for item in metadata_rows_data
        )

        page_html = occurrence_template.substitute(
            occurrence_id=html.escape(
                occurrence_id
            ),
            scientific_name_heading=scientific_name_heading,
            image_block=image_block,
            metadata_rows=metadata_rows_html,
            logo1=logo1_name,
            logo2=logo2_name,
            home_icon=home_icon_name,
        )

        output_path = os.path.join(
            website_folder,
            html_filename,
        )

        with open(
            output_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write(
                page_html
            )

        generated_pages += 1

        if (
            position == 1
            or position % 100 == 0
            or position == total_records
        ):
            print(
                f"[{position}/{total_records}] "
                f"Generated: {html_filename}"
            )

    print(
        f"✅ Occurrence pages created: "
        f"{generated_pages}"
    )

    if skipped_rows:
        print(
            f"⚠️ Rows skipped because occurrenceID was empty: "
            f"{skipped_rows}"
        )

    return occurrence_data


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    """Run the complete occurrence-page and GBIF-catalog workflow."""

    root = Tk()
    root.withdraw()

    try:

        # Select the Darwin Core workbook.
        file_path = askopenfilename(
            title="Select the Excel file",
            filetypes=[
                (
                    "Excel files",
                    "*.xlsx",
                )
            ],
        )

        if not file_path:
            return

        # Select the website root.
        #
        # This should normally be the same folder containing index.html.
        website_folder = askdirectory(
            title=(
                "Select the WEBSITE ROOT folder "
                "where occurrence pages and catalogo_gbif.html "
                "will be created"
            )
        )

        if not website_folder:
            return

        # Select logos and Home icon.
        image_filetypes = [
            (
                "Image files",
                "*.png *.jpg *.jpeg *.gif *.webp",
            )
        ]

        logo1_path = askopenfilename(
            title="Select logo1 file",
            filetypes=image_filetypes,
        )

        logo2_path = askopenfilename(
            title="Select logo2 file",
            filetypes=image_filetypes,
        )

        home_icon_path = askopenfilename(
            title="Select the Home icon",
            filetypes=image_filetypes,
        )

        if not all(
            [
                logo1_path,
                logo2_path,
                home_icon_path,
            ]
        ):
            return

        # Load Darwin Core data.
        print(
            "\n=== Occurrence-page generation started ==="
        )

        occurrences_df = pd.read_excel(
            file_path,
            sheet_name=SHEET_NAME,
            engine="openpyxl",
            dtype=object,
        )

        occurrences_df.columns = [
            str(column).strip()
            for column in occurrences_df.columns
        ]

        validate_columns(
            occurrences_df
        )

        # Copy shared assets into the website root.
        logo1_name = os.path.basename(
            logo1_path
        )

        logo2_name = os.path.basename(
            logo2_path
        )

        home_icon_name = os.path.basename(
            home_icon_path
        )

        copy_asset(
            logo1_path,
            website_folder,
            logo1_name,
        )

        copy_asset(
            logo2_path,
            website_folder,
            logo2_name,
        )

        copy_asset(
            home_icon_path,
            website_folder,
            home_icon_name,
        )

        # Generate all MZURODOBxxxxx.html pages.
        occurrence_data = generate_occurrence_pages(
            dataframe=occurrences_df,
            website_folder=website_folder,
            logo1_name=logo1_name,
            logo2_name=logo2_name,
            home_icon_name=home_icon_name,
        )

        if not occurrence_data:
            raise ValueError(
                "No occurrence pages were generated. "
                "Check the occurrenceID column."
            )

        # Generate the catalog interface.
        numbers_options = "\n".join(
            [
                f"<option value='{i}'>{i}</option>"
                for i in range(10)
            ]
        )

        catalog_html = catalog_template.substitute(
            numbers=numbers_options,
            occurrences=json.dumps(
                occurrence_data,
                ensure_ascii=False,
            ),
            logo1=logo1_name,
            logo2=logo2_name,
            home_icon=home_icon_name,
        )

        catalog_path = os.path.join(
            website_folder,
            "catalogo_gbif.html",
        )

        with open(
            catalog_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write(
                catalog_html
            )

        print(
            "✅ catalogo_gbif.html created:"
        )

        print(
            f"   {catalog_path}"
        )

        print(
            "=== Generation completed ===\n"
        )

        messagebox.showinfo(
            "Operation completed",
            (
                "Occurrence pages and GBIF Catalog were created successfully.\n\n"
                f"Occurrence pages: {len(occurrence_data)}\n"
                "Catalog page: catalogo_gbif.html\n\n"
                f"Website folder:\n{website_folder}"
            ),
        )

    except Exception as error:

        print(
            f"❌ Error: {error}"
        )

        messagebox.showerror(
            "Error",
            str(error),
        )

    finally:

        root.destroy()


if __name__ == "__main__":
    main()
