"""
Generate identifier-based HTML pages from a Darwin Core dataset.

The script reads occurrence records from an Excel workbook,
groups them by identifiedBy, summarizes the associated species
and locality data, copies the required assets, and creates one
HTML page for each identifier.
"""

import os
import re
import shutil

import pandas as pd
from jinja2 import Environment, select_autoescape
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory


# ============================================================
# DARWIN CORE CONFIGURATION
# ============================================================

SHEET_NAME = "Occurrences"

IDENTIFIER_COLUMN = "identifiedBy"
SCIENTIFIC_NAME_COLUMN = "scientificName"
CONTINENT_COLUMN = "continent"
COUNTRY_COLUMN = "country"
LOCALITY_COLUMN = "locality"

REQUIRED_COLUMNS = [
    IDENTIFIER_COLUMN,
    SCIENTIFIC_NAME_COLUMN,
    CONTINENT_COLUMN,
    COUNTRY_COLUMN,
    LOCALITY_COLUMN,
]


# ============================================================
# FILE AND FOLDER SELECTION
# ============================================================

def select_file(title, filetypes):
    """Opens a file-selection dialog and returns the selected path."""
    root = Tk()
    root.withdraw()

    selected = askopenfilename(
        title=title,
        filetypes=filetypes
    )

    root.destroy()
    return selected


def select_folder(title):
    """Opens a folder-selection dialog and returns the selected path."""
    root = Tk()
    root.withdraw()

    selected = askdirectory(title=title)

    root.destroy()
    return selected


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_text(value):
    """Converts missing values to empty strings and trims whitespace."""
    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def validate_columns(dataframe):
    """Checks that all required Darwin Core columns are present."""
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


def clean_filename(identifier):
    """
    Creates a filesystem-safe filename from one or more identifiers.

    Multiple names separated by commas are normalized, deduplicated,
    and joined with a double underscore.
    """
    identifier = normalize_text(identifier)

    identifiers = [
        name.strip()
        for name in identifier.split(",")
        if name.strip()
    ]

    identifiers = [
        re.sub(r"[^a-zA-Z0-9 ]", "", name)
        for name in identifiers
    ]

    identifiers = [
        name.replace(" ", "_")
        for name in identifiers
        if name
    ]

    identifiers = list(dict.fromkeys(identifiers))

    filename = "__".join(
        identifiers
    ).lower().lstrip("/")

    return filename or "unknown_identifier"


def safe_copy(src, dest):
    """Copies an asset with overwrite confirmation when required."""
    if not src or not os.path.isfile(src):
        raise FileNotFoundError(
            f"File not found: {src}"
        )

    if os.path.abspath(src) == os.path.abspath(dest):
        print(
            f"The files {src} and {dest} are already identical. "
            "Copy operation skipped."
        )
        return

    if os.path.exists(dest):
        overwrite = messagebox.askyesno(
            "Existing file",
            (
                f"The file '{os.path.basename(dest)}' "
                "already exists.\n"
                "Do you want to overwrite it?"
            )
        )

        if not overwrite:
            print(
                f"Existing file retained: {dest}"
            )
            return

    shutil.copy2(src, dest)

    print(
        f"File copied from {src} to {dest}"
    )


# ============================================================
# HTML PAGE GENERATION
# ============================================================

def create_html_pages_by_identifier(
    excel_file,
    output_folder,
    home_icon,
    left_logo,
    right_logo
):
    """
    Generates one HTML summary page for each identifiedBy value.

    Records are grouped by identifier, scientific name, continent,
    country, and locality. The resulting table reports the number
    of occurrence records associated with each combination.
    """
    try:
        # Load the Darwin Core records required for identifier pages.
        df = pd.read_excel(
            excel_file,
            sheet_name=SHEET_NAME,
            engine="openpyxl",
            dtype=object
        )

        print("Columns found in the Excel file:")
        print(df.columns.tolist())

        print("First rows of the Excel file:")
        print(df.head())

    except Exception as error:
        raise RuntimeError(
            f"Error while reading the Excel file: {error}"
        ) from error

    # Normalize column names before validating the required fields.
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    validate_columns(df)

    # Keep only the fields used by the identifier workflow.
    df = df[
        REQUIRED_COLUMNS
    ].copy()

    # Normalize all values used for grouping and display.
    for column in REQUIRED_COLUMNS:
        df[column] = df[column].apply(
            normalize_text
        )

    # Records without identifiedBy cannot be assigned to an identifier page.
    df = df[
        df[IDENTIFIER_COLUMN] != ""
    ].copy()

    if df.empty:
        raise ValueError(
            "No valid records were found in the identifiedBy column."
        )

    # Count occurrence records for each identifier and geographic grouping.
    df_grouped = (
        df.groupby(
            [
                IDENTIFIER_COLUMN,
                SCIENTIFIC_NAME_COLUMN,
                CONTINENT_COLUMN,
                COUNTRY_COLUMN,
                LOCALITY_COLUMN,
            ],
            dropna=False,
            as_index=False
        )
        .size()
        .rename(
            columns={
                "size": "NumberOfIndividuals"
            }
        )
    )

    print("Example of grouped data:")
    print(df_grouped.head())

    # ========================================================
    # ORIGINAL HTML TEMPLATE
    # ========================================================

    template_html = """
    <!DOCTYPE html>
    <html lang="en">

    <head>

        <meta charset="UTF-8">

        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">

        <title>{{ identifier }}</title>

        <style>

            html,
            body {
                margin: 0;
                padding: 0;
                height: 100%;
                overflow: hidden;
                font-family: Arial, Helvetica, sans-serif;
                color: #222;
                background: white;
            }

            /*
             * Fixed upper band containing the Home icon and identifier name.
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
                box-sizing: border-box;
                padding: 18px 70px 12px;
            }

            .home-icon {
                position: absolute;
                top: 20px;
                left: 20px;
                width: 40px;
                height: 40px;
                object-fit: contain;
            }

            h1 {
                margin: 8px 0 0;
                text-align: center;
                color: #821a33;
                font-size: 42px;
            }

            /*
             * Only the identifier summary table scrolls.
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
                width: min(1150px, 94%);
                margin: 0 auto;
                text-align: center;
                padding: 28px 10px 40px;
                box-sizing: border-box;
            }

            .data-table {
                width: 90%;
                margin: 25px auto;
                border-collapse: collapse;
            }

            .data-table th,
            .data-table td {
                border: 1px solid #ddd;
                padding: 9px;
            }

            .data-table th {
                background: #f3f3f3;
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

            .logos {
                display: flex;
                justify-content: space-between;
                align-items: center;
                width: min(1200px, 96%);
                margin: 0 auto;
            }

            .logos img {
                width: 150px;
                max-height: 65px;
                object-fit: contain;
            }

            @media (max-width: 700px) {

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

                h1 {
                    font-size: 30px;
                }

                .scroll-area {
                    top: 90px;
                    bottom: 80px;
                }

                .logos img {
                    width: 120px;
                    max-height: 52px;
                }

                .data-table {
                    width: 100%;
                    font-size: 14px;
                }
            }

        </style>

    </head>

    <body>

        <!-- Fixed upper band -->
        <header class="fixed-header">

            <a href="index.html">
                <img
                    src="{{ home_icon }}"
                    alt="Home"
                    class="home-icon"
                >
            </a>

            <h1>{{ identifier }}</h1>

        </header>

        <!-- Scrollable identifier summary -->
        <main class="scroll-area">

            <div class="container">

                <h2>Data</h2>

                <table class="data-table">

                    <tr>
                        <th>Species</th>
                        <th>Continent</th>
                        <th>Country</th>
                        <th>Locality</th>
                        <th>Individuals</th>
                    </tr>

                    {% for row in data %}

                    <tr>
                        <td><em>{{ row.scientificName }}</em></td>
                        <td>{{ row.continent }}</td>
                        <td>{{ row.country }}</td>
                        <td>{{ row.locality }}</td>
                        <td>{{ row.NumberOfIndividuals }}</td>
                    </tr>

                    {% endfor %}

                </table>

            </div>

        </main>

        <!-- Fixed lower band -->
        <footer class="fixed-footer">

            <div class="logos">

                <img src="{{ left_logo }}" alt="Left Logo">

                <img src="{{ right_logo }}" alt="Right Logo">

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

    template = env.from_string(
        template_html
    )

    # Create the output folder before copying shared assets.
    os.makedirs(
        output_folder,
        exist_ok=True
    )

    safe_copy(
        home_icon,
        os.path.join(
            output_folder,
            "home.png"
        )
    )

    safe_copy(
        left_logo,
        os.path.join(
            output_folder,
            "logo1.png"
        )
    )

    safe_copy(
        right_logo,
        os.path.join(
            output_folder,
            "logo2.png"
        )
    )

    generated_pages = 0

    # Generate one page for each unique identifiedBy value.
    for identifier, group in df_grouped.groupby(
        IDENTIFIER_COLUMN
    ):
        try:
            print(
                f"Identifier: {identifier}, "
                f"Number of rows found: {len(group)}"
            )

            output_file = os.path.join(
                output_folder,
                f"{clean_filename(identifier)}.html"
            )

            html_content = template.render(
                identifier=identifier,
                data=group.to_dict(
                    orient="records"
                ),
                home_icon="home.png",
                left_logo="logo1.png",
                right_logo="logo2.png"
            )

            with open(
                output_file,
                "w",
                encoding="utf-8"
            ) as file:
                file.write(
                    html_content
                )

            generated_pages += 1

            print(
                f"HTML file created for {identifier}: "
                f"{output_file}"
            )

        except Exception as error:
            print(
                "Error while creating the HTML file "
                f"for {identifier}: {error}"
            )

    print(
        f"HTML pages created in {output_folder}"
    )

    messagebox.showinfo(
        "Operation completed",
        (
            "Identifier pages were created successfully.\n\n"
            f"Pages generated: {generated_pages}\n"
            f"Output folder:\n{output_folder}"
        )
    )


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    excel_file = select_file(
        "Select the Excel file",
        [
            (
                "Excel files",
                "*.xlsx"
            )
        ]
    )

    output_folder = select_folder(
        "Select the output folder for saving HTML files"
    )

    image_filetypes = [
        (
            "Image files",
            "*.png *.jpg *.jpeg *.gif *.webp"
        )
    ]

    home_icon = select_file(
        "Select the Home icon",
        image_filetypes
    )

    left_logo = select_file(
        "Select the left logo",
        image_filetypes
    )

    right_logo = select_file(
        "Select the right logo",
        image_filetypes
    )

    if not all([
        excel_file,
        output_folder,
        home_icon,
        left_logo,
        right_logo
    ]):
        print(
            "Operation cancelled: one or more "
            "files/folders were not selected."
        )
        return

    try:
        create_html_pages_by_identifier(
            excel_file,
            output_folder,
            home_icon,
            left_logo,
            right_logo
        )

    except Exception as error:
        print(
            "\nError while generating identifier pages: "
            f"{error}\n"
        )

        messagebox.showerror(
            "Error",
            str(error)
        )


if __name__ == "__main__":
    main()

    input("Press Enter to close...")
