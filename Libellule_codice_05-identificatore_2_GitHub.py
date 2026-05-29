import pandas as pd
import os
import re
import shutil
from jinja2 import Environment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

# Function for selecting a file
def select_file(title, filetypes):

    root = Tk()
    root.withdraw()

    return askopenfilename(
        title=title,
        filetypes=filetypes
    )

# Function for selecting a folder
def select_folder(title):

    root = Tk()
    root.withdraw()

    return askdirectory(title=title)

# Function to create safe filenames
def clean_filename(identifier):

    identifiers = [
        name.strip()
        for name in identifier.split(',')
    ]

    identifiers = [
        re.sub(r"[^a-zA-Z0-9 ]", '', name)
        for name in identifiers
    ]

    identifiers = [
        name.replace(' ', '_')
        for name in identifiers
    ]

    # Remove duplicate names while preserving order
    identifiers = list(dict.fromkeys(identifiers))

    return "__".join(identifiers).lower().lstrip('/')

# Function for safely copying files
def safe_copy(src, dest):

    try:
        # Compare absolute paths to avoid unnecessary copying
        if os.path.abspath(src) != os.path.abspath(dest):

            shutil.copy(src, dest)

            print(f"File copied from {src} to {dest}")

        else:

            print(
                f"The files {src} and {dest} are already identical. "
                f"Copy operation skipped."
            )

    except Exception as e:

        print(
            f"Error while copying file from {src} to {dest}: {e}"
        )

# Function to create HTML pages grouped by identifier
def create_html_pages_by_identifier(
    excel_file,
    output_folder,
    home_file,
    home_icon,
    left_logo,
    right_logo
):

    try:
        # Read the Excel file
        df = pd.read_excel(
            excel_file,
            sheet_name='Occurrences',
            engine="openpyxl"
        )

        print("Columns found in the Excel file:")
        print(df.columns)

        print("First rows of the Excel file:")
        print(df.head())

    except Exception as e:

        print(f"Error while reading the Excel file: {e}")
        return

    # Replace missing values with empty strings
    df.fillna('', inplace=True)

    # Clean and standardize data
    df['Identifier'] = (
        df['Identifier']
        .astype(str)
        .str.strip()
    )

    df['ScientificName'] = (
        df['ScientificName']
        .astype(str)
        .str.strip()
    )

    df['Continent'] = (
        df['Continent']
        .astype(str)
        .str.strip()
    )

    df['Country'] = (
        df['Country']
        .astype(str)
        .str.strip()
    )

    df['Locality'] = (
        df['Locality']
        .astype(str)
        .str.strip()
    )

    # Group records by identifier and locality
    df_grouped = df.groupby(
        [
            'Identifier',
            'ScientificName',
            'Continent',
            'Country',
            'Locality'
        ],
        as_index=False
    ).size().rename(
        columns={'size': 'NumberOfIndividuals'}
    )

    print("Example of grouped data:")
    print(df_grouped.head())

    # HTML template
    template_html = '''
    <!DOCTYPE html>
    <html lang="en">

    <head>

        <meta charset="UTF-8">

        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">

        <title>{{ identifier }}</title>

        <style>

            .container {
                text-align: center;
            }

            .data-table {
                width: 80%;
                margin: 20px auto;
                border-collapse: collapse;
            }

            .data-table th,
            .data-table td {
                border: 1px solid #ddd;
                padding: 8px;
            }

            h1 {
                font-style: italic;
            }

            .home-icon {
                position: absolute;
                top: 10px;
                left: 10px;
                width: 40px;
                height: 40px;
            }

            .logos {
                display: flex;
                justify-content: space-between;
                margin-top: 20px;
            }

            .logos img {
                width: 100px;
            }

        </style>

    </head>

    <body>

        <a href="home.html">

            <img src="{{ home_icon }}"
                 alt="Home"
                 class="home-icon">

        </a>

        <div class="container">

            <h1>{{ identifier }}</h1>

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

                    <td>
                        <em>{{ row.ScientificName }}</em>
                    </td>

                    <td>{{ row.Continent }}</td>

                    <td>{{ row.Country }}</td>

                    <td>{{ row.Locality }}</td>

                    <td>{{ row.NumberOfIndividuals }}</td>

                </tr>

                {% endfor %}

            </table>

            <div class="logos">

                <img src="{{ left_logo }}"
                     alt="Left Logo">

                <img src="{{ right_logo }}"
                     alt="Right Logo">

            </div>

        </div>

    </body>

    </html>
    '''

    env = Environment()
    template = env.from_string(template_html)

    # Create output folder if it does not exist
    os.makedirs(output_folder, exist_ok=True)

    # Copy images to the output folder
    try:

        safe_copy(
            home_icon,
            os.path.join(output_folder, 'home.png')
        )

        safe_copy(
            left_logo,
            os.path.join(output_folder, 'logo1.png')
        )

        safe_copy(
            right_logo,
            os.path.join(output_folder, 'logo2.png')
        )

        print("Images copied successfully.")

    except Exception as e:

        print(f"Error while copying images: {e}")
        return

    # Relative paths for copied images
    home_icon_rel = 'home.png'
    left_logo_rel = 'logo1.png'
    right_logo_rel = 'logo2.png'

    # Generate HTML files
    for identifier, group in df_grouped.groupby('Identifier'):

        try:

            print(
                f"Identifier: {identifier}, "
                f"Number of rows found: {len(group)}"
            )

            safe_filename = clean_filename(identifier)

            output_file = os.path.join(
                output_folder,
                f"{safe_filename}.html"
            )

            html_content = template.render(
                identifier=identifier,
                data=group.to_dict(orient='records'),
                home_page="home.html",
                home_icon=home_icon_rel,
                left_logo=left_logo_rel,
                right_logo=right_logo_rel
            )

            with open(
                output_file,
                'w',
                encoding='utf-8'
            ) as f:

                f.write(html_content)

            print(
                f"HTML file created for {identifier}: "
                f"{output_file}"
            )

        except Exception as e:

            print(
                f"Error while creating the HTML file "
                f"for {identifier}: {e}"
            )

    print(f'HTML pages created in {output_folder}')

# File and folder selection
excel_file = select_file(
    "Select the Excel file",
    [("Excel files", "*.xlsx")]
)

output_folder = select_folder(
    "Select the output folder for saving HTML files"
)

home_file = select_file(
    "Select the Home file",
    [("HTML files", "*.html")]
)

home_icon = select_file(
    "Select the Home icon",
    [("Image files", "*.png;*.jpg;*.jpeg;*.gif")]
)

left_logo = select_file(
    "Select the left logo",
    [("Image files", "*.png;*.jpg;*.jpeg;*.gif")]
)

right_logo = select_file(
    "Select the right logo",
    [("Image files", "*.png;*.jpg;*.jpeg;*.gif")]
)

# Verify that all files and folders have been selected
if all([
    excel_file,
    output_folder,
    home_file,
    home_icon,
    left_logo,
    right_logo
]):

    create_html_pages_by_identifier(
        excel_file,
        output_folder,
        home_file,
        home_icon,
        left_logo,
        right_logo
    )

else:

    print(
        "Operation cancelled: one or more "
        "files/folders were not selected."
    )

# Keep the console window open
input("Press Enter to close...")
