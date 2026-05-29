import pandas as pd
import os
import shutil
import re
from jinja2 import Environment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

# Functions for selecting files and folders
def select_file(title, filetypes):

    root = Tk()
    root.withdraw()

    return askopenfilename(
        title=title,
        filetypes=filetypes
    )

def select_folder(title):

    root = Tk()
    root.withdraw()

    return askdirectory(title=title)

# Function to generate HTML pages
def create_html_pages(
    excel_file,
    output_folder,
    home_icon,
    left_logo,
    right_logo,
    image_folder
):

    try:
        # Read the Excel file
        df = pd.read_excel(
            excel_file,
            sheet_name='Occurrences',
            engine="openpyxl"
        )

        print(f"Data successfully loaded from: {excel_file}")

    except Exception as e:

        print(f"Error while reading the Excel file: {e}")
        return

    try:
        # Group records by species and locality
        df_grouped = df.groupby(
            ['ScientificName', 'Continent', 'Country', 'Locality'],
            as_index=False
        ).size()

    except Exception as e:

        print(f"Error while grouping data: {e}")
        return

    env = Environment()

    # Template for species pages
    species_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">

    <head>

        <meta charset="UTF-8">

        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">

        <title>{{ scientific_name }}</title>

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

            .gallery-button {
                margin-top: 20px;
                display: inline-block;
                padding: 10px 20px;
                background: #007bff;
                color: white;
                text-decoration: none;
                border-radius: 5px;
            }

        </style>

    </head>

    <body>

        <a href="home.html">
            <img src="home.png"
                 alt="Home"
                 class="home-icon">
        </a>

        <div class="container">

            <h1>{{ scientific_name }}</h1>

            <h2>Data</h2>

            <table class="data-table">

                <tr>
                    <th>Continent</th>
                    <th>Country</th>
                    <th>Locality</th>
                    <th>Individuals</th>
                </tr>

                {% for row in data %}

                <tr>
                    <td>{{ row.Continent }}</td>
                    <td>{{ row.Country }}</td>
                    <td>{{ row.Locality }}</td>
                    <td>{{ row.size }}</td>
                </tr>

                {% endfor %}

            </table>

            <a href="{{ gallery_filename }}"
               class="gallery-button">
               Gallery
            </a>

            <div class="logos">

                <img src="logo1.png"
                     alt="Left Logo">

                <img src="logo2.png"
                     alt="Right Logo">

            </div>

        </div>

    </body>

    </html>
    ''')

    # Template for gallery pages
    gallery_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">

    <head>

        <meta charset="UTF-8">

        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">

        <title>Gallery - {{ scientific_name }}</title>

        <style>

            .gallery {
                display: flex;
                flex-wrap: wrap;
                justify-content: center;
                gap: 10px;
            }

            .gallery img {
                width: 150px;
                height: auto;
                border: 1px solid #ccc;
            }

        </style>

    </head>

    <body>

        <h1>Gallery - {{ scientific_name }}</h1>

        <div class="gallery">

            {% for image in images %}

            <a href="{{ image.page_filename }}">
                <img src="{{ image.src }}" alt="">
            </a>

            {% endfor %}

        </div>

    </body>

    </html>
    ''')

    # Template for single-image pages
    image_template = env.from_string('''
    <!DOCTYPE html>
    <html lang="en">

    <head>

        <meta charset="UTF-8">

        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">

        <title>{{ image_name }}</title>

    </head>

    <body style="text-align:center">

        <h1>{{ image_name }}</h1>

        <img src="{{ src }}"
             style="max-width: 90%; height: auto;">

        <br><br>

        <a href="{{ back_link }}">
            Back to gallery
        </a>

    </body>

    </html>
    ''')

    # Create the output folder if it does not exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Copy logos and home icon if needed
    for icon, name in [
        (home_icon, 'home.png'),
        (left_logo, 'logo1.png'),
        (right_logo, 'logo2.png')
    ]:

        destination = os.path.join(output_folder, name)

        if not os.path.exists(destination):
            shutil.copy(icon, destination)

    # Generate HTML files
    for scientific_name, group in df_grouped.groupby('ScientificName'):

        scientific_name_sanitized = re.sub(
            r'[^a-zA-Z0-9_-]',
            '_',
            scientific_name
        ).lower()

        html_filename = f"{scientific_name_sanitized}.html"

        gallery_filename = (
            f"gallery_{scientific_name_sanitized}.html"
        )

        species_rows = df[
            df['ScientificName'] == scientific_name
        ]

        images = species_rows['Photo'].dropna().unique().tolist()

        image_entries = []

        for img in images:

            abs_path = os.path.join(image_folder, img)

            image_entries.append({
                'src': f"file:///{abs_path.replace(os.sep, '/')}",
                'page_filename':
                    f"img_{scientific_name_sanitized}_{img}.html",
                'filename': img
            })

        # Generate single-image pages
        for image in image_entries:

            image_html = image_template.render(
                image_name=image['filename'],
                src=image['src'],
                back_link=gallery_filename
            )

            with open(
                os.path.join(output_folder, image['page_filename']),
                'w',
                encoding='utf-8'
            ) as f:

                f.write(image_html)

        # Generate gallery page
        gallery_html = gallery_template.render(
            scientific_name=scientific_name,
            images=image_entries
        )

        with open(
            os.path.join(output_folder, gallery_filename),
            'w',
            encoding='utf-8'
        ) as f:

            f.write(gallery_html)

        # Generate species page
        species_html = species_template.render(
            scientific_name=scientific_name,
            data=group.to_dict(orient='records'),
            gallery_filename=gallery_filename
        )

        with open(
            os.path.join(output_folder, html_filename),
            'w',
            encoding='utf-8'
        ) as f:

            f.write(species_html)

    print(
        f"All HTML files were successfully created in: "
        f"{output_folder}"
    )

# File and folder selection
excel_file = select_file(
    "Select the Excel file",
    [("Excel files", "*.xlsx")]
)

output_folder = select_folder(
    "Select the output folder for saving HTML files"
)

image_folder = select_folder(
    "Select the folder containing all specimen images"
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
    image_folder,
    home_icon,
    left_logo,
    right_logo
]):

    try:

        create_html_pages(
            excel_file,
            output_folder,
            home_icon,
            left_logo,
            right_logo,
            image_folder
        )

    except Exception as e:

        print(
            f"\n❌ Error while generating HTML pages: {e}\n"
        )

else:

    print(
        "\n⚠️ Operation cancelled: one or more "
        "files/folders were not selected.\n"
    )

# Keep the console window open
input("Press any key to exit...")
