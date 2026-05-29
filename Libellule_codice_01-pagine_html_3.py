import pandas as pd
import os
import shutil
import re
from jinja2 import Environment
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory

# Function for selecting the input Excel file
def select_excel_file():
    root = Tk()
    root.withdraw()
    return askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx")])

# Function for selecting the directory containing specimen images
def select_image_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Select the folder containing the images")

# Function for selecting the directory containing logo files
def select_logo_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Select the folder containing logo files")

# Function for selecting the output directory where generated HTML files will be saved
def select_output_folder():
    root = Tk()
    root.withdraw()
    return askdirectory(title="Select the output folder for saving HTML files")

# Function to select the home page icon image
def select_home_icon():
    root = Tk()
    root.withdraw()
    return askopenfilename(title="Select the home icon (PNG)", filetypes=[("PNG files", "*.png")])

# Function for copying files with user overwrite prompt
def copy_file_with_prompt(src, dest):
    # Check if the source file exists
    if not os.path.exists(src):
        print(f"❌ ERROR: The file '{src}' does not exist! Skipping copy operation.")
        return

    # Verify that the source and destination files are not identical
    if os.path.abspath(src) == os.path.abspath(dest):
        print(f"ℹ️ The file '{src}' is already in the destination folder. No copy needed.")
        return

    print(f"🔄 copying from: {os.path.abspath(src)} to {os.path.abspath(dest)}")

    if os.path.exists(dest):
        response = messagebox.askyesno("Existing File", f"The file '{os.path.basename(dest)}' already exists.\nDo you want to overwrite it?"
)
        if not response:
            print(f"✅ File retained: {dest}")
            return

    shutil.copy(src, dest)
    print(f"✅ File copied: {dest}")



# Function to generate HTML pages
def create_html_pages(excel_file, image_folder, logo_folder, output_folder, home_icon):
    df = pd.read_excel(excel_file, engine="openpyxl")

    template_html = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Image Page</title>
        <style>
            .container { width: 80%; margin: 0 auto; text-align: center; position: relative; }
            .image { width: 100%; max-width: 600px; height: auto; }
            .data-table { width: 100%; margin: 20px 0; border-collapse: collapse; }
            .data-table th, .data-table td { border: 1px solid #ddd; padding: 8px; }
            .data-table th { background-color: #f2f2f2; }
            .italic { font-style: italic; }
            .title { font-size: 1.5em; }
            .logo-container { width: 100%; margin-top: 20px; display: flex; justify-content: space-between; }
            .logo { width: 100px; height: auto; }
            .home-icon { width: 40px; height: auto; position: absolute; top: 20px; right: 20px; }
        </style>
    </head>
    <body>
        <div class="container">
            <a href="home.html">
                <img src="home.png" alt="Home" class="home-icon">
            </a>
            <h1 class="italic title">{{ data['ScientificName'] }}</h1>
            <img src="{{ image }}" alt="Image" class="image">
            <h2>Data</h2>
            <table class="data-table">
                <tbody>
                    <tr><td>Kingdom:</td><td>{{ data['Kingdom'] }}</td></tr>
                    <tr><td>Phylum:</td><td>{{ data['Phylum'] }}</td></tr>
                    <tr><td>Class:</td><td>{{ data['Class'] }}</td></tr>
                    <tr><td>Order:</td><td>{{ data['Order'] }}</td></tr>
                    <tr><td>Family:</td><td>{{ data['Family'] }}</td></tr>
                    <tr><td>Genus:</td><td class="italic">{{ data['Genus'] }}</td></tr>
                    <tr><td>Species:</td><td class="italic">{{ data['Species'] }}</td></tr>
                    <tr><td>Identifier:</td><td>{{ data['Identifier'] }}</td></tr>
                    <tr><td>Continent:</td><td>{{ data['Continent'] }}</td></tr>
                    <tr><td>Country:</td><td>{{ data['Country'] }}</td></tr>
                    <tr><td>Location:</td><td>{{ data['Locality'] }}</td></tr>
                    <tr><td>Coordinates:</td><td>{{ data['Coordinates'] }}</td></tr>
                    <tr><td>GBIF:</td><td>{{ data['OccurrenceID'] }}</td></tr>
                </tbody>
            </table>
            <div class="logo-container">
                <img src="logo1.png" alt="Logo 1" class="logo" style="align-self: flex-start;">
                <img src="logo2.png" alt="Logo 2" class="logo" style="align-self: flex-end;">
            </div>
        </div>
    </body>
    </html>
    '''

    env = Environment()
    template = env.from_string(template_html)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    print(f"Output folder: {output_folder}")

    # Copy logos and the home icon with overwrite confirmation
    copy_file_with_prompt(os.path.join(logo_folder, 'logo1.png'), os.path.join(output_folder, 'logo1.png'))
    copy_file_with_prompt(os.path.join(logo_folder, 'logo2.png'), os.path.join(output_folder, 'logo2.png'))
    home_dest = os.path.join(output_folder, 'home.png')

    # Avoid copying the file if it is already in the destination folder
    if os.path.abspath(home_icon) != os.path.abspath(home_dest):
        copy_file_with_prompt(home_icon, home_dest)
    else:
        print(f"ℹ️ The file home.png is already in the destination folder. No copy needed.")

        logo1_path = "logo1.png"
        logo2_path = "logo2.png"
        home_icon = "home.png"


    # Loop through DataFrame rows and generate individual HTML files
    for index, row in df.iterrows():
        if pd.isna(row['Foto']):
            continue

        image_file = os.path.join(image_folder, row['Foto'])
        if not os.path.exists(image_file):
            print(f"❌ Image not found: {image_file}")
            continue

        # Relative path for the HTML file (without copying)
        image_path = row['Foto']

        data = row.to_dict()
        data['Coordinate'] = row['Coordinate'] if pd.notna(row['Coordinate']) else "Not available"

        html_content = template.render(
        image=image_path, 
        data=data, 
        logo1=logo1_path, 
        logo2=logo2_path, 
        home_icon=home_icon
)

        image_name = os.path.splitext(row['Foto'])[0]
        safe_image_name = re.sub(r'[^a-zA-Z0-9_-]', '_', image_name)
        output_file = os.path.join(output_folder, f"{safe_image_name}.html")

        print(f"✅ Generating HTML files: {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

    print(f'🎉 HTML pages created in {output_folder}')



# File and folder selection
excel_file = select_excel_file()
image_folder = select_image_folder()
logo_folder = select_logo_folder()
output_folder = select_output_folder()
home_icon = select_home_icon()

if excel_file and image_folder and logo_folder and output_folder and home_icon:
    create_html_pages(excel_file, image_folder, logo_folder, output_folder, home_icon)
else:
    print("Operation canceled: one or more files/folders were not selected.")
