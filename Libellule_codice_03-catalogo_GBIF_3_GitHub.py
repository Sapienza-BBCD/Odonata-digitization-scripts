import os
import json
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
from string import Template

# Create a hidden window for file selection
Tk().withdraw()

# Select the Excel file
file_path = askopenfilename(
    title="Select the Excel file",
    filetypes=[("Excel files", "*.xlsx")]
)

if not file_path:
    print("No file selected. The program will exit.")
    exit()

# Select the folder containing HTML files
html_folder_path = askdirectory(
    title="Select the folder containing HTML files"
)

if not html_folder_path:
    print("No folder selected. The program will exit.")
    exit()

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
    exit()

# Function to select the home icon
def select_home_icon():
    root = Tk()
    root.withdraw()

    return askopenfilename(
        title="Select the home icon (PNG)",
        filetypes=[("PNG files", "*.png")]
    )

# Select the home icon
home_icon_path = select_home_icon()

if not home_icon_path:
    print("Home icon not selected. The program will exit.")
    exit()

# Select the folder for saving the HTML file
save_folder_path = askdirectory(
    title="Select the folder for saving the HTML file"
)

if not save_folder_path:
    print("No output folder selected. The program will exit.")
    exit()

try:
    # Read data from the "Occurrences" worksheet
    occurrences_df = pd.read_excel(
        file_path,
        sheet_name='Occurrences'
    )

    # Remove extra spaces from column names
    occurrences_df.columns = occurrences_df.columns.str.strip()

    # Remove rows without OccurrenceID or Photo
    occurrences_df = occurrences_df.dropna(
        subset=['OccurrenceID', 'Photo']
    )

    # Create a dictionary mapping OccurrenceID to Photo
    occurrence_data = occurrences_df.set_index(
        'OccurrenceID'
    )['Photo'].to_dict()

except Exception as e:
    print(f"Error while reading the Excel file: {e}")
    exit()

# HTML template
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

    <!-- Home icon -->
    <img src="$home_icon"
         alt="Home"
         class="home-icon"
         onclick="location.href='index.html'" />

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

        <img src="$logo1" alt="Logo 1"/>
        <img src="$logo2" alt="Logo 2"/>

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

                resultElement.innerHTML =
                    "<a href='" + relativePath +
                    "' target='_self'>Open " +
                    htmlFile + "</a>";

            } else {

                resultElement.innerText = "Code not found";
            }
        }

    </script>

</body>

</html>
""")

# Generate number options for dropdown menus
numbers_options = "\n".join([
    f"<option value='{i}'>{i}</option>"
    for i in range(10)
])

# Replace placeholders in the HTML template
html_content = catalog_template.substitute(
    numbers=numbers_options,
    occurrences=json.dumps(occurrence_data),
    html_folder=html_folder_path.replace("\\", "/"),
    logo1=os.path.basename(logo1_path),
    logo2=os.path.basename(logo2_path),
    home_icon=os.path.basename(home_icon_path)
)

# Save the HTML file
catalogo_gbif_path = os.path.join(
    save_folder_path,
    'catalogo_gbif.html'
)

with open(catalogo_gbif_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(
    f"✅ HTML file 'catalogo_gbif.html' successfully created in: "
    f"{catalogo_gbif_path}"
)
