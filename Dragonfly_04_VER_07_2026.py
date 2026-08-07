"""
Generate species summary pages and image galleries from a Darwin Core dataset.

The script reads specimen metadata and associatedMedia values from an Excel
workbook, groups occurrence records by scientific name and locality, copies
the required specimen images into a web-compatible folder, and creates:

- one summary page for each scientific name;
- one gallery page for each scientific name;
- direct links from each gallery thumbnail to the corresponding specimen
  record page (1.html, 2.html, 3.html, ...).

IMPORTANT
---------
The specimen-page numbering must follow the same spreadsheet row order used
by the specimen-page generator. Therefore, the first occurrence record is
linked to 1.html, the second to 2.html, and so on.
"""

import os
import re
import shutil
from urllib.parse import urlparse, unquote

import pandas as pd
from jinja2 import Environment, select_autoescape
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory


# ============================================================
# DARWIN CORE CONFIGURATION
# ============================================================

SHEET_NAME = "Occurrences"
SCIENTIFIC_NAME_COLUMN = "scientificName"
CONTINENT_COLUMN = "continent"
COUNTRY_COLUMN = "country"
LOCALITY_COLUMN = "locality"
MEDIA_COLUMN = "associatedMedia"
RECORD_PAGE_COLUMN = "_recordPage"

REQUIRED_COLUMNS = [
    SCIENTIFIC_NAME_COLUMN,
    CONTINENT_COLUMN,
    COUNTRY_COLUMN,
    LOCALITY_COLUMN,
    MEDIA_COLUMN,
]


# ============================================================
# FILE AND FOLDER SELECTION
# ============================================================

def select_file(title, filetypes):
    """Open a file-selection dialog and return the selected path."""
    root = Tk()
    root.withdraw()
    selected = askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return selected


def select_folder(title):
    """Open a folder-selection dialog and return the selected path."""
    root = Tk()
    root.withdraw()
    selected = askdirectory(title=title)
    root.destroy()
    return selected


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
    missing = [c for c in REQUIRED_COLUMNS if c not in dataframe.columns]
    if missing:
        raise ValueError(
            "Missing Darwin Core columns:\n\n"
            + "\n".join(f"• {column}" for column in missing)
        )


def sanitize_filename(value):
    """Create a filesystem-safe lowercase filename component."""
    value = normalize_text(value)
    value = re.sub(r"[^a-zA-Z0-9_-]", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.lower().strip("_") or "record"


def split_media(value):
    """Split associatedMedia values separated by | or ;."""
    value = normalize_text(value)
    if not value:
        return []
    return [item.strip() for item in re.split(r"[|;]", value) if item.strip()]


def media_basename(media_value):
    """Extract a filename from a local path or URL stored in associatedMedia."""
    media_value = normalize_text(media_value)
    if not media_value:
        return ""

    parsed = urlparse(media_value)
    if parsed.scheme in {"http", "https", "file"}:
        return os.path.basename(unquote(parsed.path).replace("\\", "/"))

    clean = media_value.split("?", 1)[0].split("#", 1)[0]
    return os.path.basename(clean.replace("\\", "/"))


def is_web_url(value):
    """Return True when a media reference is an HTTP or HTTPS URL."""
    value = normalize_text(value)

    if not value:
        return False

    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"}



def copy_asset(source_path, destination_path):
    """Copy a shared visual asset."""
    if not source_path or not os.path.isfile(source_path):
        raise FileNotFoundError(f"File not found: {source_path}")
    if os.path.abspath(source_path) == os.path.abspath(destination_path):
        return
    shutil.copy2(source_path, destination_path)



# ============================================================
# HTML PAGE GENERATION
# ============================================================

def create_html_pages(
    excel_file,
    website_folder,
    home_icon,
    left_logo,
    right_logo,
):
    """
    Generate species summary pages and galleries inside the website root.

    The website root must be the same folder that contains index.html and the
    numbered specimen pages (1.html, 2.html, 3.html, ...). This guarantees
    that all relative links generated by index.html remain valid.
    """

    print("\n=== Species and gallery generation started ===")
    print(f"Excel file: {excel_file}")
    print(f"Website folder: {website_folder}")
    print("Reading Darwin Core dataset...")

    df = pd.read_excel(
        excel_file,
        sheet_name=SHEET_NAME,
        engine="openpyxl",
        dtype=object,
    )
    df.columns = [str(column).strip() for column in df.columns]
    validate_columns(df)

    # Preserve the same progressive numbering used by the specimen-page
    # generator: first record -> 1.html, second -> 2.html, and so on.
    df[RECORD_PAGE_COLUMN] = [
        f"{record_number}.html"
        for record_number in range(1, len(df) + 1)
    ]

    df = df[REQUIRED_COLUMNS + [RECORD_PAGE_COLUMN]].copy()

    for column in REQUIRED_COLUMNS:
        df[column] = df[column].apply(normalize_text)

    df = df[df[SCIENTIFIC_NAME_COLUMN] != ""].copy()

    if df.empty:
        raise ValueError("No valid records were found after cleaning the Excel file.")

    df_grouped = (
        df.groupby(
            [
                SCIENTIFIC_NAME_COLUMN,
                CONTINENT_COLUMN,
                COUNTRY_COLUMN,
                LOCALITY_COLUMN,
            ],
            dropna=False,
            as_index=False,
        )
        .size()
    )


    env = Environment(
        autoescape=select_autoescape(
            enabled_extensions=("html", "xml"),
            default_for_string=True,
        )
    )

    species_template = env.from_string("""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ scientific_name }}</title>

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
 * Fixed upper band containing the Home icon and scientific name.
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
    font-style: italic;
    color: #821a33;
    font-size: 42px;
}

/*
 * Scrollable central content.
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
    width: min(1100px, 94%);
    margin: auto;
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

.gallery-button {
    display: inline-block;
    margin-top: 20px;
    padding: 11px 22px;
    background: #821a33;
    color: white;
    text-decoration: none;
    border-radius: 6px;
}

.gallery-button:hover {
    background: #c54c00;
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
        font-size: 32px;
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

<header class="fixed-header">

    <a href="index.html">
        <img src="home.png" alt="Home" class="home-icon">
    </a>

    <h1>{{ scientific_name }}</h1>

</header>

<main class="scroll-area">

    <div class="container">

        <h2>Collection summary</h2>

        <table class="data-table">

            <tr>
                <th>Continent</th>
                <th>Country</th>
                <th>Locality</th>
                <th>Individuals</th>
            </tr>

            {% for row in data %}
            <tr>
                <td>{{ row.continent }}</td>
                <td>{{ row.country }}</td>
                <td>{{ row.locality }}</td>
                <td>{{ row.size }}</td>
            </tr>
            {% endfor %}

        </table>

        <a href="{{ gallery_filename }}" class="gallery-button">
            Open gallery ({{ specimen_count }} specimens)
        </a>

    </div>

</main>

<footer class="fixed-footer">

    <div class="logos">
        <img src="logo1.png" alt="Left Logo">
        <img src="logo2.png" alt="Right Logo">
    </div>

</footer>

</body>
</html>
""")

    gallery_template = env.from_string("""
<!DOCTYPE html>
<html lang="en">
<head>

<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gallery - {{ scientific_name }}</title>

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
 * Fixed upper band containing the gallery title and navigation controls.
 */
.fixed-header {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    z-index: 1000;
    min-height: 118px;
    background: white;
    border-bottom: 1px solid #ddd;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    box-sizing: border-box;
    padding: 12px 20px 10px;
}

h1 {
    text-align: center;
    font-style: italic;
    color: #821a33;
    margin: 4px 0 10px;
    font-size: 38px;
}

.navigation {
    text-align: center;
    margin: 0;
}

.navigation a {
    display: inline-block;
    padding: 8px 14px;
    margin: 3px;
    background: #821a33;
    color: white;
    text-decoration: none;
    border-radius: 5px;
}

.navigation a:hover {
    background: #c54c00;
}

/*
 * Scrollable image gallery.
 */
.scroll-area {
    position: fixed;
    top: 118px;
    bottom: 92px;
    left: 0;
    right: 0;
    overflow-y: auto;
    overflow-x: hidden;
    -webkit-overflow-scrolling: touch;
}

.gallery {
    width: min(1300px, 96%);
    margin: 0 auto;
    padding: 28px 10px 40px;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 18px;
    box-sizing: border-box;
}

.gallery-item {
    width: 180px;
    text-align: center;
}

.gallery-item a {
    color: #222;
    text-decoration: none;
}

.gallery-item img {
    width: 180px;
    height: 180px;
    object-fit: contain;
    border: 1px solid #ccc;
    background: #fafafa;
}

.filename {
    margin-top: 6px;
    font-size: 13px;
    overflow-wrap: anywhere;
}

.empty-gallery {
    width: 100%;
    margin: 40px auto;
    text-align: center;
    color: #821a33;
    font-size: 18px;
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
        min-height: 108px;
        padding: 10px 12px 8px;
    }

    h1 {
        font-size: 30px;
    }

    .scroll-area {
        top: 108px;
        bottom: 80px;
    }

    .gallery-item {
        width: 145px;
    }

    .gallery-item img {
        width: 145px;
        height: 145px;
    }

    .logos img {
        width: 120px;
        max-height: 52px;
    }
}

</style>
</head>

<body>

<header class="fixed-header">

    <h1>Gallery - {{ scientific_name }}</h1>

    <div class="navigation">
        <a href="{{ species_filename }}">Back to species page</a>
        <a href="index.html">Home</a>
    </div>

</header>

<main class="scroll-area">

    <div class="gallery">

        {% if images %}

        {% for image in images %}

        <div class="gallery-item">

            <a href="{{ image.record_page }}">

                <img
                    src="{{ image.src }}"
                    alt="{{ image.filename }}"
                    loading="lazy"
                >

                <div class="filename">
                    {{ image.filename }}
                </div>

            </a>

        </div>

        {% endfor %}

        {% else %}

        <div class="empty-gallery">
            No associated images could be resolved for this taxon.
        </div>

        {% endif %}

    </div>

</main>

<footer class="fixed-footer">

    <div class="logos">
        <img src="logo1.png" alt="Left Logo">
        <img src="logo2.png" alt="Right Logo">
    </div>

</footer>

</body>
</html>
""")

    os.makedirs(website_folder, exist_ok=True)
    copy_asset(home_icon, os.path.join(website_folder, "home.png"))
    copy_asset(left_logo, os.path.join(website_folder, "logo1.png"))
    copy_asset(right_logo, os.path.join(website_folder, "logo2.png"))

    missing_pages = []

    species_count = 0
    gallery_count = 0

    total_species = df_grouped[SCIENTIFIC_NAME_COLUMN].nunique()
    current_species = 0

    for scientific_name, group in df_grouped.groupby(SCIENTIFIC_NAME_COLUMN):
        current_species += 1
        print(
            f"[{current_species}/{total_species}] "
            f"Generating species page and gallery: {scientific_name}"
        )
        safe_name = sanitize_filename(scientific_name)
        species_filename = f"{safe_name}.html"
        gallery_filename = f"gallery_{safe_name}.html"

        species_rows = df[df[SCIENTIFIC_NAME_COLUMN] == scientific_name]
        image_entries = []

        for _, row in species_rows.iterrows():
            record_page = normalize_text(row[RECORD_PAGE_COLUMN])

            if not os.path.isfile(os.path.join(website_folder, record_page)):
                missing_pages.append(f"{record_page} — {scientific_name}")

            for media_value in split_media(row[MEDIA_COLUMN]):
                filename = media_basename(media_value)

                # Use absolute web URLs directly.
                if is_web_url(media_value):
                    image_source = media_value
                else:
                    # Use the existing image in the website root.
                    # No duplicate copy and no gallery_images folder.
                    image_source = filename or normalize_text(media_value)

                if not image_source:
                    continue

                image_entries.append({
                    "src": image_source,
                    "record_page": record_page,
                    "filename": filename or image_source,
                })

        # Remove duplicate image/page pairs while preserving order.
        unique_entries = []
        seen = set()
        for entry in image_entries:
            key = (entry["record_page"], entry["src"])
            if key not in seen:
                seen.add(key)
                unique_entries.append(entry)

        print(
            f"    Records for taxon: {len(species_rows)} | "
            f"Gallery images generated: {len(unique_entries)}"
        )

        gallery_html = gallery_template.render(
            scientific_name=scientific_name,
            images=unique_entries,
            species_filename=species_filename,
        )
        with open(
            os.path.join(website_folder, gallery_filename),
            "w",
            encoding="utf-8",
        ) as file:
            file.write(gallery_html)
        gallery_count += 1

        species_html = species_template.render(
            scientific_name=scientific_name,
            data=group.to_dict(orient="records"),
            gallery_filename=gallery_filename,
            specimen_count=len(species_rows),
        )
        with open(
            os.path.join(website_folder, species_filename),
            "w",
            encoding="utf-8",
        ) as file:
            file.write(species_html)
        species_count += 1

    print(f"✅ Species pages created: {species_count}")
    print(f"✅ Gallery pages created: {gallery_count}")
    print(f"✅ Website folder: {website_folder}")
    print("=== Generation completed ===\n")

    warnings = []

    if missing_pages:
        items = list(dict.fromkeys(missing_pages))
        warnings.append(
            "Specimen HTML pages not found:\n"
            + "\n".join(f"• {x}" for x in items[:20])
        )

    if warnings:
        messagebox.showwarning(
            "Completed with warnings",
            "Species pages and galleries were generated, but some resources "
            "could not be verified.\n\n" + "\n\n".join(warnings),
        )
    else:
        messagebox.showinfo(
            "Operation completed",
            "All species pages and galleries were created successfully.\n\n"
            f"Species pages: {species_count}\n"
            f"Gallery pages: {gallery_count}\n\n"
            f"Output folder:\n{website_folder}",
        )


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    """Run the graphical workflow for species-page generation."""
    excel_file = select_file(
        "Select the Excel file",
        [("Excel files", "*.xlsx")],
    )
    if not excel_file:
        return

    website_folder = select_folder(
        "Select the WEBSITE ROOT folder containing index.html and numbered specimen pages"
    )
    if not website_folder:
        return

    index_path = os.path.join(website_folder, "index.html")
    if not os.path.isfile(index_path):
        raise FileNotFoundError(
            "index.html was not found in the selected website folder.\n\n"
            "Select the same folder that contains index.html, 1.html, 2.html, etc."
        )

    image_types = [
        ("Image files", "*.png *.jpg *.jpeg *.gif *.webp *.tif *.tiff")
    ]

    home_icon = select_file("Select the Home icon", image_types)
    left_logo = select_file("Select the left logo", image_types)
    right_logo = select_file("Select the right logo", image_types)

    if not all([home_icon, left_logo, right_logo]):
        return

    try:
        create_html_pages(
            excel_file,
            website_folder,
            home_icon,
            left_logo,
            right_logo,
        )
    except Exception as error:
        print(f"❌ Error: {error}")
        messagebox.showerror("Error", str(error))


if __name__ == "__main__":
    main()

