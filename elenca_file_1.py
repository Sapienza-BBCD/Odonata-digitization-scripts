import os
from openpyxl import Workbook
from tkinter import Tk, filedialog

# Function to select a folder
def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window
    folder = filedialog.askdirectory(title="Select the folder")
    return folder

def main():
    # Folder selection through dialog window
    folder = select_folder()

    if not folder:
        print("No folder selected.")
        return

    # File type selection
    print("Choose the file type to list:")
    print("1 - HTML")
    print("2 - JPG")
    print("3 - PNG")
    print("4 - XLS / XLSX")
    print("5 - DOC / DOCX")
    print("6 - All files")

    choice = input("Enter the number corresponding to your choice: ")

    extensions = {
        "1": (".html", ".htm"),
        "2": (".jpg", ".jpeg"),
        "3": (".png",),
        "4": (".xls", ".xlsx"),
        "5": (".doc", ".docx"),
        "6": None  # All files
    }

    if choice not in extensions:
        print("Invalid choice.")
        return

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "File List"

    row = 1

    # Scan files in the selected folder
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)

        if not os.path.isfile(file_path):
            continue

        # Filter by selected extension
        if extensions[choice]:
            if not filename.lower().endswith(extensions[choice]):
                continue

        ws.cell(row=row, column=1, value=filename)
        row += 1

    # Save the Excel file
    output_path = os.path.join(folder, "File_List.xlsx")
    wb.save(output_path)

    print(f"File list successfully created: {output_path}")

if __name__ == "__main__":
    main()
