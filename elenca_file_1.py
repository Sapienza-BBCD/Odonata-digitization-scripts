import os
from openpyxl import Workbook
from tkinter import Tk, filedialog

def scegli_cartella():
    root = Tk()
    root.withdraw()  # Nasconde la finestra principale
    cartella = filedialog.askdirectory(title="Seleziona la cartella")
    return cartella

def main():
    # Selezione cartella tramite finestra
    cartella = scegli_cartella()
    if not cartella:
        print("Nessuna cartella selezionata.")
        return

    # Scelta tipo file
    print("Scegli il tipo di file da elencare:")
    print("1 - HTML")
    print("2 - JPG")
    print("3 - PNG")
    print("4 - XLS / XLSX")
    print("5 - DOC / DOCX")
    print("6 - Tutti i file")

    scelta = input("Inserisci il numero della scelta: ")

    estensioni = {
        "1": (".html", ".htm"),
        "2": (".jpg", ".jpeg"),
        "3": (".png",),
        "4": (".xls", ".xlsx"),
        "5": (".doc", ".docx"),
        "6": None  # Tutti i file
    }

    if scelta not in estensioni:
        print("Scelta non valida.")
        return

    # Crea Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Elenco file"

    row = 1
    for nome_file in os.listdir(cartella):
        percorso_file = os.path.join(cartella, nome_file)
        if not os.path.isfile(percorso_file):
            continue

        if estensioni[scelta]:
            if not nome_file.lower().endswith(estensioni[scelta]):
                continue

        ws.cell(row=row, column=1, value=nome_file)
        row += 1

    percorso_output = os.path.join(cartella, "Elenco_file.xlsx")
    wb.save(percorso_output)

    print(f"Elenco creato con successo: {percorso_output}")

if __name__ == "__main__":
    main()

