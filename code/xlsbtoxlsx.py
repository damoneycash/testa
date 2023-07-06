from pyxlsb import open_workbook as open_xlsb
from openpyxl import Workbook

def convert_xlsb_to_xlsx(input_file, output_file):
    try:
        with open_xlsb(input_file) as wb_xlsb:
            wb_xlsx = Workbook()
            for sheetname in wb_xlsb.sheets:
                with wb_xlsb.get_sheet(sheetname) as sheet_xlsb:
                    ws_xlsx = wb_xlsx.create_sheet(title=sheetname)
                    for row in sheet_xlsb.rows():
                        ws_xlsx.append([item.v for item in row])
        wb_xlsx.save(output_file)
        print(f"Le fichier {input_file} a été converti avec succès en {output_file}")
    except Exception as e:
        print(f"Une erreur s'est produite lors de la conversion du fichier {input_file}:")
        print(str(e))

# Exemple d'utilisation
input_file_path = "D:\IMPORTANT\AUTO_ENTREPRISE\MISSIONS\INRAE\APP_EXCEL\EXCEL\ps(168).xlsb"  # Remplacez par le chemin vers votre fichier XLSB
output_file_path = "D:\IMPORTANT\AUTO_ENTREPRISE\MISSIONS\INRAE\APP_EXCEL\EXCEL\ps(168).xlsx"  # Remplacez par le chemin de sortie pour le fichier XLSX converti

convert_xlsb_to_xlsx(input_file_path, output_file_path)