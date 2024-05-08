import win32com.client
import os
import re
from pathlib import Path
import time

def _set_print_area(sheet, print_area):
    sheet.PageSetup.PrintArea = print_area

def _set_fit_to_page(sheet):
    sheet.PageSetup.FitToPagesTall = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.Zoom = False
    #sheet.PageSetup.FitToPage = True

def _center_content(sheet):
    used_range = sheet.UsedRange
    used_range.VerticalAlignment = win32com.client.constants.xlVAlignCenter

def _set_page_size(sheet, paper_size, orientation):
    sheet.PageSetup.PaperSize = paper_size
    if orientation == "Landscape":
        sheet.PageSetup.Orientation = win32com.client.constants.xlLandscape
    elif orientation == "Portrait":
        sheet.PageSetup.Orientation = win32com.client.constants.xlPortrait
 
def excel_to_pdf(output_path: str | Path, tipicidad: str, type_data: str, excel_path: str, print_areas: list, sheet_names: list, page_sizes: list, orientations: list) -> None:
    excel_name = os.path.split(excel_path)[1]
    pattern1 = r"([A-Z]+[0-9+])"
    pattern2 = r"([A-Z]+-[0-9]+)"

    if re.search(pattern1, excel_name):
        code = re.search(pattern1, excel_name)[0]
    elif re.search(pattern2, excel_name):
        code = re.search(pattern2, excel_name)[0]

    excel = win32com.client.Dispatch("Excel.Application")
    #Visibility:
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    try:
        wb = excel.Workbooks.OpenXML(excel_path) #Open -> OpenXML
    except Exception as e:
        return print("Error: ", e)

    for i, sheet_name in enumerate(sheet_names):
        try:
            sheet = wb.Sheets(sheet_name)
        except Exception as e:
            print("No se encontró la hoja solicitada: ", sheet_name)
            print("Error: ", e)

        #sheet.Select()
        sheet.Activate()
        print_area = print_areas[i]

        if print_area:
            _set_print_area(sheet, print_area)

        _set_fit_to_page(sheet)
        _center_content(sheet)
        if page_sizes[i] == "A4":
            _set_page_size(sheet, win32com.client.constants.xlPaperA4, orientations[i])
        elif page_sizes[i] == "A3":
            _set_page_size(sheet, win32com.client.constants.xlPaperA3, orientations[i])
        
        print("Contenido: ", sheet_name, print_area, tipicidad)
        final_name = f"{code}_"+sheet_name+tipicidad+".pdf"
        output_filaname = output_path / type_data / final_name

        if not os.path.exists(output_path / type_data):
            os.makedirs(output_path / type_data)
        wb.ActiveSheet.ExportAsFixedFormat(0, str(output_filaname))

    wb.Close(False)
    excel.ScreenUpdating = True
    excel.DisplayAlerts = True
    excel.EnableEvents = True
    excel.Quit()
    return None

def apply_with_tipicidad(anexos_path: Path | str, type_path: str, type_name: str, configuration: dict, extension: str) -> None:
    for tipicidad in ["Tipico", "Atipico"]:
        folder_path = type_path / tipicidad
        if not os.path.exists(folder_path):
            return print("No existe la carpeta: ", folder_path)
        
        listExcels = os.listdir(folder_path)
        listExcels = [file for file in listExcels 
                      if file.endswith(extension) and not file.startswith("~$")]
        
        for excel in listExcels:
            excelPath = folder_path / excel
            print("Excel: ", excel)
            if tipicidad == "Tipico":
                sufix = "_T"
            elif tipicidad == "Atipico":
                sufix = "_A"
            try:
                excel_to_pdf(
                    output_path=    anexos_path,
                    tipicidad=      sufix,
                    type_data=      type_name,
                    excel_path=     excelPath,
                    print_areas=    configuration[type_name]["printAreas"],
                    sheet_names=    configuration[type_name]["sheetNames"],
                    page_sizes=     configuration[type_name]["pageSizes"],
                    orientations=   configuration[type_name]["orientations"],
                )
            except Exception as e:
                print("Error: ", e)
            
            time.sleep(3)

def apply_without_tipicidad(anexos_path: Path | str, type_path: str, type_name: str, configuration: dict, extension: str) -> None:
    folder_path = type_path
    if not os.path.exists(folder_path):
        return print("No existe la carpeta: ", folder_path)
    
    listExcels = os.listdir(folder_path)
    listExcels = [file for file in listExcels
                  if file.endswith(extension) and not file.startswith("~$")]
    
    for excel in listExcels:
        excelPath = folder_path / excel
        print("Excel: ", excel)
        try:
            excel_to_pdf(
                output_path=    anexos_path,
                tipicidad=      "",
                type_data=      type_name,
                excel_path=     excelPath,
                print_areas=    configuration[type_name]["printAreas"],
                sheet_names=    configuration[type_name]["sheetNames"],
                page_sizes=     configuration[type_name]["pageSizes"],
                orientations=   configuration[type_name]["orientations"],
            )
        except Exception as e:
            print("Error: ", e)
        
        time.sleep(3)
