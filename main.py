from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QButtonGroup, QErrorMessage, QMessageBox
from interfaces.main_interface import Ui_MainWindow
from tools import *
from pathlib import Path
import json

try:
    with open('./tools/config.json', 'r') as json_file:
        configuration = json.load(json_file) #TODO: Cambiar a una ruta especifica
except Exception as e:
    print("No se pudo cargar el archivo de configuración config.json: ")
    print("Error: ", e)

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.UI = Ui_MainWindow()
        self.UI.setupUi(self)

        #Butons:
        self.UI.openButton.clicked.connect(self.open_file)
        self.UI.startButton.clicked.connect(self.start_program)

    def open_file(self): #OK
        self.path_file = QFileDialog.getExistingDirectory(self, 'Open File')
        if self.path_file:
            self.UI.pathText.setText(self.path_file)

    def start_program(self):
        start_time = time.time()
        try:
            anexos_file = Path(self.path_file) / "Anexos"
        except AttributeError as e:
            error_message = QErrorMessage(self)
            return error_message.showMessage("Use Open File button first!")
        if not os.path.exists(anexos_file):
            os.makedirs(anexos_file)

        path_parts = self.path_file.split("/")
        project_path = '/'.join(path_parts[:-2])
        subarea_name = path_parts[-1]
        field_path = Path(project_path) / "7. Informacion de Campo" / subarea_name

        if not os.path.exists(field_path):
            return print(f"Falta la carpeta de campo: {field_path}")
        
        if self.UI.vehiclesCheck.isChecked():
            print("###### STARTING VEHICLES #####")
            vehicle_path = field_path / "Vehicular"
            apply_with_tipicidad(
                anexos_path=    anexos_file,
                type_path=      vehicle_path,
                type_name=      "Vehicular",
                configuration=  configuration,
                extension=      ".xlsm",
            )

        if self.UI.pedestrianCheck.isChecked():
            print("\n###### STARTING PEDESTRIAN #####")
            pedestrian_path = field_path / "Peatonal"
            apply_with_tipicidad(
                anexos_path=    anexos_file,
                type_path=      pedestrian_path,
                type_name=      "Peatonal",
                configuration=  configuration,
                extension=      ".xlsm",
            )

        if self.UI.queuesCheck.isChecked():
            print("\n###### STARTING QUEUES #####")
            queues_path = field_path / "Longitud de Cola"
            apply_with_tipicidad(
                anexos_path=    anexos_file,
                type_path=      queues_path,
                type_name=      "Longitud de Cola",
                configuration=  configuration,
                extension=      ".xlsx",
            )

        if self.UI.boardingCheck.isChecked():
            print("\n###### STARTING BOARDING #####")
            boarding_path = field_path / "Embarque y Desembarque"
            apply_with_tipicidad(
                anexos_path=    anexos_file,
                type_path=      boarding_path,
                type_name=      "Embarque y Desembarque",
                configuration=  configuration,
                extension=      ".xlsx",
            )

        if self.UI.cycleCheck.isChecked():
            print("\n###### STARTING CYCLE AND PHASES #####")
            cycle_path = field_path / "Tiempo de Ciclo Semaforico"
            apply_without_tipicidad(
                anexos_path=    anexos_file,
                type_path=      cycle_path,
                type_name=      "Tiempo de Ciclo Semaforico",
                configuration=  configuration,
                extension=      ".xlsx",
            )

        end_time = time.time()
        print("Total time: ", int(end_time - start_time), "seconds")
        self.UI.label.setText(f"STATE: Attachment files created successfully")

def main():
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()