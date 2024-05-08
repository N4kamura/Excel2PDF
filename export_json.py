import json

configuration = {
    "Vehicular": {
        "printAreas": [
            "J9:HB112","J9:HB112","J9:HB112","J9:HB112",
            "C2:BR106","C2:BR106","C2:BR106",
            "M6:W37",
        ],
        "sheetNames": [
            "N","S","E","O",
            "V_Ma","V_Ta","V_No",
            "Histograma",
        ],
        "pageSizes": [
            "A3","A3","A3","A3",
            "A4","A4","A4",
            "A4",
        ],
        "orientations": [
            "Landscape","Landscape","Landscape","Landscape",
            "Portrait","Portrait","Portrait",
            "Landscape",
        ],
    },
    "Peatonal": {
        "printAreas": [
            "K8:UZ83",
            "C2:BC86","C2:BC86","C2:BC86",
            "P2:AF30",
        ],
        "sheetNames": [
            "Data Peatonal",
            "Turno 01","Turno 02","Turno 03",
            "Histograma",
        ],
        "pageSizes": [
            "A3",
            "A4","A4","A4",
            "A4",
        ],
        "orientations": [
            "Landscape",
            "Portrait","Portrait","Portrait",
            "Landscape",
        ],
    },
    "Longitud de Cola": {
        "printAreas": [
            "B17:K59","M17:V59","AI17:AR59",
        ],
        "sheetNames": [
            "Base Data","Base Data","Base Data",
        ],
        "pageSizes": [
            "A4","A4","A4",
        ],
        "orientations": [
            "Portrait","Portrait","Portrait",
        ],
    },
    "Embarque y Desembarque": {
        "printAreas": [
            "D10:N69","D71:N130","D132:N191","D193:N252","D254:N313","D315:N374",
            "R10:AB69","R71:AB130","R132:AB191","R193:AB252","R254:AB313","R315:AB374",
            "AF10:AP69","AF71:AP130","AF132:AP191","AF193:AP252","AF254:AP313","AF315:AP374",
            "D10:N69","D71:N130","D132:N191","D193:N252","D254:N313","D315:N374",
            "R10:AB69","R71:AB130","R132:AB191","R193:AB252","R254:AB313","R315:AB374",
            "AF10:AP69","AF71:AP130","AF132:AP191","AF193:AP252","AF254:AP313","AF315:AP374",
            "D10:N69","D71:N130","D132:N191","D193:N252","D254:N313","D315:N374",
            "R10:AB69","R71:AB130","R132:AB191","R193:AB252","R254:AB313","R315:AB374",
            "AF10:AP69","AF71:AP130","AF132:AP191","AF193:AP252","AF254:AP313","AF315:AP374",
            "D10:N69","D71:N130","D132:N191","D193:N252","D254:N313","D315:N374",
            "R10:AB69","R71:AB130","R132:AB191","R193:AB252","R254:AB313","R315:AB374",
            "AF10:AP69","AF71:AP130","AF132:AP191","AF193:AP252","AF254:AP313","AF315:AP374",
        ],
        "sheetNames": [
            "Norte","Norte","Norte","Norte","Norte","Norte",
            "Norte","Norte","Norte","Norte","Norte","Norte",
            "Norte","Norte","Norte","Norte","Norte","Norte",
            "Sur","Sur","Sur","Sur","Sur","Sur",
            "Sur","Sur","Sur","Sur","Sur","Sur",
            "Sur","Sur","Sur","Sur","Sur","Sur",
            "Este","Este","Este","Este","Este","Este",
            "Este","Este","Este","Este","Este","Este",
            "Este","Este","Este","Este","Este","Este",
            "Oeste","Oeste","Oeste","Oeste","Oeste","Oeste",
            "Oeste","Oeste","Oeste","Oeste","Oeste","Oeste",
            "Oeste","Oeste","Oeste","Oeste","Oeste","Oeste",
        ],
        "pageSizes": [
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
            "A4","A4","A4","A4","A4","A4",
        ],
        "orientations": [
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
            "Portrait","Portrait","Portrait","Portrait","Portrait","Portrait",
        ],
    },
    "Tiempo de Ciclo Semaforico":{
        "printAreas": [
            "A1:W35",
        ],
        "sheetNames": [
            "Tiempo de Ciclo y Fases",
        ],
        "pageSizes": [
            "A4",
        ],
        "orientations": [
            "Portrait",
        ],
    },
}

with open('config.json', 'w') as json_file:
    json.dump(configuration, json_file, indent = 4)