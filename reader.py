# -*- coding: utf-8 -*-
import openpyxl
from pathlib import Path
import csv


NoneType = type(None)     

# Ouvrir le fichier
# Préciser le chemin du fichier source
xlsx_file = Path('/home/houssam/Bureau/data.xlsx')

# tester si le fichier est bien ouvert en affichant son chemin
print(xlsx_file)
print()

# chager le fichier .xlsx
wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active

# préciser ou le fichier .csv va etre créé, chemin + nom, a ne pas toucher le w = mode ecriture (stands for write)
file = open('/home/houssam/Bureau/reults.csv', 'w')
writer = csv.writer(file)


#Préciser le point de départ en min_row = y et min_col = x
for row in sheet.iter_rows(min_row=8, min_col=2):
    line = []
    for cell in row:
        # vérifier si le champ est de type date pour le formater en dd-mm-yyyy
        # pour modification de séparateur remplacer le '-' par ce que vous voulez
        if cell.column == 4 and cell.row != 8 and cell.value is not None:
            print((cell.value).strftime("%d-%m-%Y"), end=",")
            line.append((cell.value).strftime("%d-%m-%Y"))
        else:
            print(cell.value, end=",")
            line.append(cell.value)  
    print()
    writer.writerow(line)

# Fermer le flux du fichier csv
file.close()