import Messages
import file_io
import file_constants
from MySpreadSheet import MySpreadSheet



#metemos bienvenida
print(Messages.WELCOME)

#metemos aviso
print(Messages.ADVISE)

#metemos el input para pillar la ruta
print(Messages.INSTRUCTIONS)

#script para abrir el archivo -> en otra clase?

isCorrectFile = False
while isCorrectFile == False:

    print(Messages.PATH_ENTER)
    path = input(">>")
    # OJUUUUUU DEBUG!!!
    # path = r"C:\Users\manga\Downloads\3. INVENTARI EXTRACCIONS_06.08.21_copiaNH.xlsx"
    # OJUUUUU DEBUG!!!!
    (spread, isCorrectFile) = file_io.get_workbook(path)

    if(isCorrectFile == False):
        print(Messages.PATH_ERROR_FILE)
    else:
        myDoc = MySpreadSheet(spread)

#pillamos las id's del libro que nos interesa

(myIds, isCorrectFile) = myDoc.get_ids_form_sheet_name(file_constants.SHEET_NAME, file_constants.CELL_NAME)

if(isCorrectFile == False):
    print("La hoja no tiene el nombre correcto")
else:
    #si tiene el nombre correcto y hemos pillado id's, tonces ahora viene lo chungo-chungo, pillar las bm
    dict_full = myDoc.get_complete_list(myDoc.doc.worksheets[0], myIds)
    myDoc.set_values_to_document(file_constants.SHEET_NAME, dict_full)
    file_io.save_workbook(path, myDoc.doc)


