import Messages
import file_io
import file_constants
from MySpreadSheet import MySpreadSheet


debug = False

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

    if debug:
        input("DEBUG MODE!!! Uso del path por defecto! Presiona una tecla para continuar")
        path = r"C:\Users\manga\Downloads\Copia3_3.INVENTARI EXTRACCIONS_06.08.21_copiaNH.xlsx"
    else:
        path = input(">>")

    (spread, isCorrectFile) = file_io.get_workbook(path)

    if not isCorrectFile:
        print(Messages.PATH_ERROR_FILE)
    else:
        myDoc = MySpreadSheet(spread)

#pillamos las id's del libro que nos interesa

(myIds, isCorrectFile) = myDoc.get_ids_form_sheet_name(file_constants.SHEET_NAME, file_constants.CELL_NAME)

if(isCorrectFile == False):
    print(Messages.MESS_NOT_FOUND_LIST)
else:
    #si tiene el nombre correcto y hemos pillado id's, tonces ahora viene lo chungo-chungo, pillar las bm
    dict_full = myDoc.get_complete_list(myDoc.doc.worksheets[0], myIds)
    print(f"Diccionario:{dict_full}")
    #borramos el diccionario porque no nos interesa
    myIds.clear()

    if len(dict_full) == 0:
        print(Messages.MESS_NOT_FOUND_COLUMN)
    else:
        myDoc.set_values_to_document(file_constants.SHEET_NAME, dict_full)
        file_io.save_workbook(path, myDoc.doc)
del myDoc
input(Messages.PRESS_KEY_EXIT)


