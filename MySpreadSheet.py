from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
import Messages
import string_compares
import file_constants


class MySpreadSheet:

    def __init__(self, spread: Workbook):
        self.doc = spread
        return

    def set_values_to_document(self, sheet_name: str, full_data: dict):
        '''Setea los valores en el documento'''
        print(Messages.MESS_WRITTING)

        # declaramos para poder pillar autocomplete, como guardamos los valores en el mismo sheet
        # que hemos abierto, pillamos las cosas de aquí
        sheet: Worksheet
        sheet_name = self.__get_real_name(sheet_name)
        sheet = self.doc[sheet_name]
        self.__populate_sheet(sheet, full_data)

    def __populate_sheet(self, sheet: Workbook, full_data: dict) -> None:
        '''Rellena las celdas del sheet con los valores del diccionario.
           Tiene el control de errores para tratar con las idisiosincrasias del programa'''
        # el diccionario tiene todos los datos... esto lo pondremos en t
        cell: Cell
        for cell in full_data:
            index = 0
            # vale, esto imprime los valores por columnas
            for data in full_data[cell]:

                # sumamos 1 porque el cell que pillamos es el del valor por defecto
                sheet.cell(cell.row, cell.column + index + 1).value = data
                print("Data written: ", data)
                index += 1

            # pillamos la última celda porque, depende del valor, ponemos una observación
            # como los datos no están unificados tenemos que hacer toda esta magia.
            # si es -1 no hay valor, si es <0 tiene poco volumen y si es 0 es epicentre
            last_value = sheet.cell(cell.row, cell.column + file_constants.MAX_VARS).value
            observation_mess = self.__observation_message(last_value)
            sheet.cell(cell.row, cell.column + file_constants.MAX_VARS + 1, observation_mess)
            Messages.print_full(cell.value, full_data[cell][0], full_data[cell][1], full_data[cell][2], full_data[cell][3])

    def __observation_message(self, value_to_check) -> str:
        observation_mess =""
        if value_to_check == -1 or value_to_check is None:
            observation_mess = file_constants.NO_VOL_NAME
        elif value_to_check < 0:
            observation_mess = file_constants.POC_NAME_FULL
        elif value_to_check == 0:
            observation_mess = file_constants.SALIVA_EPI
        else:
            observation_mess = "Succes" # lo ponemos para debuggear
        return observation_mess

    def get_ids_form_sheet_name(self, sheet_name: str, cell_value: str) -> (list, bool):
        '''Pilla los valores que hemos dado desde una hoja a la que hemos dado el nombre.
           como nombre usaremos por defecto file_constatns.SHEET_NAME, en vez de hacerlo
           como constante, lo usamos como variable en la función, por si hay que cambiarlo'''
        name = self.__get_real_name(sheet_name)
        try:
            sheet = self.doc[name]
        except:
            return (None, False)

        # si no hay nombre, devolvemos lista vacía y un false
        if name == file_constants.NO_NAME:
            return (None, False)

        # y aquí rellenamos
        valid_cells = []
        (row, col, isValue) = self.__get_row_column_by_value(cell_value, sheet)
        row += 1
        print("col=" + cell_value)
        for sheet_row in sheet.iter_rows(min_row=row, min_col=col, max_col=col):
            if sheet_row is not None:
                valid_cells.append(sheet_row)

        return (valid_cells, True)

    # TODO cambiar nombre de la función por __get_real_sheet_name

    def __get_real_name(self, sheet_name: str) -> str:
        '''devuelve el nombre verdadero de la hoja hace una comparación entre el nombre de la hoja en minúscula
           y el que usamos nosotros en minúscula para luego devolver el nombre verdadero a partir del que hemos dado
           ej: nombre_hoja = "HoLa"; nombre del argumento = "Hola"
           así, esto devuelve  "HoLa"'''
        real_name = file_constants.NO_NAME

        for name in self.doc.sheetnames:
            if string_compares.equal_strings_not_case(name, sheet_name):
                return name

        return real_name

    def __get_row_column_by_value(self, cell_value: str, sheet) -> (int, int, bool):
        '''devuelve la fila y la columna que tenga el nombre que hemos pasado.
        OJU! si hay varias celdas con el mismo valor, devuelve la primera de ellas
        :param cell_value: el nombre de la celda
        :param sheet: hoja a mirar
        :return: fila, columna y si tiene valor
        '''
        col = 0
        row = 0
        is_value = False

        # por cada tupla de celdas en una sheet (cell_tuple)
        for cell_tp in sheet.rows:

            # por cada celda singular dentro de la tupla (cell_singular)
            for cell_sn in cell_tp:
                name = str(cell_sn.value).lower()
                if name == cell_value:
                    # si los nombres coinciden, pasamos las coordenadas
                    return (cell_sn.row, cell_sn.column, True)

        # también devolvemos la columna
        return (row, col, is_value)

    # pilla la lista completa de lo que nos interesa, esta es la función principal
    # metemos en un dict la id que hemos metido como clave
    # y como valor es una tupla con (BMCODE, tipo extracción [sangre, ...] número extraccón [e1, e2, ...], y volumen)
    def get_complete_list(self, sheet, ids: list) -> dict:
        final_dict = {}

        # primero, pillamos la columna bm
        (row, bmCol, isTrue) = self.__get_row_column_by_value(file_constants.CELL_BM_NAME, sheet)

        # ahora pillamos la columna del esquimot
        (row, hcCol, isTrue) = self.__get_row_column_by_value(file_constants.CELL_HC_NAME, sheet)

        # ahora pillamos la columnda del sample
        (row, samCol, isTrue) = self.__get_row_column_by_value(file_constants.CELL_SAMPLE_NAME, sheet)

        # ahora columna bis
        (row, bisCol, isTrue) = self.__get_row_column_by_value(file_constants.CELL_BIS_NAME, sheet)

        #ahora el número de las columnas donde estan las extracciones
        extract_cols = self.__get_all_ng_columns(sheet, file_constants.CELL_NG_NAME)

        # ahora, por cada fila de la columna del esquimot, miramos si el valor es igual a cualquiera de la lista
        # sumamos uno al row porque doy por hecho que la primera fila es la del título.
        for cells in sheet.iter_rows(min_row=row+1, min_col=hcCol, max_col=hcCol):
            # primer bucle, miramos los ids del esquimot
            # esto pilla la mierda del esquimot (hc), si el valor es nulo, pasamos porque no nos interesa
            if cells[0].value is None:
                continue

            for myIds in ids:
                if myIds[0].value == cells[0].value:
                    # declaramos celdas para bm, sample y si hay bisBM (así tenemos autocomplete)
                    bmCell: Cell
                    samCell: Cell
                    bisCell: Cell
                    # pillamos el bm y el tipo de extracción
                    bmCell = sheet.cell(cells[0].row, bmCol)
                    samCell = sheet.cell(cells[0].row, samCol)

                    # ahora hacemos una tupla con el nombre de extracción (e1, e2,...) y el valor
                    (ext_name, ext_value) = self.__get_best_extraction(sheet, cells[0].row, extract_cols)

                    # miramos si tiene "BM Bis Code" para saber si está repetido
                    bisCell = sheet.cell(cells[0].row, bisCol)

                    # printeamos lo que hemos pillado
                    Messages.print_full(myIds[0].value, bmCell.value, samCell.value, ext_name, ext_value, (bisCell.value is not None))

                    # si bisCell tiene valor, significa que este HC tiene varios bm, así que miramos si está en el diccionario
                    if bisCell.value is not None:
                        # si esta en el dicc, comprobamos el valor de la extracción
                        # TODO meter un tipo de constantes para esto o algo así, estaría bien porque no me entero
                        if final_dict.__contains__(myIds[0]) and final_dict[myIds[0]][3] > ext_value:
                            continue

                        #end
                    #end

                    final_dict[myIds[0]] = (bmCell.value, samCell.value, ext_name, ext_value)

        return final_dict

    # devuelve la mejor extracción de las posibles, con su identificador (e1, ...) y su valor
    # OJUUUUUU!!! este método es propenso a bugs por todas las excepciones que hay en las filas [ex] ya que,
    # aún siendo númericas (o deberían serlo) hay bastante texto que debería estar en otro sitio.
    # Necesita bastante comprobación por ahora y en caso de modificar el excel.
    # TODO cambiar partes del excel, meter textos del excel como constantes a seleccionar.
    def __get_best_extraction(self, sheet, row: int, columns: list) -> (str, int):
        # iniciamos tupla, así es más facil
        (name, best_value) = (file_constants.NO_VOL_NAME, file_constants.NO_VOL)
        index = 0
        #por cada columna en la que pillemos algo (oju, esto lo podriamos hacer de otro modo pero lo dejamos así)
        for col in columns:
            index += 1
            tempCell = sheet.cell(row, col)
            print(f"Value {index}: {tempCell.value}")
            # si no hay valor seguimos.
            # realmente antes había un break y parábamos, pero prefiero continuar
            # para pillar algún error, en plan, extracción e2 no tiene valor pero e3 sí
            if tempCell.value is None:
                continue
            elif type(tempCell.value) is str:
                if best_value == file_constants.NO_VOL and string_compares.file_constants.SALIVA_EPI.lower() in tempCell.value.lower():
                    (name, best_value) = file_constants.CELL_NG_NAME + str(index), file_constants.EPI_VOL
                continue
            elif sheet.cell(row, col - 1).value is not None:
                # pillamos la columna anterior al "ng/ul E1" que debería ser "VOL. NO DISPONIBLE"
                # sin embargo, en algunas extracciones esa columna no existe asi que HAY QUE CREARLA!!!... Ju_u
                best_value_name = sheet.cell(row, col - 1).value

                # si no es texto, da igual y seguimos, esta comprobación es importante porque, como he dicho arriba
                # a veces hay números u otras cosas porque no funca bien la columna
                if type(best_value_name) is str:
                    best_value_name = best_value_name.lower()
                    # si la cadena incluye "poc" comprobamos valores
                    # metemos como absoluto el valor
                    if file_constants.POC_NAME.lower() in best_value_name:
                        if tempCell.value > abs(best_value):
                            (name, best_value) = file_constants.CELL_NG_NAME + str(index), - tempCell.value
                        #end
                    #end
                #end

            elif tempCell.value > best_value:
                (name, best_value) = file_constants.CELL_NG_NAME + str(index), tempCell.value

        print(f"Best value: {best_value}")
        return (name, best_value)

    def __get_all_ng_columns(self, sheet, name) -> list:
        '''
        pilla los index de todas las columans que tienen el nombre seleccionado
        :param sheet: hoja donde pillamos los datos, cambiar
        :param name: nombre lo que nos interesa
        :return: lista con los índices
        '''
        index = file_constants.EXTRACTION_FIRST
        has_data = True
        columns = []

        while has_data and index < file_constants.MAX_VARS:
            cell_value = name + str(index)
            print("data: ", cell_value)
            (row, col, hasData) = self.__get_row_column_by_value(cell_value, sheet)
            index += 1
            if hasData:
                columns.append(col)

        return columns
