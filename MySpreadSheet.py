
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell


import Messages
import file_constants


class MySpreadSheet:

    def __init__(self, spread: Workbook):
        self.doc = spread
        return

    def set_values_to_document(self, sheet_name: str, full_data: dict):
        print(Messages.MESS_WRITTING)

        sheet: Worksheet
        sheet = self.doc[sheet_name]

        cell: Cell
        for cell in full_data:
            for i in range(len(full_data[cell])):
                sheet.cell(cell.row, cell.column + i + 1).value = full_data[cell][i]

            Messages.print_full(cell.value, full_data[cell][0], full_data[cell][1], full_data[cell][2], full_data[cell][3])
        return

    def get_ids_form_sheet_name(self, sheet_name: str, cell_value: str) -> (list, bool):
        try:
            sheet = self.doc[sheet_name]
        except:
            return  (None, False)

        valid_cells = []
        (row, col, isValue) = self.__get_row_column_by_value(cell_value, sheet)
        row += 1
        print("col=" + cell_value)
        for sheet_row in sheet.iter_rows(min_row=row, min_col=col, max_col=col):
            if sheet_row is not None:
                valid_cells.append(sheet_row)

        return (valid_cells, True)

    def __get_row_column_by_value(self, cell_value: str, sheet) -> (int, int, bool):
        col = 0
        row = 0
        isValue = False

        # por cada tupla de celdas en una sheet (cell_tuple)
        for cell_tp in sheet.rows:
            if isValue == True:
                break

            # por cada celda singular dentro de la tupla (cell_singular)
            for cell_sn in cell_tp:
                name = str(cell_sn.value).lower()
                if name == cell_value:
                    isValue = True
                    col = cell_sn.column
                    row = cell_sn.row
                    break

        # también devolvemos la columna
        # TODO cambiar el nombre de la función
        return (row, col, isValue)

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
                    (ext_name, ext_value) = self.__get_best_extraction(sheet, cells[0].row)

                    # miramos si tiene "BM Bis Code" para saber si está repetido
                    bisCell = sheet.cell(cells[0].row, bisCol)

                    # printeamos lo que hemos pillado
                    Messages.print_full(myIds[0].value, bmCell.value, samCell.value, ext_name, ext_value, (bisCell.value is None) == False)

                    # si bisCell tiene valor, significa que este HC tiene varios bm, así que miramos si está en el diccionario
                    if bisCell.value is not None:
                        # si esta en el dicc, comprobamos el valor de la extracción
                        # TODO meter un tipo de constantes para esto o algo así
                        if final_dict.__contains__(myIds[0]) and final_dict[myIds[0]][3] > ext_value:
                            continue
                        #end
                    #end

                    final_dict[myIds[0]] = (bmCell.value, samCell.value, ext_name, ext_value)

        return final_dict

    def __get_best_extraction(self, sheet, row: int) -> (str, int):
        # pilla los índices de las columnas
        columns = self.__get_all_ng_columns(sheet, file_constants.CELL_NG_NAME)
        # iniciamos tupla, así es más facil
        (name, value) = ("a", 0)
        index = 0

        for col in columns:
            index += 1
            tempCell = sheet.cell(row, col)
            print(tempCell.value)
            if tempCell.value is None:
                break
            elif type(tempCell.value) is str:
                (name, value) = file_constants.CELL_NG_NAME + str(index), -1
            elif tempCell.value > value:
                (name, value) = file_constants.CELL_NG_NAME + str(index), tempCell.value

        return (name, value)

    def __get_all_ng_columns(self, sheet, name) -> list:
        index = file_constants.EXTRACTION_FIRST
        hasData = True
        columns = []

        while hasData:
            cellValue = name + str(index)
            (row, col, hasData) = self.__get_row_column_by_value(cellValue, sheet)
            index += 1
            if hasData:
                columns.append(col)
            else:
                break

        return columns