WELCOME = "Bienvenido al nuevo programa del Lluís."
ADVISE = "OJU! Para que esto funcione tienes que hacer antes varias cosas."
INSTRUCTIONS = ("Primero tienes que crear una hoja llamada \"LISTA\".\n"
                "Segundo, en la celda que más rabia te de, escribes \"BUSCAR\".\n"
                "Debajo de esa celda coloca todos los HC a buscar. Y, en teoria, eso es todo.\n"
                "Si tienes más dudas, abre el archivo \"LÉEME\"\n")
END = "Pues ya está. Si Tenias la hoja abierta, ciérrala SIN guardar los cambios y vuelve a abrirla."

PRESS_KEY_EXIT = "Presiona cualquier tecla para salir"


PATH_ERROR_FILE = "El archivo no existe."
PATH_ENTER = "Introduce la ruta completa del archivo."
PATH_CORRECT = "El archivo es correcto."

MESS_WRITTING = "\n\n*** Escribiendo los valores en la hoja de Excel ***\n"
MESS_SEARCH = "\n\n*** Buscando los valores por cada HC ***\n"
MESS_GOT_FILE = "Archivo encontrado y abierto."
MESS_FAILED = "Algo ha fallado!!"
MESS_NOT_FOUND_LIST = "No se ha encontrado la hoja con nombre \"LISTA\""
MESS_NOT_FOUND_COLUMN ="No se ha encontrado la columna \"BUSCAR\" en la hoja \"LISTA\" o bien no hay ningún valor en la hoja"

BORDER = "+-----------------------------------------------------+"

def print_value_inside_border(value):
    value = "|" + str(value)
    for i in range(len(BORDER)):
        if len(value) < i:
            value += " "

    value += "|"
    print(value)
    return


def print_full(esq_code, bm_code, tipe, extrac, ext_value, isRep = False):
    full_st : str
    full_st = "| " + str(esq_code) + "\t" + str(bm_code) + "\t" + str(tipe) + "\t" + str(extrac) + "\t" + str(ext_value)
    if isRep:
        full_st += "\t OJU REP!!!"

    __print_border(full_st)
    print(full_st)
    __print_border(full_st)
    return


def __print_border(text: str):
    border =""
    for char in text:
        if char == '\t':
            border += "----"
        else:
            border += "-"

    print(border)
    return