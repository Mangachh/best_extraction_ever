
# TODO: Meter una comprobación de path aquí mismo. 2 partes: el formato es buen y que el path existe!

# metemos lo de pillar regex aquí, sólo porque no sé donde meterlo
import re

__INPUT_PROMPT = ">>"

__MESS_NO_INT = "Sólo están permitidos números enteros."
__MES_NO_FLOAT = "Sólo están permitidos números decimales."
__MES_NO_SINGLE = "Sólo está permitido introducir un carácter"
__MES_NO_LETTER = "Sólo se permite introducir una letra."





def ask_integer(header_mess: str, error_mess:str = __MESS_NO_INT) -> int:
    '''
    Pide como input un entero y no para hasta que lo consigue.
    :param header_mess: mensaje para el header
    :param error_mess: mensaje de error, hay un default
    :return: el valor entero
    '''
    value: int
    ans: str
    print(header_mess)
    correct_value = False
    while not correct_value:
        ans = input(__INPUT_PROMPT)
        try:
            value = int(ans)
            correct_value = True
        except ValueError:
            print(error_mess)
    return value


def ask_float(header_mess: str, error_mess: str = __MES_NO_FLOAT) -> float:
    '''
    Pide como input un número decimal y no para hasta que lo consigue
    :param header_mess: Header para el número
    :param error_mess: Error, tiene default
    :return:
    '''
    value: float
    ans: str
    print(header_mess)
    correct_value = False
    while not correct_value:
        ans = input(__INPUT_PROMPT)
        try:
            value = float(ans)
            correct_value = True
        except ValueError:
            print(error_mess)
    return value


def ask_single_char(header_mess: str, error_mess: str = __MES_NO_SINGLE) -> str:
    '''
    Pide como input un carácter y no para hasta que lo consigue
    :param header_mess: Header a mostrar
    :param error_mess: Mensaje de error, tiene default
    :return: caracter pillado
    '''
    correct = False
    while not correct:
        letter = input(__INPUT_PROMPT)
        if len(letter) > 1 or len(letter) == 0:
            print(error_mess)
        else:
            correct = True

    return  letter


def ask_only_letter(header_mess: str, error_mess: str = __MES_NO_LETTER) -> str:
    '''
    Pide y devuelve una sola letra.
    TODO: mirar que pille acentos y demás
    :param header_mess: Header del input
    :param error_mess: error si no es una letra
    :return: una letra
    '''
    correct = False
    print(header_mess)
    while not correct:
        letter = ask_single_char("", "")
        if letter.isdigit():
            return letter
        else:
            print(__MES_NO_LETTER)
