from openpyxl import load_workbook
from openpyxl.workbook import Workbook
import openpyxl

import Messages


def get_workbook(path:str) -> (Workbook, bool):
    isSpread: bool
    path = __path_formatted(path)

    try:
        spread = load_workbook(path)
        isSpread = True
        print(Messages.MESS_GOT_FILE)
    except Exception as inst:
        spread = None
        isSpread = False
        print(Messages.MESS_FAILED)
        print(inst)

    return (spread, isSpread)


def save_workbook(path:str, doc: Workbook):
    path = __path_formatted(path)
    try:
        print("Path " + path)
        doc.save(path)
    except Exception as inst:
        print(Messages.MESS_FAILED)
        print(inst)

    return

def __path_formatted(old_path:str) -> str:
    newPath =""
    if old_path[len(old_path) - 1] == '\"':
        for index in range(1, len(old_path) - 1):
            newPath += old_path[index]
    else:
        newPath = old_path

    return newPath







