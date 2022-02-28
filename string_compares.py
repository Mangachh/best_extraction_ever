import unicodedata
import re

__MESS_NO_EXPRESSION = "No se ha encontrado el subtexto."


def equal_strings(text_a: str, text_b: str) -> bool:
    return  text_a == text_b

def equal_strings_not_case(text_a: str, text_b: str) -> bool:
    return normalize_text(text_a).casefold() == normalize_text(text_b).casefold()

def normalize_text(text: str) -> str:
    return unicodedata.normalize('NFD', text)

def regular_substring(expresion: str, text: str):
    try:
        substring = re.search(expresion, text).group()
        return substring
    except AttributeError:
        print(__MESS_NO_EXPRESSION)
        return ""

