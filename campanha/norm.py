"""Normalização de documento, razão social e telefone."""
import re

import pandas as pd
from unidecode import unidecode

SUFIXOS = r"\b(LTDA|LTD|ME|EPP|EIRELI|S\s*A|SA|SLU|MEI|E\s+CIA|CIA|EI)\b"


def so_digitos(x):
    if pd.isna(x):
        return ""
    s = str(x)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\D", "", s)


def doc(x, tipo=None):
    """CNPJ com 14 dígitos ou CPF com 11 (zeros à esquerda). Vazio se inválido."""
    d = so_digitos(x)
    if not d or set(d) == {"0"}:
        return ""
    if tipo == "F" or (tipo is None and len(d) <= 11):
        return d.zfill(11) if len(d) <= 11 else d
    return d.zfill(14) if len(d) <= 14 else d


def raiz(d):
    return d[:8] if len(d) == 14 else d


def razao(x):
    if pd.isna(x):
        return ""
    s = unidecode(str(x)).upper()
    s = s.replace("S/A", " SA ").replace("S.A.", " SA ").replace("&", " E ")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(SUFIXOS, " ", s)
    s = re.sub(r"\b(E)\s*$", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def telefones(x):
    """Lista de números brasileiros (DDD + número) achados num texto."""
    if pd.isna(x):
        return []
    out = []
    for parte in re.split(r"[/;,|]| e |\n", str(x)):
        d = so_digitos(parte)
        if d.startswith("55") and len(d) in (12, 13):
            d = d[2:]
        if d.startswith("0") and len(d) in (11, 12):
            d = d[1:]
        if len(d) in (10, 11) and d[:2] not in ("00", "10") and int(d[:2]) >= 11:
            out.append(d)
    return out


def celular(d):
    return len(d) == 11 and d[2] == "9"


def e164(d):
    return "+55" + d if d else ""
