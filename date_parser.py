"""Interpretación conservadora de fechas históricas (p. ej. OCT.6/25F)."""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass

MONTHS = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
          "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}

@dataclass
class ParsedDate:
    original: str
    status: str  # VALIDA, RECUPERABLE, INVALIDA, AMBIGUA
    normalized: str | None = None
    interpreted: str | None = None
    reason: str | None = None
    month: int | None = None
    year: int | None = None

def _month(token: str):
    token = token.upper()
    if token in MONTHS:
        return token, False
    # Solo correcciones inequívocas: letra repetida o distancia de edición 1.
    compressed = re.sub(r"(.)\1+", r"\1", token)
    candidates = [m for m in MONTHS if _distance(compressed, m) <= 1]
    if len(candidates) == 1:
        return candidates[0], True
    return None, False

def _distance(a: str, b: str) -> int:
    if abs(len(a)-len(b)) > 1: return 2
    previous = list(range(len(b)+1))
    for i, x in enumerate(a, 1):
        current = [i]
        for j, y in enumerate(b, 1):
            current.append(min(current[-1]+1, previous[j]+1, previous[j-1]+(x != y)))
        previous = current
    return previous[-1]

def parse_date(value) -> ParsedDate:
    original = "" if value is None else str(value).strip()
    if not original:
        return ParsedDate(original, "AMBIGUA", reason="Fecha vacía")
    text = re.sub(r"\s+", "", original.upper())
    source_f = text.endswith("F")
    core = text[:-1] if source_f else text
    # Mes, día y año; se tolera la falta de punto/slash, no números con longitudes dudosas.
    match = re.fullmatch(r"([A-Z]+)[.\-_/]?(\d{1,2})[.\-_/]?(\d{2}|\d{4})", core)
    if not match:
        return ParsedDate(original, "AMBIGUA", reason="No se reconoce inequívocamente mes, día y año")
    mon_token, day_text, year_text = match.groups()
    mon, changed_month = _month(mon_token)
    if not mon:
        return ParsedDate(original, "AMBIGUA", reason="Mes no reconocido o ambiguo")
    day, year = int(day_text), int(year_text)
    if year < 100: year += 2000
    if day < 1 or day > calendar.monthrange(year, MONTHS[mon])[1]:
        return ParsedDate(original, "INVALIDA", reason="Día no válido para el mes", month=MONTHS[mon], year=year)
    normalized = f"{mon}.{day}/{str(year)[-2:]}" + ("F" if source_f else "")
    altered = changed_month or original.upper() != normalized
    return ParsedDate(original, "RECUPERABLE" if altered else "VALIDA", normalized,
                      f"{day:02d}/{MONTHS[mon]:02d}/{year}", month=MONTHS[mon], year=year)
