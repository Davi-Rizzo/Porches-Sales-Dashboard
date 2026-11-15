#!/usr/bin/env python3
"""
Agente de sanitização – Porsche Sales
-------------------------------------
Lê a planilha bruta e aplica as regras do schema.md, gerando um novo .xlsx
com as colunas originais + colunas *Sanitized logo após cada coluna de origem.

Uso:
    python3 agente_sanitizacao_porsche.py ENTRADA.xlsx SAIDA.xlsx
"""
import re
import sys
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

INVALID = "INVALID"

# --------------------------------------------------------------------------
# Utilitários numéricos
# --------------------------------------------------------------------------
UNITS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
    "seventeen": 17, "eighteen": 18, "nineteen": 19,
}
TENS = {"twenty": 20, "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60,
        "seventy": 70, "eighty": 80, "ninety": 90}


def words_to_int(text):
    """'eighty two thousand' -> 82000. Retorna None se não entender."""
    tokens = re.findall(r"[a-z]+", text.lower().replace("-", " "))
    tokens = [t for t in tokens if t != "and"]
    if not tokens:
        return None
    total, current = 0, 0
    for t in tokens:
        if t in UNITS:
            current += UNITS[t]
        elif t in TENS:
            current += TENS[t]
        elif t == "hundred":
            current = max(current, 1) * 100
        elif t == "thousand":
            total += max(current, 1) * 1000
            current = 0
        elif t == "million":
            total += max(current, 1) * 1_000_000
            current = 0
        else:
            return None
    return total + current


def parse_number(s):
    """
    Converte '1.200', '12,500', '$103.750,00', '153,200.50' em Decimal.
    Regras de separador:
      - '.' e ',' presentes: o ÚLTIMO é o decimal, o outro é milhar.
      - só um tipo, repetido: milhar.
      - só um tipo, uma vez: 3 dígitos depois = milhar; senão decimal.
    """
    m = re.search(r"\d[\d.,]*", s)
    if not m:
        return None
    num = m.group(0).rstrip(".,")
    has_dot, has_com = "." in num, "," in num
    if has_dot and has_com:
        dec = "." if num.rfind(".") > num.rfind(",") else ","
        thou = "," if dec == "." else "."
        num = num.replace(thou, "").replace(dec, ".")
    elif has_dot or has_com:
        sep = "." if has_dot else ","
        parts = num.split(sep)
        if len(parts) > 2 or len(parts[-1]) == 3:
            num = "".join(parts)          # separador de milhar
        else:
            num = ".".join(parts)         # separador decimal
    try:
        return Decimal(num)
    except Exception:
        return None


def half_up_int(d):
    return int(Decimal(d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


# --------------------------------------------------------------------------
# Datas
# --------------------------------------------------------------------------
MONTHS = {
    "january": 1, "jan": 1, "february": 2, "feb": 2, "march": 3, "mar": 3,
    "april": 4, "apr": 4, "may": 5, "june": 6, "jun": 6, "july": 7, "jul": 7,
    "august": 8, "aug": 8, "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10, "november": 11, "nov": 11, "december": 12,
    "dec": 12,
}


def _mk(y, m, d):
    try:
        return date(int(y), int(m), int(d)).isoformat()
    except ValueError:
        return INVALID


def sanitize_date(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if v is None:
        return INVALID
    s = str(v).strip()
    if m := re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s):
        return _mk(m[1], m[2], m[3])                      # YYYY-MM-DD
    if m := re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4}|\d{2})", s):
        y = int(m[3])
        y = y + 2000 if y < 100 else y
        return _mk(y, m[1], m[2])                         # MM/DD/YY(YY)
    if m := re.fullmatch(r"([A-Za-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", s):
        mon = MONTHS.get(m[1].lower())
        return _mk(m[3], mon, m[2]) if mon else INVALID   # Month DDth, YYYY
    return INVALID


# --------------------------------------------------------------------------
# Modelo
# --------------------------------------------------------------------------
CANONICAL_MODELS = [
    "Carrinho", "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Turbo",
    "911 Turbo S", "911 GT3", "911 GT3 RS", "911 Dakar", "911 Targa 4",
    "911 Targa 4S", "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster GTS", "718 Spyder RS", "Cayenne", "Cayenne S",
    "Cayenne Coupe", "Cayenne E-Hybrid", "Cayenne Turbo", "Cayenne Turbo GT",
    "Macan", "Macan S", "Macan T", "Macan GTS", "Macan Electric", "Panamera",
    "Panamera 4", "Panamera 4S", "Panamera Turbo", "Panamera Turbo S",
    "Panamera 4 E-Hybrid", "Taycan", "Taycan 4S", "Taycan GTS", "Taycan Turbo",
    "Taycan Turbo S", "Taycan Cross Turismo",
]
_MODEL_LOOKUP = {re.sub(r"[\s\-_]+", " ", m.lower()): m for m in CANONICAL_MODELS}
_UPPER_TOKENS = {"GT3", "GT4", "RS", "GTS", "GT"}


def _title_token(tok):
    if tok.upper() in _UPPER_TOKENS:
        return tok.upper()
    if re.search(r"\d", tok):                  # 4S, 718, 911...
        return tok.upper()
    if "-" in tok:                             # E-Hybrid
        return "-".join(p.capitalize() for p in tok.split("-"))
    return tok.capitalize()


def sanitize_model(v):
    if v is None or str(v).strip() == "":
        return INVALID
    s = re.sub(r"\s+", " ", str(v).strip())
    key = re.sub(r"[\s\-_]+", " ", s.lower())
    if key in _MODEL_LOOKUP:
        return _MODEL_LOOKUP[key]
    return " ".join(_title_token(t) for t in s.split(" "))   # desconhecido: title case


# --------------------------------------------------------------------------
# Ano
# --------------------------------------------------------------------------
def sanitize_year(v):
    if v is None:
        return INVALID
    year = None
    if isinstance(v, (int, float)):
        year = int(v)
    else:
        s = str(v).strip().lower()
        if m := re.fullmatch(r"(\d{4})", s):
            year = int(m[1])
        elif m := re.fullmatch(r"(\d{2})[\s\-](\d{2})", s):       # 20-24 / 20 24
            year = int(m[1] + m[2])
        else:
            toks = re.findall(r"[a-z]+", s.replace("-", " "))
            # 'twenty twenty four' -> 20|24 ; 'nineteen ninety nine' -> 19|99
            if len(toks) >= 2 and toks[0] in ("nineteen", "twenty"):
                head = UNITS.get(toks[0], TENS.get(toks[0]))
                tail = words_to_int(" ".join(toks[1:]))
                if tail is not None and 0 <= tail < 100 and toks[1] != "thousand":
                    year = head * 100 + tail
            if year is None:
                year = words_to_int(s)                            # 'two thousand twenty one'
    if year is None or not (1990 <= year <= 2035):
        return INVALID
    return str(year)


# --------------------------------------------------------------------------
# Preço
# --------------------------------------------------------------------------
def sanitize_price(v):
    if v is None:
        return INVALID
    if isinstance(v, (int, float)):
        amount = Decimal(str(v))
    else:
        s = str(v).strip().lower()
        s_clean = re.sub(r"usd|dollars?|\$", " ", s).strip()
        amount = None
        if re.search(r"\d", s_clean):
            n = parse_number(s_clean)
            if n is not None:
                if re.search(r"\d\s*k\b", s_clean):
                    n *= 1000
                amount = n
        else:
            w = words_to_int(s_clean)
            amount = Decimal(w) if w is not None else None
        if amount is None:
            return INVALID
    return str(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# --------------------------------------------------------------------------
# Quilometragem (sempre em milhas inteiras)
# --------------------------------------------------------------------------
def sanitize_mileage(v):
    if v is None:
        return INVALID
    if isinstance(v, (int, float)):
        return half_up_int(Decimal(str(v)))
    s = str(v).strip().lower()
    if s in ("new", "new car", "zero", "zero miles", "0 mi", "0 miles"):
        return 0
    is_km = bool(re.search(r"\bkm\b|kilomet", s))
    if re.search(r"\d", s):
        n = parse_number(s)
    else:
        w = words_to_int(re.sub(r"miles?|mi\b|\.|:", " ", s))
        n = Decimal(w) if w is not None else None
    if n is None:
        return INVALID
    if is_km:
        n = n * Decimal("0.621371")
    return half_up_int(n)


# --------------------------------------------------------------------------
# Pagamento
# --------------------------------------------------------------------------
def sanitize_payment(v):
    if v is None or str(v).strip() == "":
        return INVALID
    s = re.sub(r"[\s\-_]+", " ", str(v).strip().lower())
    if "ach" in s.split():
        return "ACH Payment"
    if "debit" in s:
        return "Debit Card"
    if "credit" in s:
        return "Credit Card"
    if "wire" in s:
        return "Wire Transfer"
    if "bank" in s and "transfer" in s:
        return "Bank Transfer"
    if "financ" in s or s == "finance":
        return "Financing"
    if "leas" in s:
        return "Lease"
    if "cash" in s:
        return "Cash"
    if "crypto" in s:
        return "Crypto Payment"
    return s.title()


# --------------------------------------------------------------------------
# Cidade
# --------------------------------------------------------------------------
def sanitize_city(v):
    if v is None or str(v).strip() == "":
        return INVALID
    s = re.sub(r"\s+", " ", str(v).strip())
    words = []
    for w in s.split(" "):
        lw = w.lower()
        if lw in ("st", "st."):
            words.append("St.")
        elif lw in ("ft", "ft."):
            words.append("Ft.")
        elif lw in ("mt", "mt."):
            words.append("Mt.")
        else:
            words.append(w.capitalize())
    return " ".join(words)


# --------------------------------------------------------------------------
# Estado (USPS)
# --------------------------------------------------------------------------
STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "district of columbia": "DC", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN",
    "iowa": "IA", "kansas": "KS", "kentucky": "KY", "louisiana": "LA",
    "maine": "ME", "maryland": "MD", "massachusetts": "MA", "michigan": "MI",
    "minnesota": "MN", "mississippi": "MS", "missouri": "MO", "montana": "MT",
    "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC",
    "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
}
_ABBRS = set(STATES.values())


def sanitize_state(v):
    if v is None:
        return INVALID
    s = re.sub(r"\s+", " ", str(v).strip().lower())
    if s.upper() in _ABBRS:
        return s.upper()
    return STATES.get(s, INVALID)


# --------------------------------------------------------------------------
# Status de entrega
# --------------------------------------------------------------------------
DELIVERY = {
    "delivered": "Delivered", "deliverd": "Delivered", "pending": "Pending",
    "in transit": "In Transit", "cancelled": "Cancelled", "canceled": "Cancelled",
    "awaiting delivery": "Awaiting Delivery", "awaiting pickup": "Awaiting Pickup",
    "pending approval": "Pending Approval", "pending review": "Pending Review",
    "shipped": "Shipped", "awaiting review": "Awaiting Review",
}


def sanitize_delivery(v):
    if v is None:
        return INVALID
    s = re.sub(r"[^a-z\s]", " ", str(v).lower().replace("-", " "))
    s = re.sub(r"\s+", " ", s).strip()
    return DELIVERY.get(s, INVALID)


# --------------------------------------------------------------------------
# Pipeline
# --------------------------------------------------------------------------
# (coluna de origem, coluna de saída, função)
RULES = [
    ("sale_date", "SaleDateSanitized", sanitize_date),
    ("porsche_model", "PorscheModelSanitized", sanitize_model),
    ("model_year", "ModelYearSanitized", sanitize_year),
    ("sale_price", "SalesPriceSanitized", sanitize_price),
    ("vehicle_mileage", "VehicleMileageSanitized", sanitize_mileage),
    ("payment_method", "PayMethodSanitized", sanitize_payment),
    ("city", "CitySanitized", sanitize_city),
    ("state", "StateSanitized", sanitize_state),
    ("delivery_status", "DeliveryStatusSanitized", sanitize_delivery),
]
REQUIRED = ["sale_id", "sale_date", "customer_name", "porsche_model", "model_year",
            "sale_price", "vehicle_mileage", "payment_method", "city", "state",
            "salesperson", "delivery_status"]


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header, data = list(rows[0]), rows[1:]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        sys.exit(f"Colunas obrigatórias ausentes: {missing}")
    idx = {c: i for i, c in enumerate(header)}
    rule_by_src = {s: (o, f) for s, o, f in RULES}

    # novo cabeçalho: sanitizada logo após a original
    new_header, plan = [], []
    for c in header:
        new_header.append(c)
        plan.append(("orig", idx[c]))
        if c in rule_by_src:
            out, fn = rule_by_src[c]
            new_header.append(out)
            plan.append(("san", idx[c], fn))

    out_rows = []
    for r in data:
        row = []
        for p in plan:
            row.append(r[p[1]] if p[0] == "orig" else p[2](r[p[1]]))
        out_rows.append(row)

    # ---- quality checks ----
    col = {h: i for i, h in enumerate(new_header)}
    errors = []
    for i, h in enumerate(header):                      # nenhuma coluna removida
        if h not in col:
            errors.append(f"coluna original removida: {h}")
    for s, o, _ in RULES:                               # sanitizada logo após origem
        if col[o] != col[s] + 1:
            errors.append(f"{o} não está logo após {s}")
    for r in out_rows:
        d = r[col["SaleDateSanitized"]]
        if d != INVALID and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
            errors.append(f"data fora do padrão: {d}")
        st = r[col["StateSanitized"]]
        if st != INVALID and not re.fullmatch(r"[A-Z]{2}", st):
            errors.append(f"estado fora do padrão: {st}")
        if not isinstance(r[col["VehicleMileageSanitized"]], int) and \
                r[col["VehicleMileageSanitized"]] != INVALID:
            errors.append("milhagem não inteira")
        p = r[col["SalesPriceSanitized"]]
        if p != INVALID and not re.fullmatch(r"\d+\.\d{2}", p):
            errors.append(f"preço fora do padrão: {p}")
        for _, o, _ in RULES:
            if r[col[o]] in (None, ""):
                errors.append(f"valor em branco em {o}")

    # ---- escreve saída ----
    out = openpyxl.Workbook()
    wso = out.active
    wso.title = "tratado"
    wso.append(new_header)
    for r in out_rows:
        wso.append(r)

    san_cols = {o for _, o, _ in RULES}
    head_fill = PatternFill("solid", fgColor="1F2937")
    san_fill = PatternFill("solid", fgColor="0F766E")
    bad_fill = PatternFill("solid", fgColor="FECACA")
    for j, h in enumerate(new_header, start=1):
        c = wso.cell(row=1, column=j)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = san_fill if h in san_cols else head_fill
        c.alignment = Alignment(horizontal="center")
        width = max(len(str(h)), *(len(str(wso.cell(row=r, column=j).value or ""))
                                    for r in range(2, wso.max_row + 1))) + 2
        wso.column_dimensions[get_column_letter(j)].width = min(width, 32)
        if h in san_cols:
            for r in range(2, wso.max_row + 1):
                cell = wso.cell(row=r, column=j)
                if cell.value == INVALID:
                    cell.fill = bad_fill
        elif h == "sale_date":
            for r in range(2, wso.max_row + 1):
                if isinstance(wso.cell(row=r, column=j).value, datetime):
                    wso.cell(row=r, column=j).number_format = "yyyy-mm-dd"
    wso.freeze_panes = "B2"

    # aba de relatório
    rep = out.create_sheet("relatorio")
    rep.append(["Coluna sanitizada", "Total", "INVALID", "% INVALID"])
    for _, o, _ in RULES:
        n = sum(1 for r in out_rows if r[col[o]] == INVALID)
        rep.append([o, len(out_rows), n, round(n / len(out_rows), 4)])
    for c in rep[1]:
        c.font = Font(bold=True)
    for r in range(2, rep.max_row + 1):
        rep.cell(row=r, column=4).number_format = "0.0%"
    rep.column_dimensions["A"].width = 28
    out.save(dst)

    print(f"Linhas processadas: {len(out_rows)}")
    print(f"Colunas: {len(header)} originais + {len(RULES)} sanitizadas")
    for _, o, _ in RULES:
        n = sum(1 for r in out_rows if r[col[o]] == INVALID)
        print(f"  {o:26s} INVALID: {n}")
    print("Quality checks:", "OK" if not errors else f"{len(errors)} problemas")
    for e in errors[:20]:
        print("  -", e)
    return 1 if errors else 0


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.exit(run(sys.argv[1], sys.argv[2]))
