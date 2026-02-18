import os
import platform
import re
import webbrowser
from datetime import datetime

from openpyxl import Workbook

from rma_database import (
    SET,
    get_all,
    get_next_folio,
    import_from_excel_to_db,
    write_or_update_excel_row,
)

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")


def validate_date(s: str) -> bool:
    if not s or not DATE_RE.match(s):
        return False
    try:
        datetime.strptime(s, "%d-%m-%Y")
        return True
    except ValueError:
        return False


def validate_phone(s: str) -> bool:
    if not s:
        return True
    digits = re.sub(r"[^\d]", "", s)
    return 7 <= len(digits) <= 15


def validate_folio(s: str) -> bool:
    return bool(re.match(r"^RMA-RED-\d{4}$", s or ""))


def generate_folio() -> str:
    return get_next_folio()


def open_pdf(path: str) -> None:
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            os.system(f'open "{path}"')
        else:
            webbrowser.open(f"file://{os.path.abspath(path)}")
    except OSError:
        pass


def confirm(prompt: str) -> bool:
    ans = input(f"{prompt} (S/N): ").strip().lower()
    return ans in {"s", "si"}


def summarize_entry_for_check(entry: dict, exclude=None) -> str:
    exclude = set(exclude or [])
    order = [
        ("folio", "Folio"),
        ("fecha_recepcion", "Fecha de recepción"),
        ("nombre_cliente", "Cliente"),
        ("telefono", "Teléfono"),
        ("factura", "Factura"),
        ("producto", "Producto"),
        ("marca", "Marca"),
        ("modelo", "Modelo"),
        ("serie", "Serie"),
        ("falla", "Falla"),
        ("observaciones", "Observaciones"),
        ("responsable", "Responsable"),
    ]
    lines = []
    for key, label in order:
        if key in exclude:
            continue
        lines.append(f"- {label}: {entry.get(key, '')}")
    return "\n".join(lines)


def update_excel_from_db() -> None:
    for rec in get_all():
        write_or_update_excel_row(rec)


def update_db_from_excel() -> int:
    return import_from_excel_to_db(deduplicate=True, backup=True)


def ensure_excel_exists() -> None:
    path = SET["excel_path"]
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "RMA"
        ws.append(SET["excel_headers"])
        wb.save(path)
