import json
import os

from openpyxl import Workbook

DEFAULT_SETTINGS = {
    "company_name": "REDMAYOREO",
    "logo_path": "logo.png",
    "excel_path": "rma_estatus.xlsx",
    "database_path": "rma_database.json",
    "pdf_folder": "pdfs",
    "backup_folder": "backups",
    "excel_headers": [
        "Folio",
        "Fecha Recepción",
        "Cliente",
        "Teléfono",
        "Factura",
        "Producto",
        "Marca",
        "Modelo",
        "Serie",
        "Falla",
        "Observaciones",
        "Responsable",
        "Estado",
        "Fecha de Resolución",
        "Tipo de Resolución",
        "Producto Nuevo",
        "Modelo Nuevo",
        "Serie Nueva",
    ],
}


def ensure_dirs(settings: dict) -> None:
    os.makedirs(settings["pdf_folder"], exist_ok=True)
    os.makedirs(settings["backup_folder"], exist_ok=True)


def ensure_db(path: str) -> None:
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)


def ensure_excel(path: str, headers: list) -> None:
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "RMA"
        ws.append(headers)
        wb.save(path)


def ensure_admin_password(settings_path: str, settings: dict) -> None:
    if settings.get("admin_password"):
        return

    env_password = os.getenv("RMA_ADMIN_PASSWORD")
    if env_password:
        settings["admin_password"] = env_password
    else:
        settings["admin_password"] = "admin123"
        print(
            "⚠️ No se encontró RMA_ADMIN_PASSWORD; se usó contraseña temporal 'admin123'. "
            "Cámbiala en settings.json o define la variable de entorno."
        )

    with open(settings_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def main() -> None:
    settings_path = "settings.json"
    if not os.path.exists(settings_path):
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_SETTINGS, f, ensure_ascii=False, indent=2)
        print("settings.json creado.")

    with open(settings_path, "r", encoding="utf-8") as f:
        settings = json.load(f)

    ensure_admin_password(settings_path, settings)
    ensure_dirs(settings)
    ensure_db(settings["database_path"])
    ensure_excel(settings["excel_path"], settings["excel_headers"])

    print("Instalación lista:")
    print(f"- PDF folder: {settings['pdf_folder']}")
    print(f"- Backups: {settings['backup_folder']}")
    print(f"- Excel: {settings['excel_path']}")
    print(f"- DB: {settings['database_path']}")


if __name__ == "__main__":
    main()
