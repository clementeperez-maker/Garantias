import os
from datetime import datetime

from rma_database import (
    SET,
    get_all,
    get_by_folio,
    upsert_entry,
    save_rma_data,
    write_or_update_excel_row,
)
from rma_pdf import (
    generate_no_valida_pdf,
    generate_reposicion_pdf,
    generate_rma_pdf,
    generate_validacion_pdf,
)
from rma_utils import (
    confirm,
    ensure_excel_exists,
    generate_folio,
    open_pdf,
    summarize_entry_for_check,
    update_db_from_excel,
    update_excel_from_db,
    validate_date,
    validate_folio,
    validate_phone,
)

PDF_DIR = SET["pdf_folder"]


def input_nonempty(prompt: str) -> str:
    while True:
        v = input(prompt).strip()
        if v:
            return v


def input_date(prompt: str) -> str:
    while True:
        v = input(prompt).strip()
        if validate_date(v):
            return v
        print("Formato de fecha inválido. Use DD-MM-YYYY.")


def input_phone(prompt: str) -> str:
    while True:
        v = input(prompt).strip()
        if validate_phone(v):
            return v
        print("Teléfono inválido. Deje vacío o ingrese 7-15 dígitos.")


def input_folio(prompt: str) -> str:
    while True:
        v = input(prompt).strip().upper()
        if validate_folio(v):
            return v
        print("Folio inválido. Formato: RMA-RED-0001")


def flow_create_rma() -> None:
    os.makedirs(PDF_DIR, exist_ok=True)
    ensure_excel_exists()

    folio = generate_folio()
    print(f"\nFolio asignado: {folio}")

    entry = {
        "folio": folio,
        "fecha_recepcion": input_date("Fecha de recepción (DD-MM-YYYY): "),
        "nombre_cliente": input_nonempty("Cliente: "),
        "telefono": input_phone("Teléfono (opcional): "),
        "factura": input_nonempty("Factura: "),
        "producto": input_nonempty("Producto: "),
        "marca": input_nonempty("Marca: "),
        "modelo": input_nonempty("Modelo: "),
        "serie": input_nonempty("Serie: "),
        "falla": input_nonempty("Falla: "),
        "observaciones": input("Observaciones (dejar vacío si no hay): ").strip(),
        "responsable": input_nonempty("Responsable: "),
        "estado": "En Proceso",
        "fecha_entrega": "",
        "tipo_resolucion": "",
        "producto_nuevo": "",
        "modelo_nuevo": "",
        "serie_nueva": "",
    }

    print("\nVerifique la información:")
    print(
        summarize_entry_for_check(
            entry, exclude={"estado", "fecha_entrega", "tipo_resolucion"}
        )
    )
    if not confirm("¿Los datos son correctos?"):
        print("Operación cancelada. No se guardó el RMA.")
        return

    save_rma_data(entry)
    write_or_update_excel_row(entry)

    filename = os.path.join(PDF_DIR, f"{folio}.pdf")
    generate_rma_pdf(entry, filename, SET)
    print(f"PDF generado: {filename}")
    open_pdf(filename)


def flow_validate_guarantee() -> None:
    folio = input_folio("Ingrese folio de RMA: ")
    rec = get_by_folio(folio)
    if not rec:
        print("Folio no existe en la base de datos.")
        return

    estado_resuelto = (rec.get("estado", "").lower() in ["aprobada", "válida", "valida", "no válida", "no valida"]) or rec.get("tipo_resolucion")
    if estado_resuelto:
        print("La garantía ya fue resuelta.")
        return

    print("\nTipo de validación:")
    print("1) Equipo Nuevo  2) Reparación  3) Nota de Crédito  4) No válida")
    op = input("Opción: ").strip()

    update = {
        "folio": folio,
        "fecha_recepcion": rec.get("fecha_recepcion", ""),
        "nombre_cliente": rec.get("nombre_cliente", ""),
        "telefono": rec.get("telefono", ""),
        "factura": rec.get("factura", ""),
        "producto": rec.get("producto", ""),
        "marca": rec.get("marca", ""),
        "modelo": rec.get("modelo", ""),
        "serie": rec.get("serie", ""),
        "falla": rec.get("falla", ""),
        "observaciones": rec.get("observaciones", ""),
        "responsable": rec.get("responsable", ""),
        "estado": "",
        "fecha_entrega": "",
        "tipo_resolucion": "",
        "producto_nuevo": "",
        "modelo_nuevo": "",
        "serie_nueva": "",
        "detalles_reparacion": "",
        "detalles_no_valida": "",
    }

    if op == "1":
        update["estado"] = "Aprobada"
        update["tipo_resolucion"] = "Equipo Nuevo"
        update["fecha_entrega"] = input_date("Fecha de entrega (DD-MM-YYYY): ")
        update["producto_nuevo"] = input_nonempty("Producto Nuevo: ")
        update["modelo_nuevo"] = input_nonempty("Modelo Nuevo: ")
        update["serie_nueva"] = input_nonempty("Serie Nueva: ")

        print("\nVerifique la nueva información capturada:")
        print(f"- Fecha de entrega: {update['fecha_entrega']}")
        print(f"- Producto Nuevo: {update['producto_nuevo']}")
        print(f"- Modelo Nuevo: {update['modelo_nuevo']}")
        print(f"- Serie Nueva: {update['serie_nueva']}")
        if not confirm("¿Es correcta?"):
            print("Cancelado.")
            return

        merged = {**rec, **update}
        upsert_entry(merged)
        write_or_update_excel_row(merged)

        filename = os.path.join(PDF_DIR, f"{folio}_valida.pdf")
        generate_validacion_pdf(merged, filename, SET)
        print(f"PDF generado: {filename}")
        open_pdf(filename)

    elif op == "2":
        update["estado"] = "Aprobada"
        update["tipo_resolucion"] = "Reparación"
        update["fecha_entrega"] = input_date("Fecha de entrega (DD-MM-YYYY): ")
        update["detalles_reparacion"] = input_nonempty("Detalles de la reparación: ")

        print("\nVerifique la nueva información capturada:")
        print(f"- Fecha de entrega: {update['fecha_entrega']}")
        print(f"- Detalles: {update['detalles_reparacion']}")
        if not confirm("¿Es correcta?"):
            print("Cancelado.")
            return

        merged = {**rec, **update}
        upsert_entry(merged)
        write_or_update_excel_row(merged)

        filename = os.path.join(PDF_DIR, f"{folio}_valida.pdf")
        generate_validacion_pdf(merged, filename, SET)
        print(f"PDF generado: {filename}")
        open_pdf(filename)

    elif op == "3":
        update["estado"] = "Aprobada"
        update["tipo_resolucion"] = "Nota de Crédito"
        update["fecha_entrega"] = input_date("Fecha de entrega (DD-MM-YYYY): ")

        print("\nVerifique la nueva información capturada:")
        print(f"- Fecha de entrega: {update['fecha_entrega']}")
        print("- Tipo de resolución: Nota de Crédito")
        if not confirm("¿Es correcta?"):
            print("Cancelado.")
            return

        merged = {**rec, **update}
        upsert_entry(merged)
        write_or_update_excel_row(merged)

        filename = os.path.join(PDF_DIR, f"{folio}_valida.pdf")
        generate_validacion_pdf(merged, filename, SET)
        print(f"PDF generado: {filename}")
        open_pdf(filename)

    elif op == "4":
        update["estado"] = "No válida"
        update["fecha_entrega"] = input_date("Fecha de entrega (DD-MM-YYYY): ")
        update["detalles_no_valida"] = input_nonempty("Detalles (motivo no válida): ")

        print("\nVerifique la nueva información capturada:")
        print(f"- Fecha de entrega: {update['fecha_entrega']}")
        print(f"- Detalles: {update['detalles_no_valida']}")
        if not confirm("¿Es correcta?"):
            print("Cancelado.")
            return

        merged = {**rec, **update}
        upsert_entry(merged)
        write_or_update_excel_row(merged)

        filename = os.path.join(PDF_DIR, f"{folio}_no_valida.pdf")
        generate_no_valida_pdf(merged, filename, SET)
        print(f"PDF generado: {filename}")
        open_pdf(filename)

    else:
        print("Opción inválida.")


def flow_reposicion() -> None:
    folio = input_folio("Folio de RMA: ")
    rec = get_by_folio(folio)
    if not rec:
        print("Folio no existe en la base de datos.")
        return

    if rec.get("reposicion_realizada") is True:
        print("Este RMA ya tiene reposición de mercancía registrada.")
        return

    fecha_rep = input_date("Fecha de reposición (DD-MM-YYYY): ")
    detalles_rep = input_nonempty("Información (detalles de la reposición): ")
    reemplazo = input_nonempty("Reemplazo: ")
    responsable = rec.get("responsable") or input_nonempty("Responsable: ")
    enc_alm = input_nonempty("Encargado de almacén (para firma): ")

    print("\nVerifique la nueva información capturada:")
    print(f"- Fecha de reposición: {fecha_rep}")
    print(f"- Información: {detalles_rep}")
    print(f"- Reemplazo: {reemplazo}")
    print(f"- Encargado de almacén: {enc_alm}")
    if not confirm("¿Es correcta?"):
        print("Cancelado.")
        return

    merged = {
        **rec,
        "fecha_reposicion": fecha_rep,
        "detalles_reposicion": detalles_rep,
        "reemplazo": reemplazo,
        "encargado_almacen": enc_alm,
        "responsable": responsable,
        "reposicion_realizada": True,
    }
    upsert_entry(merged)

    filename = os.path.join(PDF_DIR, f"{folio}_mercancia.pdf")
    generate_reposicion_pdf(merged, filename, SET)
    print(f"PDF generado: {filename}")
    open_pdf(filename)


def flow_info_rma() -> None:
    folio = input_folio("Folio: ")
    rec = get_by_folio(folio)
    if not rec:
        print("No existe.")
        return

    print("\nInformación del RMA:")
    basic = [
        ("Fecha de recepción", rec.get("fecha_recepcion", "")),
        ("Cliente", rec.get("nombre_cliente", "")),
        ("Teléfono", rec.get("telefono", "")),
        ("Factura", rec.get("factura", "")),
        ("Producto", rec.get("producto", "")),
        ("Marca", rec.get("marca", "")),
        ("Modelo", rec.get("modelo", "")),
        ("Serie", rec.get("serie", "")),
        ("Falla", rec.get("falla", "")),
        ("Observaciones", rec.get("observaciones", "")),
        ("Responsable", rec.get("responsable", "")),
        ("Estado", rec.get("estado", "En Proceso")),
    ]
    for k, v in basic:
        print(f"- {k}: {v}")

    est = (rec.get("estado", "En Proceso") or "").lower()
    if est in ["aprobada", "válida", "valida"]:
        print(f"- Tipo de Resolución: {rec.get('tipo_resolucion', '')}")
        print(f"- Fecha de Resolución: {rec.get('fecha_entrega', '')}")


def flow_estatus_garantia() -> None:
    folio = input_folio("Folio: ")
    rec = get_by_folio(folio)
    if not rec:
        print("No existe.")
        return

    est = (rec.get("estado", "En Proceso") or "En Proceso").strip()
    if est.lower() == "en proceso":
        print("Estatus: En Proceso")
    elif est.lower() in ["aprobada", "válida", "valida"]:
        print(f"Estatus: Aprobada -> {rec.get('tipo_resolucion', '') or '(sin tipo)'}")
    else:
        print("Estatus: No válida")


def flow_search() -> None:
    print("\nBúsqueda avanzada:")
    print("1) Por Cliente   2) Por Serie   3) Por Factura   4) Por Estado")
    op = input("Opción: ").strip()
    q = input("Valor a buscar: ").strip().lower()

    results = []
    for r in get_all():
        if op == "1" and q in r.get("nombre_cliente", "").lower():
            results.append(r)
        elif op == "2" and q in r.get("serie", "").lower():
            results.append(r)
        elif op == "3" and q in r.get("factura", "").lower():
            results.append(r)
        elif op == "4" and q == r.get("estado", "").lower():
            results.append(r)

    if not results:
        print("Sin resultados.")
        return

    for r in results:
        print(
            f"- {r.get('folio', '')} | {r.get('nombre_cliente', '')} | {r.get('producto', '')} | {r.get('estado', '')}"
        )

    folio_sel = input("\nIngrese el folio exacto a consultar: ").strip()
    elegido = get_by_folio(folio_sel)
    if not elegido:
        print("⚠️ Folio no válido.")
        return

    print("\n=== Detalle del RMA ===")
    for k, v in elegido.items():
        print(f"{k}: {v}")

    while True:
        print("\nOpciones:")
        print("1) Exportar a PDF")
        print("2) Regresar al menú principal")
        choice = input("Seleccione una opción: ").strip()

        if choice == "1":
            print("\nSeleccione tipo de PDF a exportar:")
            print("1) Solicitud inicial")
            print("2) RMA Final con firmas")
            tipo_pdf = input("Opción: ").strip()

            try:
                if tipo_pdf == "1":
                    filename = os.path.join(PDF_DIR, f"{elegido['folio']}_solicitud.pdf")
                    generate_rma_pdf(elegido, filename, SET)
                    print(f"✔️ PDF de solicitud generado: {filename}")
                    open_pdf(filename)
                elif tipo_pdf == "2":
                    estado = elegido.get("estado", "").lower()
                    if estado in ["aprobada", "válida", "valida"]:
                        filename = os.path.join(PDF_DIR, f"{elegido['folio']}_final.pdf")
                        generate_validacion_pdf(elegido, filename, SET)
                        print(f"✔️ PDF final con firmas generado: {filename}")
                        open_pdf(filename)
                    elif estado in ["no válida", "no valida"]:
                        filename = os.path.join(PDF_DIR, f"{elegido['folio']}_final.pdf")
                        generate_no_valida_pdf(elegido, filename, SET)
                        print(f"✔️ PDF de garantía no válida generado: {filename}")
                        open_pdf(filename)
                    elif elegido.get("reposicion_realizada"):
                        filename = os.path.join(PDF_DIR, f"{elegido['folio']}_final.pdf")
                        generate_reposicion_pdf(elegido, filename, SET)
                        print(f"✔️ PDF de reposición generado: {filename}")
                        open_pdf(filename)
                    else:
                        print(
                            "⚠️ Este RMA aún no tiene información suficiente para generar un final."
                        )
                else:
                    print("⚠️ Opción inválida, no se generó ningún PDF.")
            except Exception as e:
                print(f"❌ Error al generar PDF: {e}")
        elif choice == "2":
            return
        else:
            print("⚠️ Opción inválida.")


def flow_update_data() -> None:
    print("\nActualización de Datos")
    print("1) Exportar DB -> Excel")
    print("2) Importar Excel -> DB (con respaldo y deduplicación)")
    op = input("Opción: ").strip()
    if op == "1":
        update_excel_from_db()
        print("Excel actualizado desde la base de datos.")
    elif op == "2":
        count = update_db_from_excel()
        print(f"Base de datos actualizada desde Excel. Registros leídos: {count}")
    else:
        print("Opción inválida.")


def login():
    user = input("Usuario: ").strip()
    is_admin = False
    if user.lower() == "admin":
        pwd = input("Contraseña de admin: ").strip()
        if pwd == SET["admin_password"]:
            is_admin = True
        else:
            print("Contraseña incorrecta. Entrando como usuario normal.")
    session_id = f"{user}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return user, session_id, is_admin


def menu(user: str, session_id: str, is_admin: bool) -> None:
    while True:
        print("\n=== MENÚ ===")
        print("1) Crear RMA")
        if is_admin:
            print("2) Validar garantía")
            print("3) Información de RMA")
            print("4) Reposición de mercancía")
            print("5) Estatus de garantía")
            print("6) Buscador avanzado")
            print("7) Actualización de datos (Excel ↔ DB)")
            print("8) Salir")
        else:
            print("2) Salir")

        op = input("Opción: ").strip()
        if not is_admin:
            if op == "1":
                flow_create_rma()
            elif op == "2":
                break
            else:
                print("Opción inválida.")
        else:
            if op == "1":
                flow_create_rma()
            elif op == "2":
                flow_validate_guarantee()
            elif op == "3":
                flow_info_rma()
            elif op == "4":
                flow_reposicion()
            elif op == "5":
                flow_estatus_garantia()
            elif op == "6":
                flow_search()
            elif op == "7":
                flow_update_data()
            elif op == "8":
                break
            else:
                print("Opción inválida.")


if __name__ == "__main__":
    os.makedirs(PDF_DIR, exist_ok=True)
    user, session_id, is_admin = login()
    menu(user, session_id, is_admin)
