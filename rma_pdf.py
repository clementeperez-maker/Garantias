import os

from fpdf import FPDF


class PDF(FPDF):
    def __init__(self, settings, title, folio):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.settings = settings
        self.title_text = title
        self.folio_text = folio
        self.set_auto_page_break(auto=True, margin=18)
        self.set_margins(15, 18, 15)
        self.alias_nb_pages()

    def header(self):
        logo = self.settings.get("logo_path", "logo.png")
        if logo and os.path.exists(logo):
            try:
                self.image(logo, x=15, y=8, w=28)
            except RuntimeError:
                pass

        self.set_xy(-80, 10)
        self.set_font("Arial", "B", 12)
        self.cell(70, 8, self.folio_text or "", align="R", ln=1)

        self.set_y(20)
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, self.title_text, ln=1, align="C")
        self.ln(20)

    def footer(self):
        self.set_y(-18)
        self.set_font("Arial", "", 9)
        self.multi_cell(
            0,
            5,
            "Dirección: Bv. Francisco González Bocanegra 2054, Azteca, 37520 León de los Aldama, Gto.\n"
            "Teléfono: 4775958956",
            align="C",
        )

    def line_item(self, label, value):
        self.set_x(self.l_margin)
        self.set_font("Arial", "B", 12)
        self.cell(60, 7, f"{label}", ln=0)
        self.set_font("Arial", "", 12)
        self.cell(0, 7, str(value or ""), ln=1)

    def spacer(self, h=4):
        self.ln(h)

    def signatures_bottom(self, left_title, left_name, right_title, right_name):
        self.set_y(self.h - 60)
        sig_w = (self.w - self.l_margin - self.r_margin - 10) / 2.0

        self.set_x(self.l_margin)
        self.set_font("Arial", "", 12)
        self.cell(sig_w, 10, "", border="B", ln=0)
        self.cell(10, 10, "", ln=0)
        self.cell(sig_w, 10, "", border="B", ln=1)

        self.set_x(self.l_margin)
        self.set_font("Arial", "B", 12)
        self.cell(sig_w, 5, left_title, ln=0, align="C")
        self.cell(10, 5, "", ln=0)
        self.cell(sig_w, 5, right_title, ln=1, align="C")

        self.set_x(self.l_margin)
        self.set_font("Arial", "", 12)
        self.cell(sig_w, 5, left_name or "", ln=0, align="C")
        self.cell(10, 5, "", ln=0)
        self.cell(sig_w, 5, right_name or "", ln=1, align="C")

    def ensure_single_page(self):
        while self.get_y() < (self.h - 70):
            self.ln(5)


def _render_block(pdf: PDF, pairs):
    for label, value in pairs:
        pdf.line_item(label, value)
        pdf.spacer(2)


def generate_rma_pdf(data: dict, path: str, settings: dict):
    pdf = PDF(settings, "Solicitud de Garantía", data.get("folio", ""))
    pdf.add_page()
    _render_block(
        pdf,
        [
            ("Fecha de Recepción:", data.get("fecha_recepcion", "")),
            ("Cliente:", data.get("nombre_cliente", "")),
            ("Teléfono:", data.get("telefono", "")),
            ("Factura:", data.get("factura", "")),
            ("Producto:", data.get("producto", "")),
            ("Marca:", data.get("marca", "")),
            ("Modelo:", data.get("modelo", "")),
            ("Serie:", data.get("serie", "")),
            ("Falla:", data.get("falla", "")),
            ("Observaciones:", data.get("observaciones", "")),
            ("Responsable:", data.get("responsable", "")),
        ],
    )
    pdf.ensure_single_page()
    pdf.output(path)


def generate_validacion_pdf(data: dict, path: str, settings: dict):
    tipo = (data.get("tipo_resolucion") or "").strip().lower()
    pdf = PDF(settings, "RMA - Garantía", data.get("folio", ""))
    pdf.add_page()
    _render_block(
        pdf,
        [
            ("Fecha de Recepción:", data.get("fecha_recepcion", "")),
            ("Cliente:", data.get("nombre_cliente", "")),
            ("Teléfono:", data.get("telefono", "")),
            ("Factura:", data.get("factura", "")),
            ("Producto:", data.get("producto", "")),
            ("Marca:", data.get("marca", "")),
            ("Modelo:", data.get("modelo", "")),
            ("Serie:", data.get("serie", "")),
            ("Falla:", data.get("falla", "")),
            ("Observaciones:", data.get("observaciones", "")),
            ("Responsable:", data.get("responsable", "")),
            ("Fecha de Entrega:", data.get("fecha_entrega", "")),
        ],
    )
    pdf.spacer(5)
    if tipo == "equipo nuevo":
        pdf.set_font("Arial", "B", 13)
        pdf.cell(0, 8, "Datos de Reposición", ln=1, align="C")
        pdf.spacer(3)
        _render_block(
            pdf,
            [
                ("Producto Nuevo:", data.get("producto_nuevo", "")),
                ("Modelo Nuevo:", data.get("modelo_nuevo", "")),
                ("Serie Nueva:", data.get("serie_nueva", "")),
            ],
        )
    elif tipo in ["reparación", "reparacion"]:
        pdf.set_font("Arial", "B", 13)
        pdf.cell(0, 8, "Datos de la Reparación", ln=1, align="C")
        pdf.spacer(3)
        _render_block(pdf, [("Detalles Reparación:", data.get("detalles_reparacion", ""))])
    elif tipo in ["nota de crédito", "nota de credito"]:
        pdf.set_font("Arial", "B", 13)
        pdf.cell(0, 8, "Resolución", ln=1, align="C")
        pdf.spacer(3)
        _render_block(pdf, [("Detalles:", "Nota de crédito")])
    pdf.ensure_single_page()
    pdf.signatures_bottom(
        "Firma del Cliente",
        data.get("nombre_cliente", ""),
        "Firma del Responsable",
        data.get("responsable", ""),
    )
    pdf.output(path)


def generate_no_valida_pdf(data: dict, path: str, settings: dict):
    pdf = PDF(settings, "RMA - Garantía", data.get("folio", ""))
    pdf.add_page()
    _render_block(
        pdf,
        [
            ("Fecha de Recepción:", data.get("fecha_recepcion", "")),
            ("Cliente:", data.get("nombre_cliente", "")),
            ("Teléfono:", data.get("telefono", "")),
            ("Factura:", data.get("factura", "")),
            ("Producto:", data.get("producto", "")),
            ("Marca:", data.get("marca", "")),
            ("Modelo:", data.get("modelo", "")),
            ("Serie:", data.get("serie", "")),
            ("Falla:", data.get("falla", "")),
            ("Observaciones:", data.get("observaciones", "")),
            ("Responsable:", data.get("responsable", "")),
            ("Fecha de Entrega:", data.get("fecha_entrega", "")),
        ],
    )
    pdf.spacer(5)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(0, 8, "Garantía No Válida", ln=1, align="C")
    pdf.spacer(3)
    _render_block(pdf, [("Detalles:", data.get("detalles_no_valida", ""))])
    pdf.ensure_single_page()
    pdf.signatures_bottom(
        "Firma del Cliente",
        data.get("nombre_cliente", ""),
        "Firma del Responsable",
        data.get("responsable", ""),
    )
    pdf.output(path)


def generate_reposicion_pdf(data: dict, path: str, settings: dict):
    pdf = PDF(settings, "Reposición de Mercancía", data.get("folio", ""))
    pdf.add_page()
    _render_block(
        pdf,
        [
            ("Fecha de Recepción:", data.get("fecha_recepcion", "")),
            ("Cliente:", data.get("nombre_cliente", "")),
            ("Factura:", data.get("factura", "")),
            ("Producto:", data.get("producto", "")),
            ("Modelo:", data.get("modelo", "")),
            ("Serie:", data.get("serie", "")),
        ],
    )
    pdf.spacer(5)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(0, 8, "Datos de Reposición", ln=1, align="C")
    pdf.spacer(3)
    _render_block(
        pdf,
        [
            ("Fecha de Reposición:", data.get("fecha_reposicion", "")),
            ("Detalles:", data.get("detalles_reposicion", "")),
            ("Reemplazo:", data.get("reemplazo", "")),
            ("Responsable:", data.get("responsable", "")),
        ],
    )
    pdf.ensure_single_page()
    pdf.signatures_bottom(
        "Responsable",
        data.get("responsable", ""),
        "Encargado de almacén",
        data.get("encargado_almacen", ""),
    )
    pdf.output(path)
