"""
Genera un Excel de Flujo de Caja para un Airbnb.
Hojas: Parámetros | Flujo de Caja Mensual | Resumen
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel


# ── Colores ──────────────────────────────────────────────────────────────────
C_DARK   = "1A1A2E"   # fondo encabezado
C_ACCENT = "E94560"   # acento rojo-rosa
C_LIGHT  = "16213E"   # fondo filas alternas oscuro
C_BLUE   = "0F3460"   # azul medio
C_WHITE  = "FFFFFF"
C_GRAY   = "F2F2F2"
C_GREEN  = "2ECC71"
C_RED    = "E74C3C"
C_YELLOW = "F39C12"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color=C_WHITE, size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic,
                name="Calibri")

def border_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ── Helpers ───────────────────────────────────────────────────────────────────
def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def style_header_cell(cell, bg=C_DARK, fg=C_WHITE, size=11, bold=True):
    cell.fill = fill(bg)
    cell.font = font(bold=bold, color=fg, size=size)
    cell.alignment = align()
    cell.border = border_thin()

def style_data_cell(cell, bg=C_WHITE, fg="1A1A2E", bold=False,
                    h_align="right", num_fmt=None):
    cell.fill = fill(bg)
    cell.font = font(bold=bold, color=fg, size=10)
    cell.alignment = align(h=h_align)
    cell.border = border_thin()
    if num_fmt:
        cell.number_format = num_fmt


CLP = '#,##0'
PCT = '0.0%'

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 1: PARÁMETROS
# ═══════════════════════════════════════════════════════════════════════════════
def build_parametros(wb):
    ws = wb.create_sheet("Parámetros")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 20

    # anchos
    for col, w in [("A", 34), ("B", 22), ("C", 18), ("D", 2)]:
        set_col_width(ws, col, w)

    # título
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "PARÁMETROS – FLUJO DE CAJA AIRBNB"
    c.fill = fill(C_DARK)
    c.font = font(bold=True, size=15)
    c.alignment = align()

    secciones = [
        ("HORIZONTE Y OPERACIÓN", [
            ("Horizonte (meses)",               36,       None),
            ("Meses de gracia (sin ingresos)",  6,        None),
            ("ADR promedio (CLP / noche)",       38_000,   CLP),
            ("Noches promedio por mes",          21,       None),
            ("Crecimiento anual ADR (%)",        0.03,     PCT),
        ]),
        ("COSTOS OPERATIVOS (CLP/mes)", [
            ("Comisión administración/Airbnb",  0.18,     PCT),
            ("Gastos comunes",                  90_000,   CLP),
            ("Servicios (luz, agua, internet)", 60_000,   CLP),
            ("Fondo de mantención",             40_000,   CLP),
            ("Dividendo / arriendo",            274_000,  CLP),
        ]),
        ("INVERSIÓN Y FINANCIAMIENTO", [
            ("Inversión inicial (amoblado)",    3_200_000, CLP),
            ("Tasa de descuento anual (%)",     0.12,      PCT),
        ]),
    ]

    named_ranges = {}   # label -> cell address for cross-sheet references
    row = 3
    for titulo, items in secciones:
        # encabezado de sección
        ws.merge_cells(f"A{row}:C{row}")
        c = ws[f"A{row}"]
        c.value = titulo
        c.fill = fill(C_BLUE)
        c.font = font(bold=True, size=10)
        c.alignment = align(h="left")
        ws[f"B{row}"].fill = fill(C_BLUE)
        ws[f"C{row}"].fill = fill(C_BLUE)
        row += 1

        for label, valor, fmt in items:
            bg = C_GRAY if row % 2 == 0 else C_WHITE
            # etiqueta
            cA = ws[f"A{row}"]
            cA.value = label
            style_data_cell(cA, bg=bg, h_align="left", fg="1A1A2E")
            # valor
            cB = ws[f"B{row}"]
            cB.value = valor
            style_data_cell(cB, bg=bg, num_fmt=fmt, h_align="right")
            cB.fill = fill("FFFDE7")   # amarillo suave → editable
            cB.font = font(bold=True, color=C_BLUE, size=10)
            # unidad
            cC = ws[f"C{row}"]
            cC.fill = fill(bg)
            cC.border = border_thin()
            named_ranges[label] = f"'Parámetros'!$B${row}"
            row += 1
        row += 1   # espacio entre secciones

    # leyenda
    row += 1
    ws.merge_cells(f"A{row}:C{row}")
    c = ws[f"A{row}"]
    c.value = "Las celdas en amarillo son editables. Todos los demás valores se calculan automáticamente."
    c.font = font(italic=True, color="888888", size=9)
    c.alignment = align(h="left")

    return named_ranges


# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 2: FLUJO DE CAJA MENSUAL
# ═══════════════════════════════════════════════════════════════════════════════
def build_flujo(wb, params_addr):
    ws = wb.create_sheet("Flujo de Caja Mensual")
    ws.sheet_view.showGridLines = False

    # referencias a Parámetros (por índice de fila en params_addr)
    def ref(label):
        return params_addr[label]

    # ── anchos de columna ──
    headers = [
        "Mes", "Período", "ADR\n(CLP/noche)", "Noches\npromedio",
        "Ing. Brutos\n(CLP)", "Comisión\nAdmin/Airbnb",
        "Ing. Netos\n(CLP)", "Gastos\nComunes",
        "Servicios", "Fondo\nMantención",
        "Dividendo", "Total\nEgresos",
        "Flujo Neto\n(CLP)", "Flujo\nAcumulado\n(CLP)"
    ]
    col_widths = [7, 14, 14, 11, 16, 16, 14, 13, 13, 14, 13, 13, 14, 16]
    for i, (_, w) in enumerate(zip(headers, col_widths), 1):
        set_col_width(ws, get_column_letter(i), w)

    # ── título ──
    ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
    c = ws["A1"]
    c.value = "FLUJO DE CAJA MENSUAL – AIRBNB"
    c.fill = fill(C_DARK)
    c.font = font(bold=True, size=14)
    c.alignment = align()
    ws.row_dimensions[1].height = 35

    # ── encabezados de columna ──
    ws.row_dimensions[2].height = 42
    for col_i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_i, value=h)
        c.fill = fill(C_ACCENT)
        c.font = font(bold=True, size=9)
        c.alignment = align(wrap=True)
        c.border = border_thin()

    # ── referencia corta a celdas de Parámetros ──
    # Row numbers in Parámetros sheet (1-indexed, cuenta desde fila 3):
    # Horizonte     -> row 4
    # Meses gracia  -> row 5
    # ADR           -> row 6
    # Noches        -> row 7
    # Crec ADR      -> row 8
    # [sep row 9]
    # Comisión      -> row 11
    # Gastos com    -> row 12
    # Servicios     -> row 13
    # Fondo mant    -> row 14
    # Dividendo     -> row 15
    # [sep row 16]
    # Inversión     -> row 18
    # Tasa desc     -> row 19

    P = "Parámetros"   # shorthand

    horizonte = 36
    meses_gracia = 6

    inv_row = 3   # row offset counter used during build_parametros
    # We hard-code row numbers matching build_parametros logic:
    ROW = {
        "horizonte":    4,
        "gracia":       5,
        "adr":          6,
        "noches":       7,
        "crec_adr":     8,
        "comision":    11,
        "g_comunes":   12,
        "servicios":   13,
        "fondo":       14,
        "dividendo":   15,
        "inversion":   18,
        "tasa":        19,
    }

    def p(key):
        return f"'{P}'!$B${ROW[key]}"

    # ── datos mes a mes ──
    data_start = 3
    import datetime
    start_date = datetime.date(2025, 1, 1)

    for mes in range(1, horizonte + 1):
        row = data_start + mes - 1
        ws.row_dimensions[row].height = 18
        bg = "EEF2FF" if mes % 2 == 0 else C_WHITE
        en_gracia = mes <= meses_gracia

        fecha = start_date.replace(
            month=((start_date.month + mes - 2) % 12) + 1,
            year=start_date.year + (start_date.month + mes - 2) // 12
        )
        periodo = fecha.strftime("%b %Y")

        # Col A: Mes
        c = ws.cell(row=row, column=1, value=mes)
        style_data_cell(c, bg=bg, h_align="center")

        # Col B: Período
        c = ws.cell(row=row, column=2, value=periodo)
        style_data_cell(c, bg=bg, h_align="center")

        # Col C: ADR con crecimiento anual compuesto
        # ADR_mes = ADR_base * (1+g)^floor((mes-1)/12)
        año_idx = (mes - 1) // 12
        adr_formula = f"={p('adr')}*(1+{p('crec_adr')})^{año_idx}"
        c = ws.cell(row=row, column=3, value=adr_formula if not en_gracia else 0)
        style_data_cell(c, bg="FFE0E0" if en_gracia else bg, num_fmt=CLP)

        # Col D: Noches (0 en gracia)
        noches_val = f"=IF({mes}<={p('gracia')},0,{p('noches')})"
        c = ws.cell(row=row, column=4, value=noches_val)
        style_data_cell(c, bg="FFE0E0" if en_gracia else bg, h_align="center",
                        num_fmt="0")

        # Col E: Ingresos brutos = ADR * noches
        c_adr = get_column_letter(3)
        c_noch = get_column_letter(4)
        ing_brutos = f"={c_adr}{row}*{c_noch}{row}"
        c = ws.cell(row=row, column=5, value=ing_brutos)
        style_data_cell(c, bg="FFE0E0" if en_gracia else bg, num_fmt=CLP)

        # Col F: Comisión
        comision_f = f"=E{row}*{p('comision')}"
        c = ws.cell(row=row, column=6, value=comision_f)
        style_data_cell(c, bg="FFE0E0" if en_gracia else bg, num_fmt=CLP)

        # Col G: Ingresos netos
        ing_netos = f"=E{row}-F{row}"
        c = ws.cell(row=row, column=7, value=ing_netos)
        c.fill = fill("D5F5E3" if not en_gracia else "FFE0E0")
        c.font = font(bold=True, color="1A6B3A" if not en_gracia else "B71C1C",
                      size=10)
        c.alignment = align(h="right")
        c.border = border_thin()
        c.number_format = CLP

        # Col H: Gastos comunes
        c = ws.cell(row=row, column=8, value=f"={p('g_comunes')}")
        style_data_cell(c, bg=bg, num_fmt=CLP)

        # Col I: Servicios
        c = ws.cell(row=row, column=9, value=f"={p('servicios')}")
        style_data_cell(c, bg=bg, num_fmt=CLP)

        # Col J: Fondo mantención
        c = ws.cell(row=row, column=10, value=f"={p('fondo')}")
        style_data_cell(c, bg=bg, num_fmt=CLP)

        # Col K: Dividendo
        c = ws.cell(row=row, column=11, value=f"={p('dividendo')}")
        style_data_cell(c, bg=bg, num_fmt=CLP)

        # Col L: Total egresos
        c = ws.cell(row=row, column=12, value=f"=SUM(H{row}:K{row})")
        style_data_cell(c, bg=bg, bold=True, num_fmt=CLP)

        # Col M: Flujo neto
        flujo_neto = f"=G{row}-L{row}"
        c = ws.cell(row=row, column=13, value=flujo_neto)
        # conditional color via value (static at generation time, formula stays)
        c.fill = fill(bg)
        c.font = font(bold=True, color="1A6B3A" if not en_gracia else "B71C1C",
                      size=10)
        c.alignment = align(h="right")
        c.border = border_thin()
        c.number_format = CLP

        # Col N: Flujo acumulado
        if mes == 1:
            flujo_acum = f"=-{p('inversion')}+M{row}"
        else:
            flujo_acum = f"=N{row-1}+M{row}"
        c = ws.cell(row=row, column=14, value=flujo_acum)
        c.fill = fill(bg)
        c.font = font(bold=True, color="333333", size=10)
        c.alignment = align(h="right")
        c.border = border_thin()
        c.number_format = CLP

    # ── fila de totales ──
    tot_row = data_start + horizonte
    ws.row_dimensions[tot_row].height = 22
    ws.merge_cells(f"A{tot_row}:F{tot_row}")
    c = ws[f"A{tot_row}"]
    c.value = "TOTALES"
    c.fill = fill(C_DARK)
    c.font = font(bold=True, size=11)
    c.alignment = align()

    for col in range(7, 15):
        col_l = get_column_letter(col)
        c = ws.cell(row=tot_row, column=col,
                    value=f"=SUM({col_l}{data_start}:{col_l}{data_start+horizonte-1})"
                          if col != 14 else
                          f"={get_column_letter(14)}{data_start+horizonte-1}")
        c.fill = fill(C_DARK)
        c.font = font(bold=True, size=10)
        c.alignment = align(h="right")
        c.border = border_thin()
        c.number_format = CLP

    # ── Inmovilizar panel ──
    ws.freeze_panes = "C3"

    return ws, data_start, horizonte, ROW, P


# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 3: RESUMEN / KPIs
# ═══════════════════════════════════════════════════════════════════════════════
def build_resumen(wb, data_start, horizonte, ROW, P):
    ws = wb.create_sheet("Resumen")
    ws.sheet_view.showGridLines = False

    set_col_width(ws, "A", 36)
    set_col_width(ws, "B", 22)
    set_col_width(ws, "C", 18)

    # título
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "RESUMEN EJECUTIVO – FLUJO DE CAJA AIRBNB"
    c.fill = fill(C_DARK)
    c.font = font(bold=True, size=14)
    c.alignment = align()
    ws.row_dimensions[1].height = 38

    def p(key):
        return f"'{P}'!$B${ROW[key]}"

    fc_sheet = "'Flujo de Caja Mensual'"
    last_row = data_start + horizonte - 1
    first_data = data_start

    # Referencia a col N (flujo acumulado) de la hoja FC
    def fc_col(col_letter, row):
        return f"={fc_sheet}!{col_letter}{row}"

    kpis = [
        ("INVERSIÓN", [
            ("Inversión inicial (amoblado)", f"={p('inversion')}", CLP),
            ("Horizonte de análisis",        f"={p('horizonte')}&\" meses\"", "@"),
            ("Meses de gracia",              f"={p('gracia')}", "0"),
        ]),
        ("INGRESOS ESPERADOS", [
            ("ADR promedio inicial (CLP/noche)", f"={p('adr')}", CLP),
            ("Noches promedio / mes",            f"={p('noches')}", "0"),
            ("Ingreso bruto mensual promedio",
             f"={p('adr')}*{p('noches')}", CLP),
            ("Comisión administración/Airbnb",   f"={p('comision')}", PCT),
            ("Ingreso neto mensual promedio",
             f"={p('adr')}*{p('noches')}*(1-{p('comision')})", CLP),
        ]),
        ("EGRESOS FIJOS MENSUALES", [
            ("Gastos comunes",      f"={p('g_comunes')}", CLP),
            ("Servicios",           f"={p('servicios')}", CLP),
            ("Fondo mantención",    f"={p('fondo')}",     CLP),
            ("Dividendo / arriendo",f"={p('dividendo')}",  CLP),
            ("Total egresos fijos",
             f"={p('g_comunes')}+{p('servicios')}+{p('fondo')}+{p('dividendo')}",
             CLP),
        ]),
        ("INDICADORES FINANCIEROS", [
            ("Tasa de descuento anual",  f"={p('tasa')}", PCT),
            ("Tasa de descuento mensual",
             f"=(1+{p('tasa')})^(1/12)-1", "0.000%"),
            ("Flujo neto total (36 meses)",
             f"={fc_sheet}!M{last_row+1}", CLP),  # total row
            ("Flujo acumulado final",
             f"={fc_sheet}!N{last_row}", CLP),
            # VAN: inversión en mes 0 + flujos descontados
            ("VAN (Valor Actual Neto)",
             f"=NPV((1+{p('tasa')})^(1/12)-1,"
             f"{fc_sheet}!M{first_data}:{fc_sheet}!M{last_row})"
             f"-{p('inversion')}",
             CLP),
            # TIR: incluye -inversión en mes 0
            ("TIR mensual",
             f"=IFERROR(IRR(IF({{1}},{{-{p('inversion')},"
             f"{fc_sheet}!M{first_data}:{fc_sheet}!M{last_row}}})),\"N/D\")",
             "0.00%"),
            ("Payback (mes aprox.)",
             f"=IFERROR(MATCH(0,IF({fc_sheet}!N{first_data}:{fc_sheet}!N{last_row}>0,1,0),0)"
             f"+{p('gracia')},\"No recuperado\")",
             "0"),
        ]),
    ]

    row = 3
    for titulo, items in kpis:
        ws.merge_cells(f"A{row}:C{row}")
        c = ws[f"A{row}"]
        c.value = titulo
        c.fill = fill(C_BLUE)
        c.font = font(bold=True, size=10)
        c.alignment = align(h="left")
        ws[f"B{row}"].fill = fill(C_BLUE)
        ws[f"C{row}"].fill = fill(C_BLUE)
        row += 1

        for label, formula, fmt in items:
            bg = C_GRAY if row % 2 == 0 else C_WHITE
            cA = ws[f"A{row}"]
            cA.value = label
            style_data_cell(cA, bg=bg, h_align="left")

            cB = ws[f"B{row}"]
            # Algunos son fórmulas array, otros texto
            if formula.startswith("="):
                cB.value = formula
            else:
                cB.value = formula
            cB.fill = fill(bg)
            cB.font = font(bold=True, color=C_DARK, size=10)
            cB.alignment = align(h="right")
            cB.border = border_thin()
            if fmt and fmt != "@":
                cB.number_format = fmt

            cC = ws[f"C{row}"]
            cC.fill = fill(bg)
            cC.border = border_thin()
            row += 1
        row += 1

    # Nota VAN
    row += 1
    ws.merge_cells(f"A{row}:C{row}")
    c = ws[f"A{row}"]
    c.value = ("Nota: VAN positivo indica que el proyecto genera valor por "
               "encima de la tasa de descuento elegida (12% anual). "
               "TIR se calcula con flujos mensuales incluyendo la inversión inicial en t=0.")
    c.font = font(italic=True, color="888888", size=9)
    c.alignment = align(h="left", wrap=True)
    ws.row_dimensions[row].height = 36


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    # Quitar hoja default
    wb.remove(wb.active)

    params_addr = build_parametros(wb)
    ws_fc, data_start, horizonte, ROW, P = build_flujo(wb, params_addr)
    build_resumen(wb, data_start, horizonte, ROW, P)

    out = "/home/user/flujo-de-caja/airbnb/flujo_caja_airbnb.xlsx"
    wb.save(out)
    print(f"Archivo generado: {out}")


if __name__ == "__main__":
    main()
