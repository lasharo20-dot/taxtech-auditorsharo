import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="TaxTech Auditor RD",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── ESTILOS CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stSidebar"] { background: #0f1923; }
    [data-testid="stSidebar"] * { color: #e8edf2 !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stTextInput label,
    [data-testid="stSidebar"] .stSlider label { color: #94a3b8 !important; font-size: 0.78rem !important; }
    [data-testid="stSidebar"] h1 { color: #38bdf8 !important; font-size: 1rem !important; letter-spacing: 0.05em; text-transform: uppercase; }
    .block-container { padding-top: 1.5rem; }
    div[data-testid="metric-container"] { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 1rem; }
    div[data-testid="metric-container"] label { font-size: 0.75rem !important; color: #64748b !important; text-transform: uppercase; letter-spacing: 0.06em; }
    .fin-table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    .fin-table th { background: #1e3a5f; color: white; padding: 8px 12px; text-align: left; }
    .fin-table td { padding: 6px 12px; border-bottom: 1px solid #e2e8f0; }
    .fin-table tr:nth-child(even) { background: #f8fafc; }
    .fin-table .subtotal td { background: #dbeafe; font-weight: 600; }
    .fin-table .total td { background: #1e3a5f; color: white; font-weight: 700; }
    .fin-table .section-header td { background: #f1f5f9; font-weight: 700; color: #1e3a5f; letter-spacing: 0.04em; text-transform: uppercase; font-size: 0.78rem; }
    .fin-table .negative { color: #dc2626; }
    .fin-table .positive { color: #16a34a; }
    .stTabs [data-baseweb="tab"] { font-size: 0.82rem; padding: 6px 14px; }
    .stTabs [aria-selected="true"] { border-bottom: 2px solid #1e3a5f; font-weight: 700; }
    h1 { color: #1e3a5f !important; }
    h3 { color: #1e3a5f !important; font-size: 1rem !important; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# PARÁMETROS FISCALES RD 2026
# ──────────────────────────────────────────────────────────────────────────────
NATURALEZAS = {'1': 'Debito', '2': 'Credito', '3': 'Credito', '4': 'Credito', '5': 'Debito', '6': 'Debito'}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Art. 287 CTRD: Validar NCF válido y medios de pago para crédito ITBIS.',
    'representacion': 'Art. 287 CTRD: Razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Art. 318 / Reg. 139-98: Retribuciones en especie. Validar ISR sustitutivo.',
    'gasto de personal': 'Art. 287 CTRD: Cruce obligatorio con IR-4 (TSS) para admitir deducción.',
    'honorario': 'Art. 309 CTRD: Retención 10% personas físicas / 2% entre jurídicas.',
    'viaje': 'Art. 287 CTRD: Gastos de viaje. Documentación de viáticos y propósito del viaje.',
    'donacion': 'Art. 287 CTRD: Donaciones solo deducibles si beneficiario es Ley 122-05.',
    'multa': 'Art. 287 CTRD: Multas y recargos NO son gastos deducibles.',
    'perdida': 'Art. 287 CTRD: Validar naturaleza de la pérdida para deducibilidad.',
}

TASA_ITBIS = 0.18
TASA_SFS_PAT = 0.0709; TASA_AFP_PAT = 0.0710
TASA_SRL = 0.0120; TASA_INFOTEP = 0.0100
TASA_SFS_EMP = 0.0304; TASA_AFP_EMP = 0.0287
COSTO_PERCAPITA_2026 = 1691.38

# Excel styles
FILL_HDR   = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_SEC   = PatternFill(start_color="EBF3FA", end_color="EBF3FA", fill_type="solid")
FNT_TITLE  = Font(name="Calibri", size=13, bold=True, color="1F497D")
FNT_HDR    = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FNT_BODY   = Font(name="Calibri", size=10)
FNT_BOLD   = Font(name="Calibri", size=10, bold=True)
FNT_SEC    = Font(name="Calibri", size=10, bold=True, color="1F497D")
THIN       = Side(border_style="thin", color="D9D9D9")
BRD        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_RD     = 'RD$ #,##0.00'
FMT_PCT    = '0.00%'

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS EXCEL
# ──────────────────────────────────────────────────────────────────────────────
def xl_header(ws, title, subtitle, cols):
    ws["B2"] = title;   ws["B2"].font = FNT_TITLE
    ws["B3"] = subtitle; ws["B3"].font = Font(name="Calibri", size=10, italic=True)
    for i, h in enumerate(cols, 2):
        c = ws.cell(row=5, column=i, value=h)
        c.font = FNT_HDR; c.fill = FILL_HDR; c.border = BRD
        c.alignment = Alignment(horizontal="center")

def xl_col_widths(ws):
    for col in ws.columns:
        ltr = get_column_letter(col[0].column)
        if ltr != 'A':
            ws.column_dimensions[ltr].width = 45 if ltr in ['C', 'D'] else 22

def xl_money(cell, value):
    cell.value = value; cell.number_format = FMT_RD
    cell.alignment = Alignment(horizontal="right")
    cell.font = FNT_BODY; cell.border = BRD

# ──────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO DE BALANZA
# ──────────────────────────────────────────────────────────────────────────────
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.lower().endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
        elif file.name.lower().endswith('.xls'):
            df = pd.read_excel(file, engine='xlrd')
        else:
            df = pd.read_csv(file)
        df.columns = [str(c).strip().lower() for c in df.columns]
        mapeo = {}
        for col in df.columns:
            if any(x in col for x in ['código', 'codigo', 'cuenta no', 'no.', 'cuenta_no', 'cta']):
                mapeo[col] = 'codigo'
            elif any(x in col for x in ['nombre', 'descripción', 'descripcion', 'concepto']) and 'codigo' not in col:
                mapeo[col] = 'cuenta'
            elif 'cuenta' in col and 'codigo' not in col and 'cuenta' not in mapeo.values():
                mapeo[col] = 'cuenta'
            elif any(x in col for x in ['débito', 'debito', 'debe', 'cargos']):
                mapeo[col] = 'debito'
            elif any(x in col for x in ['crédito', 'credito', 'haber', 'abonos']):
                mapeo[col] = 'credito'
            elif any(x in col for x in ['saldo', 'balance', 'final', 'monto']):
                mapeo[col] = 'saldo_final'
        df = df.rename(columns=mapeo)
        if 'saldo_final' not in df.columns and 'debito' in df.columns and 'credito' in df.columns:
            df['saldo_final'] = df['debito'] - df['credito']
        col_req = {'codigo', 'cuenta', 'debito', 'credito', 'saldo_final'}
        missing = col_req - set(df.columns)
        if missing:
            st.error(f"⚠️ Columnas detectadas: {list(df.columns)}. Faltan: {missing}")
            return pd.DataFrame()
        df['codigo'] = df['codigo'].fillna('').astype(str).str.strip()
        df['codigo'] = df['codigo'].apply(lambda x: x.split('.')[0] if '.' in x else x)
        df['cuenta'] = df['cuenta'].fillna('').astype(str).str.strip()
        df = df[(df['codigo'] != '') & (~df['codigo'].str.lower().str.contains('total|resultado|suma|nan', na=False))]
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ Error al procesar la balanza: {e}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat, alertas_fisc = [], []
    for _, row in df.iterrows():
        cod, nom = str(row['codigo']), str(row['cuenta']).lower()
        nat_esp = NATURALEZAS.get(cod[0] if cod else '', None)
        saldo = row['saldo_final']
        if nat_esp == 'Debito' and saldo < 0:
            alertas_nat.append("⚠️ Saldo crédito (nat. débito)")
        elif nat_esp == 'Credito' and saldo > 0:
            alertas_nat.append("⚠️ Saldo débito (nat. crédito)")
        else:
            alertas_nat.append("✅ Correcto")
        alerta_f = ""
        for palabra, msg in PALABRAS_CRITICAS_ART287.items():
            if palabra in nom:
                alerta_f = msg; break
        alertas_fisc.append(alerta_f)
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal'] = alertas_fisc
    return df

# ──────────────────────────────────────────────────────────────────────────────
# ESTADO FINANCIERO DESDE BALANZA
# ──────────────────────────────────────────────────────────────────────────────
def fmt_rd(v): return f"RD$ {v:>14,.2f}"
def fmt_neg(v): return f"(RD$ {abs(v):>13,.2f})" if v < 0 else fmt_rd(v)

def generar_balance_general(df: pd.DataFrame) -> dict:
    """Clasifica cuentas en estructura de Balance General PCGA-RD."""
    result = {
        'activo_corriente': [],
        'activo_no_corriente': [],
        'pasivo_corriente': [],
        'pasivo_no_corriente': [],
        'patrimonio': [],
    }
    for _, row in df.iterrows():
        cod = str(row['codigo'])
        if not cod or len(cod) == 0:
            continue
        saldo = abs(row['saldo_final']) if row['saldo_final'] != 0 else 0
        nombre = row['cuenta']
        entry = (nombre, cod, saldo)
        p = cod[0]
        if p == '1':
            # Activo: subcategoría por segundo dígito
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3', '4'):
                result['activo_corriente'].append(entry)
            else:
                result['activo_no_corriente'].append(entry)
        elif p == '2':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3'):
                result['pasivo_corriente'].append(entry)
            else:
                result['pasivo_no_corriente'].append(entry)
        elif p == '3':
            result['patrimonio'].append(entry)
    return result

def generar_estado_resultados(df: pd.DataFrame) -> dict:
    """Clasifica cuentas 4, 5, 6 para P&L."""
    result = {'ingresos': [], 'costos': [], 'gastos_operacion': [], 'otros_ingresos': [], 'gastos_financieros': []}
    for _, row in df.iterrows():
        cod = str(row['codigo'])
        if not cod: continue
        p = cod[0]
        saldo = abs(row['saldo_final'])
        nombre = row['cuenta']
        entry = (nombre, cod, saldo)
        if p == '4':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3', '4', '5'):
                result['ingresos'].append(entry)
            else:
                result['otros_ingresos'].append(entry)
        elif p == '5':
            result['costos'].append(entry)
        elif p == '6':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('3', '4'):
                result['gastos_financieros'].append(entry)
            else:
                result['gastos_operacion'].append(entry)
    return result

def render_estado_html(titulo, secciones):
    """Renderiza un estado financiero como tabla HTML."""
    html = f'<table class="fin-table"><tr><th colspan="3">{titulo}</th></tr>'
    grand_total = 0
    for sec_nombre, cuentas, es_resta, mostrar_subtotal in secciones:
        html += f'<tr class="section-header"><td colspan="2">{sec_nombre}</td><td></td></tr>'
        subtotal = sum(m for _, _, m in cuentas)
        for nombre, cod, monto in cuentas:
            css = ""
            html += f'<tr class="{css}"><td style="padding-left:24px">{nombre}</td><td style="font-size:0.75rem;color:#94a3b8">{cod}</td><td style="text-align:right">{"RD$ {:,.2f}".format(monto)}</td></tr>'
        if mostrar_subtotal:
            val = -subtotal if es_resta else subtotal
            grand_total += val
            css_cls = "negative" if val < 0 else ""
            html += f'<tr class="subtotal"><td colspan="2">Subtotal {sec_nombre}</td><td style="text-align:right" class="{css_cls}">{"RD$ {:,.2f}".format(subtotal)}</td></tr>'
    html += f'<tr class="total"><td colspan="2">TOTAL</td><td style="text-align:right">RD$ {grand_total:,.2f}</td></tr>'
    html += '</table>'
    return html, grand_total

# ──────────────────────────────────────────────────────────────────────────────
# LLENADO IR-2 (Ajuste Patrimonial DGII)
# ──────────────────────────────────────────────────────────────────────────────
IR2_CASILLAS = {
    1:  ("TOTAL ACTIVOS (libros inicio ejercicio)", "1"),
    27: ("TOTAL PASIVOS (libros inicio ejercicio)", "2"),
    37: ("Inventarios (Valor Fiscal)", "13"),
    38: ("Activos Categoría 1 (Edificios)", "15"),
    39: ("Activos Categoría 2 (Maquinaria y Equipo)", "15"),
    40: ("Activos Categoría 3 (Muebles y Enseres)", "15"),
}

def calcular_casillas_ir2(df: pd.DataFrame) -> dict:
    """Extrae valores de la balanza para llenar las casillas del IR-2."""
    def suma(prefijos):
        mask = df['codigo'].apply(lambda x: any(str(x).startswith(p) for p in prefijos))
        return abs(df.loc[mask, 'saldo_final'].sum())

    total_activos    = suma(['1'])
    total_pasivos    = suma(['2'])
    patrimonio       = suma(['3'])
    inventario       = suma(['13', '130', '131', '132', '133'])
    activo_fijo_cat1 = suma(['152', '153'])  # Edificios
    activo_fijo_cat2 = suma(['154', '155'])  # Maquinaria
    activo_fijo_cat3 = suma(['156', '157', '158'])  # Muebles
    otros_activos    = suma(['14', '16', '17', '18', '19'])

    # Saldo activos fiscales (simplificado: sum categorías)
    saldo_act_fiscal = activo_fijo_cat1 + activo_fijo_cat2 + activo_fijo_cat3 + inventario + otros_activos
    saldo_pasivos    = total_pasivos
    patrimonio_fisc  = saldo_act_fiscal - saldo_pasivos
    total_no_monet   = activo_fijo_cat1 + activo_fijo_cat2 + activo_fijo_cat3 + inventario

    return {
        'cas_1':  total_activos,
        'cas_27': total_pasivos,
        'cas_31': saldo_pasivos,
        'cas_26': saldo_act_fiscal,
        'cas_32': max(patrimonio_fisc, 0),
        'cas_33': total_no_monet,
        'cas_34': min(max(patrimonio_fisc, 0), total_no_monet),
        'cas_37': inventario,
        'cas_38': activo_fijo_cat1,
        'cas_39': activo_fijo_cat2,
        'cas_40': activo_fijo_cat3,
        'cas_49': total_no_monet,
    }

# ──────────────────────────────────────────────────────────────────────────────
# EXPORTAR ESTADOS FINANCIEROS A EXCEL
# ──────────────────────────────────────────────────────────────────────────────
def exportar_estados_excel(empresa, periodo, bg, er, ir2_vals):
    wb = openpyxl.Workbook()

    # ── Hoja 1: Balance General ──────────────────────────────────────────────
    ws_bg = wb.active; ws_bg.title = "Balance General"
    xl_header(ws_bg, f"BALANCE GENERAL — {empresa.upper()}", f"Al {periodo}", ["Código", "Cuenta", "Monto (RD$)"])

    row = 6
    secciones_bg = [
        ("ACTIVO CORRIENTE", bg['activo_corriente']),
        ("ACTIVO NO CORRIENTE", bg['activo_no_corriente']),
        ("PASIVO CORRIENTE", bg['pasivo_corriente']),
        ("PASIVO NO CORRIENTE", bg['pasivo_no_corriente']),
        ("PATRIMONIO", bg['patrimonio']),
    ]
    totals = {}
    for sec_nom, cuentas in secciones_bg:
        c = ws_bg.cell(row=row, column=2, value=sec_nom)
        c.font = FNT_SEC; c.fill = FILL_SEC; c.border = BRD
        ws_bg.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        row += 1
        first = row
        for nombre, cod, monto in cuentas:
            ws_bg.cell(row=row, column=2, value=cod).font = FNT_BODY
            ws_bg.cell(row=row, column=3, value=nombre).font = FNT_BODY
            xl_money(ws_bg.cell(row=row, column=4), monto)
            ws_bg.cell(row=row, column=2).border = BRD
            ws_bg.cell(row=row, column=3).border = BRD
            if row % 2 == 0:
                for c2 in [2, 3, 4]:
                    ws_bg.cell(row=row, column=c2).fill = FILL_ZEBRA
            row += 1
        # subtotal
        sub_cell = ws_bg.cell(row=row, column=3, value=f"Total {sec_nom}")
        sub_cell.font = FNT_BOLD; sub_cell.fill = FILL_TOTAL; sub_cell.border = BRD
        tot_cell = ws_bg.cell(row=row, column=4)
        tot_cell.value = f"=SUM(D{first}:D{row-1})" if first < row else 0
        tot_cell.number_format = FMT_RD; tot_cell.font = FNT_BOLD
        tot_cell.fill = FILL_TOTAL; tot_cell.border = BRD
        tot_cell.alignment = Alignment(horizontal="right")
        ws_bg.cell(row=row, column=2).fill = FILL_TOTAL; ws_bg.cell(row=row, column=2).border = BRD
        totals[sec_nom] = f"D{row}"
        row += 2

    ws_bg.column_dimensions['B'].width = 12
    ws_bg.column_dimensions['C'].width = 48
    ws_bg.column_dimensions['D'].width = 20

    # ── Hoja 2: Estado de Resultados ─────────────────────────────────────────
    ws_er = wb.create_sheet("Estado de Resultados")
    xl_header(ws_er, f"ESTADO DE RESULTADOS — {empresa.upper()}", f"Período {periodo}", ["Código", "Cuenta", "Monto (RD$)"])

    row = 6
    secciones_er = [
        ("INGRESOS POR OPERACIONES", er['ingresos']),
        ("COSTO DE VENTAS / SERVICIOS", er['costos']),
        ("GASTOS DE OPERACIÓN", er['gastos_operacion']),
        ("GASTOS FINANCIEROS", er['gastos_financieros']),
        ("OTROS INGRESOS", er['otros_ingresos']),
    ]
    er_refs = {}
    for sec_nom, cuentas in secciones_er:
        c = ws_er.cell(row=row, column=2, value=sec_nom)
        c.font = FNT_SEC; c.fill = FILL_SEC; c.border = BRD
        ws_er.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        row += 1
        first = row
        for nombre, cod, monto in cuentas:
            ws_er.cell(row=row, column=2, value=cod).font = FNT_BODY
            ws_er.cell(row=row, column=3, value=nombre).font = FNT_BODY
            xl_money(ws_er.cell(row=row, column=4), monto)
            ws_er.cell(row=row, column=2).border = BRD
            ws_er.cell(row=row, column=3).border = BRD
            if row % 2 == 0:
                for c2 in [2, 3, 4]:
                    ws_er.cell(row=row, column=c2).fill = FILL_ZEBRA
            row += 1
        sub_cell = ws_er.cell(row=row, column=3, value=f"Subtotal {sec_nom}")
        sub_cell.font = FNT_BOLD; sub_cell.fill = FILL_TOTAL; sub_cell.border = BRD
        tot_cell = ws_er.cell(row=row, column=4)
        tot_cell.value = f"=SUM(D{first}:D{row-1})" if first < row else 0
        tot_cell.number_format = FMT_RD; tot_cell.font = FNT_BOLD
        tot_cell.fill = FILL_TOTAL; tot_cell.border = BRD
        tot_cell.alignment = Alignment(horizontal="right")
        ws_er.cell(row=row, column=2).fill = FILL_TOTAL; ws_er.cell(row=row, column=2).border = BRD
        er_refs[sec_nom] = f"D{row}"
        row += 2

    ws_er.column_dimensions['B'].width = 12
    ws_er.column_dimensions['C'].width = 48
    ws_er.column_dimensions['D'].width = 20

    # ── Hoja 3: Borrador IR-2 ─────────────────────────────────────────────────
    ws_ir2 = wb.create_sheet("Borrador IR-2")
    xl_header(ws_ir2, f"BORRADOR IR-2 — {empresa.upper()}", "Determinación Ajuste Fiscal Patrimonial (DGII)", ["N° Casilla", "Descripción", "Signo", "Monto (RD$)"])

    ir2_data = [
        (1,  "Total Activos (libros inicio ejercicio)",                      "+", ir2_vals['cas_1']),
        (2,  "Provisiones y Reservas no admitidas (fiscal)",                 "+", 0.0),
        (3,  "Impuesto sobre la Renta Diferido",                             "-", 0.0),
        (4,  "Cuentas por Cobrar no relacionadas con el giro",               "-", 0.0),
        (5,  "Costo del Terreno (Valor en Libros)",                          "-", 0.0),
        (6,  "Costo de las Acciones (Valor en Libros)",                      "-", 0.0),
        (7,  "Costo Edificio (Valor en Libros)",                             "-", 0.0),
        (8,  "Costo Construcción en Proceso (Valor en Libros)",              "-", 0.0),
        (9,  "Costo Activos en Construcción (Valor en Libros)",              "-", 0.0),
        (10, "Costo Activos Fijos Categoría 2 (Valor en Libros)",            "-", ir2_vals['cas_39']),
        (11, "Costo Activos Fijos Categoría 3 (Valor en Libros)",            "-", ir2_vals['cas_40']),
        (12, "Costo Activos Cat.2 (Arrendamientos, Valor Libros)",           "-", 0.0),
        (13, "Costo Activos Cat.3 (Arrendamientos, Valor Libros)",           "-", 0.0),
        (14, "Otros Activos (Valor en Libros)",                              "-", 0.0),
        (15, "Costo Fiscal Categoría 1 (Edificios)",                        "+", ir2_vals['cas_38']),
        (16, "Costo Fiscal Categoría 2 (Maquinaria y Equipo)",              "+", ir2_vals['cas_39']),
        (17, "Costo Fiscal Activo en Construcción",                         "+", 0.0),
        (18, "Costo Fiscal Categoría 3 (Muebles y Enseres)",                "+", ir2_vals['cas_40']),
        (19, "Costo Fiscal Cat.2 (Arrendamientos)",                         "+", 0.0),
        (20, "Costo Fiscal Cat.3 (Arrendamientos)",                         "+", 0.0),
        (21, "Costo Fiscal Terrenos",                                        "+", 0.0),
        (22, "Costo Fiscal de las Acciones",                                "+", 0.0),
        (23, "Reevaluación de Activos",                                     "-", 0.0),
        (24, "Mejoras en Propiedades Arrendadas",                           "-", 0.0),
        (25, "Costo Fiscal de Otros Activos",                               "+", 0.0),
        (26, "SALDO ACTIVOS FISCALES",                                      "=", ir2_vals['cas_26']),
        (27, "Total Pasivos (libros inicio ejercicio)",                     "=", ir2_vals['cas_27']),
        (28, "(-) ISR Diferido Pasivo",                                     "-", 0.0),
        (29, "(-) Provisiones y Reservas no admitidas (pasivo)",            "-", 0.0),
        (30, "(-) Otros Pasivos",                                           "-", 0.0),
        (31, "SALDO DE LOS PASIVOS",                                        "=", ir2_vals['cas_31']),
        (32, "PATRIMONIO FISCAL (Casilla 26 - 31)",                         "=", ir2_vals['cas_32']),
        (33, "Saldo Fiscal Activos no Monetarios (Casilla 49)",             "=", ir2_vals['cas_33']),
        (34, "BASE AJUSTE POR INFLACIÓN (Menor de Casillas 32 y 33)",       "=", ir2_vals['cas_34']),
        (35, "Ajuste Fiscal Patrimonial (% ajustado)",                      "=", 0.0),
        (37, "Inventarios (Valor Fiscal)",                                  " ", ir2_vals['cas_37']),
        (38, "Activos Categoría 1 - Edificios (Casilla 15)",                " ", ir2_vals['cas_38']),
        (39, "Activos Categoría 2 - Maquinaria (Casilla 16)",               " ", ir2_vals['cas_39']),
        (40, "Activos Categoría 3 - Muebles (Casilla 18)",                  " ", ir2_vals['cas_40']),
        (49, "TOTAL ACTIVOS NO MONETARIOS",                                 "=", ir2_vals['cas_49']),
    ]

    FILL_TOTAL_IR2 = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    r = 6
    for cas, desc, signo, monto in ir2_data:
        is_total = signo == "="
        ws_ir2.cell(row=r, column=2, value=str(cas)).font = FNT_BOLD if is_total else FNT_BODY
        ws_ir2.cell(row=r, column=3, value=desc).font = Font(name="Calibri", size=10, bold=is_total, color="FFFFFF" if is_total else "000000")
        ws_ir2.cell(row=r, column=4, value=signo).alignment = Alignment(horizontal="center")
        ws_ir2.cell(row=r, column=4).font = FNT_BOLD
        mc = ws_ir2.cell(row=r, column=5, value=monto)
        mc.number_format = FMT_RD; mc.alignment = Alignment(horizontal="right")
        mc.font = Font(name="Calibri", size=10, bold=is_total, color="FFFFFF" if is_total else "000000")
        for c2 in [2, 3, 4, 5]:
            ws_ir2.cell(row=r, column=c2).border = BRD
            if is_total:
                ws_ir2.cell(row=r, column=c2).fill = FILL_TOTAL_IR2
            elif r % 2 == 0:
                ws_ir2.cell(row=r, column=c2).fill = FILL_ZEBRA
        r += 1

    ws_ir2.column_dimensions['B'].width = 12
    ws_ir2.column_dimensions['C'].width = 52
    ws_ir2.column_dimensions['D'].width = 8
    ws_ir2.column_dimensions['E'].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.title("🏛️ TaxTech Auditor RD")
st.sidebar.markdown("---")
st.sidebar.title("Cliente")
empresa = st.sidebar.text_input("Empresa", value="Empresa de Prueba SRL")
rnc     = st.sidebar.text_input("RNC", value="1-31-12345-6")
periodo = st.sidebar.text_input("Período", value="Enero - Diciembre 2025")
anio    = st.sidebar.text_input("Año Fiscal", value="2025")

st.sidebar.markdown("---")
st.sidebar.title("Materialidad (NIA 320)")
tipo_ent = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])
tasa_ref = 0.01 if tipo_ent == "Comercial / Servicios" else 0.005
pct_mp   = st.sidebar.slider("% Materialidad Planificación", 0.5, 3.0, tasa_ref * 100, 0.1) / 100
pct_me   = st.sidebar.slider("% Materialidad Ejecución (MP×)", 50, 75, 75, 5) / 100

# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────
st.title("TaxTech Auditor — Declaraciones Juradas & Estados Financieros")
st.caption("República Dominicana · DGII · TSS · Ciclo Fiscal 2025-2026")

uploaded = st.file_uploader(
    "📂 Cargar Balanza de Comprobación (Excel / CSV)",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed"
)

if uploaded is None:
    st.info("👆 Sube una balanza de comprobación para iniciar el análisis.")
    st.markdown("""
    **Formato esperado — columnas mínimas:**
    | Código | Cuenta/Nombre | Débito | Crédito | Saldo Final |
    |--------|---------------|--------|---------|-------------|
    | 1101   | Caja y Banco  | 500,000| 0       | 500,000     |
    
    > La aplicación detecta variaciones en nombres de columnas automáticamente.
    """)
    st.stop()

df_bal = procesar_balanza(uploaded)
if df_bal.empty:
    st.stop()

df_bal = analizar_balanza(df_bal)

# KPIs de Materialidad
t_ingresos = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
t_activos  = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
t_costos   = abs(df_bal[df_bal['codigo'].str.startswith('5', na=False)]['saldo_final'].sum())
t_gastos   = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False)]['saldo_final'].sum())
t_pasivos  = abs(df_bal[df_bal['codigo'].str.startswith('2', na=False)]['saldo_final'].sum())
t_patrim   = abs(df_bal[df_bal['codigo'].str.startswith('3', na=False)]['saldo_final'].sum())
utilidad_bruta = t_ingresos - t_costos
utilidad_neta  = t_ingresos - t_costos - t_gastos
base_calc = t_ingresos if t_ingresos > 0 else t_activos
mp = base_calc * pct_mp; me = mp * pct_me

st.markdown(f"### 📌 {empresa} — {periodo}")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Ingresos",        f"RD$ {t_ingresos:,.0f}")
c2.metric("Costo de Ventas", f"RD$ {t_costos:,.0f}")
c3.metric("Utilidad Bruta",  f"RD$ {utilidad_bruta:,.0f}", delta=f"{(utilidad_bruta/t_ingresos*100 if t_ingresos else 0):.1f}%")
c4.metric("Utilidad Neta",   f"RD$ {utilidad_neta:,.0f}", delta=f"{(utilidad_neta/t_ingresos*100 if t_ingresos else 0):.1f}%")
c5.metric("MP (Materialidad)", f"RD$ {mp:,.0f}")
c6.metric("ME (Ejecución)",    f"RD$ {me:,.0f}")

st.markdown("---")

# ─── TABS PRINCIPALES ─────────────────────────────────────────────────────────
tab_bg, tab_er, tab_bal, tab_inconsist, tab_art287, tab_ir2, tab_it1, tab_tss, tab_ir17, tab_consol = st.tabs([
    "📊 Balance General",
    "📈 Estado de Resultados",
    "📋 Balanza",
    "🚨 Inconsistencias",
    "⚖️ Riesgos Art.287",
    "📝 IR-2 Borrador",
    "⚡ IT-1 (ITBIS)",
    "🏢 TSS / IR-3",
    "💸 IR-17",
    "🏛️ Consolidado",
])

bg = generar_balance_general(df_bal)
er = generar_estado_resultados(df_bal)
ir2_vals = calcular_casillas_ir2(df_bal)

# ── TAB: BALANCE GENERAL ──────────────────────────────────────────────────────
with tab_bg:
    st.markdown("### Balance General")
    col_act, col_pas = st.columns(2)

    with col_act:
        st.markdown("#### 🟦 ACTIVO")
        t_ac = sum(m for _, _, m in bg['activo_corriente'])
        t_anc = sum(m for _, _, m in bg['activo_no_corriente'])
        st.markdown("**Activo Corriente**")
        if bg['activo_corriente']:
            df_ac = pd.DataFrame(bg['activo_corriente'], columns=['Cuenta', 'Código', 'Monto'])
            df_ac['Monto'] = df_ac['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
            st.dataframe(df_ac, use_container_width=True, hide_index=True)
        st.metric("Subtotal Activo Corriente", f"RD$ {t_ac:,.2f}")
        st.markdown("**Activo No Corriente**")
        if bg['activo_no_corriente']:
            df_anc = pd.DataFrame(bg['activo_no_corriente'], columns=['Cuenta', 'Código', 'Monto'])
            df_anc['Monto'] = df_anc['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
            st.dataframe(df_anc, use_container_width=True, hide_index=True)
        st.metric("Subtotal Activo No Corriente", f"RD$ {t_anc:,.2f}")
        st.success(f"**TOTAL ACTIVO: RD$ {t_ac + t_anc:,.2f}**")

    with col_pas:
        st.markdown("#### 🟥 PASIVO & PATRIMONIO")
        t_pc  = sum(m for _, _, m in bg['pasivo_corriente'])
        t_pnc = sum(m for _, _, m in bg['pasivo_no_corriente'])
        t_pat = sum(m for _, _, m in bg['patrimonio'])
        st.markdown("**Pasivo Corriente**")
        if bg['pasivo_corriente']:
            df_pc = pd.DataFrame(bg['pasivo_corriente'], columns=['Cuenta', 'Código', 'Monto'])
            df_pc['Monto'] = df_pc['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
            st.dataframe(df_pc, use_container_width=True, hide_index=True)
        st.metric("Subtotal Pasivo Corriente", f"RD$ {t_pc:,.2f}")
        st.markdown("**Pasivo No Corriente**")
        if bg['pasivo_no_corriente']:
            df_pnc = pd.DataFrame(bg['pasivo_no_corriente'], columns=['Cuenta', 'Código', 'Monto'])
            df_pnc['Monto'] = df_pnc['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
            st.dataframe(df_pnc, use_container_width=True, hide_index=True)
        st.metric("Subtotal Pasivo No Corriente", f"RD$ {t_pnc:,.2f}")
        st.markdown("**Patrimonio**")
        if bg['patrimonio']:
            df_pat = pd.DataFrame(bg['patrimonio'], columns=['Cuenta', 'Código', 'Monto'])
            df_pat['Monto'] = df_pat['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
            st.dataframe(df_pat, use_container_width=True, hide_index=True)
        st.metric("Subtotal Patrimonio", f"RD$ {t_pat:,.2f}")
        st.error(f"**TOTAL PASIVO + PATRIMONIO: RD$ {t_pc + t_pnc + t_pat:,.2f}**")

    # Cuadre
    total_activo = t_ac + t_anc
    total_pas_pat = t_pc + t_pnc + t_pat
    diferencia = total_activo - total_pas_pat
    if abs(diferencia) < 1:
        st.success(f"✅ Balanza CUADRADA — Diferencia: RD$ {diferencia:,.2f}")
    else:
        st.warning(f"⚠️ Diferencia de cuadre: RD$ {diferencia:,.2f} — Revisar saldos o clasificación de cuentas.")

    # Exportar
    excel_bytes = exportar_estados_excel(empresa, periodo, bg, er, ir2_vals)
    st.download_button(
        "📥 Descargar Estados Financieros + IR-2 (Excel)",
        data=excel_bytes,
        file_name=f"Estados_Financieros_{empresa.replace(' ', '_')}_{anio}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── TAB: ESTADO DE RESULTADOS ─────────────────────────────────────────────────
with tab_er:
    st.markdown("### Estado de Resultados")

    t_ing  = sum(m for _, _, m in er['ingresos'])
    t_cos  = sum(m for _, _, m in er['costos'])
    t_gop  = sum(m for _, _, m in er['gastos_operacion'])
    t_gfin = sum(m for _, _, m in er['gastos_financieros'])
    t_oing = sum(m for _, _, m in er['otros_ingresos'])
    util_bruta = t_ing - t_cos
    util_oper  = util_bruta - t_gop
    util_neta  = util_oper - t_gfin + t_oing

    col_p1, col_p2 = st.columns([2, 1])
    with col_p1:
        for sec_nom, cuentas in [
            ("INGRESOS", er['ingresos']),
            ("COSTO DE VENTAS / SERVICIOS", er['costos']),
            ("GASTOS DE OPERACIÓN", er['gastos_operacion']),
            ("GASTOS FINANCIEROS", er['gastos_financieros']),
            ("OTROS INGRESOS", er['otros_ingresos']),
        ]:
            if cuentas:
                st.markdown(f"**{sec_nom}**")
                df_sec = pd.DataFrame(cuentas, columns=['Cuenta', 'Código', 'Monto'])
                df_sec['Monto'] = df_sec['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
                st.dataframe(df_sec, use_container_width=True, hide_index=True)

    with col_p2:
        st.markdown("#### 📊 Resumen")
        st.metric("Ingresos Totales",  f"RD$ {t_ing:,.2f}")
        st.metric("(-) Costo Ventas",  f"RD$ {t_cos:,.2f}")
        margen_b = (util_bruta / t_ing * 100) if t_ing else 0
        st.metric("= Utilidad Bruta",  f"RD$ {util_bruta:,.2f}", delta=f"Margen {margen_b:.1f}%")
        st.metric("(-) Gastos Operación", f"RD$ {t_gop:,.2f}")
        margen_o = (util_oper / t_ing * 100) if t_ing else 0
        st.metric("= Utilidad Operacional", f"RD$ {util_oper:,.2f}", delta=f"Margen {margen_o:.1f}%")
        st.metric("(-) Gastos Financieros", f"RD$ {t_gfin:,.2f}")
        st.metric("(+) Otros Ingresos", f"RD$ {t_oing:,.2f}")
        margen_n = (util_neta / t_ing * 100) if t_ing else 0
        if util_neta >= 0:
            st.success(f"**UTILIDAD NETA: RD$ {util_neta:,.2f}** ({margen_n:.1f}%)")
        else:
            st.error(f"**PÉRDIDA NETA: RD$ {abs(util_neta):,.2f}** ({margen_n:.1f}%)")

        # ISR estimado
        if util_neta > 0:
            isr_est = util_neta * 0.27
            st.warning(f"ISR Estimado (27%): RD$ {isr_est:,.2f}")

# ── TAB: BALANZA CRUDA ────────────────────────────────────────────────────────
with tab_bal:
    st.dataframe(df_bal, use_container_width=True)
    buf_bal = io.BytesIO()
    with pd.ExcelWriter(buf_bal, engine='openpyxl') as w:
        df_bal.to_excel(w, index=False, sheet_name="Balanza Analizada")
    st.download_button("📥 Descargar Balanza Analizada", data=buf_bal.getvalue(), file_name=f"Balanza_{empresa.replace(' ', '_')}.xlsx")

# ── TAB: INCONSISTENCIAS ──────────────────────────────────────────────────────
with tab_inconsist:
    df_err = df_bal[~df_bal['validacion_naturaleza'].str.startswith('✅')]
    if df_err.empty:
        st.success("✅ Sin inconsistencias en naturaleza de saldos.")
    else:
        st.error(f"Se detectaron {len(df_err)} cuentas con inconsistencias.")
        st.dataframe(df_err[['codigo', 'cuenta', 'debito', 'credito', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)

# ── TAB: ART. 287 ─────────────────────────────────────────────────────────────
with tab_art287:
    df_fisc = df_bal[df_bal['alerta_fiscal'] != ""]
    if df_fisc.empty:
        st.success("✅ Sin alertas fiscales Art. 287 detectadas.")
    else:
        st.warning(f"{len(df_fisc)} cuentas con exposición fiscal identificada.")
        st.dataframe(df_fisc[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal']], use_container_width=True)

# ── TAB: IR-2 BORRADOR ────────────────────────────────────────────────────────
with tab_ir2:
    st.markdown("### 📝 IR-2 — Determinación Ajuste Fiscal Patrimonial")
    st.caption("Valores calculados automáticamente desde la balanza. Revisa y ajusta manualmente donde aplique.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### Sección 1 — Saldo Activos Fiscales")
        ir2_display = [
            ("1", "Total Activos (libros inicio ejercicio)", "+", ir2_vals['cas_1']),
            ("10", "Costo Activos Cat.2 (libros)", "-", ir2_vals['cas_39']),
            ("11", "Costo Activos Cat.3 (libros)", "-", ir2_vals['cas_40']),
            ("15", "Costo Fiscal Categoría 1 (Edificios)", "+", ir2_vals['cas_38']),
            ("16", "Costo Fiscal Categoría 2 (Maquinaria)", "+", ir2_vals['cas_39']),
            ("18", "Costo Fiscal Categoría 3 (Muebles)", "+", ir2_vals['cas_40']),
            ("26", "SALDO ACTIVOS FISCALES", "=", ir2_vals['cas_26']),
        ]
        df_ir2a = pd.DataFrame(ir2_display, columns=["Casilla", "Concepto", "Signo", "Monto"])
        df_ir2a['Monto'] = df_ir2a['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
        st.dataframe(df_ir2a, use_container_width=True, hide_index=True)

    with col_b:
        st.markdown("#### Sección 1 — Saldo Pasivos y Patrimonio Fiscal")
        ir2_display_b = [
            ("27", "Total Pasivos (libros inicio ejercicio)", "=", ir2_vals['cas_27']),
            ("31", "SALDO DE LOS PASIVOS", "=", ir2_vals['cas_31']),
            ("32", "PATRIMONIO FISCAL (Cas.26 - Cas.31)", "=", ir2_vals['cas_32']),
            ("33", "Saldo Activos No Monetarios (Cas.49)", "=", ir2_vals['cas_33']),
            ("34", "BASE AJUSTE (Menor: 32 vs 33)", "=", ir2_vals['cas_34']),
            ("35", "Ajuste Fiscal Patrimonial (%)", "=", 0.0),
        ]
        df_ir2b = pd.DataFrame(ir2_display_b, columns=["Casilla", "Concepto", "Signo", "Monto"])
        df_ir2b['Monto'] = df_ir2b['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
        st.dataframe(df_ir2b, use_container_width=True, hide_index=True)

    st.markdown("#### Sección 2 — Distribución Activos No Monetarios")
    ir2_display_c = [
        ("37", "Inventarios", ir2_vals['cas_37']),
        ("38", "Activos Categoría 1 — Edificios (de Cas.15)", ir2_vals['cas_38']),
        ("39", "Activos Categoría 2 — Maquinaria (de Cas.16)", ir2_vals['cas_39']),
        ("40", "Activos Categoría 3 — Muebles (de Cas.18)", ir2_vals['cas_40']),
        ("49", "TOTAL ACTIVOS NO MONETARIOS", ir2_vals['cas_49']),
    ]
    df_ir2c = pd.DataFrame(ir2_display_c, columns=["Casilla", "Concepto", "Monto"])
    df_ir2c['Monto'] = df_ir2c['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
    st.dataframe(df_ir2c, use_container_width=True, hide_index=True)

    st.info("💡 Descarga el Excel completo desde la tab **Balance General** — incluye la hoja 'Borrador IR-2' con todas las casillas.")

    # Subir IR-2 oficial para cruzar
    st.markdown("---")
    st.markdown("#### 📂 Cargar IR-2 Oficial (DGII) para Cruce")
    f_ir2 = st.file_uploader("Sube el IR-2-XXXX.xls descargado del portal DGII", type=["xls", "xlsx"], key="ir2_oficial")
    if f_ir2:
        try:
            engine = 'xlrd' if f_ir2.name.endswith('.xls') else 'openpyxl'
            df_ir2_raw = pd.read_excel(f_ir2, engine=engine, header=None)
            st.success(f"✅ Archivo cargado: {df_ir2_raw.shape[0]} filas × {df_ir2_raw.shape[1]} columnas")

            # Extraer casillas con montos
            casillas_encontradas = []
            for idx, row in df_ir2_raw.iterrows():
                for col_idx, val in enumerate(row):
                    val_str = str(val)
                    if val_str.replace('.', '').replace('-', '').replace(',', '').strip().isdigit():
                        num = float(str(val).replace(',', ''))
                        if num != 0 and abs(num) > 100:
                            concepto = ""
                            for c2 in range(max(0, col_idx - 5), col_idx):
                                cand = str(df_ir2_raw.iat[idx, c2])
                                if len(cand) > 5 and cand != 'nan':
                                    concepto = cand; break
                            if concepto:
                                casillas_encontradas.append({'Descripción': concepto[:70], 'Monto DGII': f"RD$ {num:,.2f}"})

            if casillas_encontradas:
                st.dataframe(pd.DataFrame(casillas_encontradas[:30]), use_container_width=True, hide_index=True)
            else:
                st.info("Archivo vacío (sin montos ingresados). Usa el borrador calculado arriba para llenar las casillas.")
        except Exception as e:
            st.error(f"Error al leer el IR-2: {e}")

# ── TAB: IT-1 / ITBIS ─────────────────────────────────────────────────────────
with tab_it1:
    st.markdown("### ⚡ Liquidación ITBIS — IT-1")

    ing_grav = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
    itbis_gen = ing_grav * TASA_ITBIS
    compras   = abs(df_bal[df_bal['codigo'].str.startswith(('5', '6'), na=False) &
                            ~df_bal['cuenta'].str.lower().str.contains('personal|sueldo|salario|tss|infotep', na=False)]['saldo_final'].sum())
    itbis_sop = compras * TASA_ITBIS
    b_hon     = abs(df_bal[df_bal['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
    itbis_r100 = b_hon * TASA_ITBIS
    b_serv    = abs(df_bal[df_bal['cuenta'].str.lower().str.contains('servicio tecnico|consultoria|reparacion', na=False)]['saldo_final'].sum())
    itbis_r30  = b_serv * TASA_ITBIS * 0.30
    b_tarj    = abs(df_bal[df_bal['cuenta'].str.lower().str.contains('retencion tarjeta|adquirente|cardnet|azul', na=False)]['saldo_final'].sum())
    itbis_tarj = b_tarj if b_tarj > 0 else ing_grav * 0.60 * 0.02
    neto_itbis = itbis_gen - itbis_sop - itbis_r100 - itbis_r30 - itbis_tarj

    col_606, col_607 = st.columns(2)
    with col_606:
        f606 = st.file_uploader("Formato 606 (Compras)", type=["xlsx", "csv"], key="f606")
        if f606:
            try:
                df606 = pd.read_excel(f606) if f606.name.endswith('.xlsx') else pd.read_csv(f606)
                df606.columns = [str(c).lower().strip() for c in df606.columns]
                col_it = [c for c in df606.columns if 'itbis' in c and ('adelantado' in c or 'pagado' in c or 'facturado' in c)]
                if col_it:
                    itbis_sop = pd.to_numeric(df606[col_it[0]], errors='coerce').fillna(0).sum()
                    st.info(f"ITBIS soportado actualizado desde 606: RD$ {itbis_sop:,.2f}")
                st.success(f"606 cargado: {len(df606)} registros")
            except Exception as e:
                st.error(str(e))
    with col_607:
        f607 = st.file_uploader("Formato 607 (Ventas)", type=["xlsx", "csv"], key="f607")
        if f607:
            try:
                df607 = pd.read_excel(f607) if f607.name.endswith('.xlsx') else pd.read_csv(f607)
                df607.columns = [str(c).lower().strip() for c in df607.columns]
                col_it607 = [c for c in df607.columns if 'itbis' in c and 'facturado' in c]
                if col_it607:
                    itbis_gen = pd.to_numeric(df607[col_it607[0]], errors='coerce').fillna(0).sum()
                    st.info(f"ITBIS generado actualizado desde 607: RD$ {itbis_gen:,.2f}")
                st.success(f"607 cargado: {len(df607)} registros")
            except Exception as e:
                st.error(str(e))

    neto_itbis = itbis_gen - itbis_sop - itbis_r100 - itbis_r30 - itbis_tarj
    st.markdown("---")
    if neto_itbis > 0:
        st.error(f"🚨 ITBIS NETO A PAGAR: RD$ {neto_itbis:,.2f}")
    else:
        st.success(f"✅ SALDO A FAVOR (COMPENSABLE): RD$ {abs(neto_itbis):,.2f}")

    df_it1 = pd.DataFrame([
        ("Línea 1", "Ingresos gravados (Base 607/Balanza)", ing_grav),
        ("Línea 2", "ITBIS Bruto Generado (18%)",          itbis_gen),
        ("Línea 3", "(-) Adelanto ITBIS Compras (606)",    itbis_sop),
        ("Línea 4", "(-) ITBIS Ret. Personas Físicas",     itbis_r100),
        ("Línea 5", "(-) ITBIS Ret. Personas Jurídicas",   itbis_r30),
        ("Línea 6", "(-) Ret. Tarjetas Norma 08-04",       itbis_tarj),
        ("TOTAL",   "IMPUESTO NETO / SALDO A FAVOR",       neto_itbis),
    ], columns=["Línea", "Concepto", "RD$"])
    df_it1['RD$'] = df_it1['RD$'].apply(lambda x: f"RD$ {x:,.2f}")
    st.dataframe(df_it1, use_container_width=True, hide_index=True)

# ── TAB: TSS / IR-3 ──────────────────────────────────────────────────────────
with tab_tss:
    st.markdown("### 🏢 Nómina y TSS — Liquidación IR-3")

    nom = abs(df_bal[df_bal['cuenta'].str.lower().str.contains('personal|sueldo|salario', na=False)]['saldo_final'].sum())
    percap = abs(df_bal[df_bal['cuenta'].str.lower().str.contains('percapita|per capita|dependiente', na=False)]['saldo_final'].sum())
    num_dep = round(percap / COSTO_PERCAPITA_2026) if percap > 0 else 0

    st.markdown(f"**Base nómina detectada:** RD$ {nom:,.2f} | **Percápita registrada:** RD$ {percap:,.2f} ({num_dep} dependientes estimados)")

    df_tss = pd.DataFrame([
        ("Patronal", "SFS Patronal (7.09%)",    TASA_SFS_PAT,  nom * TASA_SFS_PAT),
        ("Patronal", "AFP Patronal (7.10%)",    TASA_AFP_PAT,  nom * TASA_AFP_PAT),
        ("Patronal", "SRL Patronal (1.20%)",    TASA_SRL,      nom * TASA_SRL),
        ("Patronal", "INFOTEP (1.00%)",         TASA_INFOTEP,  nom * TASA_INFOTEP),
        ("Empleado", "SFS Empleado (3.04%)",    TASA_SFS_EMP,  nom * TASA_SFS_EMP),
        ("Empleado", "AFP Empleado (2.87%)",    TASA_AFP_EMP,  nom * TASA_AFP_EMP),
        ("Empleado", "Percápita Dependientes",  0.0,           percap),
    ], columns=["Tipo", "Concepto", "Tasa", "Monto"])
    df_tss['Tasa'] = df_tss['Tasa'].apply(lambda x: f"{x:.2%}" if x > 0 else "—")
    df_tss['Monto'] = df_tss['Monto'].apply(lambda x: f"RD$ {x:,.2f}")
    st.dataframe(df_tss, use_container_width=True, hide_index=True)

    total_pat = nom * (TASA_SFS_PAT + TASA_AFP_PAT + TASA_SRL + TASA_INFOTEP)
    total_emp = nom * (TASA_SFS_EMP + TASA_AFP_EMP) + percap
    st.metric("Total Costo Patronal", f"RD$ {total_pat:,.2f}")
    st.metric("Total Retenciones Empleado", f"RD$ {total_emp:,.2f}")
    st.error(f"**FACTURA TSS TOTAL: RD$ {total_pat + total_emp:,.2f}**")

    f_tss = st.file_uploader("Cargar nómina detallada por empleado (Excel/CSV)", type=["xlsx", "csv"], key="nomina")
    if f_tss:
        df_nom = pd.read_excel(f_tss) if f_tss.name.endswith('.xlsx') else pd.read_csv(f_tss)
        st.success(f"Nómina cargada: {len(df_nom)} empleados")
        st.dataframe(df_nom.head(20), use_container_width=True)

# ── TAB: IR-17 ────────────────────────────────────────────────────────────────
with tab_ir17:
    st.markdown("### 💸 IR-17 — Otras Retenciones")

    def get(kw): return abs(df_bal[df_bal['cuenta'].str.lower().str.contains(kw, na=False)]['saldo_final'].sum())

    b_hon2  = get('honorario')
    b_rep   = get('reparacion|mantenimiento')
    b_veh   = get('vehiculo personal|combustible empleado')
    b_rent  = get('renta personal|alquiler personal|vivienda')
    b_esp   = get('especie')
    b_spain = get('españa|espana')
    b_can   = get('canada|canadá')
    b_ext   = get('exterior|remesa|extranjero')

    datos_ir17 = [
        ("1",  "Honorarios Personas Físicas",          "10%",  b_hon2,  b_hon2 * 0.10),
        ("2",  "Servicios Técnicos y Reparaciones",    "2%",   b_rep,   b_rep  * 0.02),
        ("8",  "Retrib. Compl. — Vehículos/Combustible","27%", b_veh,   b_veh  * 0.27),
        ("9",  "Retrib. Compl. — Alquileres/Renta",    "27%",  b_rent,  b_rent * 0.27),
        ("10", "Retrib. Compl. — Otros en Especie",    "27%",  b_esp,   b_esp  * 0.27),
        ("15", "Remesas Exterior — CDI España",        "10%",  b_spain, b_spain* 0.10),
        ("16", "Remesas Exterior — CDI Canadá",        "18%",  b_can,   b_can  * 0.18),
        ("17", "Remesas Exterior — Otros",             "27%",  b_ext,   b_ext  * 0.27),
    ]
    total_ir17 = sum(r[4] for r in datos_ir17)
    df_ir17 = pd.DataFrame(datos_ir17, columns=["Casilla", "Concepto", "Tasa", "Base Imponible", "Impuesto"])
    df_ir17['Base Imponible'] = df_ir17['Base Imponible'].apply(lambda x: f"RD$ {x:,.2f}")
    df_ir17['Impuesto']        = df_ir17['Impuesto'].apply(lambda x: f"RD$ {x:,.2f}")
    st.dataframe(df_ir17, use_container_width=True, hide_index=True)
    st.error(f"**TOTAL IR-17 A PAGAR: RD$ {total_ir17:,.2f}**")

# ── TAB: CONSOLIDADO ─────────────────────────────────────────────────────────
with tab_consol:
    st.markdown("### 🏛️ Consolidado Fiscal General")

    neto_itbis_c = itbis_gen - itbis_sop - itbis_r100 - itbis_r30 - itbis_tarj
    total_tss_c  = nom * (TASA_SFS_PAT + TASA_AFP_PAT + TASA_SRL + TASA_INFOTEP + TASA_SFS_EMP + TASA_AFP_EMP) + percap
    total_ir17_c = sum(b_hon2*0.10, b_rep*0.02, b_veh*0.27, b_rent*0.27, b_esp*0.27, b_spain*0.10, b_can*0.18, b_ext*0.27) \
                   if False else (b_hon2*0.10 + b_rep*0.02 + b_veh*0.27 + b_rent*0.27 + b_esp*0.27 + b_spain*0.10 + b_can*0.18 + b_ext*0.27)
    isr_est = max(0, utilidad_neta) * 0.27

    df_consol = pd.DataFrame([
        ("IT-1",  "ITBIS Mensual",                    f"RD$ {abs(neto_itbis_c):,.2f}", "A Pagar" if neto_itbis_c > 0 else "Saldo Favor", "✅" if neto_itbis_c <= 0 else "⚠️"),
        ("IR-3",  "TSS / Seguridad Social",           f"RD$ {total_tss_c:,.2f}",       "Estimado",     "ℹ️"),
        ("IR-17", "Otras Retenciones",                f"RD$ {total_ir17_c:,.2f}",      "A Pagar",      "⚠️" if total_ir17_c > 0 else "✅"),
        ("IR-2",  "Ajuste Patrimonial",               f"RD$ {ir2_vals['cas_34']:,.2f}","Base",         "ℹ️"),
        ("ISR",   "Impuesto Renta Estimado (27%)",    f"RD$ {isr_est:,.2f}",           "Proyección",   "⚠️" if isr_est > 0 else "✅"),
    ], columns=["Formulario", "Concepto", "Monto", "Estado", ""])
    st.dataframe(df_consol, use_container_width=True, hide_index=True)

    gran_total = abs(neto_itbis_c) + total_ir17_c + isr_est
    st.error(f"**OBLIGACIONES FISCALES TOTALES ESTIMADAS: RD$ {gran_total:,.2f}**")

    n_alertas = len(df_bal[~df_bal['validacion_naturaleza'].str.startswith('✅')])
    n_riesgos = len(df_bal[df_bal['alerta_fiscal'] != ""])
    st.markdown("---")
    st.markdown("#### Resumen de Riesgos de Auditoría")
    cr, cr2, cr3 = st.columns(3)
    cr.metric("Cuentas con inconsistencias", n_alertas)
    cr2.metric("Cuentas con riesgo Art.287", n_riesgos)
    cr3.metric("Ratio Deuda/Activo", f"{(t_pasivos/t_activos*100 if t_activos else 0):.1f}%")
