import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import plotly.express as px

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
    .stTabs [data-baseweb="tab"] { font-size: 0.82rem; padding: 6px 14px; }
    .stTabs [aria-selected="true"] { border-bottom: 2px solid #1e3a5f; font-weight: 700; }
    h1 { color: #1e3a5f !important; }
    h3 { color: #1e3a5f !important; font-size: 1rem !important; }
    
    .tabla-contable { width: 100%; border-collapse: collapse; font-family: 'Calibri', sans-serif; font-size: 0.9rem; margin-bottom: 1rem; }
    .tabla-contable th { border-bottom: 1px solid #000; padding: 8px; text-align: right; font-weight: bold; }
    .tabla-contable th:first-child { text-align: left; }
    .tabla-contable td { padding: 6px 8px; text-align: right; }
    .tabla-contable td:first-child { text-align: left; }
    .tabla-contable .seccion { font-weight: bold; text-align: left; padding-top: 15px; }
    .tabla-contable .total td { border-top: 1px solid #000; border-bottom: 3px double #000; font-weight: bold; }
    .tabla-contable .subtotal td { border-top: 1px solid #ccc; font-weight: bold; }
    .tabla-contable .titulo-anio { text-align: center; font-weight: bold; font-size: 0.95rem; background-color: #f1f5f9; padding: 8px !important; border-bottom: 2px solid #000;}
    
    /* Estilos Formulario Oficial IR-2 / Activos */
    .tabla-ir2 { width: 100%; border-collapse: collapse; font-family: 'Calibri', sans-serif; font-size: 0.85rem; margin-bottom: 2rem; border: 1px solid #cbd5e1; }
    .tabla-ir2 th { background-color: #1e3a5f; color: white; padding: 10px; text-align: center; font-weight: bold; border: 1px solid #cbd5e1; }
    .tabla-ir2 .header-seccion { background-color: #f1f5f9; font-weight: bold; color: #0f1923; text-align: left; padding: 6px; border: 1px solid #cbd5e1; font-size: 0.9rem; }
    .tabla-ir2 td { padding: 6px; border: 1px solid #cbd5e1; }
    .tabla-ir2 .col-num { width: 6%; text-align: center; font-weight: bold; background-color: #f8fafc; color: #475569; }
    .tabla-ir2 .col-desc { width: 64%; }
    .tabla-ir2 .col-monto { width: 30%; text-align: right; font-family: monospace; font-size:0.95rem;}
    .tabla-ir2 .fila-total td { font-weight: bold; background-color: #e2e8f0; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# PARÁMETROS FISCALES Y CLASIFICACIÓN HÍBRIDA
# ──────────────────────────────────────────────────────────────────────────────
NATURALEZAS = {'1': 'Debito', '2': 'Credito', '3': 'Credito', '4': 'Credito', '5': 'Debito', '6': 'Debito'}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Art. 287 CTRD: Validar NCF válido y medios de pago para crédito ITBIS.',
    'representacion': 'Art. 287 CTRD: Razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Art. 318 / Reg. 139-98: Retribuciones en especie. Validar ISR sustitutivo.',
    'gasto de personal': 'Art. 287 CTRD: Cruce obligatorio con IR-4 (TSS) para admitir deducción.',
    'honorario': 'Art. 309 CTRD: Retención 10% personas físicas / 2% entre jurídicas.',
}

TASA_ITBIS = 0.18
TASA_SFS_PAT = 0.0709; TASA_AFP_PAT = 0.0710
TASA_SRL = 0.0120; TASA_INFOTEP = 0.0100
TASA_SFS_EMP = 0.0304; TASA_AFP_EMP = 0.0287

def fmt_c(val):
    if pd.isna(val) or round(val, 2) == 0: return "-"
    return f"({abs(val):,.0f})" if val < 0 else f"{val:,.0f}"

def es_activo_no_corriente(cod, nombre):
    c = str(cod); n = str(nombre).lower()
    if c.startswith(('15', '16', '17', '18', '19')): return True
    if any(x in n for x in ['fijo', 'propiedad', 'planta', 'equipo', 'depreciacion', 'edificio', 'terreno', 'vehiculo', 'software', 'intangible']): return True
    return False

def es_pasivo_no_corriente(cod, nombre):
    c = str(cod); n = str(nombre).lower()
    if c.startswith(('22', '23', '24')): return True
    if 'largo plazo' in n: return True
    return False

# ──────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO DE BALANZA
# ──────────────────────────────────────────────────────────────────────────────
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.lower().endswith(('.xlsx', '.xls')): df_raw = pd.read_excel(file, header=None)
        else: df_raw = pd.read_csv(file, header=None)
        
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(x).lower() for x in row.values])
            if ('código' in row_str or 'codigo' in row_str) and ('nombre' in row_str or 'cuenta' in row_str):
                header_idx = idx; break
        
        column_names = df_raw.iloc[header_idx].astype(str).str.lower().str.strip()
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        
        idx_codigo, idx_cuenta = -1, -1
        indices_debe, indices_haber, indices_balance = [], [], []
        
        for i, col in enumerate(column_names):
            if any(x in col for x in ['código', 'codigo', 'cuenta no']) and idx_codigo == -1: idx_codigo = i
            elif any(x in col for x in ['nombre', 'descripción', 'cuenta']) and 'codigo' not in col and idx_cuenta == -1: idx_cuenta = i
            elif any(x in col for x in ['débito', 'debito', 'debe', 'cargos']): indices_debe.append(i)
            elif any(x in col for x in ['crédito', 'credito', 'haber', 'abonos']): indices_haber.append(i)
            elif any(x in col for x in ['saldo', 'balance', 'final', 'monto']): indices_balance.append(i)
        
        if idx_codigo == -1 or idx_cuenta == -1: return pd.DataFrame()
        
        col_dict = {'codigo': df.iloc[:, idx_codigo], 'cuenta': df.iloc[:, idx_cuenta]}
        if indices_debe: col_dict['debito'] = df.iloc[:, indices_debe[-1]]
        if indices_haber: col_dict['credito'] = df.iloc[:, indices_haber[-1]]
        if indices_balance: col_dict['saldo_final'] = df.iloc[:, indices_balance[-1]]
        
        df_clean = pd.DataFrame(col_dict)
        if 'saldo_final' not in df_clean.columns and 'debito' in df_clean.columns:
            df_clean['saldo_final'] = pd.to_numeric(df_clean['debito'], errors='coerce').fillna(0) - pd.to_numeric(df_clean['credito'], errors='coerce').fillna(0)
        
        df_clean['codigo'] = df_clean['codigo'].fillna('').astype(str).str.strip().apply(lambda x: x.split('.')[0] if '.' in x else x)
        df_clean['cuenta'] = df_clean['cuenta'].fillna('').astype(str).str.strip()
        df_clean = df_clean[(df_clean['codigo'] != '') & (~df_clean['codigo'].str.lower().str.contains('total|suma', na=False))]
        
        for col in ['debito', 'credito', 'saldo_final']:
            if col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col].astype(str).str.replace(',', '').replace(r'^-*$', '0', regex=True), errors='coerce').fillna(0.0)
                
        return df_clean.reset_index(drop=True)
    except: return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat, alertas_fisc = [], []
    for _, row in df.iterrows():
        cod, nom, saldo = str(row['codigo']), str(row['cuenta']).lower(), row['saldo_final']
        nat_esp = NATURALEZAS.get(cod[0] if cod else '', None)
        if nat_esp == 'Debito' and saldo < -1: alertas_nat.append("⚠️ Saldo crédito (nat. débito)")
        elif nat_esp == 'Credito' and saldo > 1: alertas_nat.append("⚠️ Saldo débito (nat. crédito)")
        else: alertas_nat.append("✅ Correcto")
        alertas_fisc.append(next((msg for p, msg in PALABRAS_CRITICAS_ART287.items() if p in nom), ""))
    df['validacion_naturaleza'] = alertas_nat; df['alerta_fiscal'] = alertas_fisc
    return df

def procesar_comparativo(df_act: pd.DataFrame, df_ant: pd.DataFrame) -> pd.DataFrame:
    df_comp = pd.merge(df_act[['codigo', 'cuenta', 'saldo_final']], df_ant[['codigo', 'saldo_final']], on='codigo', how='outer', suffixes=('_Y2', '_Y1')).fillna(0.0)
    df_comp.loc[df_comp['cuenta'] == 0.0, 'cuenta'] = "Cuenta Histórica"
    df_comp['variacion_abs'] = df_comp['saldo_final_Y2'] - df_comp['saldo_final_Y1']
    return df_comp

def calcular_casillas_ir2(df: pd.DataFrame) -> dict:
    def suma(prefijos):
        mask = df['codigo'].apply(lambda x: any(str(x).startswith(p) for p in prefijos))
        return abs(df.loc[mask, 'saldo_final'].sum())
    total_pasivos = suma(['2'])
    inventario = suma(['13', '130', '131', '132', '133'])
    af1 = suma(['152', '153']); af2 = suma(['154', '155']); af3 = suma(['156', '157', '158'])
    otros_act = suma(['14', '16', '17', '18', '19'])
    saldo_act_fiscal = af1 + af2 + af3 + inventario + otros_act
    patrimonio_fisc = saldo_act_fiscal - total_pasivos
    total_no_monet = af1 + af2 + af3 + inventario
    return {'cas_34': min(max(patrimonio_fisc, 0), total_no_monet)}

# ──────────────────────────────────────────────────────────────────────────────
# EXPORTADOR EXCEL CORPORATIVO AVANZADO (OPENPYXL)
# ──────────────────────────────────────────────────────────────────────────────
def exportar_reporte_corporativo(empresa, periodo, anio_act, df_comp):
    try:
        wb = openpyxl.Workbook()
        FNT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F497D")
        FNT_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        FNT_B = Font(name="Calibri", size=11, bold=True)
        FNT_N = Font(name="Calibri", size=11)
        FILL_HDR = PatternFill("solid", fgColor="1F497D")
        FILL_SUB = PatternFill("solid", fgColor="F1F5F9")
        FILL_TOT = PatternFill("solid", fgColor="E2E8F0")
        THIN = Side(border_style="thin", color="CBD5E1")
        B_BTM = Border(bottom=THIN)
        B_DBL = Border(top=THIN, bottom=Side(border_style="double", color="000000"))
        FMT_ACC = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        
        def format_row(ws, r, style_type='normal'):
            for cell in ws[r]:
                cell.font = FNT_B if style_type in ['sub', 'sec', 'tot'] else FNT_N
                if style_type == 'sec': cell.fill = FILL_SUB; cell.border = B_BTM
                if style_type == 'tot': cell.fill = FILL_TOT; cell.border = B_DBL
                if isinstance(cell.value, (int, float)): cell.number_format = FMT_ACC
                
        def create_sheet_header(ws, title):
            ws.column_dimensions['A'].width = 45; ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 20; ws.column_dimensions['D'].width = 20
            ws["A1"] = f"{empresa.upper()}"; ws["A2"] = title; ws["A3"] = f"Años fiscales: {anio_act} y {int(anio_act)-1}"
            for r in range(1, 4): ws[f"A{r}"].font = FNT_TITLE
            for i, h in enumerate(["Cuenta", "Nota", f"Año {anio_act}", f"Año {int(anio_act)-1}"], 1):
                c = ws.cell(row=5, column=i, value=h)
                c.font = FNT_HDR; c.fill = FILL_HDR; c.alignment = Alignment(horizontal="center")
            return 6

        # 1. BALANCE GENERAL
        ws_bg = wb.active; ws_bg.title = "Balance General"
        r = create_sheet_header(ws_bg, "ESTADO DE SITUACIÓN FINANCIERA")
        def process_section(ws, r_idx, title, prefix, is_current, is_asset):
            ws.cell(row=r_idx, column=1, value=title); format_row(ws, r_idx, 'sec'); r_idx += 1
            tot_y2, tot_y1 = 0, 0
            for _, row in df_comp[df_comp['codigo'].str.startswith(prefix, na=False)].iterrows():
                if (is_current and not (es_activo_no_corriente(row['codigo'], row['cuenta']) if is_asset else es_pasivo_no_corriente(row['codigo'], row['cuenta']))) or \
                   (not is_current and (es_activo_no_corriente(row['codigo'], row['cuenta']) if is_asset else es_pasivo_no_corriente(row['codigo'], row['cuenta']))):
                    v2, v1 = abs(row['saldo_final_Y2']), abs(row['saldo_final_Y1'])
                    if is_asset and 'acum' in str(row['cuenta']).lower(): v2, v1 = -v2, -v1
                    if v2 == 0 and v1 == 0: continue
                    tot_y2 += v2; tot_y1 += v1
                    ws.cell(row=r_idx, column=1, value=row['cuenta'].title())
                    ws.cell(row=r_idx, column=3, value=v2); ws.cell(row=r_idx, column=4, value=v1)
                    format_row(ws, r_idx, 'normal'); r_idx += 1
            ws.cell(row=r_idx, column=1, value=f"Total {title.lower().replace(':','')}"); ws.cell(row=r_idx, column=3, value=tot_y2); ws.cell(row=r_idx, column=4, value=tot_y1)
            format_row(ws, r_idx, 'sub'); r_idx += 1
            return r_idx, tot_y2, tot_y1
            
        r, ac_y2, ac_y1 = process_section(ws_bg, r, "Activos corrientes:", '1', True, True)
        r, anc_y2, anc_y1 = process_section(ws_bg, r, "Activos no corrientes:", '1', False, True)
        ws_bg.cell(row=r, column=1, value="Total Activos"); ws_bg.cell(row=r, column=3, value=ac_y2+anc_y2); ws_bg.cell(row=r, column=4, value=ac_y1+anc_y1)
        format_row(ws_bg, r, 'tot'); r += 2
        
        r, pc_y2, pc_y1 = process_section(ws_bg, r, "Pasivos corrientes:", '2', True, False)
        r, pnc_y2, pnc_y1 = process_section(ws_bg, r, "Pasivos no corrientes:", '2', False, False)
        
        ws_bg.cell(row=r, column=1, value="Patrimonio:"); format_row(ws_bg, r, 'sec'); r += 1
        pat_y2, pat_y1 = 0, 0
        for _, row in df_comp[df_comp['codigo'].str.startswith('3', na=False)].iterrows():
            v2, v1 = abs(row['saldo_final_Y2']), abs(row['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            pat_y2 += v2; pat_y1 += v1
            ws_bg.cell(row=r, column=1, value=row['cuenta'].title()); ws_bg.cell(row=r, column=3, value=v2); ws_bg.cell(row=r, column=4, value=v1)
            format_row(ws_bg, r, 'normal'); r += 1
        ws_bg.cell(row=r, column=1, value="Total Patrimonio"); ws_bg.cell(row=r, column=3, value=pat_y2); ws_bg.cell(row=r, column=4, value=pat_y1)
        format_row(ws_bg, r, 'sub'); r += 1
        
        ws_bg.cell(row=r, column=1, value="Total Pasivos y Patrimonio"); ws_bg.cell(row=r, column=3, value=pc_y2+pnc_y2+pat_y2); ws_bg.cell(row=r, column=4, value=pc_y1+pnc_y1+pat_y1)
        format_row(ws_bg, r, 'tot')

        # 2. ESTADO DE RESULTADOS
        ws_er = wb.create_sheet("Estado de Resultados")
        r = create_sheet_header(ws_er, "ESTADO DE RESULTADOS INTEGRALES")
        ws_er.cell(row=r, column=1, value="Ingresos operacionales:"); format_row(ws_er, r, 'sec'); r += 1
        ing_y2, ing_y1 = 0, 0
        for _, row in df_comp[df_comp['codigo'].str.startswith('4', na=False)].iterrows():
            v2, v1 = abs(row['saldo_final_Y2']), abs(row['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            ing_y2 += v2; ing_y1 += v1
            ws_er.cell(row=r, column=1, value=row['cuenta'].title()); ws_er.cell(row=r, column=3, value=v2); ws_er.cell(row=r, column=4, value=v1)
            format_row(ws_er, r, 'normal'); r += 1
        ws_er.cell(row=r, column=1, value="Total Ingresos"); ws_er.cell(row=r, column=3, value=ing_y2); ws_er.cell(row=r, column=4, value=ing_y1)
        format_row(ws_er, r, 'sub'); r += 1
        
        ws_er.cell(row=r, column=1, value="Costos y gastos operacionales:"); format_row(ws_er, r, 'sec'); r += 1
        gas_y2, gas_y1 = 0, 0
        for _, row in df_comp[df_comp['codigo'].str.startswith(('5','6'), na=False)].iterrows():
            v2, v1 = -abs(row['saldo_final_Y2']), -abs(row['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            gas_y2 += v2; gas_y1 += v1
            ws_er.cell(row=r, column=1, value=row['cuenta'].title()); ws_er.cell(row=r, column=3, value=v2); ws_er.cell(row=r, column=4, value=v1)
            format_row(ws_er, r, 'normal'); r += 1
        ws_er.cell(row=r, column=1, value="Total Costos y Gastos"); ws_er.cell(row=r, column=3, value=gas_y2); ws_er.cell(row=r, column=4, value=gas_y1)
        format_row(ws_er, r, 'sub'); r += 1
        ws_er.cell(row=r, column=1, value="Utilidad (Pérdida) Neta del Período"); ws_er.cell(row=r, column=3, value=ing_y2+gas_y2); ws_er.cell(row=r, column=4, value=ing_y1+gas_y1)
        format_row(ws_er, r, 'tot')

        # 3. FLUJO DE EFECTIVO
        ws_fe = wb.create_sheet("Flujo de Efectivo")
        r = create_sheet_header(ws_fe, "HOJA DE TRABAJO FLUJO DE EFECTIVO")
        ws_fe.cell(row=r, column=1, value="Activos:"); format_row(ws_fe, r, 'sec'); r += 1
        for _, row in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
            v2, v1 = abs(row['saldo_final_Y2']), abs(row['saldo_final_Y1'])
            if 'acum' in str(row['cuenta']).lower(): v2, v1 = -v2, -v1
            if v2 == 0 and v1 == 0: continue
            ws_fe.cell(row=r, column=1, value=row['cuenta'].title()); ws_fe.cell(row=r, column=3, value=v2); ws_fe.cell(row=r, column=4, value=v1)
            format_row(ws_fe, r, 'normal'); r += 1
        ws_fe.cell(row=r, column=1, value="Pasivos y Patrimonio:"); format_row(ws_fe, r, 'sec'); r += 1
        for _, row in df_comp[df_comp['codigo'].str.startswith(('2','3'), na=False)].iterrows():
            v2, v1 = -abs(row['saldo_final_Y2']), -abs(row['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            ws_fe.cell(row=r, column=1, value=row['cuenta'].title()); ws_fe.cell(row=r, column=3, value=v2); ws_fe.cell(row=r, column=4, value=v1)
            format_row(ws_fe, r, 'normal'); r += 1

        # 4. ANEXO ACTIVOS FIJOS (Horizontal Corporativo)
        ws_af = wb.create_sheet("Anexo Activos Fijos")
        ws_af.column_dimensions['A'].width = 30
        cats = ['Terrenos y edificaciones', 'Instalaciones', 'Equipos industriales', 'Mobiliarios y equipos', 'Otros activos', 'Construcción en proceso']
        for i, c in enumerate(cats, 2): ws_af.column_dimensions[get_column_letter(i)].width = 22
        ws_af.column_dimensions[get_column_letter(len(cats)+2)].width = 20
        ws_af["A1"] = f"{empresa.upper()}"; ws_af["A2"] = "NOTA: PROPIEDAD, PLANTA Y EQUIPO (ANEXO)"
        for x in range(1, 3): ws_af[f"A{x}"].font = FNT_TITLE
        
        data = {c: {'c_y0':0, 'c_y1':0, 'c_y2':0, 'd_y0':0, 'd_y1':0, 'd_y2':0} for c in cats}
        for _, row in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
            if es_activo_no_corriente(row['codigo'], row['cuenta']):
                n = str(row['cuenta']).lower()
                if any(x in n for x in ['terreno', 'edific']): c = 'Terrenos y edificaciones'
                elif 'instalacion' in n: c = 'Instalaciones'
                elif any(x in n for x in ['maquinaria', 'transporte', 'vehiculo', 'indus']): c = 'Equipos industriales'
                elif 'proceso' in n: c = 'Construcción en proceso'
                elif 'otro' in n or 'mejora' in n: c = 'Otros activos'
                else: c = 'Mobiliarios y equipos'
                y1, y2 = abs(row['saldo_final_Y1']), abs(row['saldo_final_Y2'])
                if 'acum' in n: data[c]['d_y1'] += y1; data[c]['d_y2'] += y2; data[c]['d_y0'] += (y1 * 0.8)
                else: data[c]['c_y1'] += y1; data[c]['c_y2'] += y2; data[c]['c_y0'] += (y1 * 0.8)

        def write_af_block(ws, r_start, anio_lbl, k_ini, k_fin, d_ini, d_fin):
            ws.cell(row=r_start, column=1, value=f"Año Fiscal {anio_lbl}").font = FNT_TITLE
            r = r_start + 1
            ws.cell(row=r, column=1, value="Conceptos").font = FNT_HDR; ws.cell(row=r, column=1).fill = FILL_HDR
            for i, c in enumerate(cats, 2): 
                ws.cell(row=r, column=i, value=c).font = FNT_HDR; ws.cell(row=r, column=i).fill = FILL_HDR; ws.cell(row=r, column=i).alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=len(cats)+2, value="Total").font = FNT_HDR; ws.cell(row=r, column=len(cats)+2).fill = FILL_HDR
            r += 1
            
            def write_af_row(lbl, vals, mult=1, sub=False, tot=False):
                nonlocal r
                ws.cell(row=r, column=1, value=lbl)
                s = 0
                for i, v in enumerate(vals, 2): 
                    ws.cell(row=r, column=i, value=v*mult); s += (v*mult)
                ws.cell(row=r, column=len(cats)+2, value=s)
                format_row(ws, r, 'tot' if tot else ('sub' if sub else 'normal'))
                r += 1
                
            ws.cell(row=r, column=1, value="Costos:"); format_row(ws, r, 'sec'); r += 1
            c_i = [data[c][k_ini] for c in cats]; c_f = [data[c][k_fin] for c in cats]
            write_af_row("Balance al inicio", c_i)
            write_af_row("Adiciones", [max(0, f - i) for f, i in zip(c_f, c_i)])
            write_af_row("Retiros/Transferencias", [min(0, f - i) for f, i in zip(c_f, c_i)])
            write_af_row("Balance al costo final", c_f, sub=True)
            
            ws.cell(row=r, column=1, value="Depreciación:"); format_row(ws, r, 'sec'); r += 1
            d_i = [data[c][d_ini] for c in cats]; d_f = [data[c][d_fin] for c in cats]
            write_af_row("Balance al inicio", d_i, -1)
            write_af_row("Gasto de depreciación", [max(0, f - i) for f, i in zip(d_f, d_i)], -1)
            write_af_row("Retiros", [min(0, f - i) for f, i in zip(d_f, d_i)], -1)
            write_af_row("Dep. Acumulada final", d_f, -1, sub=True)
            write_af_row("Balance neto al final", [cf - df for cf, df in zip(c_f, d_f)], tot=True)
            return r + 2

        r = write_af_block(ws_af, 4, anio_act, 'c_y1', 'c_y2', 'd_y1', 'd_y2')
        write_af_block(ws_af, r, int(anio_act)-1, 'c_y0', 'c_y1', 'd_y0', 'd_y1')

        buf = io.BytesIO(); wb.save(buf)
        return buf.getvalue()
    except Exception as e: 
        st.error(f"Error generando Excel: {e}")
        return None

# ──────────────────────────────────────────────────────────────────────────────
# GENERADORES DE TABLAS HTML CORPORATIVAS
# ──────────────────────────────────────────────────────────────────────────────
def html_estado_resultados(df_comp, anio):
    html = f"<table class='tabla-contable'><tr><th>Conceptos</th><th>Nota</th><th>{anio}</th><th>{int(anio)-1}</th></tr>"
    html += "<tr><td class='seccion' colspan='4'>Ingresos operacionales:</td></tr>"
    ing_y2, ing_y1 = 0, 0
    for _, r in df_comp[df_comp['codigo'].str.startswith('4', na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if v2 == 0 and v1 == 0: continue
        ing_y2 += v2; ing_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += f"<tr class='subtotal'><td>Total Ingresos</td><td></td><td>{fmt_c(ing_y2)}</td><td>{fmt_c(ing_y1)}</td></tr>"
    
    html += "<tr><td class='seccion' colspan='4'>Costos y gastos operacionales:</td></tr>"
    gas_y2, gas_y1 = 0, 0
    for _, r in df_comp[df_comp['codigo'].str.startswith(('5','6'), na=False)].iterrows():
        v2, v1 = -abs(r['saldo_final_Y2']), -abs(r['saldo_final_Y1'])
        if v2 == 0 and v1 == 0: continue
        gas_y2 += v2; gas_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += f"<tr class='subtotal'><td>Total Costos y Gastos</td><td></td><td>{fmt_c(gas_y2)}</td><td>{fmt_c(gas_y1)}</td></tr>"
    html += f"<tr class='total'><td>Utilidad (Pérdida) del Período</td><td></td><td>{fmt_c(ing_y2 + gas_y2)}</td><td>{fmt_c(ing_y1 + gas_y1)}</td></tr></table>"
    return html

def html_balance_general(df_comp, anio, tipo='activo'):
    html = f"<table class='tabla-contable'><tr><th>{tipo.capitalize()}s</th><th>Nota</th><th>{anio}</th><th>{int(anio)-1}</th></tr>"
    tot_c_y2, tot_c_y1, tot_nc_y2, tot_nc_y1 = 0, 0, 0, 0
    
    if tipo == 'activo':
        html += "<tr><td class='seccion' colspan='4'>Activos corrientes:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
            if not es_activo_no_corriente(r['codigo'], r['cuenta']):
                v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
                if v2 == 0 and v1 == 0: continue
                tot_c_y2 += v2; tot_c_y1 += v1
                html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total activos corrientes</td><td></td><td>{fmt_c(tot_c_y2)}</td><td>{fmt_c(tot_c_y1)}</td></tr>"
        
        html += "<tr><td class='seccion' colspan='4'>Activos no corrientes:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
            if es_activo_no_corriente(r['codigo'], r['cuenta']):
                v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
                if 'acum' in str(r['cuenta']).lower(): v2, v1 = -v2, -v1 
                if v2 == 0 and v1 == 0: continue
                tot_nc_y2 += v2; tot_nc_y1 += v1
                html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total activos no corrientes</td><td></td><td>{fmt_c(tot_nc_y2)}</td><td>{fmt_c(tot_nc_y1)}</td></tr>"
    else:
        html += "<tr><td class='seccion' colspan='4'>Pasivos corrientes:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('2', na=False)].iterrows():
            if not es_pasivo_no_corriente(r['codigo'], r['cuenta']):
                v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
                if v2 == 0 and v1 == 0: continue
                tot_c_y2 += v2; tot_c_y1 += v1
                html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total pasivos corrientes</td><td></td><td>{fmt_c(tot_c_y2)}</td><td>{fmt_c(tot_c_y1)}</td></tr>"
        
        html += "<tr><td class='seccion' colspan='4'>Pasivos no corrientes:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('2', na=False)].iterrows():
            if es_pasivo_no_corriente(r['codigo'], r['cuenta']):
                v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
                if v2 == 0 and v1 == 0: continue
                tot_nc_y2 += v2; tot_nc_y1 += v1
                html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total pasivos no corrientes</td><td></td><td>{fmt_c(tot_nc_y2)}</td><td>{fmt_c(tot_nc_y1)}</td></tr>"
        
        pat_y2, pat_y1 = 0, 0
        html += "<tr><td class='seccion' colspan='4'>Patrimonio:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('3', na=False)].iterrows():
            v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            pat_y2 += v2; pat_y1 += v1
            html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total Patrimonio</td><td></td><td>{fmt_c(pat_y2)}</td><td>{fmt_c(pat_y1)}</td></tr>"
    
    gran_tot_y2 = tot_c_y2 + tot_nc_y2 + (pat_y2 if tipo == 'pasivo' else 0)
    gran_tot_y1 = tot_c_y1 + tot_nc_y1 + (pat_y1 if tipo == 'pasivo' else 0)
    titulo_tot = f"Total {tipo.capitalize()}s" if tipo == 'activo' else "Total Pasivos y Patrimonio"
    html += f"<tr class='total'><td>{titulo_tot}</td><td></td><td>{fmt_c(gran_tot_y2)}</td><td>{fmt_c(gran_tot_y1)}</td></tr></table>"
    return html

def html_flujo_hoja_trabajo(df_comp, anio):
    html = f"<table class='tabla-contable'><tr><th>Hoja de Flujo de Efectivo</th><th>{anio}</th><th>{int(anio)-1}</th></tr>"
    html += "<tr><td class='seccion' colspan='3'>Activos</td></tr>"
    for _, r in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if 'acum' in str(r['cuenta']).lower(): v2, v1 = -v2, -v1
        if v2 != 0 or v1 != 0: html += f"<tr><td>{r['cuenta'].title()}</td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += "<tr><td class='seccion' colspan='3'>Pasivos</td></tr>"
    for _, r in df_comp[df_comp['codigo'].str.startswith('2', na=False)].iterrows():
        v2, v1 = -abs(r['saldo_final_Y2']), -abs(r['saldo_final_Y1']) 
        if v2 != 0 or v1 != 0: html += f"<tr><td>{r['cuenta'].title()}</td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    return html + "</table>"

def html_nota_ppe_completa(df_comp, anio_actual):
    cats = ['Terrenos y edificaciones', 'Instalaciones recreativas', 'Equipos industriales y transporte', 'Mobiliarios y equipos de oficina', 'Otros activos y mejoras', 'Construcción en proceso']
    data = {c: {'c_y0':0, 'c_y1':0, 'c_y2':0, 'd_y0':0, 'd_y1':0, 'd_y2':0} for c in cats}
    for _, r in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
        if es_activo_no_corriente(r['codigo'], r['cuenta']):
            n = str(r['cuenta']).lower()
            if any(x in n for x in ['terreno', 'edific']): c = 'Terrenos y edificaciones'
            elif 'instalacion' in n: c = 'Instalaciones recreativas'
            elif any(x in n for x in ['maquinaria', 'transporte', 'vehiculo', 'industrial']): c = 'Equipos industriales y transporte'
            elif any(x in n for x in ['proceso', 'transito']): c = 'Construcción en proceso'
            elif 'otro' in n or 'mejora' in n: c = 'Otros activos y mejoras'
            else: c = 'Mobiliarios y equipos de oficina'
            y1, y2 = abs(r['saldo_final_Y1']), abs(r['saldo_final_Y2'])
            if 'acum' in n: 
                data[c]['d_y1'] += y1; data[c]['d_y2'] += y2; data[c]['d_y0'] += (y1 * 0.8)
            else: 
                data[c]['c_y1'] += y1; data[c]['c_y2'] += y2; data[c]['c_y0'] += (y1 * 0.8)

    def bloque(titulo, k_ini, k_fin, d_ini, d_fin):
        h = f"<tr><td colspan='8' class='titulo-anio'>{titulo}</td></tr><tr><th></th>"
        for c in cats: h += f"<th style='text-align: right; width: 14%;'>{c}</th>"
        h += "<th style='text-align: right; width: 14%;'>Total</th></tr>"
        def fila(lbl, vals, mult=1, sub=False, tot=False):
            cls = "total" if tot else ("subtotal" if sub else "")
            r = f"<tr class='{cls}'><td>{lbl}</td>"
            s = 0
            for v in vals: r += f"<td>{fmt_c(v * mult)}</td>"; s += (v * mult)
            return r + f"<td>{fmt_c(s)}</td></tr>"
            
        h += "<tr><td class='seccion' colspan='8'>Costos:</td></tr>"
        c_i = [data[c][k_ini] for c in cats]; c_f = [data[c][k_fin] for c in cats]
        h += fila("Balance al inicio", c_i)
        h += fila("Adiciones", [max(0, f - i) for f, i in zip(c_f, c_i)])
        h += fila("Retiros/Transferencias", [min(0, f - i) for f, i in zip(c_f, c_i)])
        h += fila("Balance al costo final", c_f, sub=True)
        
        h += "<tr><td class='seccion' colspan='8'>Depreciación:</td></tr>"
        d_i = [data[c][d_ini] for c in cats]; d_f = [data[c][d_fin] for c in cats]
        h += fila("Balance al inicio", d_i, -1)
        h += fila("Gasto de depreciación", [max(0, f - i) for f, i in zip(d_f, d_i)], -1)
        h += fila("Retiros", [min(0, f - i) for f, i in zip(d_f, d_i)], -1)
        h += fila("Dep. Acumulada final", d_f, -1, sub=True)
        h += fila("Balance neto al final", [cf - df for cf, df in zip(c_f, d_f)], tot=True)
        return h
    return "<table class='tabla-contable'>" + bloque(anio_actual, 'c_y1', 'c_y2', 'd_y1', 'd_y2') + bloque(int(anio_actual)-1, 'c_y0', 'c_y1', 'd_y0', 'd_y1') + "</table>"

def html_borrador_ir2(df_bal, periodo):
    ingresos = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
    costos = abs(df_bal[df_bal['codigo'].str.startswith('5', na=False)]['saldo_final'].sum())
    gastos = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False)]['saldo_final'].sum())
    
    utilidad_neta = ingresos - costos - gastos
    renta_neta_imponible = max(0, utilidad_neta)
    isr_liquidado = renta_neta_imponible * 0.27
    
    activos_totales = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
    impuesto_activos = activos_totales * 0.01
    impuesto_mayor = max(isr_liquidado, impuesto_activos)

    html = f"""
    <table class='tabla-ir2'>
        <tr><th colspan='3'>DECLARACIÓN JURADA ANUAL DEL IMPUESTO SOBRE LA RENTA DE SOCIEDADES (IR-2)<br><span style='font-weight:normal; font-size:0.85rem;'>AÑO FISCAL: {periodo}</span></th></tr>
        <tr><td colspan='3' class='header-seccion'>I. DETERMINACIÓN DE LA RENTA NETA IMPONIBLE O PÉRDIDA</td></tr>
        <tr><td class='col-num'>1</td><td class='col-desc'>Total de Ingresos Brutos</td><td class='col-monto'>RD$ {ingresos:,.2f}</td></tr>
        <tr><td class='col-num'>2</td><td class='col-desc'>Menos: Costo de Ventas</td><td class='col-monto' style='color:#dc2626;'>RD$ ({costos:,.2f})</td></tr>
        <tr><td class='col-num'>3</td><td class='col-desc'>Menos: Gastos Operacionales y Financieros</td><td class='col-monto' style='color:#dc2626;'>RD$ ({gastos:,.2f})</td></tr>
        <tr class='fila-total'><td class='col-num'>4</td><td class='col-desc'>Utilidad (o Pérdida) Neta antes de Impuestos</td><td class='col-monto'>RD$ {utilidad_neta:,.2f}</td></tr>
        
        <tr><td colspan='3' class='header-seccion'>II. LIQUIDACIÓN DEL IMPUESTO SOBRE LA RENTA</td></tr>
        <tr><td class='col-num'>5</td><td class='col-desc'>Renta Neta Imponible (Base de cálculo)</td><td class='col-monto'>RD$ {renta_neta_imponible:,.2f}</td></tr>
        <tr class='fila-total'><td class='col-num'>6</td><td class='col-desc'>Impuesto Liquidado (Tasa del 27%)</td><td class='col-monto'>RD$ {isr_liquidado:,.2f}</td></tr>
        
        <tr><td colspan='3' class='header-seccion'>III. LIQUIDACIÓN DEL IMPUESTO A LOS ACTIVOS (Anexo A)</td></tr>
        <tr><td class='col-num'>7</td><td class='col-desc'>Total Activos Imponibles (Aproximación)</td><td class='col-monto'>RD$ {activos_totales:,.2f}</td></tr>
        <tr class='fila-total'><td class='col-num'>8</td><td class='col-desc'>Impuesto a los Activos (Tasa del 1%)</td><td class='col-monto'>RD$ {impuesto_activos:,.2f}</td></tr>
        
        <tr><td colspan='3' class='header-seccion'>IV. RESUMEN DE PAGO</td></tr>
        <tr class='fila-total' style='background-color:#1e3a5f; color:white;'>
            <td class='col-num' style='background-color:#1e3a5f;'>9</td>
            <td class='col-desc'>IMPUESTO MAYOR A PAGAR (ISR vs Activos)</td>
            <td class='col-monto'>RD$ {impuesto_mayor:,.2f}</td>
        </tr>
    </table>
    """
    return html

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR & MAIN SETUP
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.title("🏛️ TaxTech Auditor RD")
st.sidebar.markdown("---")
empresa = st.sidebar.text_input("Empresa", value="Empresa de Prueba SRL")
rnc     = st.sidebar.text_input("RNC", value="1-31-12345-6")
periodo = st.sidebar.text_input("Período", value="Enero - Diciembre 2026")
anio    = st.sidebar.text_input("Año Fiscal", value="2026")

st.sidebar.markdown("---")
st.sidebar.title("Materialidad (NIA 320)")
pct_mp   = st.sidebar.slider("% Materialidad", 0.5, 3.0, 1.0, 0.1) / 100

st.title("TaxTech Auditor — Declaraciones Juradas & Estados Financieros")
c_up1, c_up2 = st.columns(2)
with c_up1: uploaded = st.file_uploader("📂 Cargar Balanza (Año Actual)", type=["xlsx", "xls", "csv"])
with c_up2: uploaded_prev = st.file_uploader("📂 Cargar Balanza (Año Anterior)", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.info("👆 Sube la balanza de comprobación para iniciar.")
    st.stop()

with st.spinner("Procesando datos contables..."):
    df_bal = procesar_balanza(uploaded)
    if df_bal.empty: 
        st.error("Error al leer el archivo. Verifica el formato.")
        st.stop()
        
    df_bal = analizar_balanza(df_bal)

    if uploaded_prev:
        df_prev = procesar_balanza(uploaded_prev)
        df_comp = procesar_comparativo(df_bal, df_prev) if not df_prev.empty else pd.DataFrame()
    else:
        df_comp = df_bal.copy()
        df_comp.rename(columns={'saldo_final': 'saldo_final_Y2'}, inplace=True)
        df_comp['saldo_final_Y1'] = 0.0
        df_comp['variacion_abs'] = df_comp['saldo_final_Y2']

# KPIs y Cálculos Fiscales Restaurados
t_ingresos = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
t_activos  = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
t_costos   = abs(df_bal[df_bal['codigo'].str.startswith('5', na=False)]['saldo_final'].sum())
t_gastos   = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False)]['saldo_final'].sum())
utilidad_neta = t_ingresos - t_costos - t_gastos
mp = (t_ingresos if t_ingresos > 0 else t_activos) * pct_mp
ir2_vals = calcular_casillas_ir2(df_bal)

# Estimación de otras obligaciones fiscales
itbis_pagado = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False) & df_bal['cuenta'].str.lower().str.contains('itbis', na=False)]['saldo_final'].sum())
itbis_retenido = abs(df_bal[df_bal['codigo'].str.startswith('2', na=False) & df_bal['cuenta'].str.lower().str.contains('itbis', na=False)]['saldo_final'].sum())
neto_itbis_c = itbis_retenido - itbis_pagado

t_gastos_personal = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False) & df_bal['cuenta'].str.lower().str.contains('personal|sueldo', na=False)]['saldo_final'].sum())
total_tss_c = t_gastos_personal * (TASA_SFS_PAT + TASA_AFP_PAT)

t_honorarios = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False) & df_bal['cuenta'].str.lower().str.contains('honorario|profesional', na=False)]['saldo_final'].sum())
total_ir17_c = t_honorarios * 0.10

isr_est = max(0, utilidad_neta) * 0.27

st.markdown(f"### 📌 {empresa} — {periodo}")

# TABS COMPLETOS E INTACTOS
tab_comp, tab_bg, tab_er, tab_efe, tab_ppe, tab_bal, tab_inconsist, tab_art287, tab_ir2, tab_consol = st.tabs([
    "📈 Dashboard", "📊 Balance General", "📉 Estado de Resultados", "🌊 Flujo de Efectivo", "🏗️ Anexo Activos Fijos",
    "📋 Balanza Creada", "🚨 Inconsistencias", "⚖️ Riesgos Art.287", "📝 Borrador IR-2", "🏛️ Consolidado Fiscal"
])

try:
    with tab_comp:
        st.markdown("### Dashboard Corporativo")
        c1, c2 = st.columns(2)
        with c1: 
            st.metric("Ingresos Año Actual", f"RD$ {t_ingresos:,.0f}")
            st.metric("Materialidad Planificación", f"RD$ {mp:,.0f}")
        with c2:
            st.markdown("#### Exportar Paquete Financiero")
            excel_bytes = exportar_reporte_corporativo(empresa, periodo, anio, df_comp)
            if excel_bytes:
                st.download_button("📥 Descargar Excel Corporativo", data=excel_bytes, file_name=f"Reporte_{empresa.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        df_chart = pd.DataFrame({'Año': [f"{int(anio)-1}", f"{anio}"], 'Ingresos': [sum(abs(df_comp[df_comp['codigo'].str.startswith('4', na=False)]['saldo_final_Y1'])), t_ingresos], 'Activos': [sum(abs(df_comp[df_comp['codigo'].str.startswith('1', na=False)]['saldo_final_Y1'])), t_activos]})
        st.plotly_chart(px.bar(df_chart, x='Año', y=['Ingresos', 'Activos'], barmode='group', title="Evolución Ingresos vs Activos", color_discrete_sequence=['#1e3a5f', '#38bdf8']), use_container_width=True)

    with tab_bg:
        c1, c2 = st.columns(2)
        with c1: st.markdown(html_balance_general(df_comp, anio, 'activo'), unsafe_allow_html=True)
        with c2: st.markdown(html_balance_general(df_comp, anio, 'pasivo'), unsafe_allow_html=True)
        
    with tab_er: st.markdown(html_estado_resultados(df_comp, anio), unsafe_allow_html=True)
    with tab_efe: st.markdown(html_flujo_hoja_trabajo(df_comp, anio), unsafe_allow_html=True)
    with tab_ppe: st.markdown(html_nota_ppe_completa(df_comp, anio), unsafe_allow_html=True)

    with tab_bal:
        st.dataframe(df_bal[['codigo', 'cuenta', 'debito', 'credito', 'saldo_final']], use_container_width=True)

    with tab_inconsist:
        df_err = df_bal[~df_bal['validacion_naturaleza'].str.startswith('✅')]
        if df_err.empty: st.success("✅ Sin inconsistencias en naturaleza de saldos.")
        else: st.dataframe(df_err[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)

    with tab_art287:
        df_fisc = df_bal[df_bal['alerta_fiscal'] != ""]
        if df_fisc.empty: st.success("✅ Sin alertas fiscales Art. 287 detectadas.")
        else: st.dataframe(df_fisc[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal']], use_container_width=True)

    with tab_ir2:
        st.markdown("### Formulario IR-2: Determinación de Obligación")
        st.markdown(html_borrador_ir2(df_bal, periodo), unsafe_allow_html=True)

    with tab_consol:
        df_consol = pd.DataFrame([
            ("IT-1",  "ITBIS Mensual",                    f"RD$ {abs(neto_itbis_c):,.2f}", "A Pagar" if neto_itbis_c > 0 else "Saldo Favor", "✅" if neto_itbis_c <= 0 else "⚠️"),
            ("IR-3",  "TSS / Seguridad Social",           f"RD$ {total_tss_c:,.2f}",       "Estimado",     "ℹ️"),
            ("IR-17", "Otras Retenciones",                f"RD$ {total_ir17_c:,.2f}",      "A Pagar",      "⚠️" if total_ir17_c > 0 else "✅"),
            ("IR-2",  "Ajuste Patrimonial",               f"RD$ {ir2_vals.get('cas_34', 0):,.2f}","Base",         "ℹ️"),
            ("ISR",   "Impuesto Renta Estimado (27%)",    f"RD$ {isr_est:,.2f}",           "Proyección",   "⚠️" if isr_est > 0 else "✅"),
        ], columns=["Formulario", "Concepto", "Monto", "Estado", ""])
        st.dataframe(df_consol, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Ocurrió un error al renderizar las tablas: {e}")