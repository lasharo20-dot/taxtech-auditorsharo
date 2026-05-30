import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal",
    layout="wide"
)

# ==========================================
# 1. PARÁMETROS FISCALES Y DE CONFIGURACIÓN (CTRD / TSS / ITBIS 2026)
# ==========================================
NATURALEZAS = {
    '1': 'Debito', '2': 'Credito', '3': 'Credito',
    '4': 'Credito', '5': 'Debito', '6': 'Debito'
}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Riesgo Art. 287 CTRD: Validar deducibilidad, comprobantes con NCF válido y uso de medios de pago para crédito fiscal de ITBIS.',
    'representacion': 'Riesgo Art. 287 CTRD: Gastos de representación. Sujetos a criterios de razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Riesgo Art. 318 CTRD / Reg. 139-98: Retribuciones en especie. Validar que la empresa efectúe el pago del ISR sustitutivo correspondiente.',
    'gasto de personal': 'Riesgo Art. 287 CTRD: Cruce obligatorio con la declaración jurada de TSS (Formulario IR-4) para admitir la deducción.',
    'honorario': 'Riesgo Art. 309 CTRD: Validar aplicación de retenciones fiscales (10% a personas físicas o 2% entre personas jurídicas).'
}

MAPEO_IR2 = {
    '11': 'Anexo A - Efectivo e Inversiones Temporales',
    '12': 'Anexo A - Cuentas por Cobrar (Neto)',
    '13': 'Anexo A - Inventarios',
    '15': 'Anexo A - Propiedad, Planta y Equipo (Neto)',
    '21': 'Anexo A - Pasivos Corrientes / Cuentas por Pagar',
    '31': 'Anexo A - Capital Social y Reservas',
    '41': 'Anexo B - Ingresos por Operaciones (Locales)',
    '51': 'Anexo B - Costo de Ventas / Servicios',
    '61': 'Anexo B - Gastos de Personal (TSS / Sueldos)',
    '62': 'Anexo B - Gastos Operativos y de Administración',
    '63': 'Anexo B - Gastos Financieros'
}

# Coeficientes Oficiales de Ley
TASA_ITBIS_GENERAL = 0.18
TASA_SFS_PATRONAL = 0.0709
TASA_AFP_PATRONAL = 0.0710
TASA_SRL_PROMEDIO = 0.0120
TASA_INFOTEP = 0.0100
TASA_SFS_EMPLEADO = 0.0304
TASA_AFP_EMPLEADO = 0.0287
COSTO_PER_CAPITA_2026 = 1691.38

# --- ESTILOS COMPARTIDOS PARA EXPORTACIÓN EXCEL ---
FILL_HEADER = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F497D")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=11, bold=False)
FONT_BOLD = Font(name="Calibri", size=11, bold=True)
BORDER_THIN = Side(border_style="thin", color="D9D9D9")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)

def aplicar_estilos_base(ws, titulo, subtitulo):
    ws.views.sheetView[0].showGridLines = True
    ws["B2"] = titulo
    ws["B2"].font = FONT_TITLE
    ws["B3"] = subtitulo
    ws["B3"].font = Font(name="Calibri", size=11, italic=True)

def autoajustar_columnas(ws):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter != 'A':
            ws.column_dimensions[col_letter].width = 30 if col_letter in ['E', 'F'] else 50

# ==========================================
# 2. FUNCIONES DE LÓGICA DE NEGOCIO Y AUDITORÍA
# ==========================================
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl', dtype={'codigo': str, 'código': str})
        else:
            df = pd.read_csv(file, dtype={'codigo': str, 'código': str})
        
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        mapeo_columnas = {
            'código': 'codigo', 'cuenta': 'cuenta', 'nombre': 'cuenta', 'nombre de cuenta': 'cuenta',
            'débito': 'debito', 'crédito': 'credito', 'saldo final': 'saldo_final', 'saldo': 'saldo_final'
        }
        df = df.rename(columns=mapeo_columnas)
        
        columnas_requeridas = {'codigo', 'cuenta', 'debito', 'credito', 'saldo_final'}
        if not columnas_requeridas.issubset(set(df.columns)):
            st.error(f"⚠️ Estructura incorrecta en la balanza de comprobación.")
            return pd.DataFrame()
            
        df['codigo'] = df['codigo'].fillna('').astype(str).str.strip()
        df['codigo'] = df['codigo'].apply(lambda x: x.split('.')[0] if '.' in x else x)
        df['cuenta'] = df['cuenta'].fillna('').astype(str).str.strip()
        
        df = df[~df['codigo'].str.lower().str.contains('total|resultado|suma', na=False)]
        df = df[~df['cuenta'].str.lower().str.contains('total|resultado|suma', na=False)]
        df = df[df['codigo'] != '']
        
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
        return df
    except Exception as e:
        st.error(f"❌ Error al procesar archivo base: {str(e)}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat = []
    alertas_fisc = []
    casillas_ir2 = []
    
    for _, row in df.iterrows():
        codigo_str = row['codigo']
        nombre_cuenta = str(row['cuenta']).strip().lower()
        
        if not codigo_str or any(keyword in nombre_cuenta for keyword in ['total', 'suma']):
            alertas_nat.append("Ignorado")
            alertas_fisc.append("Sin observaciones")
            casillas_ir2.append("No Aplica")
            continue
            
        primer_digito = codigo_str[0]
        primeros_dos = codigo_str[:2]
        nat_esperada = NATURALEZAS.get(primer_digito, None)
        saldo = row['saldo_final']
        
        if nat_esperada == 'Debito' and saldo < 0:
            alertas_nat.append("Saldo Crédito inusual (Naturaleza Débito)")
        elif nat_esperada == 'Credito' and saldo > 0:
            alertas_nat.append("Saldo Débito inusual (Naturaleza Crédito)")
        else:
            alertas_nat.append("Correcto")
            
        alerta_f = "Sin observaciones"
        for palabra, mensaje in PALABRAS_CRITICAS_ART287.items():
            if palabra in nombre_cuenta:
                alerta_f = mensaje
                break
        alertas_fisc.append(alerta_f)
        
        casilla = MAPEO_IR2.get(primeros_dos, MAPEO_IR2.get(primer_digito + '1', 'Otros Conceptos No Mapeados'))
        casillas_ir2.append(casilla)
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal_rd'] = alertas_fisc
    df['casilla_ir2'] = casillas_ir2
    return df

# ==========================================
# 3. INTERFAZ DE USUARIO PRINCIPAL
# ==========================================

# --- BARRA LATERAL ---
st.sidebar.title("Configuración del Cliente")
empresa = st.sidebar.text_input("Nombre de la Empresa", value="Empresa de Prueba SRL")
periodo = st.sidebar.text_input("Período de Análisis", value="2026/05/30")

st.sidebar.markdown("---")
st.sidebar.title("Parámetros de Materialidad (NIA 320)")
tipo_entidad = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])

tasa_referencia = 0.01 if tipo_entidad == "Comercial / Servicios" else 0.005
porcentaje_mp = st.sidebar.slider("Porcentaje de Materialidad", 0.5, 3.0, tasa_referencia * 100, step=0.1) / 100
porcentaje_me = st.sidebar.slider("Porcentaje de Materialidad de Ejecución (ME)", 50, 75, 75, step=5) / 100

st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")
st.header("1. Carga de Balanza de Comprobación Base")

uploaded_file = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded_file is not None:
    df_balanza = procesar_balanza(uploaded_file)
    
    if not df_balanza.empty:
        df_balanza = analizar_balanza(df_balanza)
        
        total_activos = abs(df_balanza[df_balanza['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
        total_ingresos = abs(df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
        
        base_calculo = total_ingresos if total_ingresos > 0 else total_activos
        mp = base_calculo * porcentaje_mp
        me = mp * porcentaje_me
        
        st.markdown("---")
        st.subheader(f"📌 Informe de Auditoría Analítica: {empresa} — Período: {periodo}")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos Declarados", f"RD$ {total_ingresos:,.2f}")
        c2.metric("Total Activos Registrados", f"RD$ {total_activos:,.2f}")
        c3.metric("Materialidad Planificación (MP)", f"RD$ {mp:,.2f}")
        c4.metric("Materialidad Ejecución (ME)", f"RD$ {me:,.2f}")
        
        st.markdown("---")
        
        # PESTAÑAS PRINCIPALES
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "📋 Balanza", "🚨 Inconsistencias", "🇩🇴 Riesgos Art. 287", 
            "📋 Borrador Anual IR-2", "⚡ Mensual (Liquidación IT-1)", 
            "🏢 Nómina y TSS (IR-3)", "💸 Liquidación IR-17", "🏛️ Consolidado Fiscal General"
        ])
        
        with tab1:
            st.dataframe(df_balanza, use_container_width=True)
            
        with tab2:
            df_errores = df_balanza[(df_balanza['validacion_naturaleza'] != "Correcto") & (df_balanza['validacion_naturaleza'] != "Ignorado")]
            if not df_errores.empty:
                st.error(f"Se detectaron {len(df_errores)} cuentas con saldos inconsistentes.")
                st.dataframe(df_errores[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)
            else:
                st.success("✅ Excelente: No se han encontrado cuentas con inconsistencias en su saldo final.")
                
        with tab3:
            df_fiscal = df_balanza[df_balanza['alerta_fiscal_rd'] != "Sin observaciones"]
            if not df_fiscal.empty:
                st.warning(f"Atención: Se identificaron {len(df_fiscal)} cuentas expuestas a revisión fiscal.")
                st.dataframe(df_fiscal[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal_rd']], use_container_width=True)
            else:
                st.success("✅ Cumplimiento Inicial: No se detectaron alertas críticas.")
                
        with tab4:
            st.markdown("### 📋 Mapeo y Cruce Avanzado - Formulario Anual IR-2")
            df_ir2 = df_balanza.groupby('casilla_ir2')['saldo_final'].sum().reset_index()
            df_ir2['saldo_final'] = df_ir2['saldo_final'].apply(lambda x: abs(x))
            df_ir2.columns = ['Renglón Formulario DGII', 'Monto Acumulado (RD$)']
            st.dataframe(df_ir2, use_container_width=True)
            
        # --- EXTRACCIÓN DE VALORES COMUNES ---
        monto_ingresos_gravados = abs(df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
        itbis_ventas_generado = monto_ingresos_gravados * TASA_ITBIS_GENERAL
        compras_y_gastos_base = abs(df_balanza[(df_balanza['codigo'].str.startswith(('5', '6'), na=False)) & (~df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario|tss|infotep|percapita', na=False))]['saldo_final'].sum())
        itbis_soportado_compras = compras_y_gastos_base * TASA_ITBIS_GENERAL
        base_honorarios_fisicos = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
        itbis_ret_100_fisicas = (base_honorarios_fisicos * TASA_ITBIS_GENERAL) * 1.00
        base_servicios_juridicas = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('servicio tecnico|consultoria|reparacion', na=False)]['saldo_final'].sum())
        itbis_ret_30_juridicas = (base_servicios_juridicas * TASA_ITBIS_GENERAL) * 0.30
        cuenta_ret_tarjeta = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('retencion tarjeta|adquirente|cardnet|azul', na=False)]['saldo_final'].sum())
        itbis_ret_tarjetas_2 = cuenta_ret_tarjeta if cuenta_ret_tarjeta > 0 else (monto_ingresos_gravados * 0.60) * 0.02
        neto_itbis_resultado = itbis_ventas_generado - (itbis_soportado_compras + itbis_ret_100_fisicas + itbis_ret_30_juridicas + itbis_ret_tarjetas_2)

        gasto_nominas_global = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario', na=False)]['saldo_final'].sum())
        gasto_per_capita_balanza = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('percapita|per capita|dependiente adicional', na=False)]['saldo_final'].sum())
        costo_patronal_total = (gasto_nominas_global * TASA_SFS_PATRONAL) + (gasto_nominas_global * TASA_AFP_PATRONAL) + (gasto_nominas_global * TASA_SRL_PROMEDIO) + (gasto_nominas_global * TASA_INFOTEP)
        retenciones_empleados_total = (gasto_nominas_global * TASA_SFS_EMPLEADO) + (gasto_nominas_global * TASA_AFP_EMPLEADO) + gasto_per_capita_balanza
        total_liquidacion_ir3_tss = costo_patronal_total + retenciones_empleados_total

        b_honorarios = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
        b_reparaciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('reparacion|mantenimiento', na=False)]['saldo_final'].sum())
        b_vehiculos = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('vehiculo personal|combustible empleado', na=False)]['saldo_final'].sum())
        b_renta_vivienda = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('renta personal|alquiler personal|vivienda', na=False)]['saldo_final'].sum())
        b_otras_retribuciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('retribucion|especie', na=False) & ~df_balanza['cuenta'].str.lower().str.contains('vehiculo|renta|alquiler|vivienda', na=False)]['saldo_final'].sum())
        b_espana = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('españa|espana', na=False)]['saldo_final'].sum())
        b_canada = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('canada|canadá', na=False)]['saldo_final'].sum())
        b_exterior_general = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('exterior|remesa|extranjero', na=False) & ~df_balanza['cuenta'].str.lower().str.contains('españa|espana|canada|canadá', na=False)]['saldo_final'].sum())
        
        r_honorarios = b_honorarios * 0.10
        r_reparaciones = b_reparaciones * 0.02
        r_vehiculos = b_vehiculos * 0.27
        r_renta = b_renta_vivienda * 0.27
        r_otras_ret = b_otras_retribuciones * 0.27
        r_espana = b_espana * 0.10
        r_canada = b_canada * 0.18
        r_exterior = b_exterior_general * 0.27
        total_ir17 = r_honorarios + r_reparaciones + r_vehiculos + r_renta + r_otras_ret + r_espana + r_canada + r_exterior

        with tab5:
            st.markdown("### 🇩🇴 Módulo de Liquidación de ITBIS (Anexo A + Formulario IT-1)")
            if neto_itbis_resultado > 0:
                st.error(f"🚨 **IMPUESTO NETO A PAGAR EN IT-1:** RD$ {neto_itbis_resultado:,.2f}")
            else:
                st.success(f"🎉 **SALDO A FAVOR COMPENSABLE:** RD$ {abs(neto_itbis_resultado):,.2f}")
            
            cit1, cit2, cit3, cit4 = st.columns(4)
            cit1.metric("ITBIS Ventas (Generado)", f"RD$ {itbis_ventas_generado:,.2f}")
            cit2.metric("ITBIS Soportado (Adelantos)", f"RD$ {itbis_soportado_compras:,.2f}")
            cit3.metric("Retenciones Sufridas", f"RD$ {itbis_ret_100_fisicas + itbis_ret_30_juridicas:,.2f}")
            cit4.metric("Retención Tarjeta (2%)", f"RD$ {itbis_ret_tarjetas_2:,.2f}")
            
            buffer_it1 = io.BytesIO()
            wb_it1 = openpyxl.Workbook()
            
            ws_anexo = wb_it1.active
            ws_anexo.title = "Anexo_A_IT1"
            aplicar_estilos_base(ws_anexo, f"TAXTECH AUDITOR RD — ANEXO A DEL IT-1: {empresa.upper()}", f"Desglose Analítico de Ingresos, Adelantos y Retenciones")
            
            headers_anexo = ["Casilla", "Descripción del Concepto Operativo", "Monto Base Imponible", "ITBIS Liquidado / Retenido"]
            for col_idx, h in enumerate(headers_anexo, start=2):
                cell = ws_anexo.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_anexo = [
                ("Casilla 1", "Ingresos por Operaciones Locales Gravadas (Ventas)", monto_ingresos_gravados, itbis_ventas_generado),
                ("Casilla 12", "ITBIS Pagado en Compras Locales (Adelantos del 606)", compras_y_gastos_base, itbis_soportado_compras),
                ("Casilla 22", "ITBIS Retenido por Servicios Profesionales de Personas Físicas (100%)", base_honorarios_fisicos, itbis_ret_100_fisicas),
                ("Casilla 23", "ITBIS Retenido entre Personas Jurídicas (30% Norma 02-05)", base_servicios_juridicas, itbis_ret_30_juridicas),
                ("Casilla 30", "Retención por Operaciones con Tarjetas de Crédito (2% Norma 08-04)", 0.00, itbis_ret_tarjetas_2)
            ]
            
            for idx, (cas, con, bas, imp) in enumerate(datos_anexo):
                r_idx = 6 + idx
                ws_anexo.cell(row=r_idx, column=2, value=cas).alignment = Alignment(horizontal="center")
                ws_anexo.cell(row=r_idx, column=3, value=con)
                c_bas = ws_anexo.cell(row=r_idx, column=4, value=bas)
                c_bas.number_format = 'RD$ #,##0.00'; c_bas.alignment = Alignment(horizontal="right")
                c_imp = ws_anexo.cell(row=r_idx, column=5, value=imp)
                c_imp.number_format = 'RD$ #,##0.00'; c_imp.alignment = Alignment(horizontal="right")
                
                for c in range(2, 6):
                    cell = ws_anexo.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY; cell.border = CELL_BORDER
                    if idx % 2 == 1: cell.fill = FILL_ZEBRA
            autoajustar_columnas(ws_anexo)
            
            ws_form = wb_it1.create_sheet(title="Formulario_IT1")
            aplicar_estilos_base(ws_form, f"TAXTECH AUDITOR RD — FORMULARIO DEFINITIVO IT-1: {empresa.upper()}", f"Liquidación Final del Impuesto — Período: {periodo}")
            
            headers_form = ["Línea", "Renglón Final IT-1", "Fórmula de Amarre (DGII)", "Monto Relacionado"]
            for col_idx, h in enumerate(headers_form, start=2):
                cell = ws_form.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_form = [
                ("Línea 1", "Total ITBIS Bruto por Operaciones", "=Anexo_A_IT1!E6", "=Anexo_A_IT1!E6"),
                ("Línea 2", "(-) Menos ITBIS Deducible (Adelanto de Compras)", "=Anexo_A_IT1!E7", "=Anexo_A_IT1!E7"),
                ("Línea 3", "(-) Menos ITBIS Retenido por Personas Físicas", "=Anexo_A_IT1!E8", "=Anexo_A_IT1!E8"),
                ("Línea 4", "(-) Menos ITBIS Retenido por Personas Jurídicas", "=Anexo_A_IT1!E9", "=Anexo_A_IT1!E9"),
                ("Línea 5", "(-) Menos Retención por Tarjetas de Crédito", "=Anexo_A_IT1!E10", "=Anexo_A_IT1!E10"),
                ("TOTAL", "IMPUESTO NETO A PAGAR / SALDO A FAVOR", "=E6-SUM(E7:E10)", "=E6-SUM(E7:E10)")
            ]
            
            for idx, (lin, ren, for_am, val) in enumerate(datos_form):
                r_idx = 6 + idx
                ws_form.cell(row=r_idx, column=2, value=lin).alignment = Alignment(horizontal="center")
                ws_form.cell(row=r_idx, column=3, value=ren)
                ws_form.cell(row=r_idx, column=4, value=for_am).font = Font(name="Calibri", size=9, color="7F7F7F")
                c_val = ws_form.cell(row=r_idx, column=5, value=val)
                c_val.number_format = 'RD$ #,##0.00'; c_val.alignment = Alignment(horizontal="right")
                
                for c in range(2, 6):
                    cell = ws_form.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY if lin != "TOTAL" else FONT_BOLD
                    cell.border = CELL_BORDER
                    if idx % 2 == 1 and lin != "TOTAL": cell.fill = FILL_ZEBRA
                    if lin == "TOTAL": cell.fill = FILL_TOTAL
            
            autoajustar_columnas(ws_form)
            wb_
