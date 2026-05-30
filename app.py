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
# 1. PARAMETRIZACIÓN FISCAL DOMINICANA
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

TASA_ITBIS_GENERAL = 0.18
TASA_SFS_PATRONAL = 0.0709
TASA_AFP_PATRONAL = 0.0710
TASA_SRL_PROMEDIO = 0.0120
TASA_INFOTEP = 0.0100
TASA_SFS_EMPLEADO = 0.0304
TASA_AFP_EMPLEADO = 0.0287
COSTO_PER_CAPITA_2026 = 1691.38

# Estilos de Diseño Corporativo
FILL_HEADER = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F497D")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=11, bold=False)
FONT_BOLD = Font(name="Calibri", size=11, bold=True)
BORDER_THIN = Side(border_style="thin", color="D9D9D9")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)

# ==========================================
# 2. MOTOR DE RENDERIZADO EXCEL MOTORIZADO
# ==========================================
def armar_encabezado_excel(ws, titulo, subtitulo, headers):
    """Inicializa la hoja con el formato corporativo estándar."""
    ws.views.sheetView[0].showGridLines = True
    ws["B2"] = titulo
    ws["B2"].font = FONT_TITLE
    ws["B3"] = subtitulo
    ws["B3"].font = Font(name="Calibri", size=11, italic=True)
    
    for col_idx, text in enumerate(headers, start=2):
        cell = ws.cell(row=5, column=col_idx, value=text)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center" if "casilla" in text.lower() or "línea" in text.lower() or "porcentaje" in text.lower() else "left")

def escribir_datos_excel(ws, start_row, dataset, total_label="TOTAL", sum_column_indices=[]):
    """Pinta de forma dinámica las filas, cebras, formatos y la fila de totales finales."""
    r_idx = start_row
    for idx, fila in enumerate(dataset):
        for c_idx, valor in enumerate(fila, start=2):
            cell = ws.cell(row=r_idx, column=c_idx, value=valor)
            cell.font = FONT_BODY
            cell.border = CELL_BORDER
            if idx % 2 == 1:
                cell.fill = FILL_ZEBRA
            
            # Formateo inteligente según contenido
            if isinstance(valor, float) or isinstance(valor, int):
                if valor <= 1.0 and c_idx == 4:  # Columna de tasa/porcentaje
                    cell.number_format = '0.00%'
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.number_format = 'RD$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right")
            elif str(valor).startswith('='):
                cell.number_format = 'RD$ #,##0.00'
                cell.alignment = Alignment(horizontal="right")
        r_idx += 1
        
    # Construcción limpia de la Fila de Totales
    ws.cell(row=r_idx, column=2, value=total_label).alignment = Alignment(horizontal="center")
    for c in range(2, len(ws[5]) + 2):
        cell = ws.cell(row=r_idx, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = CELL_BORDER
        if c in sum_column_indices:
            col_letter = get_column_letter(c)
            cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{r_idx-1})"
            cell.number_format = 'RD$ #,##0.00'
            cell.alignment = Alignment(horizontal="right")
            
    # Autoajuste de columnas
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        if letter != 'A':
            ws.column_dimensions[letter].width = 30 if letter in ['E', 'F'] else 50

# ==========================================
# 3. FUNCIONES DE LÓGICA DE AUDITORÍA
# ==========================================
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl', dtype={'codigo': str, 'código': str})
        else:
            df = pd.read_csv(file, dtype={'codigo': str, 'código': str})
        
        df.columns = [str(c).strip().lower() for c in df.columns]
        mapeo = {
            'código': 'codigo', 'cuenta': 'cuenta', 'nombre': 'cuenta', 'nombre de cuenta': 'cuenta',
            'débito': 'debito', 'crédito': 'credito', 'saldo final': 'saldo_final', 'saldo': 'saldo_final'
        }
        df = df.rename(columns=mapeo)
        
        col_req = {'codigo', 'cuenta', 'debito', 'credito', 'saldo_final'}
        if not col_req.issubset(set(df.columns)):
            st.error("⚠️ Columnas requeridas ausentes en la Balanza.")
            return pd.DataFrame()
            
        df['codigo'] = df['codigo'].fillna('').astype(str).str.strip().apply(lambda x: x.split('.')[0] if '.' in x else x)
        df['cuenta'] = df['cuenta'].fillna('').astype(str).str.strip()
        df = df[(df['codigo'] != '') & (~df['codigo'].str.lower().str.contains('total|resultado|suma', na=False))]
        
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df
    except Exception as e:
        st.error(f"❌ Error base: {str(e)}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat, alertas_fisc, casillas_ir2 = [], [], []
    for _, row in df.iterrows():
        cod, nom = row['codigo'], str(row['cuenta']).strip().lower()
        if not cod or any(k in nom for k in ['total', 'suma']):
            alertas_nat.append("Ignorado"); alertas_fisc.append("Sin observaciones"); casillas_ir2.append("No Aplica")
            continue
            
        nat_esp = NATURALEZAS.get(cod[0], None)
        saldo = row['saldo_final']
        
        if nat_esp == 'Debito' and saldo < 0:
            alertas_nat.append("Saldo Crédito inusual (Naturaleza Débito)")
        elif nat_esp == 'Credito' and saldo > 0:
            alertas_nat.append("Saldo Débito inusual (Naturaleza Crédito)")
        else:
            alertas_nat.append("Correcto")
            
        alerta_f = "Sin observaciones"
        for palabra, mensaje in PALABRAS_CRITICAS_ART287.items():
            if palabra in nom:
                alerta_f = mensaje
                break
        alertas_fisc.append(alerta_f)
        casillas_ir2.append(MAPEO_IR2.get(cod[:2], MAPEO_IR2.get(cod[0] + '1', 'Otros Conceptos No Mapeados')))
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal_rd'] = alertas_fisc
    df['casilla_ir2'] = casillas_ir2
    return df

# ==========================================
# 4. CUERPO PRINCIPAL DE LA APP (STREAMLIT)
# ==========================================
st.sidebar.title("Configuración del Cliente")
empresa = st.sidebar.text_input("Nombre de la Empresa", value="Empresa de Prueba SRL")
periodo = st.sidebar.text_input("Período de Análisis", value="2026/05/30")

st.sidebar.markdown("---")
st.sidebar.title("Parámetros de Materialidad (NIA 320)")
tipo_entidad = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])
tasa_ref = 0.01 if tipo_entidad == "Comercial / Servicios" else 0.005
porcentaje_mp = st.sidebar.slider("Porcentaje de Materialidad", 0.5, 3.0, tasa_ref * 100, step=0.1) / 100
porcentaje_me = st.sidebar.slider("Porcentaje de Materialidad de Ejecución (ME)", 50, 75, 75, step=5) / 100

st.header("1. Carga de Balanza de Comprobación Base")
uploaded_file = st.file_uploader("Upload Balanza", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded_file is not None:
    df_balanza = procesar_balanza(uploaded_file)
    if not df_balanza.empty:
        df_balanza = analizar_balanza(df_balanza)
        
        total_activos = abs(df_balanza[df_balanza['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
        total_ingresos = abs(df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
        base_calc = total_ingresos if total_ingresos > 0 else total_activos
        mp = base_calc * porcentaje_mp
        me = mp * porcentaje_me
        
        st.markdown("---")
        st.subheader(f"📌 Informe de Auditoría Analítica: {empresa} — Período: {periodo}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos Declarados", f"RD$ {total_ingresos:,.2f}")
        c2.metric("Total Activos Registrados", f"RD$ {total_activos:,.2f}")
        c3.metric("Materialidad Planificación (MP)", f"RD$ {mp:,.2f}")
        c4.metric("Materialidad Ejecución (ME)", f"RD$ {me:,.2f}")
        st.markdown("---")
        
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "📋 Balanza", "🚨 Inconsistencias", "🇩🇴 Riesgos Art. 287", 
            "📋 Borrador Anual IR-2", "⚡ Mensual (Liquidación IT-1)", 
            "🏢 Nómina y TSS (IR-3)", "💸 Liquidación IR-17", "🏛️ Consolidado Fiscal General"
        ])
        
        with tab1:
            st.dataframe(df_balanza, use_container_width=True)
        with tab2:
            df_err = df_balanza[(df_balanza['validacion_naturaleza'] != "Correcto") & (df_balanza['validacion_naturaleza'] != "Ignorado")]
            if not df_err.empty:
                st.error("Se detectaron inconsistencias en la naturaleza contable de los saldos.")
                st.dataframe(df_err[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)
            else:
                st.success("✅ Excelente: No se han encontrado cuentas con inconsistencias.")
        with tab3:
            df_fisc = df_balanza[df_balanza['alerta_fiscal_rd'] != "Sin observaciones"]
            if not df_fisc.empty:
                st.warning("Se identificaron cuentas expuestas a revisión fiscal del Art. 287.")
                st.dataframe(df_fisc[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal_rd']], use_container_width=True)
            else:
                st.success("✅ Cumplimiento Inicial: No se detectaron alertas críticas.")
        with tab4:
            st.markdown("### 📋 Estructura Acumulada para Declaración Anual IR-2")
            df_ir2 = df_balanza.groupby('casilla_ir2')['saldo_final'].sum().reset_index()
            df_ir2['saldo_final'] = df_ir2['saldo_final'].apply(abs)
            df_ir2.columns = ['Renglón Formulario DGII', 'Monto Acumulado (RD$)']
            st.dataframe(df_ir2, use_container_width=True)
            
        # --- CÁLCULOS IMPOSITIVOS ---
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
        
        r_honorarios = b_honorarios * 0.10; r_reparaciones = b_reparaciones * 0.02
        r_vehiculos = b_vehiculos * 0.27; r_renta = b_renta_vivienda * 0.27; r_otras_ret = b_otras_retribuciones * 0.27
        r_espana = b_espana * 0.10; r_canada = b_canada * 0.18; r_exterior = b_exterior_general * 0.27
        total_ir17 = r_honorarios + r_reparaciones + r_vehiculos + r_renta + r_otras_ret + r_espana + r_canada + r_exterior

        with tab5:
            st.markdown("### 🇩🇴 Módulo de Liquidación de ITBIS (Anexo A + Formulario IT-1)")
            if neto_itbis_resultado > 0:
                st.error(f"🚨 **IMPUESTO NETO A PAGAR EN IT-1:** RD$ {neto_itbis_resultado:,.2f}")
            else:
                st.success(f"🎉 **SALDO A FAVOR COMPENSABLE:** RD$ {abs(neto_itbis_resultado):,.2f}")
                
            buffer_it1 = io.BytesIO()
            wb_it1 = openpyxl.Workbook()
            ws_anexo = wb_it1.active; ws_anexo.title = "Anexo_A_IT1"
            armar_encabezado_excel(ws_anexo, f"TAXTECH AUDITOR RD — ANEXO A IT-1: {empresa.upper()}", "Desglose de Ingresos, Adelantos y Retenciones", ["Casilla", "Descripción del Concepto Operativo", "Monto Base Imponible", "ITBIS Liquidado / Retenido"])
            
            datos_anexo = [
                ("Casilla 1", "Ingresos por Operaciones Locales Gravadas (Ventas)", monto_ingresos_gravados, itbis_ventas_generado),
                ("Casilla 12", "ITBIS Pagado en Compras Locales (Adelantos del 606)", compras_y_gastos_base, itbis_soportado_compras),
                ("Casilla 22", "ITBIS Retenido por Servicios Profesionales (100%)", base_honorarios_fisicos, itbis_ret_100_fisicas),
                ("Casilla 23", "ITBIS Retenido entre Personas Jurídicas (30% Norma 02-05)", base_servicios_juridicas, itbis_ret_30_juridicas),
                ("Casilla 30", "Retención por Ventas con Tarjetas de Crédito (2% Norma 08-04)", 0.00, itbis_ret_tarjetas_2)
            ]
            escribir_datos_excel(ws_anexo, 6, datos_anexo, "TOTALES ANEXO", [])
            
            ws_form = wb_it1.create_sheet(title="Formulario_IT1")
            armar_encabezado_excel(ws_form, f"TAXTECH AUDITOR RD — FORMULARIO IT-1: {empresa.upper()}", "Liquidación Final", ["Línea", "Renglón Final IT-1", "Fórmula de Amarre", "Monto Relacionado"])
            datos_form = [
                ("Línea 1", "Total ITBIS Bruto por Operaciones", "=Anexo_A_IT1!E6", "=Anexo_A_IT1!E6"),
                ("Línea 2", "(-) Menos ITBIS Deducible (Compras)", "=Anexo_A_IT1!E7", "=Anexo_A_IT1!E7"),
                ("Línea 3", "(-) Menos ITBIS Retenido por Personas Físicas", "=Anexo_A_IT1!E8", "=Anexo_A_IT1!E8"),
                ("Línea 4", "(-) Menos ITBIS Retenido por Personas Jurídicas", "=Anexo_A_IT1!E9", "=Anexo_A_IT1!E9"),
                ("Línea 5", "(-) Menos Retención por Tarjetas de Crédito", "=Anexo_A_IT1!E10", "=Anexo_A_IT1!E10"),
                ("TOTAL", "IMPUESTO NETO A PAGAR / SALDO A FAVOR", "=E6-SUM(E7:E10)", "=E6-SUM(E7:E10)")
            ]
            escribir_datos_excel(ws_form, 6, datos_form, "TOTAL LIQUIDACIÓN", [])
            wb_it1.save(buffer_it1)
            st.download_button("📥 Descargar Libro Completo IT-1 (Excel)", data=buffer_it1.getvalue(), file_name=f"Borrador_Completo_IT1_{empresa.replace(' ', '_')}.xlsx")

        with tab6:
            st.markdown("### 🏢 Módulo de Conciliación y Liquidación TSS / INFOTEP / IR-3")
            uploaded_empleados = st.file_uploader("Subir listado auxiliar de empleados (Excel/CSV)", type=["xlsx", "csv"], key="tss_empleados")
            if uploaded_empleados is not None:
                st.success("✅ Detalle de empleados cargado correctamente.")
                
            buffer_tss = io.BytesIO()
            wb_tss = openpyxl.Workbook()
            ws_tss = wb_tss.active; ws_tss.title = "Liquidación TSS"
            armar_encabezado_excel(ws_tss, f"TAXTECH AUDITOR RD — COSTO DE NOMINA: {empresa.upper()}", "Liquidación Mensual Seg. Social", ["Tipo", "Concepto de Aporte o Retención", "Porcentaje", "Monto Autocalculado"])
            
            datos_tss = [
                ("Patronal", "Seguro Familiar de Salud (SFS Patronal)", 0.0709, gasto_nominas_global * 0.0709),
                ("Patronal", "Fondo de Pensiones (AFP Patronal)", 0.0710, gasto_nominas_global * 0.0710),
                ("Patronal", "Seguro de Riesgos Laborales (SRL)", 0.0120, gasto_nominas_global * 0.0120),
                ("Patronal", "Aporte INFOTEP Obligatorio (1%)", 0.0100, gasto_nominas_global * 0.0100),
                ("Empleado", "SFS Retención Trabajador", 0.0304, gasto_nominas_global * 0.0304),
                ("Empleado", "AFP Retención Trabajador", 0.0287, gasto_nominas_global * 0.0287),
                ("Empleado", "Aporte Percápita Adicional (Dependientes)", 0.0000, gasto_per_capita_balanza)
            ]
            escribir_datos_excel(ws_tss, 6, datos_tss, "TOTAL FACTURA TSS", [5])
            wb_tss.save(buffer_tss)
            st.download_button("📥 Descargar Plantilla Auxiliar TSS (Excel)", data=buffer_tss.getvalue(), file_name=f"Plantilla_TSS_{empresa.replace(' ', '_')}.xlsx")

        with tab7:
            st.markdown("### 💸 Formulario IR-17: Declaración Jurada de Otras Retenciones")
            st.error(f"💸 **TOTAL IMPUESTO A PAGAR (IR-17):** RD$ {total_ir17:,.2f}")
            
            buffer_ir17 = io.BytesIO()
            wb_ir17 = openpyxl.Workbook()
            ws_ir17 = wb_ir17.active; ws_ir17.title = "Borrador IR-17"
            armar_encabezado_excel(ws_ir17, f"TAXTECH AUDITOR RD — MAESTRO IR-17: {empresa.upper()}", "Borrador Resumido de Casillas", ["Casilla", "Concepto Detallado (DGII)", "Tasa", "Monto Base Imponible", "Impuesto Retenido"])
            
            datos_ir17 = [
                ("Casilla 1", "Honorarios por Servicios Profesionales (Persona Física)", 0.10, b_honorarios),
                ("Casilla 2", "Otras Rentas / Servicios Técnicos y Reparaciones", 0.02, b_reparaciones),
                ("Casilla 8", "Retribuciones Complementarias - Vehículos / Combustible", 0.27, b_vehiculos),
                ("Casilla 9", "Retribuciones Complementarias - Pago de Alquileres / Renta", 0.27, b_renta_vivienda),
                ("Casilla 10", "Retribuciones Complementarias - Otros Beneficios en Especie", 0.27, b_otras_retribuciones),
                ("Casilla 15", "Remesas al Exterior - Convenio Doble Imposición España", 0.10, b_espana),
                ("Casilla 16", "Remesas al Exterior - Convenio Doble Imposición Canadá", 0.18, b_canada),
                ("Casilla 17", "Remesas al Exterior - Otras Rentas de Fuente Dominicana (General)", 0.27, b_exterior_general)
            ]
            escribir_datos_excel(ws_ir17, 6, datos_ir17, "TOTAL", [5, 6])
            wb_ir17.save(buffer_ir17)
            st.download_button("📥 Descargar Borrador Resumido IR-17 (Excel)", data=buffer_ir17.getvalue(), file_name=f"Borrador_Resumido_IR17_{empresa.replace(' ', '_')}.xlsx")
            
        with tab8:
            st.markdown("### 🏛️ Consolidado Fiscal General del Periodo")
            itbis_caja = neto_itbis_resultado if neto_itbis_resultado > 0 else 0.0
            gran_total_periodo_pagar = itbis_caja + total_liquidacion_ir3_tss + total_ir17
            st.warning(f"🏦 **EFECTIVO TOTAL ESTIMADO A TRANSFERIR (DGII / TSS):** RD$ {gran_total_periodo_pagar:,.2f}")
            
            buffer_con = io.BytesIO()
            wb_con = openpyxl.Workbook()
            ws_con = wb_con.active; ws_con.title = "Consolidado Fiscal"
            armar_encabezado_excel(ws_con, f"TAXTECH AUDITOR RD — VOLANTE DE CONSOLIDACIÓN", "Resumen Maestro de Caja", ["Formulario Oficial", "Origen / Módulo", "Estado de Cuenta", "Monto Neto Determinado"])
            
            datos_con = [
                ("Formulario IT-1", "Módulo de ITBIS / 606 / 607", "Saldo a Pagar" if neto_itbis_resultado > 0 else "Saldo a Favor", itbis_caja),
                ("Tesorería TSS", "Seguridad Social Patronal", "Costo Operativo", costo_patronal_total),
                ("Formulario IR-3", "Retenciones Empleados Nómina", "Pasivo de Retención", retenciones_empleados_total),
                ("Formulario IR-17", "Retenciones Locales / Exterior", "Impuesto Retenido", total_ir17)
            ]
            escribir_datos_excel(ws_con, 6, datos_con, "TOTAL GENERAL", [5])
            wb_con.save(buffer_con)
            st.download_button("📥 Descargar Volante Consolidado de Caja (Excel)", data=buffer_con.getvalue(), file_name=f"Consolidado_Fiscal_General_{empresa.replace(' ', '_')}.xlsx")
else:
    st.info("👋 Por favor, carga tu archivo de Balanza de Comprobación para desplegar los cálculos automáticos.")
