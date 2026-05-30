import streamlit as st
import pandas as pd
import numpy as np
import io

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal",
    layout="wide"
)

# ==========================================
# 1. PARÁMETROS FISCALES Y DE CONFIGURACIÓN (CTRD / TSS / PERCÁPITA 2026)
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

# Coeficientes Oficiales Ley 87-01 e INFOTEP
TASA_SFS_PATRONAL = 0.0709
TASA_AFP_PATRONAL = 0.0710
TASA_SRL_PROMEDIO = 0.0120
TASA_INFOTEP = 0.0100

TASA_SFS_EMPLEADO = 0.0304
TASA_AFP_EMPLEADO = 0.0287

# Costo Indexado Oficial TSS para Dependiente Adicional
COSTO_PER_CAPITA_2026 = 1691.38

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
            st.error(f"⚠️ Estructura de archivo incorrecta. Debe incluir las columnas: {list(columnas_requeridas)}")
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
        st.error(f"❌ Error crítico en el procesamiento del archivo: {str(e)}")
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
# 3. INTERFAZ DE USUARIO (INTEGRACIÓN COMPLETA UI)
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

# --- CUERPO PRINCIPAL ---
st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")
st.header("1. Carga de Balanza de Comprobación")
st.markdown("Arrastra tu archivo Excel o CSV generado desde tu software contable (Odoo, QuickBooks, etc.)")

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
        
        # Módulo de KPIs Generales
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos Declarados", f"RD$ {total_ingresos:,.2f}")
        c2.metric("Total Activos Registrados", f"RD$ {total_activos:,.2f}")
        c3.metric("Materialidad Planificación (MP)", f"RD$ {mp:,.2f}")
        c4.metric("Materialidad Ejecución (ME)", f"RD$ {me:,.2f}")
        
        st.markdown("---")
        
        # PESTAÑAS PRINCIPALES
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📋 Balanza", "🚨 Inconsistencias", "🇩🇴 Riesgos Art. 287", 
            "📋 Borrador Anual IR-2", "⚡ Mensual (IT-1, 606, 607)", 
            "🏢 Nómina y TSS (IR-3)", "💸 Liquidación IR-17 (Retenciones)"
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
            df_ir2 = df_balanza.groupby('casilla_ir2')['saldo_final'].sum().reset_index()
            df_ir2['saldo_final'] = df_ir2['saldo_final'].apply(lambda x: abs(x))
            df_ir2.columns = ['Renglón Formulario DGII', 'Monto Acumulado (RD$)']
            st.dataframe(df_ir2, use_container_width=True)
            
        with tab5:
            st.markdown("### 🇩🇴 Gestión de Impuestos Mensuales y Formatos de Envío")
            with st.expander("📊 Resumen Formulario IT-1 (ITBIS)", expanded=True):
                st.metric("ITBIS Bruto Generado (Sugerido 18%)", f"RD$ {total_ingresos * 0.18:,.2f}")
            with st.expander("📥 Desglose Analítico - Formato 606", expanded=False):
                df_606 = df_balanza[(df_balanza['codigo'].str.startswith(('5', '6'), na=False)) & (~df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario|tss|infotep', na=False))].copy()
                df_606['saldo_final'] = df_606['saldo_final'].apply(lambda x: abs(x))
                df_606['itbis_estimado_compras'] = df_606['saldo_final'] * 0.18
                st.dataframe(df_606[['codigo', 'cuenta', 'saldo_final', 'itbis_estimado_compras']], use_container_width=True)
                
        with tab6:
            st.markdown("### 🏢 Módulo de Conciliación y Liquidación TSS / INFOTEP / IR-3")
            gasto_nominas_global = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario', na=False)]['saldo_final'].sum())
            
            # Escáner dinámico para encontrar la cuenta de descuento per cápita adicional en la balanza
            gasto_per_capita_balanza = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('percapita|per capita|dependiente adicional', na=False)]['saldo_final'].sum())
            # Determinamos cuántos dependientes representa ese monto retenido
            num_dependientes_estimados = round(gasto_per_capita_balanza / COSTO_PER_CAPITA_2026) if gasto_per_capita_balanza > 0 else 0
            
            st.metric("Masa Salarial Global Detectada", f"RD$ {gasto_nominas_global:,.2f}")
            st.markdown("---")
            
            # FILA 1: DESGLOSE DE APORTES PATRONALES (COSTO EMPRESA)
            st.markdown("##### 💼 Bloque 1: Aportes Patronales (Costo de la Empresa - Gasto Deducible)")
            cp1, cp2, cp3, cp4 = st.columns(4)
            cp1.metric("SFS Patronal (7.09%)", f"RD$ {gasto_nominas_global * TASA_SFS_PATRONAL:,.2f}")
            cp2.metric("AFP Patronal (7.10%)", f"RD$ {gasto_nominas_global * TASA_AFP_PATRONAL:,.2f}")
            cp3.metric("Seguro Riesgos Laborales (1.20%)", f"RD$ {gasto_nominas_global * TASA_SRL_PROMEDIO:,.2f}")
            cp4.metric("Aporte INFOTEP (1.00%)", f"RD$ {gasto_nominas_global * TASA_INFOTEP:,.2f}")
            
            st.markdown("---")
            
            # FILA 2: DESGLOSE DE RETENCIONES AL EMPLEADO (DESCUENTOS DE NÓMINA)
            st.markdown("##### 👥 Bloque 2: Retenciones al Empleado (Pasivo / Descuentos de Nómina)")
            ce1, ce2, ce3, ce4 = st.columns(4)
            ce1.metric("SFS Retenido (3.04%)", f"RD$ {gasto_nominas_global * TASA_SFS_EMPLEADO:,.2f}")
            ce2.metric("AFP Retenido (2.87%)", f"RD$ {gasto_nominas_global * TASA_AFP_EMPLEADO:,.2f}")
            
            # IMPACTO DEL APORTE PERCAPITA ADICIONAL DETECTADO
            ce3.metric("Aporte Percápita Adicional (TSS)", f"RD$ {gasto_per_capita_balanza:,.2f}", delta=f"{num_dependientes_estimados} Dependientes", delta_color="inverse")
            
            tot_retenciones_emp = (gasto_nominas_global * TASA_SFS_EMPLEADO) + (gasto_nominas_global * TASA_AFP_EMPLEADO) + gasto_per_capita_balanza
            ce4.metric("Total Retenido a Empleados", f"RD$ {tot_retenciones_emp:,.2f}")
            
            st.markdown("---")
            st.markdown("#### 📋 Simulación de Carga Masiva TSS por Empleados")
            
            if gasto_nominas_global > 0:
                # Distribuimos el descuento per cápita detectado en la plantilla de simulación para validación analítica
                empleados_data = {
                    'Cédula / Identificación': ['001-XXXXXXX-1', '001-XXXXXXX-2', '001-XXXXXXX-3'],
                    'Nombre del Empleado': ['Personal Operativo A', 'Personal Administrativo B', 'Dirección / Gerencia C'],
                    'Sueldo Cotizable (RD$)': [gasto_nominas_global * 0.40, gasto_nominas_global * 0.35, gasto_nominas_global * 0.25],
                    'Dependientes Adicionales': [num_dependientes_estimados if num_dependientes_estimados > 0 else 0, 0, 0]
                }
                df_empleados = pd.DataFrame(empleados_data)
                
                df_empleados['SFS Empleado (3.04%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_SFS_EMPLEADO
                df_empleados['AFP Empleado (2.87%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_AFP_EMPLEADO
                df_empleados['Percápita Adicional (Descuento)'] = df_empleados['Dependientes Adicionales'] * COSTO_PER_CAPITA_2026
                df_empleados['SFS Patronal (7.09%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_SFS_PATRONAL
                df_empleados['AFP Patronal (7.10%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_AFP_PATRONAL
                df_empleados['SRL Patronal (1.20%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_SRL_PROMEDIO
                df_empleados['INFOTEP Patronal (1.00%)'] = df_empleados['Sueldo Cotizable (RD$)'] * TASA_INFOTEP
                
                st.dataframe(df_empleados.style.format({
                    'Sueldo Cotizable (RD$)': 'RD$ {:,.2f}',
                    'SFS Empleado (3.04%)': 'RD$ {:,.2f}',
                    'AFP Empleado (2.87%)': 'RD$ {:,.2f}',
                    'Percápita Adicional (Descuento)': 'RD$ {:,.2f}',
                    'SFS Patronal (7.09%)': 'RD$ {:,.2f}',
                    'AFP Patronal (7.10%)': 'RD$ {:,.2f}',
                    'SRL Patronal (1.20%)': 'RD$ {:,.2f}',
                    'INFOTEP Patronal (1.00%)': 'RD$ {:,.2f}'
                }), use_container_width=True)
                
                buffer_tss = io.BytesIO()
                with pd.ExcelWriter(buffer_tss, engine='openpyxl') as writer:
                    df_empleados.to_excel(writer, index=False, sheet_name='Borrador_TSS_IR3')
                    
                st.download_button(
                    label="📥 Descargar Plantilla Auxiliar TSS & IR-3 (Excel)",
                    data=buffer_tss.getvalue(),
                    file_name=f"Plantilla_TSS_INFOTEP_{empresa.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        with tab7:
            st.markdown("### 💸 Liquidación IR-17 (Retenciones)")
            base_honorarios = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
            base_reparaciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('reparacion|mantenimiento', na=False)]['saldo_final'].sum())
            base_retribuciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('retribucion|especie|alquiler personal|renta personal|vehiculo personal', na=False)]['saldo_final'].sum())
            
            base_espana = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('españa|espana', na=False)]['saldo_final'].sum())
            base_canada = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('canada|canadá', na=False)]['saldo_final'].sum())
            base_exterior_general = abs(df_balanza[(df_balanza['cuenta'].str.lower().str.contains('exterior|remesa|extranjero', na=False)) & (~df_balanza['cuenta'].str.lower().str.contains('españa|espana|canada|canadá', na=False))]['saldo_final'].sum())
            
            ret_honorarios = base_honorarios * 0.10
            ret_reparaciones = base_reparaciones * 0.02
            ret_retribuciones = base_retribuciones * 0.27
            ret_espana = base_espana * 0.10
            ret_canada = base_canada * 0.18
            ret_exterior = base_exterior_general * 0.27
            
            total_ir17_liquidar = ret_honorarios + ret_reparaciones + ret_retribuciones + ret_espana + ret_canada + ret_exterior
            st.metric("💵 Total a Liquidar IR-17", f"RD$ {total_ir17_liquidar:,.2f}")
            
            df_ir17_resumen = pd.DataFrame({
                'Concepto o Tipo de Retención': [
                    'Honorarios Profesionales (Persona Física - 10%)', 'Servicios Técnicos y Reparaciones (Persona Física - 2%)',
                    'Retribuciones Complementarias (27%)', 'Remesas Exterior - Convenio España (10%)',
                    'Remesas Exterior - Convenio Canadá (18%)', 'Remesas Exterior - Otros Países (27%)'
                ],
                'Monto Base Imponible (RD$)': [base_honorarios, base_reparaciones, base_retribuciones, base_espana, base_canada, base_exterior_general],
                'Impuesto Retenido (RD$)': [ret_honorarios, ret_reparaciones, ret_retribuciones, ret_espana, ret_canada, ret_exterior]
            })
            st.dataframe(df_ir17_resumen.style.format({
                'Monto Base Imponible (RD$)': 'RD$ {:,.2f}', 'Impuesto Retenido (RD$)': 'RD$ {:,.2f}'
            }), use_container_width=True)
else:
    st.info("👋 Por favor, carga tu archivo de Balanza de Comprobación para desplegar los cálculos automáticos.")
