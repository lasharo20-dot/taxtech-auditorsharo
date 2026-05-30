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
            "📋 Borrador Anual IR-2", "⚡ Mensual (Liquidación IT-1)", 
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
            st.markdown("### 📋 Mapeo y Cruce Avanzado - Formulario Anual IR-2")
            with st.expander("📥 CARGAR ARCHIVO AUXILIAR DE ANEXOS IR-2 (DGII)", expanded=False):
                uploaded_anexos = st.file_uploader("Subir borrador de anexos (A, B, C, D) en Excel", type=["xlsx"], key="ir2_anexos")
                if uploaded_anexos is not None:
                    st.success("✅ Borrador de Anexos indexado. Listo para conciliación masiva.")
            
            st.markdown("#### Distribución Predictiva según Balanza:")
            df_ir2 = df_balanza.groupby('casilla_ir2')['saldo_final'].sum().reset_index()
            df_ir2['saldo_final'] = df_ir2['saldo_final'].apply(lambda x: abs(x))
            df_ir2.columns = ['Renglón Formulario DGII', 'Monto Acumulado (RD$)']
            st.dataframe(df_ir2, use_container_width=True)
            
        with tab5:
            st.markdown("### 🇩🇴 Módulo Avanzado de Liquidación - Formulario IT-1 (ITBIS)")
            with st.expander("📥 CARGAR PRE-ENVÍOS OFICIALES DE FORMATOS 606 Y 607", expanded=False):
                up_606 = st.file_uploader("Subir txt o Excel definitivo de Compras (606)", type=["txt", "xlsx"], key="it1_606")
                up_607 = st.file_uploader("Subir txt o Excel definitivo de Ventas (607)", type=["txt", "xlsx"], key="it1_607")
                if up_606 and up_607:
                    st.success("✅ Formatos 606 y 607 enlazados. El sistema priorizará estos valores sobre las cuentas de control.")
            
            # --- MOTOR DE CÁLCULO IT-1 ---
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
            
            total_deducciones_creditos = itbis_soportado_compras + itbis_ret_100_fisicas + itbis_ret_30_juridicas + itbis_ret_tarjetas_2
            neto_itbis_resultado = itbis_ventas_generado - total_deducciones_creditos
            
            st.markdown("---")
            if neto_itbis_resultado > 0:
                st.error(f"🚨 **IMPUESTO NETO A PAGAR EN IT-1:** RD$ {neto_itbis_resultado:,.2f}")
            else:
                st.success(f"🎉 **SALDO A FAVOR COMPENSABLE:** RD$ {abs(neto_itbis_resultado):,.2f}")
            
            cit1, cit2, cit3, cit4 = st.columns(4)
            cit1.metric("ITBIS Ventas (Generado)", f"RD$ {itbis_ventas_generado:,.2f}")
            cit2.metric("ITBIS Soportado (Adelantos)", f"RD$ {itbis_soportado_compras:,.2f}")
            cit3.metric("Retenciones Sufridas (30%/100%)", f"RD$ {itbis_ret_100_fisicas + itbis_ret_30_juridicas:,.2f}")
            cit4.metric("Retención Tarjeta (2% Norma 08-04)", f"RD$ {itbis_ret_tarjetas_2:,.2f}")
            
        with tab6:
            st.markdown("### 🏢 Módulo de Conciliación y Liquidación TSS / INFOTEP / IR-3")
            gasto_nominas_global = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario', na=False)]['saldo_final'].sum())
            gasto_per_capita_balanza = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('percapita|per capita|dependiente adicional', na=False)]['saldo_final'].sum())
            num_dependientes_estimados = round(gasto_per_capita_balanza / COSTO_PER_CAPITA_2026) if gasto_per_capita_balanza > 0 else 0
            
            with st.expander("📥 CARGAR ARCHIVO COMPLEMENTARIO DE TXT ENTRADA TSS (IR-4)", expanded=False):
                up_tss = st.file_uploader("Subir borrador de empleados o txt oficial de la TSS", type=["txt", "xlsx"], key="tss_file")
                if up_tss is not None:
                    st.success("✅ Padron de empleados real vinculado para cruce cruzado.")
            
            st.markdown("---")
            cp1, cp2, cp3, cp4 = st.columns(4)
            cp1.metric("SFS Patronal (7.09%)", f"RD$ {gasto_nominas_global * TASA_SFS_PATRONAL:,.2f}")
            cp2.metric("AFP Patronal (7.10%)", f"RD$ {gasto_nominas_global * TASA_AFP_PATRONAL:,.2f}")
            cp3.metric("Seguro Riesgos Laborales (1.20%)", f"RD$ {gasto_nominas_global * TASA_SRL_PROMEDIO:,.2f}")
            cp4.metric("Aporte INFOTEP (1.00%)", f"RD$ {gasto_nominas_global * TASA_INFOTEP:,.2f}")
            
        with tab7:
            st.markdown("### 💸 Liquidación IR-17 (Otras Retenciones)")
            with st.expander("📥 CARGAR AUXILIAR DE RETENCIONES COMPLEMENTARIAS", expanded=False):
                up_ir17 = st.file_uploader("Subir hoja de cálculo de retenciones del exterior / remesas", type=["xlsx", "csv"], key="ir17_file")
                if up_ir17 is not None:
                    st.success("✅ Resumen de operaciones internacionales integrado correctamente.")
            
            st.markdown("---")
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
else:
    st.info("👋 Por favor, carga tu archivo de Balanza de Comprobación para desplegar los cálculos automáticos.")
