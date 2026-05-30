import streamlit as st
import pandas as pd
import numpy as np

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal",
    layout="wide"
)

# ==========================================
# 1. PARÁMETROS FISCALES Y DE CONFIGURACIÓN (CTRD)
# ==========================================
NATURALEZAS = {
    '1': 'Debito',  # Activos
    '2': 'Credito', # Pasivos
    '3': 'Credito', # Capital
    '4': 'Credito', # Ingresos
    '5': 'Debito',  # Costos
    '6': 'Debito'   # Gastos
}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Riesgo Art. 287 CTRD: Validar deducibilidad, comprobantes con NCF válido y uso de medios de pago para crédito fiscal de ITBIS.',
    'representacion': 'Riesgo Art. 287 CTRD: Gastos de representación. Sujetos a criterios de razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Riesgo Art. 318 CTRD / Reg. 139-98: Retribuciones en especie. Validar que la empresa efectúe el pago del ISR sustitutivo correspondiente.',
    'gasto de personal': 'Riesgo Art. 287 CTRD: Cruce obligatorio con la declaración jurada de TSS (Formulario IR-4) para admitir la deducción.',
    'honorario': 'Riesgo Art. 309 CTRD: Validar aplicación de retenciones fiscales (10% a personas físicas o 2% entre personas jurídicas).'
}

# ==========================================
# 2. FUNCIONES DE LÓGICA DE NEGOCIO Y AUDITORÍA
# ==========================================
def procesar_balanza(file) -> pd.DataFrame:
    """Lee y normaliza la balanza de comprobación previniendo caídas del sistema."""
    try:
        # Forzamos que la columna 'codigo' se lea como texto desde el inicio si es posible
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
            
        # Blindaje: Limpieza estricta y conversión de códigos flotantes/enteros a texto plano
        df['codigo'] = df['codigo'].fillna('').astype(str).str.strip()
        # Elimina el ".0" que a veces Excel agrega automáticamente a los números enteros
        df['codigo'] = df['codigo'].apply(lambda x: x.split('.')[0] if '.' in x else x)
        
        # Limpieza de valores numéricos para los importes financieros
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
        df['cuenta'] = df['cuenta'].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"❌ Error crítico en el procesamiento del archivo: {str(e)}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica algoritmos de auditoría contable y fiscal cuenta por cuenta."""
    alertas_nat = []
    alertas_fisc = []
    
    for _, row in df.iterrows():
        codigo_str = row['codigo']
        
        # Validación de seguridad en caso de filas vacías en el Excel
        if not codigo_str:
            alertas_nat.append("Código Vacío")
            alertas_fisc.append("Sin observaciones")
            continue
            
        # Extraemos de forma segura el primer dígito como texto
        primer_digito = codigo_str[0]
        nat_esperada = NATURALEZAS.get(primer_digito, None)
        saldo = row['saldo_final']
        
        if nat_esperada == 'Debito' and saldo < 0:
            alertas_nat.append("Saldo Crédito inusual (Naturaleza Débito)")
        elif nat_esperada == 'Credito' and saldo < 0:
            alertas_nat.append("Saldo Débito inusual (Naturaleza Crédito)")
        else:
            alertas_nat.append("Correcto")
            
        nombre_cuenta = row['cuenta'].lower()
        alerta_f = "Sin observaciones"
        for palabra, mensaje in PALABRAS_CRITICAS_ART287.items():
            if palabra in nombre_cuenta:
                alerta_f = mensaje
                break
        alertas_fisc.append(alerta_f)
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal_rd'] = alertas_fisc
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

porcentaje_mp = st.sidebar.slider(
    "Porcentaje de Materialidad", 
    min_value=0.5, 
    max_value=3.0, 
    value=tasa_referencia * 100, 
    step=0.1
) / 100

porcentaje_me = st.sidebar.slider(
    "Porcentaje de Materialidad de Ejecución (ME)", 
    min_value=50, 
    max_value=75, 
    value=75, 
    step=5
) / 100

# --- CUERPO PRINCIPAL ---
st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")
st.header("1. Carga de Balanza de Comprobación")
st.markdown("Arrastra tu archivo Excel o CSV generado desde tu software contable (Odoo, QuickBooks, etc.)")

uploaded_file = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded_file is not None:
    df_balanza = procesar_balanza(uploaded_file)
    
    if not df_balanza.empty:
        df_balanza = analizar_balanza(df_balanza)
        
        # Aseguramos que la máscara booleana no falle si hay códigos vacíos
        total_activos = df_balanza[df_balanza['codigo'].str.startswith('1', na=False)]['saldo_final'].sum()
        total_ingresos = df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum()
        
        base_calculo = total_ingresos if total_ingresos > 0 else total_activos
        mp = base_calculo * porcentaje_mp
        me = mp * porcentaje_me
        
        st.markdown("---")
        st.subheader(f"📌 Informe de Auditoría Analítica: {empresa} — Período: {periodo}")
        
        # Módulo de KPIs
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos Declarados", f"RD$ {total_ingresos:,.2f}")
        c2.metric("Total Activos Registrados", f"RD$ {total_activos:,.2f}")
        c3.metric("Materialidad Planificación (MP)", f"RD$ {mp:,.2f}")
        c4.metric("Materialidad Ejecución (ME)", f"RD$ {me:,.2f}")
        
        st.markdown("---")
        
        # Pestañas de Navegación
        tab1, tab2, tab3 = st.tabs(["📋 Balanza de Comprobación", "🚨 Inconsistencias de Naturaleza", "🇩🇴 Diagnóstico Fiscal (Art. 287)"])
        
        with tab1:
            st.markdown("### Vista General de Cuentas")
            st.dataframe(df_balanza, use_container_width=True)
            
        with tab2:
            st.markdown("### Cuentas con Saldos Fuera de su Naturaleza Contable")
            df_errores = df_balanza[df_balanza['validacion_naturaleza'] != "Correcto"]
            if not df_errores.empty:
                st.error(f"Se detectaron {len(df_errores)} cuentas con saldos contrarios a su dinámica operativa contable.")
                st.dataframe(df_errores[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)
            else:
                st.success("✅ Excelente: No se han encontrado cuentas con inconsistencias en su saldo final.")
                
        with tab3:
            st.markdown("### Alertas de Fiscalización y Deducciones Admitidas (DGII)")
            df_fiscal = df_balanza[df_balanza['alerta_fiscal_rd'] != "Sin observaciones"]
            if not df_fiscal.empty:
                st.warning(f"Atención: Se identificaron {len(df_fiscal)} cuentas con exposición a revisión del Art. 287 del Código Tributario.")
                st.dataframe(df_fiscal[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal_rd']], use_container_width=True)
            else:
                st.success("✅ Cumplimiento Inicial: No se detectaron cuentas con las palabras de riesgo fiscal configuradas.")
else:
    st.info("👋 Por favor, carga tu archivo de Balanza de Comprobación para desplegar los cálculos automáticos.")
