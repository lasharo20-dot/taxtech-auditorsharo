import streamlit as st
import pandas as pd

st.set_page_config(page_title="TaxTech Auditor RD", layout="wide")
st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")

# 1. PARAMETRIZACIÓN EN BARRA LATERAL
with st.sidebar:
    st.header("Configuración del Cliente")
    cliente = st.text_input("Nombre de la Empresa", value="Empresa de Prueba SRL")
    periodo = st.date_input("Período de Análisis")
    st.markdown("---")
    st.header("Parámetros de Materialidad (NIA 320)")
    tipo_empresa = st.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Pérdidas / Sin Fines de Lucro"])
    porcentaje_mat = st.slider("Porcentaje de Materialidad", 0.5, 2.0, 1.0, step=0.1)

st.subheader("1. Carga de Balanza de Comprobación")
archivo_balanza = st.file_uploader("Arrastra tu archivo Excel o CSV generado desde tu software contable (Odoo, QuickBooks, etc.)", type=["xlsx", "csv"])

if archivo_balanza is not None:
    # Lectura del archivo cargado
    try:
        if archivo_balanza.name.endswith('.csv'):
            df = pd.read_csv(archivo_balanza)
        else:
            df = pd.read_excel(archivo_balanza)
        
        st.success("Balanza cargada exitosamente.")
        
        # Limpieza estándar de nombres de columnas a minúsculas para evitar errores de sistemas
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Intentar detectar columnas numéricas de Débito y Crédito
        col_debito = [col for col in df.columns if 'deb' in col or 'debe' in col]
        col_credito = [col for col in df.columns if 'cred' in col or 'haber' in col]
        col_ingresos = [col for col in df.columns if 'ingre' in col or 'venta' in col]

        if col_debito and col_credito:
            total_debito = df[col_debito[0]].sum()
            total_credito = df[col_credito[0]].sum()
            diferencia = abs(total_debito - total_credito)
            
            # 2. INDICADORES VISUALES DE CUADRE CONTABLE
            st.markdown("### Resumen de Validación del Sistema")
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Total Débitos", f"RD$ {total_debito:,.2f}")
            kpi2.metric("Total Créditos", f"RD$ {total_credito:,.2f}")
            
            if diferencia < 0.01:
                kpi3.metric("Estado de Balanza", "✅ CUADRADA", f"Dif: RD$ {diferencia:,.2f}")
            else:
                kpi3.metric("Estado de Balanza", "❌ DESCUADRADA", f"Dif: RD$ {diferencia:,.2f}", delta_color="inverse")
            
            # 3. CÁLCULO DINÁMICO DE MATERIALIDAD (NIA 320)
            st.markdown("---")
            st.markdown("### Umbrales de Auditoría Calculados")
            
            # Si el usuario no tiene cuenta de ingresos definida, estimamos una base o usamos el total crédito/12 de forma referencial
            base_calculo = df[col_ingresos[0]].sum() if col_ingresos else total_credito * 0.30
            
            materialidad_planificacion = base_calculo * (porcentaje_mat / 100)
            materialidad_desempeno = materialidad_planificacion * 0.75
            
            mat_col1, mat_col2 = st.columns(2)
            mat_col1.info(f"**Materialidad de Planificación ({porcentaje_mat}%):** RD$ {materialidad_planificacion:,.2f}")
            mat_col2.warning(f"**Materialidad de Desempeño (75%):** RD$ {materialidad_desempeno:,.2f}")
            
        st.markdown("---")
        st.subheader("2. Vista de Datos Cargados")
        st.dataframe(df)
        
    except Exception as e:
        st.error(f"Error al procesar el archivo. Asegúrate de que las columnas tengan nombres claros como 'Debito', 'Credito' o 'Saldo'. Detalle: {e}")
