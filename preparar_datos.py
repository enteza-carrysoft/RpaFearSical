import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
from PIL import Image
import os
import time

# --- Configuración de la página y Título ---
st.set_page_config(page_title="Generador Asientos SICAL", layout="wide")

# --- Barra lateral para acciones ---
st.sidebar.title("Control")
if st.sidebar.button("🔴 Finalizar Aplicación"):
    st.sidebar.success("Cerrando la aplicación...")
    st.balloons()
    time.sleep(2)
    os._exit(0)


# Crear columnas para el encabezado
col1, col2 = st.columns([1, 8])

# Mostrar logo en la columna izquierda
with col1:
    # Es importante tener el archivo 'logo-dipu.png' en el mismo directorio
    # o proporcionar una ruta válida.
    try:
        logo = Image.open("logo-dipu.png")
        st.image(logo, width=100)
    except FileNotFoundError:
        st.warning("Logo no encontrado.")

# Título en la columna derecha
with col2:
    st.markdown("<h1 style='margin-top: 10px;'>Generador de Asientos para SICAL (Versión Mejorada)</h1>", unsafe_allow_html=True)

# --- Carga de Archivos ---
uploaded_file = st.file_uploader("📁 Cargar archivo .xls con los datos", type=["xls"])

# Cargar archivo resoluciones.csv (debe estar en el mismo directorio)
@st.cache_data
def cargar_resoluciones():
    """Carga el archivo de resoluciones desde un CSV."""
    try:
        df = pd.read_csv("resoluciones.csv", dtype=str)
        # Aseguramos que las columnas se llamen 'Codigo' y 'Texto' para el mapeo
        if len(df.columns) >= 2:
            df = df.iloc[:, :2] # Tomamos solo las dos primeras columnas
            df.columns = ["Codigo", "Texto"]
            return df
        else:
            return None
    except FileNotFoundError:
        return None

df_resoluciones = cargar_resoluciones()
if df_resoluciones is None:
    st.error("❌ No se encuentra el archivo 'resoluciones.csv' o no tiene el formato correcto (dos columnas).")
    st.stop()

# --- Datos Solicitados al Usuario ---
st.subheader("📝 Datos para la Generación del Archivo")
with st.container():
    col_datos1, col_datos2 = st.columns(2)
    with col_datos1:
        anualidad = st.text_input("Anualidad", help="Introduce el año de la anualidad. Ej: 2025")
        numero_cuenta = st.text_input("Número de cuenta")
    with col_datos2:
        codigo_economica = st.text_input("Código para Económica", value="821", help="Código que precede al año en la columna 'economica'.")
        numero_control = st.text_input("Número de control")

fecha_arqueo = st.date_input("Fecha de arqueo", value=datetime.date.today())


# --- Botón de Procesamiento y Lógica Principal ---
if st.button("🔄 Generar Archivo de Salida"):
    # Validaciones de entrada
    if uploaded_file is None:
        st.warning("⚠️ Por favor, sube el archivo .xls.")
    elif not all([numero_cuenta, numero_control, anualidad, codigo_economica]):
        st.warning("⚠️ Completa todos los campos de datos.")
    else:
        try:
            # Leer archivo Excel, omitiendo las tres primeras filas (cabecera)
            df = pd.read_excel(uploaded_file, dtype=str, skiprows=3)

            # Validar que las columnas necesarias del archivo de entrada existan
            columnas_necesarias = ["Clave Objeto", "Código", "Nombre", "Importe"]
            if not all(col in df.columns for col in columnas_necesarias):
                st.error(f"❌ El archivo de entrada debe tener las columnas: {columnas_necesarias}")
                st.stop()

            # --- INICIO: Lógica de transformación de datos ---

            # 1. Extraer el año del FEAR desde 'Clave Objeto'
            # Se asume que el año son los 4 dígitos después del primer guion.
            df['año_fear'] = df['Clave Objeto'].str.split('-', n=1).str[1].str.slice(0, 4)

            # 2. Crear la columna 'economica'
            df['economica'] = codigo_economica + df['año_fear']

            # 3. Crear la columna 'anualidad'
            df['anualidad'] = anualidad

            # 4. Crear la columna 'Paso'
            df['Paso'] = "1"

            # 5. Mapear el texto del concepto (lógica original)
            df["ClaveExtraida"] = df["Clave Objeto"].str.extract(r'(-20.*)$')[0].str.replace('-', '', n=1)
            df["ClaveExtraida"] = df["ClaveExtraida"].str.strip()
            df_resoluciones["Codigo"] = df_resoluciones["Codigo"].str.strip()
            df["Concepto"] = df["ClaveExtraida"].map(df_resoluciones.set_index("Codigo")["Texto"]).fillna("Texto no encontrado")

            # --- FIN: Lógica de transformación de datos ---

            # Renombrar columnas para que coincidan con la salida deseada
            df.rename(columns={
                "Código": "Codigo",
                "Nombre": "Nombre",
                "Importe": "importe",
                "Clave Objeto": "Codigo FEAR"
            }, inplace=True)

            # Seleccionar y ordenar las columnas para el archivo final
            columnas_finales = [
                "Codigo", "Nombre", "importe", "Codigo FEAR",
                "Concepto", "anualidad", "economica", "Paso"
            ]
            df_final = df[columnas_finales]

            # --- Generación del archivo Excel de salida ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Arqueo")

            output.seek(0)
            st.success("✅ ¡Archivo de apuntes generado correctamente!")
            st.download_button(
                label="📥 Descargar Archivo .xlsx",
                data=output,
                file_name=f"Apuntes_Generados_{fecha_arqueo.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Ocurrió un error durante el procesamiento: {str(e)}")
            st.error("Asegúrate de que el archivo .xls tiene el formato correcto y las columnas necesarias.")