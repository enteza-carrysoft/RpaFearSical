import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

#st.set_page_config(page_title="Generador Apuntes SICAL", layout="centered")
#st.title("📄 Generador de Apuntes para SICAL")

from PIL import Image
st.set_page_config(page_title="Generador Asientos SICAL", layout="wide")

# Crear columnas para el encabezado
col1, col2 = st.columns([1, 8])

# Mostrar logo en la columna izquierda
with col1:
    logo = Image.open("logo-dipu.png")
    st.image(logo, width=100)

# Título en la columna derecha
with col2:
    st.markdown("<h1 style='margin-top: 10px;'>Generador de Asientos para SICAL</h1>", unsafe_allow_html=True)

# Subida del archivo principal .xls
uploaded_file = st.file_uploader("📁 Cargar archivo .xls con los datos", type=["xls"])

# Cargar archivo resoluciones.csv (debe estar en el mismo directorio)
@st.cache_data
def cargar_resoluciones():
    try:
        df = pd.read_csv("resoluciones.csv", dtype=str)
        df.columns = ["Codigo", "Texto"]
        return df
    except FileNotFoundError:
        return None

df_resoluciones = cargar_resoluciones()
if df_resoluciones is None:
    st.error("❌ No se encuentra el archivo 'resoluciones.csv' en el directorio.")
    st.stop()

# Datos solicitados al usuario
st.subheader("📝 Datos de cabecera")
numero_cuenta = st.text_input("Número de cuenta")
fecha_arqueo = st.date_input("Fecha de arqueo", value=datetime.date.today())
numero_control = st.text_input("Número de control")

# Botón para procesar
if st.button("🔄 Generar archivo"):
    if uploaded_file is None:
        st.warning("⚠️ Por favor, sube el archivo .xls.")
    elif not numero_cuenta or not numero_control:
        st.warning("⚠️ Completa todos los campos.")
    else:
        try:
            # Leer archivo Excel, omitiendo las dos primeras filas (cabecera)
            df = pd.read_excel(uploaded_file, dtype=str, skiprows=3)

            # Validar columnas necesarias
            columnas_necesarias = ["Clave Objeto", "Código", "Nombre", "Importe"]
            if not all(col in df.columns for col in columnas_necesarias):
                st.error(f"❌ El archivo debe tener las columnas: {columnas_necesarias}")
                st.stop()

            # Extraer subcadena de 'Clave objeto'
            df["ClaveExtraida"] = df["Clave Objeto"].str.extract(r'(-20.*)$')[0].str.replace('-', '', n=1)

            # Limpiar espacios en los códigos de ambos DataFrames antes de hacer el mapeo
            df["ClaveExtraida"] = df["ClaveExtraida"].str.strip()
            df_resoluciones["Codigo"] = df_resoluciones["Codigo"].str.strip()
            df_resoluciones["Texto"] = df_resoluciones["Texto"].str.strip()

            # Realizar el mapeo limpio
            df["Texto"] = df["ClaveExtraida"].map(df_resoluciones.set_index("Codigo")["Texto"]).fillna("Texto no encontrado")

            # Preparar resultado final
            df_final = df[["Código", "Nombre", "Importe", "Clave Objeto","Texto"]]

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                meta = pd.DataFrame({
                    "Número de cuenta": [numero_cuenta],
                    "Fecha arqueo": [fecha_arqueo.strftime("%Y-%m-%d")],
                    "Número de control": [numero_control]
                })
                meta.to_excel(writer, index=False, sheet_name="Arqueo", startrow=0)
                df_final.to_excel(writer, index=False, sheet_name="Arqueo", startrow=3)

            output.seek(0)
            st.success("✅ Archivo generado correctamente.")
            st.download_button(
                label="📥 Descargar archivo",
                data=output,
                file_name=f"Arqueo_{fecha_arqueo.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
