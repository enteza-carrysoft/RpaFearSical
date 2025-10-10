import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
from PIL import Image
import os
import time
import pyautogui, pyperclip

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.02   # pausa global mínima entre acciones

TYPE_IV   = 0.005        # write interval
PRESS_IV  = 0.04         # entre pulsaciones repetidas (tab x3)
AFTER_TAB = 0.06         # pequeña espera tras tab
AFTER_ENT = 0.20         # espera tras enter (validaciones)

# --- Configuración de la página ---
st.set_page_config(page_title="Asistente SICAL", layout="wide")

# --- Constantes y Ficheros Necesarios ---
NOMBRE_FICHERO_RPA = "APUNTES_PARA_RPA.xlsx"
IMAGENES_RPA = [
    "./images/tercero.png", "./images/boton_validar.png", "./images/boton_yes.png", "./images/boton_no.png",
    "./images/boton_ok.png", "./images/boton_nuevo.png"
]

# --- Funciones Auxiliares ---

@st.cache_data
def cargar_resoluciones():
    """Carga el archivo de resoluciones desde un CSV, manejando codificación, nombres y BOM."""
    try:
        df = pd.read_csv("resoluciones.csv", dtype=str, encoding="utf-8-sig")  # elimina BOM si existe
        df.columns = df.columns.str.strip().str.capitalize()  # convierte 'codigo' -> 'Codigo'

        if "Codigo" not in df.columns or "Texto" not in df.columns:
            st.error("❌ Las columnas 'Codigo' y 'Texto' no se han encontrado correctamente.")
            st.stop()

        df = df[["Codigo", "Texto"]]
        return df
    except FileNotFoundError:
        st.error("❌ El archivo 'resoluciones.csv' no se encuentra.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar 'resoluciones.csv': {e}")
        return None

def write_fast(text):
    s = "" if text is None else str(text).strip()
    # pegar si es largo o con espacios complejos
    if len(s) >= 15:
        pyperclip.copy(s)
        pyautogui.hotkey('ctrl','v')
    else:
        pyautogui.write(s, interval=TYPE_IV)


def ejecutar_rpa_corregido(placeholder):
    """
    Función RPA con la lógica corregida para buscar el campo 'Tercero' en cada iteración.
    """
    # Nombres de las imágenes
    TERCERO_IMG = "./images/tercero.png"
    VALIDAR_IMG = "./images/boton_validar.png"
    YES_IMG = "./images/boton_yes.png"
    NO_IMG = "./images/boton_no.png"
    OK_IMG = "./images/boton_ok.png"
    NUEVO_IMG = "./images/boton_nuevo.png"
    
    CONFIDENCIA = 0.8
    TIEMPO_ENTRE_CLICS = 1.2

    try:
        placeholder.info("🤖 RPA Iniciado. Leyendo fichero de apuntes...")
        df = pd.read_excel(NOMBRE_FICHERO_RPA)
        time.sleep(1)

        placeholder.warning("⏳ El operario tiene 5 segundos para asegurarse de que la ventana de SICAL está visible y en primer plano...")
        time.sleep(5)

        for index, row in df.iterrows():
            placeholder.info(f"▶️ Procesando apunte {index + 1} de {len(df)}: {row['Nombre']}")
            
            placeholder.info("... 🔍 Buscando el campo 'Tercero' para iniciar el apunte.")
            try:
                campo_tercero_loc = pyautogui.locateCenterOnScreen(TERCERO_IMG, confidence=CONFIDENCIA)
                if campo_tercero_loc is None:
                    raise pyautogui.ImageNotFoundException
                pyautogui.click(campo_tercero_loc)
                time.sleep(0.1)
            except pyautogui.ImageNotFoundException:
                placeholder.error(f"❌ Error Crítico: No se pudo encontrar el campo 'Tercero' ({TERCERO_IMG}) en la pantalla. Proceso detenido.")
                return


#            pyautogui.write(str(row['Codigo']), interval=0.00)
#            pyautogui.press('tab')
#            time.sleep(0.1)
#            pyautogui.write(str(row['Paso']), interval=0.00)
#            pyautogui.press('tab')
#            time.sleep(0.1)
#            pyautogui.write(str(row['Concepto']), interval=0.00)
#            pyautogui.press('tab')
#            time.sleep(0.1)
#            pyautogui.write(str(row['anualidad']), interval=0.00)
#            pyautogui.press('enter')
#            time.sleep(0.1)
#            pyautogui.write(str(row['economica']), interval=0.00)
#            pyautogui.press('tab', presses=3, interval=0.00)
#            time.sleep(0.1)
#            pyautogui.write(str(row['importe']), interval=0.00)
#            pyautogui.press('tab')

#           Código mejorado para darle mas rapidez a la entrada de datos en el formulario delphi.
            write_fast(row['Codigo'])
            pyautogui.press('tab'); time.sleep(AFTER_TAB)

            write_fast(row['Paso'])
            pyautogui.press('tab'); time.sleep(AFTER_TAB)

            write_fast(row['Concepto'])
            pyautogui.press('tab'); time.sleep(AFTER_TAB)

            write_fast(row['anualidad'])
            pyautogui.press('enter'); time.sleep(AFTER_ENT)

            write_fast(row['economica'])
            pyautogui.press('tab', presses=3, interval=PRESS_IV); time.sleep(AFTER_TAB)

            # Si 'importe' necesita coma decimal española:
            importe = str(row['importe']).replace('.', ',')
            write_fast(importe)
            pyautogui.press('tab'); time.sleep(AFTER_TAB)
            time.sleep(0.1)

            placeholder.info("... Pulsando secuencia de botones para guardar.")
            pyautogui.click(pyautogui.locateCenterOnScreen(VALIDAR_IMG, confidence=CONFIDENCIA))
            time.sleep(TIEMPO_ENTRE_CLICS)            
            pyautogui.click(pyautogui.locateCenterOnScreen(YES_IMG, confidence=CONFIDENCIA))
            time.sleep(TIEMPO_ENTRE_CLICS)
            pyautogui.click(pyautogui.locateCenterOnScreen(OK_IMG, confidence=CONFIDENCIA))
            time.sleep(TIEMPO_ENTRE_CLICS)
            pyautogui.click(pyautogui.locateCenterOnScreen(NO_IMG, confidence=CONFIDENCIA))
            time.sleep(TIEMPO_ENTRE_CLICS)
            
            if index < len(df) - 1:
                placeholder.info("... Preparando para el siguiente apunte.")
                pyautogui.click(pyautogui.locateCenterOnScreen(NUEVO_IMG, confidence=CONFIDENCIA))
                time.sleep(TIEMPO_ENTRE_CLICS)
                pyautogui.click(pyautogui.locateCenterOnScreen(YES_IMG, confidence=CONFIDENCIA))
                time.sleep(1)

        placeholder.success("✅ ¡Proceso RPA finalizado! Todos los apuntes han sido introducidos.")
        st.balloons()

    except FileNotFoundError:
        placeholder.error(f"❌ Error Crítico: No se encuentra el archivo '{NOMBRE_FICHERO_RPA}'. Por favor, genéralo primero en la Pestaña 1.")
    except pyautogui.ImageNotFoundException as e:
        placeholder.error(f"❌ Error de RPA: No se pudo encontrar una imagen en pantalla. Detalle: {e}")
    except Exception as e:
        placeholder.error(f"❌ Ha ocurrido un error inesperado durante el RPA: {e}")

# --- Interfaz Principal ---

col1, col2 = st.columns([1, 8])
with col1:
    try:
        logo = Image.open("./images/logo-dipu.png")
        st.image(logo, width=100)
    except FileNotFoundError:
        st.warning("Logo no encontrado.")
with col2:
    st.markdown("<h1 style='margin-top: 10px;'>Asistente para Generación de Asientos en SICAL</h1>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["1. Generar Archivo de Apuntes", "2. Automatizar Carga con RPA"])

with tab1:
    st.header("Paso 1: Preparar la hoja de cálculo")
    st.markdown("Carga los ficheros y completa los datos para generar el archivo que usará el robot.")

    uploaded_file = st.file_uploader("📁 Cargar archivo .xls con los datos de origen", type=["xls"])
    archivo_anterior = st.file_uploader("📁 (Opcional) Cargar archivo del mes anterior para comparar importes", type=["xlsx"])

    df_resoluciones = cargar_resoluciones()

    if df_resoluciones is None:
        st.error("❌ No se encuentra el archivo 'resoluciones.csv'. Es necesario para la conversión.")
        st.stop()

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

    if st.button("🔄 Generar Archivo de Salida"):
        if not all([uploaded_file, numero_cuenta, numero_control, anualidad, codigo_economica]):
            st.warning("⚠️ Completa todos los campos y carga el archivo .xls de origen.")
        else:
            try:
                # --- Lógica principal de procesamiento ---
                df = pd.read_excel(uploaded_file, dtype=str, skiprows=1)
                columnas_necesarias = ["Clave Objeto", "Código", "Nombre", "Importe"]
                if not all(col in df.columns for col in columnas_necesarias):
                    st.error(f"❌ El archivo de entrada debe tener las columnas: {columnas_necesarias}")
                    st.stop()

                df['año_fear'] = df['Clave Objeto'].str.split('-', n=1).str[1].str.slice(0, 4)
                df['economica'] = codigo_economica + df['año_fear']
                df['anualidad'] = anualidad
                df['Paso'] = "1"
                df["ClaveExtraida"] = df["Clave Objeto"].str.extract(r'(-20.*)$')[0].str.replace('-', '', n=1)
                df["ClaveExtraida"] = df["ClaveExtraida"].str.strip()
# --- LIMPIEZA Y MAPEO ROBUSTO ---
                df["ClaveExtraida"] = df["ClaveExtraida"].str.strip()
                df_resoluciones["Codigo"] = df_resoluciones["Codigo"].astype(str).str.strip()
                df_resoluciones["Texto"] = df_resoluciones["Texto"].astype(str).str.strip()

# Detectar y avisar si hay códigos duplicados en resoluciones.csv
                if df_resoluciones["Codigo"].duplicated().any():
                    duplicados = df_resoluciones[df_resoluciones["Codigo"].duplicated(keep=False)]["Codigo"].unique().tolist()
                    st.warning(f"⚠️ Se encontraron códigos duplicados en resoluciones.csv: {', '.join(duplicados[:10])} "
                            f"{'(y más...)' if len(duplicados) > 10 else ''}. Se usará solo la primera aparición de cada uno.")
                    df_resoluciones = df_resoluciones.drop_duplicates(subset="Codigo", keep="first")

# Realizar el mapeo limpio y seguro
                mapa_resoluciones = df_resoluciones.set_index("Codigo")["Texto"]
                df["Concepto"] = df["ClaveExtraida"].map(mapa_resoluciones).fillna("Texto no encontrado")              
                df.rename(columns={"Código": "Codigo", "Nombre": "Nombre", "Importe": "importe", "Clave Objeto": "Codigo FEAR"}, inplace=True)
                
                columnas_finales = ["Codigo", "Nombre", "importe", "Codigo FEAR", "Concepto", "anualidad", "economica", "Paso"]
                df_final = df[columnas_finales].copy()

                # --- INICIO: Lógica de comparación y coloreado REVISADA ---
                if archivo_anterior:
                    st.info("Comparando importes con el archivo del mes anterior...")
                    df_anterior = pd.read_excel(archivo_anterior)
                    
                    if 'Codigo FEAR' not in df_anterior.columns:
                        st.error("El archivo del mes anterior no contiene la columna 'Codigo FEAR', necesaria para la comparación.")
                        st.stop()

                    # FIX: Handle potential duplicates in the key column of the previous month's file
                    if df_anterior['Codigo FEAR'].duplicated().any():
                        st.warning("Se han encontrado valores 'Codigo FEAR' duplicados en el archivo del mes anterior. Se usará solo la primera aparición de cada código para la comparación.")
                        df_anterior.drop_duplicates(subset=['Codigo FEAR'], keep='first', inplace=True)

                    df_final['importe_numeric'] = pd.to_numeric(df_final['importe'], errors='coerce')
                    df_anterior['importe_numeric'] = pd.to_numeric(df_anterior['importe'], errors='coerce')

                    mapa_importes_anterior = df_anterior.set_index('Codigo FEAR')['importe_numeric']
                    df_final['importe_anterior'] = df_final['Codigo FEAR'].map(mapa_importes_anterior)

                    df_final['diferencia'] = (df_final['importe_anterior'].notna()) & \
                                           (df_final['importe_numeric'] != df_final['importe_anterior'])
                    
                    num_diferencias = df_final['diferencia'].sum()
                    if num_diferencias > 0:
                        st.warning(f"Se han encontrado {num_diferencias} registros con importes diferentes al mes anterior. Se marcarán en el Excel.")
                    else:
                        st.success("No se encontraron diferencias de importes con el mes anterior.")

                    def highlight_diff(row):
                        color = 'background-color: #FFDDC1'
                        return [color if row.diferencia else '' for _ in row]

                    df_styled = df_final.style.apply(highlight_diff, axis=1)
                    
                    df_styled.to_excel(NOMBRE_FICHERO_RPA, index=False, engine='openpyxl', columns=columnas_finales)
                    
                    output = BytesIO()
                    df_styled.to_excel(output, index=False, engine='openpyxl', columns=columnas_finales)
                    output.seek(0)
                
                else:
                    df_final.to_excel(NOMBRE_FICHERO_RPA, index=False)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        df_final.to_excel(writer, index=False, sheet_name="Arqueo")
                    output.seek(0)
                # --- FIN: Lógica REVISADA ---

                st.success(f"✅ ¡Archivo generado y guardado como '{NOMBRE_FICHERO_RPA}'!")
                st.download_button(
                    label="📥 Descargar Copia del Archivo .xlsx",
                    data=output,
                    file_name=f"Apuntes_Generados_{fecha_arqueo.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Ocurrió un error durante la generación del archivo: {str(e)}")

with tab2:
    st.header("Paso 2: Cargar los apuntes en SICAL con el Robot")
    st.markdown("---")
    st.info("""
    **Instrucciones para el Operario:**

    1.  Asegúrate de que el archivo `APUNTES_PARA_RPA.xlsx` ha sido generado en la **Pestaña 1**.
    2.  **Revisa el archivo Excel**. Si hay filas coloreadas, verifica que las diferencias de importe son correctas.
    3.  Abre la aplicación **SICAL** y déjala en primer plano, totalmente visible.
    4.  Cuando todo esté listo, pulsa el siguiente botón y no muevas el ratón ni el teclado.
    """)

    rpa_status_placeholder = st.empty()

    if st.button("🤖 Iniciar Automatización (RPA)"):
        imagenes_faltantes = [img for img in IMAGENES_RPA if not os.path.exists(img)]
        if imagenes_faltantes:
            st.error(f"❌ Faltan archivos de imagen para el RPA: {', '.join(imagenes_faltantes)}.")
        else:
            ejecutar_rpa_corregido(rpa_status_placeholder)

st.sidebar.title("Control")
if st.sidebar.button("🔴 Finalizar Aplicación"):
    st.sidebar.success("Cerrando la aplicación...")
    st.balloons()
    time.sleep(2)
    os._exit(0)
