import pandas as pd
import pyautogui
from pywinauto import Desktop
import time

TERCERO="tercero.png"
IMAGEN_VALIDAR = "boton_validar.png"
IMAGEN_YES = "boton_yes.png"
IMAGEN_NO = "boton_no.png"
IMAGEN_OK="boton_ok.png"
IMAGEN_NUEVO="boton_nuevo.png"
IMAGEN_TERCERO="campo_tercero.png"
CONFIDENCIA = 0.7
ESPERA_INICIAL = 3  # tiempo para preparar pantalla antes de comenzar
TIEMPO_ENTRE_CLICS = 1.0  # tiempo entre validar y aceptar

def enfocar_ventana(nombre_parcial):
    print(f"🔄 Buscando ventana que contenga: '{nombre_parcial}'")
    ventanas = Desktop(backend="uia").windows()

    for v in ventanas:
        if nombre_parcial.lower() in v.window_text().lower():
            print(f"✅ Ventana encontrada: {v.window_text()}. Activando...")
            v.set_focus()
            return True

    print("❌ No se encontró la ventana.")
    return False

# Leer los datos desde el Excel omitiendo los 4 primeros registros
df = pd.read_excel('apuntes1.xlsx')
print(df.columns.tolist())
# Esperar para que el usuario abra y enfoque la aplicación Sical
print("Coloca el cursor en el campo inicial del primer apunte. Tienes 3 segundos...")
time.sleep(0.5)
print("🔍 Buscando el texto 'Tercero'...")
# 1) Captura toda la pantalla
pantalla = pyautogui.screenshot()
# 2) Trata de localizar el rectángulo de la imagen en la pantalla
#    Devuelve (left, top, width, height) o None, sin excepción
rect = pyautogui.locate(TERCERO, pantalla, confidence=CONFIDENCIA)

for index, row in df.iterrows():
    if rect is not None:
        print(f"Procesando apunte {index + 1} de {len(df)}...")
        centro = pyautogui.center(rect)
        pyautogui.click(centro)
        # Pequeña pausa para asegurarnos de que el foco ya está en el campo
        time.sleep(0.1)
        # Codigo
        pyautogui.write(str(row['codigo']), interval=0.01)
        pyautogui.press('tab')
        time.sleep(0.1)
#        input("Pulsa ENTER para continuar...")
#        enfocar_ventana("SICAL II 4.2 mtec40")  # o el título visible de tu app
#        time.sleep(0.5)
        # Fecha
        pyautogui.write(str(row['Paso']), interval=0.01)
        pyautogui.press('tab')
        time.sleep(0.1)

        # Texto
        pyautogui.write(str(row['Concepto']), interval=0.01)
        pyautogui.press('tab')
        time.sleep(0.1)

        # Importe
        pyautogui.write(str(row['anualidad']), interval=0.01)
        pyautogui.press('enter')  # Asume que ENTER guarda el apunte, puedes cambiarlo a click si es necesario
        time.sleep(0.1)  # Espera antes del siguiente registro

        # Economica
        pyautogui.write(str(row['economica']), interval=0.1)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        time.sleep(0.1)

        # Importe
        pyautogui.write(str(row['importe']), interval=0.05)
        pyautogui.press('tab')
        time.sleep(0.1)

        #Ahora vamos a pulsar los botones para grabar los datos.
        # === PASO 1: Buscar y hacer clic en botón 'Validar' ===
        print(f"🔍 Buscando botón 'Validar' ({IMAGEN_VALIDAR})...")
        validar = pyautogui.locateOnScreen(IMAGEN_VALIDAR, confidence=CONFIDENCIA)

        if validar:
            centro_validar = pyautogui.center(validar)
            print(f"✅ Botón 'Validar' encontrado en {centro_validar}. Haciendo clic...")
            pyautogui.click(centro_validar)
            time.sleep(TIEMPO_ENTRE_CLICS)
        else:
            print("❌ No se encontró el botón 'Validar'. Verifica la imagen y que esté visible.")
            exit()

        # === PASO 2: Buscar y hacer clic en botón 'Aceptar' ===
        print(f"🔍 Buscando botón 'Aceptar' ({IMAGEN_YES})...")
        aceptar = pyautogui.locateOnScreen(IMAGEN_YES, confidence=CONFIDENCIA)

        if aceptar:
            centro_aceptar = pyautogui.center(aceptar)
            print(f"✅ Botón 'Aceptar' encontrado en {centro_aceptar}. Haciendo clic...")
            pyautogui.click(centro_aceptar)
        else:
            print("❌ No se encontró el botón 'Aceptar'. Asegúrate de que la ventana de confirmación esté abierta y visible.")

        # Mueve el ratón 200 píxeles a la izquierda de su posición actual
        x, y = pyautogui.position()
        pyautogui.moveTo(x - 200, y)
        time.sleep(0.5) 
            # === PASO 3: Aceptar la grabación ===
        print(f"🔍 Buscando botón 'Aceptar' ({IMAGEN_OK})...")
        ok = pyautogui.locateOnScreen(IMAGEN_OK, confidence=CONFIDENCIA)
        if ok:
            centro_ok = pyautogui.center(ok)
            print(f"✅ Botón 'Aceptar' encontrado en {centro_ok}. Haciendo clic...")
            pyautogui.click(centro_ok)
        else:
            print("❌ No se encontró el botón 'Ok'. Asegúrate de que la ventana de confirmación esté abierta y visible.")

        x, y = pyautogui.position()
        pyautogui.moveTo(x - 200, y)
        time.sleep(0.5) 

            # === PASO 3.5: No incorporamos documentacion ===
        print(f"🔍 Buscando botón 'Aceptar' ({IMAGEN_NO})...")
        ok = pyautogui.locateOnScreen(IMAGEN_NO, confidence=CONFIDENCIA)
        if ok:
            centro_ok = pyautogui.center(ok)
            print(f"✅ Botón 'No' encontrado en {centro_ok}. Haciendo clic...")
            pyautogui.click(centro_ok)
        else:
            print("❌ No se encontró el botón 'No'. Asegúrate de que la ventana de confirmación esté abierta y visible.")
#        input("Pulsa ENTER para continuar...")
        # === PASO 4: Vamos a grabar un registro NUEVO ===
        x, y = pyautogui.position()
        pyautogui.moveTo(x - 200, y)
        time.sleep(0.5)  # pequeña pausa para estabilizar imagen

        print(f"🔍 Buscando botón 'Aceptar' ({IMAGEN_NUEVO})...")
        nuevo = pyautogui.locateOnScreen(IMAGEN_NUEVO, confidence=CONFIDENCIA)

        if nuevo:
            centro_nuevo = pyautogui.center(nuevo)
            print(f"✅ Botón 'Aceptar' encontrado en {centro_nuevo}. Haciendo clic...")
            pyautogui.click(centro_nuevo)
        else:
            print("❌ No se encontró el botón 'Nuevo'. Asegúrate de que la ventana de confirmación esté abierta y visible.")
        x, y = pyautogui.position()
        pyautogui.moveTo(x - 200, y)
        time.sleep(0.5)  # pequeña pausa para estabilizar imagen

        # === PASO 5: Confirmamos que se puede borrar los datos de pantalla ===
        print(f"🔍 Buscando botón 'Aceptar' ({IMAGEN_YES})...")
        yes = pyautogui.locateOnScreen(IMAGEN_YES, confidence=CONFIDENCIA)

        if yes:
            centro_yes = pyautogui.center(yes)
            print(f"✅ Botón 'Aceptar' encontrado en {centro_yes}. Haciendo clic...")
            pyautogui.click(centro_yes)
        else:
            print("❌ No se encontró el botón 'Yes'. Asegúrate de que la ventana de confirmación esté abierta y visible.")
        print("Vamos a esperar y debe estar el cursor en Tercero.")        
#        input("Pulsa ENTER para continuar...")
        time.sleep(2.0)
        rect = pyautogui.locate(TERCERO, pantalla, confidence=CONFIDENCIA)

print("✅ Todos los apuntes han sido introducidos.")
