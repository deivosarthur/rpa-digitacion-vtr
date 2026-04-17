# Proyecto : RPA "TG Bots"
# Propietario : Technogroup SPA
# Creador: Jose Huaiquiman Arce
# modificacion: Adolfo Puentes
# Modificacion: 2024-06-17 - Conexion de base de datos e integracion al bot

from db_connection import (
    obtener_ordenes,
    obtener_materiales_por_ot,
    actualizar_estado,
    tomar_orden,
)  # Importa las funciones necesarias para interactuar con la base de datos

from selenium import webdriver
import time  # Ejercicio 1 , libreria nativa de python, para controlar el tiempo de espera del time.sleep()
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unittest
from selenium.webdriver.support.ui import Select
from selenium.webdriver import (
    ActionChains,
)  # permite importar la clase ActionChains, esta clase permite automatizar interacciones de usuario más complejas en una página web
from selenium.webdriver.chrome.options import (
    Options,
)  # permite personalizar el comportamiento del navegador Chrome antes de que se inicie

import os  # Permite ver que carpeta esta direccionada la busqueda de un archivo
import csv  # para manejo de archivos csv
import re  # ¡Añade esta línea!
import sys
import json
from pathlib import Path


def esperar_modal_cerrado(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.invisibility_of_element_located((By.ID, "alert"))
        )
    except:
        pass

    try:
        WebDriverWait(driver, 5).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "ui-widget-overlay"))
        )
    except:
        pass


# from selenium.common.exceptions import TimeoutException
# from playsound import playsound # primero se ejecuta (pip install pyglet) y luego (pip install playsound)
# print(os.getcwd()) # Muestra la ruta de archivos adicionales

monitor = "1"
archivolog = f"./Digitacion/log_ejecucion_monitor{monitor}.csv"
statusorden = f"./Digitacion/status_orden_de_trabajo.csv"
os.makedirs(os.path.dirname(archivolog), exist_ok=True)


class usando_unittest(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.set_window_rect(x=450, y=0, width=800, height=800)
        self.driver.implicitly_wait(
            10
        )  # Espera implícita para evitar errores de elementos no encontrados.

    def test_a_run_1_de_2(self):
        driver = self.driver
        driver.get("http://eps.vtr.cl/eps/declaracion_consumo.do#")
        driver.execute_script("document.body.style.zoom='65%'")
        with open(
            archivolog, "a", newline=""
        ) as csvfile:  # 'w' (write): Borra el contenido existente y escribe desde cero. 'a' (append): Agrega las nuevas líneas al final del archivo.
            writer = csv.writer(csvfile)
            writer.writerow(
                ["-------------------", "--------------", "DECLARACION X BOTS (1/2)"]
            )
            writer.writerow(["Timestamp", "Accion", "Detalle"])
            print(f"NDC{monitor} -> Abriendo navegador en la pagina de inicio")
            sys.stdout.flush()  # permite enviar el print al frame del home
            writer.writerow(
                [
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                    "Navegacion",
                    "TG Bots -> Pagina de inicio cargada",
                ]
            )

        try:  # Este bloque permite al RPA esperar a que el objeto o componente termine de cargar.
            if getattr(sys, "frozen", False):
                # Si es un ejecutable, la base es la carpeta donde está el archivo .exe
                base_dir = Path(sys.executable).resolve().parent
            else:
                # Si es un script .py, la base es la carpeta donde está el archivo .py
                base_dir = Path(__file__).resolve().parent

            # --- PASO NUEVO: Cargar credenciales desde el archivo ---
            keys = base_dir / "keys.txt"
            with open(keys, "r") as f:
                credentials = json.load(f)

            user_val = credentials["user"]
            pass_val = credentials["pass"]
            # -------------------------------------------------------

            usuario_crecencial = driver.find_element(By.XPATH, f"//*[@id='username']")
            usuario_crecencial.send_keys(user_val)
            time.sleep(2)

            clave_crecencial = driver.find_element(By.XPATH, f"//*[@id='password']")
            clave_crecencial.send_keys(pass_val)
            time.sleep(2)

            btn_crecencial = driver.find_element(
                By.XPATH, f"//*[@id='btn_valida_ingreso']"
            )
            btn_crecencial.click()
            driver.execute_script("document.body.style.zoom='65%'")
            print(f"NDC{monitor} -> Credenciales y Acceso ok")
            sys.stdout.flush()
            with open(archivolog, "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(
                    [
                        time.strftime("%Y-%m-%d %H:%M:%S"),
                        "Credenciales",
                        "NDC -> Credenciales y Acceso ok",
                    ]
                )
                writer.writerow(
                    [
                        "-------------------",
                        "--------------",
                        "--------------------------------------------------------------------------------",
                    ]
                )
            time.sleep(10)  # Espera para que cargue bien la pagina

            # ------- Importación de datos desde excel DATA.xlsx (modificacion)
            #
            try:
                print(f"NDC{monitor} -> Extracción de datos desde sql")
                sys.stdout.flush()
                driver.execute_script("document.body.style.zoom='65%'")

                # 2. Recorre el diccionario y extrae los datos de las otras columnas
                ordenes = obtener_ordenes()

                for orden in ordenes:
                    id_orden = orden.id
                    ot = orden.OT

                    # INTENTAR TOMAR LA OT
                    tomada = tomar_orden(ot, f"BOT_{monitor}")
                    print(f"NDC{monitor} -> Intentando tomar OT {ot}")
                    if not tomada:
                        print(f"NDC{monitor} -> OT {ot} ya fue tomada por otro bot")
                        continue

                    print(f"NDC{monitor} -> OT {ot} tomada correctamente")
                    #
                    # -------  OT Busqueda -------
                    #
                    folio_ot = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "numero_ot"))
                    )

                    driver.execute_script("arguments[0].click();", folio_ot)
                    folio_ot.clear()
                    folio_ot.send_keys(ot)

                    # print("NDC{monitor} -> registro de orden de trabajo ok")
                    time.sleep(1)

                    wait = WebDriverWait(
                        driver, 8
                    )  # Espera un máximo de 15 segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                    btn_buscar = wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//*[@id='boton-buscar']/span")
                        )
                    )

                    btn_buscar.click()
                    esperar_modal_cerrado(driver)  # 👈 FALTA ESTO

                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[@id='trabajo-material']")
                        )
                    )

                    # VALIDACIÓN NUEVA (VA AQUÍ)
                    try:
                        WebDriverWait(driver, 3).until(
                            EC.text_to_be_present_in_element(
                                (By.TAG_NAME, "body"),
                                "La orden no existe o no está finalizada",
                            )
                        )

                        print(f"NDC{monitor} -> OT {ot} no existe o no finalizada")

                        

                        actualizar_estado(
                            ot, "Procesado", "NO_EXISTE"
                        )  # estado de revision
                        # saltar a la siguiente OT
                        continue

                    except:
                        pass
                    #
                    # -------  OT Validación trabajo no asociado -------
                    #
                    espera_alerta = WebDriverWait(driver, 2)

                    # Verifica si la alerta está presente
                    try:
                        espera_alerta.until(
                            EC.text_to_be_present_in_element(
                                (By.XPATH, "//*[@id='alert']"),
                                "Existe actividad en la OT no asociada a un trabajo",
                            )
                        )
                        # Si la alerta existe, busca y haz clic en el botón
                        btn_alerta = driver.find_element(
                            By.XPATH, "/html/body/div[8]/div[3]/div/button/span"
                        )
                        btn_alerta.click()
                        print(
                            f"NDC{monitor} -> Alerta 'Actividad en la OT no asociada' detectada y aceptada."
                        )
                        alerta_Ot_no_asociada = False
                    except:
                        alerta_Ot_no_asociada = True
                        # Si la alerta no aparece en 2 segundos, no hace nada y el código continúa
                        # print(f"NDC{monitor} -> Sin alerta, continua el proceso.")

                    if alerta_Ot_no_asociada:
                        #
                        # -------  OT Validación trabajo ya declarado -------
                        #
                        
                        estado_detectado = False
                        try:
                            WebDriverWait(driver, 6).until(
                                EC.any_of(
                                    EC.text_to_be_present_in_element(
                                        (
                                            By.XPATH,
                                            "//*[@id='div_declaraciones_realizadas']/table/tbody/tr[3]/td[6]",
                                        ),"Declarada",
                                    ),
                                    EC.text_to_be_present_in_element(
                                        (By.XPATH,"//*[@id='div_declaraciones_realizadas']/table/tbody/tr[3]/td[6]",),"En proceso",
                                    ),
                                )
                            )
                            print(
                                f"NDC{monitor} -> Error - Orden fue declarada con anterioridad."
                            )
                            actualizar_estado(
                                ot, "Procesado", "YA_DECLARADA"
                            )  # estado de revision
                            estado_detectado = True
                            
                        except:
                            print(f"NDC{monitor} -> No estaba declarada, continuar flujo normal")

                        if estado_detectado:
                            continue
                        seguir_declaracion = False
                        saltar_btn_procesar = False  
                        try:
                            wait = WebDriverWait(driver, 6)
                            seleccion_trabajo = wait.until(
                                EC.presence_of_element_located(
                                    (By.XPATH, "//*[@id='trabajo-material']")
                                )
                            )
                            seguir_declaracion = True
                            dropdown = Select(
                                seleccion_trabajo
                            )  # Create a Select object for the dropdown
                            dropdown.select_by_index(1)

                            selected_option_text = (
                                dropdown.first_selected_option.text
                            )  # Captura el dato visible del select
                            if (
                                selected_option_text
                                == "Access Points Anexo WIFI (Extensor)"
                                or selected_option_text == "Access Points Anexo"
                            ):
                                # if selected_option_text == "Access Points Anexo WIFI (Extensor)" :
                                print(
                                    f"NDC{monitor} -> 'Access Point Anexo WIFI (Extensor)' se mueve a la siguente opcion."
                                )
                                dropdown.select_by_index(
                                    2
                                )  # Pasa a la siguiente opción
                                time.sleep(2)
                            else:
                                # print(f"NDC{monitor} -> Trabajo seleccionado.")
                                time.sleep(2)
                            
                        except:
                            print(
                                f"NDC{monitor} -> Error - Actividad no termino de cargar o existe un problema con la actividad"
                            )
                            sys.stdout.flush()
                            print("-" * 50)  # Separador para mejor lectura
                            sys.stdout.flush()
                            with open(archivolog, "a", newline="") as csvfile:
                                writer = csv.writer(csvfile)
                                writer.writerow(
                                    [
                                        time.strftime("%Y-%m-%d %H:%M:%S"),
                                        " CierreOrden ",
                                        " NDC -> NO Declarada - Actividad no termino de cargar o existe un problema con la actividad",
                                    ]
                                )
                                writer.writerow(
                                    [
                                        "-------------------",
                                        "--------------",
                                        "--------------------------------------------------------------------------------",
                                    ]
                                )
                                seguir_declaracion = False
                                actualizar_estado(
                                    ot, "Procesado", "ERROR_ACTIVIDAD"
                                )  # estado de revision
                                continue
                            

                            #
                            # --- Bloque de Carga de material
                        
                        
                        time.sleep(8)
                        if seguir_declaracion:
                            materiales = obtener_materiales_por_ot(ot)

                            for material in materiales:
                                producto = material.Codigo_a_rebajar
                                cantidad = str(material.Declarado)

                                ingreso_producto = driver.find_element(
                                    By.XPATH, "//*[@id='in_codigo']"
                                )

                                time.sleep(2)

                                ingreso_producto.clear()  # Limpia el campo antes de ingresar el nuevo valor
                                ingreso_producto.send_keys(producto)

                                wait = WebDriverWait(driver, 10)
                                wait.until(
                                    EC.presence_of_element_located(
                                        (By.XPATH, "//*[@id='id_material']")
                                    )
                                )

                                ingreso_cantidad = driver.find_element(
                                    By.XPATH, "//*[@id='cantidad']"
                                )
                                ingreso_cantidad.send_keys(cantidad)
                                ingreso_cantidad.click()

                                btn_agregar_material = driver.find_element(
                                    By.XPATH, "//*[@id='ingresar_nuevo']/span"
                                )
                                btn_agregar_material.click()
                                saltar_btn_procesar = True
                                time.sleep(1)
                                    
                                try:
                                    wait = WebDriverWait(
                                        driver, 1
                                    )  # Espera un máximo de 1x segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                                    btn_buscar = wait.until(
                                        EC.element_to_be_clickable(
                                            (By.XPATH,"/html/body/div[8]/div[3]/div/button/span")
                                        )
                                    )
                                    btn_buscar.click()
                                    wait = WebDriverWait(driver, 1)  # Espera un máximo de x segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                                    btn_buscar = wait.until(
                                        EC.element_to_be_clickable(
                                            (By.XPATH, "//*[@id='in_codigo']")
                                        )
                                    )
                                    btn_buscar.clear()
                                    print(f"NDC{monitor}-> Material Repetido")

                                except:

                                    time.sleep(1)

                            
                            if (
                                saltar_btn_procesar
                            ):  # si es True ejecuta el bloque de procesar declaración
                                btn_procesar_material = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable(
                                        (
                                            By.XPATH,
                                            "//*[@id='btn_procesar_declaracion']/span",
                                        )
                                    )
                                )

                                try:
                                    WebDriverWait(driver, 10).until_not(
                                        EC.presence_of_element_located((By.CLASS_NAME, "ui-widget-overlay"))
                                    )
                                except:
                                    print(f"NDC{monitor} -> Overlay no desapareció, continúo igual")

                                driver.execute_script("arguments[0].click();", btn_procesar_material)

                                # 1. Esperar modal de validación
                                modal_validacion = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located(
                                        (By.ID, "form_validacion_declaracion")
                                    )
                                )

                                print(f"NDC{monitor} -> Modal de validación cargado")

                                # 2. Click en botón PROCESAR del modal
                                # 🔥 esperar que desaparezca overlay (CRÍTICO)
                                try:
                                    WebDriverWait(driver, 10).until_not(
                                        EC.presence_of_element_located((By.CLASS_NAME, "ui-widget-overlay"))
                                    )
                                except:
                                    print(f"NDC{monitor} -> Overlay no desapareció, continúo igual")

                                boton_procesar_modal = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.XPATH, "//span[text()='Procesar']/.."))
                                )

                                driver.execute_script("arguments[0].click();", boton_procesar_modal)
                                time.sleep(1.5)  # pequeña pausa para estabilizar backend
                                print(f"NDC{monitor} -> Click en procesar del modal")

                                # 3. Esperar resultado final
                                # Esperar a que deje de decir "Validando"
                                WebDriverWait(driver, 25).until(
                                    lambda d: (
                                        d.find_element(By.ID, "alert").is_displayed() and
                                        "Validando" not in d.find_element(By.ID, "alert").text and
                                        len(d.find_element(By.ID, "alert").text.strip()) > 10
                                    )
                                )

                                alerta = driver.find_element(By.ID, "alert")
                                texto_alerta = alerta.text

                                print(f"NDC{monitor} -> Resultado FINAL: {texto_alerta}")

                                print(
                                    f"NDC{monitor} -> ALERT DETECTADO OK, validando resultado..."
                                )
                                

                                # 4. VALIDAR ÉXITO REAL
                                if "Declaración realizada exitosamente" in texto_alerta:

                                    print(f"NDC{monitor} -> Declaración EXITOSA")

                                    # 👉 EXTRAER DOCUMENTO

                                    doc = re.search(
                                        r"Nro\. Documento:\s*(\d+)", texto_alerta
                                    )
                                    doc_num = doc.group(1) if doc else None

                                    print(f"NDC{monitor} -> Documento: {doc_num}")

                                    # 5. Click aceptar
                                    boton_ok = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable(
                                            (
                                                By.XPATH,
                                                "//button[.//span[text()='Aceptar']]",
                                            )
                                        )
                                    )
                                    
                                    boton_ok.click()

                                    esperar_modal_cerrado(driver)

                                    # 6. AHORA SÍ actualizar BD
                                    actualizar_estado(ot, "Procesado", "PROCESADO")
                                    print(
                                        f"NDC{monitor} -> OT {ot} procesada correctamente en sistema y BD"
                                    )
                                    continue

                                else:
                                    print(
                                        f"NDC{monitor} -> ERROR: no se declaró correctamente"
                                    )
                                    actualizar_estado(ot, "Error", "FALLO_PROCESO")
                                    continue

                                # 👉 esperar botón aceptar

            #
            # ------- TG Bots Final del proceso ya sea por 100% de registros procesados o error sin alternativa de continuar ------
            #
            except Exception as e:
                print(f"NDC{monitor} -> ERROR CONTROLADO: {e}")
                return
                
        finally:

            print(f"NDC{monitor} -> Proceso finalizado correctamente")
            with open(archivolog, "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(
                    [
                        time.strftime("%Y-%m-%d %H:%M:%S"),
                        " Navegación ",
                        " TG Bots -> El archivo de Excel ha sido cerrado.",
                    ]
                )
                writer.writerow(
                    [
                        "-------------------",
                        "--------------",
                        "--------------------------------------------------------------------------------",
                    ]
                )

    def tearDown(self):
        print(f"NDC{monitor} -> Cierre")
        sys.stdout.flush()
        with open(archivolog, "a", newline="") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(
                [
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                    " Navegación ",
                    " TG Bots -> Cierre",
                ]
            )
            writer.writerow(
                [
                    "-------------------",
                    "--------------",
                    "--------------------------------------------------------------------------------",
                ]
            )
        time.sleep(2)
        self.driver.close()


if __name__ == "__main__":
    unittest.main()
