# Proyecto : RPA "TG Bots"
# Propietario : Technogroup SPA
# Creador: Jose Huaiquiman Arce
# modificacion: Adolfo Puentes
#Modificacion: 2024-06-17 - Conexion de base de datos e integracion al bot

from db_connection import obtener_ordenes, obtener_materiales_por_ot, actualizar_estado

from selenium import webdriver  
import time  #Ejercicio 1 , libreria nativa de python, para controlar el tiempo de espera del time.sleep()
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
import unittest 
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains #permite importar la clase ActionChains, esta clase permite automatizar interacciones de usuario más complejas en una página web
from selenium.webdriver.chrome.options import Options #permite personalizar el comportamiento del navegador Chrome antes de que se inicie

import os  # Permite ver que carpeta esta direccionada la busqueda de un archivo
import csv # para manejo de archivos csv
import re # ¡Añade esta línea!
import sys
import json
from pathlib import Path


#from selenium.common.exceptions import TimeoutException
#from playsound import playsound # primero se ejecuta (pip install pyglet) y luego (pip install playsound)
#print(os.getcwd()) # Muestra la ruta de archivos adicionales

monitor="1"
archivolog=f"./Digitacion/log_ejecucion_monitor{monitor}.csv"
statusorden=f"./Digitacion/status_orden_de_trabajo.csv"
os.makedirs(os.path.dirname(archivolog), exist_ok=True)

class usando_unittest(unittest.TestCase):
    
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.set_window_rect(x=450, y=0, width=800, height=800)
        self.driver.implicitly_wait(10) # Espera implícita para evitar errores de elementos no encontrados.

    def test_a_run_1_de_2(self):
        driver = self.driver
        driver.get("http://eps.vtr.cl/eps/declaracion_consumo.do#")
        driver.execute_script("document.body.style.zoom='65%'")
        with open( archivolog, 'a', newline='') as csvfile: # 'w' (write): Borra el contenido existente y escribe desde cero. 'a' (append): Agrega las nuevas líneas al final del archivo. 
            writer = csv.writer(csvfile)
            writer.writerow(['-------------------','--------------','DECLARACION X BOTS (1/2)'])
            writer.writerow(['Timestamp', 'Accion', 'Detalle'])
            print(f'NDC{monitor} -> Abriendo navegador en la pagina de inicio')
            sys.stdout.flush() #permite enviar el print al frame del home
            writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'Navegacion', 'TG Bots -> Pagina de inicio cargada'])

        try: # Este bloque permite al RPA esperar a que el objeto o componente termine de cargar.
            if getattr(sys, 'frozen', False):
                # Si es un ejecutable, la base es la carpeta donde está el archivo .exe
                base_dir = Path(sys.executable).resolve().parent
            else:
                # Si es un script .py, la base es la carpeta donde está el archivo .py
                base_dir = Path(__file__).resolve().parent
                   
            # --- PASO NUEVO: Cargar credenciales desde el archivo ---
            keys = base_dir / 'keys.txt'
            with open(keys, 'r') as f:
                credentials = json.load(f)
          
            user_val = credentials['user']
            pass_val = credentials['pass']
            # -------------------------------------------------------

            usuario_crecencial = driver.find_element(By.XPATH, f"//*[@id='username']")
            usuario_crecencial.send_keys(user_val)
            time.sleep(2)
            
            clave_crecencial = driver.find_element(By.XPATH, f"//*[@id='password']")
            clave_crecencial.send_keys(pass_val) 
            time.sleep(2)

            btn_crecencial = driver.find_element(By.XPATH, f"//*[@id='btn_valida_ingreso']")
            btn_crecencial.click()
            driver.execute_script("document.body.style.zoom='65%'")
            print(f"NDC{monitor} -> Credenciales y Acceso ok")
            sys.stdout.flush()
            with open( archivolog, 'a', newline='') as csvfile: 
                writer = csv.writer(csvfile)
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'Credenciales', 'NDC -> Credenciales y Acceso ok'])
                writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
            time.sleep(10) # Espera para que cargue bien la pagina
    #
    #------- Importación de datos desde excel DATA.xlsx (modificacion)
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
                    tecnico = orden.Tecnico

                    print(f"Procesando OT: {ot}")
            #
            # -------  OT Busqueda -------
            #
                    folio_ot = driver.find_element(By.XPATH, f"//*[@id='numero_ot']")
                    folio_ot.clear()  # limpia el contenido anterior
                    folio_ot.send_keys(ot)
                    folio_ot.click()
                    #print("NDC{monitor} -> registro de orden de trabajo ok")
                    time.sleep(1)

                    wait = WebDriverWait(driver, 8)  # Espera un máximo de 15 segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                    btn_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='boton-buscar']/span")))
                    btn_buscar.click()
                    time.sleep(5)

                    # VALIDACIÓN NUEVA (VA AQUÍ)
                    try:
                        WebDriverWait(driver, 3).until(
                            EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "La orden no existe o no está finalizada")
                        )

                        print(f"NDC{monitor} -> OT {ot} no existe o no finalizada")

                        with open(statusorden, 'a', newline='') as csvfile:
                            writer = csv.writer(csvfile)
                            writer.writerow([f'{ot}; no_existe_o_no_finalizada;', time.strftime('%Y-%m-%d %H:%M:%S')])
                        actualizar_estado(ot, "NO_EXISTE")
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
                            EC.text_to_be_present_in_element((By.XPATH, "//*[@id='alert']"), "Existe actividad en la OT no asociada a un trabajo")
                        )
                        # Si la alerta existe, busca y haz clic en el botón
                        btn_alerta = driver.find_element(By.XPATH, "/html/body/div[8]/div[3]/div/button/span")
                        btn_alerta.click()
                        print(f"NDC{monitor} -> Alerta 'Actividad en la OT no asociada' detectada y aceptada.")
                        alerta_Ot_no_asociada = False   
                    except:
                        alerta_Ot_no_asociada = True
                        # Si la alerta no aparece en 2 segundos, no hace nada y el código continúa
                        #print(f"NDC{monitor} -> Sin alerta, continua el proceso.")

                    if alerta_Ot_no_asociada:
            #
            # -------  OT Validación trabajo ya declarado -------
            #
                        try:
                            WebDriverWait(driver, 6).until(
                                EC.any_of(  
                                   EC.text_to_be_present_in_element((By.XPATH, "//*[@id='div_declaraciones_realizadas']/table/tbody/tr[3]/td[6]"), "Declarada"),
                                   EC.text_to_be_present_in_element((By.XPATH, "//*[@id='div_declaraciones_realizadas']/table/tbody/tr[3]/td[6]"), "En proceso")
                                )
                            )
                            print(f"NDC{monitor} -> Error - Orden fue declarada con anterioridad.")
                            sys.stdout.flush()
                            print("-" * 50) # Separador para mejor lectura
                            sys.stdout.flush()
                            with open( archivolog, 'a', newline='') as csvfile: 
                                writer = csv.writer(csvfile)
                                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> NO Declarada - Orden fue declarada con anterioridad'])
                                writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                            with open( statusorden, 'a', newline='') as csvfile: 
                                writer = csv.writer(csvfile)
                                writer.writerow([ f'{ot} ; Declarada o En Proceso ;', time.strftime('%Y-%m-%d %H:%M:%S')])    
                        except:
                    #
                    #--- Seleccion de la actividad realizada
                    #
                            try:
                                wait = WebDriverWait(driver, 6)
                                seleccion_trabajo = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='trabajo-material']")))
                                
                                dropdown = Select(seleccion_trabajo)  # Create a Select object for the dropdown
                                dropdown.select_by_index(1)
    
                                selected_option_text = dropdown.first_selected_option.text #Captura el dato visible del select
                                if selected_option_text == "Access Points Anexo WIFI (Extensor)" or selected_option_text == "Access Points Anexo":
                                #if selected_option_text == "Access Points Anexo WIFI (Extensor)" :    
                                    print(f"NDC{monitor} -> 'Access Point Anexo WIFI (Extensor)' se mueve a la siguente opcion.")
                                    dropdown.select_by_index(2) #Pasa a la siguiente opción
                                    time.sleep(2) 
                                else:
                                    #print(f"NDC{monitor} -> Trabajo seleccionado.")
                                    time.sleep(2) 
                                seguir_declaracion = True
                            except:
                                print(f"NDC{monitor} -> Error - Actividad no termino de cargar o existe un problema con la actividad")
                                sys.stdout.flush()
                                print("-" * 50) # Separador para mejor lectura
                                sys.stdout.flush()
                                with open( archivolog, 'a', newline='') as csvfile: 
                                    writer = csv.writer(csvfile)
                                    writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> NO Declarada - Actividad no termino de cargar o existe un problema con la actividad'])
                                    writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                                seguir_declaracion = False 
                                with open( statusorden, 'a', newline='') as csvfile: 
                                    writer = csv.writer(csvfile)
                                    writer.writerow([ f'{ot} ; existe un problema con la actividad ;', time.strftime('%Y-%m-%d %H:%M:%S')]) 

                    #
                    #--- Bloque de Carga de material
                            saltar_btn_procesar = False
                            salto_a_otra_orden = False
                            time.sleep(8) 
                            if seguir_declaracion:                             
                                    materiales = obtener_materiales_por_ot(ot)

                                    for material in materiales:
                                        producto = material.Codigo_a_rebajar
                                        cantidad = str(material.Declarado)

                                        ingreso_producto = driver.find_element(By.XPATH, "//*[@id='in_codigo']")
                                        valor_del_campo = ingreso_producto.get_attribute("value")
                                        time.sleep(2)

                                        
                                        ingreso_producto.clear()  # Limpia el campo antes de ingresar el nuevo valor
                                        ingreso_producto.send_keys(producto)
                                        

                                        wait = WebDriverWait(driver, 10)
                                        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='id_material']")))

                                        ingreso_cantidad = driver.find_element(By.XPATH, "//*[@id='cantidad']")
                                        ingreso_cantidad.send_keys(cantidad)
                                        ingreso_cantidad.click()

                                        btn_agregar_material = driver.find_element(By.XPATH, "//*[@id='ingresar_nuevo']/span")
                                        btn_agregar_material.click()
                                        saltar_btn_procesar = True
                                        time.sleep(1)
                        #
                        #------- Carga de materiales -> Alerta por material repetido ---      
                        #                           

                                        try:
                                            wait = WebDriverWait(driver, 1)  # Espera un máximo de 1x segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                                            btn_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div[3]/div/button/span")))
                                            btn_buscar.click()
                                            wait = WebDriverWait(driver, 1)  # Espera un máximo de x segundos la visibillidad del boton, pero si lo encuentra pasa inmediatamente
                                            btn_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='in_codigo']")))
                                            btn_buscar.clear()                                           
                                            print(f'NDC{monitor}-> Material Repetido')
                                            
                                        except:
                                            
                                            time.sleep(1)
                    
                            else:
                                salto_a_otra_orden = False
                    #
                    #--- Procesar -> PreCierr de declaración
                    #   
                            if saltar_btn_procesar: # si es True ejecuta el bloque de procesar declaración
                                btn_procesar_material = driver.find_element(By.XPATH, f"//*[@id='btn_procesar_declaracion']/span")
                                btn_procesar_material.click()
                                #print("NDC{monitor} -> btn procesar material ok")
                                time.sleep(8)
    
                                try:  # bloque que detecta error en materiales no encontrado o fuera del estandar de trabajo
                                    WebDriverWait(driver, 2).until(
                                        EC.text_to_be_present_in_element((By.TAG_NAME, 'body'), "Debe declarar material")
                                    )
                                    boton_material_notificacion = driver.find_element(By.XPATH, "/html/body/div[8]/div[3]/div/button/span") # Usar XPATH Procesar
                                    boton_material_notificacion.click()
                                    print(f"NDC{monitor} -> Error - no detecta producto erroneo.")
                                    sys.stdout.flush()
                                    print("-" * 50) # Separador para mejor lectura
                                    sys.stdout.flush()
                                    with open( archivolog, 'a', newline='') as csvfile: 
                                        writer = csv.writer(csvfile)
                                        writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> NO Declarada - producto erroneo'])
                                        writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                                    with open( statusorden, 'a', newline='') as csvfile: 
                                        writer = csv.writer(csvfile)
                                        writer.writerow([ f'{ot} ; producto erroneo ;', time.strftime('%Y-%m-%d %H:%M:%S')])     
                                    salto_a_otra_orden = False
                                except:
                                    salto_a_otra_orden = True
                    #
                    #--- Procesar -> Validación de stock final
                    #
                            if salto_a_otra_orden: # si es True ejecuta el bloque de validacion antes de procesar ------
                                time.sleep(3)
                                try:
                                    WebDriverWait(driver, 6).until(
                                        #EC.text_to_be_present_in_element((By.TAG_NAME, 'body'), "Error Stock") # no tocar
                                        EC.text_to_be_present_in_element((By.TAG_NAME, 'body'), "sin stock suficiente")
                                    )
                                    print(f"NDC{monitor} -> No Declarada - Stock InSuficiente.")
                                    sys.stdout.flush()
                                    print("-" * 50) # Separador para mejor lectura
                                    sys.stdout.flush()
                                    with open( archivolog, 'a', newline='') as csvfile: 
                                        writer = csv.writer(csvfile)
                                        writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> NO Declarada - Stock InSuficiente'])
                                        writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                                    with open( statusorden, 'a', newline='') as csvfile: 
                                                writer = csv.writer(csvfile)
                                                writer.writerow([ f'{ot}; falta-stock ;', time.strftime('%Y-%m-%d %H:%M:%S')]) 
                                    time.sleep(3)
                                    try:
                                       boton_aceptar = WebDriverWait(driver, 6).until(
                                           EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Aceptar')]"))
                                       )
                                       
                                       # 2. Una vez encontrado, hacer el clic
                                       boton_aceptar.click()
                                       print("Se encontro el boton y se hizo clic en Aceptar.")                                      
                                    except:
                                       ActionChains(driver).send_keys(Keys.TAB).perform()
                                       ActionChains(driver).send_keys(Keys.ENTER).perform()
                                       print("Click en 'Tab-action' realizado con éxito.")
                                       sys.stdout.flush()
                                    
                                    time.sleep(3)
                                    boton_cerrar = driver.find_element(By.XPATH, "/html/body/div[6]/div[3]/div/button[2]") 
                                    boton_cerrar.click()
                                except:
                                        print(f"NDC{monitor} -> Stock correcto se procede a procesar.")
                                        try:
                                            boton_procesar2 = WebDriverWait(driver, 8).until(
                                                EC.element_to_be_clickable((By.XPATH, "/html/body/div[7]/div[3]/div/button[1]"))
                                            )

                                            # 2. Ahora que lo tenemos guardado en 'boton_procesar', hacemos el click
                                            boton_procesar2.click()
                                            print(f"NDC{monitor} -> apunto de procesar - boton procesar clickable.")
                                        except:
                                            WebDriverWait(driver, 8).until(
                                            EC.text_to_be_present_in_element((By.TAG_NAME, 'body'), "Procesar")
                                            )
                                            print(f"NDC{monitor} -> apunto de procesar Declarada - boton procesar in_element.")
                                            time.sleep(10)
                                        with open( archivolog, 'a', newline='') as csvfile: 
                                            writer = csv.writer(csvfile)
                                            writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> Declarada sin refuerzo'])
                                            print(f"NDC{monitor} -> apunto de procesar Declarada - boton xpath")
                                        try:
                                            
                                            time.sleep(10)
                                            btn_procesar_notificacion = driver.find_element(By.XPATH, f"/html/body/div[9]/div[3]/div/button/span")
                                            btn_procesar_notificacion.click()
                                            time.sleep(10)
                                            print(f"NDC{monitor} -> Material declarado correctamente.")
                                            sys.stdout.flush()
                                            print("-" * 50) # Separador para mejor lectura
                                            actualizar_estado(ot, "OK")
                                            print(f"OT {ot} actualizada a estado 'Procesado' en la base de datos.")
                                            sys.stdout.flush()
                                            with open( archivolog, 'a', newline='') as csvfile: 
                                                writer = csv.writer(csvfile)
                                                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> SI Declarada >>>>>'])
                                                writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                                            with open( statusorden, 'a', newline='') as csvfile: 
                                                writer = csv.writer(csvfile)
                                                writer.writerow([ f'{ot}; Declarada o En Proceso ;', time.strftime('%Y-%m-%d %H:%M:%S')]) 
                                        except: # Se mueve a otra pestaña si el procesamiento final se demora mas de lo esperado.
                                            print(f"NDC{monitor} -> Orden no declarada, fallo la ultima notificación.")
                                            sys.stdout.flush()
                                            print("-" * 50) # Separador para mejor lectura
                                            sys.stdout.flush()
                                            with open( archivolog, 'a', newline='') as csvfile: 
                                                writer = csv.writer(csvfile)
                                                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> Orden no declarada, fallo la ultima notificación.'])
                                                writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                                             # Obtener la ventana actual (la pestaña principal)
                                            original_window = driver.current_window_handle
                                            # Abrir una nueva pestaña usando JavaScript
                                            # Esto creará una nueva pestaña, pero el driver seguirá enfocado en la original
                                            driver.execute_script("window.open('');")
                                            # Esperar a que la nueva ventana se abra (es una buena práctica)
                                            time.sleep(4)
                                            # Obtener todas las ventanas/pestañas abiertas
                                            #all_windows = driver.window_handles
                                            # Buscar la nueva ventana y cambiar el enfoque a ella

                                            # 4. Cambiar el enfoque a la ÚLTIMA ventana abierta
                                            # [-1] selecciona el último elemento de la lista de handles
                                            driver.switch_to.window(driver.window_handles[-1])
                                           # for window_handle in all_windows:
                                           #     if window_handle != original_window:
                                           #         driver.switch_to.window(window_handle)
                                           #         break
                                            driver.get("http://eps.vtr.cl/eps/declaracion_consumo.do#")
                                            # Si quieres volver a la pestaña original más tarde, simplemente haz esto:
                                            # driver.switch_to.window(original_window)

                    else:
                        print(f"NDC{monitor} -> Error - OT no asociada a un trabajo.")
                        sys.stdout.flush()
                        print("-" * 50) # Separador para mejor lectura
                        sys.stdout.flush()
                        with open( archivolog, 'a', newline='') as csvfile: 
                            writer = csv.writer(csvfile)
                            writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' NDC -> NO Declarada - Existe actividad en la OT no asociada a un trabajo'])
                            writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
            #
            #--- Eliminación -> Campos correctamente procesados, serán eliminados
            #               
                
        #
        #--- Errores -> de extraccion de datos
        #
           
            
            except Exception as e:
                with open( archivolog, 'a', newline='') as csvfile: 
                                           writer = csv.writer(csvfile)
                                           writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' TG Bots -> Error: Ocurrió un error inesperado.'])
                                           writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
                print(f"Error: Ocurrió un error inesperado: {e}")
    #        
    #------- TG Bots Final del proceso ya sea por 100% de registros procesados o error sin alternativa de continuar ------
    #
        except Exception as e:
            time.sleep(5)
            with open( archivolog, 'a', newline='') as csvfile: 
                                           writer = csv.writer(csvfile)
                                           writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' CierreOrden ', ' TG Bots -> No se pudo encontrar o interactuar con el elemento.'])
                                           writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
            self.fail(f"No se pudo encontrar o interactuar con el elemento: {e}")
        finally:
          
                print(f"NDC{monitor} -> Proceso finalizado correctamente")
                with open( archivolog, 'a', newline='') as csvfile: 
                                           writer = csv.writer(csvfile)
                                           writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' Navegación ', ' TG Bots -> El archivo de Excel ha sido cerrado.'])
                                           writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
       

    
    def tearDown(self):
        print(f"NDC{monitor} -> Cierre")
        sys.stdout.flush()
        with open( archivolog, 'a', newline='') as csvfile: 
                                           writer = csv.writer(csvfile)
                                           writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), ' Navegación ', ' TG Bots -> Cierre'])
                                           writer.writerow(['-------------------','--------------','--------------------------------------------------------------------------------'])
        time.sleep(2)
        self.driver.close()

if __name__ == '__main__':
    unittest.main()  


