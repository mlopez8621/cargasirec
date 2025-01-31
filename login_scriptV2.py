import sys
import os
import pandas as pd
import time
import ctypes
import logging
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from logging.handlers import RotatingFileHandler
from selenium.common.exceptions import TimeoutException

def refrescar_detalle_programacion():
    try:
        # Espera a que el botón de "Cancelar" esté presente y sea clickeable
        cancelar_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarDetProgramacionPopUp"))
        )
        time.sleep(1)
        cancelar_button.click()
        time.sleep(1)
        logging.info("Botón 'Cancelar' presionado.")
    except Exception as e:
        logging.error("Error al hacer clic en el botón 'Cancelar':", e)

    try:
        # Espera adicional para asegurar que el botón esté cargado
        time.sleep(2)
        
        # Desplazar a la vista y hacer clic en el botón "Crear detalle" por su ID
        crear_detalle_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_cell0_7_btnNuevoDetProgramacion_0"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", crear_detalle_button)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_cell0_7_btnNuevoDetProgramacion_0"))).click()
        logging.info("Botón 'Crear detalle' presionado exitosamente.")
    except Exception as e:
        logging.error("Error al hacer clic en el botón 'Crear detalle':", e)

# Mostrar mensaje de finalización
def mostrar_mensaje_exito(mensaje, titulo, archivo_ruta):
    try:
        import webbrowser
        # Intentar abrir el archivo directamente
        if ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, 0x40 | 0x0) == 1:  # El botón OK retorna 1
            webbrowser.open(f'file:///{archivo_ruta}')  # Abrir el archivo en el sistema
    except Exception as e:
        logging.error(f"Error al intentar abrir el archivo: {e}")

# Definir el método para volver a la página de "Programación de Cortometrajes"
def volver_a_programacion_cortometrajes():
    # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
    ##time.sleep(1)
    try:
        programacion_cortometrajes_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='/frm/peliculas/frmLstPEL_ENC_PROGRAMACION_CORTOS.aspx']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", programacion_cortometrajes_link)
        programacion_cortometrajes_link.click()
        logging.info("Volviendo a Programación de Cortometrajes.")
    except Exception as e:
        logging.error(f"Error al hacer clic en Programación de Cortometrajes: {e}")
        logging.exception("Error al hacer clic en Programación de Cortometrajes:")

def validar_asignacion_fechas():
    try:
        error_message_element = WebDriverWait(driver, 1).until(
        EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_lblErrorTab_3"))
        )
        error_message = error_message_element.text
        if "No puede cargar fecha de inicio y fin por fuera del rango del mes y año de beneficio" in error_message:
            return False, error_message
        elif "La fecha final no puede ser menor de 8 dias continuos"  in error_message:   
            return False, error_message
        elif "Sala" in error_message and "El número de días consecutivos" in error_message:
            return False, error_message  # Captura el error con valores dinámicos
        else:
            return True    
    except Exception as e:
        logging.info("Botón Asignar Fechas presionado.")
        return True, "Botón Asignar Fechas presionado."

def limpiar():
    try:
        xpath_checkbox_select_all = "//input[@id='ContentPlaceHolder1_chkSeleccionarTodo' and @type='checkbox']"
        checkbox_select_all = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, xpath_checkbox_select_all))
        )
        logging.info("'Seleccionar todos' está presente.")

        # Verificar si el checkbox está marcado
        is_checked = driver.execute_script("return arguments[0].checked;", checkbox_select_all)
        time.sleep(1)
        if is_checked:
            logging.info("El checkbox 'Seleccionar todos' está marcado. Procediendo a desmarcarlo.")
            checkbox_select_all.click()
            time.sleep(1)
            logging.info("Checkbox 'Seleccionar todos' ahora está desmarcado.")
        else:
            logging.info("El checkbox 'Seleccionar todos' ya estaba desmarcado.")
    except Exception as e:
        logging.error(f"Error al verificar el estado de 'Seleccionar todos': {e}")
        logging.error(f"Detalles del error: {traceback.format_exc()}")
        return False

    try:
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cmbComplejo"))
        )
        select = Select(select_element)
        try:
            select.select_by_value("-1")  # Cambia id_complejo por el valor que corresponda si es necesario
            time.sleep(1)
        except:
            select.select_by_visible_text("Seleccione")
            time.sleep(1)
    except Exception as e:
        logging.error(f"Error al seleccionar la opción SELECCIONE en el menú desplegable: {e}")
        return False
    
    try:
        fecha_inicio_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaInicio"))
        )
        fecha_inicio_input.clear()  # Limpia el campo antes de ingresar la fecha
    except Exception as e:
        logging.error(f"Error al limpiar la fecha de inicio: {str(e)}")
        return False
   
    try:
        fecha_fin_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaFin"))
        )
        fecha_fin_input.clear()  # Limpia el campo antes de ingresar la fecha
        return True
    except Exception as e:
        logging.error(f"Error al ingresar la fecha final: {str(e)}")
        return False        

def asignar_fechas():
    # Hacer clic en "Asignar Fechas"
    try:
        asignar_fechas_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAgregarFechas"))
        )
        asignar_fechas_button.click()
        time.sleep(0.5)
        logging.info("Botón Asignar Fechas presionado.")
        return True
    except Exception as e:
        logging.info("Asignación de fechas:", e)
        return False        

def guardar_continuar():    
    # Hacer clic en "Guardar y continuar"
    try:
        guardar_continuar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnGuardarDetProgramacion"))
        )
        guardar_continuar_button.click()
        #time.sleep(2)
        logging.info("Botón Guardar y continuar presionado.")
        return True
    except Exception as e:
        logging.info("Guardar programacion:", e)
        return False 

def confimar():
    try:
        WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".confirm"))
            ).click()
        #time.sleep(2)
        logging.info("Programacion guardada")
        return True        
    except Exception as e:
        logging.info("Guardar programacion:", e)
        return False        
        
    

    
if __name__ == "__main__":
    # Obtener la ruta del archivo actual
    current_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(current_dir, "script.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=3),  # 5 MB por archivo, 3 backups
            logging.StreamHandler()
        ]
    )

    # Validar argumentos
    if len(sys.argv) != 3:
        logging.error("Uso: python login_script.py <ruta_excel> <nombre_hoja>")
        sys.exit(1)

    # Recibir parámetros desde la macro de Excel
    file_path = sys.argv[1]  # Ruta del archivo Excel
    sheet_name = sys.argv[2]  # Nombre de la hoja

    # Verificar si el archivo existe
    if not os.path.exists(file_path):
        logging.error(f"Error: El archivo '{file_path}' no existe.")
        sys.exit(1)

    # Configuración del servicio de ChromeDriver
    DRIVER_PATH = 'C:\\chromedriver\\chromedriver.exe'
    TIMEOUT = 20  # Tiempo máximo de espera en segundos
    service = Service(DRIVER_PATH)

    # Inicializa el controlador con el servicio configurado
    try:
        driver = webdriver.Chrome(service=service)
    except Exception as e:
        logging.error(f"Error al iniciar el navegador: {e}")
        sys.exit(1)

    # URL de la página de inicio de sesión
    url = "https://sirec.gov.co/frm/seguridad/frmLogin.aspx"
    driver.get(url)

    # Tiempo de espera para que cargue la página
    #time.sleep(1)

   # Configurar WebDriverWait
    wait = WebDriverWait(driver, 20)  # Máximo 20 segundos de espera

    # Esperar hasta que los campos estén presentes en el DOM
    username_field = wait.until(EC.presence_of_element_located((By.ID, "txtUsuario")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))

    # Ingresar las credenciales
    username = "yrojas"  # Nombre de usuario
    password = "Subrecaudo2024*"  # Contraseña

    username_field.send_keys(username)
    password_field.send_keys(password)

    # Esperar hasta que el botón de ingreso sea clickeable
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "btnIngresar")))

    # Hacer clic en el botón
    login_button.click()

    # Esperar a que el menú de "SIREC" esté presente y hacer clic para expandirlo
    try:
        sirec_menu = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'SIREC')]"))
        )
        sirec_menu.click()
        ##time.sleep(1)
    except Exception as e:
        logging.error(f"Error al localizar el menú de SIREC: {e}")

    # Hacer clic en la sección "Programación de Cortometrajes" para expandir su menú
    try:
        programacion_menu = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Programación de Cortometrajes')]"))
        )
        programacion_menu.click()
        ##time.sleep(1)
    except Exception as e:
        logging.error(f"Error al localizar el menú de Programación de Cortometrajes: {e}")

    # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
    try:
        programacion_cortometrajes_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='/frm/peliculas/frmLstPEL_ENC_PROGRAMACION_CORTOS.aspx']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", programacion_cortometrajes_link)
        programacion_cortometrajes_link.click()
    except Exception as e:
        logging.error(f"Error al hacer clic en Programación de Cortometrajes: {e}")

    # Intentar cargar el archivo
    try:
        logging.info(f"Cargando archivo: {file_path}, hoja: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'ID_COMPLEJO': str})
        logging.info("Archivo y hoja cargados exitosamente.")
    except FileNotFoundError:
        logging.error(f"Error: No se encontró el archivo '{file_path}'.")
        sys.exit(1)
    except ValueError as e:
        logging.error(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error inesperado: {e}")
        logging.exception(f"Error inesperado: {e}")
        sys.exit(1)

    # Verificar si el DataFrame está vacío
    if df.empty:
        logging.error("Error: El archivo Excel está vacío o la hoja no contiene datos.")
        sys.exit(1)

    # Verificar columnas esperadas
    columnas_esperadas = ['NOMBRE_EXHIBIDOR', 'ID_EXHIBIDOR', 'ANIO_BENEFICIO', 'MES_BENEFICIO']
    for columna in columnas_esperadas:
        if columna not in df.columns:
            logging.error(f"Error: Falta la columna esperada '{columna}' en la hoja '{sheet_name}'.")
            sys.exit(1)

    logging.info("El archivo y la hoja contienen los datos necesarios.")

    # Agregar la columna Estado si no existe
    if 'Estado' not in df.columns:
        df['Estado'] = ""

    # Conjunto de índices procesados
    indices_procesados = set()

    # Recorrer cada fila del archivo Excel y llenar el formulario
    for index, row in df.iterrows():
        start_time = time.time()
        if index in indices_procesados:
            # Si el índice ya fue procesado, continuar con el siguiente
            logging.info(f"Índice {index} ya procesado. Saltando...")
            end_time = time.time()
            duration = end_time - start_time
            df.at[index, 'Tiempo'] = duration
            logging.info(f"Tiempo tomado para la fila {index}: {duration:.2f} segundos")
            df.at[index, 'Estado'] = "Insertado"
            continue  # Omitir si ya fue procesado

        volver_a_programacion_cortometrajes()
        

        exhibidor = row['NOMBRE_EXHIBIDOR']
        id_exhibidor = row['ID_EXHIBIDOR']
        anio_beneficio = row['ANIO_BENEFICIO']
        mes_beneficio = row['MES_BENEFICIO']
        nombre_corto = row['NOMBRE_CORTO']
        certificado = row['CERTIFICADO']
        fecha_certificado = row['FECHA_CERTIFICADO']
        numero_acta = row['ACTA']
        clasificacion = row['CLASIFICACION']
        item = row['ITEM']
        ciudad = row['CIUDAD']
        id_complejo = row['ID_COMPLEJO']
        nombre_complejo = row['NOMBRE_COMPLEJO']
        id_sala = row['ID_SALA']
        nombre_sala = row['NOMBRE_SALA']
        fecha_exhi_inicial = row['FECHA_EXHI_INICIAL']  # Asegúrate de que esté en formato 'dd/mm/yyyy'
        fecha_exhi_final = row['FECHA_EXHI_FINAL']  # Asegúrate de que esté en formato 'dd/mm/yyyy'
        observaciones = row['OBSERVACIONES']

        # Imprimir las variables para verificar el formato
        logging.info(f"Exhibidor: {exhibidor}")
        logging.info(f"ID Exhibidor: {id_exhibidor}")
        logging.info(f"Año de Beneficio: {anio_beneficio}")
        logging.info(f"Mes de Beneficio: {mes_beneficio}")
        logging.info(f"Nombre Corto: {nombre_corto}")
        logging.info(f"Certificado: {certificado}")
        logging.info(f"Fecha Certificado: {fecha_certificado}")
        logging.info(f"Número de Acta: {numero_acta}")
        logging.info(f"Clasificación: {clasificacion}")
        logging.info(f"Ítem: {item}")
        logging.info(f"Ciudad: {ciudad}")
        logging.info(f"ID Complejo: {id_complejo}")
        logging.info(f"Nombre Complejo: {nombre_complejo}")
        logging.info(f"ID Sala: {id_sala}")
        logging.info(f"Nombre Sala: {nombre_sala}")
        logging.info(f"Fecha de Exhibición Inicial: {fecha_exhi_inicial}")
        logging.info(f"Fecha de Exhibición Final: {fecha_exhi_final}")
        logging.info(f"Observaciones: {observaciones}")
        logging.info("-" * 40)


        try:
            # Buscar el campo de búsqueda para ingresar el nombre del exhibidor
            #time.sleep(1)
            volver_a_programacion_cortometrajes()
            search_field = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Buscar']"))
            )
            #search_field.send_keys(exhibidor)
            # búsca por el ID de la programacion asignada al exhibidor
            search_field.send_keys(id_exhibidor)

            ##time.sleep(1)
            tabla = wait.until(
                EC.presence_of_element_located((By.ID, "grdDatos"))
            )

            # Buscar si existe un <td> con el texto 'Sin registros'
            sin_registros_element = tabla.find_elements(By.XPATH, ".//td[contains(text(), 'Sin registros')]")

            # Verificar si el elemento está presente y es visible
            if sin_registros_element and sin_registros_element[0].is_displayed():
                error_message = "Sin registros"
                logging.error(f"Error detectado: {error_message}")
                df.at[index, 'Estado'] = f"Error: {error_message}"  # Guardar el error en el DataFrame
                continue  # Pasar al siguiente registro

            # Esperar y hacer clic en el botón "Editar"
            ##time.sleep(1)
            # Elimina los espacios en blanco al inicio y al final del nombre del exhibidor
            try:
                exhibidor = exhibidor.strip()

                # Construir el XPath dinámico con el valor de 'exhibidor'
                # xpath_editar_button = f"//tr[td[text()='{id_complejo}'] or td[contains(text(), '{exhibidor}')]]//input[@type='submit' and @value='Editar']"
                xpath_editar_button = f"//tr[td[text()='{id_exhibidor}'] or td[normalize-space(.)='{exhibidor}']]//input[@type='submit' and @value='Editar']"


                # Imprimir el XPath construido para verificar
                logging.info(f"XPath construido: {xpath_editar_button}")

                editar_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath_editar_button))
                )
            except TimeoutException:
                # Si el botón no se encuentra, intenta con un espacio adicional al final
                logging.warning("No se encontró el elemento con XPath original. Intentando con un espacio adicional...")

                # Agregar un espacio al final de la variable exhibidor
                exhibidor_con_espacio = exhibidor + " "
                # xpath_editar_button = f"//tr[td[text()='{id_complejo}'] or td[contains(text(), '{exhibidor_con_espacio}')]]//input[@type='submit' and @value='Editar']"
                xpath_editar_button = f"//tr[td[text()='{id_exhibidor}'] or td[normalize-space(.)='{exhibidor}']]//input[@type='submit' and @value='Editar']"

                # Imprimir el XPath modificado para verificar
                logging.info(f"Intentando con XPath modificado: {xpath_editar_button}")

                # Volver a intentar encontrar y hacer clic en el botón "Editar"
                editar_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_editar_button))
                )

            # Si se encuentra el elemento, hacer scroll y clic
            driver.execute_script("arguments[0].scrollIntoView(true);", editar_button)
            editar_button.click()
            logging.info("Botón 'Editar' clickeado exitosamente.")

            # Navegar a Datos Programación
            #time.sleep(1)
            wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Datos Programación"))
            ).click()
            logging.info("Navegando a Datos Programación")

            #time.sleep(1)
            crear_programacion_button = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCrearProgramacion")))
            crear_programacion_button.click()
            logging.info("Botón Crear Programación presionado")

            # Ingresar el número de acta
            time.sleep(1)
            numero_acta = str(numero_acta)
            acta_field = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtNroActa")))
            acta_field.send_keys("00"+numero_acta)
            # Simular Enter para mover el foco
            acta_field.send_keys(Keys.ENTER)
            logging.info(f"Número de acta ingresado: 00{numero_acta}")
            time.sleep(1)
            # Esperar si el elemento de error aparece

            elementos = driver.find_elements(By.ID, "ContentPlaceHolder1_lblErrorTab_2")
            if elementos and elementos[0].is_displayed():
                error_message_element = wait.until(
                    EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_lblErrorTab_2"))
                )
                error_message = error_message_element.text
                if "No existe pelicula para esa acta" in error_message:
                    cancelar_button = wait.until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                    )
                    cancelar_button.click()
                    logging.error(f"Error detectado tras ingresar el número de acta: {error_message}")
                    df.at[index, 'Estado'] = f"Error: {error_message}"  # Guardar el error en el DataFrame
                    continue  # Pasar al siguiente registro

            # Seleccionar año y mes
            #time.sleep(1)
            # Asegurarse de que anio_beneficio sea una cadena
            anio_beneficio = str(anio_beneficio)
            dropdown_anio = driver.find_element(By.ID, "ddlAnio")
            dropdown_anio.find_element(By.XPATH, "//option[. = '" + anio_beneficio + "']").click()
            logging.info("Año de beneficio seleccionado")

            #time.sleep(1)
            dropdown_mes = driver.find_element(By.ID, "drpMesBeneficio")
            dropdown_mes.find_element(By.XPATH, "//option[. = '" + mes_beneficio + "']").click()
            logging.info("Mes de beneficio seleccionado")
            
            time.sleep(1)
            driver.find_element(By.ID, "ContentPlaceHolder1_btnGuardarYsalir").click()
            logging.info("Botón Guardar y salir presionado")

            # Esperar a que aparezca el mensaje de error o el mensaje de éxito
            time.sleep(2)
            try:
                # Verificar si el elemento está presente
                elementos = driver.find_elements(By.CSS_SELECTOR, ".confirm")
                if elementos:
                    # Si la lista no está vacía, significa que el elemento está presente
                    logging.info("Cargue de acta exitoso")
                    wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".confirm"))
                    ).click()
                else:
                    error_message_element = wait.until(
                    EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_lblErrorProgramacion"))
                    )
                    error_message = error_message_element.text
                    if "Nº Acta es obligatorio" in error_message:
                        cancelar_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                        )
                        cancelar_button.click()
                        logging.error(error_message)
                        df.at[index, 'Estado'] = f"Error: {error_message}"
                        continue  # Pasa al siguiente registro en el bucle principal
                    if "Debe seleccionar un mes de beneficio" in error_message:
                        cancelar_button = wait.until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                        )
                        cancelar_button.click()
                        logging.error(error_message)
                        df.at[index, 'Estado'] = f"Error: {error_message}"
                        continue  # Pasa al siguiente registro en el bucle principal
                    if "Solo puede crear una programación por periodo" in error_message:
                        cancelar_button = wait.until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                        )
                        cancelar_button.click()
                        logging.warning(error_message)
                    if "Ya exhibio esta acta, solo puede exhibir esta acta maximo 1 vez" in error_message:
                        cancelar_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                        )
                        cancelar_button.click()
                        logging.warning(error_message)
            except Exception as e:
                logging.error(f"Error al guardar el acta: {e}")
                df.at[index, 'Estado'] = f"Error: {str(e)}"
                volver_a_programacion_cortometrajes()
                continue  # Pasa al siguiente registro en el bucle principal

            # try:
            #     # Localizar el campo de entrada por su ID y enviar el valor "00"+numero_acta
            #     nro_acta_field = WebDriverWait(driver, 10).until(
            #         EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_DXFREditorcol0_I"))
            #     )
            #     nro_acta_field.clear()  # Opcional: Limpia el campo antes de escribir
            #     nro_acta_field.send_keys("00"+numero_acta)
            #     logging.info("Número de acta ingresado correctamente en el campo de programación.")
            # except Exception as e:
            #     logging.error("Error al ingresar el número de acta:", e)

            try:
                # Espera adicional para asegurar que el botón esté cargado
                #time.sleep(2)
                numero_acta_f="00"+numero_acta

                # XPath para ubicar el botón "Crear detalle" dentro de la fila que contiene "00"+numero_acta
                xpath_boton = f"//tr[td[text()='{numero_acta_f}']]//input[@type='submit' and @value='Crear detalle']"
                
               # Esperar a que el botón esté presente en la página
                boton_crear_detalle =  wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath_boton)))

                # Hacer clic en el botón
                boton_crear_detalle.click()
                logging.info("Botón 'Crear detalle' presionado exitosamente.")

            except Exception as e:
                logging.error("Error al hacer clic en el botón 'Crear detalle':", e)        

            bandera = 1
            # Agrupar por ID_EXHIBIDOR y luego iterar por cada ID_COMPLEJO asociado
            grupo_exhibidor = df[df['ID_EXHIBIDOR'] == id_exhibidor]  # Filtrar el grupo del exhibidor actual
            logging.info(f"Contenido de grupo_exhibidor ordenado por 'ITEM':\n{grupo_exhibidor[['NOMBRE_COMPLEJO']]}")
            for id_complejo_actual, grupo_complejo in grupo_exhibidor.groupby('ID_COMPLEJO'):
                logging.info(f"Procesando complejo {id_complejo_actual} del exhibidor {id_exhibidor}...")
                logging.info(f"Contenido de grupo_complejo:\n{grupo_complejo[['NOMBRE_COMPLEJO']]}")
                
                # Filtrar el grupo de salas por el mismo ID_COMPLEJO
                grupo_complejo_actual = df[df['ID_COMPLEJO'] == id_complejo_actual]
                fechas_iniciales_iguales = grupo_complejo_actual['FECHA_EXHI_INICIAL'].nunique() == 1
                fechas_finales_iguales = grupo_complejo_actual['FECHA_EXHI_FINAL'].nunique() == 1
                mas_de_una_fila = len(grupo_complejo_actual) > 1

                if fechas_iniciales_iguales and fechas_finales_iguales and mas_de_una_fila:
                    # Seleccionar todos si las fechas son iguales para el mismo complejo
                    nombre_complejo_actual = grupo_complejo_actual['NOMBRE_COMPLEJO'].iloc[0]  # O usar .unique()[0
                    logging.info(f"Las fechas son iguales (ID_COMPLEJO = {id_complejo_actual}): {nombre_complejo_actual}")
                    try:
                        # Espera a que el elemento <select> esté presente
                        select_element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cmbComplejo"))
                        )

                        # Inicializa el objeto Select
                        select = Select(select_element)

                        # Intentar seleccionar por valor
                        try:
                            select.select_by_value(str(id_complejo_actual))  # Cambia id_complejo por el valor que corresponda si es necesario
                            time.sleep(2)
                            logging.info(f"Opción seleccionada por valor correctamente {id_complejo_actual} del exhibidor {id_exhibidor}...")
                        except:
                            # Si falla, intenta seleccionar por texto visible
                            logging.info(f"Seleccionando por texto visible: {nombre_complejo_actual}")
                            select.select_by_visible_text(nombre_complejo_actual)
                            time.sleep(2)
                            logging.info("Opción seleccionada por texto visible correctamente.")

                    except Exception as e:
                        logging.error(f"Error al seleccionar la opción en el menú desplegable: {e}")
                        logging.error(f"Error index {index} NO procesado - No Selecciono complejo .")
                        df.at[index, 'Estado'] = f"Error index {index} NO procesado - No Selecciono complejo ."
                        bandera = 0
                        continue  # Pasa al siguiente registro en el bucle principal

                    # Ingresar la fecha en el campo de Fecha de Inicio
                    time.sleep(1)
                    try:
                        fecha_inicio_input = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaInicio"))
                        )
                        fecha_inicio_input.clear()  # Limpia el campo antes de ingresar la fecha
                        #time.sleep(1)
                        fecha_inicio_input.send_keys(fecha_exhi_inicial.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                        fecha_inicio_input.send_keys(Keys.ENTER)
                        logging.info("Fecha de inicio ingresada correctamente.")
                    except Exception as e:
                        logging.error(f"Error al ingresar la fecha de inicio: {str(e)}")
                        df.at[index, 'Estado'] = f"Error al ingresar la fecha de inicio: {str(e)}"  # Registrar el error en el DataFrame
                        logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                        bandera=0
                        continue  # Pasar al siguiente elemento del bucle


                    # Ingresar la fecha en el campo de Fecha de Fin
                    time.sleep(1)
                    try:
                        fecha_fin_input = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaFin"))
                        )
                        fecha_fin_input.clear()  # Limpia el campo antes de ingresar la fecha
                        #time.sleep(1)
                        fecha_fin_input.send_keys(fecha_exhi_final.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                        fecha_fin_input.send_keys(Keys.ENTER)
                        logging.info("Fecha de fin ingresada correctamente.")
                    except Exception as e:
                        logging.error(f"Error al ingresar la fecha final: {str(e)}")
                        df.at[index, 'Estado'] = f"Error al ingresar la fecha final: {str(e)}"  # Registrar el error en el DataFrame
                        logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                        bandera=0
                        continue  # Pasar al siguiente elemento del bucle
                    try:
                        # Construir el XPath como una cadena
                        xpath_checkbox = "//input[@id='ContentPlaceHolder1_chkSeleccionarTodo' and @type='checkbox']"

                        # Imprimir el XPath para verificar cómo se ve
                        logging.info(f"XPath construido: {xpath_checkbox}")
                        checkbox = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_checkbox))
                        )
                        checkbox.click()
                        time.sleep(1)
                        logging.info(f"'Seleccionar todos' aplicado correctamente para el complejo {nombre_complejo}.")
                    except Exception as e:
                           logging.error(f"Error al seleccionar el Checkbox: {str(e)}")
                           df.at[index, 'Estado'] = f"Error al seleccionar el Checkbox: {str(e)}"  # Registrar el error en el DataFrame
                           logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                           bandera=0
                           continue  # Pasar al siguiente elemento del bucle    

                    try:
                        exito = asignar_fechas()
                        if exito:
                            exito, mensaje = validar_asignacion_fechas()
                            if not exito:
                                refrescar_detalle_programacion()  
                        if exito:
                            exito = guardar_continuar()
                        if exito:
                            exito, mensaje = validar_asignacion_fechas()
                        if exito:
                            exito = confimar()    
                        if exito:
                            exito = limpiar()
                        
                        if not exito:
                            limpiar()
                            refrescar_detalle_programacion()
                            logging.error("Error en uno de los pasos del flujo de asignación de fechas.")
                            logging.error("Error en uno de los pasos del flujo de guardar la programación.")
                            df.at[index, 'Estado'] = "Error :" + str(mensaje)  # Registrar el error en el DataFrame
                            continue  # Passar al siguiente elemento del bucle
                    except Exception as e:
                        limpiar()
                        refrescar_detalle_programacion()
                        logging.error(f"Excepción durante la ejecución del flujo de asignación de fechas: {str(e)}")
                        df.at[index, 'Estado'] = f"Excepción en el flujo de asignación de fechas: {str(e)}"
                        logging.info(f"Saltando al siguiente registro debido a una excepción en el índice {index}.")
                        bandera = 0
                        continue   

                    # Marcar todos los índices del grupo como procesados
                    indices_procesados.update(grupo_complejo.index)
                    df.loc[grupo_complejo.index, 'Estado'] = "Insertado"
                    
                    
                else:
                    # Selección del checkbox
                    #time.sleep(1)
                    for idx in grupo_complejo.index:
                        fila = df.loc[idx]  # Acceder a la fila por su índice original
                        logging.info(f"Índice original: {idx}, Contenido de grupo_complejo: {fila['NOMBRE_COMPLEJO']}")
                        id_complejo_actual = fila['ID_COMPLEJO']
                        nombre_complejo_actual = fila['NOMBRE_COMPLEJO']
                        id_sala_actual = fila['ID_SALA']
                        fecha_exhi_inicial_actual = fila['FECHA_EXHI_INICIAL']  # Asegúrate de que esté en formato 'dd/mm/yyyy'
                        fecha_exhi_final_actual = fila['FECHA_EXHI_FINAL']  # Asegúrate de que esté en formato 'dd/mm/yyyy'
                        #df.loc[idx, 'Estado'] = nuevo_valor  # Actualizar el DataFrame original

                        try:
                            select_element = wait.until(
                                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cmbComplejo"))
                            )
                            select = Select(select_element)
                            try:
                                select.select_by_value(str(id_complejo_actual))  # Cambia id_complejo por el valor que corresponda si es necesario
                                #time.sleep(2)
                                logging.info(f"Opción seleccionada por valor correctamente {id_complejo_actual} del exhibidor {id_exhibidor}...")
                            except:
                                logging.info(f"Seleccionando por texto visible: {nombre_complejo_actual}")
                                select.select_by_visible_text(nombre_complejo_actual)
                                time.sleep(2)
                                logging.info("Opción seleccionada por texto visible correctamente.")

                        except Exception as e:
                            logging.error(f"Error al seleccionar la opción en el menú desplegable: {e}")
                            logging.error(f"Error index {index} NO procesado - No Selecciono complejo .")
                            df.at[index, 'Estado'] = f"Error index {index} NO procesado - No Selecciono complejo ."
                            bandera = 0
                            continue  # Pasa al siguiente registro en el bucle principal

                        # Ingresar la fecha en el campo de Fecha de Inicio
                        time.sleep(1)
                        try:
                            fecha_inicio_input = wait.until(
                                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaInicio"))
                            )
                            fecha_inicio_input.clear()  # Limpia el campo antes de ingresar la fecha
                            #time.sleep(1)
                            fecha_inicio_input.send_keys(fecha_exhi_inicial_actual.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                            fecha_inicio_input.send_keys(Keys.ENTER)
                            logging.info("Fecha de inicio ingresada correctamente.")
                        except Exception as e:
                            logging.error(f"Error al ingresar la fecha de inicio: {str(e)}")
                            df.at[index, 'Estado'] = f"Error al ingresar la fecha de inicio: {str(e)}"  # Registrar el error en el DataFrame
                            logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                            bandera=0
                            continue  # Pasar al siguiente elemento del bucle


                        # Ingresar la fecha en el campo de Fecha de Fin
                        time.sleep(1)
                        try:
                            fecha_fin_input = wait.until(
                                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaFin"))
                            )
                            fecha_fin_input.clear()  # Limpia el campo antes de ingresar la fecha
                            #time.sleep(1)
                            fecha_fin_input.send_keys(fecha_exhi_final_actual.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                            fecha_fin_input.send_keys(Keys.ENTER)
                            logging.info("Fecha de fin ingresada correctamente.")
                        except Exception as e:
                            logging.error(f"Error al ingresar la fecha final: {str(e)}")
                            df.at[index, 'Estado'] = f"Error al ingresar la fecha final: {str(e)}"  # Registrar el error en el DataFrame
                            logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                            bandera=0
                            continue  # Pasar al siguiente elemento del bucle
                        try:
                            # Esperar a que el checkbox esté presente y seleccionarlo (reemplaza 'Autocine Mas SALA: 1' si es otro nombre)
                            id_sala = str(id_sala_actual)
                            # Construir el XPath como una cadena
                            xpath_checkbox = "//label[contains(text(), '" + nombre_complejo_actual + " SALA: " + id_sala + "')]/preceding-sibling::input[@type='checkbox']"

                            # Imprimir el XPath para verificar cómo se ve
                            logging.info(f"XPath construido: {xpath_checkbox}")
                            checkbox = wait.until(
                                EC.element_to_be_clickable((By.XPATH, xpath_checkbox))
                            )
                            checkbox.click()
                            #time.sleep(2)
                            logging.info("Checkbox seleccionado.")
                        except Exception as e:
                            logging.error(f"Error al seleccionar el Checkbox: {str(e)}")
                            df.at[index, 'Estado'] = f"Error al seleccionar el Checkbox: {str(e)}"  # Registrar el error en el DataFrame
                            logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                            bandera=0
                            continue  # Pasar al siguiente elemento del bucle    

                        try:
                            exito = asignar_fechas()
                            if exito:
                                exito, mensaje = validar_asignacion_fechas()
                                if not exito:
                                    refrescar_detalle_programacion()
                            if exito:
                                exito = guardar_continuar()
                            if exito:
                                exito, mensaje = validar_asignacion_fechas()
                            if exito:
                                exito = confimar()    
                            if exito:
                                exito = limpiar()
                            
                            if not exito:
                                limpiar()
                                refrescar_detalle_programacion()
                                logging.error("Error en uno de los pasos del flujo de guardar la programación.")
                                df.at[index, 'Estado'] = "Error: " + str(mensaje)
                                
                                logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                                bandera = 0
                                continue  # Passar al siguiente elemento del bucle
                        except Exception as e:
                            limpiar()
                            refrescar_detalle_programacion()
                            logging.error(f"Excepción durante la ejecución del flujo de asignación de fechas: {str(e)}")
                            df.at[index, 'Estado'] = f"Excepción en el flujo de asignación de fechas: {str(e)}"
                            logging.info(f"Saltando al siguiente registro debido a una excepción en el índice {index}.")
                            bandera = 0
                            continue
                        indices_procesados.add(idx)
                        df.at[idx, 'Estado'] = "Insertado"
            # Marcar estado como Insertado
            if bandera == 1:
                logging.info(f"Termino el proceso para el exhibidor: {exhibidor} - index final {index} Procesado .")
                bandera = 0
            else:
                logging.error(f"Termino el proceso para el exhibidor: {exhibidor} - index {index} NO procesado - no tiene ID_COMPLEJO .")
            try:
                # Espera a que el botón de "Cancelar" esté presente y sea clickeable
                cancelar_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarDetProgramacionPopUp"))
                )
                #time.sleep(1)
                cancelar_button.click()
                #time.sleep(1)
                logging.info("Botón 'Cancelar' presionado.")
            except Exception as e:
                logging.error("Error al hacer clic en el botón 'Cancelar':", e)

            # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
            #time.sleep(1)
        except Exception as e:
            logging.error(f"Error al procesar el exhibidor {exhibidor}: {e}")
            df.at[index, 'Estado'] = f"Error: {str(e)}"  # Registrar el error en el DataFrame
            logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
            continue  # Pasar al siguiente elemento del bucle
            
        # Espera antes de continuar con la siguiente fila (si es necesario)
        #time.sleep(1)
        # Calcular el tiempo que tardó la iteración
        end_time = time.time()
        duration = end_time - start_time
        df.at[index, 'Tiempo'] = duration
        logging.info(f"Tiempo tomado para la fila {index}: {duration:.2f} segundos")

    # Crear la ruta de salida con un nombre dinámico
    base_name, _ = os.path.splitext(file_path)
    output_path = f"{base_name}_Salida.xlsx"
    df.to_excel(output_path, index=False)

    # Cerrar el navegador
    driver.quit()
    logging.info(f"Proceso completado. El archivo ha sido actualizado en {output_path}")
    logging.info("Navegador cerrado.")

    # Mensaje de finalización
    mensaje = f"Proceso completado.\nArchivo actualizado:\n{output_path}"
    mostrar_mensaje_exito(mensaje, "Proceso Terminado", output_path)