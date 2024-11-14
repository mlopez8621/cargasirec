import sys
import os
import pandas as pd
import time
import ctypes
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from logging.handlers import RotatingFileHandler

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
    time.sleep(1)
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
    time.sleep(1)

    # Localizar el campo de usuario y contraseña y enviar las credenciales
    username_field = driver.find_element(By.ID, "txtUsuario")
    password_field = driver.find_element(By.ID, "txtPassword")

    # Ingresar las credenciales
    username = "yrojas"  # Nombre de usuario
    password = "Recaudo2019*"  # Contraseña

    username_field.send_keys(username)
    password_field.send_keys(password)

    # Localizar y hacer clic en el botón de ingreso
    login_button = driver.find_element(By.ID, "btnIngresar")
    login_button.click()

    # Esperar a que el menú de "SIREC" esté presente y hacer clic para expandirlo
    try:
        sirec_menu = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'SIREC')]"))
        )
        sirec_menu.click()
        time.sleep(1)
    except Exception as e:
        logging.error(f"Error al localizar el menú de SIREC: {e}")

    # Hacer clic en la sección "Programación de Cortometrajes" para expandir su menú
    try:
        programacion_menu = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Programación de Cortometrajes')]"))
        )
        programacion_menu.click()
        time.sleep(1)
    except Exception as e:
        logging.error(f"Error al localizar el menú de Programación de Cortometrajes: {e}")

    # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
    try:
        programacion_cortometrajes_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='/frm/peliculas/frmLstPEL_ENC_PROGRAMACION_CORTOS.aspx']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", programacion_cortometrajes_link)
        programacion_cortometrajes_link.click()
    except Exception as e:
        logging.error(f"Error al hacer clic en Programación de Cortometrajes: {e}")

    # Intentar cargar el archivo
    try:
        logging.info(f"Cargando archivo: {file_path}, hoja: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
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

    # Leer el archivo y especificar que la columna 'Acta' sea leída como texto
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'ACTA': str})
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'ANIO_BENEFICIO': str})
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'CERTIFICADO': str})
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'ID_SALA': str})

    # Conjunto de índices procesados
    indices_procesados = set()

    # Recorrer cada fila del archivo Excel y llenar el formulario
    for index, row in df.iterrows():

        # Registrar el inicio del tiempo de la iteración
        start_time = time.time()

        if index in indices_procesados:
            # Si el índice ya fue procesado, continuar con el siguiente
            logging.info(f"Índice {index} ya procesado. Saltando...")
            end_time = time.time()
            duration = end_time - start_time
            df.at[index, 'Tiempo'] = duration
            logging.info(f"Tiempo tomado para la fila {index + 1}: {duration:.2f} segundos")
            df.at[index, 'Estado'] = "Insertado"
            continue  # Omitir si ya fue procesado

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
            time.sleep(1)
            search_field = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Buscar']"))
            )
            search_field.send_keys(exhibidor)

            # Esperar y hacer clic en el botón "Editar"
            time.sleep(1)
            # Elimina los espacios en blanco al inicio y al final del nombre del exhibidor
            exhibidor = exhibidor.strip()

            # Construir el XPath dinámico con el valor de 'exhibidor'
            xpath_editar_button = "//tr[td[contains(text(), '" + exhibidor + "')]]//input[@value='Editar']"

            # Imprimir el XPath construido para verificar
            logging.info(f"XPath construido: {xpath_editar_button}")

            editar_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//tr[td[contains(text(), '" + exhibidor + "')]]//input[@value='Editar']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", editar_button)
            editar_button.click()

            # Navegar a Datos Programación
            time.sleep(1)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Datos Programación"))
            ).click()
            logging.info("Navegando a Datos Programación")

            time.sleep(1)
            driver.find_element(By.ID, "ContentPlaceHolder1_btnCrearProgramacion").click()
            logging.info("Botón Crear Programación presionado")

            # Ingresar el número de acta
            time.sleep(1)
            numero_acta = str(numero_acta)
            acta_field = driver.find_element(By.ID, "ContentPlaceHolder1_txtNroActa")
            acta_field.send_keys("00"+numero_acta)
            # Simular Enter para mover el foco
            acta_field.send_keys(Keys.ENTER)
            logging.info("Número de acta ingresado")

            # Seleccionar año y mes
            time.sleep(1)
            # Asegurarse de que anio_beneficio sea una cadena
            anio_beneficio = str(anio_beneficio)
            dropdown_anio = driver.find_element(By.ID, "ddlAnio")
            dropdown_anio.find_element(By.XPATH, "//option[. = '" + anio_beneficio + "']").click()
            logging.info("Año de beneficio seleccionado")

            time.sleep(1)
            dropdown_mes = driver.find_element(By.ID, "drpMesBeneficio")
            dropdown_mes.find_element(By.XPATH, "//option[. = '" + mes_beneficio + "']").click()
            logging.info("Mes de beneficio seleccionado")
            
            time.sleep(1)
            driver.find_element(By.ID, "ContentPlaceHolder1_btnGuardarYsalir").click()
            logging.info("Botón Guardar y salir presionado")

            # Esperar a que aparezca el mensaje de error o el mensaje de éxito
            time.sleep(1)
            try:
                error_message_element = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_lblErrorProgramacion"))
                )
                error_message = error_message_element.text
                if "Solo puede crear una programación por periodo" in error_message:
                    cancelar_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                    )
                    cancelar_button.click()
                    logging.error("Error detectado: Acta ya existente. Cancelando la operación.")
                if "Ya exhibio esta acta, solo puede exhibir esta acta maximo 1 vez" in error_message:
                    cancelar_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarProgramacionPopUp"))
                    )
                    cancelar_button.click()
                    logging.error("Error detectado: Acta ya exhibida máximo una vez. Cancelando y avanzando al siguiente registro.")
            
                    # Marcar el error en la columna Estado y continuar con el próximo registro
                    df.at[index, 'Estado'] = f"Error: {error_message}"
                    volver_a_programacion_cortometrajes()
                    continue  # Pasa al siguiente registro en el bucle principal
            except Exception as e:
                # Si no se encuentra el mensaje de error, asumimos que todo está correcto
                logging.info("No se encontró mensaje de error; asumiendo que se guardó correctamente.")
                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".confirm"))
                ).click()
                logging.info("Acta Guardada Exitosamente.")

            try:
                # Localizar el campo de entrada por su ID y enviar el valor "00"+numero_acta
                nro_acta_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_DXFREditorcol0_I"))
                )
                nro_acta_field.clear()  # Opcional: Limpia el campo antes de escribir
                nro_acta_field.send_keys("00"+numero_acta)
                logging.info("Número de acta ingresado correctamente en el campo de programación.")
            except Exception as e:
                logging.error("Error al ingresar el número de acta:", e)

            try:
                # Espera adicional para asegurar que el botón esté cargado
                time.sleep(1)
                
                # Desplazar a la vista y hacer clic en el botón "Crear detalle" por su ID
                crear_detalle_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_cell0_7_btnNuevoDetProgramacion_0"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", crear_detalle_button)
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_grdDatosProgramacion_cell0_7_btnNuevoDetProgramacion_0"))).click()
                logging.info("Botón 'Crear detalle' presionado exitosamente.")
            except Exception as e:
                logging.error("Error al hacer clic en el botón 'Crear detalle':", e)        


            # Agrupar por ID_EXHIBIDOR y luego iterar por cada ID_COMPLEJO asociado
            grupo_exhibidor = df[df['ID_EXHIBIDOR'] == id_exhibidor]  # Filtrar el grupo del exhibidor actual

            for id_complejo_actual, grupo_complejo in grupo_exhibidor.groupby('ID_COMPLEJO'):
                logging.info(f"Procesando complejo {id_complejo_actual} del exhibidor {id_exhibidor}...")
                try:
                    # Espera a que el elemento <select> esté presente
                    select_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cmbComplejo"))
                    )

                    # Inicializa el objeto Select
                    select = Select(select_element)

                    # Intentar seleccionar por valor
                    try:
                        select.select_by_value(id_complejo_actual)  # Cambia id_complejo por el valor que corresponda si es necesario
                        logging.info("Opción seleccionada por valor correctamente.")
                    except:
                        # Si falla, intenta seleccionar por texto visible
                        select.select_by_visible_text(grupo_complejo.iloc[0]['NOMBRE_COMPLEJO'])
                        logging.info("Opción seleccionada por texto visible correctamente.")

                except Exception as e:
                    logging.error("Error al seleccionar la opción en el menú desplegable:", e)

                # Ingresar la fecha en el campo de Fecha de Inicio
                time.sleep(1)
                try:
                    fecha_inicio_input = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaInicio"))
                    )
                    fecha_inicio_input.clear()  # Limpia el campo antes de ingresar la fecha
                    fecha_inicio_input.click()  # Hacer clic en el campo por si hay alguna acción de apertura de calendario
                    fecha_inicio_input.send_keys(fecha_exhi_inicial.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                    time.sleep(1)
                    fecha_inicio_input.send_keys(Keys.ENTER)
                    logging.info("Fecha de inicio ingresada correctamente.")
                except Exception as e:
                    logging.error(f"Error al ingresar la fecha de inicio: {str(e)}")
                    df.at[index, 'Estado'] = f"Error: {str(e)}"  # Registrar el error en el DataFrame
                    logging.info(f"Saltando al siguiente registro debido a un error en el índice {index}.")
                    continue  # Pasar al siguiente elemento del bucle


                # Ingresar la fecha en el campo de Fecha de Fin
                time.sleep(1)
                try:
                    fecha_fin_input = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_txtFechaFin"))
                    )
                    fecha_fin_input.clear()  # Limpia el campo antes de ingresar la fecha
                    fecha_fin_input.click()  # Hacer clic en el campo para asegurarse de que esté en foco
                    fecha_fin_input.send_keys(fecha_exhi_final.strftime("%d/%m/%Y"))  # Ingresa la fecha en formato 'dd/mm/yyyy'
                    time.sleep(1)
                    fecha_fin_input.send_keys(Keys.ENTER)
                    logging.info("Fecha de fin ingresada correctamente.")
                except Exception as e:
                    logging.error("Error al ingresar la fecha de fin:", e)

                try:

                    # Filtrar el grupo de salas por el mismo ID_COMPLEJO
                    grupo_complejo = df[df['ID_COMPLEJO'] == id_complejo_actual]
                    fechas_iniciales_iguales = grupo_complejo['FECHA_EXHI_INICIAL'].nunique() == 1
                    fechas_finales_iguales = grupo_complejo['FECHA_EXHI_FINAL'].nunique() == 1

                    if fechas_iniciales_iguales and fechas_finales_iguales:
                        # Seleccionar todos si las fechas son iguales para el mismo complejo
                        logging.info(f"Aplicando 'Seleccionar todos' para el complejo {grupo_complejo.iloc[0]['NOMBRE_COMPLEJO']}...")
                        # Construir el XPath como una cadena
                        xpath_checkbox = "//input[@id='ContentPlaceHolder1_chkSeleccionarTodo' and @type='checkbox']"

                        # Imprimir el XPath para verificar cómo se ve
                        logging.info(f"XPath construido: {xpath_checkbox}")
                        checkbox = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_checkbox))
                        )
                        checkbox.click()
                        logging.info(f"'Seleccionar todos' aplicado correctamente para el complejo {grupo_complejo.iloc[0]['NOMBRE_COMPLEJO']}.")

                        # Marcar todos los índices del grupo como procesados
                        indices_procesados.update(grupo_complejo.index)
                    else:
                        # Selección del checkbox
                        time.sleep(1)
                        # Esperar a que el checkbox esté presente y seleccionarlo (reemplaza 'Autocine Mas SALA: 1' si es otro nombre)
                        id_sala = str(id_sala)
                        # Construir el XPath como una cadena
                        xpath_checkbox = "//label[contains(text(), '" + grupo_complejo.iloc[0]['NOMBRE_COMPLEJO'] + " SALA: " + id_sala + "')]/preceding-sibling::input[@type='checkbox']"

                        # Imprimir el XPath para verificar cómo se ve
                        logging.info(f"XPath construido: {xpath_checkbox}")
                        checkbox = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_checkbox))
                        )
                        checkbox.click()
                        logging.info("Checkbox seleccionado.")
                except Exception as e:
                    logging.error("Error en la selección de checkbox", e)

                # Hacer clic en "Asignar Fechas"
                try:
                    time.sleep(1)
                    asignar_fechas_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAgregarFechas"))
                    )
                    asignar_fechas_button.click()
                    logging.info("Botón Asignar Fechas presionado.")
                except Exception as e:
                    logging.info("Asignación de fechas:", e)        
                
                # Hacer clic en "Guardar y continuar"
                try:
                    time.sleep(1)
                    guardar_continuar_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnGuardarDetProgramacion"))
                    )
                    guardar_continuar_button.click()
                    logging.info("Botón Guardar y continuar presionado.")
                    time.sleep(1)
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".confirm"))
                    ).click()
                    logging.info("Programacion guardada")
                except Exception as e:
                    logging.info("Guardar programacion:", e)
                
                try:
                    # Localizar el checkbox "Seleccionar todos"
                    time.sleep(2)
                    xpath_checkbox_select_all = "//input[@id='ContentPlaceHolder1_chkSeleccionarTodo' and @type='checkbox']"
                    checkbox_select_all = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, xpath_checkbox_select_all))
                    )
                    logging.info("'Seleccionar todos' está presente.")

                    # Verificar si el checkbox está marcado
                    is_checked = driver.execute_script("return arguments[0].checked;", checkbox_select_all)

                    if is_checked:
                        logging.info("El checkbox 'Seleccionar todos' está marcado. Procediendo a desmarcarlo.")
                        checkbox_select_all.click()
                        logging.info("Checkbox 'Seleccionar todos' ahora está desmarcado.")
                    else:
                        logging.info("El checkbox 'Seleccionar todos' ya estaba desmarcado.")
                except Exception as e:
                    logging.error(f"Error al verificar el estado de 'Seleccionar todos': {e}")
                
                df.at[index, 'Estado'] = "Insertado"

            # Marcar estado como Insertado
            logging.info(f"Programación para {exhibidor} completada.")

            try:
                # Espera a que el botón de "Cancelar" esté presente y sea clickeable
                time.sleep(1)
                cancelar_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCancelarDetProgramacionPopUp"))
                )
                cancelar_button.click()
                logging.info("Botón 'Cancelar' presionado.")
            except Exception as e:
                logging.error("Error al hacer clic en el botón 'Cancelar':", e)

            # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
            time.sleep(1)
            volver_a_programacion_cortometrajes()


        except Exception as e:
            # Guardar el error en la columna Estado
            df.at[index, 'Estado'] = f"Error: {str(e)}"
            logging.error(f"Error al procesar el exhibidor {exhibidor}: {e}")

            # Hacer clic en el enlace "Programación de Cortometrajes" con desplazamiento previo
            time.sleep(1)
            volver_a_programacion_cortometrajes()
            
        # Espera antes de continuar con la siguiente fila (si es necesario)
        time.sleep(1)
        # Calcular el tiempo que tardó la iteración
        end_time = time.time()
        duration = end_time - start_time
        df.at[index, 'Tiempo'] = duration
        logging.info(f"Tiempo tomado para la fila {index + 1}: {duration:.2f} segundos")

    # Crear la ruta de salida con un nombre dinámico
    base_name, _ = os.path.splitext(file_path)
    output_path = f"{base_name}_Salida.xlsx"
    df.to_excel(output_path, index=False)

    # Aquí va toda la lógica del script...
    logging.info(f"Proceso completado. El archivo ha sido actualizado en {output_path}")

    # Cerrar el navegador
    driver.quit()
    logging.info(f"Proceso completado. El archivo ha sido actualizado en {output_path}")
    logging.info("Navegador cerrado.")

    # Mensaje de finalización
    mensaje = f"Proceso completado.\nArchivo actualizado:\n{output_path}"
    mostrar_mensaje_exito(mensaje, "Proceso Terminado", output_path)