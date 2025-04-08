import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
from difflib import SequenceMatcher
from PIL import Image, ImageTk
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
import threading
import os
import sys

usuario_pagina = None
contrasena_pagina = None
def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecciona el archivo excel",
        filetypes=(("Excel files", "*.xlsx"),)
    )
    return file_path

def seleccionar_hoja(libro_excel):
    hojas_disponibles = libro_excel.sheet_names
    print("Hojas disponibles en el archivo:")
    for idx, hoja in enumerate(hojas_disponibles, start=1):
        print(f"{idx}. {hoja}")
    
    while True:
        try:
            root = tk.Tk()
            root.withdraw()
            indice_hoja = simpledialog.askinteger("Input", f"Ingrese el número correspondiente a la hoja que desea seleccionar (1-{len(hojas_disponibles)}):", parent=root)
            if indice_hoja is None:
                raise ValueError("Entrada inválida. Por favor, ingrese un número.")
            if 1 <= indice_hoja <= len(hojas_disponibles):
                return hojas_disponibles[indice_hoja - 1]
            else:
                messagebox.showerror("Error", "Número fuera de rango. Inténtelo de nuevo.")
        except ValueError:
            messagebox.showerror("Error", "Entrada inválida. Por favor, ingrese un número.")

def Manejar_alertas(driver, timeout=3):
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alerta = driver.switch_to.alert
        alerta.accept()
        print("Alerta aceptada.")
    except TimeoutException:
        print("No hay alerta presente.")

#Esta funcion tiene como proposito en una pagina especifica elegir un autocompletado, ya que en esa pagina no funciona el solo poner el texto, hay que elegir un autocompletado
def Encontrar_autocompletado(options, text, threshold):
    Mejor_opcion = None
    alto_ratio = 0.0
    for option in options:
        ratio = SequenceMatcher(None, text, option.text).ratio()
        if ratio > alto_ratio:
            alto_ratio = ratio
            Mejor_opcion = option
    if alto_ratio >= threshold:
        return Mejor_opcion
    return None

#Esta funcion es para que el usuario pueda elegir la columna como un numero o como una letra
def index_columna(Numero_o_letra_Columna):
    if isinstance(Numero_o_letra_Columna, str):
        Numero_o_letra_Columna = Numero_o_letra_Columna.upper().strip()
        if Numero_o_letra_Columna.isdigit():
            return int(Numero_o_letra_Columna) -1 
        col_idx = 0
        for char in Numero_o_letra_Columna:
            if 'A' <= char <= 'Z':
                col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
            else:
                raise ValueError("El nombre de la columna debe estar compuesto solo por letras.")
        return col_idx
    elif isinstance(Numero_o_letra_Columna, int):
        if Numero_o_letra_Columna > 0:
            return Numero_o_letra_Columna -1 
        else:
            raise ValueError("El índice de columna debe ser un número entero positivo.")
    else:
        raise ValueError("El índice de columna debe ser una cadena o un número entero.")
    

def preguntar_columna(nombre_columna, solo_rut=False):
    if solo_rut and nombre_columna != "RUT":
        return None
    while True:
        root = tk.Tk()
        root.withdraw()
        columna_input = simpledialog.askstring("Input", f"¿En qué columna se encuentra {nombre_columna}? (Letra o número)", parent=root)
        if columna_input is None:
            messagebox.showerror("Error", "Entrada inválida. Por favor, ingrese una letra o un número.")
            continue
        try:
            col_idx = index_columna(columna_input.strip().upper())
            return col_idx
        except ValueError:
            messagebox.showerror("Error", "Entrada inválida. Por favor, ingrese una letra o un número.")

def Seleccionar_Autocompletado(input_field, text, wait, ul_xpath, ul_xpath_secondary=None):
    input_field.send_keys(text)
    time.sleep(1)
    try:
        xpath_to_use = ul_xpath_secondary if ul_xpath_secondary else ul_xpath
        options = wait.until(EC.visibility_of_all_elements_located((By.XPATH, xpath_to_use)))
        for option in options:
            if option.text.strip().lower() == text.strip().lower():
                option.click()
                return True
    except TimeoutException:
        pass
    return False

def Seleccionar_Autocompletado_Int(input_field, integer_value, wait, ul_xpath):
    input_field.send_keys(str(integer_value))
    time.sleep(1)
    try:
        options = wait.until(EC.visibility_of_all_elements_located((By.XPATH, ul_xpath)))
        for option in options:
            if option.text.strip() == str(integer_value):
                option.click()
                return True
    except TimeoutException:
        pass
    return False


def Scrapeo_Primera_pagina(data, fila_inicio, fila_fin, columna_comuna, columna_calle, columna_numero, columna_rut, columna_vhfc,columna_hfc,columna_fibra,columna_deuda):
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    try:
        driver.get("Confidencial")
        Campo_Usuario = driver.find_element(By.ID, 'username')
        Campo_Usuario.send_keys(usuario_pagina)
        Campo_Contraseña = driver.find_element(By.ID, 'password')
        Campo_Contraseña.send_keys(contrasena_pagina)
        conectar_button = driver.find_element(By.ID, 'conectar')
        conectar_button.click()
        Click_Boton = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div[1]/a')))
        Click_Boton.click()
    except UnexpectedAlertPresentException as e:
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoAlertPresentException:
            pass
        finally:
            print("En el inicio de sesion probablemente se equivoco, o esta fuera de el horario en que funciona la pagina")
            pass
    
    for index, row in data.iloc[fila_inicio-2:fila_fin-1].iterrows():
        print(f"Procesando fila {index + 2}")
        Limpiar = wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div/div[2]/table[2]/tbody/tr/td/button[2]/span"))).click()
        time.sleep(1.5)

        try:
            # Manejo de la Comuna
            comuna_value = row.iloc[columna_comuna]
            if pd.isnull(comuna_value):
                print(f"Valor Comuna es nulo, pasando a la siguiente fila...")
                continue
            print(f"Valor Comuna: {comuna_value}")
            comuna_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/table[1]/tbody/tr[1]/td[2]/input[1]')))
            if not Seleccionar_Autocompletado(comuna_field, comuna_value, wait, '/html/body/ul', '//html/body/ul[1]'):
                Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
                Manejar_alertas(driver)
                continue
            Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
            time.sleep(1)
            
            # Manejo de la Calle
            calle_value = row.iloc[columna_calle]
            if pd.isnull(calle_value):
                print(f"Valor Calle es nulo, pasando a la siguiente fila...")
                continue
            print(f"Valor Calle: {calle_value}")
            calle_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/table[1]/tbody/tr[2]/td[2]/input')))
            if not Seleccionar_Autocompletado(calle_field, calle_value, wait, '/html/body/ul[2]'):
                Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
                Manejar_alertas(driver)
                continue
            Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
            time.sleep(1)
            
            # Manejo del Número
            numero_value = row.iloc[columna_numero]
            if pd.isnull(numero_value):
                print(f"Valor Número es nulo, pasando a la siguiente fila...")
                continue
            print(f"Valor Número: {numero_value}")
            numero_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/table[1]/tbody/tr[3]/td[2]/input')))
            if not Seleccionar_Autocompletado_Int(numero_field, numero_value, wait, '/html/body/ul[3]'):
                Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
                Manejar_alertas(driver)
                continue
            Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()
            time.sleep(1)

            # Manejo del RUT
            rut_value = row.iloc[columna_rut]
            if pd.isnull(numero_value):
                print(f"Valor Rut es nulo, pasando a la siguiente fila...")
                continue
            print(f"Valor RUT: {rut_value}")
            rut_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/table[1]/tbody/tr[4]/td[2]/input')))
            rut_field.send_keys(rut_value)
            Body = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]"))).click()

            Busqueda = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[2]/div/div[2]/table[2]/tbody/tr/td/button[1]/span[2]"))).click()
            Manejar_alertas(driver)
            
            try:
                data.at[index, columna_vhfc] = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/div/div/table/tbody/tr[3]/td').text
                print(f"Extrajo {columna_vhfc} para RUT {rut_value}: {data.at[index, columna_vhfc]}")
            except NoSuchElementException:
                data.at[index, columna_vhfc] = "Sin Datos"
                print(f"No se encontró el dato {columna_vhfc} para RUT {rut_value}")
            except Exception as e:
                print(f"Error en columna_riesgo para RUT {rut_value}: {str(e)}")

            try:
                data.at[index, columna_hfc] = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/div/div/table/tbody/tr[4]/td').text
                print(f"Extrajo {columna_hfc} para RUT {rut_value}: {data.at[index, columna_hfc]}")
            except NoSuchElementException:
                data.at[index, columna_hfc] = "Sin Datos"
                print(f"No se encontró el dato {columna_hfc} para RUT {rut_value}")
            except Exception as e:
                print(f"Error en columna_riesgo para RUT {rut_value}: {str(e)}")

            try:
                data.at[index, columna_fibra] = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/div/div/table/tbody/tr[5]/td').text
                print(f"Extrajo {columna_fibra} para RUT {rut_value}: {data.at[index, columna_fibra]}")
            except NoSuchElementException:
                data.at[index, columna_fibra] = "Sin Datos"
                print(f"No se encontró el dato {columna_fibra} para RUT {rut_value}")
            except Exception as e:
                print(f"Error en columna_riesgo para RUT {rut_value}: {str(e)}")

            try:
                data.at[index, columna_deuda] = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/div/div/table/tbody/tr[6]/td').text
                print(f"Extrajo {columna_deuda} para RUT {rut_value}: {data.at[index, columna_deuda]}")
            except NoSuchElementException:
                data.at[index, columna_deuda] = "Sin Datos"
                print(f"No se encontró el dato {columna_deuda} para RUT {rut_value}")
            except Exception as e:
                print(f"Error en columna_riesgo para RUT {rut_value}: {str(e)}")

            print(f"Vhfc: {columna_vhfc}, Hfc: {columna_hfc}, Fibra: {columna_fibra}, Deuda: {columna_deuda}")

        finally:
            # Cierra ventana y limpia para la siguiente iteración
            try:
                Cierre_ventana = wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[4]/div[1]/button/span[1]"))).click()
            except:
                pass
            Limpiar = wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div/div[2]/table[2]/tbody/tr/td/button[2]/span"))).click()
    
    # Guardar el archivo Excel con los resultados
    output_file_path = filedialog.asksaveasfilename(
        title="Guardar archivo con resultados",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"),)
    )
    if output_file_path:
        data.to_excel(output_file_path, index=False)
        print(f"Resultados guardados en {output_file_path}")

    driver.quit()

def Esperar_Busqueda(driver, xpath, timeout=30):
    """Espera hasta que el campo con el estilo cambie a display: none;."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda driver: driver.find_element(By.XPATH, xpath).get_attribute('style') == 'display: none;'
        )
    except TimeoutException:
        print("Timeout: El campo de búsqueda no cambió al estado esperado.")

def Scrapeo_Segunda_pagina(data, fila_inicio, fila_fin, columna_rut, rut_formato, columna_razon_social, columna_segmento, columna_subrango):
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    # Abre la página de inicio de sesión
    driver.get("Confidencial")

    # Realiza el inicio de sesión
    Campo_Usuario = driver.find_element(By.XPATH, '/html/body/div[1]/div/form/p[1]/input')
    Campo_Usuario.send_keys(usuario_pagina)

    Campo_Contraseña = driver.find_element(By.XPATH, '/html/body/div[1]/div/form/div/div/input')
    Campo_Contraseña.send_keys(contrasena_pagina)

    inicio_sesion = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/form/p[3]/input[1]"))).click()

    # Abre la URL de segmentación en la misma pestaña
    time.sleep(1.5)
    driver.get("Confidencial")

    # Añadir las nuevas columnas si no existen
    if columna_razon_social not in data.columns:
        data[columna_razon_social] = ""
    if columna_segmento not in data.columns:
        data[columna_segmento] = ""
    if columna_subrango not in data.columns:
        data[columna_subrango] = ""

    # Ajuste: Ciclo a través del rango de filas especificado
    for index, row in data.iloc[fila_inicio-2:fila_fin-1].iterrows():
        print(f"Procesando fila {index + 2}")
        time.sleep(1.5)
        rut_value = str(row.iloc[columna_rut])  # Asegura que rut_value es una cadena
        if rut_formato == 1:
            rut_value = rut_value[:-1]  # Elimina el último dígito del RUT
        else:
            rut_value = rut_value[:-2]
        print(f"Valor RUT: {rut_value}")

        rut_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div/div/article/div/div/div/div/div[2]/label/input')))
        
        # Obtén el valor actual del campo antes de enviar el RUT
        previous_value = rut_field.get_attribute('value')
        
        rut_field.clear()
        rut_field.send_keys(rut_value)
        rut_field.send_keys(u'\ue007')  # Presiona Enter

        # Espera a que el campo de búsqueda termine (display: none)
        Esperar_Busqueda(driver, '/html/body/div[2]/div[2]/div/div/article/div/div/div/div/div[3]')

        # Extrae los datos de la página
        try:
            razon_social = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/article/div/div/div/div/table/tbody/tr/td[3]').text
            segmento = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/article/div/div/div/div/table/tbody/tr/td[4]').text
            subrango = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/article/div/div/div/div/table/tbody/tr/td[5]').text
            print(f"Razón Social: {razon_social}, Segmento: {segmento}, Subrango: {subrango}")

            # Guarda los resultados en las columnas especificadas
            data.at[index, columna_razon_social] = razon_social
            data.at[index, columna_segmento] = segmento
            data.at[index, columna_subrango] = subrango
        except NoSuchElementException:
            data.at[index, columna_razon_social] = "Sin Datos"
            data.at[index, columna_segmento] = "Sin Datos"
            data.at[index, columna_subrango] = "Sin Datos"

    # Guardar el archivo Excel con los resultados
    output_file_path = filedialog.asksaveasfilename(
        title="Guardar archivo con resultados",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"),)
    )
    os.system("cls")
    print("Se esta guardando el archivo por favor espere...")
    if output_file_path:
        data.to_excel(output_file_path, index=False)
        print(f"Resultados guardados en {output_file_path}")

    # Cierra el navegador
    driver.quit()

def Scrapeo_Tercera_pagina(data, fila_inicio, fila_fin, columna_rut, columna_riesgo, columna_protestos, columna_dicom_total, columna_fecha_evaluacion, columna_origen, columna_deuda_telcom, columna_fecha_vencimiento, columna_deuda_total_vendida, columna_lineas_dormidas, columna_q_lineas_dormidas, columna_q_total_de_lineas, columna_tamano_empresa, columna_tipo_cliente, columna_monto_recurrente_fijo, columna_nota_movil,columna_correo):
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    driver.get("Confidencial")

    time.sleep(1.5)
    Login_Usu = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/form/div[1]/input')
    Login_Usu.send_keys(usuario_pagina)
        
    Login_pass = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/form/div[2]/input')
    Login_pass.send_keys(contrasena_pagina)

    Login_Botton = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/form/div[4]/button')
    Login_Botton.click()
    time.sleep(1.5)

    driver.get("Confidencial")

    # Realiza el scraping en el rango de filas especificado
    for index, row in data.iloc[fila_inicio-2:fila_fin-1].iterrows():
        print(f"Procesando fila {index + 2}")
        time.sleep(1.5)
        
        rut_value = str(row.iloc[columna_rut])
        print(f"Valor RUT: {rut_value}")

        rut_field = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/form/div[1]/div/input')))
        rut_field.clear()
        rut_field.send_keys(rut_value)

        buscar_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/form/div[4]/button')))
        buscar_button.click()

        time.sleep(2)

        # Extrae la información deseada y la guarda en el DataFrame
        try:
            data.at[index, columna_riesgo] = driver.find_element(By.ID, 'mostrar_riesgo').text
            print(f"Extrajo {columna_riesgo} para RUT {rut_value}: {data.at[index, columna_riesgo]}")
        except NoSuchElementException:
            data.at[index, columna_riesgo] = "Sin Datos"
            print(f"No se encontró el dato {columna_riesgo} para RUT {rut_value}")
        except Exception as e:
            print(f"Error en columna_riesgo para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_protestos] = driver.find_element(By.ID, 'mostrar_protestos').text
            print(f"Extrajo {columna_protestos} para RUT {rut_value}: {data.at[index, columna_protestos]}")
        except NoSuchElementException:
            data.at[index, columna_protestos] = "Sin Datos"
            print(f"No se encontró el dato {columna_protestos} para RUT {rut_value}")
        except Exception as e:
            print(f"Error en columna_protestos para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_dicom_total] = driver.find_element(By.ID, 'mostrar_dicom').text
            print(f"Extrajo {columna_dicom_total} para RUT {rut_value}: {data.at[index, columna_dicom_total]}")
        except NoSuchElementException:
            data.at[index, columna_dicom_total] = "Sin Datos"
            print(f"No se encontró el dato {columna_dicom_total} para RUT {rut_value}")
        except Exception as e:
            print(f"Error en columna_dicom_total para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_deuda_telcom] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[1]/table[2]/tbody/tr[3]/td[4]/label').text
            print(f"Extrajo {columna_deuda_telcom} para RUT {rut_value}: {data.at[index, columna_deuda_telcom]}")
        except NoSuchElementException:
            data.at[index, columna_deuda_telcom] = "Sin Datos"
            print(f"No se encontró el dato {columna_deuda_telcom} para RUT {rut_value}")
        except Exception as e:
            print(f"Error en columna_deuda_telcom para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_fecha_evaluacion] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[2]/table/thead/tr[3]/td[1]/label').text
            print(f"Extrajo {columna_fecha_evaluacion} para RUT {rut_value}: {data.at[index, columna_fecha_evaluacion]}")
        except NoSuchElementException:
            data.at[index, columna_fecha_evaluacion] = "Sin Datos"
            print(f"No se encontro el dato {columna_fecha_evaluacion}")
        except Exception as e:
            print(f"Error en columna_fecha_evaluacion para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_origen] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[1]/td[1]/label').text
            print(f"Extrajo {columna_origen} para RUT {rut_value}: {data.at[index, columna_origen]}")
        except NoSuchElementException:
            data.at[index, columna_origen] = "Sin Datos"
            print(f"No se encontro el dato {columna_origen}")
        except Exception as e:
            print(f"Error en columna_origen para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_deuda_telcom] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[1]/td[2]/label').text
            print(f"Extrajo {columna_deuda_telcom} para RUT {rut_value}: {data.at[index, columna_deuda_telcom]}")
        except NoSuchElementException:
            data.at[index, columna_deuda_telcom] = "Sin Datos"
            print(f"No se encontro el dato {columna_deuda_telcom}")
        except Exception as e:
            print(f"Error en columna_deuda_telcom para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_fecha_vencimiento] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[1]/td[3]/label').text
            print(f"Extrajo {columna_fecha_vencimiento} para RUT {rut_value}: {data.at[index, columna_fecha_vencimiento]}")
        except NoSuchElementException:
            data.at[index, columna_fecha_vencimiento] = "Sin Datos"
            print(f"No se encontro el dato {columna_fecha_vencimiento}")
        except Exception as e:
            print(f"Error en columna_fecha_vencimiento para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_deuda_total_vendida] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[2]/td/label').text
            print(f"Extrajo {columna_deuda_total_vendida} para RUT {rut_value}: {data.at[index, columna_deuda_total_vendida]}")
        except NoSuchElementException:
            data.at[index, columna_deuda_total_vendida] = "Sin Datos"
            print(f"No se encontro el dato {columna_deuda_total_vendida}")
        except Exception as e:
            print(f"Error en columna_deuda_total_vendida para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_lineas_dormidas] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[3]/table/tbody/tr/td[1]').text
            print(f"Extrajo {columna_lineas_dormidas} para RUT {rut_value}: {data.at[index, columna_lineas_dormidas]}")
        except NoSuchElementException:
            data.at[index, columna_lineas_dormidas] = "Sin Datos"
            print(f"No se encontro el dato {columna_lineas_dormidas}")
        except Exception as e:
            print(f"Error en columna_lineas_dormidas para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_q_lineas_dormidas] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[3]/table/tbody/tr/td[2]').text
            print(f"Extrajo {columna_q_lineas_dormidas} para RUT {rut_value}: {data.at[index, columna_q_lineas_dormidas]}")
        except NoSuchElementException:
            data.at[index, columna_q_lineas_dormidas] = "Sin Datos"
            print(f"No se encontro el dato {columna_q_lineas_dormidas}")
        except Exception as e:
            print(f"Error en columna_q_lineas_dormidas para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_q_total_de_lineas] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[3]/table/tbody/tr/td[3]').text
            print(f"Extrajo {columna_q_total_de_lineas} para RUT {rut_value}: {data.at[index, columna_q_total_de_lineas]}")
        except NoSuchElementException:
            data.at[index, columna_q_total_de_lineas] = "Sin Datos"
            print(f"No se encontro el dato {columna_q_total_de_lineas}")
        except Exception as e:
            print(f"Error en columna_q_total_de_lineas para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_tamano_empresa] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[4]/table/tbody/tr[1]/td[2]').text
            print(f"Extrajo {columna_tamano_empresa} para RUT {rut_value}: {data.at[index, columna_tamano_empresa]}")
        except NoSuchElementException:
            data.at[index, columna_tamano_empresa] = "Sin Datos"
            print(f"No se encontro el dato {columna_tamano_empresa}")
        except Exception as e:
            print(f"Error en columna_tamano_empresa para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_tipo_cliente] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[4]/table/tbody/tr[2]/td[2]').text
            print(f"Extrajo {columna_tipo_cliente} para RUT {rut_value}: {data.at[index, columna_tipo_cliente]}")
        except NoSuchElementException:
            data.at[index, columna_tipo_cliente] = "Sin Datos"
            print(f"No se encontro el dato {columna_tipo_cliente}")
        except Exception as e:
            print(f"Error en columna_tipo_cliente para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_monto_recurrente_fijo] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[5]/table/tbody/tr/td/p').text
            print(f"Extrajo {columna_monto_recurrente_fijo} para RUT {rut_value}: {data.at[index, columna_monto_recurrente_fijo]}")
        except NoSuchElementException:
            data.at[index, columna_monto_recurrente_fijo] = "Sin Datos"
            print(f"No se encontro el dato {columna_monto_recurrente_fijo}")
        except Exception as e:
            print(f"Error en columna_monto_recurrente_fijo para RUT {rut_value}: {str(e)}")
        
        try:
            data.at[index, columna_nota_movil] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[4]/table/tbody/tr[3]/td[2]').text
            print(f"Extrajo {columna_nota_movil} para RUT {rut_value}: {data.at[index, columna_nota_movil]}")
        except NoSuchElementException:
            data.at[index, columna_nota_movil] = "Sin Datos"
            print(f"No se encontro el dato {columna_nota_movil}")
        except Exception as e:
            print(f"Error en columna_nota_movil para RUT {rut_value}: {str(e)}")

        try:
            data.at[index, columna_correo] = driver.find_element(By.XPATH, '/html/body/div[3]/div[4]/div[2]/div/div[2]/div[1]/div[1]/div[1]/table[1]/tbody/tr[2]/td[2]/label').text
            print(f"Extrajo {columna_correo} para RUT {rut_value}: {data.at[index, columna_correo]}")
        except NoSuchElementException:
            data.at[index, columna_correo] = "Sin Datos"
            print(f"No se encontro el dato {columna_correo}")
        except Exception as e:
            print(f"Error en columna_correo para RUT {rut_value}: {str(e)}")


        print(f"Datos procesados para RUT {rut_value}")

        time.sleep(2)  # Pequeña espera antes de procesar la siguiente fila
        driver.get("Confidencial")
        time.sleep(2)
    driver.quit()

    # Guardar el archivo Excel con los resultados
    output_file_path = filedialog.asksaveasfilename(
        title="Guardar archivo con resultados",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"),)
    )
    os.system("cls")
    print("Se esta guardando el archivo por favor espere...")
    if output_file_path:
        data.to_excel(output_file_path, index=False)
        print(f"Resultados guardados en {output_file_path}")

class RedirectText:
    def __init__(self, widget):
        self.widget = widget

    def write(self, message):
        self.widget.insert(tk.END, message)
        self.widget.yview(tk.END)

    def flush(self):
        pass

def main():
    global usuario_pagina, contrasena_pagina
    try:
        os.system("cls")

        root = tk.Tk()
        root.withdraw()
        opcion = simpledialog.askinteger("Input", "Seleccione la opción:\n1. Factibilidad Fijo\n2. Segmentacion\n3. Consulta riesgo convergente", parent=root)

        if opcion not in [1, 2, 3]:
            messagebox.showerror("Error", "Opción no válida. Por favor, ingrese 1, 2 o 3.")
            return

        if opcion in [1, 3]:
            pass
        else:
            rut_formato = simpledialog.askinteger("Input", "Seleccione el formato del rut:\n1. Rut Sin Guion\n2. Rut Con Guion", parent=root)
            if rut_formato not in [1, 2]:
                messagebox.showerror("Error", "Opción no válida. Por favor, ingrese 1 o 2.")
                return

        # Pedir usuario y contraseña
        usuario_pagina = simpledialog.askstring("Input", "Ingrese el Usuario de la página", parent=root)
        contrasena_pagina = simpledialog.askstring("Input", "Ingrese la Contraseña de la página", parent=root, show="*")

        if not usuario_pagina or not contrasena_pagina:
            messagebox.showerror("Error", "Usuario y contraseña son necesarios.")
            return

        # Seleccionar archivo Excel
        excel_file_path = seleccionar_archivo()

        # Leer libro Excel
        libro_excel = pd.ExcelFile(excel_file_path)

        # Seleccionar hoja
        hoja_seleccionada = seleccionar_hoja(libro_excel)
        data = pd.read_excel(excel_file_path, sheet_name=hoja_seleccionada)
        os.system("cls")

        # Mostrar columnas disponibles en la hoja seleccionada
        print("Columnas disponibles en la hoja seleccionada:")
        for idx, col in enumerate(data.columns, start=1):
            print(f"{idx}. {col}")

        # Preguntar al usuario por las columnas
        if opcion in [2, 3]:
            columna_rut = preguntar_columna("RUT", solo_rut=True)
        else:
            columna_comuna = preguntar_columna("Comuna")
            columna_calle = preguntar_columna("Calle")
            columna_numero = preguntar_columna("Número")
            columna_rut = preguntar_columna("RUT")

            # Verificar los índices de columna
            print(f"Índice columna Comuna: {columna_comuna}")
            print(f"Índice columna Calle: {columna_calle}")
            print(f"Índice columna Número: {columna_numero}")
            print(f"Índice columna RUT: {columna_rut}")

        # Preguntar al usuario por el rango de filas
        fila_inicio = simpledialog.askinteger("Input", "¿Cuál es la fila de inicio?", parent=root)
        fila_fin = simpledialog.askinteger("Input", "¿Cuál es la fila final?", parent=root)

        # Ejecutar la opción seleccionada
        if opcion == 1:
            columna_vhfc = "VHFC"
            columna_hfc = "HFC"
            columna_fibra = "FIBRA"
            columna_deuda = "Deuda"
            Scrapeo_Primera_pagina(data, fila_inicio, fila_fin, columna_comuna, columna_calle, columna_numero, columna_rut, columna_vhfc, columna_hfc, columna_fibra, columna_deuda)
        elif opcion == 2:
            columna_razon_social = "Razon Social"
            columna_segmento = "Segmento"
            columna_subrango = "Subrango"
            Scrapeo_Segunda_pagina(data, fila_inicio, fila_fin, columna_rut, rut_formato, columna_razon_social, columna_segmento, columna_subrango)
        elif opcion == 3:
            columna_riesgo = "Riesgo"
            columna_protestos = "Tiene Protestos"
            columna_dicom_total = "Dicom Total"
            columna_fecha_evaluacion = "Fecha Evaluación"
            columna_origen = "Origen"
            columna_deuda_telcom = "Deuda Telcom"
            columna_fecha_vencimiento = "Fecha Vencimiento"
            columna_deuda_total_vendida = "Deuda Total Vendida"
            columna_lineas_dormidas = "Porcentaje Líneas Dormidas"
            columna_q_lineas_dormidas = "Q Líneas Dormidas"
            columna_q_total_de_lineas = "Q Total De Líneas"
            columna_tamano_empresa = "Tamaño Empresa"
            columna_tipo_cliente = "Tipo Cliente"
            columna_monto_recurrente_fijo = "Monto Recurrente Fijo"
            columna_nota_movil = "Nota Movil"
            columna_correo = "Correo Empresa"
            Scrapeo_Tercera_pagina(data, fila_inicio, fila_fin, columna_rut, columna_riesgo, columna_protestos, columna_dicom_total, columna_fecha_evaluacion, columna_origen, columna_deuda_telcom, columna_fecha_vencimiento, columna_deuda_total_vendida, columna_lineas_dormidas, columna_q_lineas_dormidas, columna_q_total_de_lineas, columna_tamano_empresa, columna_tipo_cliente, columna_monto_recurrente_fijo, columna_nota_movil, columna_correo)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def ejecutar_proceso():
    threading.Thread(target=main).start()

def recurso_de_img(relative_path):
    """ Obtiene la ruta del archivo en un entorno de desarrollo o en un ejecutable. """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def crear_interfaz():
    root = tk.Tk()
    root.title("Luna Scraper v1.0")

    # Cargar la imagen
    img_path = recurso_de_img("vprime logo.png")
    img = Image.open(img_path)
    img = img.resize((150, 100), Image.LANCZOS)
    img_tk = ImageTk.PhotoImage(img)


    label_img = tk.Label(root, image=img_tk)
    label_img.image = img_tk
    label_img.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

    label_nombre = tk.Label(root, text="Luna Scraper", font=("Arial", 24, "bold"))
    label_nombre.grid(row=0, column=1, padx=10, pady=10, sticky="w")

    terminal = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20, width=80)
    terminal.grid(row=1, column=0, padx=10, pady=10, columnspan=2)

    redir = RedirectText(terminal)
    sys.stdout = redir

    btn_ejecutar_proceso = tk.Button(root, text="Ejecutar Proceso", command=ejecutar_proceso)
    btn_ejecutar_proceso.grid(row=2, column=0, pady=10, columnspan=2)

    root.mainloop()
if __name__ == "__main__":
    crear_interfaz()
