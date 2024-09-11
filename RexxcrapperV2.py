
#  .d8888b.        d8888 8888888 8888888b.  8888888b.  8888888888 Y88b   d88P Y88b   d88P 
# d88P  Y88b      d88888   888   888  "Y88b 888   Y88b 888         Y88b d88P   Y88b d88P  
# Y88b.          d88P888   888   888    888 888    888 888          Y88o88P     Y88o88P   
#  "Y888b.      d88P 888   888   888    888 888   d88P 8888888       Y888P       Y888P    
#     "Y88b.   d88P  888   888   888    888 8888888P"  888           d888b       d888b    
#       "888  d88P   888   888   888    888 888 T88b   888          d88888b     d88888b   
# Y88b  d88P d8888888888   888   888  .d88P 888  T88b  888         d88P Y88b   d88P Y88b  
#  "Y8888P" d88P     888 8888888 8888888P"  888   T88b 8888888888 d88P   Y88b d88P   Y88b 
                                                                                        
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import pandas as pd
import os

def login(driver):
    # Carga las variables de entorno desde el archivo .env
    load_dotenv('.env')

    # Accede a las variables de entorno cargadas
    username = os.getenv('ADMIN_USERNAME')
    password = os.getenv('ADMIN_PASSWORD')
    adminLoginURL = os.getenv('ADMINISTRATOR_LOGIN_URL')

    # Abre la página de inicio de sesión
    driver.get(adminLoginURL)

    # Encuentra los campos de nombre de usuario y contraseña y realiza el inicio de sesión
    username_field = driver.find_element(By.NAME, 'usuario_activo')
    password_field = driver.find_element(By.NAME, 'contrasenia')

    # Ingresa las credenciales de inicio de sesión
    username_field.send_keys(username)
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    # Espera hasta que el botón de enviar esté presente
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[8]/ul/div/form/i/input')))

def scrape_data(driver):
    # Encuentra y hace clic en el botón de enviar
    boton = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[8]/ul/div/form/i/input')
    boton.click()

    # Espera hasta que BeautifulSoup recupere todos los elementos <tr> y su contenido <td> de <tbody>
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))

    # Recupera todos los elementos <tr> y su contenido <td> de <tbody>
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table_rows = soup.select('tbody tr')

    # Inicializa una lista vacía para almacenar los datos
    data = []

    # Itera sobre los elementos <tr> y su contenido <td>
    for row in table_rows:
        cells = row.select('td')
        # Procesa los datos en cada celda
        for cell in cells:
            # Guarda el texto de cada celda en un archivo de texto
            with open('resultados.txt', 'a', encoding='utf-8') as file:
                file.write(cell.get_text() + '\n')

    # Lee el archivo de texto y procesa los datos en bloques de 7 líneas
    with open('resultados.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        for i in range(0, len(lines), 7):
            try:
                entry = {
                    "Número de Control": lines[i].strip(),
                    "Nombre": lines[i+1].strip(),
                    "Grupo": lines[i+2].strip(),
                    "Carrera": lines[i+3].strip(),
                    "Materia": lines[i+4].strip(),
                    "Clave de Materia": lines[i+5].strip(),
                    "Docente Responsable": lines[i+6].strip()
                }
                # Verifica si la entrada ya existe en la lista de datos
                if entry not in data:
                    data.append(entry)
            except IndexError:
                # Maneja cualquier error de índice que pueda ocurrir si las líneas no están en el formato esperado
                print("Error: Formato de datos inválido")

    return data

def save_to_excel(data):
    # Convierte los datos en un DataFrame de pandas
    df = pd.DataFrame(data)

    # Crea un libro de Excel
    writer = pd.ExcelWriter('archivo.xlsx', engine='xlsxwriter')

    # Itera sobre los grupos únicos y crea una hoja para cada grupo
    for group in df['Grupo'].unique():
        group_df = df[df['Grupo'] == group]
        group_df.to_excel(writer, sheet_name=group, index=False)

    # Guarda el libro de Excel
    writer.close()

def main():
    # Configura el controlador web (por ejemplo, Chrome)
    driver = webdriver.Edge()

    # Inicia sesión
    login(driver)

    # Extrae los datos
    data = scrape_data(driver)

    # Guarda los datos en un archivo de Excel
    save_to_excel(data)

    # Cierra el controlador web
    driver.quit()

if __name__ == "__main__":
    main()
