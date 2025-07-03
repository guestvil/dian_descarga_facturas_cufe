import pandas as pd
from patchright.sync_api import sync_playwright, Page
from dotenv import load_dotenv
import pymupdf
from datetime import datetime
from patchright._impl._errors import TimeoutError
import os
import json




def load_env_files():
    load_dotenv()
    dian_url = os.getenv('DIAN')
    return dian_url


def load_invoice_codes(file_path: str):
    '''Returns a list of invoices to retrieve from the DIAN page
    
    Arguments:
        file_path: an excel file
        
    Returns: 
        invoice_list: a list of invoices codes'''
    invoice_list = []
    # returns a pandas dataframe object that can be iterated over, the header=None means that columns do not have name
    df_excel = pd.read_excel(file_path, header=None, engine='openpyxl')
    # Let's now iterate over the excel file to get the codes from there. The index=True assigns an index to each column
    for row in df_excel.itertuples(index=True):
        invoice_list.append(row[2])
    # The first column in the excel is the column name, so we will remove it
    return invoice_list[1:]


def get_dian_pdfs(downloads_path: str, codes_list: list, dian_website: str, playwright_page: Page):
    '''Recieves a list of DIAN codes and retrieves the invoices from the website
    Arguments:
        downloads_path: directorio en que se descargarán las facturas de la DIAN
        codes_list: a list of invoices codes
        
    returns: nothing'''
    playwright_page.goto(dian_website)
    files_list = []
    failed_invoices = []
    for code in codes_list:
        # por algun motivo playwright no añade automáticamente la extensión .pdf a los archivos
        file_name = code + '.pdf'
        download_path = downloads_path
        # Se crean direcciones absolutas para cada uno de los archivos con que luego se guardarán
        file_path = download_path + '/' + file_name
        print('Loggin in DIAN website')
        # es mejor esperar porque la página a veces se demora
        playwright_page.wait_for_timeout(3000)
        # Este es el número de intentos que se harán de pasar el captcha
        attempts = 3
        # esta bandera nos dirá cuando una factura ha sido efectivamente descargada
        download_flag = False
        for attempt in range(attempts):
            try: 
                # Se ingresa el CUFE y se da clic en buscar
                playwright_page.get_by_placeholder('Ingrese el código CUFE o UUID').fill(code)
                playwright_page.get_by_role('button', name='Buscar').click()
                print('Se dio click en buscar')
                # Verificando si aparece el captcha
                if playwright_page.get_by_text('Falta Token de validación de captcha').is_visible():
                    # Si aparece se debe esperar, a veces toma tiempo pasarlo
                    playwright_page.wait_for_timeout(4000)
                    print('Error, esperando validación de captcha')
                    # Si el captcha no se pasa, dando click en "Buscar" suele reiniciarlo
                    playwright_page.get_by_role('button', name='Buscar').click()
                    print('Se dio click de nuevo a buscar')
                if playwright_page.get_by_text('Documento no encontrado en los registros de la DIAN.').is_visible():
                    print('El documento no está en los registros de la DIAN')
                    continue
                # Si el documento existe se descarga la factura y se le da el absolute path creado con file_path
                with playwright_page.expect_download() as download_info:
                    print('Downloading pdf')
                    playwright_page.get_by_role('link', name=' Descargar PDF ').click()
                downloaded_file = download_info.value
                downloaded_file.save_as(file_path)
                # Se cambia la vandera para indicar que el arvhivo fue descargado
                download_flag = True
                print('pdf downloaded')
                break
            except TimeoutError as e: 
                print(e)
                print('Algo pasó, intentanto de nuevo. Intento', attempt, '/', attempts)
                playwright_page.reload()
        # Se añade la ubicación del archivo a una lista para posterior uso. This appends relatives paths for the pdf reading function, I don't know why this works but providing absolute paths to the Downloads folder simply does not work
        if download_flag == True:
            files_list.append(download_path+ '/'+ file_name)
            playwright_page.get_by_role('link', name='Volver').click()
        # Si no se pudo ddescargar se añade el archivo a una lista de archivos fallidos
        else:
            print('No se pudo descargar: ', code)
            failed_invoices.append(code)
            playwright_page.goto(dian_website)
    print(f'No se pudieron descargar los siguientes cufes: {failed_invoices}')
    with open('files_paths.json', 'w', encoding='utf-8') as json_file:
        json.dump(files_list, json_file, indent=4)
    # Se retorna la lista con la ubicación abolusta de cada uno de las facturas descargadas
    return files_list


def get_payment_method(file_path_lists):    
    '''file_path_list: una lista con las direcciones absolutas de las facturas descargadas
    
    Returns: 

        file_payment_method: una lista de tuplas de la firma: (cufe, forma_de_pago)
    '''
    file_payment_method = []
    # Iterar sobre todos los archivos
    for pdf_path in file_path_lists:
        invoice_text = ''
        # Se toma cada factura y se le extrae el texto
        with pymupdf.open(pdf_path) as pdf:
            for page in pdf:
                invoice_text += page.get_text('text')
        invoice_text = invoice_text.split()
        try:
            # La información de pago se halla siempre después de la palabra "pago" en la factura
            forma_pago = invoice_text[invoice_text.index('pago:') + 1]
            # Se le quita el .pdf a la dirección de archivo y luego se toma la última parte del nombre después de /, que es el CUFE
            invoice = pdf_path.split('.')[0]
            invoice = invoice.split('/')[-1]
            # Se añaden el CUFE y la forma de pago a una tupla
            file_payment_method.append((invoice, forma_pago))
        except ValueError as e:
            print(e)
            continue
    # Se retorna la lista de tuplas conteniendo los cufes y su respectiva forma de pago
    return file_payment_method


def update_excel(file_path: str, invoices_tuples: list):
    '''Parametros: 
        file_path: dirección relativa del archivo de excel en que se encuentran los CUFES
        invoices_tuples: lista de tuplas que contienen: (cufe, forma de pago)

    Retorna: 

        Nada, la función sólo crea un nuevo excel con la información
    '''
    # Se crea el nombre para el nuevo excel
    new_path = 'updated' + file_path
    df_excel = pd.read_excel(file_path, engine='openpyxl')
    indexing = 0
    df_excel['pago'] = None
    for row in df_excel.itertuples(index=True):
        # Se itera sobre cada uno de los elementos en la segunda columna, que son los cufe
        print(row[2])
        try: 
            # se verifica que el cufe en la casilla actual y el que se tiene en la tupla sean el mismo
            if row[2] == invoices_tuples[indexing][0]:
                print(invoices_tuples[indexing][0])
                print(invoices_tuples[indexing][1])
                # se añade la forma de pago en la fila del respectivo cufe, en la columna 'pago'
                df_excel.at[row.Index, 'pago'] = invoices_tuples[indexing][1]
                indexing += 1
            else:
                # si no coinciden el cufe en el excel con el de la tupla se continúa, esto se debe a que en el excel puede haber cufes que no estaban en la DIAN o no se pudieron descargar
                continue
        except IndexError as e:
            print('Programa finalizado. Hay facturas del excel sin procesar')
            break
    # se guardan los cambios en el nuevo excel
    df_excel.to_excel(new_path, index=False, engine='openpyxl')
    print('Changes saved to excel file!')
    return None


def main():
    print(datetime.now().time())
    # directorio en que se encuentra el archivo con los códigos CUFE
    path = 'your_excel_file_here'
    new_path = 'updated_file.xlsx'
    dian_url = load_env_files()
    invoice_list = load_invoice_codes(path)
    invoices_not_read = []
    # generamos una instancia de playwright
    with sync_playwright() as playwright:
        # abrimos un navegador
        browser = playwright.chromium.launch_persistent_context(user_data_dir='', channel='chrome', headless=False, downloads_path='your_downloads_path_here', no_viewport=True, slow_mo=1000)
        # abrimos una nueva página
        dian_page = browser.new_page()
        # descargamos las facturas y recibimos la lista con la ubicación de cada factura
        pdfs_paths = get_dian_pdfs(codes_list=invoice_list, dian_website=dian_url, playwright_page=dian_page)
        # leemos los archivos y obtenemos el método de pago de cada CUFE
        payment_tuple = get_payment_method(pdfs_paths)
        for tuples in payment_tuple:
            if tuples[0] not in invoice_list:
                invoices_not_read.append(tuples[0])
        # el archivo excel es actualizado con la información de pago de cada factura
        update_excel(file_path=path, invoices_tuples=payment_tuple)
    # se imprime la hora de terminación, usado durante pruebas para determinar tiempo de ejecución
    print('Program successful', datetime.now().time())


if __name__ == '__main__':
    main()