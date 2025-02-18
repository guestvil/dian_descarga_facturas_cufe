import pandas as pd
from patchright.sync_api import sync_playwright
from dotenv import load_dotenv
import pymupdf
import os




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


def get_dian_pdfs(codes_list: list, dian_website: str, playwright_page):
    '''Recieves a list of DIAN codes and retrieves the invoices from the website
    Arguments: 
        codes_list: a list of invoices codes
        
    returns: nothing'''
    playwright_page.goto(dian_website)
    files_list = []
    for code in codes_list:
        file_name = code + '.pdf'
        download_path = '/Users/guestvil/Downloads'
        file_path = download_path + '/' + file_name
        print('Loggin in DIAN website')
        playwright_page.wait_for_timeout(3000)
        playwright_page.get_by_placeholder('Ingrese el código CUFE o UUID').fill(code)
        # playwright_page.get_by_text('Success!').wait_for(state='visible')
        playwright_page.get_by_role('button', name='Buscar').click()
        # playwright_page.get_by_text('Success!').wait_for(state='visible')
        with playwright_page.expect_download() as download_info:
            print('Downloading pdf')
            playwright_page.get_by_role('link', name=' Descargar PDF ').click()
        downloaded_file = download_info.value
        downloaded_file.save_as(file_path)
        print('pdf downloaded')
        # This appends relatives paths for the pdf reading function, I don't know why this works but providing absolute paths to the Downloads folder simply does not work
        files_list.append(download_path+ '/'+ file_name)
        playwright_page.get_by_role('link', name='Volver').click()
    return files_list


def get_payment_method(file_path_lists):
    file_payment_method = []
    for pdf_path in file_path_lists:
        invoice_text = ''
        with pymupdf.open(pdf_path) as pdf:
            for page in pdf:
                invoice_text += page.get_text('text')
        invoice_text = invoice_text.split()
        forma_pago = invoice_text[invoice_text.index('pago:') + 1]
        invoice = pdf_path.split('.')[0]
        invoice = invoice.split('/')[-1]
        file_payment_method.append((invoice, forma_pago))
    return file_payment_method


def update_excel(file_path: str, invoices_tuples: list):
    new_path = 'updated' + file_path
    df_excel = pd.read_excel(file_path, engine='openpyxl')
    indexing = 0
    df_excel['pago'] = None
    for row in df_excel.itertuples(index=True):
        print(row[2])
        try: 
            if row[2] == invoices_tuples[indexing][0]:
                print(invoices_tuples[indexing][0])
                print(invoices_tuples[indexing][1])
                df_excel.at[row.Index, 'pago'] = invoices_tuples[indexing][1]
                indexing += 1
            else:
                continue
        except IndexError as e:
            print('Programa finalizado. Hay facturas del excel sin procesar')
            break
    df_excel.to_excel(new_path, index=False, engine='openpyxl')
    print('Changes saved to excel file!')
    return None


def main():
    path = '1_enero.xlsx'
    new_path = 'updated_file.xlsx'
    dian_url = load_env_files()
    invoice_list = load_invoice_codes(path)
    invoices_not_read = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch_persistent_context(user_data_dir='', channel='chrome', headless=False, downloads_path='/Users/guestvil/Downloads', no_viewport=True, slow_mo=1000)
        dian_page = browser.new_page()
        pdfs_paths = get_dian_pdfs(codes_list=invoice_list[0:2], dian_website=dian_url, playwright_page=dian_page)
        payment_tuple = get_payment_method(pdfs_paths)
        for tuples in payment_tuple:
            if tuples[0] not in invoice_list:
                invoices_not_read.append(tuples[0])
        update_excel(file_path=path, invoices_tuples=payment_tuple)
    print('Program successful')


if __name__ == '__main__':
    main()