import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
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
    for code in codes_list:
        file_name = 'dia_invoice.pdf'
        payment_method_list = []
        playwright_page.goto(dian_website)
        playwright_page.get_by_placeholder('Ingrese el código CUFE o UUID').fill(code)
        playwright_page.get_by_role('button', name='Buscar').click()
        with playwright_page.expect_download() as download_info:
            playwright_page.get_by_role('link', name='Descargar PDF')
        downloaded_file = download_info.value
        downloaded_file.save_as('/Users/guestvil/Downloads' + file_name)




def main():
    path = '1_enero.xlsx'
    dian_url = load_env_files()
    invoice_list = load_invoice_codes(path)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False, downloads_path='/Users/guestvil/Downloads', slow_mo=1000)
        dian_window = browser.new_context()
        dian_page = dian_window.new_page()
        get_dian_pdfs(codes_list=invoice_list, dian_website=dian_url, playwright_page=dian_page)
        



if __name__ == '__main__':
    main()