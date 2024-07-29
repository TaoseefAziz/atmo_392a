import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import numpy as np
import pandas as pd
from win32com.client import Dispatch
import re
from pandas.tseries.api import guess_datetime_format
from selenium.webdriver.common.keys import Keys
import requests
import warnings

def get_search_keywords(sample_id, driver):
    pass

def search_using_excel(sample_id, driver):
    '''Genome Name / Sample Name'''
    name = '''Cellulose adapted compost microbial communities from Newby Island Compost Facility, Milpitas, CA, USA - Passage2 60B (SPAdes) (version 2)'''
    # Cellulose adapted compost microbial communities from Newby Island Compost \
    # Facility, Milpitas, CA, USA - Passage2 60B (SPAdes) (version 2)

    # Cellulose adapted compost microbial communities from Newby Island Compost \
    # Facility, Milpitas, CA, USA - Passage2 37A (SPAdes) (version 3)

    '''Alternative2 Contact Names'''
    # Steven W. Singer

    '''Contact Name'''
    # Steven W. Singer

    '''Pubmed ID'''
    # ""
    # 25136443

    driver.get("https://scholar.google.com/")

    menu_button = driver.find_element(by=By.CSS_SELECTOR, value="#gs_hdr_mnu > .gs_ico")
    menu_button.click()

    adv_search_button = driver.find_element(by=By.CSS_SELECTOR, value="#gs_hp_drw_adv > .gs_lbl")
    time.sleep(1)
    at_least_text_box = driver.find_element(by=By.ID, value="gs_asd_oq")
    at_least_text_box.send_keys(name)
    time.sleep(1)
    author_box = driver.find_element(by=By.ID, value="gs_asd_sau")
    author_box.click()
    time.sleep(1)
    published_box = driver.find_element(by=By.ID, value="gs_asd_pub")
    published_box.click()
    time.sleep(1)
    year_lower = driver.find_element(by=By.ID, value="gs_asd_ylo")
    year_lower.click()
    time.sleep(1)
    year_higher = driver.find_element(by=By.ID, value="gs_asd_yhi")
    year_higher.click()

    time.sleep(1)

def convert_to_xls(decoded_cwd, sample_id):
    xlsx_path = os.path.join(decoded_cwd,   str(sample_id) + ".xlsx") 
    xls_path = os.path.join(decoded_cwd,    str(sample_id) + ".xls")

    if not os.path.exists(xls_path):
        # Create an Excel application instance
        xl = Dispatch('Excel.Application')

        # Open the XLSX file
        wb = xl.Workbooks.Add(xlsx_path)

        # Save it as XLS format
        wb.SaveAs(xls_path, FileFormat=56)

        # Close Excel
        xl.Quit()
    else:
        pass
        # print("Already converted ",end = "")
        # print(decoded_cwd + f"\\{sample_id}.xlsx --> xls")

def create_search_keywords(dataframe):


    # sample_name_column = dataframe["Genome Name / Sample Name"].unique()
    sample_name_simple_column = dataframe["Study Name"].unique()
    try:
        alt2_contact_email_column = dataframe["Alternative2 Contact Emails"].unique()
    except KeyError:
        alt2_contact_email_column = list()
    try:
        alt2_contact_name_column = dataframe["Alternative2 Contact Names"].unique()
    except KeyError:
        alt2_contact_name_column = list()
    try:
        contact_email_column = alt2_contact_name_column = dataframe["Contact Email"].unique()
    except KeyError:
        contact_email_column = list()
    try:
        contact_name_column = alt2_contact_name_column = dataframe["Contact Name"].unique()  
    except KeyError:
        contact_name_column = list()
    try:
        sample_coll_date_column = dataframe["Sample Collection Date"].unique()  
    except KeyError:
        sample_coll_date_column = list()
    try:
        pubmed_ids = dataframe["Pubmed ID"].unique().astype(int)
    except KeyError:
        pubmed_ids = list()
    try:
        min_funding_year = dataframe["Funding Year"].min()
    except KeyError:
        min_funding_year = 1900 
    
    min_coll_year = 1900

    if len(sample_coll_date_column) > 0:
        try: 
            format = guess_datetime_format(sample_coll_date_column[0])
            dataframe['new_sample_coll_date'] = pd.to_datetime(dataframe['Sample Collection Date'].astype(str), format =format, errors='coerce')
            min_coll_year = dataframe['new_sample_coll_date'].min().year
        except TypeError:
            pass

    sample_name_set = set()
    email_set = set()
    contact_name_set = set()

    for row in sample_name_simple_column:
        row_lst = row.split()
        for word in row_lst:
            sample_name_set.add(word.lower())

    for email_str in alt2_contact_email_column:
        emails = email_str.split(";")
        for email in emails:
            email_set.add(email)
    email_set.union(set(contact_email_column))

    for name_string in alt2_contact_name_column:
        names = name_string.split(";")
        for name in names:
            contact_name_set.add(name)
    contact_name_set.union(set(contact_name_column))

    min_year = min(min_funding_year, min_coll_year)

    pubmed_ids = set(pubmed_ids)
    pubmed_ids.discard(-2147483648)

    return sample_name_set, email_set, contact_name_set, min_year, list(pubmed_ids)

def download_from_url(download_dir, url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            # Extract the filename from the URL
            filename = os.path.basename(url)
            filepath = os.path.join(download_dir, filename)

            # Save the PDF to the specified directory
            with open(filepath, "wb") as f:
                f.write(response.content)
            print(f"Downloaded: {filename}")
        else:
            print(f"Error downloading {url}: Status code {response.status_code}")
    except Exception as e:
        print(f"Error downloading {url}: {str(e)}")

def download_scholar_results(query_string, driver, decoded_sample_dir, min_year):
    driver.get("https://scholar.google.com/schhp?hl=en&authuser=1#d=gs_asd&t=1711121138190")

    search_box = driver.find_element(by=By.ID, value="gs_asd_oq")
    year_low = driver.find_element(by=By.ID, value="gs_asd_ylo")
    year_low.send_keys(str(min_year))

    search_box.send_keys(query_string)
    search_box.send_keys(Keys.RETURN)

    # Find all PDF links on the first page of results
    pdf_link_objs = driver.find_elements(by=By.CSS_SELECTOR, value="a[href$='.pdf']")
    # By.PARTIAL_LINK_TEXT, value='pdf'
    # By.PARTIAL_LINK_TEXT, value='PDF'

    pdf_urls = []

    # Extract URLs
    for pdf_link in pdf_link_objs:
        pdf_urls.append(pdf_link.get_attribute("href"))

    print(pdf_urls)

    # Download all PDFs
    for pdf_url in pdf_urls:
        
        download_from_url(decoded_sample_dir, pdf_url)

def download_pubmed_results(pubmed_ids, driver, decoded_sample_dir,sample_id):
    pubmed_ids = list(pubmed_ids)
    # print(f"pubmed_ids = {pubmed_ids} len(pubmed_ids) = {len(pubmed_ids)}")
    if len(pubmed_ids) == 0:
        print(f"sample_id = {sample_id} does NOT have any pubmed IDs")
        return
    else:
        print(f"sample_id = {sample_id}   HAS pubmed IDs: {pubmed_ids}")

    for search_id in pubmed_ids:
        driver.get("https://pubmed.ncbi.nlm.nih.gov/")

        search_box = driver.find_element(by=By.ID, value="id_term")

        search_box.send_keys(str(search_id))
        search_box.send_keys(Keys.RETURN)

        source_text_button = driver.find_element(By.CSS_SELECTOR, ".full-text-links-list:nth-child(2) > .pmc")
        source_text_button.click()

        # add_html_links =(//a[contains(text(),'html')])
        # add_xls_links =(//a[contains(text(),'xls')])

        # main_pdf_links_bylinktext = driver.find_elements(By.PARTIAL_LINK_TEXT, 'PDF')

        pdf_links = driver.find_element(By.XPATH, "(a[contains(text(),'PDF')])[2]")
        print(f"got pdf_link objects")
        print(pdf_links)

        # # pdf_links = driver.find_elements(by=By.CSS_SELECTOR, value="a[href$='.pdf']")
        # pdf_urls = []

        # # Extract URLs
        # for pdf_link in pdf_links:
        #     pdf_urls.append(pdf_link.get_attribute("href"))

        # print(f"main pdf_urls = {pdf_urls}")
        
        # # Download all PDFs
        # for pdf_url in pdf_urls:
            
        #     download_from_url(decoded_sample_dir, pdf_url)

def main():
    warnings.filterwarnings("ignore")

    list_of_sample_ids = ["Gs0045634", "Gs0047444", "Gs0053055", "Gs0053071", "Gs0053075",
        "Gs0063124", "Gs0063182", "Gs0067859","Gs0071004", "Gs0071049", "Gs0075432",
        "Gs0084162" , "Gs0099864", "Gs0103000", "Gs0110126","Gs0110174","Gs0111452",
        "Gs0112340","Gs0114293", "Gs0114298","Gs0121480" ,"Gs0128849","Gs0128850",
        "Gs0128948","Gs0128997", "Gs0131983","Gs0134627"]

    # setup
    cwd = os.getcwd()
    decoded_cwd = os.fsdecode(cwd)
    options=webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver=webdriver.Chrome(options=options)
    driver.implicitly_wait(30)  
    
    driver.get("https://scholar.google.com/")
    # print(f"options.arguments = {options.arguments}")
    # print(f"options.experimental_options = {options.experimental_options}")
    # print(f"options.to_capabilities() = {options.to_capabilities()}")

    # time.sleep(60)

    for sample_id in list_of_sample_ids:
        # print(f"Working on {sample_id}")

        # change download location
        decoded_sample_dir = os.path.join(decoded_cwd, str(sample_id))
        encoded_sample_dir = os.fsencode(decoded_sample_dir)
        prefs={"download.default_directory": decoded_sample_dir}
        options.add_experimental_option("prefs",prefs)

        if not os.path.exists(f"{sample_id}"):
            os.makedirs(encoded_sample_dir)

        # convert_to_xls
        convert_to_xls(decoded_cwd, sample_id)

        # read xls into a dataframe
        xls_path = decoded_cwd + f"\\{sample_id}.xls"
        df = pd.read_excel(xls_path, sheet_name='Sheet1')
        sample_name_set, email_set, contact_name_set, min_year, pubmed_ids = create_search_keywords(df)
        query_string = " ".join(list(sample_name_set)) + " " + " ".join(list(contact_name_set)) + " " + " ".join(list(email_set))
        # print(f"query_string = {query_string}")

        # download_scholar_results(query_string, driver, decoded_sample_dir, min_year)
        download_pubmed_results(pubmed_ids, driver, decoded_sample_dir, sample_id)
        

    driver.quit()

main()
