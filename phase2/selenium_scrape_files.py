import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import numpy as np
import pandas as pd
from win32com.client import Dispatch


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

def convert_to_xls():
    # Create an Excel application instance
    xl = Dispatch('Excel.Application')
    # Open the XLSX file
    wb = xl.Workbooks.Add('E:\\University of Arizona\\Classes\\Spring 2024\\ATMO 392a\\Files\\Gs0000103.xlsx')
    # Save it as XLS format
    wb.SaveAs('E:\\University of Arizona\\Classes\\Spring 2024\\ATMO 392a\\Files\\Gs0000103.xls', FileFormat=56)
    # Close Excel
    xl.Quit()

def main():
    # setup
    cwd = os.getcwd()
    options=webdriver.ChromeOptions()
    prefs={"download.default_directory": cwd}
    options.add_experimental_option("prefs",prefs)
    driver=webdriver.Chrome(options=options)
    driver.implicitly_wait(30)

    df = pd.read_excel("./Files/Gs0000103.xls", index_col=None, header=None, sheet_name='Sheet1')

    search_using_excel("Gs0000103",driver)

    # for sample_id in unique_sample_IDs:
    #     if not os.path.exists(f"{sample_id}.xlsx"):
    #         print(f"Working on {sample_id}")
    #         search_using_excel(sample_id, driver)

    driver.quit()

main()
