import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import numpy as np
import pandas as pd

def create_empty_file(file_path):
    with open(file_path, 'a'):
        pass  # No need to close explicitly; Python's GC will handle it

def download_sample_excel(sample_id, driver):
    
    driver.get("https://img.jgi.doe.gov/cgi-bin/m/main.cgi?section=FindGenomes&page=genomeSearch")

    # select search box
    search_text_box = driver.find_element(by=By.ID, value="autocomplete")

    # type in sample number into box
    search_text_box.send_keys(sample_id)

    # Press the Go button next to the quick search box
    submit_button = driver.find_element(by=By.NAME, value="_section_TaxonSearch_x")
    submit_button.click()

    # Click the button with the number the number of results found
    try:
        all_button = driver.find_element(by=By.CSS_SELECTOR, value="#yui-rec0 a")
        all_button.click()
    except:
        create_empty_file(f'{sample_id} (no results).xlsx')
        return

    # select metadata columns
    meta1button = driver.find_element(by=By.CSS_SELECTOR, value="#projectMetadata > input:nth-child(1)")
    meta1button.click()

    # select metadata columns (2)
    meta1button = driver.find_element(by=By.CSS_SELECTOR, value="#jgi > input:nth-child(1)")
    meta1button.click()

    # redisplay
    redisplay_button = driver.find_element(by=By.ID, value="moreGo")
    redisplay_button.click()
    time.sleep(2)

    # select all rows for export
    selectall_button = driver.find_element(by=By.CSS_SELECTOR, value=".dt-buttons:nth-child(3) > .dt-button:nth-child(2) > span")
    selectall_button.click()
    
    # export to excel
    export_button =  driver.find_element(by=By.CSS_SELECTOR, value=".dt-buttons:nth-child(8) > .buttons-excel > span")
    export_button.click()

    time.sleep(2)

    if os.path.exists(f'{sample_id}.xlsx'):
        count = 1
        while os.path.exists(f'{sample_id} ({count}).xlsx'):
            count += 1
        os.rename('IMG.xlsx',f'{sample_id} ({count}).xlsx')
    else:
        os.rename('IMG.xlsx',f'{sample_id}.xlsx')

def main():
    # setup
    cwd = os.getcwd()
    options=webdriver.ChromeOptions()
    prefs={"download.default_directory": cwd}
    options.add_experimental_option("prefs",prefs)
    driver=webdriver.Chrome(options=options)
    driver.implicitly_wait(30)

    df = pd.read_csv("JGI_OMICS_Database_Nayfach_et_al2020.csv",encoding="")
    unique_sample_IDs = df.STUDY_ID.unique()

    for sample_id in unique_sample_IDs:
        if not os.path.exists(f"{sample_id}.xlsx"):
            print(f"Working on {sample_id}")
            download_sample_excel(sample_id, driver)

    driver.quit()

main()
