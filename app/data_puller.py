"""
File:       data_puller.py
Purpose:    This file uses selenium to access Outcome Tracker and generate and pull reports
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Anno:       Anno Domini 2025
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time


def pull_data():
    service = Service(executable_path = "chromedriver.exe")
    driver = webdriver.Chrome(service = service)

    driver.get("https://www.vistashare.com/ot2/security/login/")

    input_element = driver.get_element(By.NAME, "")

    time.sleep(10)

    driver.quit()