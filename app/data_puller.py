"""
File:       data_puller.py
Purpose:    This file uses selenium to access Outcome Tracker and generate and pull reports
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Anno:       Anno Domini 2025
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import BOT_DATA
import time
import pyotp

def navigate_to_reports(_driver):
    # Click on reports btn
    reports = _driver.find_element(By.LINK_TEXT, "Reports")
    reports.click()

def pull_data():
    service = Service(executable_path = "chromedriver.exe")
    driver = webdriver.Chrome(service = service)

    # Open up website
    driver.get("https://www.vistashare.com/ot2/security/login/")

    # send username & password and click login
    username_element = driver.find_element(By.NAME, "__ac_name")
    username_element.clear()
    username_element.send_keys(BOT_DATA.bot_username)

    password_element = driver.find_element(By.NAME, "__ac_password")
    password_element.clear()
    password_element.send_keys(BOT_DATA.bot_password)

    login_element = driver.find_element(By.NAME, "login")
    login_element.click()

    try:
        # find the captcha element
        captcha_element = driver.find_element(By.NAME, "captcha")
        # if this element is found then we need to add in a 20-second sleep for someone to manually complete it.
        time.sleep(20)
    except NoSuchElementException:
        print("No Capcha")

    # generate 2fa (2-factor authentication) otp (one time password)
    totp = pyotp.TOTP(BOT_DATA.otp_secret)

    # Wait one second for the page to load
    time.sleep(1)

    # send the otp
    otp_element = driver.find_element(By.NAME, "security_code")
    otp_element.clear()
    otp_element.send_keys(totp.now())

    verify_element = driver.find_element(By.NAME, "verify")
    verify_element.click()

    # Wait for page to load
    time.sleep(5)

    # click reports btn
    navigate_to_reports(driver)

    # search for desired report
    report_search = driver.find_element(By.NAME, "rpt_search")
    report_search.clear()
    report_search.send_keys("")

    #time.sleep(300)

    driver.quit()