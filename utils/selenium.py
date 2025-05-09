from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def start_chrome(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    return driver

def go_to(driver, url):
    driver.get(url)

def find_element(driver, by, value):
    return driver.find_element(by, value)

def click_element(element):
    element.click()

def write_element(element, text):
    element.clear()
    element.send_keys(text)

def execute_js(driver, script, *args):
    return driver.execute_script(script, *args)

def close(driver):
    driver.quit()