from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def start_chrome(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--enable-unsafe-swiftshader')
    options.add_argument('--log-level=3')  # Suprime logs do ChromeDriver
    driver = webdriver.Chrome(options=options)
    return driver

def go_to(driver, url, timeout=30):
    driver.get(url)
    from selenium.webdriver.support.ui import WebDriverWait
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

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