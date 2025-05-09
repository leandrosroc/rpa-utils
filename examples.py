from tkinter import messagebox
import utils.pyautogui as upg
import utils.selenium as us
import utils.files as uf

def run_pyautogui_example():
    upg.write("Olá, mundo!", interval=0.1)
    messagebox.showinfo("PyAutoGUI", "Texto digitado!")

def run_selenium_example():
    driver = us.start_chrome(headless=True)
    us.go_to(driver, "https://www.python.org")
    title = driver.title
    us.close(driver)
    messagebox.showinfo("Selenium", f"Título da página: {title}")

def run_selenium_example_visible():
    driver = us.start_chrome(headless=False)
    us.go_to(driver, "https://www.python.org")
    title = driver.title
    us.close(driver)
    messagebox.showinfo("Selenium (visível)", f"Título da página: {title}")

def run_excel_example():
    data = [["Nome", "Idade"], ["Ana", 30], ["João", 25]]
    uf.write_xlsx("exemplo.xlsx", data)
    lidos = uf.read_xlsx("exemplo.xlsx")
    messagebox.showinfo("Excel", f"Conteúdo lido: {lidos}")

def run_csv_example():
    data = [["a", "b"], ["1", "2"]]
    uf.write_csv("exemplo.csv", data)
    lidos = uf.read_csv("exemplo.csv")
    messagebox.showinfo("CSV", f"Conteúdo lido: {lidos}")

def run_txt_example():
    uf.write_txt("exemplo.txt", "Exemplo de texto")
    lido = uf.read_txt("exemplo.txt")
    messagebox.showinfo("TXT", f"Conteúdo lido: {lido}")

def run_js_example():
    driver = us.start_chrome(headless=True)
    us.go_to(driver, "https://www.example.com")
    result = us.execute_js(driver, "return document.title;")
    us.close(driver)
    messagebox.showinfo("JavaScript", f"Título via JS: {result}")

def run_js_example_visible():
    driver = us.start_chrome(headless=False)
    us.go_to(driver, "https://www.example.com")
    result = us.execute_js(driver, "return document.title;")
    us.close(driver)
    messagebox.showinfo("JavaScript (visível)", f"Título via JS: {result}")
