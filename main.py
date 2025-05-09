import tkinter as tk
import threading

from automation_builder import AutomationBuilder
from examples import *

def run_threaded(func):
	threading.Thread(target=func).start()

root = tk.Tk()
root.title("RPA Utils")

tk.Label(root, text="Exemplos rápidos:").pack()
tk.Button(root, text="PyAutoGUI: Digitar texto", command=lambda: run_threaded(run_pyautogui_example)).pack(fill='x')
tk.Button(root, text="Selenium: Abrir página", command=lambda: run_threaded(run_selenium_example)).pack(fill='x')
tk.Button(root, text="Selenium: Abrir página (visível)", command=lambda: run_threaded(run_selenium_example_visible)).pack(fill='x')
tk.Button(root, text="Excel: Escrever/Ler", command=lambda: run_threaded(run_excel_example)).pack(fill='x')
tk.Button(root, text="CSV: Escrever/Ler", command=lambda: run_threaded(run_csv_example)).pack(fill='x')
tk.Button(root, text="TXT: Escrever/Ler", command=lambda: run_threaded(run_txt_example)).pack(fill='x')
tk.Button(root, text="Selenium: Executar JS", command=lambda: run_threaded(run_js_example)).pack(fill='x')
tk.Button(root, text="Selenium: Executar JS (visível)", command=lambda: run_threaded(run_js_example_visible)).pack(fill='x')

tk.Label(root, text="").pack()
tk.Label(root, text="Construa sua automação:").pack()
AutomationBuilder(root)

root.mainloop()