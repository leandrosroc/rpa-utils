import tkinter as tk
from tkinter import filedialog, messagebox
import threading

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

def run_threaded(func):
    threading.Thread(target=func).start()

class AutomationBuilder:
    def __init__(self, master):
        self.master = master
        self.steps = []
        self.frame = tk.Frame(master)
        self.frame.pack(padx=10, pady=10, fill='both', expand=True)
        self.listbox = tk.Listbox(self.frame, width=60)
        self.listbox.pack(side='left', fill='both', expand=True)
        self.scroll = tk.Scrollbar(self.frame, command=self.listbox.yview)
        self.scroll.pack(side='left', fill='y')
        self.listbox.config(yscrollcommand=self.scroll.set)
        self.button_frame = tk.Frame(self.frame)
        self.button_frame.pack(side='left', padx=10)
        tk.Button(self.button_frame, text="Adicionar Clique", command=self.add_click).pack(fill='x')
        tk.Button(self.button_frame, text="Adicionar Digitação", command=self.add_write).pack(fill='x')
        tk.Button(self.button_frame, text="Adicionar Comando", command=self.add_command).pack(fill='x')
        tk.Button(self.button_frame, text="Abrir Site (Selenium)", command=self.add_open_site).pack(fill='x')
        tk.Button(self.button_frame, text="Abrir Site (Selenium visível)", command=self.add_open_site_visible).pack(fill='x')
        tk.Button(self.button_frame, text="Executar JS (Selenium)", command=self.add_js).pack(fill='x')
        tk.Button(self.button_frame, text="Executar JS (Selenium visível)", command=self.add_js_visible).pack(fill='x')
        tk.Button(self.button_frame, text="Ler Arquivo Excel", command=self.add_read_excel).pack(fill='x')
        tk.Button(self.button_frame, text="Executar Automação", command=self.run_steps).pack(fill='x', pady=(10,0))
        tk.Button(self.button_frame, text="Limpar Passos", command=self.clear_steps).pack(fill='x')
    
    def add_click(self):
        top = tk.Toplevel(self.master)
        top.title("Adicionar Clique")
        tk.Label(top, text="Imagem do local a ser clicado:").grid(row=0, column=0, columnspan=2)
        img_path_var = tk.StringVar()
        def select_img():
            path = filedialog.askopenfilename(
                title="Selecione a imagem do local a ser clicado",
                filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.bmp")]
            )
            img_path_var.set(path)
        tk.Button(top, text="Selecionar Imagem", command=select_img).grid(row=1, column=0, columnspan=2)
        tk.Label(top, text="Qtd. de cliques:").grid(row=2, column=0)
        clicks_entry = tk.Entry(top)
        clicks_entry.insert(0, "1")
        clicks_entry.grid(row=2, column=1)
        def ok():
            path = img_path_var.get()
            try:
                clicks = int(clicks_entry.get())
            except Exception:
                clicks = 1
            if path:
                self.steps.append(('click_image', path, clicks))
                self.listbox.insert('end', f"Clique na imagem: {path} ({clicks}x)")
            top.destroy()
        tk.Button(top, text="OK", command=ok).grid(row=3, column=0, columnspan=2)

    def add_write(self):
        top = tk.Toplevel(self.master)
        top.title("Adicionar Digitação")
        tk.Label(top, text="Texto:").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            text = entry.get()
            self.steps.append(('write', text))
            self.listbox.insert('end', f"Digitar: {text}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()
    
    def add_command(self):
        top = tk.Toplevel(self.master)
        top.title("Adicionar Comando")
        tk.Label(top, text="Comando (ex: ctrl+c, alt+tab, enter):").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            cmd = entry.get().strip()
            if cmd:
                self.steps.append(('command', cmd))
                self.listbox.insert('end', f"Comando: {cmd}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()

    def add_open_site(self):
        top = tk.Toplevel(self.master)
        top.title("Abrir Site")
        tk.Label(top, text="URL:").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            url = entry.get()
            self.steps.append(('open_site', url))
            self.listbox.insert('end', f"Abrir site: {url}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()
    
    def add_open_site_visible(self):
        top = tk.Toplevel(self.master)
        top.title("Abrir Site (visível)")
        tk.Label(top, text="URL:").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            url = entry.get()
            self.steps.append(('open_site_visible', url))
            self.listbox.insert('end', f"Abrir site (visível): {url}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()

    def add_js(self):
        top = tk.Toplevel(self.master)
        top.title("Executar JS")
        tk.Label(top, text="JavaScript:").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            js = entry.get()
            self.steps.append(('js', js))
            self.listbox.insert('end', f"Executar JS: {js}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()
    
    def add_js_visible(self):
        top = tk.Toplevel(self.master)
        top.title("Executar JS (visível)")
        tk.Label(top, text="JavaScript:").pack()
        entry = tk.Entry(top, width=40)
        entry.pack()
        def ok():
            js = entry.get()
            self.steps.append(('js_visible', js))
            self.listbox.insert('end', f"Executar JS (visível): {js}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()
    
    def add_read_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.steps.append(('read_excel', path))
            self.listbox.insert('end', f"Ler Excel: {path}")
    
    def clear_steps(self):
        self.steps.clear()
        self.listbox.delete(0, 'end')
    
    def run_steps(self):
        def runner():
            driver = None
            driver_visible = None
            for step in self.steps:
                if step[0] == 'click':
                    upg.click(step[1], step[2])
                elif step[0] == 'click_image':
                    pos = upg.locate_on_screen(step[1])
                    if pos:
                        upg.click(pos.x, pos.y, clicks=step[2])
                    else:
                        messagebox.showerror("Erro", f"Imagem não encontrada na tela: {step[1]}")
                        break
                elif step[0] == 'write':
                    upg.write(step[1])
                elif step[0] == 'command':
                    keys = [k.strip() for k in step[1].split('+')]
                    if len(keys) == 1:
                        upg.press(keys[0])
                    else:
                        upg.hotkey(*keys)
                elif step[0] == 'open_site':
                    if not driver:
                        driver = us.start_chrome(headless=True)
                    us.go_to(driver, step[1])
                elif step[0] == 'open_site_visible':
                    if not driver_visible:
                        driver_visible = us.start_chrome(headless=False)
                    us.go_to(driver_visible, step[1])
                elif step[0] == 'js':
                    if driver:
                        us.execute_js(driver, step[1])
                elif step[0] == 'js_visible':
                    if driver_visible:
                        us.execute_js(driver_visible, step[1])
                elif step[0] == 'read_excel':
                    data = uf.read_xlsx(step[1])
                    messagebox.showinfo("Excel", f"Conteúdo: {data}")
            if driver:
                us.close(driver)
            if driver_visible:
                us.close(driver_visible)
            messagebox.showinfo("Automação", "Fluxo finalizado!")
        threading.Thread(target=runner).start()

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