import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import os
import time

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
        tk.Button(self.button_frame, text="Unir XLSX", command=self.merge_xlsx).pack(fill='x')
        tk.Button(self.button_frame, text="Unir CSV", command=self.merge_csv).pack(fill='x')
        tk.Button(self.button_frame, text="Unir TXT", command=self.merge_txt).pack(fill='x')
        tk.Button(self.button_frame, text="Adicionar Time Sleep", command=self.add_timesleep).pack(fill='x')
        tk.Button(self.button_frame, text="Não fechar Selenium", command=self.add_no_close_selenium).pack(fill='x')
        tk.Button(self.button_frame, text="Fechar Selenium", command=self.add_close_selenium).pack(fill='x')
        tk.Button(self.button_frame, text="Salvar Fluxo", command=self.save_flow).pack(fill='x', pady=(10,0))
        tk.Button(self.button_frame, text="Restaurar Fluxo", command=self.load_flow).pack(fill='x')
        tk.Button(self.button_frame, text="Executar Automação", command=self.run_steps).pack(fill='x', pady=(10,0))
        tk.Button(self.button_frame, text="Limpar Passos", command=self.clear_steps).pack(fill='x')
        
        # Frame para botões de manipulação de passos
        self.step_ctrl_frame = tk.Frame(self.frame)
        self.step_ctrl_frame.pack(side='left', padx=5, pady=5, anchor='n')
        tk.Button(self.step_ctrl_frame, text="↑ Subir", command=self.move_step_up, width=8).pack(pady=2)
        tk.Button(self.step_ctrl_frame, text="↓ Descer", command=self.move_step_down, width=8).pack(pady=2)
        tk.Button(self.step_ctrl_frame, text="🗑 Apagar", command=self.delete_step, width=8, fg="red").pack(pady=2)

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
        top = tk.Toplevel(self.master)
        top.title("Ler Excel Avançado")
        files = []

        def add_file():
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if path:
                files.append(path)
                files_list.insert('end', path)

        tk.Label(top, text="Arquivos Excel:").pack()
        files_list = tk.Listbox(top, width=50)
        files_list.pack()
        tk.Button(top, text="Adicionar Arquivo", command=add_file).pack()

        tk.Label(top, text="Linha(s) (ex: 2, 2-5):").pack()
        row_entry = tk.Entry(top)
        row_entry.pack()
        tk.Label(top, text="Coluna(s) (ex: B, B-D, 2, 2-4):").pack()
        col_entry = tk.Entry(top)
        col_entry.pack()
        tk.Label(top, text="Ação (ex: digitar, clicar, comando):").pack()
        action_entry = tk.Entry(top)
        action_entry.pack()
        tk.Label(top, text="Parâmetro da ação (ex: texto, comando):").pack()
        param_entry = tk.Entry(top)
        param_entry.pack()

        def ok():
            if files:
                step = {
                    'type': 'read_excel_advanced',
                    'files': files.copy(),
                    'rows': row_entry.get(),
                    'cols': col_entry.get(),
                    'action': action_entry.get(),
                    'param': param_entry.get()
                }
                self.steps.append(step)
                self.listbox.insert('end', f"Ler Excel Avançado: {files} Linhas:{step['rows']} Cols:{step['cols']} Ação:{step['action']} Param:{step['param']}")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()

    def add_timesleep(self):
        top = tk.Toplevel(self.master)
        top.title("Adicionar Time Sleep")
        tk.Label(top, text="Tempo em segundos:").pack()
        entry = tk.Entry(top, width=20)
        entry.insert(0, "1")
        entry.pack()
        def ok():
            try:
                secs = float(entry.get())
            except Exception:
                secs = 1
            self.steps.append(('timesleep', secs))
            self.listbox.insert('end', f"Time Sleep: {secs}s")
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()

    def add_no_close_selenium(self):
        self.steps.append(('no_close_selenium',))
        self.listbox.insert('end', "Não fechar Selenium")

    def add_close_selenium(self):
        self.steps.append(('close_selenium',))
        self.listbox.insert('end', "Fechar Selenium")

    def merge_xlsx(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if files:
            data = []
            for f in files:
                data.extend(uf.read_xlsx(f))
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                uf.write_xlsx(save_path, data)
                messagebox.showinfo("Unir XLSX", f"Arquivos unidos em: {save_path}")

    def merge_csv(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        if files:
            data = []
            for f in files:
                data.extend(uf.read_csv(f))
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if save_path:
                uf.write_csv(save_path, data)
                messagebox.showinfo("Unir CSV", f"Arquivos unidos em: {save_path}")

    def merge_txt(self):
        files = filedialog.askopenfilenames(filetypes=[("Text files", "*.txt")])
        if files:
            data = ""
            for f in files:
                data += uf.read_txt(f) + "\n"
            save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
            if save_path:
                uf.write_txt(save_path, data)
                messagebox.showinfo("Unir TXT", f"Arquivos unidos em: {save_path}")

    def save_flow(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.steps, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Salvar Fluxo", f"Fluxo salvo em: {path}")

    def load_flow(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path and os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                self.steps = json.load(f)
            self.listbox.delete(0, 'end')
            for step in self.steps:
                if isinstance(step, dict) and step.get('type') == 'read_excel_advanced':
                    self.listbox.insert('end', f"Ler Excel Avançado: {step['files']} Linhas:{step['rows']} Cols:{step['cols']} Ação:{step['action']} Param:{step['param']}")
                else:
                    self.listbox.insert('end', str(step))

    def clear_steps(self):
        self.steps.clear()
        self.listbox.delete(0, 'end')
    
    def run_steps(self):
        def runner():
            driver = None
            driver_visible = None
            close_driver = True
            close_driver_visible = True
            for step in self.steps:
                if isinstance(step, dict) and step.get('type') == 'read_excel_advanced':
                    # Simplificação: só lê e executa ação nas células especificadas
                    for file in step['files']:
                        data = uf.read_xlsx(file)
                        # Interpretação de linhas/colunas
                        def parse_range(val):
                            if '-' in val:
                                a, b = val.split('-')
                                try:
                                    return list(range(int(a), int(b)+1))
                                except:
                                    # Letras
                                    return [chr(c) for c in range(ord(a.upper()), ord(b.upper())+1)]
                            elif ',' in val:
                                return [v.strip() for v in val.split(',')]
                            elif val:
                                try:
                                    return [int(val)]
                                except:
                                    return [val]
                            return []
                        rows = parse_range(step.get('rows',''))
                        cols = parse_range(step.get('cols',''))
                        # Se não especificado, pega tudo
                        if not rows: rows = range(1, len(data))
                        if not cols: cols = range(len(data[0]))
                        for r in rows:
                            for c in cols:
                                try:
                                    # Suporte a letras de coluna
                                    if isinstance(c, str) and c.isalpha():
                                        cidx = ord(c.upper()) - ord('A')
                                    else:
                                        cidx = int(c)
                                    val = data[int(r)][cidx]
                                except Exception:
                                    val = ""
                                if step['action'] == 'digitar':
                                    upg.write(str(val))
                                elif step['action'] == 'clicar':
                                    upg.click(str(val))
                                elif step['action'] == 'comando':
                                    upg.press(str(val))
                elif step[0] == 'click':
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
                elif step[0] == 'timesleep':
                    time.sleep(step[1])
                elif step[0] == 'no_close_selenium':
                    close_driver = False
                    close_driver_visible = False
                elif step[0] == 'close_selenium':
                    if driver:
                        us.close(driver)
                        driver = None
                    if driver_visible:
                        us.close(driver_visible)
                        driver_visible = None
                    close_driver = True
                    close_driver_visible = True
            # Fechar drivers apenas se permitido
            if close_driver and driver:
                us.close(driver)
            if close_driver_visible and driver_visible:
                us.close(driver_visible)
            messagebox.showinfo("Automação", "Fluxo finalizado!")
        threading.Thread(target=runner).start()

    def move_step_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self.steps[idx-1], self.steps[idx] = self.steps[idx], self.steps[idx-1]
        txt = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx-1, txt)
        self.listbox.selection_clear(0, 'end')
        self.listbox.selection_set(idx-1)

    def move_step_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == len(self.steps)-1:
            return
        idx = sel[0]
        self.steps[idx+1], self.steps[idx] = self.steps[idx], self.steps[idx+1]
        txt = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx+1, txt)
        self.listbox.selection_clear(0, 'end')
        self.listbox.selection_set(idx+1)

    def delete_step(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.steps[idx]
        self.listbox.delete(idx)
        # Seleciona o próximo passo, se existir
        if self.listbox.size() > 0:
            next_idx = min(idx, self.listbox.size()-1)
            self.listbox.selection_set(next_idx)

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