import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import os
import time
import utils.pyautogui as upg
import utils.selenium as us
import utils.files as uf

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
        tk.Button(self.button_frame, text="Adicionar Loop", command=self.add_loop).pack(fill='x')
        tk.Button(self.button_frame, text="Não fechar Selenium", command=self.add_no_close_selenium).pack(fill='x')
        tk.Button(self.button_frame, text="Fechar Selenium", command=self.add_close_selenium).pack(fill='x')
        tk.Button(self.button_frame, text="Salvar Fluxo", command=self.save_flow).pack(fill='x', pady=(10,0))
        tk.Button(self.button_frame, text="Restaurar Fluxo", command=self.load_flow).pack(fill='x')
        self.run_btn = tk.Button(self.button_frame, text="Iniciar Automação", command=self.run_steps)
        self.run_btn.pack(fill='x', pady=(10,0))
        self.stop_btn = tk.Button(self.button_frame, text="Parar", command=self.stop_runner, state='disabled')
        self.stop_btn.pack(fill='x')
        tk.Button(self.button_frame, text="Limpar Passos", command=self.clear_steps).pack(fill='x')
        
        # Frame para botões de manipulação de passos
        self.step_ctrl_frame = tk.Frame(self.frame)
        self.step_ctrl_frame.pack(side='left', padx=5, pady=5, anchor='n')
        tk.Button(self.step_ctrl_frame, text="↑ Subir", command=self.move_step_up, width=8).pack(pady=2)
        tk.Button(self.step_ctrl_frame, text="↓ Descer", command=self.move_step_down, width=8).pack(pady=2)
        tk.Button(self.step_ctrl_frame, text="🗑 Apagar", command=self.delete_step, width=8, fg="red").pack(pady=2)

        self._runner_thread = None
        self._stop_flag = threading.Event()
        self.driver = None
        self.driver_visible = None
        self.close_driver = True
        self.close_driver_visible = True

        # Caixa de log
        self.log_frame = tk.Frame(self.frame)
        self.log_frame.pack(side='left', padx=5, fill='both', expand=True)
        tk.Label(self.log_frame, text="Log de Execução:").pack(anchor='w')
        self.log_text = tk.Text(self.log_frame, width=40, height=30, state='disabled', bg="#f8f8f8")
        self.log_text.pack(fill='both', expand=True)
        self.log_scroll = tk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.log_scroll.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=self.log_scroll.set)

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')

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

    def add_loop(self):
        top = tk.Toplevel(self.master)
        top.title("Adicionar Loop")
        tk.Label(top, text="Índice inicial do passo (0-based, deixe vazio para 0):").pack()
        start_entry = tk.Entry(top)
        start_entry.pack()
        tk.Label(top, text="Índice final do passo (inclusivo, deixe vazio para último):").pack()
        end_entry = tk.Entry(top)
        end_entry.pack()
        tk.Label(top, text="Repetições (deixe vazio para infinito):").pack()
        repeat_entry = tk.Entry(top)
        repeat_entry.pack()
        def ok():
            # Ajuste: permite campos vazios para pegar todo o fluxo, mesmo se steps estiver vazio
            try:
                start_str = start_entry.get().strip()
                end_str = end_entry.get().strip()
                if len(self.steps) == 0:
                    messagebox.showerror("Erro", "Não há passos para criar loop.")
                    return
                start = int(start_str) if start_str != "" else 0
                end = int(end_str) if end_str != "" else len(self.steps)-1
                # Permite loop de todo o fluxo se steps não está vazio
                if start < 0 or end < 0 or start > end or end >= len(self.steps):
                    raise Exception()
            except Exception:
                messagebox.showerror("Erro", "Índices inválidos.")
                return
            reps = repeat_entry.get()
            if reps.strip() == "":
                reps = None
            else:
                try:
                    reps = int(reps)
                except Exception:
                    messagebox.showerror("Erro", "Repetições inválidas.")
                    return
            loop_steps = self.steps[start:end+1]
            for _ in range(end, start-1, -1):
                del self.steps[_]
                self.listbox.delete(_)
            loop_block = {'type': 'loop', 'steps': loop_steps, 'repeat': reps}
            self.steps.insert(start, loop_block)
            desc = f"LOOP {'∞' if reps is None else reps}x: {len(loop_steps)} passos"
            self.listbox.insert(start, desc)
            top.destroy()
        tk.Button(top, text="OK", command=ok).pack()

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

    def _normalize_steps(self, steps):
        normalized = []
        for step in steps:
            if isinstance(step, list):
                # Converte listas simples para tuplas
                if len(step) > 0 and isinstance(step[0], str):
                    normalized.append(tuple(step))
                else:
                    normalized.append(step)
            elif isinstance(step, dict) and step.get('type') == 'loop':
                # Recursivo para steps dentro do loop
                step = dict(step)
                step['steps'] = self._normalize_steps(step['steps'])
                normalized.append(step)
            else:
                normalized.append(step)
        return normalized

    def load_flow(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path and os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.steps = self._normalize_steps(loaded)
            self.listbox.delete(0, 'end')
            for step in self.steps:
                if isinstance(step, dict) and step.get('type') == 'read_excel_advanced':
                    self.listbox.insert('end', f"Ler Excel Avançado: {step['files']} Linhas:{step['rows']} Cols:{step['cols']} Ação:{step['action']} Param:{step['param']}")
                elif isinstance(step, dict) and step.get('type') == 'loop':
                    desc = f"LOOP {'∞' if step['repeat'] is None else step['repeat']}x: {len(step['steps'])} passos"
                    self.listbox.insert('end', desc)
                else:
                    self.listbox.insert('end', str(step))

    def clear_steps(self):
        self.steps.clear()
        self.listbox.delete(0, 'end')
    
    def run_steps(self):
        if self._runner_thread and self._runner_thread.is_alive():
            messagebox.showinfo("Automação", "Já está em execução.")
            return
        self._stop_flag.clear()
        self.run_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.driver = None
        self.driver_visible = None
        self.close_driver = True
        self.close_driver_visible = True
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, 'end')
        self.log_text.config(state='disabled')
        self.log("Iniciando automação...")
        def runner():
            try:
                self._execute_steps(self.steps)
                self.log("Fluxo finalizado!")
                messagebox.showinfo("Automação", "Fluxo finalizado!")
            except StopIteration:
                self.log("Execução interrompida pelo usuário.")
                messagebox.showinfo("Automação", "Execução interrompida pelo usuário.")
            finally:
                self.run_btn.config(state='normal')
                self.stop_btn.config(state='disabled')
        self._runner_thread = threading.Thread(target=runner)
        self._runner_thread.start()

    def stop_runner(self):
        self._stop_flag.set()

    def _execute_steps(self, steps):
        for idx, step in enumerate(steps):
            if self._stop_flag.is_set():
                self.log("Execução interrompida pelo usuário.")
                raise StopIteration()
            # Log do passo atual
            self.log(f"Executando passo {idx+1}/{len(steps)}: {self._describe_step(step)}")
            if isinstance(step, dict) and step.get('type') == 'loop':
                reps = step.get('repeat')
                count = 0
                self.log(f"Iniciando loop ({'∞' if reps is None else reps}x) com {len(step['steps'])} passos")
                while reps is None or count < reps:
                    if self._stop_flag.is_set():
                        self.log("Execução interrompida pelo usuário dentro do loop.")
                        raise StopIteration()
                    self.log(f"Repetição do loop: {count+1}")
                    self._execute_steps(step['steps'])
                    count += 1
                self.log("Fim do loop")
            elif isinstance(step, dict) and step.get('type') == 'read_excel_advanced':
                self.log("Executando leitura avançada de Excel (não exibido no log detalhado).")
                for file in step['files']:
                    data = uf.read_xlsx(file)
                    def parse_range(val):
                        if '-' in val:
                            a, b = val.split('-')
                            try:
                                return list(range(int(a), int(b)+1))
                            except:
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
                    if not rows: rows = range(1, len(data))
                    if not cols: cols = range(len(data[0]))
                    for r in rows:
                        for c in cols:
                            if self._stop_flag.is_set():
                                raise StopIteration()
                            try:
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
            elif isinstance(step, tuple) and step[0] == 'click':
                self.log(f"Clique: {step[1]} ({step[2]}x)")
                upg.click(step[1], step[2])
            elif isinstance(step, tuple) and step[0] == 'click_image':
                self.log(f"Procurando imagem na tela: {step[1]}")
                pos = upg.locate_on_screen(step[1])
                if pos:
                    self.log(f"Imagem encontrada em {pos.x}, {pos.y}. Clicando {step[2]}x.")
                    upg.click(pos.x, pos.y, clicks=step[2])
                else:
                    self.log(f"Imagem não encontrada: {step[1]}")
                    messagebox.showerror("Erro", f"Imagem não encontrada na tela: {step[1]}")
                    break
            elif isinstance(step, tuple) and step[0] == 'write':
                self.log(f"Digitando: {step[1]}")
                upg.write(step[1])
            elif isinstance(step, tuple) and step[0] == 'command':
                self.log(f"Comando: {step[1]}")
                keys = [k.strip() for k in step[1].split('+')]
                if len(keys) == 1:
                    upg.press(keys[0])
                else:
                    upg.hotkey(*keys)
            elif isinstance(step, tuple) and step[0] == 'open_site':
                self.log(f"Abrindo site (headless): {step[1]}")
                if not self.driver:
                    self.driver = us.start_chrome(headless=True)
                us.go_to(self.driver, step[1])
            elif isinstance(step, tuple) and step[0] == 'open_site_visible':
                self.log(f"Abrindo site (visível): {step[1]}")
                if not self.driver_visible:
                    self.driver_visible = us.start_chrome(headless=False)
                us.go_to(self.driver_visible, step[1])
            elif isinstance(step, tuple) and step[0] == 'js':
                self.log(f"Executando JS (headless): {step[1]}")
                if self.driver:
                    us.execute_js(self.driver, step[1])
            elif isinstance(step, tuple) and step[0] == 'js_visible':
                self.log(f"Executando JS (visível): {step[1]}")
                if self.driver_visible:
                    us.execute_js(self.driver_visible, step[1])
            elif isinstance(step, tuple) and step[0] == 'read_excel':
                self.log(f"Lendo Excel: {step[1]}")
                data = uf.read_xlsx(step[1])
                messagebox.showinfo("Excel", f"Conteúdo: {data}")
            elif isinstance(step, tuple) and step[0] == 'timesleep':
                self.log(f"Aguardando {step[1]} segundos...")
                time.sleep(step[1])
            elif isinstance(step, tuple) and step[0] == 'no_close_selenium':
                self.log("Não fechar Selenium ao final do fluxo.")
                self.close_driver = False
                self.close_driver_visible = False
            elif isinstance(step, tuple) and step[0] == 'close_selenium':
                self.log("Fechando Selenium.")
                if self.driver:
                    us.close(self.driver)
                    self.driver = None
                if self.driver_visible:
                    us.close(self.driver_visible)
                    self.driver_visible = None
                self.close_driver = True
                self.close_driver_visible = True
        if self.close_driver and self.driver:
            self.log("Fechando Selenium (headless) ao final do fluxo.")
            us.close(self.driver)
            self.driver = None
        if self.close_driver_visible and self.driver_visible:
            self.log("Fechando Selenium (visível) ao final do fluxo.")
            us.close(self.driver_visible)
            self.driver_visible = None

    def _describe_step(self, step):
        # Retorna uma string amigável para o log
        if isinstance(step, tuple):
            return str(step)
        elif isinstance(step, dict) and step.get('type') == 'loop':
            return f"LOOP ({'∞' if step.get('repeat') is None else step.get('repeat')}x) [{len(step.get('steps', []))} passos]"
        elif isinstance(step, dict) and step.get('type') == 'read_excel_advanced':
            return f"Ler Excel Avançado: {step.get('files')}"
        else:
            return str(step)

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
        if self.listbox.size() > 0:
            next_idx = min(idx, self.listbox.size()-1)
            self.listbox.selection_set(next_idx)
