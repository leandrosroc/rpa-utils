# RPA Utils

Ferramenta gráfica para criar automações RPA facilmente, integrando PyAutoGUI, Selenium (Chrome), manipulação de arquivos (Excel, CSV, TXT) e comandos de teclado.

## Instalação

1. **Clone ou baixe este repositório.**
2. **Instale as dependências:**

```bash
pip install pyautogui selenium openpyxl
```

> **Observação:**  
> - Para automação com Selenium, é necessário ter o Google Chrome instalado.
> - O Selenium recente já faz o download automático do ChromeDriver, não sendo necessário baixar manualmente.
> - No Linux, pode ser necessário instalar dependências extras do PyAutoGUI (ex: `python3-tk`, `scrot`, etc).

## Como executar

No terminal, navegue até a pasta do projeto e execute:

```bash
python main.py
```

## Como usar

- **Exemplos rápidos:**  
  Use os botões para testar automações prontas (digitar texto, abrir site, ler arquivos, etc).

- **Construa sua automação:**  
  Adicione passos como:
  - Clique em imagem (escolha uma imagem da tela para clicar no centro dela, pode definir quantos cliques)
  - Digitação de texto
  - Comando de teclado (ex: `ctrl+c`, `alt+tab`, `enter`)
  - Abrir site no Chrome (Selenium)
  - Executar JavaScript no navegador
  - Ler arquivo Excel

  Depois, clique em **Executar Automação** para rodar o fluxo criado.

## Observações

- Para automações que interagem com a tela, mantenha a área visível e não mova janelas durante a execução.
- Para comandos de teclado, use o formato:  
  - Simples: `enter`, `tab`, `esc`  
  - Combinado: `ctrl+c`, `alt+tab`, etc.

---

Desenvolvido para facilitar a criação de robôs RPA sem necessidade de programação.
