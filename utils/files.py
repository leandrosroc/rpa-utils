import csv
import openpyxl

def read_csv(path):
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.reader(f))

def write_csv(path, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(rows)

def read_txt(path):
    with open(path, encoding='utf-8') as f:
        return f.read()

def write_txt(path, text):
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)

def read_xlsx(path, sheet=0):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[sheet]
    return [[cell.value for cell in row] for row in ws.iter_rows()]

def write_xlsx(path, data, sheet_name='Sheet1'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in data:
        ws.append(row)
    wb.save(path)