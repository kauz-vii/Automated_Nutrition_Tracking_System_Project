import pyperclip
import time
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.styles import PatternFill
import os
import sys

FILE = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "Macros.xlsx")

green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

last_text = ""

def find_or_create_today(ws, today):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 1).value == today:
            return row
    new_row = ws.max_row + 1
    ws.cell(new_row, 1).value = today
    return new_row

def update_excel(cal, pro, carb, fat):
    wb = load_workbook(FILE)
    ws = wb["Sheet1"]

    today = datetime.now().strftime("%d-%m-%Y")
    row = find_or_create_today(ws, today)

    ws.cell(row, 2).value = cal
    ws.cell(row, 3).value = pro
    ws.cell(row, 4).value = carb
    ws.cell(row, 5).value = fat

    if cal < 2200:
        ws.cell(row, 6).fill = green
    else:
        ws.cell(row, 6).fill = red

    ws.cell(row, 7).fill = green
    ws.cell(row, 8).fill = green
    ws.cell(row, 9).fill = green

    wb.save(FILE)

while True:
    try:
        text = pyperclip.paste()
    except:
        time.sleep(0.3)
        continue

    if text != last_text and "FINAL_MACROS:" in text:
        try:
            data = text.split("FINAL_MACROS:")[1].strip()
            cal, pro, carb, fat = map(int, data.split(","))
            update_excel(cal, pro, carb, fat)
        except:
            pass

        last_text = text

    time.sleep(1)