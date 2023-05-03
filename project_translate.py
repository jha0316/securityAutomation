import openpyxl
import time
from googletrans import Translator
from datetime import datetime

now=datetime.now().strftime("%Y-%m-%d")
workbook=openpyxl.load_workbook('boannews_2023-04-03.xlsx')
worksheet=workbook.active

translator=Translator()

for row in worksheet.iter_rows():
    for cell in row:
        translated_text=translator.translate(cell.value,dest='en').text
        cell.value=translated_text
    
    time.sleep(1)

workbook.save(f'translated_{now}.xlsx')


