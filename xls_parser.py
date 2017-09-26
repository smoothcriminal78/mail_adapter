from openpyxl import load_workbook
from config import *
from io import BytesIO
import pandas as pd
import sys
import os

def parse(data):
    wb = load_workbook(filename=BytesIO(data))
    # wb = load_workbook('attachment.xlsx')
    sheet = wb.get_sheet_by_name('Лист1')
    df = pd.DataFrame(sheet.values)

    # product -> (qty, price)
    goods = {row[0] : (row[1], row[2]) for i, row in df[1:].iterrows()}

    f=open(os.path.join(os.getcwd(), 'goods'), 'a', encoding='UTF-8')
    print('\n'.join(['{} {} {}'.format(k,*v) for k,v in goods.items()]), file=f)
