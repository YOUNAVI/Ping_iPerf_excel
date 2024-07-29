import os
import pandas as pd
import re
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

def pingfiling(filename, commandline):
    filename_list = filename.split('_')
    date = filename_list[1]
    time = filename_list[2]

    fread = open(filename + '.txt', 'r', encoding='utf-8')
    fwrite = open(filename + '_' + 'temp.txt', 'w', encoding='utf-8')

    fwrite.write("byte,time(ms),TTL\n")
    fread.seek(0)
    lines = len(fread.readlines())
    fread.seek(0)

    for i, v in enumerate(fread.readlines()):
        if (v.__contains__("만료")):
            fwrite.write(',' + v.strip('\n') + ',\n')

        if len(re.findall(r'\d+', v)) == 7:
            fwrite.write(','.join(re.findall(r'\d+', v)[4:]) + "\n")

        if (v.__contains__("통계")):
            fwrite.write("all,receive,loss\n")

        if (v.__contains__("보냄")):
            fwrite.write(','.join(re.findall(r'\d+', v)[:3]) + "\n")

        if (v.__contains__("왕복")):
            fwrite.write("min,max,avg\n")

        if (v.__contains__("최소")):
            fwrite.write(','.join(re.findall(r'\d+', v)[:3]) + "\n")
            
    fwrite.close()
    fread.close()

    df = pd.read_csv(filename + '_' + 'temp.txt', encoding='utf-8', header=0)
    df.index += 1
    df.to_excel(filename + '.xlsx', index = True)

    wb = load_workbook(filename + '.xlsx', data_only=True)
    ws = wb['Sheet1']

    border_thick = Side(border_style='thin')

    for i in range(lines - 1, lines + 2):
        ws.merge_cells(start_row = i, start_column = 2, end_row = i, end_column = 6)
        ws.cell(row = i, column = 1).border = Border(left = border_thick, right = border_thick, top = border_thick, bottom = border_thick)
        ws.cell(row = i, column = 1).font = Font(bold = True)

    ws.cell(row = lines - 1, column = 1).value = "CMD"
    ws.cell(row = lines, column = 1).value = "일시"
    ws.cell(row = lines + 1, column = 1).value = "비고"
    ws.cell(row = lines - 1, column = 2).value = commandline
    ws.cell(row = lines, column = 2).value = date + ' ' + time

    for i in ['A', 'B', 'C', 'D']:
        for cell in range(len(ws[i])):
            ws[i + str(cell + 1)].alignment = Alignment(horizontal = 'center')
            if type(ws[i + str(cell + 1)].value) == str:
                try:
                    ws[i + str(cell + 1)].value = int(ws[i + str(cell + 1)].value)
                except:
                    continue

    wb.save('results/' + filename + '.xlsx')
    os.makedirs('results/txt', exist_ok=True)
    shutil.move(filename + '.txt', 'results/txt')
    os.remove(filename + '_' + 'temp.txt')
    os.remove(filename + '.xlsx')


if __name__ == "__main__":
    pingfiling("test", commandline="ping google.com -n 5")