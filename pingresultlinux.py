import os
import pandas as pd
import re
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

def pingfilinglinux(filename, commandline):
    filename_list = filename.split('_')
    date = filename_list[1]
    time = filename_list[2]

    fread = open(filename + '.txt', 'r', encoding='utf-8')
    fwrite = open(filename + '_' + 'temp.txt', 'w', encoding='utf-8')

    fwrite.write("seq,TTL,time(ms)\n")
    fread.seek(0)

    for i, v in enumerate(fread.readlines()):
        if (v.__contains__("Unreachable")):
            try:
                fwrite.write(re.findall(r'\d+', v)[-1] + ',unreachable' + ',\n')
            except:
                continue
        
        if (v.__contains__("exceeeded")):
            try:
                fwrite.write(re.findall(r'\d+', v)[-1] + ',timeout' + ',\n')
            except:
                continue

        if(v.__contains__("ttl")):
            try:
                fwrite.write(','.join(re.findall(r'[0-9.]+', v)[-3:]) + '\n')
            except:
                continue

        if (v.__contains__("statistics")):
            fwrite.write("all,receive,loss\n")

        if (v.__contains__("transmitted")):
            try:
                fwrite.write(','.join(re.findall(r'[0-9.]+', v)[:3]) + "\n")
            except:
                continue

        if (v.__contains__("rtt")):
            try:
                fwrite.write("min,avg,max\n")
                fwrite.write(','.join(re.findall(r'[0-9.]+', v)[:3]) + "\n")
            except:
                continue
            
    fwrite.close()
    fread.close()

    df = pd.read_csv(filename + '_' + 'temp.txt', encoding='utf-8', header=0)
    df.index += 1
    df.to_excel(filename + '.xlsx', index = True)

    lastindex = df.index[-1]

    wb = load_workbook(filename + '.xlsx', data_only=True)
    ws = wb['Sheet1']

    border_thick = Side(border_style='thin')

    for i in range(lastindex + 2, lastindex + 5):
        ws.merge_cells(start_row = i, start_column = 2, end_row = i, end_column = 6)
        ws.cell(row = i, column = 1).border = Border(left = border_thick, right = border_thick, top = border_thick, bottom = border_thick)
        ws.cell(row = i, column = 1).font = Font(bold = True)

    ws.cell(row = lastindex + 2, column = 1).value = "CMD"
    ws.cell(row = lastindex + 3, column = 1).value = "일시"
    ws.cell(row = lastindex + 4, column = 1).value = "비고"
    ws.cell(row = lastindex + 2, column = 2).value = commandline
    ws.cell(row = lastindex + 3, column = 2).value = date + ' ' + time

    for i in ['A', 'B', 'C', 'D']:
        for cell in range(len(ws[i])):
            ws[i + str(cell + 1)].alignment = Alignment(horizontal = 'center')
            if type(ws[i + str(cell + 1)].value) == str:
                try:
                    ws[i + str(cell + 1)].value = int(ws[i + str(cell + 1)].value)
                except:
                    try:
                        ws[i + str(cell + 1)].value = float(ws[i + str(cell + 1)].value)
                    except:
                        continue

    wb.save('results/' + filename + '.xlsx')
    os.makedirs('results/txt', exist_ok=True)
    shutil.move(filename + '.txt', 'results/txt')
    os.remove(filename + '_' + 'temp.txt')
    os.remove(filename + '.xlsx')


if __name__ == "__main__":
    pingfilinglinux("test", commandline="ping google.com -c 5")