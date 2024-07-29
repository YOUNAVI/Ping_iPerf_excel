import pandas as pd
import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

def iperfiling(filename, commandline) :
    filename_list = filename.split('_')
    date = filename_list[1]
    time = filename_list[2]

    fread = open(filename + '.txt', 'r', encoding='utf-8')
    fwrite = open(filename + '_' + 'temp.txt', 'w', encoding='utf-8')

    fwrite.write("interval,sec,transfer,bytes,bitrate,bits/sec,remarks\n")
    fread.seek(0)
    lines = len(fread.readlines())
    fread.seek(0)

    for i, v in enumerate(fread.readlines()):
        v = arrange(v)
        if len(v) == 7:
            fwrite.write(','.join(v[1:]) + '\n')

        if i == lines - 4:
            fwrite.write('-,-,-,-,-,-,-\n')

        if v.__contains__('sender') or v.__contains__('receiver'):
            fwrite.write(','.join(v[1:]) + '\n')
            
    fwrite.close()
    fread.close()

    df = pd.read_csv(filename+ '_temp.txt', sep=',', encoding='utf-8', header=0)
    df.index += 1
    df.to_excel(filename + '.xlsx', index = True)

    wb = load_workbook(filename + '.xlsx', data_only=True)
    ws = wb['Sheet1']

    border_thick = Side(border_style='thin')

    for i in range(lines - 4, lines - 1):
        ws.merge_cells(start_row = i, start_column = 2, end_row = i, end_column = 8)
        ws.cell(row = i, column = 1).border = Border(left = border_thick, right = border_thick, top = border_thick, bottom = border_thick)
        ws.cell(row = i, column = 1).font = Font(bold = True)

    ws.cell(row = lines - 4, column = 1).value = "CMD"
    ws.cell(row = lines - 3, column = 1).value = "일시"
    ws.cell(row = lines - 2, column = 1).value = "비고"
    ws.cell(row = lines - 4, column = 2).value = commandline
    ws.cell(row = lines - 3, column = 2).value = date + ' ' + time

    for i in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
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

def iperfiling_udp(filename, commandline) :
    filename_list = filename.split('_')
    date = filename_list[1]
    time = filename_list[2]

    fread = open(filename + '.txt', 'r', encoding='utf-8')
    fwrite = open(filename + '_' + 'temp.txt', 'w', encoding='utf-8')

    fwrite.write("interval,sec,transfer,bytes,bitrate,bits/sec,total,remarks,empty1,empty2\n")
    fread.seek(0)
    lines = len(fread.readlines())
    fread.seek(0)

    for i, v in enumerate(fread.readlines()):
        v = arrange(v)
        if len(v) == 8:
            fwrite.write(','.join(v[1:]) + '\n')

        if i == lines - 4:
            fwrite.write('interval,sec,transfer,bytes,bitrate,bits/sec,Jitter(ms),lost/total,loss,remarks\n')

        if v.__contains__('sender') or v.__contains__('receiver'):
            v[10] = v[10].strip('()')
            v.remove('ms')
            fwrite.write(','.join(v[1:]) + '\n')
            
    fwrite.close()
    fread.close()

    df = pd.read_csv(filename+ '_temp.txt', sep=',', encoding='utf-8', header=0)
    df.index += 1
    df.to_excel(filename + '.xlsx', index = True)

    wb = load_workbook(filename + '.xlsx', data_only=True)
    ws = wb['Sheet1']

    border_thick = Side(border_style='thin')

    for i in range(lines - 4, lines - 1):
        ws.merge_cells(start_row = i, start_column = 2, end_row = i, end_column = 11)
        ws.cell(row = i, column = 1).border = Border(left = border_thick, right = border_thick, top = border_thick, bottom = border_thick)
        ws.cell(row = i, column = 1).font = Font(bold = True)

    ws.cell(row = lines - 4, column = 1).value = "CMD"
    ws.cell(row = lines - 3, column = 1).value = "일시"
    ws.cell(row = lines - 2, column = 1).value = "비고"
    ws.cell(row = lines - 4, column = 2).value = commandline
    ws.cell(row = lines - 3, column = 2).value = date + ' ' + time

    for i in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
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

def arrange(line):
    return line.replace('[', '').replace(']', '').split()

if __name__ == "__main__":
    iperfiling("test", commandline="iperf -c localhost -t 5")
