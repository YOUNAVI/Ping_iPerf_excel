import os
import subprocess
import platform
from datetime import datetime
from pingresult import pingfiling
from iperfresult import iperfiling, iperfiling_udp

if __name__ == '__main__':
    if platform.system() == 'Windows':
        os.system("cls")

    elif platform.system == 'Linux':
        os.system("clear")

    os.makedirs("results", exist_ok=True)
    
    while(1):
        print("주의사항: tee, >, '> hello.txt', logfile, output 등 파일로 출력하는 옵션을 절대 주지 마십시오.")
        cwd = str(os.getcwd())
        command_line = input(f"Python: {cwd}>")
        command_list = command_line.split(' ')
        now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"{command_list[0]}_{now}_result"

        if command_list[0] == 'iperf3' or command_list[0] == 'iperf':
            data = subprocess.Popen(command_line + " --forceflush", stdout=subprocess.PIPE, encoding='euc-kr')
        else:
            data = subprocess.Popen(command_line, stdout=subprocess.PIPE, encoding='euc-kr')

        f = open(filename + ".txt", 'w', encoding="utf-8")
        f.write(command_line)
        while data.poll() == None:
            out = data.stdout.readline()
            f.write(out)
            print(out, end='')

        f.close()

        if command_list[0] == 'ping':
            pingfiling(filename = filename, commandline = command_line)
        elif command_list[0] == 'iperf3' or 'iperf':
            if command_list.__contains__('-u'):
                iperfiling_udp(filename = filename, commandline = command_line)
            else:
                iperfiling(filename = filename, commandline = command_line)

        print(f"filesaved: {filename}")