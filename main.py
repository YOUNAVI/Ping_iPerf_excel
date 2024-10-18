import os
import subprocess
import platform
from datetime import datetime
from pingresult import pingfiling
from pingresultlinux import pingfilinglinux
from iperfresult import iperfiling, iperfiling_udp
from iperfresultlinux import iperfilinglinux, iperfilinglinux_udp
from hrpingresult import hrpingfiling

if __name__ == '__main__':
    if platform.system() == 'Windows':
        os.system("cls")

    elif platform.system() == 'Linux':
        os.system("clear")

    os.makedirs("results", exist_ok=True)
    
    while(1):
        print("ping, iperf, hrping to excel")
        print("주의사항: tee, >, '> hello.txt', logfile, output 등 파일로 출력하는 옵션을 절대 주지 마십시오.")
        cwd = str(os.getcwd())
        command_line = input(f"Python: {cwd}> ")
        command_list = command_line.split(' ')
        now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"{command_list[0]}_{now}_result"

        if command_list[0] == 'iperf3' or command_list[0] == 'iperf' or command_list[0] == 'hrping':
            command_list.append("--forceflush")
            data = subprocess.Popen(command_list, stdout=subprocess.PIPE, encoding='euc-kr')
        else:
            data = subprocess.Popen(command_list, stdout=subprocess.PIPE, encoding='euc-kr')

        if command_line.__contains__("iperf3 -s"):
            print(f"filesaved: iperf3_server.txt\n")
            f = open(f"iperf3_server.txt", 'w', encoding = "utf-8")
        else:
            f = open(filename + ".txt", 'w', encoding="utf-8")
        f.write(command_line)
        while data.poll() == None:
            out = data.stdout.readline()
            f.write(out)
            print(out, end='')

        f.close()

        if command_list[0] == 'ping':
            if platform.system() == 'Windows':
                pingfiling(filename = filename, commandline = command_line)
            elif platform.system() == 'Linux':
                pingfilinglinux(filename = filename, commandline = command_line)
        elif command_list[0] == 'iperf3' or command_list[0] == 'iperf':
            if command_list.__contains__('-u'):
                if platform.system() == 'Windows':
                    iperfiling_udp(filename = filename, commandline = command_line)
                elif platform.system() == 'Linux':
                    iperfilinglinux_udp(filename = filename, commandline = command_line)    
            else:
                if platform.system() == 'Windows':
                    iperfiling(filename = filename, commandline = command_line)
                elif platform.system() == 'Linux':
                    iperfilinglinux(filename = filename, commandline = command_line)
        elif command_list[0] == 'hrping':
            hrpingfiling(filename = filename, commandline = command_line)

        print(f"filesaved: {filename}\n")