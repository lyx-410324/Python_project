#!/usr/bin/env python
# @Time:
# @Author:LYX
# @File:scanport.py
# @Software:PyCharm
#TS_pyserial.py
import sys
#print ("Python Version {}".format(str(sys.version).replace('\n', '')))
#Python Version 3.6.8 |Anaconda, Inc.| (default, Feb 21 2019, 18:30:04) [MSC v.1916 64 bit (AMD64)]

import serial.tools.list_ports
import time

'''Function design, single entry single exit principle'''
def scan_ports_info(Handle_print=0):
    ''''''
    ports_list = list(serial.tools.list_ports.comports())
    if Handle_print:
        print("[info]ports number:",len(ports_list))

    if len(ports_list) <= 0:
        print("The Serial port can't find!")
    else:
        for i in range(len(ports_list)):
            port_list = list(ports_list[i])
            port_serialName = port_list[0]
            if Handle_print:
                print('\n[info]port list', port_list)
                print('[info]port serial Name',port_serialName)
    return ports_list
def getportlist():
    port_serialName = list()
    ports = scan_ports_info(Handle_print=0)
    for i in range(len(ports)):
        port_list = list(ports[i])
        port_serialName.insert(i, port_list[0])  # = port_list[0]
        print('\n[info]port list', port_list)
        print('[info]port serial Name', port_serialName[i])
    print(port_serialName)
    return port_serialName


from time import sleep

def recv(serial):
    while True:
        data = serial.read_all()
        if data == '':
            time.sleep(0.02)
            continue
        else:
            break
    return data





if __name__=="__main__":
    getportlist()
    '''
    port_serialName = list()
    ports=scan_ports_info(Handle_print=0)
    for i in range(len(ports)):
        port_list = list(ports[i])
        port_serialName.insert(i,port_list[0]) #= port_list[0]
        print('\n[info]port list', port_list)
        print('[info]port serial Name', port_serialName[i])
    print (port_serialName)
    '''
