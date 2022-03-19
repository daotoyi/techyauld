# -*- coding: utf-8 -*-
"""
Created on Thu Jan 11 15:46:02 2018

@author: wenhuas
"""

import csv
from datetime import datetime
import matplotlib.dates as mdate
import matplotlib.pyplot as plt
from matplotlib import style

style.use('fivethirtyeight')
style.use('seaborn-whitegrid')

def process(file,new_file_name):
    file = file
    with open(file,'r') as f:
        line = csv.reader(f, delimiter='|')
        lines = []
        for row in line:
            if len(row) == 3:       #4G SIM7600CE
                if len(str(row)) == 52:
                    s = row
                    #print(s)
                    s[0] = s[0].split(' ')[1]
                    s[2] = s[2].split(' ')[2]
                    s[1] = s[1].split(' ')[1] + ' ' + s[1].split(' ')[2]
                    l = s[0] + ',' + s[1] + ',' + s[2]
                    l += '\n'
                    lines.append(l)
            elif len(row) == 2:     #2G SIM800C
                if len(str(row)) == 37:
                    s = row
                    s[0] = s[0].split(' ')[1]
                    s[1] = s[1].split(' ')[1] + ' ' + s[1].split(' ')[2]
                    l = s[0] + ',' + s[1]
                    l += '\n'
                    lines.append(l)
    with open(new_file_name,'w') as f:
        for line in lines:
            f.write(line)

def plot(file,label,opt):
    filename = file
    label = label
    t = date = []
    s = []
    m = []
    with open(filename,'r') as csvfile:
        plots = csv.reader(csvfile, delimiter=',')
        for row in plots:
            date.append(str(row[1]))
            s.append(int(row[0]))
            m.append(int(row[2]))  #used in 4G mode
        t = [datetime.strptime(d, '%Y-%m-%d %H:%M:%S') for d in date]
        plt.gca().xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
        #plt.gca().xaxis.set_major_locator(mdate.DateLocator())
    if opt == "m":
        plt.plot(t,m, label=label,linewidth=2)
    elif opt == "s":
        plt.plot(t,s,label=label,linewidth=0.8)   

def temp(file,label):
    file = file
    label = label
    t = date = []
    tm = []
    with open(file,'r') as f:
        for row in f.readlines():
            row = row.strip()
            if len(row) == 28:
                row = row[:20] + row[24:]
                date.append(str(row))
            elif len(row) == 5:
                tm.append(float(int(row) / 1000))
            else:
                continue
            t = [datetime.strptime(d, '%a %b %d %H:%M:%S %Y') for d in date]
            plt.gca().xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d'))
    plt.plot(t,tm,label=label,linewidth=1)

class IoT:

    def __init__(self,file,label,file2):
        self.file = file
        self.label = label
        self.file2 = file2
               
    def strength(self,opt):
        plot(self.file,self.label,opt)

    def mode(self,opt):
        plot(self.file,self.label,opt)
        
    def temp(self):
        temp(self.file2,self.label)
    
if __name__ == '__main__':
    ############# 4G IoT
    process('7_24h_str_mod.txt','4G.txt')
    process('7_24h_str_mod_a.txt','4G_a.txt')
    
    p = IoT('4G.txt','4G','cpu_temp.txt')
    p_a = IoT('4G_a.txt','4G_auxiliary','cpu_temp_a.txt')
    
    fig = plt.figure()
    m=s=""
    
#'''  #the place separate 2G & 4G
    #1  mode 
    ax1 = plt.subplot2grid((7,1), (0,0), rowspan=2, colspan=1)
    p.mode('m')
    p_a.mode('m')
    
    plt.ylabel('mode')
    plt.title('IoT test\nnormal temperature & 7_24h')
    plt.legend()
    
    #2  strength
    ax2 = plt.subplot2grid((7,1), (2,0), rowspan=3, colspan=1)
    p.strength('s')
    p_a.strength('s')
    
    plt.ylabel('strength')
    plt.legend()   
    
    #3  temperature
    ax3 = plt.subplot2grid((7,1), (5,0), rowspan=2, colspan=1)
    p.temp()
    p_a.temp()
    
    plt.xlabel('time')
    plt.ylabel('temp')
    plt.legend()
    plt.gcf().autofmt_xdate()
    #plt.xticks(pd.date_range(data.index[0],data.index[-1],freq='1day'))
    plt.show() 
    
'''  #the place separate 2G & 4G    
    #############  2G IoT 
    ''
    process('7_24h_str_mod.txt','2G.txt')
    process('7_24h_str_mod_a.txt','2G_a.txt')    
    
    p = IoT('2G.txt','2G','cpu_temp.txt')
    p_a = IoT('2G_a.txt','2G_auxiliary','cpu_temp_a.txt')  
    
    #1  temperature
    ax1 = plt.subplot2grid((6,1), (0,0), rowspan=2, colspan=1)
    p.temp()
    p_a.temp()
    
    plt.title('IoT test\nnormal temperature & 7_24h')
    plt.ylabel('temp')
    
    #2  strength
    ax2 = plt.subplot2grid((6,1), (2,0), rowspan=4, colspan=1)
    p.strength('s')
    p_a.strength('s')
    
    plt.ylabel('strength')
    plt.legend()   
    plt.gcf().autofmt_xdate()
    plt.show()
'''  #the place separate 2G & 4G          