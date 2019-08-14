import pandas as pd
import xlrd
import docx2txt

from tkinter import simpledialog
import re
from tkinter import messagebox
def readxls(d):
    dfs=None
    dfs = xlrd.open_workbook(d)
    dfs = dfs.sheet_by_index(0)
    f = []
    for j in range(0, dfs.nrows):
        ur=[]
        for u in range(0, dfs.ncols):
            ur.append(str(dfs.cell_value(j, u)).upper().strip())
        f.append(ur)
    return f
def readasxlsx(d):
    dfs=None


    r=""
    if(d[-3:]=="csv"):
        dfs=pd.read_csv(d)
        for k in dfs:
            for u in dfs[k]:
                r=r+str(u)

    else:
        dfs = xlrd.open_workbook(d)
        dfs=dfs.sheet_by_index(0)
        for j in range(0,dfs.nrows):
            for u in range(0,dfs.ncols):
                r=r+str(dfs.cell_value(j,u))
    return r
def readastxt(d):
    r=docx2txt.process(d)
    return r
def process(rt):
    l=list(set(re.findall(r'[UP][RL][KP]{0,1}\d{2}C[HER]\d{3,4}',rt)))
    le = list(set(re.findall(r'[UP][RL][KP]{0,1}18C[S]\d{3,4}', rt)))
    p=list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}E[EC]\d{3,4}',rt)))
    e = list(set(re.findall(r'[UP][RL][KP]{0,1}\d{2}RA\d{3,4}', rt)))
    f = list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}FP\d{3,4}', rt)))
    f2 = list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}ISD\d{3,4}', rt)))
    f3 = list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}M[TE]\d{3,4}', rt)))
    f6 = list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}BCA\d{3,4}', rt)))
    f4 = list(set(re.findall(r'[UP][LR][KP]{0,1}\d{2}B[IM]\d{3,4}', rt)))
    f5 = list(set(re.findall(r'ULKITCS\d{3,4}', rt)))
    return l+p+f+e+f2+f3+f4+f5+f6+le
def load():
    de=record()
    subt = readastxt1("C:/attendance/files/file.txt")
    return de,subt
def record():
    de = {}
    ul = 0
    for j in range(1,201):
        v="C:/attendance/files/New folder (2)/course%d.XLS" % j
        ul=ul+1


        ug=None
        try:
            ug = readxls(v)

        except:
            print(j)
            continue
        if(len(ug[4][1])>10):

            de[ug[4][1].replace("  "," ")]=[]
            for k in ug:

                if(len(k[9])>2 and len(k[4])>2):
                    de[ug[4][1].strip().upper().replace("  "," ")].append([k[4],k[9]])
    return de
def readastxt1(d):
    file = open(d, "r")
    cont = file.readlines()
    file.close()
    p=[]
    subt = {}
    for j in range(0,len(cont)):
        cont[j]=cont[j].replace("\t"," ")
        cont[j] = cont[j].replace("\n", "")
        cont[j] = cont[j].replace("Print PDF", "")
        cont[j] = cont[j].replace("Export to Excel", "")
        d=re.findall('^(.*)Batch (.*)',cont[j])
        if(len(d)>0 and len(d[0])>0):
            r = d[0][1].strip().split(' ')
            sub = d[0][0].upper() + "BATCH " + r[0]
            k = ""
            for j in r:
                if not j.isdigit():
                    k = k + j + " "
            subt[sub]=k.strip()
    return subt
def gettimetable(dir,day):
    i=2
    day=day.upper().strip()
    l=[]
    s=""
    daylist=["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY"]
    p=daylist.index(day)+1
    src = dir + str(p) + ".txt"
    file = open(src, "r")
    cont = file.readlines()
    for j in cont:
        j=j.replace("\n","")
        j=j.upper().strip()
        l.append(j)
    return l