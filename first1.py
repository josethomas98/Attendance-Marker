from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import calendar
from tkinter import colorchooser
from tkinter import simpledialog
import tkinter.scrolledtext as tkscrolled
import datetime
import docx2txt
import pandas as pd
import xlrd
import re
import numpy as np
from tkinter.font import Font
from pandas import DataFrame
import supporter
import csv
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
year=["misc","18","17","16","15"]
dic={year[0]:[],year[1]:[],year[2]:[],year[3]:[],year[4]:[]}
entire=[]
total=0
now = datetime.datetime.now()
final={}
faculty={}
var1=var2=var3=var4=0
info=['']*5
info[0]="Mr.XXXXXXX \n(Head of Department) \nDepartment of Computer Science"
info[1]="Mr.XXXXXXX \n(Head of Department) \nDepartment of Mechanical Engineering"
info[2]="PERMISSION FOR PHISICAL ATTENDENCE"
info[3]="please sir we are in shortage of attendence please give is some"
info[4]="%d-%d-%d" %(now.day,now.month,now.year)
ug=[]
ug=supporter.readxls('C:/attendance/refer1.xlsx')
de={}
subt={}
de,subt=supporter.load()

def rename():
    result = simpledialog.askstring("Rename","enter the new name of Project")
    root.title("%s" %(result))

def fullscreen():
    root.geometry("%dx%d+0+0" %(root.winfo_screenwidth(),root.winfo_screenheight()))
def customscreen():
    global var1,var2,var3,var4
    cs=Toplevel()
    cs.iconbitmap('abcdef.ico')
    var3 = IntVar(cs)
    var4 = IntVar(cs)
    var1 = IntVar(cs)
    var2 = IntVar(cs)
    var1.set("500")
    var2.set("500")
    cs.title("Modify Screen Size")
    cs.geometry("300x300+250+250")
    cs.resizable(width=False,height=False)
    cs_l1=Label(cs,text="change the values and enter ok")
    cs_l1.grid(row=0,columnspan=2,padx=60,rowspan=3,ipady=30)
    cs_l2=Label(cs,text="Height")
    cs_l2.grid(row=4,column=0,pady=10)
    cs_l3 = Label(cs, text="Width")
    cs_l3.grid(row=5, column=0,pady=10)
    spin1=Spinbox(cs,from_=1,to=cs.winfo_screenwidth(),textvariable=var1)
    spin1.grid(row=4,column=1,pady=10)
    spin2 = Spinbox(cs, from_=1, to=cs.winfo_screenwidth(),textvariable=var2)
    spin2.grid(row=5, column=1,pady=10)
    c1=Checkbutton(cs,text="Prevent resize of width",variable=var3)
    c1.grid(row=6,columnspan=2,pady=5)
    c1 = Checkbutton(cs, text="Prevent resize of height",variable=var4)
    c1.grid(row=7, columnspan=2,pady=5)
    cs_apply=Button(cs,text="Apply",command=csdone)
    cs_apply.grid(row=8,column=0)
    cs_cancel=Button(cs,text="Cancel",command=cs.destroy)
    cs_cancel.grid(row=8,column=1)
    cs.mainloop()
def csdone():
    global var1,var2,var3,var4
    try:
        csh=int(var1.get())
        csw =int(var2.get())
        if(csw<0 or csw>root.winfo_screenwidth()):
            result=messagebox.showwarning("invalid width","Width must be between 0 an %d" %(root.winfo_screenwidth()))
            print(1/0)
        if (csh < 0 or csh > root.winfo_screenheight()):
            result = messagebox.showwarning("invalid Height","Height must be between 0 an %d" % (root.winfo_screenheight()))
            print(1/0)
        root.resizable(width=(var3.get()+1)%2, height=(var4.get()+1)%2)
        root.geometry("%dx%d+0+0" % (csw,csh))
    except:
        result = messagebox.showwarning("ERROR", "Wrong Input")
def bgcolor():
    clr = colorchooser.askcolor(title="select color")
    root.configure(background=clr[1])


def pagelayout():
    global info
    cur_date=""
    pg=Toplevel()
    pg.iconbitmap('abcdef.ico')
    pg.config(bg="#fcef99")
    pg.geometry("800x%d+100+0" %(pg.winfo_screenheight()-90))
    pg.title("Page contents")
    pg.resizable(width=False, height=False)
    pgt=[None]*5
    ch=[None]*5
    pgc=[None]*5
    sh=[None]*5 #y?
    for j in range(5):
        ch[j]=IntVar()
        sh[j]=StringVar()
        if(j<3):
            pgt[j] = Text(pg, height=6, width=80,wrap=WORD)
        elif(j==3):
            pgt[j] = tkscrolled.ScrolledText(pg, height=10, width=80,wrap=WORD)
        else:
            pgt[j]=Text(pg,height=1,width=20,wrap=WORD)
        pgt[j].insert(INSERT, info[j])

    def activate(i):
        if ch[i].get()==1:
            pgt[i].config(state=DISABLED)
        else:
            pgt[i].config(state=NORMAL)
    pgc[0] = Checkbutton(pg, text="Disable from :",bd=1,variable=ch[0],command=lambda: activate(0)).pack()
    pgt[0].pack()
    pgc[1] = Checkbutton(pg, text="Disable To :", bd=1, variable=ch[1], command=lambda: activate(1)).pack()
    pgt[1].pack()
    pgc[2] = Checkbutton(pg, text="Disable subject :", bd=1, variable=ch[2], command=lambda: activate(2)).pack()
    pgt[2].pack()
    pgc[3] = Checkbutton(pg, text="Disable Body :", bd=1, variable=ch[3], command=lambda: activate(3)).pack()
    pgt[3].pack()
    pgc[4] = Checkbutton(pg, text="Disable Current Date :", bd=1, variable=ch[4],command=lambda: activate(4)).place(x=70,y=580)
    pgt[4].place(x=270,y=580)
    def pgsave():
        for j in range(5):
            if(ch[j].get()==0):
                info[j]=pgt[j].get(1.0,END)
            else:
                info[j]=""
        pg.destroy()
    pg_save=Button(pg,text="Save",command=pgsave)
    pg_cancel=Button(pg,text="cancel",command=pg.destroy)
    pg_cancel.place(x=270,y=630)
    pg_save.place(x=70, y=630)


    pg.mainloop()
li=[]
waste = ["misc", "first year", "second year", "third year", "fourth year", "total"]

def update_status():
    for j in range(0,6):
        status_t[j].config(state=NORMAL)
        status_t[j].delete(1.0,END)
        if(j==5):
            status_t[j].insert(INSERT,str(total))
        else:
            status_t[j].insert(INSERT, str(len(dic[year[j]])))
        status_t[j].config(state=DISABLED)
    canvas.config(scrollregion=(0, 0, 1000, (total+2) * 40+10))
def open_file():
   global entire,total,year,dic
   g=[]
   try:
       result =  filedialog.askopenfile(initialdir="/", title="select file", filetypes=(("XLS files", ".XLS"),("text files", ".txt"),("CSV File",".csv"),("Microsoft Excel Worksheet",".xlsx"),("Microsoft Document",".docx")))
       if(result==None):
           t=messagebox.askquestion("confirmation","Would you like to reselect")
           if(t=="yes"):
               open_file()
       else:
           result = str(result)
           result = result.replace('\'', '-')
           d = re.findall(r'\bname=-(.*)- mode', result)
           d=str(d[0])
           r=""
           if(d[-4:]=="docx"):

               r=supporter.readastxt(d)
           elif(d[-4:]=="xlsx" or d[-3:]=="csv" or d[-3:]=="XLS"):
               r=supporter.readasxlsx(d)

           elif(d[-3:]=="txt"):
               file=open("%s"%(d),"r")
               r=file.read()
           r=r.upper()
           g=supporter.process(r)
           g.sort()
           for item in g:
               if (item not in entire) and len(item)>2:
                   entire.append(item)
                   total+=1
                   if year[1] in item[:-4]:
                       dic[year[1]].append(item)
                       root_listbox.insert(END,item)
                   elif(year[2] in item[:-4]):
                       dic[year[2]].append(item)
                       root_listbox.insert(END,item)
                   elif(year[3] in item[:-4]):
                       dic[year[3]].append(item)
                       root_listbox.insert(END,item)
                   elif(year[4] in item[:-4]):
                       dic[year[4]].append(item)
                       root_listbox.insert(END,item)
                   else:
                       dic[year[0]].append(item)
                       root_listbox.insert(END,item)

           update_status()
   except Exception as e:
        print(e)

def populateMethod(m,student,staff):
    status_t[6] = Text(root, width=70, height=10,yscrollcommand=True)
    status_t[6].config(state=NORMAL)
    status_t[6].grid(row=4, column=7, columnspan=20, rowspan=2000, sticky=N)
    stu=""
    for j in student[m]:
        stu=stu+"  "+j[0] +"    "+j[1]+"\n"
    status_t[6].insert(INSERT,"\n "+m+"\n  "+staff[m]+"\n\n"+stu)
    status_t[6].config(state=DISABLED)
def proceed():
    global entire, total, year, dic,de,subt,final,faculty
    canvas.delete("all")
    month=['January', 'February', 'March', 'April', 'May', 'June', 'July','August', 'September', 'October', 'November', 'December']
    #dateinfo=month[int(info[4].split("-")[0])-2]+" "+info[4].split("-")[1]+", "+info[4].split("-")[2]
    #day=datetime.datetime.strptime(dateinfo,'%B %d, %Y').strftime('%A').upper()

    sublist=supporter.gettimetable("C:/attendance/files/timetable/","thursday")
    final={}
    faculty={}

    #for k in de:
     #   if k in sublist:
      #      print(k)

    for j in sublist:
        if j in de:
            temp=[]
            for k in de[j]:
                if(k[0] in entire):
                    temp.append(k)
            if(len(temp)>0):
                final[j]=temp
                faculty[j]=subt[j]
    for j in  final:
        print(j,faculty[j])
        print("")
        print(final[j])
    buttonlist=[Button(canvas, text="Quit", command=quit, anchor=W)]*len(final)
    x=10
    y=10
    for j in range(0,len(final)):
        buttonlist[j] = Button(canvas, text=list(final.keys())[j], command= lambda m=list(final.keys())[j]: populateMethod(m,final,faculty), anchor=W,bg="white")
        buttonlist[j].configure(width=50, activebackground="red", relief=FLAT)
        button1_window = canvas.create_window(x,y,anchor=NW, window=buttonlist[j])
        if(j%2==0):
            x=400
        else:
            x=10
            y=y+45

    #for j in range(1,total+1):
     #   for k in range(0, 9):
      #      canvas.create_rectangle(10 + k * 107, 20 + j * 40, 10 + 107 * (k + 1), 20 + (j+1) * 40,fill="white")


def manual():
    global entire, total,year, dic
    def mansave():
        global entire, total, year, dic
        v=man_t.get(1.0,END).upper()
        g=[]
        g=supporter.process(v)
        for item in g:
            if (item not in entire) and len(item) > 2:
                entire.append(item)
                total=total+1
                if year[1] in item[:-4]:
                    dic[year[1]].append(item)
                    root_listbox.insert(END, item)
                elif (year[2] in item[:-4]):
                    dic[year[2]].append(item)
                    root_listbox.insert(END, item)
                elif (year[3] in item[:-4]):
                    dic[year[3]].append(item)
                    root_listbox.insert(END, item)
                elif (year[4] in item[:-4]):
                    dic[year[4]].append(item)
                    root_listbox.insert(END, item)
                else:
                    dic[year[0]].append(item)
                    root_listbox.insert(END, item)
        update_status()
        man.destroy()
    man=Toplevel()
    man.geometry("200x300+%d+%d" %((man.winfo_screenwidth()/2-50),(man.winfo_screenheight()/2-150)))
    man.resizable(width=False, height=False)
    man.title("Manual Input")
    my_font = Font(family="Arial", size=22, weight="bold")
    man_t=tkscrolled.ScrolledText(man, height=7, width=10,wrap=WORD,font=my_font)
    man_t.place(x=0,y=20)
    manb1=Button(man,text="Save",fg="green",width=10,command=mansave)
    manb1.pack(side=LEFT,anchor="s")
    manb2 = Button(man, text="Cancel", fg="red", width=10,command=man.destroy)
    manb2.pack(side=RIGHT, anchor="s")


    man.mainloop()

def removeselected():
    global entire, total, year, dic
    if root_listbox.curselection():
        val=root_listbox.get(root_listbox.curselection())
        if val in entire:
            entire.remove(val)
            total-=1
            for j in range(0,5):
                if val in dic[year[j]]:
                    dic[year[j]].remove(val)
            update_status()
        root_listbox.delete(root_listbox.curselection()[0])
        canvas.config(scrollregion=(0, 0, 1000,total*50))

def export():
    global final,faculty
    src = "C:/Users/Stino Thomas/PycharmProjects/first1/OUTPUT/"
    with open(src+'complete_details.csv','w') as writeFile:
        writer= csv.writer(writeFile)
        r=['','','','DETAILS']
        writer.writerow(r)
        count = 1
        r = ['slno', 'Name']
        writer.writerow(r)
        r=['','','','FIRST YEAR']
        writer.writerow(r)

        for j in dic[year[1]]:
            r=[count,j]
            count=count+1
            writer.writerow(r)
        r = ['', '', '', 'SECOND YEAR']
        writer.writerow(r)
        for j in dic[year[2]]:
            r = [count, j]
            count = count + 1
            writer.writerow(r)
        r = ['', '', '', 'THIRD YEAR']
        writer.writerow(r)
        for j in dic[year[3]]:
            r = [count, j]
            count = count + 1
            writer.writerow(r)
        r = ['', '', '', 'FOURTH YEAR']
        writer.writerow(r)
        for j in dic[year[4]]:
            r = [count, j]
            count = count + 1
            writer.writerow(r)
        r = ['', '', '', 'OTHER']
        writer.writerow(r)
        for j in dic[year[0]]:
            r = [count, j]
            count = count + 1
            writer.writerow(r)


        writeFile.close()
    if(len(final)>0):

        for j in final:
            tsrc=src+faculty[j]+" - "+j+".csv"
            with open(tsrc, 'w') as writeFile:
                writer = csv.writer(writeFile)
                count=1
                r=['','','','','',j]
                writer.writerow(r)
                r = ['', '', '', 'FACULTY', '', faculty[j]]
                writer.writerow(r)
                for p in final[j]:
                    r=[count,p[0],p[1]]
                    count=count+1
                    writer.writerow(r)
                writeFile.close()


root=Tk()


root.iconbitmap('abcdef.ico')
root.title("samp1e01")
window_width=700
window_height=700
screen_width=root.winfo_screenwidth()
screen_height=root.winfo_screenheight()
x_cordinate=(screen_width/2) - (window_width/2)
y_cordinate=(screen_height/2) - (window_height/2)
root.geometry("%dx%d" %(screen_width,screen_height))
root.wm_minsize(width=500,height=430)


main_menu=Menu(root)
root.config(menu=main_menu)
root.bind('<Alt_L><q>', lambda e:fullscreen())
root.bind('<Alt_L><w>', lambda e:customscreen())
root.bind('<Alt_L><r>', lambda e:rename())
root.bind('<Alt_L><i>', lambda e:pagelayout())
#creating a file menu button
file_menu=Menu(main_menu)
main_menu.add_cascade(menu=file_menu,label="File")
file_menu.add_command(label="Rename",command=rename,accelerator="ALT+R")
file_menu.add_command(label="Close ",command=quit,accelerator="ALT+f4")
file_menu.add_command(label="Import",command=open_file)
#creating a insert
insert_menu=Menu(main_menu)
main_menu.add_cascade(menu=insert_menu,label="Insert")


#creating view menu
view_menu=Menu(main_menu)
main_menu.add_cascade(menu=view_menu,label="View")
view_menu.add_command(label="Full Screen",command=fullscreen,accelerator="ALT+Q")
view_menu.add_command(label="Custom Screen Size",command=customscreen,accelerator="ALT+W")
view_menu.add_command(label="Change Background Color",command=bgcolor)




#creating page layout menu
page_menu=Menu(main_menu)
main_menu.add_cascade(menu=page_menu,label="Page Layout")
page_menu.add_command(label="contents",command=pagelayout,accelerator="ALT+i")



my_font = Font(family="ALGERIAN", size=32, weight="bold")
root_label1=Label(root,text="ATTENDANCE",font=my_font,bd=10,bg="#edf2f9",width=root.winfo_width()//4)
root_label1.grid(row=0,column=0,columnspan=100)
x_image_for_button = PhotoImage(file='man.png')
x_image_for_button1 = PhotoImage(file='img1.png')
root_b1=Button(root,image=x_image_for_button,width=140,height=140,command=manual,relief=SOLID)
root_b1.grid(row=2,column=1,pady=10)
root_b2=Button(root,image=x_image_for_button1,width=140,height=140,command=open_file,relief=SOLID)
root_b2.grid(row=3,column=1)
root_label2=Label(root,text="Status Bar",width=40,bg="gray",fg="black",font=Font(family="Arial", size=11, weight="bold"))
root_label2.grid(row=4,column=0,columnspan=7,pady=20,sticky="s")
status_t=[None]*7
status_l=[None]*6

for j in range(6):
    status_t[j]=Text(root,width=30,height=1)
    status_t[j].grid(row=4+j+1,column=3,columnspan=3,sticky="nw",pady=2)


for j in range(6):
    status_t[j].insert(INSERT,"0")
    status_t[j].config(state=DISABLED)
    status_l[j]=Label(root,text=waste[j])
    status_l[j].grid(row=4 + j+1, column=1, columnspan=3, sticky="nw",pady=6)
root_listbox=Listbox(root,width=30,height=20,relief=SOLID,font=Font(family="Arial", size=9, weight="bold"),yscrollcommand=True)
root_listbox.grid(row=2,column=5,pady=20,rowspan=2)
x_image_for_button3 = PhotoImage(file='cancel.png')
root_b3=Button(root,image=x_image_for_button3,width=20,height=30,command=removeselected)
root_b3.grid(row=2,column=6,sticky='nw',pady=20,rowspan=2)
x_image_for_button4 = PhotoImage(file='update.png')
root_b4=Button(root,image= x_image_for_button4,width=20,height=30,command=proceed)
root_b4.grid(row=3,column=6,sticky='nw',pady=20)



root_canvas=Frame(root,bg="white",width=700,height=350,relief=SUNKEN)
root_canvas.grid_rowconfigure(0, weight=2)
root_canvas.grid_columnconfigure(0, weight=2)
xscrollbar = Scrollbar(root_canvas, orient=HORIZONTAL)
xscrollbar.grid(row=1, column=0, sticky=E+W)

yscrollbar = Scrollbar(root_canvas)
yscrollbar.grid(row=0, column=1, sticky=N+S)

canvas = Canvas(root_canvas, bd=4,bg="white",width=880,height=265,scrollregion=(0,0,1000,8),highlightbackground="black",
                xscrollcommand=xscrollbar.set,
                yscrollcommand=yscrollbar.set)
canvas.config(background="#f9ffad")
canvas.grid(row=0, column=0, sticky=N+S+W)
xscrollbar.config(command=canvas.xview)
yscrollbar.config(command=canvas.yview)
root_canvas.grid(row=2,column=8,rowspan=80,columnspan=20,sticky=N+E,pady=20,padx=10)
fl=PhotoImage(file='download1.png')
download=Button(root,image=fl,width=140,height=140,command=export,relief=SOLID)
download.grid(row=5, column=23, columnspan=30, rowspan=200, sticky=N)
root.mainloop()




