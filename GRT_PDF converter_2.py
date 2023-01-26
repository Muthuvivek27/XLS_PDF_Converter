from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as mb
from tkinter import ttk
from PyPDF2 import PdfFileReader
import pandas as pd
import shutil
from sys import exit
from pathlib import Path
import win32com.client
from win32com.client.gencache import EnsureDispatch
from pywintypes import com_error
import os, glob
from pathlib import Path
global Srcfdr


def PDF_CONVERT(Srcfdr):
    global pgno
    global name,l,l2,del_index

    Srcfdr1 = Srcfdr + "\Excel\*.xls"

    a = glob.glob(Srcfdr1)
    filename = []
    for i in a: filename.append(Path(i))

    print(filename)
    #l.destroy()
    #l = Label(sts, text=filename,bg='#FBEEE6')
    #l.grid(column=0, row=1, padx=20)
    #sts.update()
    c=int(1)
    for i in filename:
        WB_PATH = i
        # PDF path when saving\n",
        PATH_TO_PDF = os.path.join(i.parent, i.stem + '.pdf')
        #print(PATH_TO_PDF)
        l.destroy()
        l = Label(sts, text=i.stem, bg='#FBEEE6')
        l.grid(column=0, row=1, padx=20)
        sts.update()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            #print('Start conversion to PDF')
            l2.destroy()
            l2 = Label(sts, text="Start conversion to PDF", bg='#FBEEE6')
            l2.grid(column=0, row=2, padx=20)
            sts.update()
            wb = excel.Workbooks.Open(WB_PATH)
            if del_index == 1:
                print("varala?????")
                l2.destroy()
                l2 = Label(sts, text="Deleting Index", bg='#FBEEE6')
                l2.grid(column=0, row=2, padx=20)
                sts.update()
                try:
                    #excel = EnsureDispatch("Excel.Application")
                    sheet = wb.Worksheets("Index")
                    #sheet.Delete
                    excel.DisplayAlerts = False
                    sheet.Delete()
                    print("aacha??")
                except:
                    mb.showwarning("Warning", 'Index Sheet Not found!')
            wb.ExportAsFixedFormat(0, PATH_TO_PDF, IgnorePrintAreas=True)
            pdf1=open(PATH_TO_PDF, 'rb')
            pdf = PdfFileReader(pdf1)
            print(pdf.getNumPages())
            pgno.append(pdf.getNumPages())
            print(i.stem)
            name.append(i.stem)
            pdf1.close()

        except com_error as e:
            #print('failed.')
            l2.destroy()
            l2 = Label(sts, text="Failed", bg='#FBEEE6')
            l2.grid(column=0, row=2, padx=20)

            button1 = Button(sts, text="Exit", command=exit, width=10, state="normal", bg='#AED6F1').grid(row=3,
                                                                                                             column=1)
            sts.update()
        else:
            #print('Succeeded.')
            l2.destroy()
            l2 = Label(sts, text="Succeeded", bg='#FBEEE6')
            l2.grid(column=0, row=2, padx=20)
            sts.update()
        finally:
            #wb.save = true
            wb.Close(SaveChanges=1)
            #excel.Quit()



def PDF_Start():
    global pgno
    global name,l,l2
    pgno = []
    name = []
    Srcfdr = folderpath
    Srcfdr1 = Srcfdr + "\*.xls"
    temp1 = Srcfdr + "\Excel"
    output1 = Srcfdr + "\PDF"
    output2 = Srcfdr + "\Excel\*.*"
    print(output2)
    os.makedirs(temp1, exist_ok=True)
    os.makedirs(output1, exist_ok=True)

    # os.mkdir(temp1)
    # os.mkdir(output1)

    full = glob.glob(Srcfdr1)
    x = len(full)
    print(x)
    while x != 0:
        a = glob.glob(Srcfdr1)
        filename = []
        for i in a: filename.append(Path(i))
        k = 1
        for i in filename:
            dest = shutil.move(i, Path(temp1, i.name))
            k = k + 1
            if (k == 4):
                break
        PDF_CONVERT(Srcfdr)
        b = glob.glob(output2)
        filename2 = []
        for i in b: filename2.append(Path(i))
        for i in filename2:
            dest2 = shutil.move(i, Path(output1, i.name))
        full = glob.glob(Srcfdr1)
        x = len(full)
    Srcfdr3 = output1 + "\*.xls"
    a = glob.glob(Srcfdr3)
    filename3 = []
    for i in a: filename3.append(Path(i))
    for i in filename3:
        dest = shutil.move(i, Path(temp1, i.name))

    l.destroy()
    l = Label(sts, text="Your Files have been Converted to PDF and is available in Folder: \t \n " + output1 + ":", bg='#FBEEE6')
    l.grid(column=0, row=1, padx=20)
    button1 = Button(sts, text="Exit", command=exit, width=10, state="normal", bg='#AED6F1').grid(row=3,
                                                                                                     column=0)
    sts.update()
    #print("Your Files have been Converted to PDF and is available in Folder: \t \n " + output1 + ":")
    #print(name)
    #print(pgno)
    pagenumber = pd.DataFrame()
    pagenumber['Cabinet'] = name
    pagenumber['No_pgs'] = pgno

    pagenumber.to_excel(Srcfdr + "\\" + 'Pagenumber.xlsx', index=False)


def folder():
    global folderpath
    global Folderpath
    global crtpjt
    folderpath = filedialog.askdirectory(initialdir='C:\\')
    Folderpath.insert(0,folderpath)
    crtpjt.update()
def status():
    global sts,test,l,l2,del_index
    test="Starting"
    try:
        sts.destroy()
    except:
        print("")

    sts = Tk()
    sts.title('Status')
    sts.geometry('450x250')
    sts.configure(bg='#FBEEE6')
    l = Label(sts, text=test, bg='#FBEEE6')
    l.grid(column=0, row=1, padx=20)
    l2 = Label(sts, text="Starting", bg='#FBEEE6')
    l2.grid(column=0, row=2, padx=20)
    del_index = del_index.get()
    PDF_Start()
    sts.mainloop()



def homescreen():
    global Folderpath, folderpath
    global test
    global crtpjt,del_index

    folderpath = ''

    crtpjt = Tk()
    crtpjt.title('Home')
    crtpjt.geometry('450x300')
    crtpjt.configure(bg='#FBEEE6')
    Folderpath = Entry(crtpjt)
    del_index = IntVar(crtpjt)
    Folderpath.insert(0, '')
    Folderpath.grid(column=1, row=1, ipadx=35)
    l_title = Label(crtpjt, text="      GRT Post Gen Automation", bg='#FBEEE6').grid(column=0, row=0, padx=20, pady=20,columnspan = 3)
    l = Label(crtpjt, text="Select Folder", bg='#FBEEE6').grid(column=0, row=1, padx=20)
    l3 = Label(crtpjt, text="  ", bg='#FBEEE6').grid(column=1, row=4)

    button3 = Button(crtpjt, text="Browse", command=folder, width=10, state="normal", bg='#AED6F1').grid(row=1,
                                                                                                         column=4,
                                                                                                         padx=10)
    l = Label(crtpjt, text="\n", bg='#FBEEE6').grid(column=0, row=3, padx=20)
    l1 = Label(crtpjt, text="Delete Index Sheet:", bg='#FBEEE6').grid(column=0, row=4, padx=20)
    c1 = Checkbutton(crtpjt, variable=del_index, bg='#FBEEE6')
    c1.grid(column=1, row=4, padx=10)
    l = Label(crtpjt, text="\n", bg='#FBEEE6').grid(column=0, row=5, padx=20)
    button2 = Button(crtpjt, text="Convert", command=status, width=12, state="normal", bg='#AED6F1').grid(
        row=6, column=0)
    button1 = Button(crtpjt, text="Exit", command=exit, width=10, state="normal", bg='#AED6F1').grid(row=6,
                                                                                                     column=1)
    l5 = Label(crtpjt, text= "\n\n" + "V2.02", bg='#FBEEE6',font=("Courier", 7)).grid(column=4, row=7, padx=20)

    crtpjt.mainloop()

if __name__ == '__main__':
   homescreen()
