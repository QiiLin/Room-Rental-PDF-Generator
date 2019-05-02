import sys
import os
import win32com.client
from mailmerge import MailMerge
from tkinter import *


def action(event):
    with MailMerge(os.getcwd() + "\\test.docx") as doc:
    ## print out all field that is ready to merge
        print (doc.get_merge_fields())
        doc.merge(StartYear= StartYear_v.get(),
        Amount= Amount_v.get(),
        RentalInfo=RentalInfo_v.get(),
        StartMonth=StartMonth_v.get(),
        IDB=IDB_v.get(),
        PartyB=PartyB_v.get(),
        IDA=IDA_v.get(),
        PartyA=PartyA_v.get(),
        StartDay=StartDay_v.get(),
        TeleA=TelephoneA_v.get(),
        TeleB=TelephoneB_v.get()
        )
        doc.write('output.docx')
        doc.close();
    wdFormatPDF = 17
    in_file = os.path.abspath(os.getcwd() + "\\output.docx")
    out_file = os.path.abspath("test.pdf")
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    os.remove(os.getcwd()+'\\output.docx')

root = Tk();
root.geometry("300x280")
root.title("SimpleGen")

## create label
StartYear = Label(root, text="StartYear")
Amount = Label(root, text="Amount")
RentalInfo = Label(root, text="RentalInfo")
StartMonth = Label(root, text="StartMonth")
IDB = Label(root, text="IDB")
PartyB = Label(root, text="PartyB")
IDA = Label(root, text="IDA")
PartyA = Label(root, text="PartyA")
StartDay = Label(root, text="StartDay")
TelephoneA = Label(root, text="TelephoneA")
TelephoneB = Label(root, text="TelephoneB")

## place lable 
StartYear.grid(row=0, sticky=E);
StartMonth.grid(row=1, sticky=E);
StartDay.grid(row=2, sticky=E);
Amount.grid(row=3, sticky=E);
RentalInfo.grid(row=4, sticky=E);
PartyA.grid(row=5, sticky=E);
IDA.grid(row=6, sticky=E);
PartyB.grid(row=7, sticky=E);
IDB.grid(row=8, sticky=E);
TelephoneA.grid(row=9, sticky=E);
TelephoneB.grid(row=10, sticky=E);

## bind entry to value
StartYear_v = StringVar()
Amount_v = StringVar()
RentalInfo_v = StringVar()
StartMonth_v = StringVar()
IDB_v = StringVar()
PartyB_v = StringVar()
IDA_v = StringVar()
PartyA_v = StringVar()
StartDay_v = StringVar()
TelephoneA_v = StringVar();
TelephoneB_v = StringVar();

StartYear_f = Entry(textvariable=StartYear_v)
Amount_f = Entry(textvariable=Amount_v)
RentalInfo_f = Entry(textvariable=RentalInfo_v)
StartMonth_f = Entry(textvariable=StartMonth_v)
IDB_f = Entry(textvariable=IDB_v)
PartyB_f = Entry(textvariable=PartyB_v)
IDA_f = Entry(textvariable=IDA_v)
PartyA_f = Entry(textvariable=PartyA_v)
StartDay_f = Entry(textvariable=StartDay_v)
TelephoneA_f = Entry(textvariable=TelephoneA_v);
TelephoneB_f = Entry(textvariable=TelephoneB_v);

## place entry in grid
StartYear_f.grid(row=0, column=1);
StartMonth_f.grid(row=1, column=1);
StartDay_f.grid(row=2, column=1);
Amount_f.grid(row=3, column=1);
RentalInfo_f.grid(row=4, column=1);
PartyA_f.grid(row=5, column=1);
IDA_f.grid(row=6, column=1);
PartyB_f.grid(row=7, column=1);
IDB_f.grid(row=8, column=1);
TelephoneA_f.grid(row=9, column=1);
TelephoneB_f.grid(row=10, column=1);

gener = Button(root, text = "Generate output.pdf", fg= "black")
gener.bind("<Button-1>", action)
gener.grid(row=11, columnspan=2)


root.mainloop()

