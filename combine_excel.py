import sys
import csv
import glob
import pandas as pd
import os
from tkinter import *
from tkinter import filedialog

window = Tk()
window.title("Merge Excel")
window.geometry('350x200')

#Header 
heading = Label(window, text="Merge Excel", font=("Arial Bold", 15), anchor='w')
heading.grid(column=0, row=0)

#Source directory Label 
source_lbl = Label(window, text="Excel Source Directory")
source_lbl.grid(column=0, row=3)

#destination directory Label
destination_lbl = Label(window, text="File Destination", anchor='w')
destination_lbl.grid(column=0, row=5)

#Number of rows(header) in each excel file
Label(window, text="Number of headers").grid(row = 7, column = 0) 
#txt = Entry(window,width=10)
 
#txt.grid(column=1, row=0)
 
def clicked_destination():
 
    window.filename =  filedialog.asksaveasfilename(initialdir = "/",title = "Select file",filetypes = (("Excel file","*.xlsx"),("all files","*.*")))
    #window.filename = filedialog.askdirectory()
    #print(window.filename)
    destination_lbl.configure(text= window.filename+".xlsx")

def clicked_source():

 
    window.directory = filedialog.askdirectory()
    source_lbl.configure(text= window.directory)

def combine(*args):

    row_count = int(header.get())
    print(row_count)
    #filenames
    path = source_lbl.cget("text")
    #path = os.path.dirname(os.path.realpath(__file__))
    #path =r'C:\Users\akhil\Desktop\New folder'
    

    excel_names = glob.glob(path + "/*.xlsx")
    
    #read the excel files
    excels = [pd.ExcelFile(name) for name in excel_names]

    #turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

    #delete the header rows
    frames[1:] = [df[row_count:] for df in frames[1:]]
    #frames[1:] = [df[rows_count_int:] for df in frames[1:]]
    #frames[1:] = [df[1:] for df in frames[1:]]

    #concatenate them..
    combined = pd.concat(frames)

    
    destination_dir = destination_lbl.cget("text")
    print(destination_dir)
    new_dir = destination_dir.rsplit('/',1)
    filename = new_dir[1]
    destination_dir = new_dir[0]+'/'
    
    #filename = "Combined.xlsx"
    os.chdir(destination_dir)
    #write the combined file
    combined.to_excel(filename, header=False, index=False)

def change_dropdown(*args):
        header_val = header.get()

btn1 = Button(window, text="Excel Source", command=clicked_source)
btn1.grid(column=1, row=3)
btn2 = Button(window, text="File Destinaton", command=clicked_destination)
btn2.grid(column=1, row=5)
header = StringVar(window)
choices = {1,2,3}
header.set(1)
popupMenu = OptionMenu(window, header, *choices)
popupMenu.grid(row = 7, column =1)
header.trace('w', change_dropdown)

combine = Button(window, text="Combine", command=combine)



combine.grid(column=1, row=10)

window.mainloop()
