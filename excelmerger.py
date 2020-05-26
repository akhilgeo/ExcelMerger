import sys
import csv
import glob
import pandas as pd
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

#GUI properties
window = Tk()
window.title("Merge Excel")
window.geometry('350x200')
window.configure(background='white')

ttk.Style().configure("TButton", padding=6, relief="flat", background="white", foreground="#000")
#Header 

heading = Label(window, text="Merge Excel", font=("Arial Bold", 15), anchor='w',background='white')
heading.grid(column=0, row=0)

#Source directory Label 
source_lbl = Label(window, text="Excel Source Directory")
source_lbl.grid(column=1, row=3, sticky=W)

#destination directory Label
destination_lbl = Label(window, text="File Destination", anchor='w')
destination_lbl.grid(column=1, row=5, sticky=W)

 
def clicked_destination():
 
    window.filename =  filedialog.asksaveasfilename(initialdir = "/",title = "Select file",filetypes = (("Excel file","*.xlsx"),("all files","*.*")))
    destination_lbl.configure(text= window.filename+".xlsx")

def clicked_source():

 
    window.directory = filedialog.askdirectory()
    source_lbl.configure(text= window.directory)

def combine(*args):
    try:
   
        #filenames
        path = source_lbl.cget("text")
        excel_names = glob.glob(path + "/*.xlsx")
        
        #read the excel files
        excels = [pd.ExcelFile(name) for name in excel_names]

        #turn them into dataframes
        frames = [pd.read_excel(excel_name, ) for excel_name in excel_names]

        combined = pd.concat(frames, sort=False)

        
        destination_dir = destination_lbl.cget("text")
        print(destination_dir)
        new_dir = destination_dir.rsplit('/',1)
        filename = new_dir[1]
        destination_dir = new_dir[0]+'/'
        
        #filename = "Combined.xlsx"
        os.chdir(destination_dir)
        #write the combined file
        combined.to_excel(filename, index=False)
        messagebox.showinfo("Successful")
    except IOError:
        messagebox.showerror("Error", "Cannot Read File - Please close the source files")
    except NameError as name:
        messagebox.showerror("Error","Enter Valida Source or Destination")
    except:
        messagebox.showerror("Error","Something went wrong, Try again")
      
btn1 = ttk.Button(window, text="Select Source        ", command=clicked_source)
btn1.grid(column=0, row=3, columnspan=4, sticky=W)
btn2 = ttk.Button(window, text="Select Destination", command=clicked_destination)
btn2.grid(column=0, row=5, sticky=W, columnspan=4)

combine = ttk.Button(window, text="Combine", command=combine)

combine.grid(column=1, row=12, sticky=W)

window.mainloop()
