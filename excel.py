import sys
import csv
import glob
import pandas as pd
import os

rows_count=int(input("Enter the number of header rows"))
#rows_count_int = int(rows_count)

#filenames
path = os.path.dirname(os.path.realpath(__file__))
#path =r'C:\Users\akhil\Desktop\New folder'


excel_names = glob.glob(path + "/*.xlsx")
#excel_names = ["xlsx1.xlsx", "xlsx2.xlsx", "xlsx3.xlsx"]

#read the excel files
excels = [pd.ExcelFile(name) for name in excel_names]

#turn them into dataframes
frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

#delete the header rows
frames[1:] = [df[rows_count:] for df in frames[1:]]
#frames[1:] = [df[rows_count_int:] for df in frames[1:]]
#frames[1:] = [df[1:] for df in frames[1:]]

#concatenate them..
combined = pd.concat(frames)

#write the combined file
combined.to_excel("combined.xlsx", header=False, index=False)
