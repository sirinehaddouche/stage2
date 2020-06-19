import tkinter as tk  
from functools import partial  
import random
import xlrd
      
       
def call_result(label_result):  
  loc = "coupon.xlsx"
  wb = xlrd.open_workbook(loc)
  Sheet = wb.sheet_names()
  sheet = wb.sheet_by_index(0)
  #print(sheet.nrows)
  v=(random.randint(0, 5000))
  result=sheet.cell_value(v, 0) 
  label_result.config(text="Voici ton lien :  " +result)
  return
       
root = tk.Tk()  
root.geometry('50000x200')  
      
root.title('Code')   
labelResult = tk.Label(root)  
labelResult.grid(row=7, column=2)  
call_result = partial(call_result, labelResult)  
buttonCal = tk.Button(root, text="Obtenir un nouveau lien", command=call_result).grid(row=3, column=0)  
root.mainloop()
