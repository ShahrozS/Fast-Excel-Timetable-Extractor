import openpyxl
import re
import string
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
wb = load_workbook('Timetable.xlsx')
n = 0
sheets = wb.sheetnames
wsMonday = wb[sheets[17]]

# rows = wsf.iter_rows(min_row = 3,max_row = 51 , min_col =2 ,max_col =10)
def check(classname,val):
    if classname in val:    
        return True;
    else:
        return False;   


       
                        


def PrintByDay(classname,day,SheetNum):
    print("              ",day)
    print()
    temp = 1
    print("Time ","     ","    Class","             ","Sub")
    print()
    wsDay = wb[sheets[SheetNum]]
    for col in range(2,11):
        for row in range(5,51):
            char = get_column_letter(col)
           
              
            if wsDay[char+str(row)].value!= None and check(classname,wsDay[char+str(row)].value):
            
             print(wsDay[char+'3'].value  , "      ",wsDay["A"+str(row)].value,"     ", wsDay[char+str(row)].value.replace('\n', ' ') ,"       ", end= " ")
             print()
             print()
             temp = 0 

    if temp == 1:
             print("OFF DAY!!")  
    print()                          






classname = input("Enter Class Name(e.g. BAI-3A): ")

# Monday
PrintByDay(classname,"MONDAY",17)
#Tuesday
PrintByDay(classname,"TUESDAY",18)
#wednesday
PrintByDay(classname,"WEDNESDAY",19)
#thursday
PrintByDay(classname,"THURSDAY",20)
#Friday
PrintByDay(classname,"FRIDAY",21)


  
    
# for b,c,d,e,f,g,h,i in rows:

#     if b.value!=None and check(b.value):
#         char = get_column_letter()
#         print(wsf[char+"1"])
#         print(b.value) 
#     elif c.value!=None and check(c.value):
#         print(c.value)
#     elif d.value!=None and check(d.value):
#         print(d.value)
#     elif e.value!=None and check(e.value):
#         print(e.value)
#     elif f.value!=None and check(f.value):
#         print(f.value)
#     elif g.value!=None and check(g.value):
#         print(g.value)
#     elif h.value!=None and check(h.value):
#         print(h.value)
#     elif i.value!=None and check(i.value):
#         print(i.value)
    
    
