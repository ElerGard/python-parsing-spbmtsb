from cProfile import label
from tkinter import *
import pyodbc
import os
import pandas as pd
from datetime import datetime

conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\db.accdb;")
cursor = conn.cursor()
cursor.execute("select distinct БазисПоставки from Главная where БазисПоставки <> ''")
all_basis = [item[0] for item in cursor.fetchall()]

cursor.execute("select distinct Товар from Главная where Товар <> ''")
all_tools  = [item[0] for item in cursor.fetchall()]

all_tools = sorted(all_tools)
all_basis = sorted(all_basis)

tmp_tools = all_tools.copy()
added_tools = []

tmp_basis = all_basis.copy()
added_basis = []

def checkkey(event):
       
    value = event.widget.get()
      
    # if value == "":
    #     tmp_tools = all_tools.copy()
    # else:
    #     tmp_tools = []
    #     for item in all_tools:
    #         if value.lower() in item.lower():
    #             tmp_tools.append(item)                
    # update()
   
def update():
    lb.delete(0, "end")
    lb2.delete(0, "end")
    lb3.delete(0, "end")
    lb4.delete(0, "end")
    
    for item in tmp_tools:
        lb.insert("end", item)
        
    for item in added_tools:
        lb2.insert("end", item)
        
    for item in tmp_basis:
        lb3.insert("end", item)
        
    for item in added_basis:
        lb4.insert("end", item)
        
    lbl3.configure(text="")
  
def add_tool(event):
    cs = lb.curselection()
    
    added_tools.append(tmp_tools[cs[0]])
    added_tools.sort()
    tmp_tools.pop(cs[0])
    tmp_tools
    
    update()

def add_basis(event):
    cs = lb3.curselection()
    
    added_basis.append(tmp_basis[cs[0]])
    added_basis.sort()
    tmp_basis.pop(cs[0])
    tmp_basis
    
    update()
    
def del_tool(event):
    cs = lb2.curselection()
    
    tmp_tools.append(added_tools[cs[0]])
    tmp_tools.sort()
    added_tools.pop(cs[0])
    
    update()

def del_basis(event):
    cs = lb4.curselection()
    
    tmp_basis.append(added_basis[cs[0]])
    tmp_basis.sort()
    added_basis.pop(cs[0])
    
    update()

def reset_selected_tools():
    # e.delete(0, END)
    # tmp_tools = all_tools.copy()
    # added_tools = []
    update()

def reset_selected_basis():
    # e.delete(0, END)
    # tmp_basis = all_basis.copy()
    # added_basis = []
    update()
    
def write_to_excel(excel_name, df):
    if os.path.exists(excel_name):
        return
    else:
        writer = pd.ExcelWriter(excel_name, engine="openpyxl")
        df.to_excel(writer, index=False)
        writer.save()
    
def export():
    now = datetime.now()
    date_time = "Экспорт " + now.strftime("%d.%m.%Y_%H.%M")
    for tool in added_tools:
        for basis in added_basis:
            cursor.execute(f"select * from Главная where Товар = '{tool}' and БазисПоставки = '{basis}' and СреднЦена <> 0 and Дата between #{date_from.get()}# and #{date_to.get()}#")
            res = cursor.fetchall()
            if (res == []):
                lbl3.configure(text="Нет соответствий")
                continue
            df = pd.DataFrame({"Код инструмента": [item[1] for item in res],
                            "Наименование инструмента": [item[2] for item in res],
                            "Сокращенное название": [item[16] for item in res],
                            "Базис поставки": [item[3] for item in res],
                            "Объем Договоров в тоннах": [item[4] for item in res],
                            "Объем Договоров, руб.": [item[5] for item in res],
                            "Средневзвешенная": [item[9] for item in res],
                            "Лучшее предложение": [item[12] for item in res],
                            "Лучший спрос": [item[13] for item in res],
                            "Количество Договоров, шт.": [item[14] for item in res],
                            "Дата": [item[15] for item in res]})
            df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y')
            
            lbl3.configure(text="Экспорт файла...")
            write_to_excel(date_time + ".xlsx", df)
            lbl3.configure(text="Экспорт завершен")
    return
    
root = Tk()
root.title("SPB")
root.geometry("550x700")

lbl = Label(root, text="Название\nинструмента")
lbl.grid(column=0, row=0)

e = Entry(root, width=35)
e.grid(column=0, row=2)
e.bind("<KeyRelease>", checkkey)
  
lb = Listbox(root, height=10, width=35)
lb.grid(column=0, row=3)
lb.bind("<Double-1>", add_tool)


button = Button(root, text="Сброс", state = DISABLED, command=reset_selected_tools)
button.grid(column=1, row=3)

lb2 = Listbox(root, height=10, width=35)
lb2.grid(column=3, row=3)
lb2.bind("<Double-1>", del_tool)

#-------------------
lbl2 = Label(root, text="Базис\nпоставки")
lbl2.grid(column=0, row=4)

e2 = Entry(root, width=35)
e2.grid(column=0, row=5)
e2.bind("<KeyRelease>", checkkey)
  
lb3 = Listbox(root, height=10, width=35)
lb3.grid(column=0, row=6)
lb3.bind("<Double-1>", add_basis)

button2 = Button(root, text="Сброс", state = DISABLED, command=reset_selected_basis)
button2.grid(column=1, row=6)

lb4 = Listbox(root, height=10, width=35)
lb4.grid(column=3, row=6)
lb4.bind("<Double-1>", del_basis)



lbl4 = Label(root, text="Дата от")
lbl4.grid(column=0, row=8)

date_from = Entry(root, width=35)
date_from.grid(column=0, row=9)

lbl5 = Label(root, text="Дата до")
lbl5.grid(column=0, row=10)

date_to = Entry(root, width=35)
date_to.grid(column=0, row=11)

lbl6 = Label(root, text="Формат даты - месяц/день/год")
lbl6.grid(column=0, row=12)

button3 = Button(root, text="Экспорт", command=export)
button3.grid(column=1)

lbl3 = Label(root, text="")
lbl3.grid(column=1)

update()

root.mainloop()