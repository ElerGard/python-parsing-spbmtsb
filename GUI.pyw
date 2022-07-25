from glob import glob
from tkinter import *
import pyodbc
import os
import pandas as pd
from datetime import datetime, date
from tkcalendar import DateEntry
from data import *

def autocompleteFirst(event):
    value = e.get()

    if value == "":
        data = unselected_tools
    else:
        data = []
        for item in unselected_tools:
            if value.lower() in item.lower():
                data.append(item)

    lb.delete(0, "end")
    for item in data:
        lb.insert("end", item)
    

def autocompleteSecond(event):
    value = e2.get()

    if value == "":
        data = unselected_basis
    else:
        data = []
        for item in unselected_basis:
            if value.lower() in item.lower():
                data.append(item)

    lb3.delete(0, "end")
    for item in data:
        lb3.insert("end", item)
  
def add_tool(event):
    cs = lb.curselection()[0]
    selected_item = lb.get(0, "end")[cs]
    lb2.insert("end", selected_item)
    lb.delete([cs])
    unselected_tools.remove(selected_item)
    unselected_tools.sort()

def add_basis(event):
    cs = lb3.curselection()[0]
    selected_item = lb3.get(0, "end")[cs]
    lb4.insert("end", selected_item)
    unselected_basis.remove(selected_item)
    unselected_basis.sort()
    
    autocompleteSecond(event)
    
def del_tool(event):
    cs = lb2.curselection()[0]
    
    unselected_tools.append(lb2.get(0, "end")[cs])
    unselected_tools.sort()
    lb2.delete(cs, cs)
    
    autocompleteFirst(event)

def del_basis(event):
    cs = lb4.curselection()[0]
    
    unselected_basis.append(lb4.get(0, "end")[cs])
    unselected_basis.sort()
    lb4.delete(cs, cs)
    
    autocompleteSecond(event)

def reset_selected_tools():
    global unselected_tools
    
    e.delete(0, "end")
    lb.delete(0, "end")
    lb2.delete(0, "end")
    
    for item in all_tools:
        lb.insert("end", item)
        
    unselected_tools = all_tools.copy()

def reset_selected_basis():
    global unselected_basis
    
    e2.delete(0, "end")
    lb3.delete(0, "end")
    lb4.delete(0, "end")
    
    for item in all_basis:
        lb3.insert("end", item)
        
    unselected_basis = all_basis.copy()
    
def write_to_excel(excel_name, df):
    if os.path.exists(excel_name):
        with pd.ExcelWriter(excel_name,mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name="Sheet1",header=None, startrow=writer.sheets["Sheet1"].max_row,index=False)
    else:
        writer = pd.ExcelWriter(excel_name, engine="openpyxl")
        df.to_excel(writer, index=False)
        writer.save()
    
def export():
    global conn
    cursor = conn.cursor()
    isFinded = False
    now = datetime.now()
    date_time = "Экспорт " + now.strftime("%d.%m.%Y_%H.%M")
    lbl3.configure(text="Экспорт данных...")
    for tool in lb2.get(0, "end"):
        for basis in lb4.get(0, "end"):
            cursor.execute(f"select * from Главная where Товар = '{tool}' and БазисПоставки = '{basis}' and СреднЦена <> 0 and Дата between #{date_from.get()}# and #{date_to.get()}#")
            res = cursor.fetchall()
            if (res == []):
                continue
            isFinded = True
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
            df["Дата"] = pd.to_datetime(df["Дата"], format="%d.%m.%Y")
            write_to_excel(f"Export\{date_time}.xlsx", df)
            lbl3.configure(text="Экспорт завершен")
    if not(isFinded):
        lbl3.configure(text="Нет соответствий")
    cursor.close()
    return
    
def update():
    for item in all_tools:
        lb.insert("end", item)
        
    for item in all_basis:
        lb3.insert("end", item)

def updateDB():
    global conn
    directory = "Downloads"
    root2 = Tk()
    scrollbar = Scrollbar(root2)
    scrollbar.pack(side=RIGHT, fill=Y)
    textbox = Text(root2)
    textbox.pack()
    textbox.tag_config("finished", background="green", foreground="white")
    textbox.tag_config("warning", background="red")
    textbox.insert("end", getFilesFromAllPages("https://spimex.com/markets/oil_products/trades/results/?page=page-", directory))
    if (all_tools == []):
        textbox.insert("end", "Файл с ресурсами отсутствует\n", "warning")
    else:
        cursor = conn.cursor()
        for tov in all_tools:
            if (tov == "Газы углеводородные сжиженные"):
                    tov = "СУГ"
            cursor.execute(f"SELECT Товар FROM Главная Where Товар='{tov}';")
            x = cursor.fetchone()
            if not(x):
                textbox.insert("end", f"В БД обновлён ресурс: {tov}\n")
                cursor.execute(f"UPDATE Главная SET Главная.Товар = '{tov}' WHERE InStr(НаименованиеИнструмента,'{tov}')<>0;")
        st = insertDB(directory, conn)
        if (st == "Файл базы данных не найден\n"):
            textbox.insert("end", st, "warning")
        else:
            textbox.insert("end", st)
        cursor.close()
    textbox.insert("end", "Обновление БД завершено", "finished")
    
    textbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=textbox.yview)
    
def popup(text):
    root = Tk()
    root.title("Warning")
    root.geometry("300x150")
    lbl = Label(root, text=text)
    lbl.place(x=25, y=50)
    root.mainloop()

try:
    conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\db.accdb;")
    cursor = conn.cursor()
    cursor.execute("select distinct БазисПоставки from Главная where БазисПоставки <> ''")
    all_basis = [item[0] for item in cursor.fetchall()]
    cursor.close()
    all_tools  = parseResourcesFromFile()
    if (all_tools == []):
        popup("Файл с ресурсами пуст или отстутствует")
    
    all_tools = sorted(all_tools)
    all_basis = sorted(all_basis)

    unselected_tools = all_tools.copy()

    unselected_basis = all_basis.copy()
    
    if not(os.path.isdir("Export")): 
            os.makedirs("Export")

    root = Tk()
    root.title("SPB")
    root.geometry("575x600")

    button3 = Button(root, text="Обновить базу данных", command=updateDB)
    button3.grid(row=9, column=4)
    
    # button3 = Button(root, text="Добавить ресурс", command=updateDB)
    # button3.grid(row=11, column=4)

    scrollbar = Scrollbar(root)
    scrollbar.grid(row=2, column=1, sticky="ns")

    scrollbar2 = Scrollbar(root)
    scrollbar2.grid(row=2, column=5, sticky="ns")

    scrollbar3 = Scrollbar(root)
    scrollbar3.grid(row=5, column=1, sticky="ns")

    scrollbar4 = Scrollbar(root)
    scrollbar4.grid(row=5, column=5, sticky="ns")

    lbl = Label(root, text="Название\nресурса")
    lbl.grid(column=0, row=0)

    e = Entry(root, width=35)
    e.grid(column=0, row=1)
    e.bind("<KeyRelease>", autocompleteFirst)
    
    lb = Listbox(root, height=10, width=35, yscrollcommand=scrollbar.set)
    lb.grid(column=0, row=2)
    lb.bind("<Double-1>", add_tool)

    button = Button(root, text="Сброс", command=reset_selected_tools)
    button.grid(column=2, row=2, padx=33)

    lb2 = Listbox(root, height=10, width=35, yscrollcommand=scrollbar2.set)
    lb2.grid(column=4, row=2)
    lb2.bind("<Double-1>", del_tool)

    #-------------------
    lbl2 = Label(root, text="Базис\nпоставки")
    lbl2.grid(column=0, row=3)

    e2 = Entry(root, width=35)
    e2.grid(column=0, row=4)
    e2.bind("<KeyRelease>", autocompleteSecond)
    
    lb3 = Listbox(root, height=10, width=35, yscrollcommand=scrollbar3.set)
    lb3.grid(column=0, row=5)
    lb3.bind("<Double-1>", add_basis)

    button2 = Button(root, text="Сброс", command=reset_selected_basis)
    button2.grid(column=2, row=5)

    lb4 = Listbox(root, height=10, width=35,  yscrollcommand=scrollbar4.set)
    lb4.grid(column=4, row=5)
    lb4.bind("<Double-1>", del_basis)

    lbl4 = Label(root, text="Дата от")
    lbl4.grid(column=0, row=8)

    date_from = DateEntry(root, selectmode="day", date_pattern="mm/dd/y", mindate=date(2015, 10, 19), maxdate=date.today(), year=2015,month=10,day=19)
    date_from.grid(column=0, row=9)

    lbl5 = Label(root, text="Дата до")
    lbl5.grid(column=0, row=10)

    date_to = DateEntry(root, selectmode="day", date_pattern="mm/dd/y", mindate=date(2015, 10, 19), maxdate=date.today())
    date_to.grid(column=0, row=11)

    button3 = Button(root, text="Экспорт", command=export)
    button3.grid(column=2)

    lbl3 = Label(root, text="")
    lbl3.grid(column=2)

    scrollbar.config(command=lb.yview)
    scrollbar2.config(command=lb2.yview)
    scrollbar3.config(command=lb3.yview)
    scrollbar4.config(command=lb4.yview)



    update()
    
    root.resizable(0, 0)
    root.mainloop()
    
    conn.close()
    
except pyodbc.Error as e:
    print("DB not found")