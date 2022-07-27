from tkinter import *
import pyodbc
import os
import pandas as pd
from datetime import datetime, date
from tkcalendar import DateEntry
from data import *

def autocomplete(text, unselected, list) -> None:
    value = text.get()

    if value == "":
        data = unselected
    else:
        data = []
        for item in unselected:
            if value.lower() in item.lower():
                data.append(item)

    list.delete(0, "end")
    for item in data:
        list.insert("end", item)

def add(list1, list2, unselected) -> None:
    cs = list1.curselection()[0]
    selected_item = list1.get(0, "end")[cs]
    list2.insert("end", selected_item)
    list1.delete([cs])
    unselected.remove(selected_item)
    unselected.sort()
    
def delete(list1, list2, unselected, text, func) -> None:
    cs = list2.curselection()[0]
    
    unselected.append(list2.get(0, "end")[cs])
    unselected.sort()
    list2.delete(cs, cs)
    
    func(text, unselected, list1)

def reset_lists(text, list1, list2, all, flag):
    text.delete(0, "end")
    list1.delete(0, "end")
    list2.delete(0, "end")
    
    for item in all:
        list1.insert("end", item)
    if (flag == 1):
        global unselected_tools
        unselected_tools = all.copy()
    if (flag == 2):
        global unselected_basis
        unselected_basis = all.copy()
    
def write_to_excel(excel_name, df) -> None:
    if os.path.exists(excel_name):
        with pd.ExcelWriter(excel_name,mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name="Sheet1",header=None, startrow=writer.sheets["Sheet1"].max_row,index=False)
    else:
        writer = pd.ExcelWriter(excel_name, engine="openpyxl")
        df.to_excel(writer, index=False)
        writer.save()
    
def export() -> None:
    global conn
    cursor = conn.cursor()
    isFinded = False
    now = datetime.now()
    date_time = "Export " + now.strftime("%d.%m.%Y_%H.%M")
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
    
def update() -> None:
    for item in all_tools:
        lb.insert("end", item)
        
    for item in all_basis:
        lb3.insert("end", item)

def updateRecources() -> None:
    global conn
    global all_tools
    all_tools  = parseResourcesFromFile()
    
    root2 = Tk()
    scrollbar = Scrollbar(root2)
    scrollbar.pack(side=RIGHT, fill=Y)
    textbox = Text(root2)
    textbox.pack()
    textbox.tag_config("finished", background="green", foreground="white")
    textbox.tag_config("warning", background="red")
    if (all_tools == []):
        textbox.insert("end", "Файл с ресурсами отсутствует или пуст\n", "warning")
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
        cursor.close()
    textbox.insert("end", "Ресурсы обновлены", "finished")
    
    textbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=textbox.yview)

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
    st = insertDB(directory, conn)
    if (st == "Файл базы данных не найден\n"):
        textbox.insert("end", st, "warning")
    else:
        textbox.insert("end", st)
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
    
    button4 = Button(root, text="Обновить ресурсы", command=updateRecources)
    button4.grid(row=11, column=4)

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
    
    lb2 = Listbox(root, height=10, width=35, yscrollcommand=scrollbar2.set)
    lb2.grid(column=4, row=2)
    lb2.bind("<Double-1>", lambda event: delete(lb, lb2, unselected_tools, e, autocomplete))
    
    lb = Listbox(root, height=10, width=35, yscrollcommand=scrollbar.set)
    lb.grid(column=0, row=2)
    lb.bind("<Double-1>", lambda event: add(lb, lb2, unselected_tools))
    
    e = Entry(root, width=35)
    e.grid(column=0, row=1)
    e.bind("<KeyRelease>", lambda event: autocomplete(e, unselected_tools, lb))

    button = Button(root, text="Сброс", command=lambda : reset_lists(e, lb, lb2, all_tools, 1))
    button.grid(column=2, row=2, padx=33)

    #-------------------
    lbl2 = Label(root, text="Базис\nпоставки")
    lbl2.grid(column=0, row=3)

    lb4 = Listbox(root, height=10, width=35,  yscrollcommand=scrollbar4.set)
    lb4.grid(column=4, row=5)
    lb4.bind("<Double-1>", lambda event: delete(lb3, lb4, unselected_basis, e2, autocomplete))

    lb3 = Listbox(root, height=10, width=35, yscrollcommand=scrollbar3.set)
    lb3.grid(column=0, row=5)
    lb3.bind("<Double-1>", lambda event: add(lb3, lb4, unselected_basis))

    e2 = Entry(root, width=35)
    e2.grid(column=0, row=4)
    e2.bind("<KeyRelease>", lambda event: autocomplete(e2, unselected_basis, lb3))

    button2 = Button(root, text="Сброс", command=lambda : reset_lists(e2, lb3, lb4, all_basis, 2))
    button2.grid(column=2, row=5)

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