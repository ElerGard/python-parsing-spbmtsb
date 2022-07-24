from time import sleep
from requests_html import HTMLSession
from urllib import request
import re
import os
import pyodbc
import pandas as pd

def getFilesFromPage(r, folder_to_save_files):
    links = r.html.find(".page-content__tabs__block:first-child .accordeon-inner__item-title")
    for link in links:
        file_link = "https://spimex.com" + link.attrs['href']
        path_to_file = os.path.join(folder_to_save_files, re.search("\w*[.]\w+", link.attrs["href"])[0])
        
        if not(os.path.isdir(folder_to_save_files)): 
            os.makedirs(folder_to_save_files)
            
        if os.path.exists(path_to_file):
            return False
        request.urlretrieve(file_link, path_to_file)
        print(re.search("\w*[.]\w+", link.attrs["href"])[0] + " dowloaded")
        sleep(3)
    return True
    
def getFilesFromAllPages(url, folder_to_save_files):
    session = HTMLSession()
    page_number = 1
    while (True):
        page_url = url + str(page_number)
        r = session.get(page_url)
        print("url: " + page_url)
        if not(getFilesFromPage(r, folder_to_save_files)):
            print("Files downloaded")
            return
        page_number += 1
        sleep(20)
        if (r.html.find(".bx-pag-next > a") == []):
            break
    return

def insertDB(directory, db_filename):
    tmp = [
        "Газы углеводородные сжиженные", # -> СУГ
        "АИ-100-К5",
        "АИ-92-К5",
        "АИ-92-К4",
        "АИ-95-К5",
        "АИ-98-К5",
        "АИ-80-К5",
        "ДТ-А-К5",
        "ДТ-З-К5",
        "ДТ-Е-К5",
        "ДТ-Л-К5",
        "Керосин",
        "М100 0,5",
        "М100 3,0",
        "М100 1,5",
        "М100 3,5",
        "Масло",
        "TKM-16",
        "АИ-98-К5-Евро",
        "Мазут"
    ]   
    try:
        conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\" + db_filename + ";")
        cursor = conn.cursor()
        
        for filename in reversed(os.listdir(directory)):
            path_to_file = os.path.join(directory, filename)
            xls = pd.ExcelFile(path_to_file)
            sheetX = xls.parse(0) 
            
            if sheetX.empty:
                print('File is empty!')
                continue
            
            old_date = re.search("\d{8}", filename)[0]
            date = f"{old_date[:4]}.{old_date[4:6]}.{old_date[6:]}"
            cursor.execute(f"select Дата from Главная where Дата=#{old_date[4:6]}/{old_date[6:]}/{old_date[:4]}#")
            res = cursor.fetchall()
            if (res != []):
                print(f"{filename} already in DB")
                return
            sheetX = sheetX.dropna(axis=1,how='all')
            
            row_start = sheetX.loc[sheetX[sheetX.columns[0]] == "Код\nИнструмента"]._stat_axis[0]
            row_end = sheetX.loc[sheetX[sheetX.columns[0]] == "Итого:"]._stat_axis[0]
            row_tmp = row_start + 2
            
            while (row_tmp != row_end):
                sheetX.loc[row_tmp]
                for i in range(len(sheetX.loc[row_tmp]._values)):
                    if (sheetX.loc[row_tmp]._values[i] == "-"):
                        sheetX.loc[row_tmp]._values[i] = "0"
                
                tovar = ""
                for tov in tmp:
                    x = re.search(tov, sheetX.loc[row_tmp]._values[1], re.IGNORECASE)
                    if (x):
                        tovar = x[0]
                        if (tovar == tmp[0]):
                            tovar = "СУГ"
                        break
                sheetX.loc[row_tmp]._values[6] = str(sheetX.loc[row_tmp]._values[6]).replace('.',',')
                # TODO cursor.executemany(-, -)
                cursor.execute("insert into Главная (КодИнструмента, НаименованиеИнструмента, БазисПоставки, ОбъемДоговоровЕИ, ОбъемДоговоровРуб, ИзмРынРуб, ИзмРынПроц, МинЦена, СреднЦена, МаксЦена, РынЦена, ЛучшПредложение, ЛучшСпрос, КоличествоДоговоров, Дата, Товар) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                            (sheetX.loc[row_tmp]._values[0], sheetX.loc[row_tmp]._values[1], sheetX.loc[row_tmp]._values[2], sheetX.loc[row_tmp]._values[3],
                                sheetX.loc[row_tmp]._values[4], sheetX.loc[row_tmp]._values[5], sheetX.loc[row_tmp]._values[6], sheetX.loc[row_tmp]._values[7],
                                sheetX.loc[row_tmp]._values[8], sheetX.loc[row_tmp]._values[9], sheetX.loc[row_tmp]._values[10], sheetX.loc[row_tmp]._values[11], 
                                sheetX.loc[row_tmp]._values[12], sheetX.loc[row_tmp]._values[13], date, tovar))
                cursor.commit()
                row_tmp += 1
                
            # print(filename + " type = " + str(typeOfTable(sheetX)))
            # print(xls.loc[xls[xls.columns[0]] == one[0][0]])
            print(f"{filename} inserted in database")
        cursor.close()
        conn.close() 
    except pyodbc.Error as e:
        print("DB not found")
        
def main():
    directory = "Downloads"
    getFilesFromAllPages("https://spimex.com/markets/oil_products/trades/results/?page=page-", directory)
    insertDB(directory, "db.accdb")
    
main()