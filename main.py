from time import sleep
from requests_html import HTMLSession
from urllib import request
import re
import os

session = HTMLSession()

def getFilesFromPage(r, folder_to_save_files):
    links = r.html.find(".page-content__tabs__block:first-child .accordeon-inner__item-title")
    for link in links:
        file_link = "https://spimex.com" + link.attrs['href']
        filepath = folder_to_save_files + '/' + re.search("\w*[.]\w+", link.attrs["href"])[0]
        if not(os.path.isdir(folder_to_save_files)): 
            os.makedirs(folder_to_save_files)
            
        request.urlretrieve(file_link, filepath)
    return
    
def getFilesFromAllPages(url, folder_to_save_files):
    page_number = 1
    while (True):
        page_url = url + str(page_number)
        r = session.get(page_url)
        print("url: " + page_url)
        getFilesFromPage(r, folder_to_save_files)
        page_number += 1
        sleep(20)
        if (r.html.find(".bx-pag-next > a") == []):
            break
        
    return

def main():
    getFilesFromAllPages("https://spimex.com/markets/oil_products/trades/results/?page=page-", "Downloads2")
    
main()
