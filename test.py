import re
import requests
from bs4 import BeautifulSoup
import random
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
def twitter_List():
    #Excel_Read
    file_Location = ("TWITTER.xls")
    read_book = open_workbook(file_Location)
    copy_book = copy(read_book)
    write_book = xlrd.open_workbook(file_Location)
    sheet = write_book.sheet_by_index(0)
    sheet.cell_value(0, 0)
    for data in range(sheet.nrows):
        text = sheet.cell_value(data, 0) #pass the Excel list one by one to the text
        url = 'https://google.com/search?q=' + text
        A = ("Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36",
        )
        Agent = A[random.randrange(len(A))]
        headers = {'user-agent': Agent}
        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, 'lxml')
        ret_val = soup.find_all('h3')
        for Twit_list in range(len(ret_val)):
            twitter_str = " "
            k = ret_val[Twit_list].text
            pattern = re.compile(r"@\w+")
            final_val = pattern.findall(k)
            list_Twit = twitter_str.join(final_val)
            if len(list_Twit) == 0:
                list_Twit = "N/A"
            else:
                list_Twit = list_Twit
                Write_book_val = copy_book.get_sheet(0)
                Write_book_val.write(data, 1, list_Twit)
                copy_book.save(file_Location)
                print(list_Twit)
                break
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    twitter_List()