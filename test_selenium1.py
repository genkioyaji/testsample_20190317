#20190307
# https://note.nkmk.me/python-openpyxl-usage/
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl 

class EX_LIST:
    def __init__(self,file_1):
        self.file_1 = file_1
    def get_value_list(self,t_2d):
        return([[cell.value for cell in row] for row in t_2d])


    def get_list_2d(self,sheet, start_row, end_row, start_col, end_col):
        return self.get_value_list(sheet.iter_rows(min_row=start_row,
                                          max_row=end_row,
                                          min_col=start_col,
                                          max_col=end_col))
    
    def make_list(self):
        wb = openpyxl.load_workbook(self.file_1)
        sheet_names = wb.get_sheet_names()
        sheet = wb.get_sheet_by_name(sheet_names[0])           
        l_2d = self.get_list_2d(sheet, 2, 4, 1, 9)
        return l_2d
#        print(l_2d) 
class FORM_SEND:
    def __init__(self,item,text1):
        self.item = item
        self.text1 =text1        
    def formsend(self):
        try:
            driver.find_element_by_name(self.item).send_keys(self.text1)
        except:
            print("error ")
            time.sleep(3)
            driver.quit()

class Date_form(FORM_SEND):
    def __init__(self,item,text):
        super().__init__(item,text)
        
if __name__=='__main__':
    test1 = EX_LIST('selenium_test.xlsx')
    result =test1.make_list()
    print(result)
        
    url_1= 'http://example.selenium.jp/reserveApp/'
    driver = webdriver.Chrome()
    driver.get(url_1)

    form_ex = FORM_SEND('gnam',result[0][8])
    form_ex.formsend()
#    item_names = [reserve_y, ]
    time.sleep(5)
    driver.quit()   

