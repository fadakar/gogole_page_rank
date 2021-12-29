import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import xlrd
import xlwt
from xlutils.copy import copy
import os
import sys
from termcolor import colored
from pyfiglet import *

from selenium.webdriver.chrome.options import Options
from key_generator import check_key


def get_root_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    elif __file__:
        return os.path.dirname(__file__)


class Row:
    def __init__(self, keyword, target, page, value):
        self.keyword = keyword
        self.target = target
        self.page = page
        self.value = value


class GoogleRank:
    __delay = 3
    driver = None
    data = None
    current_page = 1
    max_page_search = 3
    result = []
    __sheet = None
    __wb = None
    __license = ''
    __license_file_name = 'key.license'

    def __init__(self):
        self.load_license()
        self.check_license()

    def load_excel(self):
        ROOT_DIR = get_root_dir()
        excel_path = os.path.join(ROOT_DIR, 'keywords.xls')
        self.wb = xlrd.open_workbook(excel_path)
        self.__sheet = self.wb.sheet_by_index(0)
        w = self.__sheet.col_values(0, 0)
        t = self.__sheet.col_values(1, 0)
        try:
            pages = self.__sheet.col_values(2, 0)
            counters = self.__sheet.col_values(3, 0)
        except:
            pages = []
            counters = []
        finally:
            self.data = {'keywords': w, 'targets': t, 'pages': pages, 'counters': counters}

    def load_license(self):
        ROOT_DIR = get_root_dir()
        license_file_path = os.path.join(ROOT_DIR, self.__license_file_name)
        file = open(license_file_path, 'r')
        self.__license = file.read()
        file.close()

    def check_license(self):
        if not check_key(self.__license):
            raise ''

    def load_driver(self):
        ROOT_DIR = get_root_dir()
        driver_path = os.path.join(ROOT_DIR, 'chromedriver.exe')
        s = Service(driver_path)
        chrome_options = Options()
        chrome_options.add_argument('--log-level=3')
        self.driver = webdriver.Chrome(service=s, options=chrome_options)

    def search_in_google(self, keyword):
        URL = "https://www.google.com/search?q="
        self.driver.get(URL + keyword)
        time.sleep(self.__delay)

    def go_to_next_page(self):
        self.current_page += 1
        link = self.driver.find_element(By.CSS_SELECTOR, '[aria-label="Page ' + str(self.current_page) + '"]')
        link.click()
        time.sleep(self.__delay)

    def find_in_result(self, keyword, target, counter=0):
        container = self.driver.find_element(By.ID, 'rso')
        if not container:
            raise 'Captcha Coming...'
        links = container.find_elements(By.CSS_SELECTOR, '.g > div:not(.g)')
        found = False
        for link in links:
            counter += 1
            if target in link.text:
                found = True
                break
        if found:
            row = Row(keyword, target, self.current_page, counter)
            self.result.append(row)
            print("[" + keyword + "][" + target + "] = " + "[ " + str(self.current_page) + " ][ " + str(counter) + " ]")
        else:
            if self.current_page < self.max_page_search:
                self.go_to_next_page()
                self.find_in_result(keyword, target, counter)
            else:
                row = Row(keyword, target, '', 'Not Found')
                self.result.append(row)
                print("[" + keyword + "][" + target + "] = Not Found")

    def write_result(self):
        i = 0
        wb = copy(self.wb)
        sheet = wb.get_sheet(0)
        for row in self.result:
            sheet.write(i, 0, row.keyword)
            sheet.write(i, 1, row.target)
            sheet.write(i, 2, row.page)
            sheet.write(i, 3, row.value)
            i += 1
        wb.save('keywords.xls')

    def run(self):
        try:
            self.load_driver()
            self.load_excel()
            keywords = self.data['keywords']
            targets = self.data['targets']
            i = -1
            for keyword in keywords:
                i += 1
                if len(self.data['counters']) > i and self.data['counters'][i] != '':
                    self.result.append(Row(keyword, targets[i], self.data['pages'][i], self.data['counters'][i]))
                    continue
                self.current_page = 1
                self.search_in_google(keyword)
                self.find_in_result(keyword, targets[i])
        except Exception as ex:
            print(ex)
            pass
        finally:
            self.write_result()
            self.driver.quit()
            # print('[Error][500]: please contact Administrator')


if __name__ == '__main__':
    # try:
    ranking = GoogleRank()
    f = Figlet()
    print(colored(f.renderText('Page Rank'), 'red'))
    print(colored('\tCreated By Gholamreza Fadakar', 'red'))
    print()
    print()
    print()
    ranking.run()
    # except Exception as ex:
    #     print(ex)
    #     pass
