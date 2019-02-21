from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time
import xlsxwriter
import xlrd
import queue
import threading

class conChrome:
    chrome_max = 1
    interval = 0.00001
    timeout = 240
    chrome_options = Options()
    service_args = ['--ignore-ssl-errors=true','--ssl-protocol=TLSv1','--load-images=no','--disk-cache=yes']

    def __init__(self):
        self.q_chrome = queue.Queue()
        self.chrome_options.add_argument("--headless")
        self.q_output = queue.Queue()
        self.q_error = queue.Queue()

    def getPage(self,dict_input):
        browser = self.q_chrome.get()
        url = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vCI_CorpManager/stockid/' + dict_input['com_code'] + '.phtml'
        seniorExecutive_name = dict_input['seniorExecutive_name']
        com_code = dict_input['com_code']
        try:
            openCEOPage(url,browser)
            current_handle = browser.current_window_handle
            handles = browser.window_handles
            new_handle = None
            if jumpIntoResume(current_handle,handles,new_handle,browser,seniorExecutive_name):
                dict_return = recordCEOInfo(browser,seniorExecutive_name,com_code)
                if dict_return['found']:
                    dict_return['com_code'] = com_code
                    dict_return['seniorExecutive_name'] = seniorExecutive_name
                    self.q_output.put(dict_return)
                else:
                    self.q_error.put(dict_input)
            else:
                self.q_error.put(dict_input)
        except:
            self.q_error.put(dict_input)
        self.q_chrome.put(browser)

    def writeLogAndData(self):
        workbook = xlsxwriter.Workbook('output.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': 1})
        worksheet.write('A1', 'Code', bold)
        worksheet.write('B1', 'CEO_name', bold)
        worksheet.write('C1', '任职日期', bold)
        worksheet.write('D1', '离职日期', bold)
        row = 1
        col = 0
        while not self.q_output.empty():
            dict_output = self.q_output.get()
            worksheet.write_string(row, col, dict_output['com_code'])
            worksheet.write_string(row, col + 1, dict_output['seniorExecutive_name'])
            worksheet.write_string(row, col + 2, dict_output['start_time'])
            worksheet.write_string(row, col + 3, dict_output['end_time'])
            row += 1
        workbook.close()
        f = open('./log.txt','w+')
        while not self.q_error.empty():
            dict_error = self.q_error.get()
            err = dict_error['com_code'] + ' ' + dict_error['seniorExecutive_name'] + ' message not found or netWork err\r'
            f.writelines(err)
        f.close()

    def open_chrome(self):
        def open_threading():
            browser = webdriver.Chrome(service_args=conChrome.service_args,chrome_options=conChrome.chrome_options)
            browser.implicitly_wait(conChrome.timeout)
            browser.set_page_load_timeout(conChrome.timeout)
            self.q_chrome.put(browser)
        
        th = []
        for i in range(conChrome.chrome_max):
            t = threading.Thread(target=open_threading)
            th.append(t)
        for i in th:
            i.start()
            time.sleep(conChrome.interval)
        for i in th:
            i.join()
    
    def close_chrome(self):
        th = []
        def close_threading():
            browser = self.q_chrome.get()
            browser.quit()
        for i in range(self.q_chrome.qsize()):
            t = threading.Thread(target=close_threading)
            th.append(t)
        for i in th:
            i.start()
        for i in th:
            i.join()

def switch_handle(handles,current_handle):
    for handle in handles:
        if handle != current_handle:
            return handle

def switch_page(current_handle,handles,new_handle,browser):
    current_handle = browser.current_window_handle
    handles = browser.window_handles
    new_handle = switch_handle(handles,current_handle)
    browser.close()
    browser.switch_to.window(new_handle)
    return None

def jumpIntoResume(current_handle,handles,new_handle,browser,seniorExecutive_name):
    hreflists = browser.find_elements_by_tag_name('a')
    found = False
    for href in hreflists:
        if href.text == seniorExecutive_name:
            found = True
            href.click()
            time.sleep(0.5)
            break
    if not found:
        return False
    else:
        switch_page(current_handle,handles,new_handle,browser)
        return True

def recordCEOInfo(browser, seniorExecutive_name, com_code):
    table = browser.find_element_by_xpath("//table[@id='Table3']")
    dict_return = {}
    trs = table.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
    for tr in trs:
        tds = tr.find_elements_by_tag_name('td')
        if tds[1].find_element_by_tag_name('div').text == '总裁':
            start_time = tds[2].find_element_by_tag_name('div').text
            end_time = tds[3].find_element_by_tag_name('div').text
            dict_return['start_time'] = start_time
            dict_return['end_time'] = end_time
            dict_return['found'] = True
            return dict_return
    dict_return['found'] = False
    return dict_return

def openCEOPage(url,browser):
    print(url)
    try:
        browser.get(url)
    except TimeoutException:
        print('stop load')
        browser.execute_script('window.stop ? window.stop() : document.execC    ommand("Stop");')
        print('exit try')

if __name__ == "__main__":
    row = 1
    col = 0
    input = []
    inputbook = xlrd.open_workbook('./input.xlsx')
    sheet = inputbook.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i == 0:
            continue
        com_code = sheet.cell(row, col).value
        seniorExecutive_name = sheet.cell(row, col + 1).value
        dict_input = {'com_code': com_code, 'seniorExecutive_name': seniorExecutive_name}
        input.append(dict_input)
        row += 1
    inputbook.release_resources()
    cur = conChrome()
    conChrome.chrome_max = 3
    cur.open_chrome()
    print ('chrome num is ',cur.q_chrome.qsize())
    th = []
    for i in input:
        t = threading.Thread(target=cur.getPage,args=(i,))
        th.append(t)
    for i in th:
        i.start()
    for i in th:
        i.join()
    
    cur.close_chrome()
    cur.writeLogAndData()