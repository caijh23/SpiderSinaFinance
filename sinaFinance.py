from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
import xlsxwriter
import xlrd

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

def openHomePage(browser,url):
    browser.set_page_load_timeout(240)
    try:
        browser.get(url)
    except TimeoutException:
        print('stop load')
        browser.execute_script('window.stop ? window.stop() : document.execCommand("Stop");')
        print('exit try')
    return None

def searchCompany(browser,search_content):
    print('start search input box')
    input = browser.find_element_by_id('suggest01_input')
    input.send_keys(search_content)
    button = browser.find_element_by_id('topSearchSubmit')
    button.send_keys(Keys.ENTER)
    return None

def jumpIntoCompany(current_handle,handles,new_handle,browser,search_content):
    href = browser.find_element_by_xpath("//div/label/a/span[contains(text(), %s)]" % search_content)
    href.click()
    switch_page(current_handle,handles,new_handle,browser)

def jumpIntoSeniorExecutive(current_handle,handles,new_handle,browser):
    seniorExecutive = browser.find_element_by_xpath("//li/a[contains(text(),'公司高管')]")
    seniorExecutive.click()
    switch_page(current_handle,handles,new_handle,browser)

def jumpIntoResume(current_handle,handles,new_handle,browser,seniorExecutive_name):
    hreflists = browser.find_elements_by_tag_name('a')
    for href in hreflists:
        if href.text == seniorExecutive_name:
            href.click()
            break
    switch_page(current_handle,handles,new_handle,browser)

def spiderInfo(browser,search_content,seniorExecutive_name, row, col):
    education = browser.find_element_by_xpath("//table[@id='Table1']/tbody/tr[1]/td[4]/div")
    resumeInfo = browser.find_element_by_xpath("//table[@id='Table1']/tbody/tr[2]/td[2]")
    writeIntoXlsx(education.text,resumeInfo.text,search_content,seniorExecutive_name, row, col)

def spider(search_content, seniorExecutive_name, row, col):
    browser = webdriver.Chrome()
    url = 'https://finance.sina.com.cn/'
    current_handle = None
    handles = None
    new_handle = None
    openHomePage(browser,url)
    searchCompany(browser,search_content)
    switch_page(current_handle,handles,new_handle,browser)
    jumpIntoCompany(current_handle,handles,new_handle,browser,search_content)
    jumpIntoSeniorExecutive(current_handle,handles,new_handle,browser)
    jumpIntoResume(current_handle,handles,new_handle,browser,seniorExecutive_name)
    spiderInfo(browser,search_content,seniorExecutive_name, row, col)
    browser.quit()


def main():
    workbook = xlsxwriter.Workbook('output.xlsx')
    global worksheet
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    worksheet.write('A1', 'Code', bold)
    worksheet.write('B1', 'CEO_name', bold)
    worksheet.write('C1', '学历', bold)
    worksheet.write('D1', '简历', bold)
    row = 1
    col = 0
    inputbook = xlrd.open_workbook('./input.xlsx')
    sheet = inputbook.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i == 0:
            continue
        search_content = sheet.cell(row, col).value
        seniorExecutive_name = sheet.cell(row, col + 1).value
        spider(search_content,seniorExecutive_name,row,col)
        row += 1
    workbook.close()

def writeIntoXlsx(education,resumeInfo,search_content,seniorExecutive_name, row, col):
    worksheet.write_string(row, col, search_content)
    worksheet.write_string(row, col + 1, seniorExecutive_name)
    worksheet.write_string(row, col + 2, education)
    worksheet.write_string(row, col + 3, resumeInfo)

if __name__=="__main__":
    main()