from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import xlrd

def login(url):
    browser.get(url)
    time.sleep(1)
    browser.find_element_by_xpath('//input [@id="UName"]').send_keys(account_cols[i])
    browser.find_element_by_xpath('//input [@id="PassWord"]').send_keys('123456')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@id="LoginButton"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/a'))).click()
    return browser

def question():
    def do_test():
        browser.find_element_by_xpath('//*[@id="buzhou2ss"]').click()
        time.sleep(1)
        browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[4]/div/dl[1]/div/dd[2]/input').click()
        browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[4]/div/dl[2]/div/dd[1]/input').click()
        browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[4]/div/dl[3]/div/dd[2]/input').click()
        browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[4]/div/dl[4]/div/dd[3]/input').click()
        browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[4]/div/dl[5]/div/dd[2]/input').click()
        browser.find_element_by_xpath('//*[@id="input_button"]').click()
    window = browser.current_window_handle
    browser.find_element_by_xpath('//a [@id="stuCss"]').click()
    browser.find_element_by_link_text('让学生掌握常见疾病的家庭护理方法。').click()
    time.sleep(10)
    handles = browser.window_handles
    for handle in handles:
        if browser.current_window_handle != handle:
            browser.switch_to.window(handle)
    browser.find_element_by_xpath('/html/body/form/div[3]/div[3]/div[1]/p/video').click()
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/form/div[3]/div[3]/div[1]/p/video').click()
    time.sleep(1)
    try:
        do_test()
    except:
        question()
    time.sleep(2)
    browser.close()
    browser.switch_to.window(window)

if __name__ == '__main__':
    browser = webdriver.Chrome()
    url = 'http://henan.xueanquan.com/login.html'
    data = xlrd.open_workbook(r'D:\pythonproject\safetytest\4class.xls')
    sheet0 = data.sheet_by_name("Sheet0")
    number_cols = sheet0.col_values(0)
    name_cols = sheet0.col_values(1)
    account_cols = sheet0.col_values(3)
    wait = WebDriverWait(browser, 10)
    for i in range(57,sheet0.nrows):
        try:
            login(url)
            question()
            progress = str(round((((i + 1) / sheet0.nrows) * 100), 2))
            print("序号"+str(int(number_cols[i]))+name_cols[i] + "的让学生掌握常见疾病的家庭护理方法专题已完成，当前进度" + progress + "%")
        except:
            time.sleep(2)
            progress = str(round((((i + 1) / sheet0.nrows) * 100), 2))
            print("序号"+str(int(number_cols[i]))+name_cols[i] + "的账号出现异常,当前进度" + progress + "%")
            browser.close()
            pass
    print("您的安全教育专题已完成，感谢您的使用——高二十班张权艺出品")


