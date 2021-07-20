from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import openpyxl
import time

file = pd.read_excel('E:\sabudh INTERNSHIP ps2\seleniumbot\sampledata.xlsx', engine='openpyxl')
n = (len(file['sentence']))
sentences1 = []
sentences2 = []
sentences3 = []
sentences4 = []
sentences5 = []



driver=webdriver.Chrome(executable_path="C:\Program Files (x86)\chromedriver.exe")
driver.get("https://quillbot.com/")

for i in range(5): 

    element = driver.find_element_by_xpath('//*[@id="inputText"]')
    element.send_keys(file['sentence'][i])

    element=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[1]/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div/button/span/div')
    action = ActionChains(driver).click(on_element = element).perform()
    time.sleep(10)
    element = driver.find_element_by_xpath('//*[@id="editable-content-within-article~0"]/div')
    print('1st')
    print(element.text)
    sentences1.append(element.text)
    driver.implicitly_wait(10)
    time.sleep(5)

    element = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/button[2]')
    action = ActionChains(driver).click(on_element = element).perform()
    time.sleep(10)
    element = driver.find_element_by_xpath('//*[@id="editable-content-within-article~0"]/div')
    print('2nd')
    print(element.text)
    sentences2.append(element.text)
    driver.implicitly_wait(10)
    time.sleep(5)

    element = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/button[2]')
    action = ActionChains(driver).click(on_element = element).perform()
    time.sleep(10)
    element = driver.find_element_by_xpath('//*[@id="editable-content-within-article~0"]/div')
    sentences3.append(element.text)
    print('3rd')
    print(element.text)
    driver.implicitly_wait(10)
    time.sleep(5)

    element = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/button[2]')
    action = ActionChains(driver).click(on_element = element).perform()
    time.sleep(10)
    element = driver.find_element_by_xpath('//*[@id="editable-content-within-article~0"]/div')
    print('4th')
    print(element.text)
    sentences4.append(element.text)
    driver.implicitly_wait(10)
    time.sleep(5)

    element = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/button[2]')
    action = ActionChains(driver).click(on_element = element).perform()
    time.sleep(10)
    element = driver.find_element_by_xpath('//*[@id="editable-content-within-article~0"]/div')
    sentences5.append(element.text)
    print('5th')
    print(element.text)
    time.sleep(5)

    driver.refresh()

df = pd.DataFrame()

df['1st Reframed sentences'] = sentences1
df['2nd Reframed sentences'] = sentences2
df['3rd Reframed sentences'] = sentences3
df['4th Reframed sentences'] = sentences4
df['5th Reframed sentences'] = sentences5

df.index += 1

print(df)

df.to_csv('rephrase.csv')

'''
with pd.ExcelWriter('sampledata.xlsx') as writer:
    writer.book = openpyxl.load_workbook('sampledata.xlsx')
    df.to_excel(writer)
'''

driver.close()
