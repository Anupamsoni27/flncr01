import time

from selenium import webdriver
import openpyxl
from selenium.webdriver.common.by import By
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup as soup
# from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.utils import ChromeType

chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--no-sandbox')
# chrome_options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install(),chrome_options= chrome_options)
driver.get("https://www.dhirtekbusinessresearch.com/admin")

driver.maximize_window()

username_element=driver.find_element(By.CLASS_NAME,"form-control")
s = "a@dbrc.com"
username_element.send_keys(s)


password_element=driver.find_element(By.NAME,"pass")
password_element.send_keys("123456")
driver.find_element(By.CLASS_NAME,"log-buton").click()

driver.get("https://www.dhirtekbusinessresearch.com/admin/collateralReport")

driver.get("https://www.dhirtekbusinessresearch.com/admin/collateralReport/addcrtemp")


df = pd.read_excel("Collateral.xlsx",engine="openpyxl")
# print(df)
for i in range(0,len(df["Menu Description 1"])):
    df["Menu Description 1"][i]=df["Menu Description 1"][i].replace('_x000D_','')
for i in range(0,len(df["Menu Description 2"])):
    df["Menu Description 2"][i]=df["Menu Description 2"][i].replace('_x000D_','')
# print(df)

for i in range(0,len(df["Meta Title"])):
    driver.find_element(By.NAME,"metatitle").send_keys(df["Meta Title"][i])
    driver.find_element(By.NAME,"metakey").send_keys(df["Meta Keywords"][i])
    driver.find_element(By.NAME,"metadesc").send_keys(df["Meta Description"][i])
    industry= driver.find_element(By.XPATH,"//select[@id='indstry']")
    industry.send_keys(df["Industry"][i])
    sub_industry=driver.find_element(By.XPATH,"//select[@name='subindustry']")
    driver.find_element(By.NAME,"titleEnglish").send_keys(df["Title"][i])
    driver.find_element(By.NAME,"shortTitle").send_keys(df["Short Title"][i])
    driver.find_element(By.NAME,"reportCode").send_keys(df["Report Code"][i])
    df['Number of Pages'] = df['Number of Pages'].astype(str)
    driver.find_element(By.NAME,"noPages").send_keys(df["Number of Pages"][i])
    button = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[2]/div[3]/input[1]")
    driver.execute_script("arguments[0].click();", button)
    button = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[2]/div[3]/input[3]")
    driver.execute_script("arguments[0].click();", button)
    button = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[2]/div[3]/input[4]")
    driver.execute_script("arguments[0].click();", button)
    df['Single User Price'] = df['Single User Price'].astype(str)
    df['Multi User Price'] = df['Multi User Price'].astype(str)
    df['Corporate Price'] = df['Corporate Price'].astype(str)
    driver.find_element(By.NAME,"price1").send_keys(df["Single User Price"][i])
    driver.find_element(By.NAME,"price2").send_keys(df["Multi User Price"][i])
    driver.find_element(By.NAME,"price3").send_keys(df["Corporate Price"][i])
    editorFrame = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[4]/div/div/div/div/div/iframe")
    driver.switch_to.frame(editorFrame)
    body =driver.find_element(By.TAG_NAME,"p")
    text=df["Description"][i]
    driver.execute_script("arguments[0].innerHTML = '"+text+"'", body)
    driver.switch_to.default_content()
    button=driver.find_element(By.CLASS_NAME,"btn-success")
    driver.execute_script("arguments[0].click();", button)
    driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[5]/div[2]/div[1]/div[1]/div/textarea").send_keys(df["Menu Title 1"][i])
    driver.switch_to.default_content()
    editorFrame = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[5]/div[2]/div[2]/div/div/div/div/div/div/iframe")
    driver.switch_to.frame(editorFrame)
    body =driver.find_element(By.TAG_NAME,"p")
    txt=df["Menu Description 1"][i]
    txt=txt.replace("'","\\'")
    txt=txt.replace("\n","")
    driver.execute_script("arguments[0].innerHTML= '"+txt+"'", body)
    driver.switch_to.default_content()
    driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[5]/div[3]/div/div[1]/div[1]/div/textarea").send_keys(df["Menu Title 2"][i])
    editorFrame = driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[5]/div[3]/div/div[2]/div/div/div/div/div/div/iframe")
    driver.switch_to.frame(editorFrame)
    body = driver.find_element(By.TAG_NAME,"p")
    txt=df["Menu Description 2"][i]
    # txt=txt.replace("'","\\'")
    # txt=txt.replace("\n","<br>")
    # driver.execute_script("arguments[0].innerHTML= '"+txt+"'", body)
    body.send_keys(txt)
    driver.switch_to.default_content()
    button=driver.find_element(By.XPATH,"/html/body/div[1]/div/section[2]/div/div/div/div/div[1]/form/div/div[3]/div/div[6]/div/center/button[1]")
    driver.execute_script("arguments[0].click();", button)
    driver.get("https://www.dhirtekbusinessresearch.com/admin/collateralReport/addcrtemp")
    time.sleep(3)