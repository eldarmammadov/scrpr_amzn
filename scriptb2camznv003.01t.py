import pandas as pd

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService

import openpyxl



def get_lastrow():
    global xl_file
    xl_file = pd.read_excel('./outputFile/output.xlsx', sheet_name=1)
    last_row = (xl_file['URL'].index[-1])+1
    #print(last_row)
    return last_row

def get_url(ind):
    url_xlfile=xl_file.loc[ind]["URL"]
    #print(url_xlfile,ind)
    return url_xlfile

# going to url(amzn) via Selenium WebDriver
chrome_options = Options()
chrome_options.headless = False
chrome_options.add_argument("start-maximized")
# options.add_experimental_option("detach", True)
chrome_options.add_argument("--no-sandbox")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument('--disable-blink-features=AutomationControlled')

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

lst_stock=[]
lst_dlevirey=[]
lst_price=[]
lst_rating=[]

ls_rw=get_lastrow()
try:
    for i in range(0,ls_rw):
        driver.get(get_url(i))

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "sp-cc-accept"))).click()
            #print('accepted cookies')
        except Exception as e:
            pass
            #print('no cookie button!')

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@data-action-type="SELECT_LOCATION"]'))).click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//option[contains(@value,"ZA")]'))).click()
            driver.refresh()
        except:
            pass

        def fnd_stock():
            a= WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@id="availabilityInsideBuyBox_feature_div"]/div/div[@id="availability"]/span')))
            #print(a.get_attribute('innerText'),'inside f')
            return a

        try:
            v_stcok=fnd_stock()
            #print(v_stcok.get_attribute('innerText'),'inside try')
            lst_stock.append(v_stcok.get_attribute('innerText'))
        except:
            #print('Currently unavailable')
            lst_stock.append('Currently unavailable')

        def fnd_delivery_duration():
            a = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//div[@id="mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE"]/span/span')))
            #print(a.text, a.get_attribute('innerText'))
            return a

        try:
            ddm=fnd_delivery_duration()
            #print(ddm.get_attribute('innerText'))
            lst_dlevirey.append(ddm.get_attribute('innerText'))
        except:
            #print('deliver to South Africa')
            lst_dlevirey.append('Currently unavailable')

        def fnd_price():
            a= driver.find_element(By.XPATH,'//span[contains(@class,"aok-align-center")]/span')
            #print(a.get_attribute('innerText'))
            return a

        try:
            v_price=fnd_price()
            #print(v_price.get_attribute('innerText'))
            lst_price.append(v_price.get_attribute('innerText'))
        except:
            #print('price could not scraped')
            lst_price.append('price could not scraped')

        def fnd_rate():
            a=driver.find_element(By.XPATH,'//span[@id="acrCustomerReviewText"]')
            return a

        try:
            v_rate=fnd_rate()
            #print(v_rate.text,v_rate.get_attribute('innerText'))
            lst_rating.append(v_rate.get_attribute('innerText'))
        except:
            lst_rating.append('rateless')
except:
    pass

#print(lst_stock)
#print(lst_dlevirey)
#print(lst_rating)
#print(lst_price)
xl_file['STATUS']=lst_stock
xl_file['DELIVERY']=lst_dlevirey
xl_file['Ratings']=lst_rating
xl_file['PRICE']=lst_price
with pd.ExcelWriter('./outputFile/output.xlsx', engine='openpyxl', mode='r+', if_sheet_exists='overlay') as writer:
    xl_file.to_excel(writer,sheet_name='URLS P1',startrow=0, index=False, header=True)

driver.quit()