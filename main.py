from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import re
import urllib.parse
import pandas as pd

# Опции хрома
chrome_options = Options()
chrome_options.add_argument("--headless")  # отлючаем ГУИ
chrome_options.add_experimental_option("prefs", {"geolocation": "disabled"})
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
options = webdriver.ChromeOptions()
options.add_argument("--disable-popup-blocking")
driver = webdriver.Chrome(options=options)

# Указываем путь к хромдрайверу
webdriver_service = Service(r"chromedriver.exe")  

# Выбираем хром
driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)

# Указываем имя файла xls, где столбец А - список значений ИКЗ
df = pd.read_excel('', usecols='A', header=None)

column_values = df.iloc[:,0].values
print(column_values)

# Вписываем URL
url = "https://zakupki.gov.ru/epz/contract/search/results.html"

driver.get(url)

data = []

for value in column_values:
    driver.get(url)
 
# Ищем, убиваем модальные окна
    wait = WebDriverWait(driver, 10)
    try:
        WebDriverWait(driver, 10).until(
        EC.presence_of_element_located([By.ID, "modal-region"])
        )
        modalwindow = driver.find_element(By.ID, "modal-region")
        if modalwindow:
            autobutton = driver.find_element(By.CSS_SELECTOR, ".btn-close.closePopUp")
            autobutton.click()
    except Exception as e:
        print('no modal window')

    search_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located([By.CSS_SELECTOR, "#searchString"])
    )
    search_field.send_keys(value)
    
    wait = WebDriverWait(driver, 10)

# работа с поиском
    
    search_button = driver.find_element(By.CSS_SELECTOR, "#quickSearchForm_header > section.content.content-search-registry-bar > div > div > div > div:nth-child(2) > div > div > button")
    
    search_button.click()
   
    wait = WebDriverWait(driver, 10)
    try:
        noRecords = driver.find_element(By.CSS_SELECTOR, "#quickSearchForm_header .noRecords")
        if noRecords:   
            print('no contracts found')
            page_data = {
                    'Номер в реестре': 'информация отсутствует',
                    'ИКЗ': value,
                    'Цена контракта': 'информация отсутствует',
                    'Стоимость исполненных поставщиком (подрядчиком, исполнителем) обязательств': 'информация отсутствует',
                    'Фактически оплачено': 'информация отсутствует',
                    'Дата заключения контракта' : 'информация отсутствует',
                    'Срок исполнения': 'информация отсутствует',
                    'URL': 'информация отсутствует'
                }
            data.append(page_data)
            continue
    except Exception as e:
        print('record found')

    articles = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#quickSearchForm_header > section.content.content-search-registry-block > div > div > div.col-9.search-results > div.search-registry-entrys-block > div > div.row.no-gutters.registry-entry__form.mr-0 > div.col-8.pr-0.mr-21px > div.registry-entry__header > div > div.registry-entry__header-mid__number > a")))
# экстрактим хрефы
    hrefs = [article.get_attribute("href") for article in articles if article.get_attribute("href")]

    for href in hrefs:
        original_url = href
        new_url_base = "https://zakupki.gov.ru/epz/contract/contractCard/process-info.html"
       
        parsed_url = urllib.parse.urlparse(original_url)
        params = urllib.parse.parse_qs(parsed_url.query)
        reestr_number = params['reestrNumber'][0]

        new_url = f"{new_url_base}?reestrNumber={reestr_number}"
        print(reestr_number)

        if new_url:
            driver.get(new_url)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
# Номер в реестре       
            reestrNumber = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(3) > div.cardMainInfo.row > div.sectionMainInfo.borderRight.col-6 > div.sectionMainInfo__header > div > span.cardMainInfo__purchaseLink.distancedText > a')
# Статус       
            status = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(3) > div.cardMainInfo.row > div.sectionMainInfo.borderRight.col-6 > div.sectionMainInfo__header > div > span.cardMainInfo__state.distancedText')
            print(status.text)
# Дата заключения контракта
            wait = WebDriverWait(driver, 10)
            try:
                contDate = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(3) > div.cardMainInfo.row > div.sectionMainInfo.borderRight.col-3.colSpaceBetween > div.date.mt-auto > div:nth-child(1) > span.cardMainInfo__content')
                if contDate: 
                    contDate = contDate.text
                print(contDate)
            except Exception as e:
                print('contDate: not found')
                payed = 0  
# Срок исполнения        
            wait = WebDriverWait(driver, 10)   
            try: 
                contDateDone = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(3) > div.cardMainInfo.row > div.sectionMainInfo.borderRight.col-3.colSpaceBetween > div.date.mt-auto > div:nth-child(2) > span.cardMainInfo__content')
                if contDateDone: 
                    contDateDone = contDateDone.text
                print(contDateDone)
            except Exception as e:
                print('contDateDone: not found')
                payed = 0
# Вывод цены контракта
            contPrice = 0 
            wait = WebDriverWait(driver, 10)
            try:            
                contPrice = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(3) > div.cardMainInfo.row > div.sectionMainInfo.borderRight.col-3.colSpaceBetween > div.price > span.cardMainInfo__content.cost')
                if contPrice:
                    contPrice = contPrice.text
                    print(contPrice)
            except Exception as e:
                print('contPrice not found')
                payAmount = 0              
# Проверка таба с ценами 
            wait = WebDriverWait(driver, 10)
            try:
                activeTab = driver.find_element(By.CSS_SELECTOR, 'body > div.cardWrapper.outerWrapper > div > div.cardHeaderBlock > div:nth-child(5) > div > a.tabsNav__item.tabsNav__item_active')
                if activeTab.text != "ИСПОЛНЕНИЕ (РАСТОРЖЕНИЕ) КОНТРАКТА": 
                   print('Вкладка "ИСПОЛНЕНИЕ (РАСТОРЖЕНИЕ) КОНТРАКТА" отсутствует')
                   continue
            except Exception as e:
                print('ничего не найдено :()')
                page_data = {
                    'Номер в реестре': reestrNumber.text,
                    'ИКЗ': value,
                    'Статус': status.text,
                    'Цена контракта': contPrice,
                    'Стоимость исполненных поставщиком (подрядчиком, исполнителем) обязательств': 'информация отсутствует',
                    'Фактически оплачено': 'информация отсутствует',
                    'Дата заключения контракта' :contDate,
                    'Срок исполнения': contDateDone,
                    'URL': new_url
                }
                data.append(page_data)
                continue
            

            print(f"Title of the new page: {driver.title}")

            page_title = driver.title
# Вывод стоимости и факта оплаты из таба 
            payAmount = 0
            payed = 0
            wait = WebDriverWait(driver, 10)  
            try:
                priceSections = driver.find_elements(By.CSS_SELECTOR, "body > div.cardWrapper.outerWrapper > div > div.mb-5.pb-3 > div.container > div > div > section.blockInfo__section.section")

                for priceSection in priceSections:
                    section__title = priceSection.find_element(By.CSS_SELECTOR, ".section__title")                    
                    if section__title.text == "Стоимость исполненных поставщиком (подрядчиком, исполнителем) обязательств, ₽": 
                        payAmount = priceSection.find_element(By.CSS_SELECTOR, ".section__info").text
                        print(payAmount)
                    if section__title.text == "Фактически оплачено, ₽": 
                        payed = priceSection.find_element(By.CSS_SELECTOR, ".section__info").text
                        print(payed)
            except Exception as e:
                print('priceSections: not found')
                print(e) 

# page_url
            page_data = {
                'Номер в реестре': reestrNumber.text,
                'ИКЗ': value,
                'Статус': status.text,
                'Цена контракта': contPrice,
                'Стоимость исполненных поставщиком (подрядчиком, исполнителем) обязательств': payAmount,
                'Фактически оплачено': payed,
                'Дата заключения контракта' :contDate,
                'Срок исполнения': contDateDone,
                'URL': new_url
                
            }
            data.append(page_data)

# страница назад
        driver.back()
        element_locator = (By.CSS_SELECTOR, "#searchString")
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(element_locator))
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)

 # закрываем браузер
    driver.quit()

# конвертируем
    df = pd.DataFrame(data)

# создаем xls
    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer, index=False)
    writer.close()
print('###############################################')
print('########### Парсинг завершен успено ###########')
print('###############################################')