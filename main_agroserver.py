from bs4 import BeautifulSoup
import lxml
from selenium import webdriver
import openpyxl
from openpyxl.styles import Alignment
import time
from fake_useragent import UserAgent
from selenium_stealth import stealth

# ====================================================
book = openpyxl.Workbook()
sheet = book.active
sheet.column_dimensions['A'].width = 50
sheet.column_dimensions['B'].width = 60
sheet.column_dimensions['C'].width = 120
sheet.column_dimensions['D'].width = 18
sheet.column_dimensions['E'].width = 35
sheet.column_dimensions['F'].width = 20
sheet.column_dimensions['G'].width = 45
sheet.column_dimensions['H'].width = 80
sheet["A1"] = 'Название категории'
sheet["B1"] = 'Название объявления'
sheet["C1"] = 'текст объявления'
sheet["D1"] = 'Цена'
sheet["E1"] = 'Имя продавца'
sheet["F1"] = 'телефон'
sheet["G1"] = 'Адрес'
sheet["H1"] = 'Ссылка на объявление'
# ====================================================
useragent = UserAgent()
# Настройка драйвера
options = webdriver.ChromeOptions()
options.add_argument("start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument(f"user-agent={useragent.random}")
options.headless = True
driver = webdriver.Chrome(options=options)
stealth(driver,
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        )
row = 2
# =================================================================  5.61.58.211:4008
for num_page in range(1, 24):
    count = 0
    url = f"https://agroserver.ru/semena/Y2l0eT18cmVnaW9uPXxjb3VudHJ5PXxtZXRrYT18c29ydD0x/{num_page}/"
    driver.get(url)
    main_page = driver.page_source
    soup = BeautifulSoup(main_page, "lxml")
    while soup.find(class_="notfound"):
        time.sleep(600)
        driver.close()
        options.add_argument(f"user-agent={useragent.random}")
        driver = webdriver.Chrome(options=options)
        driver.get(url=url)
        main_page = driver.page_source
        soup = BeautifulSoup(main_page, "lxml")
    with open("main_page.html", "w", encoding="utf-8") as file_main_page:
        file_main_page.write(main_page)
    with open("main_page.html", "r", encoding="utf-8") as file_main_page:
        soup = BeautifulSoup(file_main_page, "lxml")
        link_on_product_list = soup.findAll(class_="line")
    print(f"Страница {num_page} со списком товаров записана")
    for link_on_product_elem in link_on_product_list:
        text = ''
        link_on_product = link_on_product_elem.find(class_="th").find('a').get("href")
        url_page = 'https://agroserver.ru' + link_on_product
        driver.get(url_page)
        page = driver.page_source
        soup_page = BeautifulSoup(page, "lxml")
        while soup_page.find(class_="notfound"):
            time.sleep(600)
            driver.close()
            options.add_argument(f"user-agent={useragent.random}")
            driver = webdriver.Chrome(options=options)
            driver.get(url=url_page)
            page = driver.page_source
            soup_page = BeautifulSoup(main_page, "lxml")
        with open("page.html", "w", encoding="utf-8") as file_page:
            file_page.write(page)
        with open("page.html", "r", encoding="utf-8") as file_page:
            soup_page = BeautifulSoup(page, "lxml")
            for br in soup_page('br'):
                br.replace_with('\n')
            try:
                category = soup_page.find(class_="like_block_new").find(class_="title").text
            except:
                category = 'Нет категории'
            try:
                name = soup_page.find(class_ = "bhead").text
            except:
                name = 'Нет имени'
            try:
                text_list = soup_page.find(class_="text").findAll("p")
            except:
                text_list = ['Нет описания']
            try:
                price = soup_page.find(class_="mprice").text
            except:
                price = 'Нет цены'
            try:
                name_seller = soup_page.find(class_ = "bl org ico_company").find("a").text
            except:
                name_seller = 'Нет имени продавца'
            try:
                phone = soup_page.find(class_="phones_all").find("div").text
            except Exception:
                phone = soup_page.find(class_="bl phone ico_call").text
            except:
                phone = 'Нет номера'
            try:
                address = soup_page.find(class_="bl ico_location").text.strip()
            except:
                address = 'Нет адреса'
        link = url_page
        sheet.row_dimensions[row].height = 70
        sheet[row][0].value = category
        sheet[row][1].value = name
        try:
            for text_elem in text_list:
                text += text_elem.text.strip() + '\n'
            sheet[row][2].value = text
        except:
            print("Нет описания (Неверно записано)")
        sheet[row][3].value = price.replace('цена:', '')
        sheet[row][4].value = name_seller
        sheet[row][5].value = phone
        sheet[row][6].value = address
        sheet[row][7].value = link
        for cell in sheet[row]:
            if cell.value:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrapText=True)
        row += 1
        book.save('agroserver.xlsx')
        count += 1
        print(f"Записана {count} страница с товаром, и сохранена в excel файл")
        time.sleep(5)
book.close()
driver.quit()