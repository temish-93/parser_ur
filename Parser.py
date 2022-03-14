# программа парсинга списка товаров с сайта МК без селениума

import pandas as pd
import requests
from bs4 import BeautifulSoup
import random
import time
import ctypes                                                                                                           #для простых диалоговых окон
from tkinter import *                                                                                                   #создание интерфейса
from tkinter.ttk import *
import tkinter.filedialog as fd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService

def quit():
    window.destroy()

def choose_file():
    global file_for_monitoring
    filetypes = (("Excel файл", "*.xls *.xlsx *.xlsm"),
                 ("Любой", "*"))
    file_for_monitoring = fd.askopenfilename(title="Выбрать файл", initialdir="/",filetypes=filetypes)
    Button_quit = Button(window, text="Запустить парсер", command=quit)
    Button_quit.grid(column=0, row=3, sticky=W)

def choice_search_file(event):
    global competitor_for_monitoring
    competitor_for_monitoring = competitors.get()
    lbl = Label(window, text="Выберите конкурента для парсинга:", font=("Arial Bold", 10))
    lbl.grid(column=0, row=0, sticky=W)
    competitor1 = Combobox(window, values=[competitor_for_monitoring], state="disabled", width=50)
    competitor1.grid(column=1, row=0, sticky=W)
    competitor1.current(0)  # установите вариант по умолчанию
    search_file = Button(window, text="Выбрать файл", command=choose_file)
    search_file.grid(column=1, row=2, sticky=W)
    window.mainloop()

window = Tk()
window.title("Parser of competitors online / Парсер конкурентов онлайн")
window.geometry('800x500')
lbl = Label(window, text="Выберите конкурента для парсинга:", font=("Arial Bold", 10))
lbl.grid(column=0, row=0, sticky=W)
lbl2 = Label(window, text="Укажите путь к файлу с товарами:", font=("Arial Bold", 10))
lbl2.grid(column=0, row=2, sticky=W)
lbl3 = Label(window, text="Если в списке нет необходимого магазина конкурента,", font=("Arial Bold", 10))
lbl3.place(relx=.01, rely=.92)
lbl4 = Label(window, text="обратитесь к разработчику для актуализации базы адресов", font=("Arial Bold", 10))
lbl4.place(relx=.01, rely=.96)
lbl5 = Label(window, text="developed by Potapov / 2022", font=("Arial Bold", 10))
lbl5.place(relx=.785, rely=.96)
competitors = Combobox(window, values=['Выберите из списка', 'Ватсонс', 'Впрок', 'Магнит Косметик', 'Подружка'],
                   state="readonly", width=50)
competitors.grid(column=1, row=0)
competitors.current(0)  # установите вариант по умолчанию
competitors.bind("<<ComboboxSelected>>", choice_search_file)
window.mainloop()


df_parser_result = pd.DataFrame({'Конкурент': [],
                   'Код Товара': [],
                   'Категория': [],
                   'Наименование Товара': [],
                   'Цена, руб.': [],
                   'Акц цена, руб.': [],
                   'Цена в приложении, руб.': []
                   })

df_file_for_monitoring = pd.read_excel(file_for_monitoring).drop_duplicates().reset_index(drop=True) #наш файл для парсинга
try:
    df_product_library = pd.read_excel(r'C:\Users\User\Desktop\Parser\data\product_library.xlsx') #наша база товаров личный ПК
except:
    df_product_library = pd.read_excel(r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
                                       r'\5. Мониторинги цен конкурентов\Parser\build_windows\product_library.xlsx')  # наша база товаров ПК УР
df_file_for_monitoring2 = df_file_for_monitoring.merge(df_product_library, on=["Код Товара"]) #объединем файл поиска с нашей базой товаров
df_file_for_monitoring2 = df_file_for_monitoring2.fillna('Нет')


if competitor_for_monitoring == 'Подружка':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_подружка'].str.contains('http')]\
    .drop_duplicates()#оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)                                    #количество строк в нем
    #i2 = 50                                               #тестовое количество строк в нем
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_подружка']  # адрес страницы, с которой будет поступать информация
        response = requests.get(url, timeout=15)
        time.sleep(random.randint(2, 10))  # время ожидания в секундах
        response.encoding = 'utf-8'
        headers = (response.headers)
        soup = BeautifulSoup(response.text, 'lxml')
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']  # адрес страницы, с которой будет поступать информация
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        i3 += 1
        try:
            Podrygka_product_price_rub = soup.find('span', class_='price__item price__item--old').text
            Podrygka_product_price_rub_action = soup.find('span', class_='price__item price__item--current').text
        except:
            try:
                Podrygka_product_price_rub = soup.find('span', class_='price__item price__item--current').text
                Podrygka_product_price_rub_action = ''
            except:
                Podrygka_product_price_rub = 'Нет в наличии'
                Podrygka_product_price_rub_action = ''


        new_row = {'Конкурент':'Подружка',
                'Код Товара':kod_tovara,
                'Категория': kategotya,
                'Наименование Товара':tovar_name,
                'Цена, руб.':Podrygka_product_price_rub,
                'Акц цена, руб.':Podrygka_product_price_rub_action}


        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
    df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
    df_parser_result = df_parser_result.replace([r'\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)  # вывод окна с информацией

elif competitor_for_monitoring == 'Впрок':
    # настройка селениума и браузера
    CHROMEDRIVER_PATH = (r'C:\SOFT\chromedriver.exe')  # это наш путь к драйверу на личном ПК
    #    CHROMEDRIVER_PATH = (r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
    #          r'\5. Мониторинги цен конкурентов\Parser\build_windows\chromedriver.exe')  # это наш путь к драйверу на ПК УР

    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager'  # ожидание загрузки стр полностью
    options.add_argument('headless')  # для открытия headless-браузера
    prefs = {"profile.default_content_setting_values.geolocation": 2,  # отключаем геолокацию
             "profile.default_content_setting_values.notifications": 2,  # отключаем показ уведомлений
             "profile.managed_default_content_settings.images": 2}  # отключаем показ изображений
    # options.add_experimental_option("prefs", prefs)  # применяем настройки профиля
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = ChromeService(executable_path=CHROMEDRIVER_PATH)
    browser = webdriver.Chrome(options=options, service=service)
    browser.maximize_window()

    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_search_впрок'].str.contains('http')]\
    .drop_duplicates() #оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)                                    #количество строк в нем
#    i2 = 3                                               #тестовое количество строк в нем
    i3 = 0

    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_search_впрок']  # адрес страницы, с которой будет поступать информация
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']  # адрес страницы, с которой будет поступать информация
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        browser.get(url)  # Переход на страницу поиска товара
        time.sleep(3)  # время ожидания в секундах
        try:
            browser.find_element_by_class_name('col-xs-10').click()
            time.sleep(2)  # время ожидания в секундах
            try:
                price_shop = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[5]/'
                                                           'span').text.replace('.', ',')
                price_pril = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[7]/'
                                                           'div').text.replace('.', ',')
            except:
                price_shop = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[4]/'
                                                           'span').text.replace('.', ',')
                price_pril = ''

            new_row = {'Конкурент':'Впрок',
                   'Код Товара':kod_tovara,
                   'Категория':kategotya,
                   'Наименование Товара':tovar_name,
                   'Цена, руб.':price_shop,
                   'Цена в приложении, руб.':price_pril
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                      regex=True)  # удаляем лишние пробелы и переходы строк
            i3 += 1

        except:
            time.sleep(0)  # время ожидания в секундах
            i3 += 1

        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor') # - удалить в проде!!!

    browser.close()
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)  # вывод окна с информацией

elif competitor_for_monitoring == 'Ватсонс':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[
        df_file_for_monitoring2['url_search_ватсонс'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)                                    #количество строк в нем
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[
            i3, 'url_search_ватсонс']  # адрес страницы, с которой будет поступать информация
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
        response = requests.get(url, headers=headers, timeout=15)
        time.sleep(random.randint(2, 10))  # время ожидания в секундах
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')
        kod_tovara = df_file_for_monitoring3.loc[
            i3, 'Код Товара']  # адрес страницы, с которой будет поступать информация
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        i3 += 1
        try:
            watsons_product_price_rub = soup.find('div', class_='product-tile__price product-tile__'
                                                    'price--old js-variant-old-price').text.replace(' руб', '')
            watsons_product_price_rub_action = soup.find('div', class_='product-tile__price product-tile__'
                                                        'price--discounted js-variant-price').text.replace(' руб', '')
        except:
            try:
                watsons_product_price_rub = soup.find('div', class_='product-tile__price product-tile__'
                                                        'price--original js-variant-price').text.replace(' руб', '')
                watsons_product_price_rub_action = ''
            except:
                watsons_product_price_rub = 'Нет в наличии'
                watsons_product_price_rub_action = ''

        new_row = {'Конкурент': 'Ватсонс',
                   'Код Товара': kod_tovara,
                   'Категория': kategotya,
                   'Наименование Товара': tovar_name,
                   'Цена, руб.': watsons_product_price_rub,
                   'Акц цена, руб.': watsons_product_price_rub_action}

        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
        df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
        df_parser_result = df_parser_result.replace([r'\n', r'            '], '', regex=True)  # удаляем лишние пробелы и переходы строк
        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)  # вывод окна с информацией

elif competitor_for_monitoring == 'Магнит Косметик':
    # настройка селениума и браузера
    CHROMEDRIVER_PATH = (r'C:\SOFT\chromedriver.exe')  # это наш путь к драйверу на личном ПК
    #    CHROMEDRIVER_PATH = (r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
    #          r'\5. Мониторинги цен конкурентов\Parser\build_windows\chromedriver.exe')  # это наш путь к драйверу на ПК УР

    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager'  # ожидание загрузки стр полностью
    #options.add_argument('headless')  # для открытия headless-браузера
    prefs = {"profile.default_content_setting_values.geolocation": 2,  # отключаем геолокацию
             "profile.default_content_setting_values.notifications": 2,  # отключаем показ уведомлений
             "profile.managed_default_content_settings.images": 2}  # отключаем показ изображений
    # options.add_experimental_option("prefs", prefs)  # применяем настройки профиля
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = ChromeService(executable_path=CHROMEDRIVER_PATH)
    browser = webdriver.Chrome(options=options, service=service)
    browser.maximize_window()

    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_мк'].str.contains('http')]\
    .drop_duplicates() #оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)                                    #количество строк в нем
#    i2 = 3                                               #тестовое количество строк в нем
    i3 = 0
    FAVORITE_SHOP = 55569
    FAVORITE_SHOP_str = str(FAVORITE_SHOP)



    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_мк']  # адрес страницы, с которой будет поступать информация
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']  # адрес страницы, с которой будет поступать информация
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        browser.get(url)  # Переход на страницу товара
        browser.add_cookie({"name": "FAVORITE_SHOP", "value": FAVORITE_SHOP_str})
        time.sleep(10)  # время ожидания в секундах
        try:
            shop_city = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                  'header/div/div[1]/div[1]/div/div/a').text
            shop_address = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                     'header/div/div[4]/div[1]/a/span').text
            price_shop = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[1]/div/div[2]/'
                                                           'div/div[2]/div[1]/div[4]/div[1]/div/div[1]').text

            new_row = {'Конкурент':'Магнит Косметик',
                   'Код Товара':kod_tovara,
                   'Категория':kategotya,
                   'Наименование Товара':tovar_name,
                   'Цена, руб.':price_shop,
                   'Адрес конкурента': shop_address
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                      regex=True)  # удаляем лишние пробелы и переходы строк
            i3 += 1
#            FAVORITE_SHOP += 1
#            FAVORITE_SHOP_str = str(FAVORITE_SHOP)

        except:
            browser.navigate().refresh()
            time.sleep(8)  # время ожидания в секундах
            shop_city = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                      'header/div/div[1]/div[1]/div/div/a').text
            shop_address = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                         'header/div/div[4]/div[1]/a/span').text
            price_shop = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[1]/div/div[2]/'
                                                       'div/div[2]/div[1]/div[4]/div[1]/div/div[1]').text

            new_row = {'Конкурент': 'Магнит Косметик',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена, руб.': price_shop,
                       'Адрес конкурента': shop_address
                       }
            i3 += 1
#            FAVORITE_SHOP += 1
#            FAVORITE_SHOP_str = str(FAVORITE_SHOP)

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                        regex=True)  # удаляем лишние пробелы и переходы строк

        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')  # - удалить в проде!!!


    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)  # вывод окна с информацией
    browser.close()





else:
    ctypes.windll.user32.MessageBoxW(0, "Остальные конкуренты находятся в разработке", "Информация", 0) #вывод окна с информацией