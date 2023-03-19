import json
from re import search
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
import random
import time
import ctypes  # для простых диалоговых окон
from tkinter import *  # создание интерфейса
from tkinter.ttk import *
import tkinter.filedialog as fd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By


def quit():
    global number_UR_for_monitoring
    global all_region_for_monitoring
    number_UR_for_monitoring = str(number_UR.get())
    all_region_for_monitoring = str(all_region.get())
    window.destroy()


def choose_file():
    global file_for_monitoring
    filetypes = (("Excel файл", "*.xls *.xlsx *.xlsm"),
                 ("Любой", "*"))
    file_for_monitoring = fd.askopenfilename(title="Выбрать файл", initialdir="/", filetypes=filetypes)
    Button_quit = Button(window, text="Запустить парсер", command=quit)
    Button_quit.grid(column=0, row=5, sticky=W)


def choice_search_file(event):
    global competitor_for_monitoring
    competitor_for_monitoring = competitors.get()
    lbl = Label(window, text="Выберите конкурента для парсинга:", font=("Arial Bold", 10))
    lbl.grid(column=0, row=0, sticky=W)
    competitor1 = Combobox(window, values=[competitor_for_monitoring], state="disabled", width=50)
    competitor1.grid(column=1, row=0, sticky=W)
    competitor1.current(0)  # установите вариант по умолчанию
    search_file = Button(window, text="Выбрать файл", command=choose_file)
    search_file.grid(column=1, row=4, sticky=W)
    window.mainloop()


window = Tk()
window.title("Parser of competitors online / Парсер конкурентов онлайн")
window.geometry('800x500')
lbl = Label(window, text="Выберите конкурента:", font=("Arial Bold", 10))
lbl.grid(column=0, row=0, sticky=W)

lbl2_1 = Label(window, text="Введите номер магазина", font=("Arial Bold", 10))
lbl2_1.grid(column=0, row=1, sticky=W)
number_UR = StringVar()
lbl2_2 = Entry(textvariable=number_UR)
lbl2_2.grid(row=1, column=1, sticky=W)

all_region = BooleanVar()
all_region_checkbutton = Checkbutton(text="все регионы", variable=all_region)
all_region_checkbutton.grid(column=1, row=3, sticky=W)
all_region_label = Label(textvariable=all_region)
lbl2_3 = Label(window, text="Или выберите мониторинг по всем регионам", font=("Arial Bold", 10))
lbl2_3.grid(column=0, row=3, sticky=W)

lbl3 = Label(window, text="Укажите путь к файлу с товарами:", font=("Arial Bold", 10))
lbl3.grid(column=0, row=4, sticky=W)
lbl4 = Label(window, text="Если в списке нет необходимого магазина конкурента,", font=("Arial Bold", 10))
lbl4.place(relx=.01, rely=.92)
lbl5 = Label(window, text="обратитесь к разработчику для актуализации базы адресов", font=("Arial Bold", 10))
lbl5.place(relx=.01, rely=.96)
lbl6 = Label(window, text="developed by Potapov / 2022", font=("Arial Bold", 10))
lbl6.place(relx=.785, rely=.96)
competitors = Combobox(window, values=['Выберите из списка',
                                       'Fix-price',
                                       'Spar',
                                       'Бахетле',
                                       'Впрок',
#                                       'Галамарт',
                                       'Дикси',
                                       'Золотое яблоко',
                                       'Караван',
#                                       'Лента - not work',
                                       'Лэтуаль',
                                       'Магнит',
                                       'Магнит Косметик',
                                       'Магнолия',
                                       'Макси',
                                       'Оптима',
                                       'Подружка',
                                       'Полушка',
                                       'РивГош',
#                                       'Семишагофф - not work',
                                       'Скарлетт - wait'
                                       ],
                       state="readonly", width=50)
competitors.grid(column=1, row=0)
competitors.current(0)  # установите вариант по умолчанию
competitors.bind("<<ComboboxSelected>>", choice_search_file)
window.mainloop()

#######################################################################
# настройка селениума и браузера
CHROMEDRIVER_PATH = r'C:\SOFT\chromedriver.exe'  # это наш путь к драйверу на личном ПК
#CHROMEDRIVER_PATH = (r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
#          r'\5. Мониторинги цен конкурентов\Parser\build_windows\chromedriver.exe')  #это наш путь к драйверу на ПК УР
options = webdriver.ChromeOptions()
options.page_load_strategy = 'eager'  # ожидание загрузки стр полностью
options.add_argument('--disable-blink-features=AutomationControlled')  # отключение опции режима веб драйвера
options.add_argument('headless')  # для открытия headless-браузера

options.add_argument('--no-sandbox')
options.add_argument('start-maximized')
options.add_argument("--disable-infobars")
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option('excludeSwitches', ['enable-automation'])

prefs = {"profile.default_content_setting_values.geolocation": 2,  # отключаем геолокацию
         "profile.default_content_setting_values.notifications": 2,  # отключаем показ уведомлений
         "profile.managed_default_content_settings.images": 2}  # отключаем показ изображений
options.add_experimental_option("prefs", prefs)  # применяем настройки профиля
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument('--proxy-server=http://%s' % proxy_random)
service = ChromeService(executable_path=CHROMEDRIVER_PATH)
browser = webdriver.Chrome(options=options, service=service)
browser.maximize_window()
########################################################################

df_parser_result = pd.DataFrame({'Конкурент': [],
                                 'Код Товара': [],
                                 'Категория': [],
                                 'Наименование Товара': [],
                                 'Цена УР на дату мониторинга, руб.': [],
                                 'Цена, руб.': [],
                                 'Акц цена, руб.': [],
                                 'Цена в приложении, руб.': [],
                                 'Цена по карте, руб.': [],
                                 'Магазин УР': []
                                 })

df_file_for_monitoring = pd.read_excel(file_for_monitoring).drop_duplicates().reset_index(
    drop=True)  # наш файл для парсинга
try:
    df_product_library = pd.read_excel(
        r'C:\Users\Artem\Desktop\Parser\data\product_library.xlsx')  # наша база товаров личный ПК
#        r'C:\Users\User\Desktop\Parser\data\product_library.xlsx')  # наша база товаров личный ПК
except:
    df_product_library = pd.read_excel(r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
                                       r'\5. Мониторинги цен конкурентов\Parser\build_windows\product_library.xlsx')
df_file_for_monitoring2 = df_file_for_monitoring.merge(df_product_library, on=["Код Товара"])  # объединяем файл поиска с нашей базой товаров
df_file_for_monitoring2 = df_file_for_monitoring2.fillna('Нет')
try:
    df_competitor_library = pd.read_excel(
        r'C:\Users\Artem\Desktop\Parser\data\competitor_library.xlsx')  # наша база товаров личный ПК
#        r'C:\Users\User\Desktop\Parser\data\competitor_library.xlsx')  # наша база товаров личный ПК
except:
    df_competitor_library = pd.read_excel(
        r'V:\SHOP_HOME\Common\УПРАВЛЕНИЕ ПРОДАЖАМИ\УПРАВЛЕНИЕ ФОРМАТАМИ\ЦЕНООБРАЗОВАНИЕ'
        r'\5. Мониторинги цен конкурентов\Parser\build_windows\competitor_library.xlsx')  # наша база товаров ПК УР

if competitor_for_monitoring == 'Fix-price':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_fix-price'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)  # количество строк в нем
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_fix-price']  # адрес страницы, с которой будет поступать информация
        kod_tovara = df_file_for_monitoring3.loc[
            i3, 'Код Товара']  # адрес страницы, с которой будет поступать информация
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        browser.get(url)  # Переход на страницу поиска товара
        time.sleep(8)  # время ожидания в секундах
        try:
            price_rub_shop = browser.find_element_by_xpath('//*[@id="__layout"]/div/div/div[3]/div/div/div/div/'
                                                           'div/div[2]/div[2]/div').text.replace(' ₽', '')
        except:
            price_rub_shop = 'Нет в наличии'

        if len(price_rub_shop) > 0:
            discount_percentage = (int(price_UR) - int(price_rub_shop))/int(price_UR)
            if discount_percentage > 0.5:
                price_rub_shop = ''

        new_row = {'Конкурент': 'Fix-price',
                   'Код Товара': kod_tovara,
                   'Категория': kategotya,
                   'Наименование Товара': tovar_name,
                   'Цена УР на дату мониторинга, руб.': price_UR,
                   'Цена, руб.': price_rub_shop}

        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
        df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                    regex=True)  # удаляем лишние пробелы и переходы строк
        i3 += 1
        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

    browser.close()
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Галамарт':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Spar', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_spar'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    Spar_url_base = 'https://myspar.ru/'
    browser.get(Spar_url_base)
    time.sleep(2)

    for i in range(0, i22):
        SPAR_CITY = ("{:03.0f}".format(df_file_competitor.loc[i33, 'SPAR_CITY']))
        browser.add_cookie({"name": "BITRIX_SM_SPAR_TP_CITY", "value": SPAR_CITY})
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_spar']
                browser.get(url)
                time.sleep(random.randint(3, 5))
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1
                city = browser.find_element_by_xpath('/html/body/header/div[1]/div/div/div[2]/span').text.replace(' ', '')
                try:
                    check_action = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[1]/'
                                                             'div[1]/div[2]/span').text.replace(' ', '')
                except:
                    check_action = ''
                if len(check_action) > 1:
                    price_rub_action = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[2]/div[2]/'
                                                            'div/div[1]/span/span[2]/span[1]').text.replace('90 ', '')
                    price_rub_shop = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[2]/'
                                                        'div[2]/div/div[1]/span/span[1]/span[1]').text.replace(' ', '')

                else:
                    price_rub_shop = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[2]/div[2]/'
                                                               'div/div[1]/span/span/span[1]').text.replace('90 ', '')
                    if len(price_rub_shop) > 3:
                        price_rub_shop = price_rub_shop.replace('00 ', '')
                    price_rub_action = ''

                price_rub_shop = price_rub_shop.replace(' ', '')
                price_rub_action = price_rub_action.replace(' ', '')

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': competitor,
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Подружка':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_подружка'].str.contains('http')]\
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)  # количество строк в нем
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_подружка']
        response = requests.get(url, timeout=100)
        time.sleep(random.randint(2, 4))
        response.encoding = 'utf-8'
        headers = response.headers
        soup = BeautifulSoup(response.text, 'lxml')
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        i3 += 1
        try:
            price_rub_shop = soup.find('span', class_='price__item price__item--old').text.replace('р.', '')
            price_rub_action = soup.find('span', class_='price__item price__item--current').text.replace('р.', '')
        except:
            try:
                price_rub_shop = soup.find('span', class_='price__item price__item--current').text.replace('р.', '')
                price_rub_action = ''
            except:
                price_rub_shop = 'Нет в наличии'
                price_rub_action = ''

        price_rub_shop = price_rub_shop[1:-2]
        price_rub_action = price_rub_action[1:-2]

        if len(price_rub_action) > 0:
            discount_percentage = (int(price_rub_shop) - int(price_rub_action))/int(price_rub_shop)
            if discount_percentage > 0.6:
                price_rub_action = ''

        if len(price_rub_shop) > 0:
            discount_percentage = (int(price_UR) - int(price_rub_shop))/int(price_UR)
            if discount_percentage > 0.5:
                price_rub_shop = ''

        new_row = {'Конкурент': 'Подружка',
                   'Код Товара': kod_tovara,
                   'Категория': kategotya,
                   'Наименование Товара': tovar_name,
                   'Цена УР на дату мониторинга, руб.': price_UR,
                   'Цена, руб.': price_rub_shop,
                   'Акц цена, руб.': price_rub_action}

        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
#        df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
#        df_parser_result = df_parser_result.replace([r'\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Впрок':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[
        df_file_for_monitoring2['url_search_впрок'].str.contains('http')].drop_duplicates()
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)  # количество строк в нем
    i3 = 0

    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_search_впрок']
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        browser.get(url)
        time.sleep(5)
        try:
            browser.find_element_by_class_name('col-xs-10').click()
            time.sleep(5)
            try:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[5]/'
                                                               'span').text.replace('.', ',')
                price_rub_pril = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[7]/'
                                                               'div').text.replace('.', ',')
            except:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[1]/div[2]/div[4]/'
                                                               'span').text.replace('.', ',')
                price_rub_pril = ''

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Впрок',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Цена в приложении, руб.': price_rub_pril
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                        regex=True)  # удаляем лишние пробелы и переходы строк
            i3 += 1

        except:
            time.sleep(0)
            i3 += 1

        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

    browser.close()
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Ватсонс':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[
        df_file_for_monitoring2['url_search_ватсонс'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)  # количество строк в нем
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_search_ватсонс']
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
        response = requests.get(url, headers=headers, timeout=15)
        time.sleep(random.randint(2, 10))  # время ожидания в секундах
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        i3 += 1
        try:
            price_rub_shop = soup.find('div', class_='product-tile__price product-tile__'
                                                     'price--old js-variant-old-price').text.replace(' руб', '')
            price_rub_action = soup.find('div', class_='product-tile__price product-tile__'
                                                       'price--discounted js-variant-price').text.replace(' руб', '')
        except:
            try:
                price_rub_shop = soup.find('div', class_='product-tile__price product-tile__'
                                                         'price--original js-variant-price').text.replace(' руб', '')
                price_rub_action = ''
            except:
                price_rub_shop = 'Нет в наличии'
                price_rub_action = ''

        if len(price_rub_shop) > 0:
            discount_percentage = (int(price_UR) - int(price_rub_shop))/int(price_UR)
            if discount_percentage > 0.5:
                price_rub_shop = ''

        new_row = {'Конкурент': 'Ватсонс',
                   'Код Товара': kod_tovara,
                   'Категория': kategotya,
                   'Наименование Товара': tovar_name,
                   'Цена УР на дату мониторинга, руб.': price_UR,
                   'Цена, руб.': price_rub_shop,
                   'Акц цена, руб.': price_rub_action}

        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
        df_parser_result = df_parser_result.replace([r'\nр.\n'], '',
                                                    regex=True)  # удаляем лишние пробелы и переходы строк
        df_parser_result = df_parser_result.replace([r'\n', r'            '], '',
                                                    regex=True)  # удаляем лишние пробелы и переходы строк
        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Скарлетт':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_scarlett'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_scarlett']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            price_rub_shop = soup.find('div', class_='product-detail__price-value').text.replace(' ', '')
            price_rub_action = ''

            price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Скарлетт',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Магнит Косметик_selenium':
    df_file_competitor = df_competitor_library.loc[df_competitor_library['competitor'] == 'Магнит Косметик']. \
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_мк'].str.contains('http')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)  # количество строк в нем

    for i in range(0, i22):
        FAVORITE_SHOP = df_file_competitor.loc[i33, 'FAVORITE_SHOP_for_MK']
        FAVORITE_SHOP_str = str(FAVORITE_SHOP)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        shop_city = str(df_file_competitor.loc[i33, 'city'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            options = webdriver.ChromeOptions()
            options.page_load_strategy = 'eager'  # ожидание загрузки стр полностью
            options.add_argument('--disable-blink-features=AutomationControlled')  # отключение опции режима веб драйвера
            options.add_argument('headless')  # для открытия headless-браузера
            prefs = {"profile.default_content_setting_values.geolocation": 2,  # отключаем геолокацию
                     "profile.default_content_setting_values.notifications": 2,  # отключаем показ уведомлений
                     "profile.managed_default_content_settings.images": 2}  # отключаем показ изображений
            options.add_experimental_option("prefs", prefs)  # применяем настройки профиля

            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument('--no-sandbox')
            options.add_argument('start-maximized')
            options.add_argument("--disable-infobars")
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_experimental_option("useAutomationExtension", False)
            options.add_experimental_option('excludeSwitches', ['enable-automation'])

            service = ChromeService(executable_path=CHROMEDRIVER_PATH)
            browser = webdriver.Chrome(options=options, service=service)
            MK_url_base = 'https://magnitcosmetic.ru/'
            browser.get(MK_url_base)
            time.sleep(2)

            url = df_file_for_monitoring3.loc[i3, 'url_мк']
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            price_rub_shop = ''
            check = ''
            browser.add_cookie({"name": "FAVORITE_SHOP", "value": FAVORITE_SHOP_str})
            browser.get(url)
            time.sleep(random.randint(6, 9))
            try:
                shop_city = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                          'header/div/div[1]/div[1]/div/div/a').text
                shop_address = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/'
                                                             'header/div/div[4]/div[1]/a/span').text
                price_rub_shop = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[1]/div/div[2]/div/'
                                                               'div[2]/div[1]/div[4]/div[1]/div/div[1]').text
                if len(price_rub_shop) > 0:
                    check = 'ok'
                elif check != 'ok':
                    time.sleep(random.randint(5, 7))
                    price_rub_shop = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[1]/div/div[2]'
                                                                   '/div/div[2]/div[1]/div[4]/div[1]/div/div[1]').text
                    if len(price_rub_shop) > 0:
                        check = 'ok'
                    else:
                        price_rub_shop = "Нет в наличии"

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Магнит Косметик',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': '',
                           'Цена по карте, руб.': '',
                           'Магазин УР': shop_ur,
                           'Город': shop_city,
                           'Адрес конкурента': shop_address,
                           'FAVORITE_SHOP': FAVORITE_SHOP_str}

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
#               df_parser_result = df_parser_result.concat(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'], '',
                                                        regex=True)
                i3 += 1
                browser.quit()

            except:
                i3 += 1

            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Магнит_selenium':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Magnit', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_magnit'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    Spar_url_base = 'https://dostavka.magnit.ru/'
    browser.get(Spar_url_base)
    time.sleep(2)

    for i in range(0, i22):
        shopId = ("{:06.0f}".format(df_file_competitor.loc[i33, 'shopId_for_Magnit']))
        browser.add_cookie({"name": "shopId", "value": shopId})
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_magnit']
                browser.get(url)
                time.sleep(random.randint(3, 5))
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1
                city = browser.find_element_by_xpath('//*[@id="js-header-id"]/div/div[2]/div/div[2]/div/div/'
                                                     'div/div/div[1]').text.replace(' ', '')
                try:
                    check_action = browser.find_element_by_xpath('//*[@id="js-main-id"]/div[2]/div/div'
                                                                 '[2]/section[1]/div[1]').text.replace(' ', '')
                except:
                    check_action = ''
                if len(check_action) > 1:
                    price_rub_action = browser.find_element_by_xpath('//*[@id="js-main-id"]/div[2]/div/div[2]/'
                                                'section[1]/div[2]/div[2]/div[1]/div[1]').text.replace('90 ', '')
                    price_rub_shop = browser.find_element_by_xpath('//*[@id="js-main-id"]/div[2]/'
                                            'div/div[2]/section[1]/div[2]/div[2]/div[1]/div[2]').text.replace(' ', '')

                else:
                    price_rub_shop = browser.find_element_by_xpath('//*[@id="js-main-id"]/div[2]/div/div[2]'
                                                    '/section[1]/div[2]/div[2]/div[1]/div').text.replace('90 ', '')
#                    if len(price_rub_shop) > 3:
#                        price_rub_shop = price_rub_shop.replace('00 ', '')
                    price_rub_action = ''

                price_rub_shop = price_rub_shop.replace(' ', '')
                price_rub_action = price_rub_action.replace(' ', '')

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': competitor,
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Магнит':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Magnit', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_magnit'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0

    for i in range(0, i22):
        shopId = ("{:06.0f}".format(df_file_competitor.loc[i33, 'shopId_for_Magnit']))
        cookies = dict(shopId = str(shopId))
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_magnit']
                headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                     ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
                response = requests.get(url, headers=headers, timeout=100, cookies=cookies)
#                response = requests.get(url, headers=headers, timeout=100)
                time.sleep(random.randint(1, 2))
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'lxml')
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1

                try:
                    price_rub_shop = soup.find('div', class_='m-price__old').text.replace(' ', '')
                    price_rub_action = soup.find('div', class_='m-price__current is-discounted').text.replace(' ', '')
                except:
                    price_rub_shop = soup.find('div', class_='m-price__current').text.replace(' ', '')
                    price_rub_action = ''

                price_rub_shop = price_rub_shop.replace('₽', '')

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Магнит',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Дикси':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Дикси', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_dixy'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0

    for i in range(0, i22):
#        price_id_for_dixy = ("{:02.0f}".format(df_file_competitor.loc[i33, 'price_id_for_dixy']))
        price_id_for_dixy = (df_file_competitor.loc[i33, 'price_id_for_dixy'])
        cookies = dict(price_id = str(price_id_for_dixy))
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_dixy']
                headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                     ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
                response = requests.get(url, headers=headers, timeout=100, cookies=cookies)
#                response = requests.get(url, headers=headers, timeout=100)
                time.sleep(random.randint(2, 3))
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'lxml')
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1

                price_rub_shop_full = soup.find('div', class_='card-prices').text.replace(r'.', ',')
                price_rub_shop = price_rub_shop_full.split('шт')[0]
                price_rub_action = price_rub_shop_full.split('шт')[1]

                # price_rub_action = soup.find('span', class_='price_value').text.replace(r'.', ',')
                # price_rub_shop_full = soup.find('div', class_='prices_block').text.replace(r'.', ',')
                # print(price_rub_action, price_rub_shop_full)
                # price_rub_shop_full = price_rub_shop_full.replace(r'В наличии ', '')
                # price_rub_shop_full = price_rub_shop_full.replace(r'шт ', '')
                # price_rub_shop_full = price_rub_shop_full.replace(r'/', '')
                # price_rub_shop_full = price_rub_shop_full.replace(r'\n', '')
                # price_rub_shop_full = price_rub_shop_full.replace(r'\xa0', '')
                # price_rub_shop = price_rub_shop_full.split(' ')[1]
                # if ',' not in price_rub_shop:
                #     price_rub_shop = price_rub_action
                #     price_rub_action = ''

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Дикси',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Spar':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Spar', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_spar'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    Spar_url_base = 'https://myspar.ru/'
    browser.get(Spar_url_base)
    time.sleep(3)

    for i in range(0, i22):
        SPAR_CITY = ("{:03.0f}".format(df_file_competitor.loc[i33, 'SPAR_CITY']))
        browser.add_cookie({"name": "BITRIX_SM_SPAR_TP_CITY", "value": SPAR_CITY})
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_spar']
                browser.get(url)
                time.sleep(random.randint(3, 5))
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1
                try:
                    check_action = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[1]/'
                                                                 'div[1]/div[1]/div').text.replace(' ', '')
                except:
                    check_action = ''
                if len(check_action) > 1:
                    price_rub_action = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[1]'
                                                    '/div[2]/div[2]/div/div[1]/span/span[2]').text.replace('90 ', '')
                    price_rub_shop = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[1]/div[2]'
                                                    '/div[2]/div/div[1]/span/span[1]/span[1]').text.replace(' ', '')

                else:
                    price_rub_shop = browser.find_element_by_xpath('/html/body/main/div/div/div/div[2]/div[1]'
                                                        '/div[2]/div[2]/div/div[1]/span/span').text.replace('90 ', '')
                    if len(price_rub_shop) > 3:
                        price_rub_shop = price_rub_shop.replace('00 ', '')
                    price_rub_action = ''

                price_rub_shop = price_rub_shop.replace(' ', '')
                price_rub_action = price_rub_action.replace(' ', '')
                price_rub_shop = price_rub_shop.replace('шт', '')
                price_rub_action = price_rub_action.replace('шт', '')

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': competitor,
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Spar-soup not work':
    df_file_competitor = df_competitor_library[df_competitor_library['competitor'].str.contains('Spar', na=False)].\
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_spar'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0

    for i in range(0, i22):
        SPAR_CITY = ("{:03.0f}".format(df_file_competitor.loc[i33, 'SPAR_CITY']))
        cookies = dict(BITRIX_SM_SPAR_TP_CITY = str(SPAR_CITY))
        time.sleep(1)
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        competitor = str(df_file_competitor.loc[i33, 'competitor'])
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            try:
                url = df_file_for_monitoring3.loc[i3, 'url_spar']
                headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                     ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
                response = requests.get(url, headers=headers, timeout=100, cookies=cookies)
#                response = requests.get(url, headers=headers, timeout=100)
                time.sleep(random.randint(2, 3))
                response.encoding = 'utf-8'
                soup = BeautifulSoup(response.text, 'lxml')
                kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
                kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
                tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
                price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
                i3 += 1

                try:
                    price_rub_shop = soup.find('span', class_='prices__old').text
                    price_rub_action = soup.find('span', class_='prices__cur js-item-price').text
                except:
                    price_rub_shop = soup.find('span', class_='prices__cur js-item-price').text
                    price_rub_action = ''

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': competitor,
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Магазин УР': shop_ur.replace('.0', '')
                           }

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
#                df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
#                df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

            except:
                i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Бахетле':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_bahetle'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_bahetle']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                action = soup.find('div', class_='flag-item special_offer').text
            except:
                action = ''
            if action == 'Акция':
                price_rub_shop = ''
                price_rub_action = soup.find('div', class_='product-item-detail-price').text.replace(' ', '')
                price_rub_action = price_rub_action.replace(r'\nр.\n', '')
                price_rub_action = price_rub_action.replace(r'\n', '')
                price_rub_action = price_rub_action.replace('\xa0', '')
                price_rub_action = price_rub_action.replace(r'₽', '')
                price_rub_action = price_rub_action.replace(r'1шт', '')
                price_rub_action = price_rub_action.replace(r'/', '')

                if int(price_rub_action) > 2000:
                    if len(price_rub_action) == 14:
                        price_rub_action = price_rub_action[:4]
                    elif len(price_rub_action) == 15:
                        price_rub_action = price_rub_action[:5]
                    else:
                        price_rub_action = price_rub_action[:6]
                else:
                    price_rub_action = price_rub_action
            else:
                price_rub_shop = soup.find('div', class_='product-item-detail-price').text.replace(' ', '')
                price_rub_action = ''
                price_rub_shop = price_rub_shop.replace(r'\nр.\n', '')
                price_rub_shop = price_rub_shop.replace(r'\n', '')
                price_rub_shop = price_rub_shop.replace('\xa0', '')
                price_rub_shop = price_rub_shop.replace(r'₽', '')
                price_rub_shop = price_rub_shop.replace(r'1шт', '')
                price_rub_shop = price_rub_shop.replace(r'/', '')

                if int(price_rub_shop) > 2000:
                    if len(price_rub_shop) == 14:
                        price_rub_shop = price_rub_shop[:4]
                    elif len(price_rub_shop) == 15:
                        price_rub_shop = price_rub_shop[:5]
                    else:
                        price_rub_shop = price_rub_shop[:6]
                else:
                    price_rub_shop = price_rub_shop

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Бахетле',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Полушка':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_polushka'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_polushka']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                action = soup.find('div', class_='flag-item specialoffer').text
            except:
                action = ''
            if action == 'Акция':
                price_rub_shop = soup.find(
                    'div', class_='product-detail-price-item product-item-detail-price-old').text.replace(' ', '')
                price_rub_action = soup.find(
                    'div', class_='product-detail-price-item product-item-detail-price-current').text.replace(' ', '')
            else:
                price_rub_shop = soup.find(
                    'div', class_='product-detail-price-item product-item-detail-price-current').text.replace(' ', '')
                price_rub_action = ''

            price_rub_shop = price_rub_shop.replace('.', ',')
            price_rub_action = price_rub_action.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Полушка',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Караван':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_karavan'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_karavan']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            price_rub_shop = soup.find('span', class_='ty-price-num').text.replace(' ', '')
            price_rub_action = ''
            price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Караван',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Макси':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_maxi_arhangelsk'].str \
        .contains('ht')].drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_maxi_arhangelsk']
            browser.get(url)
            time.sleep(3)
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            action = ''
            action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                           'p/span[1]').text.replace(' ', '')
            if search('%', action):
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                           'p/span[1]').text.replace(' ', '')
                price_rub_shop = price_rub_shop.split('-')[0]
                price_rub_action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                           'div/p/b').text.replace(' ', '')
                price_rub_shop = price_rub_shop.replace('.', ',')
                price_rub_action = price_rub_action.replace('.', ',')
            else:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                           'div/p/b').text.replace(' ', '')
                price_rub_action = ''
                price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Макси',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action,
                       'Магазин УР': 1811
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1


    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_maxi_vologda'].str \
        .contains('ht')].drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_maxi_vologda']
            browser.get(url)
            time.sleep(3)
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            action = ''
            action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                   'p/span[1]').text.replace(' ', '')
            if search('%', action):
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                               'p/span[1]').text.replace(' ', '')
                price_rub_shop = price_rub_shop.split('-')[0]
                price_rub_action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                                 'div/p/b').text.replace(' ', '')
                price_rub_shop = price_rub_shop.replace('.', ',')
                price_rub_action = price_rub_action.replace('.', ',')
            else:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                               'div/p/b').text.replace(' ', '')
                price_rub_action = ''
                price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Макси',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action,
                       'Магазин УР': 963
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1

    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_maxi_cherepovec'].str \
        .contains('ht')].drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_maxi_cherepovec']
            browser.get(url)
            time.sleep(3)
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            action = ''
            action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                   'p/span[1]').text.replace(' ', '')
            if search('%', action):
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                               'p/span[1]').text.replace(' ', '')
                price_rub_shop = price_rub_shop.split('-')[0]
                price_rub_action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                                 'div/p/b').text.replace(' ', '')
                price_rub_shop = price_rub_shop.replace('.', ',')
                price_rub_action = price_rub_action.replace('.', ',')
            else:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                               'div/p/b').text.replace(' ', '')
                price_rub_action = ''
                price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Макси',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action,
                       'Магазин УР': 1789
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1

    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_maxi_yaroslavl'].str \
        .contains('ht')].drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_maxi_yaroslavl']
            browser.get(url)
            time.sleep(3)
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            action = ''
            action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                   'p/span[1]').text.replace(' ', '')
            if search('%', action):
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/div/'
                                                               'p/span[1]').text.replace(' ', '')
                price_rub_shop = price_rub_shop.split('-')[0]
                price_rub_action = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                                 'div/p/b').text.replace(' ', '')
                price_rub_shop = price_rub_shop.replace('.', ',')
                price_rub_action = price_rub_action.replace('.', ',')
            else:
                price_rub_shop = browser.find_element_by_xpath('//*[@id="modal-root"]/div/div/div[2]/'
                                                               'div/p/b').text.replace(' ', '')
                price_rub_action = ''
                price_rub_shop = price_rub_shop.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Макси',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action,
                       'Магазин УР': 493
                       }

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽', r'1шт', r'/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Магнолия':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_magnolia'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_magnolia']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                price = soup.find('div', class_='cost prices detail')
                price = str(price).count('шт')
                if price == 2:
                    price_rub_shop = soup.find('div', class_='cost prices detail') \
                            .text.split('шт', 1)[0].lstrip().replace('Р/шт', '')
                    price_rub_action = soup.find('div', class_='cost prices detail') \
                            .text.split('шт')[1].replace('Р/шт', '')
                if price == 1:
                    price_rub_shop = soup.find('div', class_='cost prices detail').text.replace('Р/шт', '')
                    price_rub_action = ''
                if price == 0:
                    price = soup.find('div', class_='cost prices detail')
                    price = str(price).count('кг')
                    if price == 2:
                        price_rub_shop = soup.find('div', class_='cost prices detail') \
                            .text.split('кг', 1)[0].lstrip().replace('Р/шт', '')
                        price_rub_action = soup.find('div', class_='cost prices detail') \
                            .text.split('кг')[1].replace('Р/шт', '')
                    if price == 1:
                        price_rub_shop = soup.find('div', class_='cost prices detail').text.replace('Р/кг', '')
                        price_rub_action = ''

            except:
                price_rub_shop = 'Нет в наличии'
                price_rub_action = ''

            price_rub_shop = price_rub_shop.replace('.', ',')
            price_rub_action = price_rub_action.replace('.', ',')
            price_rub_action = price_rub_action.replace(' ', '')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Магнолия',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'Старая',
                                                     r'Розничнаяцена', r'Р/'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
        except:
            i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Семишагофф':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_semishagoff'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_semishagoff']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                price_rub_shop = soup.find('div', class_='tovar-price__last').text.replace(' ', '')
                price_rub_action = soup.find('div', class_='tovar-price__cur').text.replace(' ', '')
            except:
                price_rub_shop = soup.find('div', class_='tovar-price__cur').text.replace(' ', '')
                price_rub_action = ''

            price_rub_shop = price_rub_shop.replace('.', ',')
            price_rub_action = price_rub_action.replace('.', ',')

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Семишагофф',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', '\xa0', r'₽'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')

        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Оптима':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_оптима'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_оптима']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(3, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                price = soup.find('div', class_='description__right-buy')
                price = str(price).count('₽')
                if price == 2:
                    price_rub_shop = soup.find('div', class_='description__right-buy') \
                            .text.split('₽', 1)[1].lstrip().replace('Купить', '')
                    price_rub_action = soup.find('div', class_='description__right-buy') \
                            .text.split('₽')[0].replace('Купить', '')
                if price == 1:
                    price_rub_shop = soup.find('div', class_='description__right-buy').text.replace('Купить', '')
                    price_rub_action = ''
            except:
                price_rub_shop = 'Нет в наличии'
                price_rub_action = ''

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Оптима',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', r' ₽', '\xa0', 'Цена: ', r'₽',
                                                     r'Сообщить о поступлении', r'Распродано'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
        except:
            i3 += 1
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'РивГош':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_ривгош'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_ривгош']
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
        response = requests.get(url, headers=headers, timeout=100)
        time.sleep(random.randint(2, 4))
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        i3 += 1

        try:
            price_rub_shop = soup.find('s', class_='price pl-5 pr-0 sm:pr-5 sm:pl-0').text.replace('\xa0₽', '')
            price_rub_action = soup.find('span', class_='price first').text.replace('\xa0₽', '')
            price_rub_card = ''
        except:
            try:
                price_rub_shop = soup.find('product-price', class_='my-5 sm:my-0 lg:my-5') \
                    .text.split(' Полная цена ', 1)[1].lstrip().replace('\xa0₽', '')
                price_rub_card = soup.find('div', class_='price first').text.replace('\xa0₽', '')
                price_rub_action = ''
            except:
                price_rub_shop = 'Нет в наличии'
                price_rub_card = ''
                price_rub_action = ''

        if len(price_rub_shop) > 0:
            discount_percentage = (int(price_UR) - int(price_rub_shop))/int(price_UR)
            if discount_percentage > 0.5:
                price_rub_shop = ''

        new_row = {'Конкурент': 'Ривгош',
                   'Код Товара': kod_tovara,
                   'Категория': kategotya,
                   'Наименование Товара': tovar_name,
                   'Цена УР на дату мониторинга, руб.': price_UR,
                   'Цена, руб.': price_rub_shop,
                   'Акц цена, руб.': price_rub_action,
                   'Цена по карте, руб.': price_rub_card}

        df_parser_result = df_parser_result.append(new_row, ignore_index=True)
        df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
        df_parser_result = df_parser_result.replace([r'\n', r' ₽'], '', regex=True)
        df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Золотое яблоко':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_goldapple'].str.contains('ht')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_goldapple']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(3, 4))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1
            try:
                price_rub_shop = soup.find('span', class_='old-price').text.replace('\xa0₽', '')
                price_rub_action = soup.find('span', class_='special-price').text.replace('\xa0₽', '')
                price_rub_card = soup.find('span', class_='best-loyalty-price').text.replace('\xa0₽', '')
            except:
                try:
                    price_rub_shop = soup.find('span', class_='old-price').text.replace('\xa0₽', '')
                    price_rub_card = soup.find('div', class_='price first').text.replace('\xa0₽', '')
                    price_rub_action = ''
                except:
                    price_rub_shop = 'Нет в наличии'
                    price_rub_card = ''
                    price_rub_action = ''

            price_rub_shop = price_rub_shop.replace(' ', '')
            price_rub_action = price_rub_action.replace(' ', '')

            if price_rub_shop == price_rub_action:
                price_rub_action = ''

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Золотое яблоко',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Акц цена, руб.': price_rub_action,
                       'Цена по карте, руб.': price_rub_card}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)
            df_parser_result = df_parser_result.replace([r'\n', r' ₽', '\xa0'], '', regex=True)
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
        except:
            i3 += 0
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Лэтуаль':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_лэтуаль_json'].str
         .contains('ht')].drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring3.index)
    i3 = 0
    for i in range(0, i2):
        try:
            url = df_file_for_monitoring3.loc[i3, 'url_лэтуаль_json']
            response = requests.get(url, timeout=90)
            time.sleep(random.randint(2, 4))
            letu_json = response.json()
            try:
                action = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]['result']
                                    ['skuList'][0]['priceWithMaxDCard']['adjustments'][0]['pricingModel'])
                if 'карте' in action:
                    price_rub_card = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]
                                        ['result']['skuList'][0]['priceWithMaxDCard']['amount'])
                    price_rub_action = ''
                else:
                    price_rub_action = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]
                                        ['result']['skuList'][0]['priceWithMaxDCard']['amount'])
                    price_rub_card = ''
            except:
                price_rub_action = ''
                price_rub_card = ''

            price_rub_shop = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]['skuList'][0]
                                 ['priceWithMaxDCard']['rawTotalPrice'])
            price_rub_action = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]['skuList'][0]
                                 ['priceWithMaxDCard']['amount'])
#            price_rub_action = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]
#                                ['result']['skuList'][0]['priceWithMaxDCard']['amount'])
#            price_rub_shop = (letu_json['contents'][0]['mainContent'][0]['contents'][0]['productContent'][0]['result']
#                                        ['skuList'][0]['priceWithMaxDCard']['rawTotalPrice'])
            if price_rub_action == price_rub_shop:
                price_rub_action = ''
            else:
                price_rub_action = price_rub_action
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            i3 += 1

            if len(price_rub_shop) > 0:
                discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                if discount_percentage > 0.5:
                    price_rub_shop = ''

            new_row = {'Конкурент': 'Лэтуаль',
                       'Код Товара': kod_tovara,
                       'Категория': kategotya,
                       'Наименование Товара': tovar_name,
                       'Цена УР на дату мониторинга, руб.': price_UR,
                       'Цена, руб.': price_rub_shop,
                       'Цена по карте, руб.': price_rub_card,
                       'Акц цена, руб.': price_rub_action}

            df_parser_result = df_parser_result.append(new_row, ignore_index=True)
            df_parser_result = df_parser_result.replace([r'\nр.\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
            df_parser_result = df_parser_result.replace([r'\n'], '', regex=True)  # удаляем лишние пробелы и переходы строк
            df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
        except:
            i3 += 1

    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)

elif competitor_for_monitoring == 'Лента - поиск по ШК':
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[
        df_file_for_monitoring2['url_лента_search'].str.contains('http')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)  # количество строк в нем
    i3 = 0

    for i in range(0, i2):
        url = df_file_for_monitoring3.loc[i3, 'url_лента_search']
        kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
        kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
        tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
        price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
        browser.get(url)
        time.sleep(random.randint(1, 2))
        browser.add_cookie({"name": "DeliveryOptions", "value": 'Pickup'})
        browser.add_cookie({"name": "CityCookie", "value": 'spb'})
        browser.add_cookie({"name": "Store", "value": '0007'})
        browser.get(url)
        time.sleep(random.randint(5, 8))
        try:
            browser.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[2]/div/div[2]/div/div[3]/div[2]/div[1]'
                '/div/div/div/a/div[1]/div[2]').click()
            time.sleep(random.randint(5, 8))
            url_lenta = browser.current_url
            try:
                shop_city = 'СПБ'
                shop_address = browser.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div/header/div[2]/div[1]/div/div[1]/div/div/div[2]').text
                action = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/main/article/div/div/div'
                                                       '/div[1]/div[2]/div[1]').text
                price_rub_shop = browser.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/div'
                    '[1]/div[2]/div[1]/div[1]/span[1]').text
                price_rub_action = ''
                price_rub_card = browser.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/'
                    'div[1]/div[2]/div[2]/div[1]/span[1]').text

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Лента',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Цена по карте, руб.': price_rub_card,
                           'Акция': action,
                           'Город': shop_city,
                           'Адрес конкурента': shop_address,
                           'Ссылка на товар': url_lenta}
                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'],
                                                            '', regex=True)  # удаляем лишние пробелы и переходы строк
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
                i3 += 1

            except:
                shop_city = 'Мурманск'
                shop_address = ''
                action = ''
                price_shop = ''
                price_rub_action = ''
                price_rub_card = ''

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Лента',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Цена по карте, руб.': price_rub_card,
                           'Акция': action,
                           'Город': shop_city,
                           'Адрес конкурента': shop_address,
                           'Ссылка на товар': url_lenta}

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'],
                                                            '', regex=True)  # удаляем лишние пробелы и переходы строк
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
                i3 += 1

        except:
            price_shop = "Нет в наличии"
            i3 += 1

    browser.quit()

    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Лента':
    df_file_competitor = df_competitor_library.loc[df_competitor_library['competitor'] == 'Лента']. \
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_лента'].str.contains('http')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)  # количество строк в нем

    for i in range(0, i22):
        Store = ("{:04.0f}".format(df_file_competitor.loc[i33, 'Store_for_lenta']))
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        shop_city = str(df_file_competitor.loc[i33, 'city'])
        CityCookie = str(df_file_competitor.loc[i33, 'Store_for_lenta'])
        url = 'https://lenta.com/'
        browser.get(url)
        time.sleep(5)
        browser.add_cookie({"name": "DeliveryOptions", "value": 'Pickup'})
        browser.add_cookie({"name": "CityCookie", "value": CityCookie})
        browser.add_cookie({"name": "Store", "value": Store})
        time.sleep(2)
        i33 += 1
        i3 = 0

        for i in range(0, i2):
            url = df_file_for_monitoring3.loc[i3, 'url_лента']
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            browser.get(url)
            time.sleep(random.randint(4, 6))
            try:
                shop_address = browser.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div/header/div[2]/div[1]/div/div[1]/div/div/div[2]').text
                action = ''
                action = browser.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[2]/div[1]').text
                if 1 < len(action) < 8:
                    price_rub_shop = browser.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/'
                        'div[1]/div[2]/div[1]/div[1]/span[1]').text
                    price_rub_action = browser.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/'
                        'div[1]/div[2]/div[2]/div[1]/span[1]').text
                    price_rub_card = ''
                else:
                    price_rub_shop = browser.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/div'
                        '[1]/div[2]/div[1]/div[1]/span[1]').text
                    price_rub_action = ''
                    price_rub_card = browser.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div/main/article/div/div/div/div[1]/div[3]/div[2]/div[1]/div/div/'
                        'div[1]/div[2]/div[2]/div[1]/span[1]').text

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Лента',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Цена по карте, руб.': price_rub_card,
                           'Магазин УР': shop_ur,
                           'Акция': action,
                           'Город': shop_city,
                           'Адрес конкурента': shop_address}

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'],
                                                            '', regex=True)  # удаляем лишние пробелы и переходы строк
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
                i3 += 1

            except:
                i3 += 1

    browser.quit()
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

elif competitor_for_monitoring == 'Лента2':
    df_file_competitor = df_competitor_library.loc[df_competitor_library['competitor'] == 'Лента']. \
        reset_index(drop=True)
    df_file_competitor['shop'] = df_file_competitor['shop'].apply(str)
    if all_region_for_monitoring == 'True':
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    elif 0 < len(number_UR_for_monitoring) < 5:
        df_file_competitor = df_file_competitor.loc[df_file_competitor['shop'] == number_UR_for_monitoring]
        df_file_competitor = df_file_competitor.loc[df_file_competitor['tag_parsing_priority'] == 1]. \
            reset_index(drop=True)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Выбранная связка конкурент - магазин УР отсутствует в базе", "Информация",
                                         0)
    i22 = len(df_file_competitor.index)  # количество строк в нем
    i33 = 0
    df_file_for_monitoring3 = df_file_for_monitoring2.loc[df_file_for_monitoring2['url_лента'].str.contains('http')] \
        .drop_duplicates()  # оставим фильтр только если есть ссылка в базе
    df_file_for_monitoring3 = df_file_for_monitoring3.reset_index(drop=True)
    i2 = len(df_file_for_monitoring2.index)  # количество строк в нем

    for i in range(0, i22):
        Store = ("{:04.0f}".format(df_file_competitor.loc[i33, 'Store_for_lenta']))
        shop_ur = str(df_file_competitor.loc[i33, 'shop'])
        shop_city = str(df_file_competitor.loc[i33, 'city'])
        CityCookie = str(df_file_competitor.loc[i33, 'Store_for_lenta'])
        url = 'https://lenta.com/'

#        browser.get(url)
#        time.sleep(5)
#        browser.add_cookie({"name": "DeliveryOptions", "value": 'Pickup'})
#        browser.add_cookie({"name": "CityCookie", "value": CityCookie})
#        browser.add_cookie({"name": "Store", "value": Store})
#        time.sleep(2)

        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                 ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
        response = requests.get(url, headers=headers, timeout=100)
        time.sleep(random.randint(1, 2))
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')

        i33 += 1
        i3 = 0

        for i in range(0, i2):
            url = df_file_for_monitoring3.loc[i3, 'url_лента']
            kod_tovara = df_file_for_monitoring3.loc[i3, 'Код Товара']
            kategotya = df_file_for_monitoring3.loc[i3, 'Категория']
            tovar_name = df_file_for_monitoring3.loc[i3, 'Наименование Товара']
            price_UR = df_file_for_monitoring3.loc[i3, 'Price_UR']
            headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                                     ' Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
            response = requests.get(url, headers=headers, timeout=100)
            time.sleep(random.randint(2, 3))
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'lxml')
            print(soup)
            try:
                shop_address = soup.find('span', class_='ty-price-num').text
                action = ''
                action = soup.find('span', class_='ty-price-num').text
                if 1 < len(action) < 8:
                    price_rub_shop = soup.find('span', class_='ty-price-num').text
                    price_rub_action = soup.find('span', class_='ty-price-num').text
                    price_rub_card = ''
                else:
                    price_rub_shop = soup.find('span', class_='ty-price-num').text
                    price_rub_action = ''
                    price_rub_card = soup.find('span', class_='ty-price-num').text

                if len(price_rub_shop) > 0:
                    discount_percentage = (int(price_UR) - int(price_rub_shop)) / int(price_UR)
                    if discount_percentage > 0.5:
                        price_rub_shop = ''

                new_row = {'Конкурент': 'Лента',
                           'Код Товара': kod_tovara,
                           'Категория': kategotya,
                           'Наименование Товара': tovar_name,
                           'Цена УР на дату мониторинга, руб.': price_UR,
                           'Цена, руб.': price_rub_shop,
                           'Акц цена, руб.': price_rub_action,
                           'Цена по карте, руб.': price_rub_card,
                           'Магазин УР': shop_ur,
                           'Акция': action,
                           'Город': shop_city,
                           'Адрес конкурента': shop_address}

                df_parser_result = df_parser_result.append(new_row, ignore_index=True)
                df_parser_result = df_parser_result.replace([r' руб,', r'\n', r'Цена в магазине от ', r'в приложении'],
                                                            '', regex=True)  # удаляем лишние пробелы и переходы строк
                df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
                i3 += 1

            except:
                i3 += 1

    browser.quit()
    df_parser_result.to_excel(file_for_monitoring, index=False, sheet_name='monitor')
    ctypes.windll.user32.MessageBoxW(0, "Парсер закончил работу", "Информация", 0)
    browser.close()

else:
    ctypes.windll.user32.MessageBoxW(0, "Остальные конкуренты находятся в разработке", "Информация", 0)
