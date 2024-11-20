import sys
import subprocess
import time
import random
from bs4 import BeautifulSoup
import re
import nodriver as uc
import sqlite3
import datetime
# from selenium import webdriver
# from selenium.webdriver import ActionChains
# from selenium.common import NoSuchElementException, ElementNotInteractableException
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
import win32com.client  # pywin32 это работа с Exel через COM напрямую
# from PyQt6 import QtCore, QtGui, QtWidgets
# from PyQt6.QtWidgets import QMessageBox


'''
Переустановил пакет websockets с версии 14.1 на версию 13 - иначе ошибка
'''


def get_all_product_from_sale_page(driver, produkti, item_description_box_l, sect_name):
    """функция считывает ВСЕ товары из полученного блока товаров
    в разделе типа Акции и возвращает список словарей с параметрами каждого товара"""
    prod = []  # sale_value, sale_percent, name, artikul, section
    art = {} #href name value artikul points section
    for item_description in item_description_box_l:

        item_artikul_field = item_description.find(class_='product-article_number').text.strip()
        regul = re.compile(r"[0-9]+")
        item_articul = regul.findall(item_artikul_field)[0]
        art['artikul'] = item_articul

        item_sale_percent = item_description.find(class_='flag_sale-count')
        if item_sale_percent != None:
            params = string_to_numbers(item_sale_percent.text)
            art['sale_percent'] = params[0]
        else:
            art['sale_percent'] = 0

        item_price = item_description.find(class_='product-price_current')
        if item_price != None:
            params = string_to_numbers(item_price.text)
            art['sale_value'] = params[0]
        else:
            art['sale_value'] = 0

        #название товара
        item_name_all = item_description.find_all('a', class_='product-name')
        tmp_txt = item_name_all[0].text
        item_name = ' '.join(tmp_txt.split())
        art['name'] = re.sub(r"\n", " ", item_name)
        art['section'] = sect_name
        #запоминаем название товара вдруг у него есть палитра цветов
        articul_group_name = art['name']
        #проверяем, есть ли выпадающий список с палитрой цветов у этого товара
        #если есть, то добавляем цвета и их артикулы
        item_variants = item_description.find(class_='amw-product-viewer-item__variants')
        # ===========================================================================================================
        # ===========================================================================================================
        #  Ниже - не переделано на nodriver
        # ===========================================================================================================
        #
        if item_variants != None:
            print(f"============= Есть товар с вариантами!!! склад {sect_name} код {item_articul}")
            #если есть надпись например "6 Цветов" - значит есть выпадающее меню с палитрой и кодами
            #наводим мышку на бокс этого товара чтобы скрипты обновили бокс
            #но наводиться там не на что нет уникальных меток или ссылок - все блоки товаров одинаковые
            #поэтому перебираем все блоки сравниваем артикул находим нужный и наводим курсор на артикул
            # block_all = driver.find_elements(By.CLASS_NAME, 'amw-product-viewer-item__ordering-number')
            # ss = len(block_all)
            # for block in block_all:
            #     reg = re.compile(r"[0-9]+")
            #     block_articul = regul.findall(block.text)[0]
            #     #если нужный блок с артикулом найден то выходим из цикла и используем этот блок для наведения
            #     if block_articul == item_articul:
            #         break
            #
            # ActionChains(driver).scroll_to_element(block).perform()
            # #time.sleep(1)
            # ActionChains(driver).move_to_element(block).perform()
            #
            # #перечитываем страницу чтобы захватить изменения
            # bs1 = BeautifulSoup(driver.page_source, 'html.parser')
            # color_list = bs1.find_all(class_='variants-select-component--variant-ditails')
            # #ss = len(color_list)
            # for color_item in color_list:
            #     color_name = color_item.find(class_="variants-select-component--variant-name").text
            #     color_name = re.sub(r"\n", " ", color_name).strip()
            #     color_kod = color_item.find(class_="variants-select-component--variant-desc").text
            #     color_kod = re.sub(r"\n", " ", color_kod).strip()
            #     #добавляем разные цвета к базовому названию
            #     art['name'] = articul_group_name + " " + color_name
            #     art['artikul'] = color_kod
            #     #проверка на повторение артикула в ценнике, в пазделах Каталога есть повторения
            #     find = False
            #     for produkt in produkti:
            #         if art['artikul'] == produkt['artikul']: find=True
            #
            #     if not find:
            #         prod.append(art.copy())
            #         print(f"Добавил товар в список: {art['artikul']} {art['name']}")
            #     else:
            #         print(f"Товар повторился в списке, отбрасываю: {art['artikul']} {art['name']}")
            # site_header = driver.find_element(By.CLASS_NAME, "site-header ")
            # ActionChains(driver).move_to_element(site_header).perform()
            #time.sleep(1)

        else: #если цветов у товара нет то добавляем только его
            # проверка на повторение артикула в ценнике, в разделах Каталога есть повторения
            find = False
            for produkt in produkti:
                if art['artikul'] == produkt['artikul']:
                    produkt['section'] = produkt['section'] + ", " + sect_name
                    find = True

            if not find:
                prod.append(art.copy())
                print(f"Добавил товар в список: {art['artikul']} {art['name']}")
            else:
                print(f"Товар повторился в списке, добавляю склад: {art['artikul']} {art['name']}")

    return prod  # ключи словаря артикула: sale_value, sale_percent, name, artikul, section

def sales_to_exel_telegram(produkti):
    '''получает список продуктов Каталога,
     Открывает Exel, создаёт новую книгу
     выводит на лист весь список товаров с баллами и ценой'''
    # ключи словаря артикула: sale_value, sale_percent, name, artikul, section
    com_object = win32com.client
    exel = com_object.Dispatch("Excel.Application")

    exel.Visible = True
    exel.Workbooks.Add()
    exel.Columns("A:A").Select()
    exel.Selection.ColumnWidth = 10
    #делаем столбец А текстовым
    exel.Selection.NumberFormat = "@"
    exel.Columns("B:B").Select()
    exel.Selection.ColumnWidth = 53
    exel.Columns("C:C").Select()
    exel.Selection.ColumnWidth = 9
    exel.Columns("D:D").Select()
    exel.Selection.ColumnWidth = 10
    # exel.Columns("E:E").Select()
    # exel.Selection.ColumnWidth = 12

    exel.Range("A1").Select()
    exel.Cells(1, 1).Value = f"Курс теньге: {0}"
    exel.Cells(1, 1).Font.Bold = True
    exel.Cells(2, 1).Value = f"Процент доставки: {0}"
    exel.Cells(2, 1).Font.Bold = True
    exel.Cells(3, 1).Value = f"Единый курс: {0}"
    exel.Cells(3, 1).Font.Bold = True
    exel.Cells(5, 1).Value = 'Артикул'
    exel.Cells(5, 2).Value = 'Наименование'
    exel.Cells(5, 3).Value = 'Скидка'
    exel.Cells(5, 4).Value = 'Теньге'
    # exel.Cells(5, 5).Value = 'Рубли'
    exel.ActiveSheet.Name = 'Акции'

    i = 5
    section_name = produkti[0]['section']
    # ключи словаря артикула: sale_value, sale_percent, name, artikul, section
    for produkt in sorted(produkti, key=lambda x: x['section']):
        i += 1
        exel.Cells(i, 1).Value = produkt['artikul']
        exel.Cells(i, 2).Value = produkt['name']
        exel.Cells(i, 3).Value = produkt['sale_percent']
        exel.Cells(i, 4).Value = produkt['sale_value']
        # if dostavka == 0: exel.Cells(i, 5).Value = round(produkt['value']/edin)
        # if edin == 0: exel.Cells(i, 5).Value = round(produkt['value']/tenge + produkt['value']/tenge*dostavka/100)
        exel.Cells(i, 6).Value = produkt['section']

def sales_to_bot_db(produkti): # sale_value, sale_percent, name, artikul, section
    kod_real = [] # список для запоминания кода
    # =================================================================================================
    with sqlite3.connect("dist\\AmVrnKotlBot_telegram_clients.db") as db:
        # ====================================================================================================
        cur = db.cursor()
        # обнуляем все скидки в базе
        sbros = cur.execute("UPDATE tblPriceList SET skidka=?, skidka_price=?", ("", "",))
        db.commit()
        i = 0 # счётчик для нумерации строк лога
        for produkt in produkti:
            kod = produkt['artikul']
            kod_real.append(kod)  # запоминаем код для сравнения потом
            articul_name = produkt['name']
            articul_sale_percent = produkt['sale_percent']
            articul_sale_price = produkt['sale_value']
            articul_sklad = produkt['section']

            # ищем код в таблице если он есть - обновляем его данные если его нет - добавляем в таблицу
            poisk = cur.execute("SELECT * FROM tblPriceList WHERE kod=?", (kod,))
            check_kod = poisk.fetchone()

            if check_kod is None:
                # db.execute("INSERT INTO tblPriceList (kod, name, points, price_npa, price_full, href, ended, aliases, "
                #            "name_short) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)", (kod, articul_name, articul_points,
                #                                                              articul_price_npa, articul_price_full,
                #                                                              articul_href, articul_ended, articul_aliases,
                #                                                              articul_name_short))
                # db.commit()
                print(f"{i}=============КОДА НЕТ В БАЗЕ {kod} {articul_name}")
            else:
                # "UPDATE users SET email='bob@newemail.com' WHERE name='Bob'"
                sale = cur.execute("UPDATE tblPriceList SET skidka=?, skidka_price=?, sklad=? \
                              WHERE kod=?", (articul_sale_percent, articul_sale_price, kod, articul_sklad,))
                db.commit()
                print(f"{i} +++++ {kod}, скидки кода обновлены")
            i += 1

        print("Закончил работу.")

    return
def string_to_numbers(str_am):
    ''' получает строку вида: '\n Промый ы\n -₸ 70,00\n  ины\n  ₸ 70,00\n Всег\n  -25,00
    8722 - это html-знак 'минус'
    находит все цифры в строке и возвращает список цифр float
    '''
    str_am = re.sub(r"\n", "", str_am)  # убираем все переводы строки
    str_am = re.sub(r"\s", "", str_am)  # убираем все пробелы
    # print(str_am) # 'Промежуточныйитогкорзины-₸70,00Итогкорзины₸70,00Всегобалловвкорзине-25,00'
    # str_am = 'Пр−₸70,00И₸70.00Всо−25,00'
    str_am = str_am + "Z"  # нужно для корректной работы цикла перебора если последний знак строки - цифра
    digits = []
    digit = ''
    prev_char = ''
    minus_flag = bool(0)
    for char in str_am:
        if re.fullmatch("\d", char) != None:  # если цифра - добавляем знак к числу
            digit += char
            prev_char = 'digit'
        elif char == f"{chr(8722)}":  # если html "минус" - ставим признак
            minus_flag = bool(1)
        elif char == "," and prev_char == 'digit':
            digit += '.'
        elif char == "." and prev_char == 'digit':
            digit += '.'
        elif prev_char == 'digit':
            number = float(digit)
            if minus_flag: number = 0 - number
            digits.append(number)
            digit = ''
            prev_char = ''
            minus_flag = bool(0)
        else:
            prev_char = ''
    return digits # возвращает список цифр формата float

async def main():
    # считывает раздел Акции
    '''Открывает сайт,  переходит в Акции
    устанавливает склад, считывает позиции,
    перебирает все склады.
    возвращает список словарей'''
    # ======= считываем позиции Каталога ==================

    driver = await uc.start()

    # открываем Сайт
    site_base_url = 'https://www.kz.amway.com'
    site_path = ''
    site_url = site_base_url + site_path

    tab = await driver.get(site_url)

    try:
        element = await tab.wait_for(text="ada-button-frame", timeout=40)
        if element:
            print("Главная страница загрузилась.", element.text)
        else:
            print("Главная страница не загрузилась!!!!!!!!")
    except Exception as e:
        print(f"Ошибка ожидания загрузки Главной страницы!: {e}")
        return

    # Подтверждаем геолокацию
    try:
        button_da = await tab.find(" Да", best_match=True)
        await button_da.mouse_move()
        await driver.wait(random.uniform(1, 3))
        await button_da.click()
        await driver.wait(2)
    except:
        print("Сбой подтверждения геолокации!")
        return

    # жмём на крестик убираем баннер куки внизу страницы
    try:
        disclamer_button = await tab.find("disclaimer__button", best_match=True)
        await disclamer_button.mouse_move()
        await driver.wait(random.uniform(1, 3))
        await disclamer_button.click()
    except:
        print("Сбой убирания баннера куки!")
        pass

    await driver.wait(random.uniform(1, 3))
    # находим и жмём меню Акции на главной странице
    user_vhod_link = await tab.find('/promo-page', best_match=True)
    await user_vhod_link.mouse_move()
    await driver.wait(random.uniform(1, 3))
    await user_vhod_link.click()

    try:
        element = await tab.wait_for(text="product-item simple-card", timeout=30)
        if element:
            print("Страница Акции загрузилась.", element.text)
        else:
            print("Страница Акции не загрузилась!!!!!!!!")
    except Exception as e:
        print(f"Ошибка ожидания загрузки Страница Акции!: {e}")
        return

    skladi = ["Актобе", "Астана", "Алматы"]
    skladi_list = ["Актюбинская обл.", "Астана", "Алматы"]

    # бесконечный цикл для считывания с задержками по времени
    while True:

        produkti = []

        for i in range(3):

            print(f"Выбираю город: {skladi[i]}")
            # находим и жмём геолокацию
            user_geo = await tab.find('user-location---cityName', best_match=True, timeout=5)
            await user_geo.mouse_move()
            await driver.wait(random.uniform(1, 3))
            await user_geo.click()

            # находим и жмём поле ввода города
            input_geo = await tab.find("searchControlInput---3aCkg_0", timeout=5)
            await input_geo.mouse_move()
            await driver.wait(random.uniform(1, 3))
            await input_geo.click()

            krestik = await tab.find("search-autocomplete__clear-btn", best_match=True, timeout=5)
            await krestik.mouse_move()
            await driver.wait(random.uniform(1, 3))
            await krestik.click()

            # Выбираю склад
            await driver.wait(random.uniform(1, 3))
            await input_geo.send_keys(skladi[i])
            await driver.wait(5)

            # находим и жмём город в списке
            city = await tab.find("search-autocomplete-results__item", timeout=5) # skladi_list[i])
            await city.mouse_move()
            await driver.wait(random.uniform(1, 3))
            await city.click()

            await driver.wait(5)

            # находим и жмём кнопку Выбрать
            vibor_but = await tab.find("selector---button---36kyL_0", timeout=5)
            await vibor_but.mouse_move()
            await driver.wait(random.uniform(1, 3))
            await vibor_but.click()

            # Жду загрузку страницы Акции
            try:
                element = await tab.wait_for(text="product-article_number", timeout=30)
                if element:
                    print("Страница Акции загрузилась.", element.text)
                else:
                    print("Страница Акции не загрузилась!!!!!!!!")
            except Exception as e:
                print(f"Ошибка ожидания загрузки Страница Акции!: {e}")
                return

            await driver.wait(10)

            # парсим страницу
            bs = BeautifulSoup(await tab.get_content(), 'html.parser')

            # словарь для хранения параметров товаров
            # articul = {}
            # список для хранения всех товаров

            section_name = skladi[i]

            # проверяем есть ли страницы, листаем если есть, читаем все товары
            pagination_info = bs.find('p', class_=re.compile('^pagination---paginationInfo.*'))

            if pagination_info == None:  # если страниц нет - считываем одну эту страницу
                item_description_boxes = bs.find_all(
                    class_='product-item simple-card')  # amw-product-viewer-item__description-box
                dd = len(item_description_boxes)
                print(f"Количество товаров в разделе: {len(item_description_boxes)}")
                if dd == 0:
                    print("Товаров нет на странице!")
                    await driver.wait(10)
                    continue
                ppp = get_all_product_from_sale_page(driver, produkti, item_description_boxes, section_name)

                produkti += ppp
                # return produkti
            else:  # если страницы есть то считываем, листаем, считываем
                print(f'======= Страница Акции, склад {section_name} - есть вторая страница!!!')
                # ss = pagination_info.text[9:10]
                # page_current = int(pagination_info.text[9:10])  # текущая страница
                # ss = pagination_info.text
                # page_count = int(pagination_info.text[14:15])  # всего страниц
                # page_number = 1
                # while page_number <= page_count:
                #     item_description_boxes = bs.find_all(
                #         class_='amw-product-viewer-item__inner')  # amw-product-viewer-item__inner
                #     dd = len(item_description_boxes)
                #     ppp = get_all_product_from_sale_page(driver, produkti, item_description_boxes, section_name)
                #     produkti += ppp
                #
                #     # input("Листай страницу и ввод:")
                #
                #     button_next_page = await tab.find("aqa-pagination-forward-button")
                #     await button_next_page.mouse_move()
                #     await driver.wait(2)
                #     await button_next_page.click()
                #
                #     site_header = await tab.find("site-header ")
                #     await site_header.mouse_move()
                #     await driver.wait(1)
                #
                #     bs = BeautifulSoup(await tab.get_content(), 'html.parser')
                #     page_number += 1

        # sales_to_exel_telegram(produkti)
        sales_to_bot_db(produkti)

        current_datetime = datetime.datetime.now()
        time_now = current_datetime.strftime("%H:%M:%S")
        date_now = current_datetime.strftime("%d/%m/%y")
        print(f"{date_now} {time_now} ======================== Отработал цикл ==================")

        await driver.wait(86400)  # 86400 секунд = 1 сутки

# Закрываем Хром
    # driver.close()
    # process.terminate()
    #    return produkti

if __name__ == '__main__':
    # since asyncio.run never worked (for me)
    uc.loop().run_until_complete(main())
    print("Программа завершилась!!!")
    time.sleep(86400)