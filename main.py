import os
import time

import dotenv
import requests
import xml.etree.ElementTree as ET

import schedule
from openpyxl.workbook import Workbook
from bs4 import BeautifulSoup

import sheets

dotenv.load_dotenv()

xml_link = os.getenv('XML_LINK')

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br, zstd',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Cache-Control': 'no-cache',
    'Dnt': '1',
    'User-Agent':
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36'}


def download_xml(url):
    try:
        response = requests.get(url, headers=headers)
    except Exception as e:
        time.sleep(10)
        return download_xml(url)

    return response.text


def get_text_or_empty(element):
    if element is not None and element.text is not None:
        return element.text.replace('.00', '')
    return ''


def xml_to_xlsx(xml_data):
    root = ET.fromstring(xml_data)

    data = []
    data.append(
        ['Название (не учитывать)', 'Поставщик (не учитывать)', 'Цена',
         'Ссылка (не учитывать)',
         'Описание (не учитывать)', 'Ширина', 'Высота (не учитывать)', 'Глубина (не учитывать)', 'Тип',
         'Цвет (не учитывать)', 'Скорость доставки',
         ])

    for offer in root.findall('.//offer'):
        name = get_text_or_empty(offer.find('name'))
        vendor = get_text_or_empty(offer.find('vendor'))
        price = get_text_or_empty(offer.find('price'))
        link = get_text_or_empty(offer.find('url'))
        description = get_text_or_empty(offer.find('description'))
        width = get_text_or_empty(offer.find('param[@name="Ширина"]'))
        height = get_text_or_empty(offer.find('param[@name="Высота"]'))
        depth = get_text_or_empty(offer.find('param[@name="Глубина"]'))
        offer_type = get_text_or_empty(offer.find('param[@name="Тип"]'))
        color = get_text_or_empty(offer.find('param[@name="Цвет"]'))
        delivery_time = get_text_or_empty(offer.find('param[@name="Сроки поставки"]'))
        soup = BeautifulSoup(description, 'html.parser')
        description = soup.text
        data.append([name, vendor, price, link, description, width, height, depth, offer_type, color, delivery_time])

    return data


def save_xlsx(workbook, filename):
    workbook.save(filename)


def main():
    print('Update started')
    xml_data = download_xml(xml_link)
    print('Link downloaded')
    workbook = xml_to_xlsx(xml_data)
    print('Data getted')
    sheets.write_to_db(workbook)
    print('Sheet saved')

schedule.every().day.at("11:00").do(main)
schedule.every().day.at("23:00").do(main)

print('Script started!')
while True:
    schedule.run_pending()
    time.sleep(60)
