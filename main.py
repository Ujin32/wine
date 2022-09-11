import datetime
import os
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

from dotenv import load_dotenv

from jinja2 import Environment, FileSystemLoader, select_autoescape

import pandas


def selection_word(delta):
    mod = delta % 100
    if mod == 1:
        word = 'год'
        return word
    elif mod in range(2, 5):
        word = 'года'
        return word
    elif mod in range(11, 21):
        word = 'лет'
        return word
    else:
        word = 'лет'
        return word


def calculate_year():
    now = datetime.datetime.now()
    start_event = 1920
    event_year = now.year
    delta = event_year - start_event
    return delta


def open_excel(file_path):
    excel_data_df = pandas.read_excel(
        file_path,
        sheet_name='Лист1',
        na_values=' ',
        keep_default_na=False
    )
    excel_transform = excel_data_df.to_dict(orient='records')
    return excel_transform


def parse_excel(file_path):
    excel_transform = open_excel(file_path)
    uniq_categori = set()
    for wine in excel_transform:
        if 'Категория' in wine:
            uniq_categori.add(wine['Категория'])
    wine_categori = defaultdict(list)
    for element in uniq_categori:
        wine_categori[element] = []
    for element in range(len(excel_transform)):
        for categ in wine_categori.keys():
            if excel_transform[element]['Категория'] == categ:
                wine_categori[categ].append(excel_transform[element])
    return wine_categori


def main():
    load_dotenv()
    file_path = os.getenv('FILE_PATH', 'DEFAULT_FILE_PATH')
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        delta=calculate_year(),
        word_for_delta=selection_word(calculate_year()),
        wine_categori=parse_excel(file_path)
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
