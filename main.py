import argparse
import datetime
import os
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pprint import pprint

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
    uniq_categoris = set()
    for wine in excel_transform:
        if 'Категория' in wine:
            uniq_categoris.add(wine['Категория'])
    wine_categoris = defaultdict(list)
    for categori in uniq_categoris:
        for element in range(len(excel_transform)):
            if excel_transform[element]['Категория'] == categori:
                wine_categoris[categori].append(excel_transform[element]) 
    return wine_categoris


def main():
    load_dotenv()
    file_path = os.getenv('FILE_PATH', 'DEFAULT_FILE_PATH')
    parser = argparse.ArgumentParser(description='Путь до excel файла')
    parser.add_argument('--excel_path', default=file_path)
    args = parser.parse_args()
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        delta=calculate_year(),
        word_for_delta=selection_word(calculate_year()),
        wine_categori=parse_excel(args.excel_path)
    )
    pprint(parse_excel(args.excel_path))
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
