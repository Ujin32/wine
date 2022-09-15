import argparse
import datetime
import os
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

from dotenv import load_dotenv

from jinja2 import Environment, FileSystemLoader, select_autoescape

import pandas


def choosing_word(age_winery):
    division_remainder = age_winery % 100
    if division_remainder == 1:
        word = 'год'
        return word
    elif division_remainder in range(2, 5):
        word = 'года'
        return word
    elif division_remainder in range(11, 21):
        word = 'лет'
        return word
    else:
        word = 'лет'
        return word


def calculate_age_winery():
    year_foundation = 1920
    current_year = datetime.datetime.now().year
    age_winery = current_year - year_foundation
    return age_winery


def parse_excel(file_path):
    wines = pandas.read_excel(file_path, na_values=None, keep_default_na=False).to_dict(orient='records')
    wine_categories = defaultdict(list)
    for wine in wines:
        wine_categories[wine['Категория']].append(wine)
    return wine_categories


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
    age_winery = calculate_age_winery()
    rendered_page = template.render(
        age_winery=age_winery,
        word_for_age_winery=choosing_word(age_winery),
        wine_categories=parse_excel(args.excel_path)
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
