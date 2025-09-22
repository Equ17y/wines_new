from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
from collections import defaultdict
import argparse


def load_excel_data(filename, sheet_name):
    return pandas.read_excel(
        filename,
        sheet_name=sheet_name,
        na_values='Nan',
        keep_default_na=False
    )


def group_drinks_by_category(drinks_data):
    drinks = defaultdict(list)
    for drink in drinks_data:
        category = drink['Категория']
        drinks[category].append(drink)
    return drinks


def get_drinks(filename, sheet_name):
    excel_data = load_excel_data(filename, sheet_name)
    drinks_data = excel_data.to_dict(orient='records')
    drinks_by_category = group_drinks_by_category(drinks_data)
    return dict(sorted(drinks_by_category.items()))


def calculate_winery_age(foundation_year):
    current_year = datetime.now().year
    return current_year - foundation_year


def get_year_suffix(years):

    last_digit = years % 10
    two_last_digits = years % 100

    if last_digit == 1 and two_last_digits != 11:
        return "год"
    elif 2 <= last_digit <= 4 and not (11 <= two_last_digits <= 19):
        return "года"
    else:
        return "лет"


def get_year_formatted(foundation_year):
    years = calculate_winery_age(foundation_year)
    suffix = get_year_suffix(years)
    return f"{years} {suffix}"


def get_template(path, template_name):
    env = Environment(
        loader=FileSystemLoader(path),
        autoescape=select_autoescape(['html', 'xml'])
    )

    return env.get_template(template_name)


def render_template(template, winery_age, drinks_by_category):
    return template.render(
        winery_age=winery_age,
        drinks_by_category=drinks_by_category
    )


def save_html(homepage, rendered_page):
    with open(homepage, 'w', encoding="utf8") as file:
        file.write(rendered_page)


def main():
    parser = argparse.ArgumentParser(description='Генератор винного сайта')
    parser = argparse.ArgumentParser(description='Конструктор винного сайта')
    parser.add_argument('--excel-file', default='wine3.xlsx')
    parser.add_argument('--sheet-name', default='Лист1')
    parser.add_argument('--template', default='template.html')
    parser.add_argument('--output', default='index.html')

    args = parser.parse_args()
    foundation_year = 1920

    drinks_by_category = get_drinks(args.excel_file, args.sheet_name)
    winery_age = get_year_formatted(foundation_year)
    template = get_template('.', args.template)
    rendered_page = render_template(template, winery_age, drinks_by_category)
    save_html(args.output, rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __name__ == '__main__':
    main()