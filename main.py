from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
from collections import defaultdict


def get_drinks(filename, table_sheet_name):
    excel_data_df = pandas.read_excel(
        filename,
        sheet_name=table_sheet_name,
        na_values='Nan',
        keep_default_na=False
    )
    excel_drinks = excel_data_df.to_dict(orient='records')
    drinks = defaultdict(list)
    for drink in excel_drinks:
        category = drink['Категория']
        drinks[category].append(drink)
    sorted_drinks = {key: value for key, value in sorted(drinks.items())}
    return sorted_drinks


def get_year_formatted(foundation_year):
    current_year = datetime.now().year
    experience_years = current_year - foundation_year

    last_digit = experience_years % 10
    two_last_digits = experience_years % 100

    if last_digit == 1 and two_last_digits != 11:
        suffix = "год"
    elif 2 <= last_digit <= 4 and not (11 <= two_last_digits <= 19):
        suffix = "года"
    else:
        suffix = "лет"

    return f"{experience_years} {suffix}"


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
    foundation_year = 1920
    filename = 'wine3.xlsx'
    sheet_name = 'Лист1'
    template_name = 'template.html'
    homepage_name = 'index.html'
    filepath = '.'

    drinks_by_category = get_drinks(filename, sheet_name)
    winery_age = get_year_formatted(foundation_year)

    template = get_template(filepath, template_name)
    rendered_page = render_template(template, winery_age, drinks_by_category)

    save_html(homepage_name, rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __name__ == '__main__':
    main()