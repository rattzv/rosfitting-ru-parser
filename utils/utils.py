import os
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import csv


def print_template(message) -> str:
    current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    message = f"\r{current_date}: {message}"
    return message


def download_file():
    try:

        url = 'https://rosfitting.ru/price/'
        response = requests.get(url)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')
        xlsx_link = soup.find('a', text='скачать xlsx')
        if xlsx_link:
            xlsx_url = xlsx_link['href']
        else:
            return False

        response = requests.get(xlsx_url)

        with open('price.xlsx', 'wb') as file:
            file.write(response.content)
        return True
    except Exception as e:
        print("Unhandled Exception: ".format(e))


def check_reports_folder_exist() -> str:
    try:
        root_folder = os.environ.get('PROJECT_ROOT')
        reports_folder = os.path.join(root_folder, "reports")

        if not os.path.exists(reports_folder):
            os.makedirs(reports_folder)
        return reports_folder
    except Exception as e:
        print(f"Could not find or create reports folder: {e}")


def write_products_to_csv(products):
    reports_folder = check_reports_folder_exist()

    current_datetime = datetime.now()
    current_datetime_str = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")

    csv_file_path = os.path.join(reports_folder, f"products_{current_datetime_str}.csv")
    headers = list(vars(products[0]).keys())

    with open(csv_file_path, mode='w', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(headers)

        for product in products:
            values = [getattr(product, prop) for prop in headers]
            writer.writerow(values)
