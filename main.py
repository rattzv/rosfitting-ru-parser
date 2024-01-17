import os

import openpyxl

from utils.utils import download_file, print_template, write_products_to_csv
from model.product import Product
from sheets import Flancy_09G2S, Flancy_10h17n13m2t, Flancy_12H18N10T_Ispolnenie, Flancy_12H18N10T_Kitaj
from sheets import Flancy_st_20, Krany_nerzh_Rossiya, Krepezh, Otvody_12H18N10T, Perekhody_12H18N10T
from sheets import Trojniki_12H18N10T, Zaglushki_12H18N10T, Flancy_12H18N10T_Rossiya_Lite


os.environ['PROJECT_ROOT'] = os.path.dirname(os.path.abspath(__file__))


def start():
    products: [Product] = []

    if not download_file():
        print_template("Stop. File download error!")
        return False

    workbook = openpyxl.load_workbook('price.xlsx')

    products.extend(Flancy_09G2S.parser(workbook))
    products.extend(Flancy_10h17n13m2t.parser(workbook))
    products.extend(Flancy_12H18N10T_Ispolnenie.parser(workbook))
    products.extend(Flancy_12H18N10T_Kitaj.parser(workbook))
    products.extend(Flancy_12H18N10T_Rossiya_Lite.parser(workbook))
    products.extend(Flancy_st_20.parser(workbook))
    products.extend(Krany_nerzh_Rossiya.parser(workbook))
    products.extend(Krepezh.parser(workbook))
    products.extend(Otvody_12H18N10T.parser(workbook))
    products.extend(Perekhody_12H18N10T.parser(workbook))
    products.extend(Trojniki_12H18N10T.parser(workbook))
    products.extend(Zaglushki_12H18N10T.parser(workbook))

    print(print_template(f"Done! Found {len(products)} products."))
    print(print_template("Save to file..."))

    write_products_to_csv(products)

    print(print_template("Completed."))


if __name__ == '__main__':
    start()
