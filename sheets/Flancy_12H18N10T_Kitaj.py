from model.product import Product


def parser(workbook):
    sheet_name = "Фланцы12Х18Н10Т Китай"

    products = []
    count_list = [6, 10, 16, 25, 40, 63, 160]

    sheet = workbook[sheet_name]

    name = sheet['A3'].value
    for row in sheet.iter_rows(min_row=6, max_row=26, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:4]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)

    name = sheet['G3'].value
    for row in sheet.iter_rows(min_row=6, max_row=26, min_col=7, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:7]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)
    workbook.close()
    return products

