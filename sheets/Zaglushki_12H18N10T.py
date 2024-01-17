from model.product import Product


def parser(workbook):
    sheet_name = "Заглушки 12Х18Н10Т"

    products = []
    count_list = [6, 10, 16, 25, 40, 63, 160]

    sheet = workbook[sheet_name]

    name = f"{sheet['A1'].value}, {sheet['A2'].value}"
    for index in [1, 3]:
        for row in sheet.iter_rows(min_row=6, max_row=20, min_col=index, values_only=True):
            type_size = row[0]
            price = row[1]

            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['F2'].value
    for row in sheet.iter_rows(min_row=6, max_row=24, min_col=6, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:5]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)

    return products


