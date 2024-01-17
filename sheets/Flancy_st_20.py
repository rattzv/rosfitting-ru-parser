from model.product import Product


def parser(workbook):
    sheet_name = "Фланцы ст.20"

    products = []
    count_list = [6, 10, 16, 25, 40, 63, 160]

    sheet = workbook[sheet_name]

    name = f"{sheet['A1'].value}, {sheet['A2'].value}"
    for row in sheet.iter_rows(min_row=5, max_row=30, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:4]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)

    name = f"{sheet['A1'].value}, {sheet['G2'].value}"
    for row in sheet.iter_rows(min_row=5, max_row=30, min_col=7, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:7]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)

    workbook.close()
    return products
