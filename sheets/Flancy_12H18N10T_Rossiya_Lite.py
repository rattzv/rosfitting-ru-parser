from model.product import Product


def parser(workbook):
    sheet_name = "Фланцы12Х18Н10Т Россия Литье"

    products = []
    count_list = [6, 10, 16, 25, 40, 63, 160]

    sheet = workbook[sheet_name]

    name = f"{sheet['A1'].value } {sheet['A2'].value}"
    for row in sheet.iter_rows(min_row=5, max_row=22, values_only=True):
        diameter = row[0]
        prices = row[1:]
        for index, price in enumerate(prices[:4]):
            price = price if price != "-" else None
            product = Product(name=name, diameter=diameter, count=count_list[index], price=price)
            products.append(product)

    workbook.close()
    return products
