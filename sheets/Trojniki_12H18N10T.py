from model.product import Product


def parser(workbook):
    sheet_name = "Тройники 12Х18Н10Т"

    products = []

    sheet = workbook[sheet_name]

    name = sheet['A2'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=6, max_row=28, min_col=index, values_only=True):
            type_size = row[0]
            price = row[1]

            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A32'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=36, max_row=48, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]

            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A51'].value
    for index in [1, 3]:
        for row in sheet.iter_rows(min_row=54, max_row=59, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]

            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)
    return products
