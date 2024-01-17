from model.product import Product


def parser(workbook):
    sheet_name = "Крепеж"

    products = []

    sheet = workbook[sheet_name]

    name = sheet['A1'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=3, max_row=12, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]
            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A13'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=15, max_row=21, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None or type_size == "-":
                continue

            price = row[1]
            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A22'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=24, max_row=26, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]
            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A27'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=29, max_row=31, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]
            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    name = sheet['A32'].value
    for index in [1, 3, 5]:
        for row in sheet.iter_rows(min_row=34, max_row=37, min_col=index, values_only=True):
            type_size = row[0]
            if type_size is None:
                continue

            price = row[1]
            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)

    workbook.close()
    return products
