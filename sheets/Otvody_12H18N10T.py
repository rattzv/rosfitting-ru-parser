from model.product import Product


def parser(workbook):
    sheet_name = "Отводы 12Х18Н10Т"

    products = []

    sheet = workbook[sheet_name]

    name = sheet['A2'].value
    for index in [1, 3, 5, 8, 10, 12]:
        name = name if index < 8 else sheet['H2'].value
        for row in sheet.iter_rows(min_row=6, max_row=39, min_col=index, values_only=True):
            type_size = row[0]
            price = row[1]

            price = price if price != "-" else None
            product = Product(name=name, price=price, type_size=type_size)
            products.append(product)
    return products
