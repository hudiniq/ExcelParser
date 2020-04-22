import pandas as pd
from itertools import islice

df = pd.read_excel("sample.xlsx")

headers = []
orders_position = []
order_counter = 0
orders = []


#def name_headers():
for i in range(0, len(df.columns)):
    headers.append(str(i))

df.columns = headers

#def find_start():
for index, row in df.iterrows():
    if row[0] == "COD. ARTICOLO / ITEM CODE": 
        start = index + 1

#def find_orders():
for index, row in islice(df.iterrows(), start, None):
    try:
        int(row[0])
        orders_position.append(index)
    except ValueError:
        pass

#def fill_orders():
"""
Za vsako narocilo za katerega sem shranil pozicijo, gre cez vrstice
med njegovo pozicijo in pozicijo naslednjega (-3 ker zadnje 3 vrstice niso artikli),
zapise artikle v item ki ga zapise v list orderjev.
Ce pri zadnjem vrne IndexError, ko isce naslednji element, izvrsi kodo ki gleda konec dokumenta.
"""
for order in orders_position:
    try:
        for index, row in islice(df.iterrows(), order + 1, orders_position[orders_position.index(order) + 1] - 3):
            item = {
                "supplier_code" : df.iloc[order, 0],
                "lotto" : row[5],
                "opis" : row[12],
                "kolicina" : int(row[9] / 1000),
                "code" : None,
            }
            orders.append(item)
            print(item)
    except IndexError:
        for index, row in islice(df.iterrows(), order + 1, len(df) - 6):
            item = {
                "supplier_code" : df.iloc[order, 0],
                "lotto" : row[5],
                "opis" : row[12],
                "kolicina" : row[9],
                "code" : None,
            }
            orders.append(item)
            print(item)

# print(orders)

# print(df.iloc[start:])
# print(orders_position)
