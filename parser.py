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
for index in df.index:
    if df.iloc[index, 0] == "COD. ARTICOLO / ITEM CODE": 
        start = index + 1

#def find_orders():
for index in islice(df.index, start, None):
# for index in islice(df.iterrows(), start, None):
    try:
        int(df.iloc[index, 0])
        # int(row[0])
        for index2 in islice(df.index, index, None):
            if df.iloc[index2, 0] == "UN":
                order_end = index2
                break
        orders_position.append((index, order_end - 1))
    except ValueError:
        pass

# #def fill_orders():
for order, order_end in orders_position:

    code_row = df.iloc[order_end]

    for c in code_row:
        try:
            code = int(c)
        except ValueError:
            pass

    for index in islice(df.index, order + 1, order_end):
        item = {
            "supplier_code" : df.iloc[order, 0],
            "lotto" : df.iloc[index, 5],
            "opis" : df.iloc[index, 12],
            "kolicina" : int(df.iloc[index, 9] / 1000),
            "code" : code,
        }
        orders.append(item)


df_out = pd.DataFrame(orders)

df_out.to_excel("output.xlsx", index=False)
