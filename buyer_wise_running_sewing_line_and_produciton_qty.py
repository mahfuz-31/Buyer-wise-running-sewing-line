import pandas as pd # type: ignore

df = pd.read_csv('D:/1. Work/1. Daily/Buyer wise running sewing line/Book1.csv')

floor_names = ['JAL-1', 'JAL-2', 'JAL-3', 'FFL-1', 'FFL-2', 'FFL-3', 'JFL-1', 'JFL-2', 'JKL-1', 'JKL-2', 'JKL-3', 'JKL-4', 'JKL-5', 'DBL', 'MFL', 'MFL-1', 'MFL-2',
'MFL-3', 'MFL-4', 'FFL2-1', 'FFL2-2', 'FFL2-3', 'FFL2-4', 'FFL2-5', 'JKL-U2-1', 'JKL-U2-2', 'JKL-U2-3', 'JKL-U2-4', 'JKL-U2-5']

df = df[['Unit', 'Line', 'Order No', 'Total']]

buyers = []
for index, row in df.iterrows():
    found = False
    if row['Unit'] in floor_names:
        found = True
        # filter the buyer
        order = row['Order No']
        words = order.split('::')
        words2 = words[0].split()
        buyer = words2[1]
        buyers.append(buyer)
    if found is False:
        df.drop(index, inplace=True)

df['Buyers'] = buyers
df.drop(['Order No'], axis=1, inplace=True)

buyers = set(buyers)

from typing import Dict
production: Dict[str, int] = {}

total_production = 0
for index, row in df.iterrows():
    num = int(row['Total'])
    buyer = row['Buyers']
    if buyer in production:
        production[buyer] += num
    else:
        production[buyer] = num
    total_production += num

print(total_production)

production_df = pd.DataFrame()
production_df['Buyer'] = production.keys()
production_df['Production Qty.'] = production.values()

df = df.drop('Total', axis=1)

with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, sheet_name='Unit-Buyer-Line', index=False)
    production_df.to_excel(writer, sheet_name='Production', index=False)