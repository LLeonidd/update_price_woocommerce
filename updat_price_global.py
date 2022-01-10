import pandas as pd

pd.set_option('display.max_rows', None)

class Price:
    def __init__(self, _file='price.xlsx'):
        self.file_path = _file
        self.file = pd.ExcelFile(self.file_path)
        self.settings = [
            {
                'name': 'Рулонные жалюзи',
                'name_sheet': 'рулонные ткани',
                'columns': {'name': 'НАИМЕНОВАНИЕ', 'price': 'Unnamed: 9', }
            },
            {
                'name': 'Вертикальные жалюзи',
                'name_sheet': 'Вертикальные ткани',
                'columns': {'name': 'НАИМЕНОВАНИЕ ', 'price': 'Unnamed: 6', }
            },
            {
                'name': 'Рулонные жалюзи Зебра',
                'name_sheet': 'Ткань зебра',
                'columns': {'name': 'НАИМЕНОВАНИЕ', 'price': 'Unnamed: 8', }
            },
        ]

    def get_dataframe_by_sheet(self, sheet):
        return self.file.parse(sheet)

    def get_all(self):
        df_summary = pd.DataFrame()
        for rule in self.settings:
            frame = self.get_dataframe_by_sheet(rule['name_sheet'])
            frame['type'] = rule['name']
            columns = [val for val in rule['columns'].values()]
            columns.append('type')
            df = frame[~frame[rule['columns']['name']].isnull()].filter(items=columns)

            for i, data in df.iterrows():
                # remove the space at the end of the line
                df.loc[i, rule['columns']['name']] = data[rule['columns']['name']].strip() \
                    .replace('  ', ' ') \
                    .replace('*', '') \
                    .replace(' Б/О', ' BLACK-OUT') \
                    .replace(' B/O', ' BLACK-OUT') \
                    .replace(' В/О', ' BLACK-OUT') \
                    .replace(' БО',' BO')\
                    .upper()
                # round price
                try:
                    df.loc[i, rule['columns']['price']] = round(data[rule['columns']['price']])
                except ValueError:
                    df.drop(index=i, inplace=True)


            #rename columns for concat
            df.rename(columns={
                rule['columns']['name']: 'name',
                rule['columns']['price']: 'price',
            }, inplace=True)

            df_summary = pd.concat([df, df_summary], ignore_index=True)

        return df_summary


class WooExport:
    def __init__(self, _file='export.csv'):
        self.file_path = _file
        self.file = pd.read_csv(self.file_path)

    def get_dataframe(self):
        return self.file



xl = Price()
ex = WooExport()

#print(xl.get_all())
#print(ex.get_dataframe())

df_export = ex.get_dataframe()
not_in_catalog = []

for i, f_price in xl.get_all().iterrows():
    _product = df_export[(df_export['Name'] == f_price['name']) & (df_export['Categories'] == f_price['type'])]
    if _product.count().ID > 0:
        df_export.loc[_product.index.values, 'Attribute 4 value(s)'] = f_price['price']
    else:
        not_in_catalog.append(f_price.values)
    pass


df_export.to_csv('woo_import.csv', index=False)
#print(df_export[['Name','Attribute 4 value(s)' ]])
for i in not_in_catalog: print(i)
print(not_in_catalog.__len__())