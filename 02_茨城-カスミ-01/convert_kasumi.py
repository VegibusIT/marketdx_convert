import pandas as pd
import openpyxl
import datetime
import copy
import os

original_xlsx_name = 'original'
original_xlsx = 'original.xlsx'
xlsx_file_start = f'{original_xlsx_name}_start.xlsx'
export_file_name = 'cart_final.xlsx'  # * 出力ファイル名

days_list = {
    '月曜': '3/13',
    '火曜': '3/14',
    '水曜': '3/15',
    '木曜': '3/16',
    '金曜': '3/17',
    '土曜': '3/18',
    '日曜': '3/19',
}

wb = openpyxl.load_workbook(f'{original_xlsx}', data_only=True)
ws = wb.worksheets[0]
ws.delete_rows(0, 6)
wb.save(f'{xlsx_file_start}')

df = pd.read_excel(f'./{xlsx_file_start}', header=0)
print('df: ', df)

df = df.fillna(0)
print('df: ', df)
df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']] = df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']].astype('int')
df = df.loc[~((df['月曜'] == 0) & (df['火曜'] == 0) & (df['水曜'] == 0) & (df['木曜'] == 0) & (df['金曜'] == 0) & (df['土曜'] == 0) & (df['日曜'] == 0))]
df = df[['ID', '品目（量目は目安です）', '出荷元（生産者）', '生産地', 'やさいバス数量', '商品入数', '単位', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']]
df_quantity = df['やさいバス数量']
print('df_quantity: ', df_quantity)
df = df.rename(columns={'品目（量目は目安です）': '商品名', '出荷元（生産者）': '生産者名', '生産地': '産地（都道府県）', 'やさいバス数量': '入数'})
df.insert(7, 'JANコード', '')
df.insert(8, '掲載単価', '')
df[['ID', '入数', '商品入数']] = df[['ID', '入数', '商品入数']].astype('int')
df['商品入数'] = df['商品入数'].apply(str)

df['規格'] = df['商品入数'].str.cat(df['単位'])

df = df.drop(['商品入数', '単位'], axis=1)
# * 最終列に空発列を追加
df[''] = ''
col = df.pop('規格')
df.insert(loc=5, column='規格', value=col)
print('df: ', df.head(10))
df = df.rename(columns=days_list)
df.to_excel(f'./{export_file_name}', index=False)

file = f'./{export_file_name}'
wb = openpyxl.load_workbook(file)
ws = wb['Sheet1']
ws.insert_cols(9, 2)
ws.insert_rows(0, 3)
wb.save(f'{file}')

os.remove(f'./{xlsx_file_start}')
