import pandas as pd
import openpyxl
import datetime
import copy

xlsx_file_name = 'asahimachi'
xlsx_file = f'{xlsx_file_name}.xlsx'
xlsx_file_start = f'{xlsx_file}_start.xlsx'
export_file_name = f'{xlsx_file}_final.xlsx'  # * 出力ファイル名

days_list = {
    '月曜': '3/6',
    '火曜': '3/7',
    '水曜': '3/8',
    '木曜': '3/9',
    '金曜': '3/10',
    '土曜': '3/11',
}

wb = openpyxl.load_workbook(xlsx_file, data_only=True)
ws = wb.worksheets[0]
ws.delete_rows(0, 6)
wb.save(xlsx_file_start)

df = pd.read_excel(f'./{xlsx_file_start}', header=0)

df = df.fillna(0)
df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜']] = df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜']].astype('int')
df = df.loc[~((df['月曜'] == 0) & (df['火曜'] == 0) & (df['水曜'] == 0) & (df['木曜'] == 0) & (df['金曜'] == 0) & (df['土曜'] == 0))]
df = df[['ID', '品目（量目は目安です）', '出荷元（生産者）', '生産地', 'ロット', '商品入数', '単位', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '数量']]
df = df.rename(columns={'品目（量目は目安です）': '商品名', '出荷元（生産者）': '生産者名', '生産地': '産地（都道府県）', 'ロット': '入数'})
df.insert(7, 'JANコード', '')
df.insert(8, '掲載単価', '')
df[['ID', '入数', '商品入数', '数量']] = df[['ID', '入数', '商品入数', '数量']].astype('int')
df['商品入数'] = df['商品入数'].apply(str)

# * 規格というカラムを作って、指定の場所に挿入
df['規格'] = df['商品入数'].str.cat(df['単位'])
df = df.drop(['商品入数', '単位'], axis=1)
col = df.pop('規格')
df.insert(loc=5, column='規格', value=col)

# * 数量カラムと入数カラムを入れ替え
df = df.drop('入数', axis=1)
col = df.pop('数量')
df.insert(loc=4, column='入数', value=col)
print('df: ', df.head(20))

df = df.rename(columns=days_list)
df.to_excel(f'./{export_file_name}', index=False)

file = f'./{export_file_name}'
wb = openpyxl.load_workbook(file)
ws = wb['Sheet1']
ws.insert_cols(9, 2)
ws.insert_cols(17, 1)
ws.insert_rows(0, 3)
wb.save(file)
