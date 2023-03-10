import pandas as pd
import openpyxl
import datetime
import copy

export_file_name = 'midorino_final'  # * 出力ファイル名
days_list = {
    '月曜': '2/27',
    '火曜': '2/28',
    '水曜': '3/1',
    '木曜': '3/2',
    '金曜': '3/3',
    '土曜': '3/4',
    '日曜': '3/5',
}

df = pd.read_csv('./midorino_0227.csv', header=0, sep=',')
df = df.fillna(0)
df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']] = df[['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']].astype('int')
df = df.loc[~((df['月曜'] == 0 ) & (df['火曜'] == 0 ) & (df['水曜'] == 0 ) & (df['木曜'] == 0 ) & (df['金曜'] == 0 ) & (df['土曜'] == 0 ) & (df['日曜'] == 0 ))]
# lf = df.loc[~((df['ID'] == 18654) | (df['ID'] == 18656))]
df = df[['ID', '品目（量目は目安です）', '出荷元（生産者）', '生産地', 'ロット', '商品入数', '単位', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']]
df = df.rename(columns={'品目（量目は目安です）': '商品名', '出荷元（生産者）': '生産者名', '生産地': '産地（都道府県）', 'ロット': '入数'})
df.insert(7, 'JANコード', '')
df.insert(8, '掲載単価', '')
df[['ID', '入数', '商品入数']] = df[['ID', '入数', '商品入数']].astype('int')
df['商品入数'] = df['商品入数'].apply(str)

df['規格'] = df['商品入数'].str.cat(df['単位'])

df = df.drop(['商品入数', '単位'], axis=1)
df[''] = ''
col = df.pop('規格')
df.insert(loc=5, column='規格', value=col)
print('df: ', df.head(10))
df = df.rename(columns=days_list)
df.to_excel(f'./{export_file_name}.xlsx', index=False)

file = f'./{export_file_name}.xlsx'
wb = openpyxl.load_workbook(file)
ws = wb['Sheet1']
ws.insert_cols(9, 2)
ws.insert_rows(0, 3)
wb.save(file)
