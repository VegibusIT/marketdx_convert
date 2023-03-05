import pandas as pd
import copy
import datetime
import openpyxl

# time = datetime.datetime.now()
# time_now = time.strftime('%Y-%m-%d %H:%M:%S')
xlsx_file_name = 'original'
xlsx_file = f'{xlsx_file_name}.xlsx'
xlsx_file_start = f'{xlsx_file}_start.xlsx'
export_file_name = f'{xlsx_file}_final.xlsx'  # * 出力ファイル名

wb = openpyxl.load_workbook(xlsx_file, data_only=True)
wb.save(xlsx_file_start)

# extract target_connect_id and make it List
df = pd.read_excel(f'./{xlsx_file_start}', header=0)

extract_col = ['最終納品先店舗名', '販売日', '商品名', '販売単価', '出荷確定数', '生産者名', '産地市町村名', 'JANコード', 'ID', '商品入数', '単位', 'やさいバス数量']

df = df[extract_col]
df['商品入数'] = df['商品入数'].apply(str)
df['規格'] = df['商品入数'].str.cat(df['単位'])
df = df.sort_values(by=['最終納品先店舗名', '商品名'])
rename_col = {'販売単価': '掲載単価', '産地市町村名': '産地（都道府県）'}
df = df.rename(columns=rename_col)
extract_col = ['販売日', '商品名', '出荷確定数', 'ID', '生産者名', '産地（都道府県）', '規格', 'JANコード', '掲載単価', 'やさいバス数量', '商品入数']

# * df分割
list = {
    'minamisuna': '南砂店',
    'makuharishintoshin': '幕張新都心店',
    'asahichuo': '旭中央店',
    'shinonome': '東雲店',
    'kaihinmakuhari': '海浜幕張店',
    'kasai': '葛西店',
    'chousi': '銚子店',
    'kamatori': '鎌取店',
    'tateyama': '館山店',
    'kamogawa': '鴨川店',
}

for k, v in list.items():
    shop_list = df['最終納品先店舗名'].to_list()
    if v in shop_list:
        df_shop = df.groupby('最終納品先店舗名').get_group(v)
    else:
        continue
    df_shop = df_shop[extract_col]

    product_name = [k for k in dict.fromkeys(df_shop['商品名']).keys()]

    # todo 年を削除して、月日にする
    df_shop = df_shop.sort_values('販売日')
    # print('df_shop: ', df_shop)
    days = [k.strftime('%-m/%-d') for k in dict.fromkeys(df_shop['販売日']).keys()]  # todo For Windows %#m/%#d

    order_df = pd.DataFrame(columns=['ID', '生産者名', '産地（都道府県）', '入数', '規格', 'JANコード', '掲載単価', *days, ''], index=product_name)
    order_df = order_df.fillna(0)
    # * 最終列に空発列を追加
    order_df[''] = ''
    for i, v in df_shop.iterrows():
        order_df.at[v[1], v[0].strftime('%-m/%-d')] = v['出荷確定数']  # todo For Windows %#m/%#d
        order_df.at[v[1], 'ID'] = v['ID']
        order_df.at[v[1], '生産者名'] = v['生産者名']
        order_df.at[v[1], '産地（都道府県）'] = v['産地（都道府県）']
        order_df.at[v[1], '入数'] = v['やさいバス数量']
        order_df.at[v[1], '規格'] = v['規格']
        order_df.at[v[1], 'JANコード'] = v['JANコード']
        order_df.at[v[1], '掲載単価'] = v['掲載単価']

    order_df.reset_index(inplace=True)
    order_df = order_df.rename(columns={'index': '商品名'})
    col = order_df.pop('ID')
    order_df.insert(loc=0, column='ID', value=col)
    print('order_list: ', order_df.head(10))

    order_df.to_excel(f'./{k}.xlsx', index=False)  # * index_labelは、product_idにする

    wb = openpyxl.load_workbook(f'./{k}.xlsx')
    ws = wb.worksheets[0]
    ws.insert_cols(9, 2)
    ws.insert_rows(0, 3)
    wb.save(f'./{k}.xlsx')
