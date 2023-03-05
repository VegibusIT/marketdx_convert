import pandas as pd
import datetime
import copy

# time = datetime.datetime.now()
# time_now = time.strftime('%Y-%m-%d %H:%M:%S')
xlsx_file_name = 'original'
xlsx_file = f'{xlsx_file_name}.xlsx'
xlsx_file_start = f'{xlsx_file}_start.xlsx'
export_file_name = f'{xlsx_file}_final.xlsx'  # * 出力ファイル名

# extract target_connect_id and make it List
df = pd.read_excel(f'./{xlsx_file}', header=0)

extract_col = ['最終納品先店舗名', '販売日', '商品名', '販売単価', '出荷確定数', '生産者名', '産地市町村名', 'JANコード', 'ID', '商品入数', '単位', '数量']

df = df[extract_col]
df['商品入数'] = df['商品入数'].apply(str)
df['規格'] = df['商品入数'].str.cat(df['単位'])
df = df.sort_values(by=['最終納品先店舗名', '商品名'])
rename_col = {'販売単価': '掲載単価', '産地市町村名': '産地（都道府県）', '数量': '入数'}
df = df.rename(columns=rename_col)
extract_col = ['販売日', '商品名', '出荷確定数', 'ID', '生産者名', '産地（都道府県）', '規格', 'JANコード', '掲載単価', '入数']

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

    order_df = pd.DataFrame(columns=['ID', '生産者名', '産地（都道府県）', '入数', '規格', 'JANコード', '掲載単価', *days], index=product_name)
    order_df = order_df.fillna(0)
    for i, v in df_shop.iterrows():
        # print('v: ', v)
        order_df.at[v[1], v[0].strftime('%-m/%-d')] = v[2]  # todo For Windows %#m/%#d
        order_df.at[v[1], 'ID'] = v[3]
        order_df.at[v[1], '生産者名'] = v[4]
        order_df.at[v[1], '産地（都道府県）'] = v[5]
        order_df.at[v[1], '入数'] = v[9]
        order_df.at[v[1], '規格'] = v[6]
        order_df.at[v[1], 'JANコード'] = v[7]
        order_df.at[v[1], '掲載単価'] = v[8]
    print('order_list: ', order_df)

    # print('order_list: ', order_list)
    order_df.to_excel(f'./{k}.xlsx', index_label='商品名')  # * index_labelは、product_idにする
