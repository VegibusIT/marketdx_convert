import pandas as pd
import datetime
import copy

# time = datetime.datetime.now()
# time_now = time.strftime('%Y-%m-%d %H:%M:%S')

with open('./original.csv') as f:
    # extract target_connect_id and make it List
    df = pd.read_csv(f, header=0, sep=',')

    extract_col = ['最終納品先店舗名', '販売日', '商品名', '仕入単価（店着原価）', '仕入単価（生産者手取）', '販売単価', '出荷確定数', '生産者名', '産地市町村名', ]
    df = df[extract_col]
    df = df.sort_values(by=['最終納品先店舗名', '商品名'])
    extract_col = ['販売日', '商品名', '出荷確定数']

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
        df_shop = df.groupby('最終納品先店舗名').get_group(v)
        df_shop = df_shop[extract_col]
        product_name = [k for k in dict.fromkeys(df_shop['商品名']).keys()]

        # todo 年を削除して、月日にする
        days = [key for key in dict.fromkeys(df_shop['販売日']).keys()]
        order_list = pd.DataFrame(columns=days, index=product_name)
        order_list = order_list.fillna(0)
        for i, v in df_shop.iterrows():
            order_list.at[v[1], v[0]] = v[2]

        print('order_list: ', order_list)
        order_list.to_excel(f'./{k}.xlsx', index_label='商品名')  # * index_labelは、product_idにする
