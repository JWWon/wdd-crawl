import os
import sys
import requests
import re
import pandas as pd

from save_to_xlsx import append_df_to_excel

# APPLICATION INFO
client_id = "ppYCMnYAz3em2lZ4Oisn"
client_secret = "bUstOMZXpg"

# CONSTS
baseURL = "https://openapi.naver.com/v1/search/local.json"
headers = {"X-Naver-Client-Id": client_id,
           "X-Naver-Client-Secret": client_secret}
filename = 'naver_data.xlsx'
sheet_name = 'Data'

keywords = pd.read_csv('keyword.csv')
writer = pd.ExcelWriter(filename, engine='openpyxl')

# DELETE FILE
try:
    os.remove(filename)
except OSError:
    pass

# CREATE REQUEST
for _, row in keywords.iterrows():
    keyword = row['keyword']
    print(f'## KEYWORD: {keyword}')

    for i in range(34):
        params = {"query": keyword, "display": 30, "start": i * 30 + 1}
        res = requests.get(baseURL, params=params, headers=headers)

        if res.status_code == 200:
            data = res.json()
            items = data['items']
            items_len = len(items)
            print(f'  count: {i}, results: {items_len}')

            for item in items:
                titles = re.sub('(<b>|</b>)', ' ', item['title'])
                titles.strip()
                del item['title']
                item['title'] = titles

            # convert items to dataframe
            df = pd.DataFrame(items)
            append_df_to_excel(
                filename, df, sheet_name=sheet_name, index=False)

            if items_len < 30:
                break
