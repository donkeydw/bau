# -*- coding: utf-8 -*-
import requests
import time
from sys import exit
import csv
import os
import pandas as pd

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36'}
today = ((time.strftime("%Y%m%d %H:%M:%S", time.localtime()))[:8])

start_date = input('Please enter the start date in YYYYMMDD format (eg: 20200601) to search official notice.\n'
             'Please directly press Enter key if you would like to take today as the start date: ')
end_date = input('Please enter the end date in in YYYYMMDD format (eg: 20200601) to search official notice.\n'
             'Please directly press Enter key if you would like to take today as the end date: ')
if len(start_date) == 0:
    start_date = today
if len(end_date) == 0:
    end_date = today
url = 'https://www.ser-ag.com/sheldon/official_notices/v2/find.json?firstDate=' + start_date + '&lastDate=' + end_date \
      + '&pageNumber=0&pageSize=2000&sortAttribute=dateTime&sortDirection=desc'
# url = 'https://www.ser-ag.com/sheldon/official_notices/v2/find.json?firstDate=20200619&lastDate=20200626&pageNumber=0&pageSize=2000&sortAttribute=dateTime&sortDirection=desc'
try:
    url_output = requests.get(url, headers=headers)
    dict_full = url_output.json()
except:
    print('No result. Please confirm if the date format is correct.')
    print('Program exit. Please retry it later.')
    exit_flag = input('Please enter E to exit: ')
    while exit_flag != 'E':
        exit_flag = input('Please enter E to exit: ')
    else:
        exit()
print('Getting the announcement list.....')
path_date = os.getcwd() + os.path.normpath("\SIX_%s" % today[4:])
if not os.path.exists(path_date):
    os.mkdir(path_date)
csv_path = os.path.normpath("%s\%s" % (path_date, 'SIX_announcement_list.csv'))
delisting_path = os.path.normpath("%s\%s" % (path_date, 'SIX_delisting.csv'))
newlisting_path = os.path.normpath("%s\%s" % (path_date, 'SIX_newlisting.csv'))
namechange_path = os.path.normpath("%s\%s" % (path_date, 'SIX_namechange.csv'))
csvfile = open(csv_path, 'w', encoding='utf-8', newline='')
writer = csv.writer(csvfile)
writer.writerow(['ID', 'Date', 'Notice Type', 'Issuer', 'Title', 'ISIN'])

item_list = dict_full.get('itemList')
if item_list:
    for item in item_list:
        item_id = str(item.get('noticeId'))
        item_date = str(item.get('date'))
        if item.get('noticeType') == 'A':
            item_type = 'Connexor event'
        elif item.get('noticeType') == 'FL':
            item_type = 'First Listing Notice'
        elif item.get('noticeType') == 'M':
            item_type = 'Further notice'
        elif item.get('noticeType') == 'EX':
            item_type = 'Ex-Dividend notice'
        elif item.get('noticeType') == 'PZ':
            item_type = 'Provisional Notice'
        elif item.get('noticeType') == 'DE':
            item_type = 'Delisting Notice'
        item_issuer = item.get('contact')
        item_title = item.get('title')
        if item.get('isin'):
            item_isin = item.get('isin')
        else:
            item_isin = 'N/A'
        item_data = [(item_id, item_date, item_type, item_issuer, item_title, item_isin)]
        writer.writerows(item_data)
else:
    print('No result is returned with requested date.')
csvfile.close()
# idv notice info
delisting, newlisting = [], []
df_announcement = pd.DataFrame(pd.read_csv(csv_path))
for announcement in df_announcement.index:
    i = (df_announcement.loc[announcement]).tolist()
    if i[2] == 'Delisting Notice':
        delisting.append(i[0])
    if i[2] == 'First Listing Notice':
        newlisting.append(i[0])
print('Processing.....')
df_delisting = pd.DataFrame(columns=('ISIN', 'Valor number', 'Company name', 'Security type', 'Valor_symbol', 'Last_listed_date'))
for item_id in delisting:
    url_idv = 'https://www.ser-ag.com/sheldon/official_notices/v2/details/' + str(item_id) + '.json'
    t = 0
    while t < 5:
        try:
            url_idv_output = requests.get(url_idv, headers=headers, timeout=5)
            dict_idv = url_idv_output.json()
        except:
            print('Connection error occurred. Retrying...')
            time.sleep(5)
            t += 1
        else:
            t = 5
    item_idv_text = dict_idv.get('itemList')[0].get('text')
    delisting_idv_list = item_idv_text.split('\n')
    for i in delisting_idv_list[1:]:
        if '|' in i:
            delisting_details = i.split('|')
            dict_delisting = {'ISIN': delisting_details[1].strip(), 'Valor number': delisting_details[0].strip(),
                              'Company name': delisting_details[2].strip(),
                              'Security type': delisting_details[3].strip(), 'Valor_symbol':
                                  delisting_details[4].strip(),
                              'Last_listed_date': delisting_details[5].strip().replace('.', '/')}
        else:
            delisting_details = i.split()
            company_name = ' '.join(delisting_details[2: -3])
            dict_delisting = {'ISIN': delisting_details[1].strip(), 'Valor number': delisting_details[0].strip(),
                              'Company name': company_name, 'Security type': delisting_details[-3].strip(),
                              'Valor_symbol': delisting_details[-2].strip(),
                              'Last_listed_date': delisting_details[-1].strip().replace('.', '/')}
        df_delisting = df_delisting.append(dict_delisting, ignore_index=True)
df_delisting.to_csv(delisting_path, encoding='utf-8', index=False)

df_newlisting = pd.DataFrame(columns=('ISIN', 'Valor number', 'Company name', 'Exchange Code', 'Valor_symbol',
                                      'Currency', 'First_listed_date', 'Last_listed_date'))
for item_id in newlisting:
    url_idv = 'https://www.ser-ag.com/sheldon/official_notices/v2/details/' + str(item_id) + '.json'
    t = 0
    while t < 5:
        try:
            url_idv_output = requests.get(url_idv, headers=headers, timeout=5)
            dict_idv = url_idv_output.json()
        except:
            print('Connection error occurred. Retrying...')
            time.sleep(5)
            t += 1
        else:
            t = 5
    item_idv_text = dict_idv.get('itemList')[0].get('text')
    if 'WTC_SUBSCR' in item_idv_text:
        newlisting_idv_list = item_idv_text.split('\n')
        for i in newlisting_idv_list[1:]:
            newlisting_details = i.split('|')
            if 'WTC_SUBSCR' in newlisting_details:
                dict_newlisting = {'ISIN': newlisting_details[1].strip(), 'Valor number': newlisting_details[0].strip(),
                    'Company name': newlisting_details[3].strip(), 'Exchange Code': '',
                    'Valor_symbol': newlisting_details[2].strip(), 'Currency': newlisting_details[4].strip(),
                    'First_listed_date': newlisting_details[5].strip().replace('.', '/'),
                    'Last_listed_date': newlisting_details[6].strip().replace('.', '/')}
                df_newlisting = df_newlisting.append(dict_newlisting, ignore_index=True)

    if 'There are no Shares today' in item_idv_text:
        continue
    else:
        newlisting_idv_list = item_idv_text.split('\n')[item_idv_text.split('\n').index('Shares'):]
        for i in newlisting_idv_list[2:]:
            newlisting_details = i.split('|')
            dict_newlisting = {'ISIN': newlisting_details[1].strip(), 'Valor number': newlisting_details[0].strip(),
                'Company name': newlisting_details[3].strip(), 'Exchange Code': newlisting_details[-1].strip(),
                'Valor_symbol': newlisting_details[2].strip(), 'Currency': newlisting_details[4].strip(),
                'First_listed_date': newlisting_details[5].strip().replace('.', '/'),
                'Last_listed_date': newlisting_details[6].strip().replace('.', '/')}
            df_newlisting = df_newlisting.append(dict_newlisting, ignore_index=True)
df_newlisting.to_csv(newlisting_path, encoding='utf-8', index=False)

print('Data grab has been completed. Please verify the output.')
print('Program will exit.')
exit_flag = input('Please enter E to exit: ')
while exit_flag != 'E':
    exit_flag = input('Please enter E to exit: ')
else:
    exit()