from bs4 import BeautifulSoup
import requests
import pandas as pd
import time
import random


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36'}


def get_operatating_cash_flow(ticker_code: str):
    url = f'https://finance.yahoo.com/quote/{ticker_code}/cash-flow?p={ticker_code}'
    

    r = requests.get(url, headers=headers)
    html = r.content
    soup = BeautifulSoup(html, 'html.parser')
    try:
        c = soup.find('div', {'title': 'Operating Cash Flow'})
        q = c.find_parent('div', {'class': 'D(tbr) fi-row Bgc($hoverBgColor):h'})
        r = q.find_all('div')[3]
        return r.get_text()
    except AttributeError:
        c = soup.find('div', {'title': 'Cash Flows from Used in Operating Activities Direct'})
        q = c.find_parent('div', {'class': 'D(tbr) fi-row Bgc($hoverBgColor):h'})
        r = q.find_all('div')[3]
        return r.get_text()

with pd.ExcelWriter('clean_tickers3.xlsx', engine='openpyxl') as writer:
    df = pd.read_excel('Ticker-Symbols.xlsx',sheet_name='Sheet1')
    rem = len(df)
    for k, tick in df.iterrows():
        time.sleep(random.randint(4,7))
        try:
            df.loc[k,'operating_cash_flow'] = str(get_operatating_cash_flow(tick['Ticker'].replace(':', '.')))
        except AttributeError:
            df.loc[k,'operating_cash_flow'] = 'N/A'
        print(rem)
        rem -= 1
    print('Completed!')
    
    df.to_excel(writer, sheet_name='tic', index=False)

