import requests
from bs4 import BeautifulSoup
import pandas as pd
import ssl
import time as t
from urllib.parse import urljoin

ssl._create_default_https_context = ssl._create_unverified_context

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'
}

base_url = 'https://banki.ru/'

companies = {
    'СБСЖ': {
        'pages': 50,  
        'url': 'insurance/responses/company/sberbankstrahovaniezhizni/',
    },
    'СБС': {
        'pages': 16,
        'url': 'insurance/responses/company/sberbankstrahovanie/',
    },
}

def parse_reviews():
    df = pd.DataFrame(columns=[
        'Компания', 'url жалобы', 'Заголовок', 'Статус', 'Текст', 'Время', 'Оценка', 'Оценка выплат',
    ])
    
    for company in companies:
        #print(f'Обрабатываем компанию: {company}')
        company_config = companies[company]
        
        for page in range(1, company_config['pages'] + 1):
            url = urljoin(base_url, company_config['url'])
            if page > 1:
                url = urljoin(url, f'?page={page}')
            
            try:
                #print(f'Загрузка страницы: {url}')
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                soup = BeautifulSoup(response.text, 'html.parser')
                
                articles = soup.find_all('article')
                if not articles:
                    print(f'Нет статей на странице {page}')
                    break
                
                for article in articles:
                    try:
                        link_elem = article.find('a', {'data-test': 'responses-header'})
                        href = link_elem.get('href') if link_elem else None
                        header = link_elem.text.strip() if link_elem else None
                        
                        text_elem = article.find('div', {'data-test': 'responses-message'})
                        text = text_elem.text.strip() if text_elem else None
                        
                        time_elem = article.find('time', {'data-test': 'responses-datetime'})
                        time = time_elem.get('datetime') if time_elem else None
                        
                        rating_elem = article.find('span', {'data-test': 'responses-rating-grade'})
                        rating = rating_elem.text.strip() if rating_elem else None
                        
                        rating_payouts_elem = article.find('strong', {'class': 'font-size-medium'})
                        rating_payouts = rating_payouts_elem.text.strip() if rating_payouts_elem else None
                        
                        rating_status_elem = article.find('span', {'data-test': 'responses-status'})
                        rating_status = rating_status_elem.text.strip() if rating_status_elem else None
                        
                        full_url = urljoin(base_url, href) if href else None
                        
                        row = [company, full_url, header, rating_status, text, time, rating, rating_payouts]
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(f'Ошибка при парсинге статьи: {e}')
                        continue
                
                t.sleep(0.1)
                
            except requests.exceptions.RequestException as e:
                print(f'Ошибка при загрузке страницы: {e}')
                continue

    df.to_excel('Отзывы.xlsx', index=False)
    #print(f'Всего собрано отзывов: {len(df)}')
