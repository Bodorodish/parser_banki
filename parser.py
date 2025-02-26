import requests
from bs4 import BeautifulSoup
import pandas as pd
import ssl
import time as t
from urllib.parse import urljoin
import os

ssl._create_default_https_context = ssl._create_unverified_context

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'
}

base_url = 'https://banki.ru/'

companies = {
    'СБСЖ': {
        'url': 'insurance/responses/company/sberbankstrahovaniezhizni/',
    },
    'СБС': {
        'url': 'insurance/responses/company/sberbankstrahovanie/',
    },
}

def parse_reviews():
    max_pages = 5

    if os.path.exists('Отзывы.xlsx'):
        existing_df = pd.read_excel('Отзывы.xlsx', parse_dates=['Время'], engine='openpyxl')
        existing_urls = set(existing_df['url жалобы'].dropna())
    else:
        existing_df = pd.DataFrame(columns=[
            'Компания', 'url жалобы', 'Заголовок', 'Статус', 'Текст', 'Время', 'Оценка', 'Оценка выплат',
        ])
        existing_urls = set()
    
    new_df = pd.DataFrame(columns=existing_df.columns)
    
    for company in companies:
        company_config = companies[company]
        page = 1
        stop_parsing = False
        
        while not stop_parsing and page <= max_pages:
            print(page)
            url = urljoin(base_url, company_config['url'])
            if page > 1:
                url = urljoin(url, f'?page={page}')
            
            try:
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                soup = BeautifulSoup(response.text, 'html.parser')
                
                articles = soup.find_all('article')
                if not articles:
                    break  # нет статей -> выходим
                
                current_page_reviews = []
                for article in articles:
                    try:
                        link_elem = article.find('a', {'data-test': 'responses-header'})
                        href = link_elem.get('href') if link_elem else None
                        header = link_elem.text.strip() if link_elem else None
                        
                        text_elem = article.find('div', {'data-test': 'responses-message'})
                        text = text_elem.text.strip() if text_elem else None
                        
                        time_elem = article.find('time', {'data-test': 'responses-datetime'})
                        time_str = time_elem.get('datetime') if time_elem else None
                        time = pd.to_datetime(time_str) if time_str else pd.NaT
                        
                        rating_elem = article.find('span', {'data-test': 'responses-rating-grade'})
                        rating = rating_elem.text.strip() if rating_elem else None
                        
                        rating_payouts_elem = article.find('strong', {'class': 'font-size-medium'})
                        rating_payouts = rating_payouts_elem.text.strip() if rating_payouts_elem else None
                        
                        rating_status_elem = article.find('span', {'data-test': 'responses-status'})
                        rating_status = rating_status_elem.text.strip() if rating_status_elem else None
                        
                        full_url = urljoin(base_url, href) if href else None
                        
                        if not full_url:
                            continue
                        
                        if full_url in existing_urls:
                            stop_parsing = True
                            break
                        
                        row = {
                            'Компания': company,
                            'url жалобы': full_url,
                            'Заголовок': header,
                            'Статус': rating_status,
                            'Текст': text,
                            'Время': time,
                            'Оценка': rating,
                            'Оценка выплат': rating_payouts,
                        }
                        current_page_reviews.append(row)
                    
                    except Exception as e:
                        print(f'Ошибка при парсинге статьи: {e}')
                        continue
                
                if current_page_reviews:
                    temp_df = pd.DataFrame(current_page_reviews)
                    new_df = pd.concat([new_df, temp_df], ignore_index=True)
                
                if stop_parsing:
                    break
                
                page += 1
                t.sleep(0.1)
            
            except requests.exceptions.RequestException as e:
                print(f'Ошибка при загрузке страницы {url}: {e}')
                break
    
    if not new_df.empty:
        final_df = pd.concat([new_df, existing_df], ignore_index=True)
        final_df.drop_duplicates(subset=['url жалобы'], keep='first', inplace=True)
        final_df.sort_values(by='Время', ascending=False, inplace=True)
    else:
        final_df = existing_df
    
    final_df.to_excel('Отзывы.xlsx', index=False)
    print(f'Всего отзывов после обновления: {len(final_df)}')


if __name__ == "__main__":
    print("Начало парсинга...")
    parse_reviews()
    print("Готово")