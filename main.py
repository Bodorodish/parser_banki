import schedule
import time
from parser import parse_reviews
from statistic import ReviewAnalyzer

def job():
    print("Начинаем парсинг отзывов...")
    parse_reviews()

    print("Начинаем анализ отзывов...")
    analyzer = ReviewAnalyzer('Отзывы.xlsx')
    analyzer.process_reviews()

    print("Задача завершена!")

schedule.every(1).minute.do(job)

if __name__ == "__main__":
    print("Система готова к работе. Задачи будут выполняться через каждую минуту.")
    while True:
        schedule.run_pending()
        time.sleep(60)
