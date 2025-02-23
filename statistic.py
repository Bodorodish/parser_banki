import os
import pandas as pd
import re
from collections import defaultdict
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image

class ReviewAnalyzer:
    def __init__(self, file_path):
        self.df = pd.read_excel(file_path)
        self.keyword_categories = {
            'выплаты': ['выплат', 'деньги', 'компенсац', 'возмещен'],
            'сроки': ['срок', 'долго', 'ждал', 'задержк'],
            'обслуживание': ['сотрудник', 'менеджер', 'вежлив', 'груб', 'обслуж'],
            'документы': ['документ', 'справк', 'бумаг', 'оформлен'],
            'страховой_случай': ['отказ', 'непризнан', 'страховой случай']
        }
        
    def _classify_rating(self, rating):
        try:
            if isinstance(rating, str):
                rating = rating.replace(',', '.')
            numeric_rating = float(rating)
            
            if numeric_rating >= 4:
                return 'Отлично'
            elif 3 <= numeric_rating < 4:
                return 'Среднее'
            else:
                return 'Плохо'
        except (ValueError, TypeError) as e:
            print(f"Ошибка при преобразовании оценки: {rating}. Ошибка: {e}")
            return 'Не определено'

    def _analyze_text(self, text):
        if pd.isnull(text):
            return []
        text = text.lower()
        found_categories = set()
        
        for category, keywords in self.keyword_categories.items():
            for keyword in keywords:
                if re.search(rf'\b{keyword}', text):
                    found_categories.add(category)
                    break
        return list(found_categories)

    def process_reviews(self):
        self.df['Категория'] = self.df['Оценка'].apply(self._classify_rating)
        self.df['Проблемы'] = self.df['Текст'].apply(self._analyze_text)
        
        bad_reviews = self.df[self.df['Категория'] == 'Плохо']
        print(f"\nНайдено плохих отзывов: {len(bad_reviews)}")
        
        problem_stats = defaultdict(int)
        for problems in bad_reviews['Проблемы']:
            for problem in problems:
                problem_stats[problem] += 1
        
        if problem_stats:
            print("\nРаспределение проблем:")
            for problem, count in sorted(problem_stats.items(), key=lambda x: -x[1]):
                print(f"- {problem}: {count} упоминаний")
        else:
            print("\nНет проблем в плохих отзывах.")
        
        output_file = 'Отзывы.xlsx'
        self.df.to_excel(output_file, index=False)

        self._add_visualizations(output_file, problem_stats)
        
        print(f"\nФайл сохранен: {output_file}")
        
        return self.df
    
    def _add_visualizations(self, output_file, problem_stats):
        plt.figure(figsize=(10, 6))
        rating_counts = self.df['Категория'].value_counts()
        sns.barplot(x=rating_counts.index, y=rating_counts.values, palette="viridis")
        plt.title('Распределение оценок')
        plt.xlabel('Категория')
        plt.ylabel('Количество')
        plt.tight_layout()
        
        chart_path = 'rating_distribution.png'
        plt.savefig(chart_path)
        plt.close()

        plt.figure(figsize=(10, 6))
        problems, counts = zip(*sorted(problem_stats.items(), key=lambda x: -x[1]))
        sns.barplot(x=list(counts), y=list(problems), palette="magma")
        plt.title('Распределение проблем в плохих отзывах')
        plt.xlabel('Количество упоминаний')
        plt.ylabel('Проблемы')
        plt.tight_layout()
        
        problem_chart_path = 'problem_distribution.png'
        plt.savefig(problem_chart_path)
        plt.close()

        wb = load_workbook(output_file)
        sheet = wb.active

        img1 = Image(chart_path)
        img2 = Image(problem_chart_path)
        sheet.add_image(img1, 'L1')
        sheet.add_image(img2, 'L32')
        wb.save(output_file)

        #os.remove(chart_path)
        #os.remove(problem_chart_path)


if __name__ == "__main__":
    analyzer1 = ReviewAnalyzer('Отзывы.xlsx')
    analyzed_df = analyzer1.process_reviews()
