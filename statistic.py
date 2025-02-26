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
        self.problem_stats = defaultdict(int)
        self.rating_counts = None

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
        
        for problems in bad_reviews['Проблемы']:
            for problem in problems:
                self.problem_stats[problem] += 1
        
        if self.problem_stats:
            print("\nРаспределение проблем:")
            for problem, count in sorted(self.problem_stats.items(), key=lambda x: -x[1]):
                print(f"- {problem}: {count} упоминаний")
        else:
            print("\nНет проблем в плохих отзывах.")
        
        output_file = 'Отзывы.xlsx'
        self.df.to_excel(output_file, index=False)

        self._add_visualizations(output_file)
        self._generate_recommendations(output_file)
        
        print(f"\nФайл сохранен: {output_file}")
        return self.df

    def _add_visualizations(self, output_file):
        plt.figure(figsize=(10, 6))
        self.rating_counts = self.df['Категория'].value_counts()
        sns.barplot(x=self.rating_counts.index, y=self.rating_counts.values, palette="viridis")
        plt.title('Распределение оценок')
        plt.xlabel('Категория')
        plt.ylabel('Количество')
        plt.tight_layout()
        
        chart_path = 'rating_distribution.png'
        plt.savefig(chart_path)
        plt.close()

        plt.figure(figsize=(10, 6))
        if self.problem_stats:
            problems, counts = zip(*sorted(self.problem_stats.items(), key=lambda x: -x[1]))
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
        img2 = Image(problem_chart_path) if self.problem_stats else None
        sheet.add_image(img1, 'L1')
        if img2:
            sheet.add_image(img2, 'L32')
        wb.save(output_file)

        os.remove(chart_path)
        if img2:
            os.remove(problem_chart_path)

    def _generate_recommendations(self, output_file):
        generator = RecommendationGenerator(
            problem_stats=self.problem_stats,
            rating_counts=self.rating_counts.to_dict(),
            df=self.df
        )
        generator.add_to_excel(output_file)


class RecommendationGenerator:
    def __init__(self, problem_stats, rating_counts, df):
        self.problem_stats = problem_stats
        self.rating_counts = rating_counts
        self.df = df
        self.recommendation_templates = {
            'выплаты': [
                "Автоматизировать процесс проверки выплат",
                "Внедрить систему уведомлений о статусе выплат"
            ],
            'сроки': [
                "Оптимизировать внутренние процессы для сокращения сроков",
                "Ввести KPI для контроля времени обработки запросов"
            ],
            'обслуживание': [
                "Провести тренинг для сотрудников по работе с клиентами",
                "Внедрить систему оценки качества обслуживания"
            ],
            'документы': [
                "Разработать шаблоны документов для клиентов",
                "Создать интерактивный гид по оформлению документов"
            ],
            'страховой_случай': [
                "Пересмотреть критерии признания страховых случаев",
                "Внедрить многоэтапную проверку спорных случаев"
            ],
            'general': [
                "Провести глубинные интервью с клиентами",
                "Реализовать систему NPS-оценки"
            ]
        }

    def generate(self):
        recommendations = []
        
        for problem, _ in sorted(self.problem_stats.items(), 
                               key=lambda x: x[1], reverse=True)[:3]:
            if problem in self.recommendation_templates:
                recs = self.recommendation_templates[problem]
                recommendations.append({
                    'Категория': problem,
                    'Рекомендация': recs[0],
                    'Дополнительные меры': "; ".join(recs[1:])
                })
        
        if self.rating_counts.get('Плохо', 0) > 10:
            recs = self.recommendation_templates['general']
            recommendations.append({
                'Категория': 'Общее улучшение',
                'Рекомендация': recs[0],
                'Дополнительные меры': "; ".join(recs[1:])
            })
            
        return pd.DataFrame(recommendations)

    def add_to_excel(self, output_file):
        df_rec = self.generate()
        if not df_rec.empty:
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
                df_rec.to_excel(writer, sheet_name='Рекомендации', index=False)


if __name__ == "__main__":
    analyzer1 = ReviewAnalyzer('Отзывы.xlsx')
    analyzed_df = analyzer1.process_reviews()
