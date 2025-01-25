import pandas as pd
import numpy as np
from typing import Dict, List, Tuple
import matplotlib.pyplot as plt
import seaborn as sns
import os

class OlistAnalyzer:
    def __init__(self, products_df: pd.DataFrame, items_df: pd.DataFrame, output_dir: str = 'reports'):
        """
        Args:
            products_df: DataFrame с информацией о продуктах и категориях
            items_df: DataFrame с информацией о продажах
            output_dir: Директория для сохранения отчетов и визуализаций
        """
        self.products_df = products_df
        self.items_df = items_df
        self.output_dir = output_dir
        
        # Создаем директорию для отчетов, если её нет
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
    def calculate_category_metrics(self) -> pd.DataFrame:
        """
        Расчет расширенных метрик по категориям:
        - Количество продаж (order_item_id)
        - Общая выручка (price)
        - Средний чек (avg_ticket)
        - Количество уникальных продуктов (product_diversity)
        - Доля в общей выручке (revenue_share)
        """
        # Объединяем данные о продажах с данными о категориях
        sales_data = self.items_df.merge(
            self.products_df[['product_id', 'normalized_category']],
            on='product_id',
            how='left'
        )
        
        # Группируем по категориям и считаем базовые метрики
        metrics = sales_data.groupby('normalized_category').agg({
            'order_item_id': 'count',      # Количество продаж
            'price': ['sum', 'mean'],      # Общая выручка и средний чек
            'product_id': 'nunique'        # Разнообразие продуктов
        }).reset_index()
        
        # Переименовываем колонки для удобства
        metrics.columns = [
            'normalized_category',
            'total_sales',
            'total_revenue',
            'avg_ticket',
            'product_diversity'
        ]
        
        # Добавляем процент от общей выручки
        metrics['revenue_share'] = metrics['total_revenue'] / metrics['total_revenue'].sum() * 100
        
        # Добавляем ранг категории по выручке
        metrics['revenue_rank'] = metrics['total_revenue'].rank(ascending=False)
        
        return metrics.sort_values('total_revenue', ascending=False)
    
    def plot_category_distribution(self, metrics_df: pd.DataFrame, top_n: int = 10) -> None:
        """
        Создание и сохранение визуализаций по категориям
        
        Args:
            metrics_df: DataFrame с метриками категорий
            top_n: количество топовых категорий для отображения
        """
        # 1. Распределение выручки по категориям
        plt.figure(figsize=(12, 6))
        plot_data = metrics_df.head(top_n)
        
        sns.barplot(
            data=plot_data,
            x='normalized_category',
            y='revenue_share',
            color='skyblue'
        )
        
        plt.xticks(rotation=45, ha='right')
        plt.title('Топ категорий по выручке (%)')
        plt.tight_layout()
        
        plt.savefig(os.path.join(self.output_dir, 'revenue_distribution.png'))
        plt.close()
        
        # 2. Сравнение среднего чека по категориям
        plt.figure(figsize=(12, 6))
        sns.barplot(
            data=plot_data,
            x='normalized_category',
            y='avg_ticket',
            color='lightgreen'
        )
        
        plt.xticks(rotation=45, ha='right')
        plt.title('Средний чек по категориям')
        plt.tight_layout()
        
        plt.savefig(os.path.join(self.output_dir, 'avg_ticket_distribution.png'))
        plt.close()
        
    def save_metrics_report(self, metrics_df: pd.DataFrame) -> None:
        """Сохранение отчета с метриками в Excel"""
        report_path = os.path.join(self.output_dir, 'category_metrics_report.xlsx')
        metrics_df.to_excel(report_path, index=False)

    def generate_detailed_excel_report(self) -> None:
        """
        Создание подробного Excel-отчета с несколькими листами:
        - Общие метрики по категориям
        - Анализ продаж
        - Анализ среднего чека
        - Анализ популярности категорий
        - Анализ доставки
        """
        # Создаем writer для Excel
        report_path = os.path.join(self.output_dir, 'detailed_analysis_report.xlsx')
        with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
            # 1. Общие метрики по категориям
            category_metrics = self.calculate_category_metrics()
            category_metrics.to_excel(writer, sheet_name='Общие метрики', index=False)
            
            # 2. Анализ продаж по месяцам
            sales_by_month = self.analyze_sales_by_month()
            sales_by_month.to_excel(writer, sheet_name='Продажи по месяцам', index=True)
            
            # 3. Анализ среднего чека
            avg_ticket_analysis = self.analyze_average_ticket()
            avg_ticket_analysis.to_excel(writer, sheet_name='Средний чек', index=False)
            
            # 4. Популярность категорий
            category_popularity = self.analyze_category_popularity()
            category_popularity.to_excel(writer, sheet_name='Популярность категорий', index=False)
            
            # 5. Анализ доставки
            delivery_analysis = self.analyze_delivery_metrics()
            delivery_analysis.to_excel(writer, sheet_name='Метрики доставки', index=False)
            
            # Получаем workbook и добавляем форматирование
            workbook = writer.book
            
            # Форматы для чисел и процентов
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            percent_format = workbook.add_format({'num_format': '0.00%'})
            
            # Применяем форматирование к каждому листу
            for worksheet in writer.sheets.values():
                worksheet.set_column('A:Z', 15)  # Ширина колонок
                
    def analyze_sales_by_month(self) -> pd.DataFrame:
        """Анализ продаж по месяцам"""
        # Объединяем данные о заказах и товарах
        sales_data = self.items_df.merge(
            self.products_df[['product_id', 'normalized_category']],
            on='product_id'
        )
        
        # Добавляем информацию о дате заказа
        sales_data = sales_data.merge(
            self.datasets['orders'][['order_id', 'order_purchase_timestamp']],
            on='order_id'
        )
        
        # Конвертируем дату и создаем колонку месяц-год
        sales_data['order_date'] = pd.to_datetime(sales_data['order_purchase_timestamp'])
        sales_data['month_year'] = sales_data['order_date'].dt.to_period('M')
        
        # Группируем по месяцам и категориям
        monthly_sales = sales_data.groupby(['month_year', 'normalized_category']).agg({
            'order_id': 'count',
            'price': 'sum'
        }).reset_index()
        
        return monthly_sales
    
    def analyze_average_ticket(self) -> pd.DataFrame:
        """Анализ среднего чека по категориям"""
        avg_ticket = self.calculate_category_metrics()[
            ['normalized_category', 'avg_ticket', 'total_sales']
        ]
        
        # Добавляем дополнительные метрики
        avg_ticket['ticket_rank'] = avg_ticket['avg_ticket'].rank(ascending=False)
        avg_ticket['relative_to_mean'] = avg_ticket['avg_ticket'] / avg_ticket['avg_ticket'].mean()
        
        return avg_ticket.sort_values('avg_ticket', ascending=False)
    
    def analyze_category_popularity(self) -> pd.DataFrame:
        """Анализ популярности категорий"""
        popularity = self.calculate_category_metrics()[
            ['normalized_category', 'total_sales', 'revenue_share', 'product_diversity']
        ]
        
        # Добавляем метрики популярности
        total_sales = popularity['total_sales'].sum()
        popularity['sales_share'] = popularity['total_sales'] / total_sales
        popularity['popularity_score'] = (
            0.4 * popularity['sales_share'] + 
            0.4 * popularity['revenue_share'] / 100 + 
            0.2 * popularity['product_diversity'] / popularity['product_diversity'].max()
        )
        
        return popularity.sort_values('popularity_score', ascending=False)
    
    def analyze_delivery_metrics(self) -> pd.DataFrame:
        """Анализ метрик доставки"""
        # Объединяем данные о заказах и доставке
        delivery_data = self.datasets['orders'].copy()
        delivery_data['order_purchase_timestamp'] = pd.to_datetime(
            delivery_data['order_purchase_timestamp']
        )
        delivery_data['order_delivered_customer_date'] = pd.to_datetime(
            delivery_data['order_delivered_customer_date']
        )
        
        # Вычисляем время доставки
        delivery_data['delivery_time'] = (
            delivery_data['order_delivered_customer_date'] - 
            delivery_data['order_purchase_timestamp']
        ).dt.total_seconds() / 86400  # переводим в дни
        
        # Группируем по категориям
        delivery_metrics = delivery_data.merge(
            self.items_df.merge(
                self.products_df[['product_id', 'normalized_category']],
                on='product_id'
            ),
            on='order_id'
        ).groupby('normalized_category').agg({
            'delivery_time': ['mean', 'std', 'min', 'max'],
            'order_id': 'count'
        }).reset_index()
        
        return delivery_metrics 