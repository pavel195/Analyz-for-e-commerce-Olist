from data_loader import OlistDataLoader
from analyzer import OlistAnalyzer
import pandas as pd
import os

def main():
    # Получаем абсолютный путь к директории с данными
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(os.path.dirname(current_dir), 'Olist')
    reports_path = os.path.join(os.path.dirname(current_dir), 'reports')
    
    # Инициализация загрузчика данных с указанием файла справочника
    data_loader = OlistDataLoader(
        data_path,
        mapping_file=os.path.join(data_path, 'category_mapping.xlsx')
    )
    data_loader.load_data()
    
    # Подготовка данных о категориях
    products_with_categories = data_loader.prepare_product_categories()
    
    # Выводим маппинг категорий
    print("\nCategory Mapping Example:")
    mapping_sample = products_with_categories[
        ['product_category_name', 'product_category_name_english', 'normalized_category']
    ].drop_duplicates().head()
    print(mapping_sample)
    
    # Создание анализатора с указанием директории для отчетов
    analyzer = OlistAnalyzer(
        products_df=products_with_categories,
        items_df=data_loader.datasets['items'],
        output_dir=reports_path
    )
    
    # Расчет метрик
    category_metrics = analyzer.calculate_category_metrics()
    
    # Вывод результатов
    print("\nTop Categories by Revenue:")
    print(category_metrics[
        ['normalized_category', 'total_sales', 'total_revenue', 
         'avg_ticket', 'revenue_share']
    ].head())
    
    # Построение визуализаций
    analyzer.plot_category_distribution(category_metrics)
    
    # Сохранение отчета
    analyzer.save_metrics_report(category_metrics)
    
    # Генерация подробного Excel-отчета
    print("\nГенерация подробного отчета...")
    analyzer.generate_detailed_excel_report()
    
    print(f"\nОтчеты и визуализации сохранены в директории: {reports_path}")

if __name__ == "__main__":
    main() 