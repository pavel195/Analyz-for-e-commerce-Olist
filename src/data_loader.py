import pandas as pd
from typing import Dict, Optional
import os
from category_mapper import CategoryMapper

class OlistDataLoader:
    def __init__(self, data_path: str, mapping_file: str = 'category_mapping.xlsx'):
        self.data_path = data_path
        self.datasets: Dict[str, pd.DataFrame] = {}
        self.category_mapper = CategoryMapper(mapping_file)
        
    def load_data(self) -> None:
        """Загрузка всех необходимых датасетов"""
        files = {
            'orders': 'olist_orders_dataset.csv',
            'items': 'olist_order_items_dataset.csv',
            'products': 'olist_products_dataset.csv',
            'sellers': 'olist_sellers_dataset.csv',
            'categories': 'product_category_name_translation.csv'
        }
        
        for key, filename in files.items():
            self.datasets[key] = pd.read_csv(f"{self.data_path}/{filename}")
            
    def prepare_product_categories(self) -> pd.DataFrame:
        """Подготовка данных о категориях продуктов с использованием справочника"""
        products_df = self.datasets['products']
        categories_df = self.datasets['categories']
        
        # Создаем справочник категорий, если его еще нет
        if not os.path.exists(self.category_mapper.mapping_file):
            self.category_mapper.create_initial_mapping(categories_df)
        else:
            self.category_mapper.load_mapping()
        
        # Объединяем таблицы
        merged_df = products_df.merge(
            categories_df,
            on='product_category_name',
            how='left'
        )
        
        # Применяем маппинг категорий
        merged_df['normalized_category'] = merged_df['product_category_name'].apply(
            self.category_mapper.get_normalized_category
        )
        
        return merged_df 