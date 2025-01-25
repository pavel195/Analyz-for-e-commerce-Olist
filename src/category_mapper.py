import pandas as pd
from typing import Optional
import os

class CategoryMapper:
    def __init__(self, mapping_file: str = 'category_mapping.xlsx'):
        self.mapping_file = mapping_file
        self.mapping_df: Optional[pd.DataFrame] = None
        
    def create_initial_mapping(self, categories_df: pd.DataFrame) -> None:
        """Создание начального справочника категорий"""
        mapping_data = {
            'original_portuguese_category': categories_df['product_category_name'],
            'english_translation': categories_df['product_category_name_english'],
            'normalized_category': categories_df['product_category_name_english'].apply(
                lambda x: str(x).title().replace(' ', '_') if pd.notna(x) else 'Unknown'
            )
        }
        
        self.mapping_df = pd.DataFrame(mapping_data).drop_duplicates()
        
        # Сохраняем справочник в Excel
        self.save_mapping()
        
    def load_mapping(self) -> None:
        """Загрузка существующего справочника"""
        if os.path.exists(self.mapping_file):
            self.mapping_df = pd.read_excel(self.mapping_file)
        else:
            raise FileNotFoundError(f"Mapping file {self.mapping_file} not found")
            
    def save_mapping(self) -> None:
        """Сохранение справочника в Excel"""
        if self.mapping_df is not None:
            self.mapping_df.to_excel(self.mapping_file, index=False)
            
    def get_normalized_category(self, portuguese_category: str) -> str:
        """Получение нормализованной категории по португальскому названию"""
        if self.mapping_df is None:
            self.load_mapping()
            
        match = self.mapping_df[
            self.mapping_df['original_portuguese_category'] == portuguese_category
        ]
        
        if len(match) > 0:
            return match.iloc[0]['normalized_category']
        return 'Unknown' 