import os
import json
import appdirs


class PathManager:
    """Класс для управления сохранением/загрузкой путей между сессиями"""
    
    def __init__(self):
        # Определяем путь к конфигурационному файлу в пользовательской директории
        self.config_dir = appdirs.user_config_dir("ID+TG", "DocumentProcessor")
        self.config_file = os.path.join(self.config_dir, "paths_config.json")
        
        # Создаем директорию конфигурации, если она не существует
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
    
    def save_paths(self, paths_dict):
        """
        Сохраняет пути и настройки в конфигурационный файл
        
        Args:
            paths_dict (dict): Словарь с путями и настройками
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(paths_dict, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения конфигурации: {e}")
    
    def load_paths(self):
        """
        Загружает сохраненные пути и настройки из конфигурационного файла
        
        Returns:
            dict: Словарь с путями и настройками или пустой словарь, если файл не найден
        """
        if not os.path.exists(self.config_file):
            return {}
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Ошибка загрузки конфигурации: {e}")
            return {}
    
    def clear_paths(self):
        """Удаляет конфигурационный файл"""
        if os.path.exists(self.config_file):
            try:
                os.remove(self.config_file)
            except Exception as e:
                print(f"Ошибка удаления конфигурации: {e}")