from collections import defaultdict
import re

class drillingFileReader:

    def readFile(path):
        
        lines = ""
        
        with open(path, "r") as file:
            
            for line in file:
                
                lines += line;
        
        return drillingFileReader.parse_drill_file(lines)
        
    def parse_drill_file(file_content):
        # Словарь для хранения диаметров инструментов
        tool_diameters = {}
        # Итоговый словарь с координатами
        result = defaultdict(list)
        # Флаг начала координат
        coordinates_started = False
        
        # Регулярные выражения
        tool_pattern = re.compile(r'T(\d+)C(\d+\.\d+)')
        coordinate_pattern = re.compile(r'X(-?\d+\.\d+)Y(-?\d+\.\d+)')
        
        for line in file_content.strip().split('\n'):
            line = line.strip()
            
            # Пропускаем пустые строки и комментарии
            if not line or line.startswith(';'):
                continue
              
            # Если встретили символ %, переходим к обработке координат
            if line.startswith('%'):
                coordinates_started = True
                continue
            
            # Если еще не начали обработку координат, собираем информацию о диаметрах
            elif not coordinates_started:
                match = tool_pattern.match(line)
                if match:
                    tool_id = match.group(1)
                    diameter = float(match.group(2))
                    tool_diameters["T" + tool_id] = diameter
                continue
                
            # Обработка координат
            if coordinates_started:
                # Проверяем, указан ли новый инструмент
                if line.startswith('T'):
                    current_tool = line.strip()
                    continue
                    
                # Ищем координаты
                match = coordinate_pattern.match(line)
                if match:
                    x = float(match.group(1))
                    y = float(match.group(2))
                    # Получаем диаметр из словаря по текущему инструменту
                    diameter = tool_diameters[current_tool]
                    result[diameter].append((x, y))
                    
        return result