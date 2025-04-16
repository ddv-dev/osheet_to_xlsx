import xlsxwriter
import json


class osheet_to_exlsx:
    __filepath = ""
    def __init__(self, filepath):
        self.__filepath = filepath
        
    def __load_file(self, file_path=__filepath):
        with open(file_path, 'rb') as f:
            content = b''.join(f.readlines())
            print(content)
        return content

    def __split_str(self, bytes):
         # Инициализация счетчика для отслеживания уровня вложенности фигурных скобок
        left_count = 0

        # Список для хранения фрагментов строки, разделенных по балансу скобок
        content_list = []

        # Проход по каждому байту в переданной строке `content`
        for i in range(0, len(bytes)):
        # Если скобки сбалансированы (left_count == 0), начинаем новый фрагмент
            if left_count == 0:
                content_list.append(b'')
        # Если текущий символ — '{' (код 123 в ASCII), увеличиваем счетчик
            if bytes[i] == 123:
                left_count += 1
        # Если текущий символ — '}' (код 125 в ASCII), уменьшаем счетчик
            elif bytes[i] == 125:
                left_count -= 1
        # Добавляем текущий байт к последнему фрагменту в content_list
            content_list[-1] += bytes[i].to_bytes(1, byteorder='little', signed=False)

        # Список для хранения только тех фрагментов, которые содержат '{'    
        return_list = []

        # Проверяем каждый фрагмент на наличие '{'
        for i in range(0, len(content_list)):    
            for j in range(0, len(content_list[i])):

                # Если найден символ '{', добавляем фрагмент в return_list (в виде строки)
                if content_list[i][j] == 123:
                
                    # Декодируем байтовую строку в UTF-8 (с заменой некорректных символов)
                    return_list.append(content_list[i].decode('utf-8', 'replace'))
                    break
        # Прерываем внутренний цикл после первой найденной '{'
        return return_list


    def convert(self):
        wb = xlsxwriter.Workbook('convert_file.xlsx')
        sheets = []
        content = self.__split_str(self.__load_file())
    
        
        for _content in content:
            _content = json.loads(_content)
            if 'sheets' in _content:
                for key, value in _content['sheets'].items():
                    sheets.append(value['title'])
        sheet_num = 0
        for _content in content:
            _content = json.loads(_content)
            if 'cells' in _content:
                sh = wb.add_worksheet(sheets[sheet_num])
                for row, row_value in _content['cells'].items():
                    for col, value in row_value.items():
                        if 'v' in value:                
                            sh.write(int(row), int(col), value['v'])
                sheet_num += 1
        wb.close()
    
    
