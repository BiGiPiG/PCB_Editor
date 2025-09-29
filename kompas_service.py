import os

import win32com.client
import pythoncom


class KompasService:
    def __init__(self):
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            self.kompas = None
            name = "Kompas.Application.7"

            try:
                self.kompas = win32com.client.Dispatch(name)
                self.kompas.Visible = True
            except Exception as e:
                print(f"Не удалось подключиться к {name}: {e}")

            if self.kompas is None:
                raise Exception("Не удалось подключиться к КОМПАС-3D")

        except Exception as e:
            print(f"Ошибка инициализации: {e}")
            self.cleanup()
            raise

    def create_fragment(self, path):
        """Метод для создания фрагмента"""
        try:
            new_doc = self.kompas.Documents.Add(2)  # 1 = фрагмент

            if not path.lower().endswith('.frw'):
                path += '.frw'

            new_doc.SaveAs(path)
            print(f"Создан файл: {path}")

        except Exception as e:
            print(f"Ошибка при создании фрагмента: {e}")

    def open_fragment(self, path):
        """Метод для открытия фрагмента"""
        try:
            if not os.path.exists(path):
                print(f"Файл не существует: {path}")
                return False

            if not path.lower().endswith('.frw'):
                path_with_ext = path + '.frw'
                if os.path.exists(path_with_ext):
                    path = path_with_ext
                else:
                    print(f"Файл с расширением .frw не найден: {path}")
                    return False

            print(f"Открываем файл: {path}")
            document = self.kompas.Documents.Open(path)
            if document:
                print("Файл успешно открыт")
                return True
            else:
                print("Не удалось открыть файл")
                return False

        except Exception as e:
            print(f"Ошибка при открытии фрагмента: {e}")
            return False

    def cleanup(self):
        """Очистка COM объектов"""
        try:
            if hasattr(self, 'kompas') and self.kompas:
                self.kompas.Quit()
                self.kompas = None
            pythoncom.CoUninitialize()
        except:
            pass

    def __del__(self):
        self.cleanup()
