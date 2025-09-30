import os

import win32com.client
import pythoncom
from win32com.client import gencache


class KompasService:
    def __init__(self):
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            self.kompas = None
            self.kompas_api7_module = None
            name = "Kompas.Application.7"

            try:
                self.kompas = win32com.client.Dispatch(name)
                self.kompas.Visible = True
                self.kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
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

            doc = self.kompas.Documents.Add(2)

            doc2d = self.kompas_api7_module.IKompasDocument2D(doc)

            views = doc2d.ViewsAndLayersManager.Views[0]
            m1 = views.MacroObjects.Add()

            l1 = m1.LineSegments.Add()

            l1.Style = 3
            l1.X1 = -5
            l1.Y1 = 0
            l1.X2 = 5
            l1.Y2 = 0

            l1.Update()

            l2 = m1.LineSegments.Add()

            l2.Style = 3
            l2.X1 = 0
            l2.Y1 = -5
            l2.X2 = 0
            l2.Y2 = 5

            l1.Update()

            c1 = m1.Circles.Add()

            c1.Style = 3
            c1.Xc = 0
            c1.Yc = 0
            c1.Radius = 3

            c1.Update()

            m1.Name = "Ноль станка"
            m1.Update()

            if not path.lower().endswith('.frw'):
                path += '.frw'

            doc.SaveAs(path)
            print(f"Создан файл с нулевой точкой: {path}")

        except Exception as e:
            print(f"Ошибка при создании фрагмента: {e}")
            import traceback
            traceback.print_exc()

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
