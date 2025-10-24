import os

import win32com.client
import pythoncom, win32com.client.connect, win32com.server.util
from pythoncom import VT_EMPTY
from win32com.client import gencache, VARIANT, Dispatch
from postprocessor import Hole, Postprocessor

class KompasService:
    def __init__(self):
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            self.kompas = None
            self.kompas_api7_module = None
            self.doc = None
            self.property_mng = None
            name = "Kompas.Application.7"

            try:
                self.kompas = win32com.client.Dispatch(name)
                self.kompas.Visible = True
                self.kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
                self.property_mng = self.kompas_api7_module.IPropertyMng(self.kompas)
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

            self.doc = self.kompas.Documents.Add(2)

            self.add_type_property()
            self.create_start_point()

            self.doc.SaveAs(path)
            print(f"Создан файл с нулевой точкой: {path}")

        except Exception as e:
            print(f"Ошибка при создании фрагмента: {e}")
            import traceback
            traceback.print_exc()

    def add_type_property(self):
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)
        if self.property_mng.GetProperty(doc2d, "Тип") is None:
            prop = self.property_mng.AddProperty(doc2d, VARIANT(VT_EMPTY, None))
            prop.Name = "Тип"
            prop.Update()


    def create_start_point(self):
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)

        macro = container.MacroObjects.Add()

        m1 = self.kompas_api7_module.IDrawingContainer(macro)

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

        l2.Update()

        c1 = m1.Circles.Add()

        c1.Style = 3
        c1.Xc = 0
        c1.Yc = 0
        c1.Radius = 3

        c1.Update()
        
        a1 = m1.Arcs.Add()

        a1.Style = 3
        a1.Xc = 0
        a1.Yc = 0
        a1.Angle1 = 0
        a1.Angle2 = 90
        a1.Direction = False
        a1.Radius = 3

        a1.Update()
        
        a2 = m1.Arcs.Add()

        a2.Style = 3
        a2.Xc = 0
        a2.Yc = 0
        a2.Angle1 = 180
        a2.Angle2 = 270
        a2.Direction = False
        a2.Radius = 3

        a2.Update()
        
        c1 = m1.Colourings.Add()
        
        b1 = self.kompas_api7_module.IBoundariesObject(c1)
        
        b1.AddBoundaries(l1, False)
        b1.AddBoundaries(l2, False)
        b1.AddBoundaries(a1, False)
        b1.AddBoundaries(a2, False)
        
        c1.Update()
        
        macro.AddDefaultHotPoint(0, 0)
        
        macro.Update()

        macro.Name = "Ноль станка"

        keeper = self.kompas_api7_module.IPropertyKeeper(macro)
        
        prop = self.property_mng.GetProperty(doc2d, "Тип")
        
        keeper.SetPropertyValue(prop, "Ноль станка", False)
        print("Ноль станка успешно создан")
        
    def draw_border(self, lines):
    
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)

        macro = container.MacroObjects.Add()
        macro.Name = "Границы"
        
        m = self.kompas_api7_module.IDrawingContainer(macro)
        
        for line in lines:
            
            lineSegment = m.LineSegments.Add()
            
            lineSegment.X1 = line.x1
            lineSegment.Y1 = line.y1
            lineSegment.X2 = line.x2
            lineSegment.Y2 = line.y2
            
            lineSegment.Update()
        
        macro.Update()
        
        prop = self.property_mng.GetProperty(doc2d, "Тип")

        keeper = self.kompas_api7_module.IPropertyKeeper(macro)
        keeper.SetPropertyValue(prop, "Границы", False)
        

    def create_holes(self, holes):
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)

        macro = container.MacroObjects.Add()
        macro.Name = "Отверстия"
        
        m1 = self.kompas_api7_module.IDrawingContainer(macro)

        for diam in holes:
            for hole in holes[diam]:
                l = m1.Circles.Add()
                l.Xc = hole[0]
                l.Yc = hole[1]
                l.Radius = diam / 2
                l.Style = 2
                l.Update()
                c = m1.Colourings.Add()
                c.Color1 = 255
                b = self.kompas_api7_module.IBoundariesObject(c)
                b.AddBoundaries(l, False)
                c.Update()

        macro.Update()
        
        prop = self.property_mng.GetProperty(doc2d, "Тип")

        keeper = self.kompas_api7_module.IPropertyKeeper(macro)
        keeper.SetPropertyValue(prop, "Отверстия", False)
        
    def create_drilling_program(self, startPoint, points, depth, overrun, feedrate):
        
        m = self.kompas_api7_module.IDrawingContainer(points)
        
        diams = m.Circles
        
        result = self.kompas_api7_module.IMacroObject(startPoint[0]).GetPlacement()
        
        dX = result[1]
        dY = result[2]
        
        result = self.kompas_api7_module.IMacroObject(points).GetPlacement()
        
        dX -= result[1]
        dY -= result[2]
        
        holes = []
        
        for diam in diams:
            
            
        
            diam = self.kompas_api7_module.ICircle(diam)
            holes.append(Hole(round(diam.Xc - dX, 4), round(diam.Yc - dY, 4), round(diam.Radius*2, 4)))
            
        return Postprocessor.drilling(holes, depth, overrun, feedrate)
        

    def get_macros(self):
        """Метод для получения макро объектов"""
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)
        macro_objects = container.MacroObjects

        macros = list()
        for macro in macro_objects:
            macros.append(macro)

        return macros

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
            self.doc = self.kompas.Documents.Open(path)
            self.add_type_property()
            if self.doc:
                print("Файл успешно открыт")
                return True
            else:
                print("Не удалось открыть файл")
                return False

        except Exception as e:
            print(f"Ошибка при открытии фрагмента: {e}")
            return False

    def find_macro_by_type(self, type):
        prop = self.property_mng.GetProperty(self.doc, "Тип")
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.kompas.ActiveDocument)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)
        macro_objects = container.MacroObjects

        return [obj for obj in macro_objects
            if self.kompas_api7_module.IPropertyKeeper(obj).GetPropertyValue(prop, None, True, True)[1] == type]

    def delete_macro(self, property):
        """метод для удаления макро объекта по его свойству"""
        print(f"delete_macro {property}")

    def rename_macro(self, property):
        """метод для переименования макро объекта по его свойству"""
        print(f"rename_macro {property}")

    def cleanup(self):
        """Очистка COM объектов"""
        try:
            if hasattr(self, 'kompas') and self.kompas:
                #self.kompas.Quit()
                self.kompas = None
            pythoncom.CoUninitialize()
        except:
            pass

    def __del__(self):
        self.cleanup()
   
    def advise_kompas_event(self, event, event_handler, event_source):

        app_notify_event = BaseEvent(event, event_handler, event_source)
        app_notify_event.advise()
       
class BaseEvent(object):

    _public_methods_ = ["__on_event"]

    def __init__(self, event, event_handler, event_source):

        self.__event = event
        self.__event_handler = event_handler
        self.__connection = None
        self.event_source = event_source

    def __del__(self):

        if not (self.__connection is None):
            self.__connection.Disconnect()
            del self.__connection

    def _invokeex_(self, command_id, locale_id, flags, params, result, exept_info):
        return self.__on_event(command_id, params)

    def _query_interface_(self, iid):
        if iid == self.__event.CLSID:
            return win32com.server.util.wrap(self)

    def advise(self):
        if self.__connection is None and not (self.event_source is None):
            self.__connection = win32com.client.connect.SimpleConnection(self.event_source, self, self.__event.CLSID)

    def unadvise(self):
        if self.__connection is not None and self.event_source is not None:
            self.__connection.Disconnect()
            self.__connection = None

    def __on_event(self, command_id, params):
        return self.__event_handler(command_id, params)
