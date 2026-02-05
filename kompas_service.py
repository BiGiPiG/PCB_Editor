import os

import win32com.client
import pythoncom, win32com.client.connect, win32com.server.util
from pythoncom import VT_EMPTY
from win32com.client import gencache, VARIANT, Dispatch
from postprocessor import Hole, Line, Arc, Postprocessor
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QProgressDialog, QMessageBox
from collections import defaultdict
from _thread import *
import time

class KompasService:
    def __init__(self):
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            self.kompas = None
            self.kompas_api7_module = None
            self.kompas5 = None
            self.kompas_api5_module = None
            self.doc = None
            self.property_mng = None
            name = "Kompas.Application.5"

            try:
                
                self.kompas_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
                self.kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
                
                self.kompas5 = self.kompas_api5_module.KompasObject(Dispatch(name)._oleobj_.QueryInterface(self.kompas_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
                self.kompas = self.kompas5.ksGetApplication7()
                
                self.kompas.Visible = True
                
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
               
    def draw_tracks(self, lines):
    
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)

        macro = container.MacroObjects.Add()
        macro.Name = "Дорожки"
        
        m = self.kompas_api7_module.IDrawingContainer(macro)
        
        pd = QProgressDialog("Operation in progress.", "Cancel", 0, len(lines))
        pd.setWindowModality(Qt.WindowModal)
        
        i = 0
        
        for line in lines:
            
            pd.setValue(i)
            i+=1
            
            lineSegment = m.LineSegments.Add()
            
            lineSegment.X1 = line.X1
            lineSegment.Y1 = line.Y1
            lineSegment.X2 = line.X2
            lineSegment.Y2 = line.Y2
            
            lineSegment.Update()
        
        macro.Update()
        
        prop = self.property_mng.GetProperty(doc2d, "Тип")

        keeper = self.kompas_api7_module.IPropertyKeeper(macro)
        keeper.SetPropertyValue(prop, "Дорожки", False)
               
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
        
    def draw_mask(self, lines):
    
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)

        macro = container.MacroObjects.Add()
        macro.Name = "Маска"
        
        m = self.kompas_api7_module.IDrawingContainer(macro)
        
        pd = QProgressDialog("Operation in progress.", "Cancel", 0, len(lines))
        pd.setWindowModality(Qt.WindowModal)
        
        i = 0
        
        for line in lines:
            
            pd.setValue(i)
            i+=1
            
            lineSegment = m.LineSegments.Add()
            
            lineSegment.Style = 2
            
            lineSegment.X1 = line.X1
            lineSegment.Y1 = line.Y1
            lineSegment.X2 = line.X2
            lineSegment.Y2 = line.Y2
            
            lineSegment.Update()
        
        macro.Update()
        
        prop = self.property_mng.GetProperty(doc2d, "Тип")

        keeper = self.kompas_api7_module.IPropertyKeeper(macro)
        keeper.SetPropertyValue(prop, "Маска", False)
        
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
        
    def create_mask_trajectory(self, startPoint, mask, offset):
    
        doc = self.kompas.Documents.Add(2, False)
        
        tempDoc = self.kompas.Documents.Add(2, True)
       
        doc2d = self.kompas_api7_module.IKompasDocument2D(doc)

        try:

            views = self.kompas_api7_module.IKompasDocument2D(tempDoc).ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)
            
            macro = container.MacroObjects.Add()
            macro.Name = "Траектория маски"
            
            container = self.kompas_api7_module.IDrawingContainer(macro)
            
            line = container.LineSegments.Add()
            
            line.X1 = 1
            line.Y1 = 1
            
            line.Update()
            
            macro.Update()

            (result, X, Y, angle, mirror) = self.kompas_api7_module.IMacroObject(mask).GetPlacement()

            views = doc2d.ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)

            macro_container = self.kompas_api7_module.IDrawingContainer(mask)
            
            lines = macro_container.LineSegments
            
            contours = self.split_into_contours(lines)
            print(contours)
            for i, cont in enumerate(contours):
            
                contour = container.DrawingContours.Add()
                icontour = self.kompas_api7_module.IContour(contour)
            
                arr = []
                
                for l in cont:
                
                    arr.append(l)
                
                icontour.CopySegments(arr, False)
                
                contour.Update()
                
                equidistant = container.Equidistants.Add()
                
                equidistant.BaseObject = contour
                equidistant.LeftRadius = offset
                equidistant.RightRadius = offset
                equidistant.Style = 6
                equidistant.Side = 0
                
                equidistant.Update()
                
                doc2d5 = self.kompas5.TransferInterface(doc, 1, 0)
                
                g1 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Side = 1
                equidistant.Update()
                
                g2 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Delete()
                
                g17 = self.kompas5.TransferReference(g1, doc.Reference)
                g27 = self.kompas5.TransferReference(g2, doc.Reference)
                
                g171 = self.kompas_api7_module.IDrawingObject1(g17.Objects(0))
                l1 = g171.GetCurve2D().Length
                
                g271 = self.kompas_api7_module.IDrawingObject1(g27.Objects(0))
                l2 = g271.GetCurve2D().Length
                
                g = None
                
                if l1 > l2:
                    g = g17
                    g27.Delete()
                else:
                    g = g27
                    g17.Delete()
                
                contour.Delete();
                
                doc2d5.ksMoveObj(g.Reference, X, Y)
                
                #doc2d5.ksDestroyObjects(g.Reference)
                
                arr = container.Objects(0)
                
                g.AddObjects(arr)
                
                doc2d5.ksWriteGroupToClip(g.Reference, 0)
                
                doc2d5 = self.kompas5.TransferInterface(tempDoc, 1, 0)
                
                group = doc2d5.ksReadGroupFromClip()
                
                doc2d5.ksStoreTmpGroup(group)
                
                group = self.kompas5.TransferReference(group, tempDoc.Reference)
                
                arr = group.Objects()
                print(arr)
                macro.AddObjects(group)
            
            prop = self.property_mng.GetProperty(tempDoc, "Тип")

            keeper = self.kompas_api7_module.IPropertyKeeper(macro)
            keeper.SetPropertyValue(prop, "Траектория маски", False)
            
            line.Delete()
            
            macro.Update()
            
        except ValueError as e:
            QMessageBox.warning(None, "Ошибка построения", e)
        
        border = self.find_macro_by_type("Границы")
        
        doc.Close(0)
        
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
       
    def create_border_trajectory(self, startPoint, border, tool_diameter, offset, contour_type):
       
        doc = self.kompas.Documents.Add(2, False)
       
        doc2d = self.kompas_api7_module.IKompasDocument2D(doc)

        try:

            views = self.kompas_api7_module.IKompasDocument2D(self.doc).ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)
            
            macro = container.MacroObjects.Add()
            macro.Name = "Траектория границ"
            
            container = self.kompas_api7_module.IDrawingContainer(macro)
            
            line = container.LineSegments.Add()
            
            line.X1 = 1
            line.Y1 = 1
            
            line.Update()
            
            macro.Update()

            (result, X, Y, angle, mirror) = self.kompas_api7_module.IMacroObject(border).GetPlacement()

            views = doc2d.ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)

            macro_container = self.kompas_api7_module.IDrawingContainer(border)
            
            
            
            lines = macro_container.LineSegments
            
            contours = self.split_into_contours(lines)
            print(contours)
            for i, cont in enumerate(contours):
            
                contour = container.DrawingContours.Add()
                icontour = self.kompas_api7_module.IContour(contour)
            
                arr = []
                
                for l in cont:
                
                    arr.append(l)
                
                icontour.CopySegments(arr, False)
                
                contour.Update()
                
                equidistant = container.Equidistants.Add()
                
                equidistant.BaseObject = contour
                equidistant.LeftRadius = tool_diameter/2 + offset
                equidistant.RightRadius = tool_diameter/2 + offset
                equidistant.Style = 6
                equidistant.Side = 0
                
                equidistant.Update()
                
                doc2d5 = self.kompas5.TransferInterface(doc, 1, 0)
                
                g1 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Side = 1
                equidistant.Update()
                
                g2 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Delete()
                
                g17 = self.kompas5.TransferReference(g1, doc.Reference)
                g27 = self.kompas5.TransferReference(g2, doc.Reference)
                
                g171 = self.kompas_api7_module.IDrawingObject1(g17.Objects(0))
                l1 = g171.GetCurve2D().Length
                
                g271 = self.kompas_api7_module.IDrawingObject1(g27.Objects(0))
                l2 = g271.GetCurve2D().Length
                
                g = None
                
                if contour_type == "Внешний":
                    if l1 > l2:
                        g = g17
                        g27.Delete()
                    else:
                        g = g27
                        g17.Delete()
                else:
                    if l1 < l2:
                        g = g17
                        g27.Delete()
                    else:
                        g = g27
                        g17.Delete()
                
                contour.Delete();
                
                doc2d5.ksMoveObj(g.Reference, X, Y)
                
                #doc2d5.ksDestroyObjects(g.Reference)
                
                arr = container.Objects(0)
                
                g.AddObjects(arr)
                
                doc2d5.ksWriteGroupToClip(g.Reference, 0)
                
                doc2d5 = self.kompas5.TransferInterface(self.doc, 1, 0)
                
                group = doc2d5.ksReadGroupFromClip()
                
                doc2d5.ksStoreTmpGroup(group)
                
                group = self.kompas5.TransferReference(group, self.doc.Reference)
                
                arr = group.Objects()
                print(arr)
                macro.AddObjects(group)
            
            prop = self.property_mng.GetProperty(self.doc, "Тип")

            keeper = self.kompas_api7_module.IPropertyKeeper(macro)
            keeper.SetPropertyValue(prop, "Траектория границ", False)
            
            line.Delete()
            
            macro.Update()
            
        except ValueError as e:
            QMessageBox.warning(None, "Ошибка построения", e)
        
        doc.Close(0)
        
    def create_tracks_trajectory(self, startPoint, border, tool_diameter, line_count, overlap_percent):
       
        doc = self.kompas.Documents.Add(2, False)
       
        doc2d = self.kompas_api7_module.IKompasDocument2D(doc)

        try:

            views = self.kompas_api7_module.IKompasDocument2D(self.doc).ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)
            
            macro = container.MacroObjects.Add()
            macro.Name = "Траектория дорожек"
            
            container = self.kompas_api7_module.IDrawingContainer(macro)
            
            line = container.LineSegments.Add()
            
            line.X1 = 1
            line.Y1 = 1
            
            line.Update()
            
            macro.Update()

            (result, X, Y, angle, mirror) = self.kompas_api7_module.IMacroObject(border).GetPlacement()

            views = doc2d.ViewsAndLayersManager.Views.View(0)
            container = self.kompas_api7_module.IDrawingContainer(views)

            macro_container = self.kompas_api7_module.IDrawingContainer(border)
            
            
            
            lines = macro_container.LineSegments
            
            contours = self.split_into_contours(lines)
            print(contours)
            for i, cont in enumerate(contours):
            
                contour = container.DrawingContours.Add()
                icontour = self.kompas_api7_module.IContour(contour)
            
                arr = []
                
                for l in cont:
                
                    arr.append(l)
                
                icontour.CopySegments(arr, False)
                
                contour.Update()
                
                equidistant = container.Equidistants.Add()
                
                equidistant.BaseObject = contour
                equidistant.LeftRadius = tool_diameter/2
                equidistant.RightRadius = tool_diameter/2
                equidistant.Style = 6
                equidistant.Side = 0
                
                equidistant.Update()
                
                doc2d5 = self.kompas5.TransferInterface(doc, 1, 0)
                
                g1 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Side = 1
                equidistant.Update()
                
                g2 = doc2d5.ksApproximationCurve(equidistant.Reference, 0.005, False, 0, 0)
                
                equidistant.Delete()
                
                g17 = self.kompas5.TransferReference(g1, doc.Reference)
                g27 = self.kompas5.TransferReference(g2, doc.Reference)
                
                g171 = self.kompas_api7_module.IDrawingObject1(g17.Objects(0))
                l1 = g171.GetCurve2D().Length
                
                g271 = self.kompas_api7_module.IDrawingObject1(g27.Objects(0))
                l2 = g271.GetCurve2D().Length
                
                g = None
                
                if l1 > l2:
                    g = g17
                    g27.Delete()
                else:
                    g = g27
                    g17.Delete()
                
                contour.Delete();
                
                doc2d5.ksMoveObj(g.Reference, X, Y)
                
                #doc2d5.ksDestroyObjects(g.Reference)
                
                arr = container.Objects(0)
                
                g.AddObjects(arr)
                
                doc2d5.ksWriteGroupToClip(g.Reference, 0)
                
                doc2d5 = self.kompas5.TransferInterface(self.doc, 1, 0)
                
                group = doc2d5.ksReadGroupFromClip()
                
                doc2d5.ksStoreTmpGroup(group)
                
                group = self.kompas5.TransferReference(group, self.doc.Reference)
                
                arr = group.Objects()
                print(arr)
                macro.AddObjects(group)
            
            prop = self.property_mng.GetProperty(self.doc, "Тип")

            keeper = self.kompas_api7_module.IPropertyKeeper(macro)
            keeper.SetPropertyValue(prop, "Траектория дорожек", False)
            
            line.Delete()
            
            macro.Update()
            
        except ValueError as e:
            QMessageBox.warning(None, "Ошибка построения", e)
        
        doc.Close(0)
     
    def create_milling_program(self, startPoint, macro, depth, overrun, feedrate):
        
        result = self.kompas_api7_module.IMacroObject(startPoint[0]).GetPlacement()
        
        dX = result[1]
        dY = result[2]
        
        result = self.kompas_api7_module.IMacroObject(macro).GetPlacement()
        
        dX -= result[1]
        dY -= result[2]
        
        contours = self.kompas_api7_module.IDrawingContainer(macro).Objects(0)
        
        segments = []
        
        ths = []
        
        j = 0
        
        for contour in contours:
            
            ths.append(None)
            
            start_new_thread(self.add_contur_segments, (contour.Reference, j, ths, self.doc.Reference, dX, dY))
            
            j+=1

        _next = True
        
        while _next:
            
            _next = False
            
            for i in range(0, len(ths)):
                if  ths[i] == None:
                    _next = True
            
            time.sleep(2)
        
        for i in range(0, len(ths)):
        
            segments = segments + ths[i]
        
        return Postprocessor.cutTrajectory(segments, depth, overrun, feedrate, 0)
     
    def add_contur_segments(self, contour, j, ths, doc, dX, dY):
        
        kompas5 = self.kompas_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.kompas_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
        
        segments = []
        
        contour = kompas5.TransferReference(contour, doc)
        
        contour = self.kompas_api7_module.IContour(contour)
            
        for i in range(0, contour.Count):
        
            segment = contour.Segment(i)
        
            segment_type = segment.Type
        
            if segment_type == 10707:
                
                line = self.kompas_api7_module.IContourLineSegment(segment)
                
                segments.append(Line(round(line.X1 - dX, 5), round(line.Y1 - dY, 5), round(line.X2 - dX, 5), round(line.Y2 - dY, 5)))
                
            elif segment_type == 10708:
            
                arc = self.kompas_api7_module.IContourArc(segment)
                
                segments.append(Arc(round(arc.X1 - dX, 5), round(arc.Y1 - dY, 5), round(arc.X2 - dX, 5), round(arc.Y2 - dY, 5), round(arc.Xc, 5), round(arc.Yc, 5), arc.Direction, round(arc.Radius, 5)))
        
        ths[j] = segments
     
    def split_into_contours(self, lines):
        """
        Разделяет коллекцию линий на замкнутые контуры.
        
        Args:
            lines: список объектов/словарей с полями X1, Y1, X2, Y2
        
        Returns:
            list of lists: каждый подсписок — линии одного контура
        """
        
        # 1. Строим граф связности: точка → список смежных точек и соответствующих линий
        graph = defaultdict(list)  # {(x, y): [(соседняя_точка, линия), ...]}
        line_to_points = {}  # для обратной связи: линия → её точки
        
        for line in lines:
        
            line = self.kompas_api7_module.ILineSegment(line)
            
            p1 = (line.X1, line.Y1)
            p2 = (line.X2, line.Y2)
            
            graph[p1].append((p2, line))
            graph[p2].append((p1, line))
            line_to_points[line.Reference] = (p1, p2)
        
        contours = []
        used_lines = []  # уже использованные линии
        
        # 2. Обходим граф, собирая контуры
        for line in lines:
        
            line = self.kompas_api7_module.ILineSegment(line)
        
            if line in used_lines:
                continue
                
            # Начинаем новый контур с этой линии
            contour = []
            current_line = line
            start_point = (current_line.X1, current_line.Y1)
            current_point = (current_line.X2, current_line.Y2)
            
            contour.append(current_line)
            used_lines.append(current_line)
            
            # 3. Двигаемся по контуру, пока не вернёмся в начало
            while current_point != start_point:
                # Ищем следующую линию, выходящую из current_point
                next_line = None
                next_point = None
                
                for neighbor, line_candidate in graph[current_point]:
                    if line_candidate not in used_lines:
                        next_line = line_candidate
                        next_point = neighbor
                        break
                
                if next_line is None:
                    # Не удалось продолжить контур (разрыв или ошибка данных)
                    break
                    
                contour.append(next_line)
                used_lines.append(next_line)
                current_point = next_point
            
            # Только если вернулись в старт — это замкнутый контур
            if current_point == start_point:
                contours.append(contour)
        
        return contours
       
    def select_macro(self, macro):
        print("select")
        document2D1 = self.kompas_api7_module.IKompasDocument2D1(self.doc)
    
        selectionManager = document2D1.SelectionManager
        selectionManager.UnselectAll()
        selectionManager.Select(macro)

    def get_macros(self):
        """Метод для получения макро объектов"""
        doc2d = self.kompas_api7_module.IKompasDocument2D(self.doc)

        views = doc2d.ViewsAndLayersManager.Views.View(0)

        container = self.kompas_api7_module.IDrawingContainer(views)
        macro_objects = container.MacroObjects

        return [macro for macro in macro_objects]

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
        property.Delete()

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
