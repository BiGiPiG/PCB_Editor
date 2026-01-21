
class Postprocessor:

    def isolationTrajectory(contur, depth, overrun, feedrate):
    
        out = "G21\n" #Еденицы измерения ММ
        out += "G90\n" #Абсолютные кординаты
        out += "G94\n" #Режим подачи мм/мин
        
        out += "G1 F" + str(feedrate) + "\n" #Задаем подачу
        
        out += "G0 Z15\n" #Поднимаемся на безопасную высоту
        
        out += "M0\n" 
        
        out += "M3\n" 
        
        #Запоминаем начальные кординаты
        fx = contur[0].X1
        fy = contur[0].Y1
    
        out += "G0 X" + str(fx) + " Y" + str(fy) + "\n" #Переходим в начало траектории
    
        out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
    
        out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
    
        #Цикл для перебора всех элементов контура
        for figure in contur:
            
            #Если начало новой фигуры не совпадает с предыдущей
            if figure.X1 != fx or figure.Y1 != fy:
            
                out += "G0 Z" + str(overrun) + "\n" #Поднимаемся на безопасную высоту
                out += "G0 X" + str(figure.X1) + " Y" + str(figure.Y1) + "\n" # Переходим к началу новой фигуры
                out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
                out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
            
            #Если фигура - отрезок
            if isinstance(figure, Line):
                
                #Линейная интерполяция
                out += "G1 X" + str(figure.X2) + " Y" + str(figure.Y2) + "\n"
            
            #Если фигура - дуга
            if isinstance(figure, Arc):
            
                #Если по часовой
                if figure.dir:
            
                    #Круговая интерполяция по часовой стрелке
                    out += "G2 X" + str(figure.X2) + " Y" + str(figure.Y2) + " I" + str(figure.X1 - figure.ox) + " J" + str(figure.Y1 - figure.oy) + "\n"
                
                #Иначе
                else:
                    
                    #Круговая интерполяция против часовой стрелке
                    out += "G3 X" + str(figure.X2) + " Y" + str(figure.Y2) + " I" + str(figure.X1 - figure.ox) + " J" + str(figure.Y1 - figure.oy) + "\n"
               
            #Запоминаем конечную точку фигуры
            fx = figure.X2
            fy = figure.Y2
    
        out += "G0 Z15\n" #Поднимаемся на безопасную высоту
    
        out += "M5\n"
        
        return out;
        
        
    def cutTrajectory(contur, depth, overrun, feedrate, depthPerPass):
    
        out = "G21\n" #Еденицы измерения ММ
        out += "G90\n" #Абсолютные кординаты
        out += "G94\n" #Режим подачи мм/мин
        
        out += "G1 F" + str(feedrate) + "\n" #Задаем подачу
        
        out += "G0 Z15\n" #Поднимаемся на безопасную высоту
        
        out += "M0\n" 
        
        out += "M3\n" 
        
        #Запоминаем начальные кординаты
        fx = contur[0].X1
        fy = contur[0].Y1
    
        out += "G0 X" + str(fx) + " Y" + str(fy) + "\n" #Переходим в начало траектории
    
        out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
    
        out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
    
        #Цикл для перебора всех элементов контура
        for figure in contur:
            
            #Если начало новой фигуры не совпадает с предыдущей
            if figure.X1 != fx or figure.Y1 != fy:
            
                out += "G0 Z" + str(overrun) + "\n" #Поднимаемся на безопасную высоту
                out += "G0 X" + str(figure.X1) + " Y" + str(figure.Y1) + "\n" # Переходим к началу новой фигуры
                out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
                out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
            
            #Если фигура - отрезок
            if isinstance(figure, Line):
                
                #Линейная интерполяция
                out += "G1 X" + str(figure.X2) + " Y" + str(figure.Y2) + "\n"
            
            #Если фигура - дуга
            if isinstance(figure, Arc):
            
                #Если по часовой
                if figure.dir:
            
                    #Круговая интерполяция по часовой стрелке
                    #out += "G3 X" + str(figure.X2) + " Y" + str(figure.Y2) + " I" + str(figure.X1 - figure.ox) + " J" + str(figure.Y1 - figure.oy) + "\n"
                    out += "G2 X" + str(figure.X2) + " Y" + str(figure.Y2) + " R" + str(figure.r) + "\n"
                
                #Иначе
                else:
                    
                    #Круговая интерполяция против часовой стрелке
                    #out += "G2 X" + str(figure.X2) + " Y" + str(figure.Y2) + " I" + str(figure.X1 - figure.ox) + " J" + str(figure.Y1 - figure.oy) + "\n"
                    out += "G3 X" + str(figure.X2) + " Y" + str(figure.Y2) + " R" + str(figure.r) + "\n"
               
            #Запоминаем конечную точку фигуры
            fx = figure.X2
            fy = figure.Y2
    
        out += "G0 Z15\n" #Поднимаемся на безопасную высоту
    
        out += "M5\n"
        
        return out;
        
        
    def drilling(points, depth, overrun, feedrate):
        
        out = "G21\n" #Еденицы измерения ММ
        out += "G90\n" #Абсолютные кординаты
        out += "G94\n" #Режим подачи мм/мин
        
        out += "G1 F" + str(feedrate) + "\n" #Задаем подачу
        
        out += "G0 Z30\n" #Поднимаемся на безопасную высоту
        
        out += "M0\n" 
        
        out += "M3\n" 
        
        holes = dict()
        
        for point in points:
        
            if point.diam in holes:
            
                holes[point.diam].append(point)
                
            else:
            
                holes[point.diam] = [point]
        
        count = 0
        
        for key in holes:
        
            out += "(tool " + str(key) + " diam)\n"
        
            for point in holes[key]:
            
                out += "G0 X" + str(point.x) + " Y" + str(point.y) + "\n" # Переходим к следующему отверстию
                out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
                out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
                out += "G0 Z" + str(overrun) + "\n" #Поднимаемся на безопасную высоту
             
            count += 1
             
            out += "G0 Z30\n" #Поднимаемся на безопасную высоту
            
            if count != len(holes):
                out += "M0\n" 
             
        out += "M5\n"
             
        return out


class Line: #Линия
    
    def __init__(self, X1, Y1, X2, Y2):
        self.X1 = X1
        self.Y1 = Y1
        self.X2 = X2
        self.Y2 = Y2
        
class Arc: #Дуга
    
    def __init__(self, X1, Y1, X2, Y2, ox, oy, dir, r):
        self.X1 = X1
        self.Y1 = Y1
        self.X2 = X2
        self.Y2 = Y2
        self.ox = ox
        self.oy = oy
        self.dir = dir
        self.r = r
        
class Hole: #Отверстие

    def __init__(self, x, y, diam):
        self.x = x
        self.y = y
        self.diam = diam
        
      
# Удалить после отладки
#traectory = []

#traectory.append(Line(5, 5, 10, 10))
#traectory.append(Line(10, 10, 20, 10))
#traectory.append(Line(20, 10, 20, 30))
#traectory.append(Arc(20, 30, 40, 50, 10, 10, True))
#traectory.append(Line(80, 60, 120, 130))

#print(Postprocessor.isolationTrajectory(traectory, -0.05, 2, 60)) #Вывод комманд

#traectory.append(Hole(0, 0, 0.8))
#traectory.append(Hole(10, 10, 0.8))
#traectory.append(Hole(20, 20, 1))
#traectory.append(Hole(10, 20, 1))
#traectory.append(Hole(5, 40, 0.8))

#print(Postprocessor.drilling(traectory, -0.05, 2, 60))
