
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
        fx = contur[0].x1
        fy = contur[0].y1
    
        out += "G0 X" + str(fx) + " Y" + str(fy) + "\n" #Переходим в начало траектории
    
        out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
    
        out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
    
        #Цикл для перебора всех элементов контура
        for figure in contur:
            
            #Если начало новой фигуры не совпадает с предыдущей
            if figure.x1 != fx or figure.y1 != fy:
            
                out += "G0 Z" + str(overrun) + "\n" #Поднимаемся на безопасную высоту
                out += "G0 X" + str(figure.x1) + " Y" + str(figure.y1) + "\n" # Переходим к началу новой фигуры
                out += "G0 Z1\n" #Опускаемся к обрабатываемому материалу
                out += "G1 Z" + str(depth) + "\n" #Врезаемся в заготовку
            
            #Если фигура - отрезок
            if isinstance(figure, Line):
                
                #Линейная интерполяция
                out += "G1 X" + str(figure.x2) + " Y" + str(figure.y2) + "\n"
            
            #Если фигура - дуга
            if isinstance(figure, Arc):
            
                #Если по часовой
                if figure.dir:
            
                    #Круговая интерполяция по часовой стрелке
                    out += "G2 X" + str(figure.x2) + " Y" + str(figure.y2) + " I" + str(figure.x1 - figure.ox) + " J" + str(figure.y1 - figure.oy) + "\n"
                
                #Иначе
                else:
                    
                    #Круговая интерполяция против часовой стрелке
                    out += "G3 X" + str(figure.x2) + " Y" + str(figure.y2) + " I" + str(figure.x1 - figure.ox) + " J" + str(figure.y1 - figure.oy) + "\n"
               
            #Запоминаем конечную точку фигуры
            fx = figure.x2
            fy = figure.y2
    
        out += "G0 Z15\n" #Поднимаемся на безопасную высоту
    
        out += "M5\n"
    
        print(out)
        
        return out;


class Line:
    
    def __init__(self, x1, y1, x2, y2):
        self.x1 = x1
        self.y1 = y1
        self.x2 = x2
        self.y2 = y2
        
class Arc:
    
    def __init__(self, x1, y1, x2, y2, ox, oy, dir):
        self.x1 = x1
        self.y1 = y1
        self.x2 = x2
        self.y2 = y2
        self.ox = ox
        self.oy = oy
        self.dir = dir
      
      
traectory = []

traectory.append(Line(5, 5, 10, 10))
traectory.append(Line(10, 10, 20, 10))
traectory.append(Line(20, 10, 20, 30))
traectory.append(Arc(20, 30, 40, 50, 10, 10, True))
traectory.append(Line(80, 60, 120, 130))

Postprocessor.isolationTrajectory(traectory, -0.05, 2, 60)
