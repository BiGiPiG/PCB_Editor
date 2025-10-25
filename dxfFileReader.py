import ezdxf
from postprocessor import Line

class DXFReader:

    def readFile(path):
    
        doc = ezdxf.readfile(path)
        
        msp = doc.modelspace()
        
        lines = []
        
        for e in msp:
            
            if e.dxftype() == "LINE":
    
                line = Line(e.dxf.start[0], e.dxf.start[1], e.dxf.end[0], e.dxf.end[1])
                
                lines.append(line)
            
        return lines