from shapely.geometry import Polygon, MultiPolygon, Point, LineString
from PySide6.QtCore import QAbstractListModel, Qt, QModelIndex, QSize, QRect, QMargins, Signal, QThread, Signal, Slot, QObject, QMetaObject, QRectF
from PySide6.QtCore import QPoint, QPointF
from shapely.ops import unary_union
import math
def normalizeCoords(height, zoom_factor, x0, y0, x1, y1, page):
    x0, y0, x1, y1 = (
        x0 * zoom_factor-2,
        y0 * zoom_factor,
        x1 * zoom_factor+2,
        y1 * zoom_factor,
    )
    y0 += (height + 5) * page
    y1 += (height + 5) * page
    
    rect_polygon = Polygon([
                    (x0, y0),
                    (x1, y0),
                    (x1, y1),
                    (x0, y1)
                ])
    
    return x0, y0, x1, y1, rect_polygon
    
#def getPolygons() 
#def createP

def shapelyPolygonToQPolygonF(self, polygon):
    qpolygon = QPolygonF()
    if not polygon.is_empty:
        for x, y in polygon.exterior.coords:
            qpolygon.append(QPointF(x, y))
    return qpolygon





def checkAndJoinPolygons(self, clusters, newcluster):
    removes = []
    intersects = []
    for cluster in clusters:                
        if(newcluster.intersects(cluster)):
            newcluster = unary_union([cluster, newcluster])
            removes.append(cluster)
        elif(len(removes)>0): break
    if len(removes)==0:
        clusters.append(newcluster)
    else:
        for rem in removes:
            clusters.remove(rem)
        clusters.append(newcluster)
        





def calculate_distance(self, point1, point2):
    x1, y1 = point1
    x2, y2 = point2
    distance = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
    return distance      

def find_closest_spans(height, zoom_factor, pageinitsel, pageendsel, mousepos1, mousepos2):
    closest_span1 = None
    closest_span1cont = None
    closest_span1page = None
    closest_span2 = None
    closest_span2cont = None
    closest_span2page = None
    min_distance = float('inf')
    min_distance2 = float('inf')
    ponto1 = Point([mousepos1.x, mousepos1.y])
    ponto2 = Point([mousepos2.x, mousepos2.y])
    inispan = [None, -1]
    endspan = [None, -1]
    cont = 0
    for page in range(pageinitsel, pageendsel + 1):
        for block in self.pixmaps[page][1]['blocks']:
            if block['type'] != 0:
                continue
            x0b, y0b, x1b, y1b, rect_polygon_b = normalizeCoords(height, zoom_factor, *block['bbox'], page)
            if(rect_polygon_b.contains(ponto1) or rect_polygon_b.contains(ponto2)):
                for line in block['lines']:
                    for span in line['spans']:
                        x0s, y0s, _, _, rect_polygon_s = normalizeCoords(height, zoom_factor, *span['bbox'], page)
                        if(rect_polygon_s.contains(ponto1)):
                            inispan = [span, cont, page]
                        if(rect_polygon_s.contains(ponto2)):
                            endspan = [span, cont, page]
                        if(inispan[0]!=None and endspan[0]!=None):
                            return inispan[0], inispan[1], inispan[2], endspan[0], endspan[1], endspan[2]
                        cont += 1
            else:
                for line in block['lines']:
                    cont += len(line['spans'])
            distance = min(self.calculate_distance((mousepos1.x, mousepos1.y), ((x0b+x1b)/2, (y0b+y1b)/2)),
                            self.calculate_distance((mousepos1.x, mousepos1.y), (x0b, y0b)),
                            self.calculate_distance((mousepos1.x, mousepos1.y), (x1b, y1b)),
                            self.calculate_distance((mousepos1.x, mousepos1.y), (x0b, y1b)),
                            self.calculate_distance((mousepos1.x, mousepos1.y), (x1b, y0b)),
                            )
            distance2 = min(self.calculate_distance((mousepos2.x, mousepos2.y), ((x0b+x1b)/2, (y0b+y1b)/2)),
                            self.calculate_distance((mousepos2.x, mousepos2.y), (x0b, y0b)),
                            self.calculate_distance((mousepos2.x, mousepos2.y), (x1b, y1b)),
                            self.calculate_distance((mousepos2.x, mousepos2.y), (x0b, y1b)),
                            self.calculate_distance((mousepos2.x, mousepos2.y), (x1b, y0b)),
                            )
            if inispan[0]==None and distance < min_distance:
                closest_span1page = page
                min_distance = distance
                closest_span1 = block
                closest_span1cont = cont
            #print(distance2)
            if endspan[0]==None and distance2 < min_distance2:
                min_distance2 = distance2
                closest_span2 = block
                closest_span2cont = cont
                closest_span2page = page 
        if(inispan[0]==None):
            cont = closest_span1cont
            min_distance = float('inf')
            for line in closest_span1['lines']:
                for span in line['spans']:
                    x0s, y0s, _, _, rect_polygon_s = self.normalizeCoords(*span['bbox'], page)
                    distance = self.calculate_distance((mousepos1.x, mousepos1.y), (x0s, y0s))
                    if distance < min_distance:
                        min_distance = distance
                        closest_span1 = span
                        closest_span1cont = cont
                    cont += 1
        if(endspan[0]==None):
            cont = closest_span2cont
            min_distance2 = float('inf')
            for line in closest_span2['lines']:
                for span in line['spans']:
                    x0s, y0s, x1s, y1s, rect_polygon_s = self.normalizeCoords(*span['bbox'], closest_span2page)
                    distance = self.calculate_distance((mousepos2.x, mousepos2.y), (x0s, y0s))
                    if distance < min_distance2:
                        min_distance2 = distance
                        closest_span2 = span
                        closest_span2cont = cont
                    cont += 1
    if(inispan[0]==None):                
        inispan = [closest_span1, closest_span1cont, closest_span1page]
    if(endspan[0]==None):
        endspan = [closest_span2, closest_span2cont, closest_span2page]
        
    return inispan[0], inispan[1], inispan[2], endspan[0], endspan[1], endspan[2]