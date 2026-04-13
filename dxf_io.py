# tsukasamiyashita/dxfflowchartmiya/DxfFlowchartMiya-1559b301efcf479ed8934d5e23d0d8559520cac5/dxf_io.py
import ezdxf
import math
from PyQt6.QtWidgets import QGraphicsLineItem, QGraphicsPathItem, QGraphicsEllipseItem
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QPainterPath, QPen, QColor
from graphics import NodeItem, EdgeItem, DxfBackgroundItem, clip_line_to_node

def export_dxf_file(scene, path, version_string):
    """シーン内のフローチャート要素をDXFファイルとしてエクスポートする"""
    base_bg = None
    for item in scene.items():
        if getattr(item, 'filepath', None) and type(item).__name__ == 'DxfBackgroundItem':
            base_bg = item
            break
            
    if base_bg and base_bg.filepath:
        try:
            doc = ezdxf.readfile(base_bg.filepath)
        except Exception:
            doc = ezdxf.new(version_string)
    else:
        doc = ezdxf.new(version_string)
        
    msp = doc.modelspace()
    
    for layer_name, color in [("FC_NODE", ezdxf.colors.BLUE), ("FC_EDGE", ezdxf.colors.GREEN), ("FC_TEXT", ezdxf.colors.WHITE)]:
        if layer_name not in doc.layers: doc.layers.add(layer_name, color=color)
    
    for item in scene.items():
        is_preview = False
        try:
            if item == getattr(scene, 'preview_node', None) or item in getattr(scene, 'preview_items', []): 
                is_preview = True
        except RuntimeError: pass
        
        # 背景レイヤーのDXF要素やプレビュー図形は出力対象外
        if is_preview or isinstance(item, DxfBackgroundItem): 
            continue
        
        if isinstance(item, NodeItem):
            x, y = item.scenePos().x(), -item.scenePos().y()
            t = item.node_type
            hw, hh = item.w / 2, item.h / 2
            if t == "process": 
                coords = [(x-hw, y+hh), (x+hw, y+hh), (x+hw, y-hh), (x-hw, y-hh)]
            elif t == "decision": 
                dw, dh = hw + 10, hh + 10
                coords = [(x, y+dh), (x+dw, y), (x, y-dh), (x-dw, y)]
            elif t == "data":
                skew = hh
                coords = [(x-hw+skew/2, y+hh), (x+hw+skew/2, y+hh), (x+hw-skew/2, y-hh), (x-hw-skew/2, y-hh)]
            elif t == "terminal":
                coords = [(x-hw+hh, y+hh), (x+hw-hh, y+hh), (x+hw, y+hh/2), (x+hw, y-hh/2), (x+hw-hh, y-hh), (x-hw+hh, y-hh), (x-hw, y-hh/2), (x-hw, y+hh/2)]
            else: 
                coords = [(x-hw, y+hh), (x+hw, y+hh), (x+hw, y-hh), (x-hw, y-hh)]
                
            msp.add_lwpolyline(coords, close=True, dxfattribs={'layer': 'FC_NODE'})
            if item.text_item.toPlainText():
                ls = item.text_item.toPlainText().split('\n')
                sy = y + (len(ls)-1)*7.5
                for i, l in enumerate(ls): 
                    msp.add_text(l, dxfattribs={'height': 12, 'layer': 'FC_TEXT'}).set_placement((x, sy-i*15), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
        elif isinstance(item, EdgeItem):
            pts = [item.source_node.scenePos()] + [wp.scenePos() for wp in item.waypoints] + [item.target_node.scenePos()]
            if len(pts) >= 2: 
                pts[0] = clip_line_to_node(pts[0], pts[1], item.source_node)
                pts[-1] = clip_line_to_node(pts[-1], pts[-2], item.target_node)
                
            for i in range(len(pts)-1): 
                msp.add_line((pts[i].x(), -pts[i].y()), (pts[i+1].x(), -pts[i+1].y()), dxfattribs={'layer': 'FC_EDGE'})
                
            if item.raw_text:
                pos = item.get_auto_text_pos() + (item.text_item.manual_offset if item.text_item.manual_offset else QPointF(0,0))
                mx, my = pos.x() + item.text_item.boundingRect().width()/2, -(pos.y() + item.text_item.boundingRect().height()/2)
                ls = item.raw_text.split('\n')
                sy = my + (len(ls)-1)*6
                for i, l in enumerate(ls): 
                    msp.add_text(l, dxfattribs={'height': 10, 'layer': 'FC_TEXT'}).set_placement((mx, sy-i*12), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
            if item.arrow in ["end", "both"] and hasattr(item, 'arrow_p1') and hasattr(item, 'arrow_p2'):
                angle = math.atan2(item.arrow_p2.y() - item.arrow_p1.y(), item.arrow_p2.x() - item.arrow_p1.x())
                arrow_size = 12 + item.line_width * 1.5
                ap1 = item.arrow_p2 - QPointF(math.cos(angle + math.pi / 6) * arrow_size, math.sin(angle + math.pi / 6) * arrow_size)
                ap2 = item.arrow_p2 - QPointF(math.cos(angle - math.pi / 6) * arrow_size, math.sin(angle - math.pi / 6) * arrow_size)
                msp.add_lwpolyline([(item.arrow_p2.x(), -item.arrow_p2.y()), (ap1.x(), -ap1.y()), (ap2.x(), -ap2.y())], close=True, dxfattribs={'layer': 'FC_EDGE'})
            
            if item.arrow in ["start", "both"] and hasattr(item, 'start_arrow_p1') and hasattr(item, 'start_arrow_p2'):
                angle = math.atan2(item.start_arrow_p2.y() - item.start_arrow_p1.y(), item.start_arrow_p2.x() - item.start_arrow_p1.x())
                arrow_size = 12 + item.line_width * 1.5
                ap1 = item.start_arrow_p2 - QPointF(math.cos(angle + math.pi / 6) * arrow_size, math.sin(angle + math.pi / 6) * arrow_size)
                ap2 = item.start_arrow_p2 - QPointF(math.cos(angle - math.pi / 6) * arrow_size, math.sin(angle - math.pi / 6) * arrow_size)
                msp.add_lwpolyline([(item.start_arrow_p2.x(), -item.start_arrow_p2.y()), (ap1.x(), -ap1.y()), (ap2.x(), -ap2.y())], close=True, dxfattribs={'layer': 'FC_EDGE'})
                
    doc.saveas(path)

def import_dxf_background(filepath):
    """DXFファイルを読み込み、背景レイヤー用のアイテム群を生成する"""
    try:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
    except Exception as e:
        raise Exception(f"DXFファイルの読み込みに失敗しました: {e}")
    
    bg_item = DxfBackgroundItem(filepath)
    pen = QPen(QColor(150, 150, 150, 180), 1) # 背景用に少し薄く表示
    
    entities = []
    for e in msp:
        if e.dxftype() == 'INSERT' and hasattr(e, 'virtual_entities'):
            try:
                entities.extend(e.virtual_entities())
            except Exception:
                pass
        else:
            entities.append(e)
            
    for e in entities:
        if e.dxftype() == 'LINE':
            item = QGraphicsLineItem(e.dxf.start.x, -e.dxf.start.y, e.dxf.end.x, -e.dxf.end.y)
            item.setPen(pen)
            bg_item.addToGroup(item)
            
        elif e.dxftype() in ('LWPOLYLINE', 'POLYLINE'):
            path = QPainterPath()
            pts = list(e.get_points('xy')) if hasattr(e, 'get_points') else [p for p in e.points()]
            if pts:
                path.moveTo(pts[0][0], -pts[0][1])
                for p in pts[1:]:
                    path.lineTo(p[0], -p[1])
                if e.closed:
                    path.closeSubpath()
                item = QGraphicsPathItem(path)
                item.setPen(pen)
                bg_item.addToGroup(item)
                
        elif e.dxftype() == 'CIRCLE':
            cx, cy = e.dxf.center.x, -e.dxf.center.y
            r = e.dxf.radius
            item = QGraphicsEllipseItem(cx - r, cy - r, r * 2, r * 2)
            item.setPen(pen)
            bg_item.addToGroup(item)
            
        # 複雑な図形（テキストやスプライン）は簡易背景用としてはスキップまたは近似対応とする
    
    return bg_item