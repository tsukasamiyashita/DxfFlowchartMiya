import ezdxf
import math
from PyQt6.QtWidgets import QGraphicsItemGroup
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QColor, QPainterPath, QBrush
import re
from ezdxf.path import Command
from graphics import (NodeItem, EdgeItem, clip_line_to_node, 
                       CadLineItem, CadPolylineItem, CadEllipseItem, CadArcItem, CadTextItem, CadPathItem, CadHatchItem)

def to_qpainterpath(path):
    """ezdxf.path.Path を PyQt6.QtGui.QPainterPath に変換する"""
    qpath = QPainterPath()
    if len(path) == 0: return qpath
    
    qpath.moveTo(path.start.x, -path.start.y)
    for cmd in path:
        if cmd.type == Command.LINE_TO:
            qpath.lineTo(cmd.end.x, -cmd.end.y)
        elif cmd.type == Command.CURVE3_TO:
            qpath.quadTo(cmd.ctrl.x, -cmd.ctrl.y, cmd.end.x, -cmd.end.y)
        elif cmd.type == Command.CURVE4_TO:
            qpath.cubicTo(cmd.ctrl1.x, -cmd.ctrl1.y, cmd.ctrl2.x, -cmd.ctrl2.y, cmd.end.x, -cmd.end.y)
    if path.is_closed:
        qpath.closeSubpath()
    return qpath

def export_scene_to_dxf(scene, path, version_string):
    doc = ezdxf.new(version_string)
    msp = doc.modelspace()
    
    for layer_name, color in [("FC_NODE", ezdxf.colors.BLUE), ("FC_EDGE", ezdxf.colors.GREEN), ("FC_TEXT", ezdxf.colors.WHITE)]:
        if layer_name not in doc.layers: doc.layers.add(layer_name, color=color)
        
    def process_item(item):
        if not item.isVisible(): return
        
        is_preview = False
        try:
            if item == getattr(scene, 'preview_node', None) or item in getattr(scene, 'preview_items', []): 
                is_preview = True
        except RuntimeError: pass
        if is_preview: return

        if type(item) == QGraphicsItemGroup:
            for child in item.childItems():
                process_item(child)
            return

        trans = item.sceneTransform()

        if isinstance(item, NodeItem):
            hw, hh = item.w / 2, item.h / 2
            t = item.node_type
            
            if t == "process": 
                coords = [QPointF(-hw, hh), QPointF(hw, hh), QPointF(hw, -hh), QPointF(-hw, -hh)]
            elif t == "decision": 
                dw, dh = hw + 10, hh + 10
                coords = [QPointF(0, dh), QPointF(dw, 0), QPointF(0, -dh), QPointF(-dw, 0)]
            elif t == "data":
                skew = hh
                coords = [QPointF(-hw+skew/2, hh), QPointF(hw+skew/2, hh), QPointF(hw-skew/2, -hh), QPointF(-hw-skew/2, -hh)]
            elif t == "terminal":
                coords = [QPointF(-hw+hh, hh), QPointF(hw-hh, hh), QPointF(hw, hh/2), QPointF(hw, -hh/2), QPointF(hw-hh, -hh), QPointF(-hw+hh, -hh), QPointF(-hw, -hh/2), QPointF(-hw, hh/2)]
            else: 
                coords = [QPointF(-hw, hh), QPointF(hw, hh), QPointF(hw, -hh), QPointF(-hw, -hh)]
                
            scene_coords = [trans.map(p) for p in coords]
            dxf_coords = [(p.x(), -p.y()) for p in scene_coords]
            msp.add_lwpolyline(dxf_coords, close=True, dxfattribs={'layer': 'FC_NODE'})
            
            if item.text_item.toPlainText():
                ls = item.text_item.toPlainText().split('\n')
                base_h = 12
                line_spacing = 15
                for i, l in enumerate(ls): 
                    ly = (len(ls)-1)*(-line_spacing/2) + i*line_spacing
                    sp = trans.map(QPointF(0, ly))
                    msp.add_text(l, dxfattribs={'height': base_h, 'layer': 'FC_TEXT'}).set_placement((sp.x(), -sp.y()), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
        elif isinstance(item, EdgeItem):
            pts = [item.source_node.scenePos()] + [wp.scenePos() for wp in item.waypoints] + [item.target_node.scenePos()]
            p0 = clip_line_to_node(pts[0], pts[1], item.source_node) if len(pts)>1 else pts[0]
            pn = clip_line_to_node(pts[-1], pts[-2], item.target_node) if len(pts)>1 else pts[-1]
            
            dxf_pts = [(p0.x(), -p0.y())]
            for wp in item.waypoints:
                dxf_pts.append((wp.scenePos().x(), -wp.scenePos().y()))
            dxf_pts.append((pn.x(), -pn.y()))
            
            for i in range(len(dxf_pts)-1):
                msp.add_line(dxf_pts[i], dxf_pts[i+1], dxfattribs={'layer': 'FC_EDGE'})
                
            if item.raw_text:
                pos = item.get_auto_text_pos() + (item.text_item.manual_offset if item.text_item.manual_offset else QPointF(0,0))
                mx = pos.x() + item.text_item.boundingRect().width()/2
                my = pos.y() + item.text_item.boundingRect().height()/2
                ls = item.raw_text.split('\n')
                base_h = 10
                line_spacing = 12
                for i, l in enumerate(ls): 
                    ly = my - (len(ls)-1)*(line_spacing/2) + i*line_spacing
                    msp.add_text(l, dxfattribs={'height': base_h, 'layer': 'FC_TEXT'}).set_placement((mx, -ly), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
            def draw_arrow(p1, p2):
                angle = math.atan2(p2.y() - p1.y(), p2.x() - p1.x())
                arrow_size = 12 + item.line_width * 1.5
                ap1 = p2 - QPointF(math.cos(angle + math.pi / 6) * arrow_size, math.sin(angle + math.pi / 6) * arrow_size)
                ap2 = p2 - QPointF(math.cos(angle - math.pi / 6) * arrow_size, math.sin(angle - math.pi / 6) * arrow_size)
                msp.add_lwpolyline([(p2.x(), -p2.y()), (ap1.x(), -ap1.y()), (ap2.x(), -ap2.y())], close=True, dxfattribs={'layer': 'FC_EDGE'})

            if item.arrow in ["end", "both"] and hasattr(item, 'arrow_p1') and hasattr(item, 'arrow_p2'):
                draw_arrow(item.arrow_p1, item.arrow_p2)
            if item.arrow in ["start", "both"] and hasattr(item, 'start_arrow_p1') and hasattr(item, 'start_arrow_p2'):
                draw_arrow(item.start_arrow_p1, item.start_arrow_p2)

        elif hasattr(item, 'is_cad_item') and item.is_cad_item:
            layer = item.cad_layer if hasattr(item, 'cad_layer') else "0"
            if layer not in doc.layers:
                doc.layers.add(layer)
            
            dxf_attribs = {'layer': layer}
            r, g, b = item.cad_color.red(), item.cad_color.green(), item.cad_color.blue()
            dxf_attribs['true_color'] = ezdxf.colors.rgb2int((r, g, b))
            
            if isinstance(item, CadLineItem):
                p1 = trans.map(item.line().p1())
                p2 = trans.map(item.line().p2())
                msp.add_line((p1.x(), -p1.y()), (p2.x(), -p2.y()), dxfattribs=dxf_attribs)
                
            elif isinstance(item, CadPolylineItem):
                dxf_pts = [(trans.map(p).x(), -trans.map(p).y()) for p in item.points_data]
                msp.add_lwpolyline(dxf_pts, close=item.is_closed, dxfattribs=dxf_attribs)
                
            elif isinstance(item, CadEllipseItem):
                center_scene = trans.map(item.rect().center())
                rx = item.rect().width() / 2 * math.sqrt(trans.m11()**2 + trans.m12()**2)
                msp.add_circle((center_scene.x(), -center_scene.y()), rx, dxfattribs=dxf_attribs)
                
            elif isinstance(item, CadArcItem):
                center_scene = trans.map(item.arc_rect.center())
                rx = item.arc_rect.width() / 2 * math.sqrt(trans.m11()**2 + trans.m12()**2)
                sa = 360 - (item.start_angle + item.span)
                ea = 360 - item.start_angle
                if sa < 0: sa += 360
                if ea < 0: ea += 360
                msp.add_arc((center_scene.x(), -center_scene.y()), rx, sa, ea, dxfattribs=dxf_attribs)
                
            elif isinstance(item, CadTextItem):
                local_pos = QPointF(0, item.text_size)
                scene_pos = trans.map(local_pos)
                height = item.text_size * math.sqrt(trans.m22()**2 + trans.m21()**2)
                if height <= 0: height = 12
                dxf_attribs['height'] = height
                msp.add_text(item.toPlainText(), dxfattribs=dxf_attribs).set_placement((scene_pos.x(), -scene_pos.y()))

    for item in scene.items():
        if item.parentItem() is None:
            process_item(item)

    doc.saveas(path)

def import_dxf_as_items(path):
    """DXFファイルを読み込み、高精度なCADアイテムのリストとして返す"""
    if not ezdxf: return []
    try:
        doc = ezdxf.readfile(path)
    except Exception as e:
        print(f"Failed to read DXF: {e}")
        return []
        
    msp = doc.modelspace()
    items = []

    def get_color(entity):
        true_color = entity.dxf.get('true_color')
        if true_color is not None:
            try:
                rgb = ezdxf.colors.int2rgb(true_color)
                return QColor(rgb[0], rgb[1], rgb[2])
            except: pass
        
        aci_color = entity.dxf.get('color', 256)
        if aci_color == 256: # ByLayer
            layer_obj = doc.layers.get(entity.dxf.layer)
            aci_color = layer_obj.color if layer_obj else 7
        
        if aci_color < 256:
            try:
                r, g, b = ezdxf.colors.aci2rgb(aci_color)
                return QColor(r, g, b)
            except: pass
        return QColor(Qt.GlobalColor.black)

    def get_width(entity):
        # lineweightは1/100mm単位。デフォルトを1とする。
        lw = entity.dxf.get('lineweight', -1)
        if lw <= 0: return 1.0
        return max(0.1, lw / 25.0) # 簡易的なスケーリング

    def process_entities(entities):
        for entity in entities:
            try:
                # INSERT (ブロック参照) の場合は展開して再帰的に処理する
                if entity.dxftype() == 'INSERT':
                    # virtual_entities() はすべての変換（位置、回転、スケール）を適用した状態で子要素を返す
                    process_entities(entity.virtual_entities())
                    continue

                etype = entity.dxftype()
                color = get_color(entity)
                layer = entity.dxf.layer
                width = get_width(entity)

                if etype in ('TEXT', 'MTEXT'):
                    if etype == 'TEXT':
                        txt = entity.dxf.text
                        pos = entity.dxf.insert if not entity.dxf.hasattr('align_point') else entity.dxf.align_point
                    else:
                        # MTEXTの書式コードを除去し、改行を適切に処理
                        txt = re.sub(r'\\[a-zA-Z0-9]+|{[^{}]*}', '', entity.text)
                        pos = entity.dxf.insert
                    
                    height = entity.dxf.get('height', 10)
                    rot = entity.dxf.get('rotation', 0)
                    
                    # アライメント
                    align = Qt.AlignmentFlag.AlignLeft
                    if etype == 'TEXT':
                        h, v = entity.dxf.get('halign', 0), entity.dxf.get('valign', 0)
                        if h == 1: align |= Qt.AlignmentFlag.AlignHCenter
                        elif h == 2: align |= Qt.AlignmentFlag.AlignRight
                        if v == 1: align |= Qt.AlignmentFlag.AlignBottom
                        elif v == 2: align |= Qt.AlignmentFlag.AlignVCenter
                        elif v == 3: align |= Qt.AlignmentFlag.AlignTop
                    else:
                        att = entity.dxf.get('attachment_point', 1)
                        if att in (2, 5, 8): align |= Qt.AlignmentFlag.AlignHCenter
                        elif att in (3, 6, 9): align |= Qt.AlignmentFlag.AlignRight
                        if att in (4, 5, 6): align |= Qt.AlignmentFlag.AlignVCenter
                        elif att in (7, 8, 9): align |= Qt.AlignmentFlag.AlignBottom
                        else: align |= Qt.AlignmentFlag.AlignTop

                    item = CadTextItem(txt, pos.x, -pos.y, height, color, layer, rotation=-rot, alignment=align)
                    items.append(item)
                    continue

                # 図形要素 (LINE, CIRCLE, ARC, SPLINE等)
                if etype == 'LINE':
                    if entity.dxf.start.isclose(entity.dxf.end, abs_tol=1e-5): continue
                
                # make_path は複雑な図形も高精度のベジェ/ラインパスに変換する
                paths = ezdxf.path.make_path(entity)
                for p in paths:
                    qpath = to_qpainterpath(p)
                    if qpath.length() < 0.005: continue
                        
                    if etype == 'HATCH' and entity.dxf.get('solid_fill', 0):
                        items.append(CadHatchItem(qpath, color, layer))
                    else:
                        items.append(CadPathItem(qpath, color, layer, width))
            except Exception as e:
                # 個別のアイテム失敗で全体を止めない
                pass

    process_entities(msp)
    return items
