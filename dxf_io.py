# tsukasamiyashita/dxfflowchartmiya/DxfFlowchartMiya-1559b301efcf479ed8934d5e23d0d8559520cac5/dxf_io.py
import ezdxf
import math
from PyQt6.QtWidgets import QGraphicsLineItem, QGraphicsPathItem, QGraphicsEllipseItem
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QPainterPath, QPen, QColor
from graphics import NodeItem, EdgeItem, clip_line_to_node, DxfFrameItem
import unicodedata

def export_dxf_file(scene, path, version_string, template_path=None):
    """シーン内のフローチャート要素をDXFファイルとしてエクスポートする"""
    doc = None
    offset_x, offset_y = 0, 0
    scale = 1.0
    
    # 図枠アイテムを探す
    frame_item = None
    for item in scene.items():
        if isinstance(item, DxfFrameItem):
            frame_item = item
            break
            
    if template_path and ezdxf:
        try:
            doc = ezdxf.readfile(template_path)
            if frame_item:
                # キャンバス上での図枠の移動量と倍率
                offset_x = frame_item.scenePos().x()
                offset_y = -frame_item.scenePos().y()
                scale = frame_item.scale()
                if scale <= 0: scale = 1.0
        except Exception as e:
            print(f"Template load failed: {e}")
            doc = None
            
    if not doc:
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
        # プレビュー図形は出力対象外
        if is_preview: 
            continue
        
        if isinstance(item, NodeItem):
            # 図枠の移動分を差し引いて倍率で補正し、元のDXF座標系に合わせる
            pos = item.scenePos()
            raw_x, raw_y = (pos.x() - offset_x) / scale, (-pos.y() - offset_y) / scale
            
            x, y = raw_x, raw_y
            t = item.node_type
            # 図形自体のサイズもDXF座標系（実寸）に合わせるために倍率で割る
            hw, hh = (item.w / 2) / scale, (item.h / 2) / scale
            if t == "process": 
                coords = [(x-hw, y+hh), (x+hw, y+hh), (x+hw, y-hh), (x-hw, y-hh)]
            elif t == "decision": 
                dw, dh = hw + (10 / scale), hh + (10 / scale)
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
                # テキストの高さ・位置も倍率で割る
                base_h = 12 / scale
                line_spacing = 15 / scale
                sy = y + (len(ls)-1)*(base_h * 0.625) 
                for i, l in enumerate(ls): 
                    msp.add_text(l, dxfattribs={'height': base_h, 'layer': 'FC_TEXT'}).set_placement((x, sy-i*line_spacing), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
        elif isinstance(item, EdgeItem):
            # 座標計算時に図枠のオフセットと倍率を考慮
            src_pos = item.source_node.scenePos()
            tgt_pos = item.target_node.scenePos()
            
            # ウェイポイントがある場合の最初と最後のセグメント
            p0 = src_pos
            p1 = item.waypoints[0].scenePos() if item.waypoints else tgt_pos
            p_last = item.waypoints[-1].scenePos() if item.waypoints else src_pos
            pn = tgt_pos
            
            cp0 = clip_line_to_node(p0, p1, item.source_node)
            cpn = clip_line_to_node(pn, p_last, item.target_node)
            
            # DXF座標系に変換
            pts = [QPointF((cp0.x() - offset_x) / scale, (-cp0.y() - offset_y) / scale)]
            for wp in item.waypoints:
                wp_pos = wp.scenePos()
                pts.append(QPointF((wp_pos.x() - offset_x) / scale, (-wp_pos.y() - offset_y) / scale))
            pts.append(QPointF((cpn.x() - offset_x) / scale, (-cpn.y() - offset_y) / scale))
                
            for i in range(len(pts)-1): 
                msp.add_line((pts[i].x(), pts[i].y()), (pts[i+1].x(), pts[i+1].y()), dxfattribs={'layer': 'FC_EDGE'})
                
            if item.raw_text:
                pos = item.get_auto_text_pos() + (item.text_item.manual_offset if item.text_item.manual_offset else QPointF(0,0))
                # テキスト位置とサイズを倍率で補正
                mx = (pos.x() - offset_x + item.text_item.boundingRect().width()/2) / scale
                my = (-(pos.y() + item.text_item.boundingRect().height()/2) - offset_y) / scale
                
                ls = item.raw_text.split('\n')
                base_h = 10 / scale
                line_spacing = 12 / scale
                sy = my + (len(ls)-1)*(base_h * 0.6)
                for i, l in enumerate(ls): 
                    msp.add_text(l, dxfattribs={'height': base_h, 'layer': 'FC_TEXT'}).set_placement((mx, sy-i*line_spacing), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
                    
            if item.arrow in ["end", "both"] and hasattr(item, 'arrow_p1') and hasattr(item, 'arrow_p2'):
                ap2_raw = QPointF((item.arrow_p2.x() - offset_x) / scale, (-item.arrow_p2.y() - offset_y) / scale)
                angle = math.atan2(item.arrow_p2.y() - item.arrow_p1.y(), item.arrow_p2.x() - item.arrow_p1.x())
                # 矢印サイズも倍率で割る
                arrow_size = (12 + item.line_width * 1.5) / scale
                ap1 = ap2_raw - QPointF(math.cos(angle + math.pi / 6) * arrow_size, -math.sin(angle + math.pi / 6) * arrow_size)
                ap2 = ap2_raw - QPointF(math.cos(angle - math.pi / 6) * arrow_size, -math.sin(angle - math.pi / 6) * arrow_size)
                msp.add_lwpolyline([(ap2_raw.x(), ap2_raw.y()), (ap1.x(), ap1.y()), (ap2.x(), ap2.y())], close=True, dxfattribs={'layer': 'FC_EDGE'})
            
            if item.arrow in ["start", "both"] and hasattr(item, 'start_arrow_p1') and hasattr(item, 'start_arrow_p2'):
                ap2_raw = QPointF((item.start_arrow_p2.x() - offset_x) / scale, (-item.start_arrow_p2.y() - offset_y) / scale)
                angle = math.atan2(item.start_arrow_p2.y() - item.start_arrow_p1.y(), item.start_arrow_p2.x() - item.start_arrow_p1.x())
                arrow_size = (12 + item.line_width * 1.5) / scale
                ap1 = ap2_raw - QPointF(math.cos(angle + math.pi / 6) * arrow_size, -math.sin(angle + math.pi / 6) * arrow_size)
                ap2 = ap2_raw - QPointF(math.cos(angle - math.pi / 6) * arrow_size, -math.sin(angle - math.pi / 6) * arrow_size)
                msp.add_lwpolyline([(ap2_raw.x(), ap2_raw.y()), (ap1.x(), ap1.y()), (ap2.x(), ap2.y())], close=True, dxfattribs={'layer': 'FC_EDGE'})
                
    doc.saveas(path)

def import_dxf_as_items(path):
    """DXFファイルを読み込み、QGraphicsItem のリストとして返す"""
    if not ezdxf: return []
    try:
        doc = ezdxf.readfile(path)
    except Exception as e:
        print(f"Failed to read DXF: {e}")
        return []
        
    msp = doc.modelspace()
    from PyQt6.QtWidgets import QGraphicsLineItem, QGraphicsPathItem, QGraphicsEllipseItem, QGraphicsTextItem
    from PyQt6.QtGui import QPen, QColor, QPainterPath
    from PyQt6.QtCore import Qt, QRectF
    
    items = []
    
    for entity in msp:
        # カラー取得 (ACI -> RGB)
        color = QColor(Qt.GlobalColor.gray) # デフォルト
        if entity.dxf.color < 256:
            try:
                # ezdxf での ACI から RGB への標準的な変換
                r, g, b = ezdxf.colors.aci2rgb(entity.dxf.color)
                color = QColor(r, g, b)
            except Exception:
                pass
        
        pen = QPen(color, 1)
        
        if entity.dxftype() == 'LINE':
            start, end = entity.dxf.start, entity.dxf.end
            line = QGraphicsLineItem(start.x, -start.y, end.x, -end.y)
            line.setPen(pen)
            items.append(line)
            
        elif entity.dxftype() == 'LWPOLYLINE':
            path = QPainterPath()
            pts = entity.get_points()
            if pts:
                path.moveTo(pts[0][0], -pts[0][1])
                for i in range(1, len(pts)):
                    path.lineTo(pts[i][0], -pts[i][1])
                if entity.closed:
                    path.closeSubpath()
            p_item = QGraphicsPathItem(path)
            p_item.setPen(pen)
            items.append(p_item)
            
        elif entity.dxftype() == 'CIRCLE':
            center = entity.dxf.center
            radius = entity.dxf.radius
            ellipse = QGraphicsEllipseItem(center.x - radius, -center.y - radius, radius * 2, radius * 2)
            ellipse.setPen(pen)
            items.append(ellipse)
            
        elif entity.dxftype() == 'ARC':
            center = entity.dxf.center
            radius = entity.dxf.radius
            start_angle = entity.dxf.start_angle
            end_angle = entity.dxf.end_angle
            path = QPainterPath()
            span = end_angle - start_angle
            if span < 0: span += 360
            # Qtの角度は反時計回り、0度は右方向、DXFも同様だがY反転に注意
            # ezdxf の ARC は反時計回り。
            # QPainterPath.arcTo(rect, startAngle, sweepLength)
            path.arcMoveTo(center.x - radius, -center.y - radius, radius * 2, radius * 2, start_angle)
            path.arcTo(center.x - radius, -center.y - radius, radius * 2, radius * 2, start_angle, span)
            a_item = QGraphicsPathItem(path)
            a_item.setPen(pen)
            items.append(a_item)
            
        elif entity.dxftype() == 'TEXT':
            txt = entity.dxf.text
            pos = entity.dxf.insert
            t_item = QGraphicsTextItem(txt)
            t_item.setDefaultTextColor(color)
            f = t_item.font()
            f.setPointSizeF(entity.dxf.height * 0.75) # 簡易的なスケール調整
            t_item.setFont(f)
            t_item.setPos(pos.x, -pos.y)
            items.append(t_item)
            
    return items
