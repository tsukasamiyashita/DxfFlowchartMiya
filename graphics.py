# tsukasamiyashita/dxfflowchartmiya/DxfFlowchartMiya-dac99957653817cc9b44f6154480ff40d3256b04/graphics.py
import uuid
import math
from PyQt6.QtWidgets import (QGraphicsScene, QGraphicsView, QGraphicsPathItem, 
                             QGraphicsTextItem, QGraphicsItem, QGraphicsEllipseItem, 
                             QGraphicsItemGroup, QDialog, QVBoxLayout, QGraphicsLineItem,
                             QHBoxLayout, QLabel, QTextEdit, QPushButton, QStyle, QStyleOptionGraphicsItem)
from PyQt6.QtCore import Qt, QRectF, QPointF, QLineF
from PyQt6.QtGui import (QPen, QBrush, QColor, QPainter, QPainterPath, 
                         QPainterPathStroker, QTransform, QUndoCommand, QPolygonF)

GRID_SIZE = 20

class TextEditDialog(QDialog):
    def __init__(self, parent, title, label_text, initial_text):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(400, 300)
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(label_text))
        
        self.editor = QTextEdit()
        self.editor.setPlainText(initial_text)
        
        self.editor.setStyleSheet("""
            QTextEdit { 
                border: 2px solid #3b82f6; 
                border-radius: 4px;
                font-family: 'Segoe UI', 'Meiryo', sans-serif; 
                font-size: 18px; 
                padding: 10px;
            }
        """)
        
        layout.addWidget(self.editor)
        
        btns = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)
        self.ok_btn.setStyleSheet("background-color: #3b82f6; color: white; border: none; font-weight: bold;")
        
        self.cancel_btn = QPushButton("キャンセル")
        self.cancel_btn.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(self.ok_btn)
        btns.addWidget(self.cancel_btn)
        layout.addLayout(btns)
        
    def get_text(self):
        return self.editor.toPlainText()

class SceneStateCommand(QUndoCommand):
    def __init__(self, main_window, old_state, new_state, description):
        super().__init__(description)
        self.main_window = main_window
        self.old_state = old_state
        self.new_state = new_state
        self.is_first_redo = True 

    def undo(self):
        self.main_window.load_scene_json(self.old_state, clear_scene=True, generate_new_ids=False, is_undo_redo=True)

    def redo(self):
        if self.is_first_redo:
            self.is_first_redo = False
            return
        self.main_window.load_scene_json(self.new_state, clear_scene=True, generate_new_ids=False, is_undo_redo=True)

class FlowchartView(QGraphicsView):
    def __init__(self, scene):
        super().__init__(scene)
        self.setRenderHints(QPainter.RenderHint.Antialiasing | 
                           QPainter.RenderHint.TextAntialiasing | 
                           QPainter.RenderHint.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.zoom_factor = 1.15
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.setRubberBandSelectionMode(Qt.ItemSelectionMode.IntersectsItemShape)
        self.setMouseTracking(True)
        self.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.FullViewportUpdate) # 描画精度優先

    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0: self.scale(self.zoom_factor, self.zoom_factor)
            else: self.scale(1 / self.zoom_factor, 1 / self.zoom_factor)
        else:
            super().wheelEvent(event)

    def leaveEvent(self, event):
        if hasattr(self.scene(), 'hide_preview_node'):
            self.scene().hide_preview_node()
        super().leaveEvent(event)

class NodeItem(QGraphicsPathItem):
    def __init__(self, x, y, text="Node", node_type="process", node_id=None, bg_color="#E1F5FE", text_color="#000000", w=100, h=50, line_color=None):
        super().__init__()
        self.node_type = node_type
        self.node_id = node_id if node_id else str(uuid.uuid4())
        self.edges = []
        self.bg_color = QColor(bg_color)
        self.text_color = QColor(text_color)
        self.line_color = QColor(line_color) if line_color else None
        self.font_family = "ＭＳ ゴシック"
        self.w = w
        self.h = h
        self._is_highlighted = False
        
        self.setPos(x, y)
        self.setBrush(QBrush(self.bg_color))
        self.update_pen()
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable | QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        
        self.text_item = QGraphicsTextItem(text)
        self.text_item.setParentItem(self)
        self.text_item.setDefaultTextColor(self.text_color)
        
        f = self.text_item.font()
        f.setFamily(self.font_family)
        self.text_item.setFont(f)
        
        self.update_path()
        self.set_text(text)

    def set_font_family(self, family):
        self.font_family = family
        f = self.text_item.font()
        f.setFamily(family)
        self.text_item.setFont(f)
        self._update_text_pos()

    def update_path(self):
        path = QPainterPath()
        hw, hh = self.w / 2, self.h / 2
        if self.node_type == "process":
            path.addRect(QRectF(-hw, -hh, self.w, self.h))
        elif self.node_type == "decision":
            path.moveTo(0, -hh-10); path.lineTo(hw+10, 0); path.lineTo(0, hh+10); path.lineTo(-hw-10, 0); path.closeSubpath()
        elif self.node_type == "data":
            skew = hh
            path.moveTo(-hw+skew/2, -hh); path.lineTo(hw+skew/2, -hh); path.lineTo(hw-skew/2, hh); path.lineTo(-hw-skew/2, hh); path.closeSubpath()
        elif self.node_type == "terminal":
            path.addRoundedRect(QRectF(-hw, -hh, self.w, self.h), hh, hh)
        else:
            path.addRect(QRectF(-hw, -hh, self.w, self.h))
        self.setPath(path)
        self._update_text_pos()
        for edge in self.edges: edge.update_position()

    def set_size(self, w, h):
        if self.w != w or self.h != h:
            self.w, self.h = w, h
            self.update_path()

    def _update_text_pos(self):
        r = self.boundingRect(); tr = self.text_item.boundingRect()
        self.text_item.setPos(r.center().x() - tr.width()/2, r.center().y() - tr.height()/2)

    def set_text(self, text):
        self.text_item.setHtml(f"<div align='center'>{text.replace(chr(10), '<br>')}</div>")
        self._update_text_pos()

    def set_bg_color(self, color: QColor):
        self.bg_color = color; self.setBrush(QBrush(self.bg_color))

    def set_text_color(self, color: QColor):
        self.text_color = color; self.text_item.setDefaultTextColor(self.text_color)

    def set_line_color(self, color: QColor):
        self.line_color = color; self.update_pen()

    def update_pen(self):
        if self.line_color:
            color = self.line_color
        else:
            color = Qt.GlobalColor.black
            if self.scene() and hasattr(self.scene(), 'main_window'):
                if not self.scene().main_window.is_light_theme:
                    color = QColor("#E0E0E0")
        self.default_pen = QPen(color, 2)
        self.update()

    def set_highlight(self, active: bool):
        if self._is_highlighted != active: self._is_highlighted = active; self.update() 

    def add_edge(self, edge):
        self.edges.append(edge)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self.scene():
            if getattr(self.scene(), 'snap_to_grid', True):
                return QPointF(round(value.x()/GRID_SIZE)*GRID_SIZE, round(value.y()/GRID_SIZE)*GRID_SIZE)
            return value
        elif change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            for edge in self.edges: edge.update_position()
        elif change == QGraphicsItem.GraphicsItemChange.ItemSceneHasChanged:
            self.update_pen()
        return super().itemChange(change, value)

    def paint(self, painter, option, widget=None):
        if self._is_highlighted: painter.setPen(QPen(QColor("#FF5722"), 3, Qt.PenStyle.DashLine))
        elif self.isSelected(): painter.setPen(QPen(QColor("#3B82F6"), 3))
        else: painter.setPen(self.default_pen)
        painter.setBrush(self.brush()); painter.drawPath(self.path())

    def mouseDoubleClickEvent(self, event):
        parent_win = self.scene().main_window if self.scene() and hasattr(self.scene(), 'main_window') else None
        dialog = TextEditDialog(parent_win, "テキスト編集", "ノード名:", self.text_item.toPlainText())
        if dialog.exec():
            new_text = dialog.get_text()
            self.set_text(new_text)
            if parent_win: parent_win.push_undo_state("テキスト変更")
        super().mouseDoubleClickEvent(event)

class WaypointItem(QGraphicsEllipseItem):
    def __init__(self, x, y, edge):
        super().__init__(-6, -6, 12, 12)
        self.edge = edge
        self.setPos(x, y)
        self.orig_x = x
        self.orig_y = y
        self.setBrush(QBrush(QColor("#FF9800"))); self.setPen(QPen(Qt.GlobalColor.white, 2))
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.setZValue(1)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self.scene():
            if getattr(self.scene(), 'snap_to_grid', True):
                return QPointF(round(value.x()/GRID_SIZE)*GRID_SIZE, round(value.y()/GRID_SIZE)*GRID_SIZE)
            return value
        elif change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            self.edge.update_position()
        return super().itemChange(change, value)

    def paint(self, painter, option, widget=None):
        painter.setPen(QPen(QColor("#3B82F6"), 2) if self.isSelected() else self.pen())
        painter.setBrush(self.brush()); painter.drawEllipse(self.rect())

    def mouseDoubleClickEvent(self, event):
        self.edge.remove_waypoint(self); super().mouseDoubleClickEvent(event)
        
    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event); self.ungrabMouse(); self.edge.check_waypoint_straightness(self)

class EdgeTextItem(QGraphicsTextItem):
    def __init__(self, text, edge):
        super().__init__(text)
        self.edge = edge
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable | QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        self.setParentItem(edge)
        self.update_style()
        self.manual_offset = None; self._is_dragging = False

    def update_style(self):
        color = QColor("#333333")
        if self.scene() and hasattr(self.scene(), 'main_window'):
            if not self.scene().main_window.is_light_theme:
                color = QColor("#F8F9FA")
        self.setDefaultTextColor(color)

    def mousePressEvent(self, event): self._is_dragging = True; super().mousePressEvent(event)
    def mouseReleaseEvent(self, event): self._is_dragging = False; super().mouseReleaseEvent(event)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self._is_dragging:
            base_pos = self.edge.get_auto_text_pos()
            if base_pos is not None: self.manual_offset = value - base_pos
        elif change == QGraphicsItem.GraphicsItemChange.ItemSceneHasChanged:
            self.update_style()
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        parent_win = self.scene().main_window if self.scene() and hasattr(self.scene(), 'main_window') else None
        dialog = TextEditDialog(parent_win, "エッジのテキスト編集", "線上のテキスト:", self.edge.raw_text)
        if dialog.exec():
            new_text = dialog.get_text()
            self.edge.set_text(new_text)
            if parent_win: parent_win.push_undo_state("エッジテキスト変更")

    def paint(self, painter, option, widget=None):
        option.state &= ~QStyle.StateFlag.State_Selected
        super().paint(painter, option, widget)
        if self.isSelected():
            painter.setPen(QPen(QColor("#3B82F6"), 1, Qt.PenStyle.DashLine)); painter.setBrush(Qt.BrushStyle.NoBrush); painter.drawRect(self.boundingRect())

def clip_line_to_node(p_start: QPointF, p_end: QPointF, node: NodeItem) -> QPointF:
    line = QLineF(p_start, p_end); polygon = node.mapToScene(node.path().toFillPolygon())
    best_p, min_dist = p_start, float('inf')
    for i in range(polygon.count()):
        p_a, p_b = polygon.at(i), polygon.at((i + 1) % polygon.count())
        intersect_type, ip = line.intersects(QLineF(p_a, p_b))
        if intersect_type == QLineF.IntersectionType.BoundedIntersection:
            dist = QLineF(p_start, ip).length()
            if dist < min_dist: min_dist = dist; best_p = ip
    return best_p

class EdgeItem(QGraphicsPathItem):
    def __init__(self, source_node, target_node, label="", width=2, style="solid", routing="straight", arrow="end", line_color=None):
        super().__init__()
        self.source_node = source_node
        self.target_node = target_node
        self.raw_text = label
        self.waypoints = []
        self._drag_start_pos = None; self._potential_waypoint_index = -1
        
        self.line_width = width
        self.line_style = style
        self.routing = routing
        self.arrow = arrow
        self.line_color = QColor(line_color) if line_color else None
        self.font_family = "ＭＳ ゴシック"
        self.update_pen()
        
        self.setZValue(-1); self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.text_item = EdgeTextItem("", self)
        self._set_label_html(label)
        self.update_position()

    def set_font_family(self, family):
        self.font_family = family
        f = self.text_item.font()
        f.setFamily(family)
        self.text_item.setFont(f)
        self.update_position()

    def set_line_color(self, color: QColor):
        self.line_color = color; self.update_pen()

    def update_pen(self):
        style_map = {"solid": Qt.PenStyle.SolidLine, "dash": Qt.PenStyle.DashLine, "dot": Qt.PenStyle.DotLine}
        ps = style_map.get(self.line_style, Qt.PenStyle.SolidLine)
        
        if self.line_color:
            color = self.line_color
        else:
            color = Qt.GlobalColor.black
            if self.scene() and hasattr(self.scene(), 'main_window'):
                if not self.scene().main_window.is_light_theme:
                    color = QColor("#E0E0E0")
        
        self.default_pen = QPen(color, self.line_width, ps)
        self.setPen(self.default_pen)

    def boundingRect(self): return super().boundingRect().adjusted(-10, -10, 10, 10)
    def shape(self): stroker = QPainterPathStroker(); stroker.setWidth(20); return stroker.createStroke(super().shape())

    def _set_label_html(self, text):
        self.raw_text = text
        if text: 
            self.text_item.setHtml(f"<div style='font-weight: bold; font-family: {self.font_family}; text-align: center;'>{text.replace(chr(10), '<br>')}</div>")
            self.text_item.show()
        else: 
            self.text_item.setHtml("")
            self.text_item.hide()

    def set_text(self, text): 
        self._set_label_html(text)
        self.update_position()

    def get_auto_text_pos(self):
        if not self.source_node or not self.target_node: return None
        path = self.path()
        if path.isEmpty(): return QPointF(0, 0)
        c = path.pointAtPercent(0.5)
        r = self.text_item.boundingRect()
        return QPointF(c.x() - r.width()/2, c.y() - r.height()/2 - 15)

    def _get_orthogonal_path(self, p1, p2):
        path = QPainterPath(); path.moveTo(p1)
        mid_y = (p1.y() + p2.y()) / 2
        path.lineTo(p1.x(), mid_y)
        path.lineTo(p2.x(), mid_y)
        path.lineTo(p2)
        return path

    def update_position(self):
        if not self.source_node or not self.target_node: return
        self.prepareGeometryChange()
        
        pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
        path = QPainterPath()
        p_before_end = pts[0]
        
        if self.routing == "orthogonal" and not self.waypoints:
            path = self._get_orthogonal_path(pts[0], pts[-1])
            p_before_end = QPointF(pts[-1].x(), (pts[0].y() + pts[-1].y()) / 2)
        else:
            path.moveTo(pts[0])
            for i in range(1, len(pts)):
                if self.routing == "orthogonal":
                    mid_y = (pts[i-1].y() + pts[i].y()) / 2
                    path.lineTo(pts[i-1].x(), mid_y)
                    path.lineTo(pts[i].x(), mid_y)
                    if i == len(pts) - 1: p_before_end = QPointF(pts[i].x(), mid_y)
                else:
                    if i == len(pts) - 1: p_before_end = pts[i-1]
                path.lineTo(pts[i])

        self.setPath(path)
        
        p_after_start = pts[1]
        if self.routing == "orthogonal":
            mid_y_start = (pts[0].y() + pts[1].y()) / 2
            p_after_start = QPointF(pts[0].x(), mid_y_start)
        
        self.start_arrow_p1 = p_after_start
        self.start_arrow_p2 = clip_line_to_node(p_after_start, pts[0], self.source_node)

        self.arrow_p1 = p_before_end
        self.arrow_p2 = clip_line_to_node(p_before_end, pts[-1], self.target_node)
        
        if self.raw_text:
            base = self.get_auto_text_pos()
            if base is not None: self.text_item.setPos(base + self.text_item.manual_offset if self.text_item.manual_offset else base)

    def paint(self, painter, option, widget=None):
        painter.setPen(QPen(QColor("#3B82F6"), max(3, self.line_width)) if self.isSelected() else self.default_pen)
        painter.drawPath(self.path())
        
        if self.arrow in ["end", "both"] and hasattr(self, 'arrow_p1') and hasattr(self, 'arrow_p2'):
            angle = math.atan2(self.arrow_p2.y() - self.arrow_p1.y(), self.arrow_p2.x() - self.arrow_p1.x())
            arrow_size = 12 + self.line_width * 1.5
            arrow_p1 = self.arrow_p2 - QPointF(math.cos(angle + math.pi / 6) * arrow_size, math.sin(angle + math.pi / 6) * arrow_size)
            arrow_p2 = self.arrow_p2 - QPointF(math.cos(angle - math.pi / 6) * arrow_size, math.sin(angle - math.pi / 6) * arrow_size)
            
            painter.setBrush(QBrush(QColor("#3B82F6") if self.isSelected() else self.default_pen.color()))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawPolygon(QPolygonF([self.arrow_p2, arrow_p1, arrow_p2]))

        if self.arrow in ["start", "both"] and hasattr(self, 'start_arrow_p1') and hasattr(self, 'start_arrow_p2'):
            angle = math.atan2(self.start_arrow_p2.y() - self.start_arrow_p1.y(), self.start_arrow_p2.x() - self.start_arrow_p1.x())
            arrow_size = 12 + self.line_width * 1.5
            arrow_p1 = self.start_arrow_p2 - QPointF(math.cos(angle + math.pi / 6) * arrow_size, math.sin(angle + math.pi / 6) * arrow_size)
            arrow_p2 = self.start_arrow_p2 - QPointF(math.cos(angle - math.pi / 6) * arrow_size, math.sin(angle - math.pi / 6) * arrow_size)
            
            painter.setBrush(QBrush(QColor("#3B82F6") if self.isSelected() else self.default_pen.color()))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawPolygon(QPolygonF([self.start_arrow_p2, arrow_p1, arrow_p2]))
            
        if self.isSelected():
            painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(QBrush(QColor("#3B82F6")))
            pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
            for i in range(len(pts)-1):
                painter.drawEllipse(QPointF((pts[i].x()+pts[i+1].x())/2, (pts[i].y()+pts[i+1].y())/2), 5.0, 5.0)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemSceneHasChanged:
            self.update_pen()
        return super().itemChange(change, value)

    def mousePressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier: super().mousePressEvent(event); return
        if event.button() == Qt.MouseButton.LeftButton and self.routing != "orthogonal":
            pos = event.scenePos()
            pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
            for i in range(len(pts) - 1):
                if math.hypot(pos.x() - (pts[i].x()+pts[i+1].x())/2, pos.y() - (pts[i].y()+pts[i+1].y())/2) < 30:
                    self._drag_start_pos = pos; self._potential_waypoint_index = i; event.accept(); return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_start_pos and (event.scenePos() - self._drag_start_pos).manhattanLength() > 5:
            pos = event.scenePos()
            sx, sy = round(pos.x()/GRID_SIZE)*GRID_SIZE, round(pos.y()/GRID_SIZE)*GRID_SIZE
            wp = WaypointItem(sx, sy, self); self.waypoints.insert(self._potential_waypoint_index, wp)
            self.scene().items_ref.append(wp); self.scene().addItem(wp); self.update_position()
            self._drag_start_pos = None; self._potential_waypoint_index = -1
            wp.grabMouse(); wp.setPos(sx, sy); event.accept(); return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event): self._drag_start_pos = None; self._potential_waypoint_index = -1; super().mouseReleaseEvent(event)
    
    def mouseDoubleClickEvent(self, event):
        parent_win = self.scene().main_window if self.scene() and hasattr(self.scene(), 'main_window') else None
        dialog = TextEditDialog(parent_win, "エッジのテキスト編集", "線上のテキスト:", self.raw_text)
        if dialog.exec():
            new_text = dialog.get_text()
            self.set_text(new_text)
            if parent_win:
                parent_win.push_undo_state("エッジテキスト変更")
        super().mouseDoubleClickEvent(event)

    def remove_waypoint(self, wp):
        if wp in self.waypoints:
            self.waypoints.remove(wp); self.scene().removeItem(wp)
            if wp in self.scene().items_ref: self.scene().items_ref.remove(wp)
            self.update_position(); self.scene().main_window.push_undo_state("ウェイポイント削除")
            
    def check_waypoint_straightness(self, wp):
        if wp not in self.waypoints or self.routing == "orthogonal": return
        idx = self.waypoints.index(wp)
        p1 = self.source_node.scenePos() if idx == 0 else self.waypoints[idx - 1].scenePos()
        p2 = self.target_node.scenePos() if idx == len(self.waypoints) - 1 else self.waypoints[idx + 1].scenePos()
        line = QLineF(p1, p2); length = line.length()
        if length == 0: self.remove_waypoint(wp); return
        dist = abs((p2.x()-p1.x())*(p1.y()-wp.scenePos().y()) - (p1.x()-wp.scenePos().x())*(p2.y()-p1.y())) / length
        dot = (wp.scenePos().x()-p1.x())*(p2.x()-p1.x()) + (wp.scenePos().y()-p1.y())*(p2.y()-p1.y())
        if dist < 15.0 and 0 <= dot <= length ** 2: self.remove_waypoint(wp)
        self.scene().main_window.push_undo_state("線の変形")

# === CAD Items ===

class CadBase:
    def init_cad(self, color, layer, node_id):
        self.is_cad_item = True
        self.cad_color = QColor(color)
        self.cad_layer = layer
        self.node_id = node_id if node_id else str(uuid.uuid4())
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable | QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        self.cad_width = 1

    def set_cad_color(self, color):
        self.cad_color = color
        if hasattr(self, 'setPen') and not isinstance(self, QGraphicsTextItem):
            pen = self.pen()
            pen.setColor(color)
            self.setPen(pen)
        elif isinstance(self, QGraphicsTextItem):
            self.setDefaultTextColor(color)

    def set_cad_width(self, width):
        self.cad_width = width
        if hasattr(self, 'setPen') and not isinstance(self, QGraphicsTextItem):
            pen = self.pen()
            # 線の太さが極端に細い場合に消えないよう、最低幅を保証し、かつ拡大縮小の影響を受けにくいCosmeticを設定
            pen.setWidthF(max(1.0, width))
            pen.setCosmetic(True)
            self.setPen(pen)

    def paint_selection_box(self, painter):
        if self.isSelected():
            rect = self.boundingRect()
            if rect.width() < 1 and rect.height() < 1: return
            lod = QStyleOptionGraphicsItem.levelOfDetailFromTransform(painter.worldTransform())
            pen_w = 1.0 / lod if lod > 0 else 1.0
            painter.setPen(QPen(QColor("#FF5722"), pen_w, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(rect)

class CadLineItem(QGraphicsLineItem, CadBase):
    def __init__(self, x1, y1, x2, y2, color, layer, width=1, node_id=None):
        super().__init__(x1, y1, x2, y2)
        self.init_cad(color, layer, node_id)
        self.set_cad_width(width)
        self.setPen(QPen(self.cad_color, width))

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadPolylineItem(QGraphicsPathItem, CadBase):
    def __init__(self, points, is_closed, color, layer, width=1, node_id=None):
        super().__init__()
        self.init_cad(color, layer, node_id)
        self.points_data = points
        self.is_closed = is_closed
        self.set_cad_width(width)
        path = QPainterPath()
        if points:
            path.moveTo(points[0].x(), points[0].y())
            for p in points[1:]: path.lineTo(p.x(), p.y())
            if is_closed: path.closeSubpath()
        self.setPath(path)
        self.setPen(QPen(self.cad_color, width))

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadEllipseItem(QGraphicsEllipseItem, CadBase):
    def __init__(self, x, y, w, h, color, layer, width=1, node_id=None):
        super().__init__(x, y, w, h)
        self.init_cad(color, layer, node_id)
        self.set_cad_width(width)
        self.setPen(QPen(self.cad_color, width))

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadArcItem(QGraphicsPathItem, CadBase):
    def __init__(self, x, y, w, h, start_angle, span, color, layer, width=1, node_id=None):
        super().__init__()
        self.init_cad(color, layer, node_id)
        self.arc_rect = QRectF(x, y, w, h)
        self.start_angle = start_angle
        self.span = span
        self.set_cad_width(width)
        path = QPainterPath()
        path.arcMoveTo(self.arc_rect, start_angle)
        path.arcTo(self.arc_rect, start_angle, span)
        self.setPath(path)
        self.setPen(QPen(self.cad_color, width))

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadPathItem(QGraphicsPathItem, CadBase):
    def __init__(self, path, color, layer, width=1, node_id=None):
        super().__init__(path)
        self.init_cad(color, layer, node_id)
        self.set_cad_width(width)
        self.setPen(QPen(self.cad_color, width))

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadHatchItem(QGraphicsPathItem, CadBase):
    def __init__(self, path, color, layer, node_id=None):
        super().__init__(path)
        self.init_cad(color, layer, node_id)
        self.setBrush(QBrush(self.cad_color))
        self.setPen(Qt.PenStyle.NoPen)

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

class CadTextItem(QGraphicsTextItem, CadBase):
    def __init__(self, text, x, y, size, color, layer, rotation=0, alignment=Qt.AlignmentFlag.AlignLeft, node_id=None):
        super().__init__()
        self.init_cad(color, layer, node_id)
        self.text_size = size
        self.setDefaultTextColor(self.cad_color)
        f = self.font()
        f.setPointSizeF(size if size > 0 else 10)
        self.setFont(f)
        self._set_text(text)
        self.setPos(x, y)
        self.setRotation(rotation)
        # 簡易的なアライメント対応
        if alignment & Qt.AlignmentFlag.AlignRight:
            self.setPos(x - self.boundingRect().width(), y)
        elif alignment & Qt.AlignmentFlag.AlignHCenter:
            self.setPos(x - self.boundingRect().width() / 2, y)

    def _set_text(self, text):
        self.setHtml(f"<div>{text.replace(chr(10), '<br>')}</div>")

    def paint(self, painter, option, widget=None):
        option.state &= ~QStyle.StateFlag.State_Selected
        super().paint(painter, option, widget)
        self.paint_selection_box(painter)

# === Scene ===

class FlowchartScene(QGraphicsScene):
    def __init__(self, main_window):
        super().__init__(main_window)
        self.main_window = main_window
        self.source_node = None
        self.items_ref = [] 
        self.preview_node = None
        self.preview_items = []
        self.draw_grid = True
        self.snap_to_grid = True

    def hide_preview_node(self):
        try:
            if self.preview_node: self.preview_node.hide()
            for pi in self.preview_items: pi.hide()
        except RuntimeError:
            self.preview_node = None
            self.preview_items = []

    def update_preview_node(self, pos=None, tool=None):
        if tool is None: tool = self.main_window.current_tool

        if tool not in ["process", "decision", "data", "terminal"]:
            try:
                if self.preview_node:
                    self.removeItem(self.preview_node)
                    self.preview_node = None
            except RuntimeError: self.preview_node = None

        if tool != "paste":
            try:
                if self.preview_items:
                    for pi in self.preview_items: 
                        if pi.scene() == self: self.removeItem(pi)
                    self.preview_items = []
            except RuntimeError: self.preview_items = []

        if tool in ["process", "decision", "data", "terminal"]:
            try:
                if self.preview_node and self.preview_node.node_type != tool:
                    self.removeItem(self.preview_node); self.preview_node = None
            except RuntimeError: self.preview_node = None

            if not self.preview_node:
                self.preview_node = NodeItem(0, 0, text="Node", node_type=tool)
                self.preview_node.setOpacity(0.5); self.preview_node.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
                self.preview_node.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
                self.preview_node.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False); self.preview_node.setZValue(1000)
                self.addItem(self.preview_node)
            try:
                self.preview_node.show()
                if pos is not None:
                    sx, sy = round(pos.x() / GRID_SIZE) * GRID_SIZE, round(pos.y() / GRID_SIZE) * GRID_SIZE
                    self.preview_node.setPos(sx, sy)
            except RuntimeError: self.preview_node = None

        elif tool == "paste" and self.main_window.clipboard_data:
            try:
                if not self.preview_items:
                    id_map = {}
                    
                    for c in self.main_window.clipboard_data.get("cad_items", []):
                        ctype = c.get("type")
                        color, layer = c.get("color", "#000000"), c.get("layer", "0")
                        item = None
                        if ctype == "line": item = CadLineItem(c["x1"], c["y1"], c["x2"], c["y2"], color, layer)
                        elif ctype == "polyline":
                            pts = [QPointF(p["x"], p["y"]) for p in c["points"]]
                            item = CadPolylineItem(pts, c["closed"], color, layer)
                        elif ctype == "ellipse": item = CadEllipseItem(c["rect_x"], c["rect_y"], c["w"], c["h"], color, layer)
                        elif ctype == "arc": item = CadArcItem(c["rect_x"], c["rect_y"], c["w"], c["h"], c["start_angle"], c["span"], color, layer)
                        elif ctype == "text": item = CadTextItem(c["text"], 0, 0, c["size"], color, layer)
                        
                        if item:
                            if "transform" in c:
                                t = c["transform"]
                                if len(t) == 9: item.setTransform(QTransform(t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7], t[8]))
                            item.setOpacity(0.5); item.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
                            item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False); item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False); item.setZValue(1000)
                            item.orig_x, item.orig_y = c.get("x", 0), c.get("y", 0)
                            self.addItem(item); self.preview_items.append(item)
                    
                    for n in self.main_window.clipboard_data.get("nodes", []):
                        node = NodeItem(n["x"], n["y"], n["text"], n["type"], str(uuid.uuid4()), n.get("bg_color", "#E1F5FE"), n.get("text_color", "#000000"), line_color=n.get("line_color"))
                        node.setOpacity(0.5); node.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
                        node.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
                        node.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False); node.setZValue(1000)
                        node.orig_x, node.orig_y = n["x"], n["y"]
                        self.addItem(node); self.preview_items.append(node); id_map[n["id"]] = node

                    for e in self.main_window.clipboard_data.get("edges", []):
                        src, tgt = id_map.get(e["source"]), id_map.get(e["target"])
                        if src and tgt:
                            edge = EdgeItem(src, tgt, e.get("label", ""), e.get("width", 2), e.get("style", "solid"), e.get("routing", "straight"), e.get("arrow", "end"), line_color=e.get("line_color"))
                            edge.setOpacity(0.5); edge.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
                            edge.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False); edge.setZValue(1000)
                            if e.get("text_offset"): edge.text_item.manual_offset = QPointF(e.get("text_offset")["x"], e.get("text_offset")["y"])
                            for w in e.get("waypoints", []):
                                wp = WaypointItem(w["x"], w["y"], edge)
                                wp.setOpacity(0.5); wp.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
                                wp.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False); wp.setZValue(1000)
                                wp.orig_x, wp.orig_y = w["x"], w["y"]
                                edge.waypoints.append(wp); self.addItem(wp); self.preview_items.append(wp)
                            src.add_edge(edge); tgt.add_edge(edge)
                            self.addItem(edge); self.preview_items.append(edge); edge.update_position()
                
                if pos is not None and self.main_window.clipboard_base_pos:
                    sx, sy = round(pos.x() / GRID_SIZE) * GRID_SIZE, round(pos.y() / GRID_SIZE) * GRID_SIZE
                    dx, dy = sx - self.main_window.clipboard_base_pos.x(), sy - self.main_window.clipboard_base_pos.y()
                    
                    for item in self.preview_items:
                        item.show()
                        if isinstance(item, (NodeItem, WaypointItem)) or (hasattr(item, 'is_cad_item') and item.is_cad_item): 
                            if hasattr(item, 'orig_x') and hasattr(item, 'orig_y'):
                                item.setPos(item.orig_x + dx, item.orig_y + dy)
                            
                    for item in self.preview_items:
                        if isinstance(item, EdgeItem):
                            item.update_position()
                            
            except RuntimeError: self.preview_items = []

    def drawBackground(self, painter, rect):
        super().drawBackground(painter, rect)
        painter.setPen(QPen(QColor(230, 230, 230) if self.main_window.is_light_theme else QColor(50, 50, 50), 1, Qt.PenStyle.SolidLine))
        left, top = int(rect.left()) - (int(rect.left()) % GRID_SIZE), int(rect.top()) - (int(rect.top()) % GRID_SIZE)
        
        # 補助線（薄いグリッド）
        lines = [QLineF(x, rect.top(), x, rect.bottom()) for x in range(left, int(rect.right()), GRID_SIZE)]
        lines.extend([QLineF(rect.left(), y, rect.right(), y) for y in range(top, int(rect.bottom()), GRID_SIZE)])
        painter.drawLines(lines)
        
        # 主線 (100pxおきに少し強調)
        painter.setPen(QPen(QColor(210, 210, 210) if self.main_window.is_light_theme else QColor(70, 70, 70), 1))
        major_lines = [QLineF(x, rect.top(), x, rect.bottom()) for x in range(left, int(rect.right()), GRID_SIZE * 5) if x % (GRID_SIZE * 5) == 0]
        major_lines.extend([QLineF(rect.left(), y, rect.right(), y) for y in range(top, int(rect.bottom()), GRID_SIZE * 5) if y % (GRID_SIZE * 5) == 0])
        painter.drawLines(major_lines)

    def mousePressEvent(self, event):
        self.main_window.is_moving = False
        tool = self.main_window.current_tool
        
        if tool == "select" and event.button() == Qt.MouseButton.LeftButton:
            if not (event.modifiers() & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier)):
                item = self.itemAt(event.scenePos(), QTransform())
                if item:
                    base_item = item
                    while base_item.parentItem(): 
                        base_item = base_item.parentItem()
                    self.clearSelection()
                    base_item.setSelected(True)
                else:
                    self.clearSelection()
        
        if tool in ["process", "decision", "data", "terminal"] and event.button() == Qt.MouseButton.LeftButton:
            pos = event.scenePos()
            if self.snap_to_grid:
                pos = QPointF(round(pos.x()/GRID_SIZE)*GRID_SIZE, round(pos.y()/GRID_SIZE)*GRID_SIZE)
            self.clearSelection()
            node = NodeItem(pos.x(), pos.y(), text="Node", node_type=tool)
            if hasattr(self.main_window, 'cb_font'):
                node.set_font_family(self.main_window.cb_font.currentText())
            self.items_ref.append(node); self.addItem(node); self.main_window.push_undo_state(f"ノード追加 ({tool})")
            return

        if tool == "connect" and event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.scenePos(), QTransform())
            while item and not isinstance(item, NodeItem): item = item.parentItem()
            if isinstance(item, NodeItem):
                if self.source_node is None:
                    self.source_node = item; self.source_node.set_highlight(True)
                    self.main_window.statusBar().showMessage("エッジ接続モード: 2つ目のノードをクリック")
                elif item != self.source_node:
                    edge = EdgeItem(self.source_node, item, routing=self.main_window.cb_routing.currentData(), arrow=self.main_window.cb_arrow.currentData())
                    if hasattr(self.main_window, 'cb_font'):
                        edge.set_font_family(self.main_window.cb_font.currentText())
                    self.source_node.add_edge(edge); item.add_edge(edge)
                    self.items_ref.append(edge); self.addItem(edge)
                    self.source_node.set_highlight(False); self.source_node = None
                    self.main_window.push_undo_state("エッジ接続")
                    self.main_window.statusBar().showMessage("エッジ接続モード: 次の1つ目のノードをクリック")
                return
            else:
                if self.source_node: self.source_node.set_highlight(False); self.source_node = None
                self.clearSelection()
                return

        if tool == "paste" and event.button() == Qt.MouseButton.LeftButton:
            if self.main_window.clipboard_data and self.main_window.clipboard_base_pos:
                sx, sy = round(event.scenePos().x() / GRID_SIZE) * GRID_SIZE, round(event.scenePos().y() / GRID_SIZE) * GRID_SIZE
                dx, dy = sx - self.main_window.clipboard_base_pos.x(), sy - self.main_window.clipboard_base_pos.y()
                self.main_window.scene.clearSelection()
                try:
                    for pi in self.preview_items: 
                        if pi.scene() == self: self.removeItem(pi)
                except RuntimeError: pass
                self.preview_items = []
                self.main_window.load_scene_json(self.main_window.clipboard_data, offset_x=dx, offset_y=dy, clear_scene=False, generate_new_ids=True)
                self.main_window.push_undo_state("貼り付け")
                self.main_window.set_tool("select")
                self.main_window.btn_select.setChecked(True)
            return

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.selectedItems(): self.main_window.is_moving = True
        tool = self.main_window.current_tool
        if tool in ["process", "decision", "data", "terminal", "paste"]: self.update_preview_node(event.scenePos(), tool)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        if getattr(self.main_window, 'is_moving', False):
            self.main_window.push_undo_state("移動"); self.main_window.is_moving = False

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace): self.main_window.delete_selected_items()
        super().keyPressEvent(event)