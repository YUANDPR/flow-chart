import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QMenu, QAction, QVBoxLayout, QWidget, QPushButton, QFileDialog, QGraphicsTextItem,
    QToolBar, QGraphicsPathItem, QGraphicsLineItem, QMessageBox, QInputDialog
)
from PyQt5.QtCore import Qt, QPointF, QLineF, QRectF
from PyQt5.QtGui import QBrush, QPen, QColor, QPainter, QPainterPath, QCursor
from enum import Enum
import math

# ====================== 枚举定义 ======================
class LineType(Enum):
    """定义连线类型"""
    SINGLE = "single"
    DOUBLE = "double"
    TRIPLE = "triple"  # 修正枚举名称拼写
    DASHED = "dashed"

    def get_pen(self, color=Qt.black, width=2):
        """获取对应线型的画笔"""
        pen = QPen(color, width)
        if self == LineType.DASHED:
            pen.setStyle(Qt.DashLine)
        elif self == LineType.DOUBLE:
            pen.setWidth(4)
        elif self == LineType.TRIPLE:
            pen.setWidth(6)
        return pen


# ====================== 可拖拽方块类 ======================
class DraggableBlock(QGraphicsRectItem):
    """可拖拽的方块组件"""
    _next_id = 1

    def __init__(self, name, x, y, width=100, height=60, block_id=None):
        super().__init__(0, 0, width, height)
        self.name = name
        self.id = block_id if block_id is not None else DraggableBlock._next_id
        if block_id is None:
            DraggableBlock._next_id += 1
        else:
            DraggableBlock._next_id = max(DraggableBlock._next_id, block_id + 1)
        self.setPos(x, y)
        self.connections = []
        self._init_ui()

    def _init_ui(self):
        self.setBrush(QBrush(QColor(173, 216, 230)))
        self.setPen(QPen(Qt.darkBlue, 2))
        self.setFlags(QGraphicsRectItem.ItemIsMovable |
                      QGraphicsRectItem.ItemSendsGeometryChanges |
                      QGraphicsRectItem.ItemIsSelectable)
        self.text = QGraphicsTextItem(f"ID: {self.id}\n{self.name}", self)
        self.text.setPos(5, 5)

    def get_center(self):
        return self.mapToScene(self.rect().center())

    def itemChange(self, change, value):
        if change == QGraphicsRectItem.ItemPositionHasChanged:
            for conn in self.connections:
                conn.update_line()
        return super().itemChange(change, value)

    def contextMenuEvent(self, event):
        menu = QMenu()
        edit_action = menu.addAction("编辑")
        action = menu.exec_(event.screenPos())
        if action == edit_action:
            new_name, ok = QInputDialog.getText(
                None, "编辑方块", "输入名称:", text=self.name)
            if ok:
                self.name = new_name
                self.text.setPlainText(f"ID: {self.id}\n{self.name}")


# ====================== 连接线类 ======================
class Connection(QGraphicsLineItem):
    def __init__(self, start_block, end_block, line_type):
        super().__init__()
        self.setFlag(QGraphicsLineItem.ItemIsSelectable)  # 启用选中
        self.start_block = start_block
        self.end_block = end_block
        self.line_type = line_type
        start_block.connections.append(self)
        end_block.connections.append(self)
        self.update_line()
        self.setPen(line_type.get_pen())
        self.setZValue(-1)

    def update_line(self):
        start_point = self.start_block.get_center()
        end_point = self.end_block.get_center()
        self.setLine(QLineF(start_point, end_point))

    def update_style(self, line_type):
        self.line_type = line_type
        self.setPen(line_type.get_pen())

    def contextMenuEvent(self, event):
        menu = QMenu()
        delete_action = menu.addAction("删除")
        action = menu.exec_(event.screenPos())
        if action == delete_action:
            self.start_block.connections.remove(self)
            self.end_block.connections.remove(self)
            self.scene().removeItem(self)


# ====================== 画布类 ======================
class Canvas(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.scene = QGraphicsScene()
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setDragMode(QGraphicsView.RubberBandDrag)
        self.blocks = []
        self.dragging_block = None
        self.preview_line = None
        self.current_line_type = LineType.SINGLE
        self._draw_grid()

    def _draw_grid(self):
        grid_pen = QPen(QColor(220, 220, 220), 1, Qt.DotLine)
        for x in range(-1000, 1000, 20):
            self.scene.addLine(x, -1000, x, 1000, grid_pen)
        for y in range(-1000, 1000, 20):
            self.scene.addLine(-1000, y, 1000, y, grid_pen)

    def contextMenuEvent(self, event):
        menu = QMenu()
        new_block_action = menu.addAction("新建方块")
        action = menu.exec_(self.mapToGlobal(event.pos()))
        if action == new_block_action:
            pos = self.mapToScene(event.pos())
            block = DraggableBlock("New Block", pos.x(), pos.y())
            self.scene.addItem(block)
            self.blocks.append(block)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            item = self.itemAt(event.pos())
            if isinstance(item, DraggableBlock):
                self._start_connection(item)
                return
        super().mousePressEvent(event)

    def _start_connection(self, block):
        self.dragging_block = block
        start_point = block.get_center()
        self.preview_line = QGraphicsLineItem()
        self.preview_line.setPen(QPen(Qt.gray, 2, Qt.DashLine))
        self.scene.addItem(self.preview_line)
        self.preview_line.setLine(QLineF(start_point, start_point))

    def mouseMoveEvent(self, event):
        if self.dragging_block:
            end_pos = self.mapToScene(event.pos())
            start_point = self.dragging_block.get_center()
            self.preview_line.setLine(QLineF(start_point, end_pos))
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.dragging_block:
            end_pos = self.mapToScene(event.pos())
            end_block = None
            for item in self.scene.items():
                if isinstance(item, DraggableBlock) and item != self.dragging_block:
                    if item.contains(item.mapFromScene(end_pos)):
                        end_block = item
                        break
            if end_block:
                connection = Connection(
                    self.dragging_block,
                    end_block,
                    self.current_line_type
                )
                self.scene.addItem(connection)
            self.scene.removeItem(self.preview_line)
            self.dragging_block = None
            self.preview_line = None
        else:
            super().mouseReleaseEvent(event)


# ====================== 主窗口类 ======================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.canvas = Canvas()
        self.selected_connection = None
        self._init_ui()

    def _init_ui(self):
        self.setWindowTitle("智能框图工具")
        self.setGeometry(100, 100, 1200, 800)
        toolbar = self.addToolBar("工具")
        self._create_line_type_buttons(toolbar)
        control_panel = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self._create_button("导出Excel", self._export))
        layout.addWidget(self._create_button("导入Excel", self._import))
        control_panel.setLayout(layout)
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(control_panel)
        main_layout.addWidget(self.canvas)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def _create_line_type_buttons(self, toolbar):
        self.line_actions = {}
        for line_type in LineType:
            action = QAction(line_type.value.capitalize(), self)
            action.triggered.connect(
                lambda _, lt=line_type: self._change_line_type(lt))
            toolbar.addAction(action)
            self.line_actions[line_type] = action
        self.line_actions[LineType.SINGLE].setChecked(True)

    def _change_line_type(self, line_type):
        self.canvas.current_line_type = line_type
        for item in self.canvas.scene.selectedItems():
            if isinstance(item, Connection):
                item.update_style(line_type)
        for action in self.line_actions.values():
            action.setChecked(False)
        self.line_actions[line_type].setChecked(True)

    def _create_button(self, text, callback):
        btn = QPushButton(text)
        btn.clicked.connect(callback)
        return btn

    def _export(self):
        try:
            modules = []
            for block in self.canvas.blocks:
                modules.append({
                    "ID": block.id,
                    "Name": block.name,
                    "X": block.x(),
                    "Y": block.y(),
                    "Width": block.rect().width(),
                    "Height": block.rect().height()
                })
            connections = []
            for item in self.canvas.scene.items():
                if isinstance(item, Connection):
                    connections.append({
                        "StartID": item.start_block.id,
                        "EndID": item.end_block.id,
                        "Type": item.line_type.value
                    })
            with pd.ExcelWriter("diagram.xlsx") as writer:
                pd.DataFrame(modules).to_excel(writer, sheet_name="Modules", index=False)
                pd.DataFrame(connections).to_excel(writer, sheet_name="Connections", index=False)
            QMessageBox.information(self, "导出成功", "数据已保存到diagram.xlsx")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def _import(self):
        try:
            path, _ = QFileDialog.getOpenFileName(
                self, "打开文件", "", "Excel文件 (*.xlsx)")
            if not path:
                return
            modules_df = pd.read_excel(path, sheet_name="Modules")
            self.canvas.scene.clear()
            self.canvas.blocks = []
            id_map = {}
            for _, row in modules_df.iterrows():
                block = DraggableBlock(
                    name=row["Name"],
                    x=row["X"],
                    y=row["Y"],
                    width=row.get("Width", 100),
                    height=row.get("Height", 60),
                    block_id=row["ID"]
                )
                self.canvas.scene.addItem(block)
                self.canvas.blocks.append(block)
                id_map[row["ID"]] = block
            connections_df = pd.read_excel(path, sheet_name="Connections")
            for _, row in connections_df.iterrows():
                start_block = id_map[row["StartID"]]
                end_block = id_map[row["EndID"]]
                connection = Connection(
                    start_block,
                    end_block,
                    LineType(row["Type"])
                )
                self.canvas.scene.addItem(connection)
            QMessageBox.information(self, "导入成功",
                                    f"已导入 {len(modules_df)} 个模块和 {len(connections_df)} 条连线")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入失败: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())