import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QMenu, QAction, QVBoxLayout, QWidget, QPushButton, QFileDialog, QGraphicsTextItem,
    QToolBar, QMessageBox, QInputDialog, QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QGraphicsPathItem,
    QGraphicsLineItem
)
from PyQt5.QtCore import Qt, QPointF, QLineF, QRectF, QSizeF
from PyQt5.QtGui import QBrush, QPen, QColor, QPainter, QTransform, QCursor, QPainterPath
from enum import Enum


# ====================== 枚举定义 ======================
class LineType(Enum):
    SINGLE = "虚线"  # 黑色虚线
    DOUBLE = "双实线"  # 绿色双实线
    TRIPLE = "三实线"  # 黄色三实线
    QUADRUPLE = "四实线"  # 红色四实线

    def get_offset(self):
        return {
            LineType.SINGLE: [0],
            LineType.DOUBLE: [-4, 4],
            LineType.TRIPLE: [-6, 0, 6],
            LineType.QUADRUPLE: [-6, -2, 2, 6]
        }[self]

    def get_color(self):
        return {
            LineType.SINGLE: Qt.black,
            LineType.DOUBLE: Qt.green,
            LineType.TRIPLE: QColor(255, 204, 0),
            LineType.QUADRUPLE: Qt.red
        }[self]

    def get_style(self):
        return Qt.DashLine if self == LineType.SINGLE else Qt.SolidLine


# ====================== 方块编辑对话框 ======================
class BlockEditDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("方块属性")
        layout = QFormLayout()

        self.name_edit = QLineEdit()
        self.id_edit = QLineEdit()
        self.width_edit = QLineEdit()
        self.height_edit = QLineEdit()

        layout.addRow("名称:", self.name_edit)
        layout.addRow("ID:", self.id_edit)
        layout.addRow("宽度:", self.width_edit)
        layout.addRow("高度:", self.height_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        self.setLayout(layout)


# ====================== 可拖拽方块类 ======================
class DraggableBlock(QGraphicsRectItem):
    _next_id = 1
    resize_handle_size = 8

    def __init__(self, name, x, y, width=100, height=60, block_id=None):
        super().__init__(0, 0, width, height)
        self.name = name
        self.id = block_id if block_id is not None else DraggableBlock._next_id
        self.setPos(x, y)
        self.connections = []
        self._init_ui()
        self.setAcceptHoverEvents(True)
        self.resizing = False

        if block_id is None:
            DraggableBlock._next_id += 1
        else:
            DraggableBlock._next_id = max(DraggableBlock._next_id, block_id + 1)

    def _init_ui(self):
        self.setBrush(QBrush(QColor(173, 216, 230)))
        self.setPen(QPen(Qt.darkBlue, 2))
        self.setFlags(QGraphicsRectItem.ItemIsMovable |
                      QGraphicsRectItem.ItemSendsGeometryChanges |
                      QGraphicsRectItem.ItemIsSelectable)
        self._update_text()

    def _update_text(self):
        if hasattr(self, 'text'):
            self.scene().removeItem(self.text)
        self.text = QGraphicsTextItem(f"ID: {self.id}\n{self.name}", self)
        self.text.setPos(5, 5)

    def get_center(self):
        return self.mapToScene(self.rect().center())

    def itemChange(self, change, value):
        # 位置或尺寸变化时更新所有连接线
        if change in [QGraphicsRectItem.ItemPositionHasChanged,
                      QGraphicsRectItem.ItemTransformHasChanged]:
            for conn in self.connections:
                conn.update_line()
        return super().itemChange(change, value)

    def contextMenuEvent(self, event):
        if not self.isSelected():
            self.setSelected(True)
        menu = QMenu()
        edit_action = menu.addAction("编辑方块")
        action = menu.exec_(event.screenPos())

        if action == edit_action:
            self.edit_properties()

    def edit_properties(self):
        dialog = BlockEditDialog()
        dialog.name_edit.setText(self.name)
        dialog.id_edit.setText(str(self.id))
        dialog.width_edit.setText(str(self.rect().width()))
        dialog.height_edit.setText(str(self.rect().height()))

        if dialog.exec_() == QDialog.Accepted:
            # 更新方块属性
            self.name = dialog.name_edit.text()
            new_id = int(dialog.id_edit.text())
            new_width = float(dialog.width_edit.text())
            new_height = float(dialog.height_edit.text())

            # 检查ID冲突
            if new_id != self.id:
                existing_ids = {block.id for block in self.scene().items()
                                if isinstance(block, DraggableBlock)}
                if new_id in existing_ids:
                    QMessageBox.warning(None, "错误", "ID已存在！")
                    return

            self.id = new_id
            self.setRect(0, 0, new_width, new_height)
            self._update_text()

            # 更新所有相关连接线
            for conn in self.connections:
                conn.update_line()

    def delete_block(self):
        # 删除所有关联的连接线
        for conn in self.connections.copy():
            conn.delete_connection()
        self.scene().removeItem(self)


# ====================== 连接线类 ======================
class Connection(QGraphicsPathItem):
    def __init__(self, start_block, end_block, line_type):
        super().__init__()
        self.start_block = start_block
        self.end_block = end_block
        self.line_type = line_type
        self.setFlag(QGraphicsPathItem.ItemIsSelectable)
        self.setZValue(-1)
        self.update_line()

    def update_line(self):
        path = QPainterPath()
        start = self.start_block.get_center()
        end = self.end_block.get_center()

        pen = QPen(self.line_type.get_color(), 2)
        pen.setStyle(self.line_type.get_style())

        for offset in self.line_type.get_offset():
            line_path = QLineF(start, end)
            if offset != 0:
                angle = line_path.angle()
                normal = QLineF.fromPolar(abs(offset), angle + 90).p2()
                line_path.translate(normal * (offset / abs(offset)))
            path.moveTo(line_path.p1())
            path.lineTo(line_path.p2())

        self.setPath(path)
        self.setPen(pen)

    def contextMenuEvent(self, event):
        menu = QMenu()
        delete_action = menu.addAction("删除")
        action = menu.exec_(event.screenPos())
        if action == delete_action:
            self.delete_connection()

    def delete_connection(self):
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
        self.draw_grid()

    def draw_grid(self):
        grid_pen = QPen(QColor(220, 220, 220), 1, Qt.DotLine)
        for x in range(-1000, 1000, 20):
            self.scene.addLine(x, -1000, x, 1000, grid_pen)
        for y in range(-1000, 1000, 20):
            self.scene.addLine(-1000, y, 1000, y, grid_pen)

    def contextMenuEvent(self, event):
        # 转换坐标系并检查该位置是否有图元
        item = self.itemAt(event.pos())

        if item is None:
            # 如果当前位置没有图元，则显示画布的菜单
            menu = QMenu()
            new_block_action = menu.addAction("新建方块")
            action = menu.exec_(self.mapToGlobal(event.pos()))
            if action == new_block_action:
                self.create_new_block(event.pos())
        else:
            # 否则，允许图元自行处理其上下文菜单事件
            super().contextMenuEvent(event)

    def create_new_block(self, pos):
        dialog = BlockEditDialog()
        dialog.name_edit.setText("New Block")
        dialog.id_edit.setText(str(DraggableBlock._next_id))
        dialog.width_edit.setText("100")
        dialog.height_edit.setText("60")

        if dialog.exec_() == QDialog.Accepted:
            name = dialog.name_edit.text()
            block_id = int(dialog.id_edit.text())
            width = float(dialog.width_edit.text())
            height = float(dialog.height_edit.text())

            # 检查ID是否重复
            existing_ids = {block.id for block in self.blocks}
            if block_id in existing_ids:
                QMessageBox.warning(self, "错误", "ID已存在！")
                return

            scene_pos = self.mapToScene(pos)
            block = DraggableBlock(name, scene_pos.x(), scene_pos.y(),
                                   width, height, block_id)
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
            # 查找目标方块
            for item in self.scene.items():
                if isinstance(item, DraggableBlock) and item != self.dragging_block:
                    if item.contains(item.mapFromScene(end_pos)):
                        end_block = item
                        break
            if end_block:
                # 创建新连接
                connection = Connection(
                    self.dragging_block,
                    end_block,
                    self.current_line_type
                )
                self.scene.addItem(connection)
                # 双向绑定
                self.dragging_block.connections.append(connection)
                end_block.connections.append(connection)
            # 清理预览线
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
        self._init_ui()

    def _init_ui(self):
        self.setWindowTitle("智能框图工具")
        self.setGeometry(100, 100, 1200, 800)

        # 工具栏
        toolbar = self.addToolBar("工具")
        line_types = [
            LineType.SINGLE,
            LineType.DOUBLE,
            LineType.TRIPLE,
            LineType.QUADRUPLE
        ]
        for lt in line_types:
            action = QAction(lt.value.capitalize(), self)
            action.triggered.connect(lambda _,lt=lt: self.set_line_type(lt))
            toolbar.addAction(action)

        # 控制面板
        control_panel = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self._create_button("导出Excel", self._export))
        layout.addWidget(self._create_button("导入Excel", self._import))
        control_panel.setLayout(layout)

        # 主界面
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(control_panel)
        main_layout.addWidget(self.canvas)
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def _create_button(self, text, callback):
        btn = QPushButton(text)
        btn.clicked.connect(callback)
        return btn

    def set_line_type(self, line_type):
        self.canvas.current_line_type = line_type
        # 更新选中连线的样式
        for item in self.canvas.scene.selectedItems():
            if isinstance(item, Connection):
                item.line_type = line_type
                item.update_line()

    def _export(self):
        try:
            # 收集模块数据
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
            # 收集连接数据
            connections = []
            for item in self.canvas.scene.items():
                if isinstance(item, Connection):
                    connections.append({
                        "StartID": item.start_block.id,
                        "EndID": item.end_block.id,
                        "Type": item.line_type.value
                    })
            # 写入Excel
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
            # 读取模块
            modules_df = pd.read_excel(path, sheet_name="Modules")
            self.canvas.scene.clear()
            self.canvas.blocks = []
            self.canvas.draw_grid()
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
            # 读取连接
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
                # 双向绑定
                start_block.connections.append(connection)
                end_block.connections.append(connection)
            QMessageBox.information(self, "导入成功",
                                    f"已导入 {len(modules_df)} 个模块和 {len(connections_df)} 条连线")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入失败: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
