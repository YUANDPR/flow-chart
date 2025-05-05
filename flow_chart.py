import sys
import os
import pandas as pd
import random
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QMenu, QAction, QVBoxLayout, QWidget, QPushButton, QFileDialog, QGraphicsTextItem,
    QToolBar, QMessageBox, QInputDialog, QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QGraphicsPathItem,
    QGraphicsLineItem, QGraphicsPixmapItem, QSlider, QHBoxLayout, QLabel
)
from PyQt5.QtCore import Qt, QPointF, QLineF, QRectF, QSizeF
from PyQt5.QtGui import QBrush, QPen, QColor, QPainter, QTransform, QCursor, QPainterPath, QIcon, QPixmap
from enum import Enum

def resource_path(relative_path):
    """ 获取资源文件的绝对路径，用于 PyInstaller 打包后的访问 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def generate_scattered_position(placed_positions, width, height,
                                center=(0, 0),
                                spread=1000, spacing=200):
    """
    在指定区域内生成一个不与已有矩形重叠的 (x, y) 位置。

    参数:
        placed_positions: 已放置的 (x, y, w, h) 列表
        width, height: 当前模块的宽高
        center: 分布中心点 (x, y)
        spread: 随机分布范围（半径）
        spacing: 模块之间最小间距

    返回:
        (x, y): 合适的位置
    """
    cx, cy = center
    attempts = 0
    # 第一个模块直接放在中心
    if not placed_positions:
        return cx, cy
    while attempts < 500:
        x = cx + random.randint(-spread, spread)
        y = cy + random.randint(-spread, spread)
        if not is_overlapping(x, y, width, height, placed_positions, spacing):
            return x, y
        attempts += 1
    # 如果找不到合适位置，返回中心点
    return cx, cy

def is_overlapping(x, y, width, height, placed_positions, spacing=20):
    for px, py, pw, ph in placed_positions:
        # 添加 spacing 边距检查
        if not (x + width < px - spacing or     # 当前模块在已有模块左边
                x > px + pw + spacing or         # 当前模块在已有模块右边
                y + height < py - spacing or     # 当前模块在已有模块上边
                y > py + ph + spacing):          # 当前模块在已有模块下边
            return True
    return False

# ====================== 枚举定义 ======================
class LineType(Enum):
    SINGLE = "单实线"  # 黑色实线线
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
        return Qt.SolidLine if self == LineType.SINGLE else Qt.SolidLine

    def to_number(self):
        mapping = {
            LineType.SINGLE: 4,
            LineType.DOUBLE: 3,
            LineType.TRIPLE: 2,
            LineType.QUADRUPLE: 1
        }
        return mapping[self]

    def from_number(number):
        mapping = {
            4: LineType.SINGLE,
            3: LineType.DOUBLE,
            2: LineType.TRIPLE,
            1: LineType.QUADRUPLE
        }
        return mapping[number]

# ====================== 方块编辑对话框 ======================
class BlockEditDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("工作组属性")
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
        edit_action = menu.addAction("编辑工作组")
        # delete_action = menu.addAction("删除")
        action = menu.exec_(event.screenPos())

        if action == edit_action:
            self.edit_properties()
        # if action == delete_action:
        #     self.delete_block()

    def edit_properties(self):
        dialog = BlockEditDialog()
        dialog.name_edit.setText(self.name)
        dialog.id_edit.setText(str(self.id))
        dialog.width_edit.setText(str(self.rect().width()))
        dialog.height_edit.setText(str(self.rect().height()))

        if dialog.exec_() == QDialog.Accepted:
            # 更新工作组属性
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
        # delete_action = menu.addAction("删除")
        # action = menu.exec_(event.screenPos())
        # if action == delete_action:
        #     self.delete_connection()

    def delete_connection(self):
        self.start_block.connections.remove(self)
        self.end_block.connections.remove(self)
        self.scene().removeItem(self)


# ====================== 画布类 ======================
class Canvas(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.scene = QGraphicsScene()
        self.setRenderHint(QPainter.Antialiasing, False)
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setDragMode(QGraphicsView.RubberBandDrag)
        self.blocks = []
        self.dragging_block = None
        self.preview_line = None
        self.current_line_type = LineType.SINGLE

        # 添加背景图片
        self.background_image = None
        self._load_background_image(resource_path("background.png"))  # 替换为你的图片路径

        # self.draw_grid()

    def fit_background_to_view(self):
        if not self.background_image:
            return

        view_width = self.viewport().width()
        view_height = self.viewport().height()

        pixmap = self.background_image.pixmap()
        if pixmap.isNull():
            return

        img_w = pixmap.width()
        img_h = pixmap.height()
        if img_w == 0 or img_h == 0:
            return

        # 按宽度拉伸，保持宽高比
        scale = view_width / img_w
        scaled_h = img_h * scale

        # 设置缩放
        self.background_image.setScale(scale)

        # 设置图片中心为 (0, 0)，即 scene 原点
        x = -img_w * scale / 2
        y = -img_h * scale / 2
        self.background_image.setPos(x, y)

    def _load_background_image(self, image_path):
        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            # 设置缩放模式为快速缩放，避免模糊
            scaled_pixmap = pixmap.scaled(
                pixmap.size() * 2,  # 可选：提高分辨率
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation  # 或 Qt.FastTransformation
            )
            self.background_image = QGraphicsPixmapItem(scaled_pixmap)
            self.background_image.setZValue(-2)
            self.scene.addItem(self.background_image)
            self.fit_background_to_view()

    def center_background_image(self):
        if self.background_image:
            rect = self.background_image.boundingRect()
            scene_rect = self.scene.sceneRect()
            x = scene_rect.width() / 2 - rect.width() / 2
            y = scene_rect.height() / 2 - rect.height() / 2
            self.background_image.setPos(x, y)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit_background_to_view()

    # def draw_grid(self):
    #     grid_pen = QPen(QColor(220, 220, 220), 1, Qt.DotLine)
    #     for x in range(-1000, 1000, 20):
    #         self.scene.addLine(x, -1000, x, 1000, grid_pen)
    #     for y in range(-1000, 1000, 20):
    #         self.scene.addLine(-1000, y, 1000, y, grid_pen)

    # def contextMenuEvent(self, event):
    #     # 转换坐标系并检查该位置是否有图元
    #     item = self.itemAt(event.pos())
    #
    #     if item is None:
    #         # 如果当前位置没有图元，则显示画布的菜单
    #         menu = QMenu()
    #         new_block_action = menu.addAction("新建工作组")
    #         action = menu.exec_(self.mapToGlobal(event.pos()))
    #         if action == new_block_action:
    #             self.create_new_block(event.pos())
    #     else:
    #         # 否则，允许图元自行处理其上下文菜单事件
    #         super().contextMenuEvent(event)

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

    # def mousePressEvent(self, event):
    #     if event.button() == Qt.LeftButton:
    #         item = self.itemAt(event.pos())
    #         if isinstance(item, DraggableBlock):
    #             self._start_connection(item)
    #             return
    #     super().mousePressEvent(event)

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
        self.setWindowTitle("重大工作组流程图工具")
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowIcon(QIcon(resource_path("logo.png")))  # 替换为你的图标文件路径
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

        # 缩放控件
        zoom_layout = QHBoxLayout()
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(10)
        self.zoom_slider.setMaximum(300)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(self._zoom_canvas)

        self.zoom_label = QLabel("100%")
        self.zoom_label.setFixedWidth(40)

        zoom_layout.addWidget(self.zoom_slider)
        zoom_layout.addWidget(self.zoom_label)

        layout.addLayout(zoom_layout)

        control_panel.setLayout(layout)
        # 主界面
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(control_panel)
        main_layout.addWidget(self.canvas)
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def _zoom_canvas(self, value):
        scale_factor = value / 100.0
        self.canvas.resetTransform()
        self.canvas.scale(scale_factor, scale_factor)
        self.zoom_label.setText(f"{value}%")

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
            # 弹出保存文件对话框
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存文件",
                "diagram.xlsx",  # 默认文件名
                "Excel 文件 (*.xlsx)",  # 文件类型过滤
                options=options
            )

            if not file_path:
                return  # 用户取消操作

            # 收集模块数据
            modules = []
            for block in self.canvas.blocks:
                modules.append({
                    "序号": block.id,
                    "group": block.name,
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
                        "起始编号": item.start_block.id,
                        "结束编号": item.end_block.id,
                        "线型": item.line_type.to_number()
                    })

            # 写入Excel
            with pd.ExcelWriter(file_path) as writer:
                pd.DataFrame(modules).to_excel(writer, sheet_name="group", index=False)
                pd.DataFrame(connections).to_excel(writer, sheet_name="relation", index=False)

            QMessageBox.information(self, "导出成功", f"数据已保存到\n{file_path}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def _import(self):
        try:
            path, _ = QFileDialog.getOpenFileName(
                self, "打开文件", "", "Excel文件 (*.xlsx)")
            if not path:
                return

            CENTER_X, CENTER_Y = 0, 0  # 分布中心点
            SPREAD_RADIUS = 800  # 分散半径
            MIN_SPACING = 100  # 模块之间最小间距

            # 读取模块
            modules_df = pd.read_excel(path, sheet_name="group")
            for item in self.canvas.scene.items():
                if isinstance(item, (DraggableBlock, Connection)):
                    self.canvas.scene.removeItem(item)
            self.canvas.blocks = []
            # self.canvas.draw_grid()
            id_map = {}
            placed_positions = []
            for _, row in modules_df.iterrows():
                block_id = row["序号"]
                name = row.get("group", "未知模块")
                x = row.get("X", None)
                y = row.get("Y", None)
                width = row.get("Width", 100)
                height = row.get("Height", 60)

                # 如果不存在 X/Y/Width/Height 列，则使用默认值
                if x is None or y is None:
                    x, y = generate_scattered_position(
                        placed_positions, width, height,
                        center=(CENTER_X, CENTER_Y),
                        spread=SPREAD_RADIUS,
                        spacing=MIN_SPACING
                    )


                block = DraggableBlock(name, x, y, width, height, block_id=block_id)
                self.canvas.scene.addItem(block)
                self.canvas.blocks.append(block)
                id_map[block.id] = block
                # 最后再加入 placed_positions
                placed_positions.append((x, y, width, height))

            # 读取连接
            connections_df = pd.read_excel(path, sheet_name="relation")
            for _, row in connections_df.iterrows():
                start_block = id_map[row["起始编号"]]
                end_block = id_map[row["结束编号"]]
                line_type_number = int(row["线型"])  # 假设 Excel 中列名为 "线型"
                line_type = LineType.from_number(line_type_number)
                connection = Connection(
                    start_block,
                    end_block,
                    line_type
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
