import sys
import random
import ctypes
import numpy as np
from mss import mss
from PIL import Image
import cv2
import pyautogui
import time
import os
import logging
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QImage, QPixmap, QKeyEvent, QPainter, QColor
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QDesktopWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog
from ultralytics import YOLO

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Windows API结构定义
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]

class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("mi", MOUSEINPUT)]

INPUT_MOUSE = 0
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000

user32 = ctypes.windll.user32
screen_w = user32.GetSystemMetrics(0)
screen_h = user32.GetSystemMetrics(1)

def win32_move(x, y, rand_offset=True):
    """Windows API鼠标移动函数"""
    try:
        if rand_offset:  # 添加随机偏移避免绝对精准
            x += random.randint(-3, 3)
            y += random.randint(-3, 3)

        # 转换为Windows绝对坐标系
        x_abs = int(x * 65535 / screen_w)
        y_abs = int(y * 65535 / screen_h)

        # 构建输入事件
        mi = MOUSEINPUT(
            dx=x_abs,
            dy=y_abs,
            mouseData=0,
            dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
            time=0,
            dwExtraInfo=ctypes.pointer(ctypes.c_ulong(0))
        )

        input_struct = INPUT(INPUT_MOUSE, mi)
        user32.SendInput(1, ctypes.pointer(input_struct), ctypes.sizeof(input_struct))
    except Exception as e:
        logging.error(f"鼠标移动出错: {e}")

def check_admin_permission():
    """检查是否具有管理员权限，若没有则请求权限"""
    if not ctypes.windll.shell32.IsUserAnAdmin():
        logging.info("正在请求管理员权限...")
        # 获取当前脚本的完整路径（处理空格）
        script_path = os.path.abspath(sys.argv[0])

        # 使用更可靠的参数格式
        params = f'"{sys.executable}" "{script_path}"'

        # 调用ShellExecuteEx获取返回值
        class SHELLEXECUTEINFOW(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("fMask", ctypes.c_ulong),
                ("hwnd", ctypes.c_void_p),
                ("lpVerb", ctypes.c_wchar_p),
                ("lpFile", ctypes.c_wchar_p),
                ("lpParameters", ctypes.c_wchar_p),
                ("lpDirectory", ctypes.c_wchar_p),
                ("nShow", ctypes.c_int),
                ("hInstApp", ctypes.c_void_p),
                ("lpIDList", ctypes.c_void_p),
                ("lpClass", ctypes.c_wchar_p),
                ("hKeyClass", ctypes.c_void_p),
                ("dwHotKey", ctypes.c_ulong),
                ("hMonitor", ctypes.c_void_p),
                ("hProcess", ctypes.c_void_p)
            ]

        sei = SHELLEXECUTEINFOW()
        sei.cbSize = ctypes.sizeof(SHELLEXECUTEINFOW)
        sei.lpVerb = "runas"
        sei.lpFile = sys.executable
        sei.lpParameters = f'"{script_path}"'
        sei.nShow = 1  # SW_SHOWNORMAL

        if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei)):
            error_code = ctypes.windll.kernel32.GetLastError()
            logging.error(f"权限请求失败，错误代码: 0x{error_code:08X}")
            input("按回车退出...")
            sys.exit(1)
        else:
            logging.info("权限请求成功，重新启动程序...")
            sys.exit(0)

def load_model(model_path):
    """加载YOLO模型"""
    try:
        model = YOLO(model_path)
        return model
    except Exception as e:
        logging.error(f"加载模型失败: {e}")
        sys.exit(1)

class SettingsWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # 检测框是否显示
        self.show_box_checkbox = QCheckBox("显示检测框")
        self.show_box_checkbox.setChecked(self.main_window.show_box)
        self.show_box_checkbox.stateChanged.connect(self.update_settings)
        layout.addWidget(self.show_box_checkbox)

        # 检测框颜色
        color_options = ["红色", "绿色", "蓝色"]
        self.color_combobox = QComboBox()
        self.color_combobox.addItems(color_options)
        self.color_combobox.setCurrentIndex(self.get_color_index(self.main_window.box_color))
        self.color_combobox.currentIndexChanged.connect(self.update_settings)
        layout.addWidget(self.color_combobox)

        # 置信度大小
        self.confidence_spinbox = QDoubleSpinBox()
        self.confidence_spinbox.setRange(0.0, 1.0)
        self.confidence_spinbox.setSingleStep(0.01)
        self.confidence_spinbox.setValue(self.main_window.conf_threshold)
        self.confidence_spinbox.valueChanged.connect(self.update_settings)
        layout.addWidget(self.confidence_spinbox)

        # 更改所使用的模型
        self.model_button = QPushButton("选择模型")
        self.model_button.clicked.connect(self.select_model)
        layout.addWidget(self.model_button)

        # 开关检测功能
        self.detection_switch = QPushButton("关闭检测功能" if self.main_window.detection_enabled else "开启检测功能")
        self.detection_switch.clicked.connect(self.toggle_detection)
        layout.addWidget(self.detection_switch)

        # 保存设置按钮
        save_button = QPushButton("保存设置")
        save_button.clicked.connect(self.save_settings)
        layout.addWidget(save_button)

        self.setLayout(layout)
        self.setWindowTitle("设置")
        self.setGeometry(100, 100, 300, 250)

    def get_color_index(self, color):
        """根据颜色获取对应的索引"""
        if color == QColor(255, 0, 0):
            return 0
        elif color == QColor(0, 255, 0):
            return 1
        elif color == QColor(0, 0, 255):
            return 2
        return 0

    def update_settings(self):
        self.main_window.show_box = self.show_box_checkbox.isChecked()
        color_index = self.color_combobox.currentIndex()
        self.main_window.box_color = self.get_color_by_index(color_index)
        self.main_window.conf_threshold = self.confidence_spinbox.value()

    def get_color_by_index(self, index):
        """根据索引返回颜色"""
        if index == 0:
            return QColor(255, 0, 0)
        elif index == 1:
            return QColor(0, 255, 0)
        elif index == 2:
            return QColor(0, 0, 255)
        return QColor(255, 0, 0)

    def select_model(self):
        file_dialog = QFileDialog()
        model_path, _ = file_dialog.getOpenFileName(self, "选择模型文件", "", "模型文件 (*.pt)")
        if model_path:
            self.main_window.model = load_model(model_path)

    def toggle_detection(self):
        self.main_window.detection_enabled = not self.main_window.detection_enabled
        self.detection_switch.setText("关闭检测功能" if self.main_window.detection_enabled else "开启检测功能")

    def save_settings(self):
        self.close()

class DetectionOverlay(QMainWindow):
    def __init__(self):
        super().__init__()

        # 初始化参数
        self.screen_width, self.screen_height = pyautogui.size()
        self.detect_size = 640
        self.show_box = True
        self.box_color = QColor(255, 0, 0)
        self.conf_threshold = 0.4
        self.target_class = "head"
        self.detection_enabled = True

        # 设置窗口大小
        self.setGeometry((self.screen_width - self.detect_size) // 2, (self.screen_height - self.detect_size) // 2, self.detect_size, self.detect_size)

        # 窗口设置
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # 初始化覆盖图像和标签
        self.overlay_image = QImage(self.detect_size, self.detect_size, QImage.Format_ARGB32)
        self.detection_label = QLabel(self)
        self.detection_label.setPixmap(QPixmap.fromImage(self.overlay_image))
        self.detection_label.setGeometry(0, 0, self.detect_size, self.detect_size)

        # 使用绝对路径加载模型
        model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cf_best.pt')
        self.model = load_model(model_path)

        self.sct = mss()

        # 添加输入延迟随机化参数
        self.move_delay = 0.033  # 基础延迟
        self.last_move_time = 0

        # 设置按钮
        settings_button = QPushButton("设置", self)
        settings_button.setGeometry(10, 10, 80, 30)
        settings_button.clicked.connect(self.open_settings)

    def open_settings(self):
        self.settings_window = SettingsWindow(self)
        self.settings_window.show()

    def update_detection(self):
        if not self.detection_enabled:
            self.overlay_image.fill(Qt.transparent)
            self.detection_label.setPixmap(QPixmap.fromImage(self.overlay_image))
            self.detection_label.update()
            self.update()
            return

        try:
            # 截图和检测逻辑
            self.monitor = {
                "left": (self.screen_width - self.detect_size) // 2,
                "top": (self.screen_height - self.detect_size) // 2,
                "width": self.detect_size,
                "height": self.detect_size
            }
            frame = np.array(self.sct.grab(self.monitor))
            frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
            results = self.model(frame, verbose=False, imgsz=self.detect_size)[0]

            # 绘制逻辑
            self.overlay_image.fill(Qt.transparent)
            painter = QPainter(self.overlay_image)
            painter.setPen(self.box_color)

            current_time = time.time()
            detected = False  # 标记是否检测到目标
            for data in results.boxes.data.tolist():
                xmin, ymin, xmax, ymax, conf, cls_id = data[:6]
                if conf < self.conf_threshold or results.names[int(cls_id)] != self.target_class:
                    continue

                detected = True  # 检测到目标
                # 从推理结果中获取原始图像尺寸
                orig_shape = results.orig_shape
                xmin = int(xmin * self.detect_size / orig_shape[1])
                ymin = int(ymin * self.detect_size / orig_shape[0])
                xmax = int(xmax * self.detect_size / orig_shape[1])
                ymax = int(ymax * self.detect_size / orig_shape[0])

                center_x = (xmin + xmax) // 2
                center_y = (ymin + ymax) // 2

                # 添加移动冷却和随机延迟
                if current_time - self.last_move_time > self.move_delay * random.uniform(0.8, 1.2):
                    global_center_x = self.monitor["left"] + center_x
                    global_center_y = self.monitor["top"] + center_y
                    win32_move(global_center_x, global_center_y)
                    self.last_move_time = current_time

                if self.show_box:
                    if 0 <= xmin < self.detect_size and 0 <= ymin < self.detect_size and \
                            0 <= xmax < self.detect_size and 0 <= ymax < self.detect_size:
                        try:
                            painter.drawRect(int(xmin), int(ymin), int(xmax - xmin), int(ymax - ymin))
                            label_text = f"{results.names[int(cls_id)]} {conf:.2f}"
                            painter.drawText(int(xmin), int(ymin) - 10, label_text)
                        except Exception as e:
                            logging.error(f"绘制检测框和标注信息出错: {e}")
                    else:
                        logging.warning(f"检测框坐标超出窗口范围: xmin={xmin}, ymin={ymin}, xmax={xmax}, ymax={ymax}")

            if not detected:
                logging.info("未检测到符合条件的目标。")

            painter.end()
            self.detection_label.setPixmap(QPixmap.fromImage(self.overlay_image))
            self.detection_label.update()
            self.update()

        except Exception as e:
            logging.error(f"更新检测结果时出错: {e}", exc_info=True)

if __name__ == "__main__":
    check_admin_permission()
    app = QApplication(sys.argv)
    overlay = DetectionOverlay()
    overlay.show()

    timer = QTimer()
    timer.timeout.connect(overlay.update_detection)
    timer.start(100)

    sys.exit(app.exec_())