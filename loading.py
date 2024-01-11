from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QVBoxLayout, QDialog
from PySide6.QtGui import QMovie


# 创建一个封装了加载动画的QDialog类
class LoadingAnimationDialog(QDialog):
    def __init__(self, gif_path, parent=None):
        super().__init__(parent, Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)  # 设置窗口背景透明
        self.setModal(True)  # 设置为模态窗口
        self.init_ui(gif_path)

    def init_ui(self, gif_path):
        self.loading_label = QLabel(self)
        self.loading_movie = QMovie(gif_path)
        self.loading_label.setMovie(self.loading_movie)

        layout = QVBoxLayout()
        layout.addWidget(self.loading_label)
        self.setLayout(layout)
        self.loading_movie.start()

    def closeEvent(self, event):
        self.loading_movie.stop()
        event.accept()


# 封装加载动画的类
class LoadingAnimation:
    def __init__(self, gif_path, parent=None):
        self.dialog = LoadingAnimationDialog(gif_path, parent)

    def show(self):
        self.dialog.show()

    def hide(self):
        self.dialog.close()