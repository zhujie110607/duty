import os

from PySide6.QtCore import Signal, QThread
from PySide6.QtWidgets import QWidget, QApplication, QFileDialog
import qtmodern.styles
import qtmodern.windows
import common
from loading import LoadingAnimation

from ui.aa_ui import Ui_Form


class LongTaskThread(QThread):
    finished = Signal()

    def __init__(self):
        super().__init__()

    def run(self):
        try:
            from py_file.gitdata import GitData
            gitdata = GitData()
            gitdata.summarize()
        except Exception as e:
            common.show_message(e, 0)
        finally:
            self.finished.emit()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.bind()

    def bind(self):
        self.ui.pushButton.clicked.connect(self.click)
        self.ui.longTaskThread = LongTaskThread()
        self.ui.finished = self.ui.longTaskThread.finished.connect(self.on_finished)

    def click(self):
        # 打开文件对话框获取文件路径
        common.path = QFileDialog.getOpenFileName(self, "请选择数据源表文件", '.', "Excel文件(*.xlsx)")[0]
        if not common.path:
            return
        current_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前文件所在目录路径
        self.ui.LoadingAnimation = LoadingAnimation(os.path.join(current_dir, 'gif', '33.gif'))
        self.ui.pushButton.setEnabled(False)
        self.ui.LoadingAnimation.show()
        self.ui.longTaskThread.start()

    def on_finished(self):
        self.ui.pushButton.setEnabled(True)
        self.ui.LoadingAnimation.hide()


if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    qtmodern.styles.dark(app)
    mw = qtmodern.windows.ModernWindow(window)
    mw.move(app.primaryScreen().size().width() / 2 - window.size().width() / 2,
            app.primaryScreen().size().height() / 2 - window.size().height() / 2)
    mw.show()
    app.exec()
