from PyQt5.QtWidgets import QWidget, QHBoxLayout
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QPen
from PyQt5.QtCore import Qt, pyqtSignal
from enum import IntEnum
from utils import g_cur_path


class ViewIndex(IntEnum):
    RQ2TB = 1
    TC2TB = 2
    RP2PDF = 3
    VAT_V_RQ = 4
    VAT_V_TC = 5


class SidePix(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedWidth(100)

    def paintEvent(self, e):
        qp = QPainter()
        qp.begin(self)
        self._draw_widget(qp)
        qp.end()

    def _draw_widget(self, qp):
        size = self.size()
        w = size.width()
        h = size.height()

        qp.setPen(QColor(255, 255, 255))
        qp.setBrush(QColor(255, 255, 255))
        qp.drawRect(0, 0, w, h)

        pen = QPen(QColor(20, 20, 20), 1, Qt.SolidLine)
        qp.setPen(pen)
        qp.setBrush(Qt.NoBrush)
        qp.drawRect(0, 0, w - 1, h - 1)

        pix = QPixmap(g_cur_path + '\\res\\sf2.png')
        pix_w = 98
        pix_h = 200
        qp.drawPixmap(w / 2 - pix_w / 2, h / 2 - pix_h / 2, pix_w, pix_h, pix)


class MyWidget(QWidget):
    change_signal = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon(g_cur_path + '\\res\\casco.png'))
        base_layout = QHBoxLayout(self)
        sp = SidePix()
        base_layout.addWidget(sp)
        self.true_widget = true_widget = QWidget()
        base_layout.addWidget(true_widget)

    def setLayout(self, q_layout):
        self.true_widget.setLayout(q_layout)
