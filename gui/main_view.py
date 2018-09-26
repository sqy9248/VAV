from PyQt5.QtWidgets import (QMainWindow, QAction, qApp, QHBoxLayout, QVBoxLayout, QLabel, QTextEdit, QPushButton)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QObject, pyqtSignal
from .utils import MyWidget, ViewIndex
from .path_view import Tc2TbView, Rp2PdfView
from utils import logger, g_cur_path
import logging


about_text = """Version: 0.0.1
Build on 2018-9-26
Copyright © 2018 CASCO
Author: Heisenberg
E-mail: songqingyang@casco.com.cn

需要你的电脑上已安装：
MS Word（2010或更高版本）;
Adobe Acrobat

发现Bug或有改进意见请联系我！
"""


class MyLogHandler(logging.Handler):
    def __init__(self, sig_emitter):
        super().__init__()
        self.sigEmitter = sig_emitter

    def emit(self, log_record):
        self.sigEmitter.msg_sig.emit(self.format(log_record))


class Emitter(QObject):
    msg_sig = pyqtSignal(str)


class AboutBox(MyWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('About')
        self.resize(400, 350)

        v_layout = QVBoxLayout()
        self._label = label = QLabel('About VAV')
        label.setAlignment(Qt.AlignHCenter)
        v_layout.addWidget(label)
        self._te = te = QTextEdit()
        # self._doc = doc = te.document()
        # doc.contentsChanged.connect(self._text_area_changed)
        te.setText(about_text)
        te.setReadOnly(True)
        v_layout.addWidget(te)
        v_layout.addStretch(1)

        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        bn_close = QPushButton('Close')
        bn_close.clicked.connect(self.close)
        h_layout.addWidget(bn_close)
        v_layout.addLayout(h_layout)

        self.setLayout(v_layout)
        self.setStyleSheet("QLabel{color:rgb(20,20,20);font-size:17px;font-weight:bold;font-family:Roman times;}"
                           "QTextEdit{background-color:rgb(240,240,240);font-size:12px;font-family:Roman times}")


class InitView(MyWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        v_layout = QVBoxLayout()
        # lb1 = QLabel('<a href=1>需求转表格')
        # lb1.linkActivated.connect(self._change)
        # v_layout.addWidget(lb1)
        lb2 = QLabel('<a href=2>用例转表格')
        lb2.linkActivated.connect(self._change)
        v_layout.addWidget(lb2)
        lb3 = QLabel('<a href=3>报告转PDF')
        lb3.linkActivated.connect(self._change)
        v_layout.addWidget(lb3)
        # lb4 = QLabel('<a href=4>VAT验证需求')
        # lb4.linkActivated.connect(self._change)
        # v_layout.addWidget(lb4)
        # lb5 = QLabel('<a href=5>VAT验证用例')
        # lb5.linkActivated.connect(self._change)
        # v_layout.addWidget(lb5)
        v_layout.addStretch(1)
        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        quit_btn = QPushButton('Quit')
        quit_btn.clicked.connect(qApp.quit)
        h_layout.addWidget(quit_btn)
        v_layout.addLayout(h_layout)
        self.setLayout(v_layout)
        self.setStyleSheet("QLabel{color:rgb(20,20,20);font-size:15px;font-weight:bold;font-family:Roman times;}")

    def _change(self, index):
        self.change_signal.emit(int(index))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self._init_ui()

    def _init_ui(self):
        self.setGeometry(200, 200, 300, 280)
        self.setWindowTitle('VAV')
        self.setWindowIcon(QIcon(g_cur_path + '\\res\\casco.png'))
        self._init_ui_menu()
        self._about_box = AboutBox()
        self._init_view = init_view = InitView(self)
        init_view.change_signal.connect(self._change_view_by_index)
        self.setCentralWidget(init_view)

        # Console handler
        dummy_emitter = Emitter()
        dummy_emitter.msg_sig.connect(self.statusBar().showMessage)
        console_handler = MyLogHandler(dummy_emitter)
        formatter = logging.Formatter('%(levelname)-8s: %(message)s')
        # formatter.datefmt = '%H:%M:%S'
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    def _init_ui_menu(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu('&File')
        # rq2tb_act = QAction('&需求转表格', self)
        # rq2tb_act.triggered.connect(self._change_to_rq2tb_view)
        # file_menu.addAction(rq2tb_act)
        tc2tb_act = QAction('&用例转表格', self)
        tc2tb_act.triggered.connect(self._change_to_tc2tb_view)
        file_menu.addAction(tc2tb_act)
        rp2pdf_act = QAction('&报告转PDF', self)
        rp2pdf_act.triggered.connect(self._change_to_rp2pdf_view)
        file_menu.addAction(rp2pdf_act)
        # vat_v_rq_act = QAction('&VAT验证需求', self)
        # vat_v_rq_act.triggered.connect(self._change_to_vrq_view)
        # file_menu.addAction(vat_v_rq_act)
        # vat_v_tc_act = QAction('&VAT验证用例', self)
        # vat_v_tc_act.triggered.connect(self._change_to_vtc_view)
        # file_menu.addAction(vat_v_tc_act)
        quit_act = QAction('&Quit', self)
        quit_act.setShortcut('Ctrl+Q')
        quit_act.triggered.connect(qApp.quit)
        file_menu.addAction(quit_act)
        help_menu = menu_bar.addMenu('&Help')
        about_act = QAction('&About', self)
        about_act.triggered.connect(self._show_about)
        help_menu.addAction(about_act)

    def _show_about(self):
        self._about_box.show()

    def _change_view_by_index(self, index):
        if index == ViewIndex.RQ2TB:
            self._change_to_rq2tb_view()
        elif index == ViewIndex.TC2TB:
            self._change_to_tc2tb_view()
        elif index == ViewIndex.RP2PDF:
            self._change_to_rp2pdf_view()
        elif index == ViewIndex.VAT_V_RQ:
            self._change_to_vrq_view()
        elif index == ViewIndex.VAT_V_TC:
            self._change_to_vtc_view()
        else:
            raise ValueError('wrong view index')

    # def _change_to_rq2tb_view(self):
    #     self.resize(660, 280)
    #     view = Rq2TbView(self)
    #     self.setCentralWidget(view)

    def _change_to_tc2tb_view(self):
        self.resize(660, 280)
        self.setCentralWidget(Tc2TbView())

    def _change_to_rp2pdf_view(self):
        self.resize(660, 280)
        self.setCentralWidget(Rp2PdfView())

    # def _change_to_vrq_view(self):
    #     self.resize(660, 280)
    #     self.setCentralWidget(VRqView())
    #
    # def _change_to_vtc_view(self):
    #     self.resize(660, 280)
    #     self.setCentralWidget(VTcView())
