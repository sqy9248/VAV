from PyQt5.QtWidgets import (QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QLineEdit, QFileDialog, QGridLayout,
    QComboBox, QMessageBox, qApp, QMenu, QTreeWidget)
from PyQt5.QtGui import QFont, QCursor
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from .utils import MyWidget
from word2pdf import RpDoc
import pythoncom
from enum import IntEnum
from tc2tb import TCDoc
from utils import g_cfg, change_path_to_word_style
import os


ATTACHMENTS = 'attachments'


class MsgType(IntEnum):
    ABOUT = 0
    ABOUT_QT = 1
    CRITICAL = 2
    INFORMATION = 3
    QUESTION = 4
    WARNING = 5


class Worker(QThread):
    # 一切有关GUI的，只能在主线程调用
    msgbox_sig = pyqtSignal(MsgType, str, str)

    def __init__(self, func, args):
        super(Worker, self).__init__()
        self.func = func
        self.args = args

    def run(self):
        try:
            self.func(*self.args)
        except Exception as e:
            self._msgbox(MsgType.WARNING, 'VAV', str(e))
        else:
            self._msgbox(MsgType.INFORMATION, 'VAV', 'Done!')

    def _msgbox(self, msg_type, title, text):
        self.msgbox_sig.emit(msg_type, title, text)


class PathView(MyWidget):
    _title = 'title'
    _out_ext = '.*'

    def __init__(self, parent=None):
        super().__init__(parent)
        # self._in_file = ''
        self._out_file = ''
        self._les = []
        self._bn_le = {}
        v_layout = QVBoxLayout()
        title_lb = QLabel(self._title)
        font = QFont()
        font.setBold(True)
        title_lb.setFont(font)
        title_lb.setAlignment(Qt.AlignCenter)
        v_layout.addWidget(title_lb)

        # 主布局
        self._main_layout = main_layout = QGridLayout()
        lb = QLabel('Directory:')
        main_layout.addWidget(lb, 0, 0)
        self._le = le = QLineEdit()
        self._les.append(le)
        le.textChanged.connect(self._handle_text_changed)
        main_layout.addWidget(le, 0, 2)
        browse_btn = QPushButton('Browse')
        self._bn_le[browse_btn] = le
        browse_btn.clicked.connect(self._on_clicked_browse)
        main_layout.addWidget(browse_btn, 0, 3)

        v_layout.addLayout(main_layout)
        v_layout.addStretch(1)
        h_layout2 = QHBoxLayout()
        h_layout2.addStretch(1)
        self._load_btn = load_btn = QPushButton('Load')
        load_btn.setShortcut('Ctrl+l')
        load_btn.setToolTip('Ctrl+l')
        load_btn.clicked.connect(self._load_btn_pressed)
        load_btn.setEnabled(False)
        h_layout2.addWidget(load_btn)
        quit_btn = QPushButton('Quit')
        quit_btn.clicked.connect(qApp.quit)
        quit_btn.setToolTip('Ctrl+q')
        h_layout2.addWidget(quit_btn)
        v_layout.addLayout(h_layout2)
        self.setLayout(v_layout)
        self.setStyleSheet("QLabel{color:rgb(20,20,20);font-size:13px;font-family:Roman times;}"
                           "QPushButton{font-size:13px;font-family:Roman times}")

        self._thread = thread = Worker(self._load, ())
        thread.msgbox_sig.connect(self._on_msgbox)
        thread.finished.connect(self._load_complete)

    def _on_clicked_browse(self):
        sender = self.sender()
        le = self._bn_le[sender]
        file_name, _ = QFileDialog.getOpenFileName(None, '打开', './')
        le.setText(file_name)

    def _load_btn_pressed(self):
        input_file = self._le.text()
        base_name, ext = os.path.splitext(input_file)
        init_name = base_name + self._out_ext
        file_name, _ = QFileDialog.getSaveFileName(None, '输出', init_name, '*'+self._out_ext)
        if file_name != '':
            self._out_file = change_path_to_word_style(file_name)
            self._load_btn.setEnabled(False)
            self._thread.start()

    def _load_complete(self):
        self._load_btn.setEnabled(True)

    def _load(self):
        pass

    def _handle_text_changed(self):
        for le in self._les:
            if not le.text():
                self._load_btn.setEnabled(False)
                return
        self._load_btn.setEnabled(True)

    def _on_msgbox(self, msg_type, title, text):
        if msg_type == MsgType.ABOUT:
            QMessageBox.about(self, title, text)
        elif msg_type == MsgType.INFORMATION:
            QMessageBox.information(self, title, text, QMessageBox.Ok)
        elif msg_type == MsgType.WARNING:
            QMessageBox.warning(self, title, text, QMessageBox.Cancel)
        else:
            raise ValueError('wrong message type')


# class Rq2TbView(PathView):
#     _title = '需 求 转 表 格'
#
#     def __init__(self, parent=None):
#         super().__init__(parent)
#
#     def _init_main_layout(self):
#         main_layout = QHBoxLayout()
#         lb = QLabel('Directory:')
#         main_layout.addWidget(lb)
#         self._le = le = QLineEdit()
#         self._les.append(le)
#         le.textChanged.connect(self._handle_text_changed)
#         main_layout.addWidget(le)
#         browse_btn = QPushButton('Browse')
#         self._bn_le[browse_btn] = le
#         browse_btn.clicked.connect(self._on_clicked_browse)
#         main_layout.addWidget(browse_btn)
#         return main_layout


class Tc2TbView(PathView):
    _title = '用 例 转 表 格'
    _out_ext = '.xls'

    def __init__(self, parent=None):
        super().__init__(parent)

    def _load(self):
        # 在线程中调用COM要初始化
        pythoncom.CoInitialize()
        path = self._le.text()
        doc = TCDoc(path)
        doc.export(self._out_file)


class Rp2PdfView(PathView):
    _title = '报 告 转 PDF'
    _out_ext = '.pdf'

    def __init__(self, parent=None):
        super().__init__(parent)
        self._save_view = sv = SaveView()
        sv.new_item_sig.connect(self._add_cb_item)

        main_layout = self._main_layout
        lb2 = QLabel('Attachments: ')
        main_layout.addWidget(lb2, 1, 0)
        self._cb = cb = QComboBox()
        items = g_cfg.options(ATTACHMENTS)
        cb.addItems(items)
        # 必须先addItems再connect，何哉？
        cb.currentTextChanged.connect(self._on_change_template)
        # cb.customContextMenuRequested.connect(self._generate_menu)
        main_layout.addWidget(cb, 1, 1)
        self._le2 = le2 = QLineEdit()
        self._on_change_template()
        main_layout.addWidget(le2, 1, 2)
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._on_clicked_save)
        main_layout.addWidget(save_btn, 1, 3)

    def _load(self):
        # 在线程中调用COM要初始化
        pythoncom.CoInitialize()
        path = self._le.text()
        xls_names = self._le2.text().split(';')
        doc = RpDoc(path, xls_names)
        doc.export(self._out_file)

    def _on_change_template(self):
        t = self._cb.currentText()
        value = g_cfg.get(ATTACHMENTS, t)
        self._le2.setText(value)

    def _on_clicked_save(self):
        self._save_view.set_name(self._cb.currentText())
        self._save_view.set_value(self._le2.text())
        self._save_view.show()

    def _add_cb_item(self, new_item):
        item_total = self._cb.count()
        has_item = False
        for i in range(item_total):
            if self._cb.itemText(i) == new_item:
                has_item = True
                break
        if not has_item:
            self._cb.addItem(new_item)
            self._cb.setCurrentText(new_item)
            # self._cb.setFixedWidth(len(new_item))
        self._on_change_template()

    # def _generate_menu(self, pos):
    #     menu = QMenu()
    #     menu.addAction('del')
    #     action = menu.exec_(self._cb.mapToGlobal(pos))


class SaveView(MyWidget):
    new_item_sig = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.resize(500, 280)
        v_layout = QVBoxLayout()
        lb1 = QLabel('Name: ')
        v_layout.addWidget(lb1)
        self._le1 = le1 = QLineEdit()
        v_layout.addWidget(le1)
        lb2 = QLabel('Value: ')
        v_layout.addWidget(lb2)
        self._le = le = QLineEdit()
        v_layout.addWidget(le)
        v_layout.addStretch(1)
        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._on_clicked_save)
        h_layout.addWidget(save_btn)
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.close)
        h_layout.addWidget(close_btn)
        v_layout.addLayout(h_layout)
        self.setLayout(v_layout)

    def set_name(self, name):
        self._le1.setText(name)

    def set_value(self, value):
        self._le.setText(value)

    def _on_clicked_save(self):
        name = self._le1.text()
        if name == '':
            QMessageBox.warning(self, 'VAV', '\'Name: \'不能为空', QMessageBox.Cancel)
        else:
            value = self._le.text()
            if value == '':
                QMessageBox.warning(self, 'VAV', '\'Value: \'不能为空', QMessageBox.Cancel)
            else:
                g_cfg.set(ATTACHMENTS, name, value)
                self.new_item_sig.emit(name)
                self.close()


# class VRqView(PathView):
#     _title = 'VAT 验 证 需 求'
#
#     def __init__(self, parent=None):
#         super().__init__(parent)
#
#     def _init_main_layout(self):
#         grid_layout = QGridLayout()
#
#         lb1 = QLabel('VAT Directory:')
#         grid_layout.addWidget(lb1, 0, 0)
#         self._le1 = le1 = QLineEdit()
#         self._les.append(le1)
#         le1.textChanged.connect(self._handle_text_changed)
#         grid_layout.addWidget(le1, 0, 1)
#         browse_btn1 = QPushButton('Browse')
#         self._bn_le[browse_btn1] = le1
#         browse_btn1.clicked.connect(self._on_clicked_browse)
#         grid_layout.addWidget(browse_btn1, 0, 2)
#
#         lb2 = QLabel('Requirement Directory:')
#         grid_layout.addWidget(lb2, 1, 0)
#         self._le2 = le2 = QLineEdit()
#         self._les.append(le2)
#         le2.textChanged.connect(self._handle_text_changed)
#         grid_layout.addWidget(le2, 1, 1)
#         browse_btn2 = QPushButton('Browse')
#         self._bn_le[browse_btn2] = le2
#         browse_btn2.clicked.connect(self._on_clicked_browse)
#         grid_layout.addWidget(browse_btn2, 1, 2)
#
#         lb3 = QLabel('VAT first row:')
#         grid_layout.addWidget(lb3, 2, 0)
#         combo = QComboBox()
#         items = [str(i) for i in range(1, 7)]
#         combo.addItems(items)
#         combo.setCurrentText('4')
#         grid_layout.addWidget(combo, 2, 1)
#
#         lb4 = QLabel('VAT sheet name:')
#         grid_layout.addWidget(lb4, 3, 0)
#         le3 = QLineEdit()
#         le3.textChanged.connect(self._handle_text_changed)
#         self._les.append(le3)
#         grid_layout.addWidget(le3, 3, 1)
#
#         return grid_layout
#
#
# class VTcView(PathView):
#     _title = 'VAT 验 证 用 例'
#
#     def __init__(self, parent=None):
#         super().__init__(parent)
#
#     def _init_main_layout(self):
#         grid_layout = QGridLayout()
#
#         lb1 = QLabel('VAT Directory:')
#         grid_layout.addWidget(lb1, 0, 0)
#         self._le1 = le1 = QLineEdit()
#         self._les.append(le1)
#         le1.textChanged.connect(self._handle_text_changed)
#         grid_layout.addWidget(le1, 0, 1)
#         browse_btn1 = QPushButton('Browse')
#         self._bn_le[browse_btn1] = le1
#         browse_btn1.clicked.connect(self._on_clicked_browse)
#         grid_layout.addWidget(browse_btn1, 0, 2)
#
#         lb2 = QLabel('Test Case Directory')
#         grid_layout.addWidget(lb2, 1, 0)
#         self._le2 = le2 = QLineEdit()
#         self._les.append(le2)
#         le2.textChanged.connect(self._handle_text_changed)
#         grid_layout.addWidget(le2, 1, 1)
#         browse_btn2 = QPushButton('Browse')
#         self._bn_le[browse_btn2] = le2
#         browse_btn2.clicked.connect(self._on_clicked_browse)
#         grid_layout.addWidget(browse_btn2, 1, 2)
#
#         lb3 = QLabel('VAT first row:')
#         grid_layout.addWidget(lb3, 2, 0)
#         combo = QComboBox()
#         items = [str(i) for i in range(1, 7)]
#         combo.addItems(items)
#         combo.setCurrentText('4')
#         grid_layout.addWidget(combo, 2, 1)
#
#         lb4 = QLabel('VAT sheet name:')
#         grid_layout.addWidget(lb4, 3, 0)
#         le3 = QLineEdit()
#         le3.textChanged.connect(self._handle_text_changed)
#         self._les.append(le3)
#         grid_layout.addWidget(le3, 3, 1)
#
#         return grid_layout
