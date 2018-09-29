from win32com.client.dynamic import Dispatch, ERRORS_BAD_CONTEXT
import winerror
import os
import win32api
from PyUserInput.pykeyboard import PyKeyboard
import time
from utils import logger, change_path_to_word_style, f_int
import shutil

ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)

EMBEDDED_OLE_OBJ = 1
FORMAT_DOCX = 16
FORMAT_PDF = 17
MARKDOWN_NONE = 0
FINAL_VISION = 0


_p_join = os.path.join
_p_isfile = os.path.isfile
_p_splitext = os.path.splitext


class RpDoc(object):
    temp_path = 'd:\\vav_temp'

    def __init__(self, path):
        self._path = path = change_path_to_word_style(path)
        if not os.path.isabs(path):
            logger.warning('%s不是有效路径' % path)
            raise ValueError('路径无效')
        self._base_name, self._extension = _p_splitext(path)
        if self._extension not in ('.doc', '.docx'):
            logger.warning('%s不是有效文件类型,请选择.doc或.docx类型的文件' % self._extension)
            raise ValueError('文件类型无效')
        # shutil.rmtree(self.temp_path)
        if not os.path.isdir(self.temp_path):
            os.mkdir(self.temp_path)

    def _clear_temp_path(self):
        if os.path.isdir(self.temp_path):
            shutil.rmtree(self.temp_path)

    def export(self, out_path):
        word = Dispatch('Word.Application')
        version = word.Version
        if f_int(version) < 13:
            logger.warning('Word 版本应该为2013或更高')
            raise IOError('Word 版本过低')
        word.Visible = 0
        word.DisplayAlerts = 0
        excel = Dispatch('Excel.Application')
        excel.Visible = 0
        excel.DisplayAlerts = 0
        self._extract_attachments(word, excel)
        excel.Quit()
        self._set_word_revision_view_final(word)
        word.ActiveDocument.SaveAs(out_path, FORMAT_PDF)
        word.ActiveDocument.Save()
        word.Quit()
        self._pdf_add_attachment(out_path)
        self._clear_temp_path()

    @staticmethod
    def _set_word_revision_view_final(word):
        version = word.Version
        # TBD: 兼容低版本office
        if version == '16.0':
            word.ActiveWindow.View.RevisionsFilter.Markup = MARKDOWN_NONE
            word.ActiveWindow.View.RevisionsFilter.View = FINAL_VISION
        # elif version == '10.0':
        #     word.ActiveWindow.View.ShowRevisionsAndComments = False
        #     word.ActiveWindow.View.RevisionsView = FINAL_VISION
        else:
            logger.warning('word 版本过低，PDF可能留有标记')
            # raise IOError('不支持的Word版本:%s.支持16,10.' % version)

    def _extract_attachments(self, word, excel):
        doc = word.Documents.Open(self._path)
        shapes = doc.InlineShapes
        k = PyKeyboard()
        for shape in shapes:
            if shape.Type == EMBEDDED_OLE_OBJ:
                field = shape.Field.Code.Text.lstrip().rstrip()
                # ole object 用复制粘贴提取
                if field == 'EMBED Package':
                    shape.Select()
                    word.Selection.Copy()
                    win32api.ShellExecute(0, 'open', self.temp_path, '', '', 1)
                    time.sleep(0.5)
                    k.press_key(k.control_key)
                    k.tap_key('v')
                    k.release_key(k.control_key)
                    time.sleep(0.1)
                # worksheet object 用另存为提取
                elif field in ('EMBED Excel.Sheet.12', 'EMBED Excel.Sheet.8'):
                    shape.OLEFormat.Open()
                    name = shape.OLEFormat.IconLabel
                    file_path = _p_join(self.temp_path, name)
                    excel.ActiveWorkbook.SaveAs(Filename=file_path)
                    excel.ActiveWorkbook.Close()
        word.ActiveDocument.Save()

    def _pdf_add_attachment(self, pdf_path):
        logger.info('开始向PDF插入附件')
        pd_doc = Dispatch('AcroExch.PDDoc')
        assert _p_isfile(pdf_path)
        if pd_doc.Open(pdf_path):
            jso = pd_doc.GetJSObject()
            attachments = os.listdir(self.temp_path)
            for name in attachments:
                path = _p_join(self.temp_path, name)
                jso.importDataObject(name, path)
            jso.saveAs(pdf_path)
        pd_doc.Close()
        logger.info('向PDF插入附件成功')


# if __name__ == '__main__':
#
#     doc_path = 'f:\\python_project\\vat_tool\\RA_12008_TSRS-KA_CASCO_TSRS ATO Function Test Report.doc'
#     if os.path.isdir(RpDoc.temp_path):
#         shutil.rmtree(RpDoc.temp_path)
#     os.mkdir(RpDoc.temp_path)
#     a = RpDoc(doc_path)
#     a.export('d:\\11.pdf')



