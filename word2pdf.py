"""
change doc/docx format to pdf, and add attachment

"""
# !usr/bin/python
from win32com.client.dynamic import Dispatch, ERRORS_BAD_CONTEXT
import winerror
import zipfile
import os
import shutil
from utils import logger, change_path_to_word_style, g_cfg, Config
ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)

_p_join = os.path.join
_p_isfile = os.path.isfile
_p_splitext = os.path.splitext

FORMAT_DOCX = 16
FORMAT_PDF = 17
MARKDOWN_NONE = 0
FINAL_VISION = 0


class RpDoc(object):
    _xls_dir = 'word\\embeddings'

    def __init__(self, path, xls_names):
        self._word = None

        self._path = path = change_path_to_word_style(path)
        if not os.path.isabs(path):
            logger.warning('%s不是有效路径' % path)
            raise ValueError('路径无效')
        self._base_name, self._extension = _p_splitext(path)
        # self._pdf_path = self._base_name + '.pdf'
        if self._extension not in ('.doc', '.docx'):
            logger.warning('%s不是有效文件类型,请选择.doc或.docx类型的文件' % self._extension)
            raise ValueError('文件类型无效')

        self._xls_names = xls_names
        self._dir_path, self._file_name = os.path.split(path)
        self._temp_path = _p_join(self._dir_path, 'vav_temp')
        self._xls_path = _p_join(self._temp_path, self._xls_dir)

    def _create_pdf(self, pdf_path):
        logger.info('开始创建PDF')
        self._word2pdf(self._path, pdf_path)
        assert _p_isfile(pdf_path)
        logger.info('创建PDF成功')

    def _extract_xls(self):
        logger.info('开始提取附件')
        if self._extension == '.doc':
            docx_path = self._base_name + '.docx'
            self._doc2docx(self._path, docx_path)
        else:
            docx_path = self._path
        assert _p_isfile(docx_path)
        zip_path = self._base_name + '.zip'
        if _p_isfile(zip_path):
            os.remove(zip_path)
        os.rename(docx_path, zip_path)
        file_zip = zipfile.ZipFile(zip_path, 'r')
        self._clear_temp_path()
        for file_name in file_zip.namelist():
            bn, ext = _p_splitext(file_name)
            if ext in ('.xls', '.xlsx'):
                file_zip.extract(file_name, self._temp_path)
        assert os.path.isdir(self._xls_path)
        file_zip.close()
        os.rename(zip_path, docx_path)
        xls_names = os.listdir(self._xls_path)
        if len(xls_names) != len(self._xls_names):
            logger.warning('配置中附件数量:%d,文件中附件数量:%d' % (len(self._xls_names), len(xls_names)))
            raise ValueError('文件中的附件数量与配置不符')

        for i, name in enumerate(xls_names):
            old_path = _p_join(self._xls_path, name)
            bn, ext = _p_splitext(name)
            new_name = self._xls_names[i] + ext
            new_path = _p_join(self._xls_path, new_name)
            os.rename(old_path, new_path)
        logger.info('提取附件成功')

    def _clear_temp_path(self):
        if os.path.isdir(self._temp_path):
            shutil.rmtree(self._temp_path)

    def export(self, out_file):
        self._word = word = Dispatch('Word.Application')
        word.Visible = 0
        word.DisplayAlerts = 0
        self._create_pdf(out_file)
        self._extract_xls()
        word.Quit()
        self._word = None
        self._pdf_add_attachment(out_file)
        if g_cfg.get(Config.GENERAL, Config.SAVE_TEMP) != 'True':
            self._clear_temp_path()

    def _doc2docx(self, doc_path, docx_path):
        doc = self._word.Documents.Open(doc_path)
        if os.path.exists(docx_path):
            os.remove(docx_path)
        doc.SaveAs(docx_path, FORMAT_DOCX)
        doc.Close()

    def _word2pdf(self, doc_path, pdf_path):
        doc = self._word.Documents.Open(doc_path, ReadOnly=1)
        self._word.ActiveWindow.View.RevisionsFilter.Markup = MARKDOWN_NONE
        self._word.ActiveWindow.View.RevisionsFilter.View = FINAL_VISION
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        # doc.SaveAs(pdf_path, 17)
        doc.SaveAs(pdf_path, FORMAT_PDF)
        doc.Close()

    def _pdf_add_attachment(self, pdf_path):
        logger.info('开始向PDF插入附件')
        pd_doc = Dispatch('AcroExch.PDDoc')
        assert _p_isfile(pdf_path)
        if pd_doc.Open(pdf_path):
            jso = pd_doc.GetJSObject()
            xls_names = os.listdir(self._xls_path)
            for name in xls_names:
                path = _p_join(self._xls_path, name)
                jso.importDataObject(name, path)
            jso.saveAs(pdf_path)
        logger.info('向PDF插入附件成功')


# if __name__ == '__main__':
#     the_path = 'f:\\python_project\\vat_tool\\RA_12008_TSRS-KA_CASCO_TSRS ATO Function Test Report.doc'
#     xls_names = ['1', '2', '3', '4']
#     out_path = 'f:\\python_project\\vat_tool\\RA_12008_TSRS-KA_CASCO_TSRS ATO Function Test Report.pdf'
#     idoc = RpDoc(the_path, xls_names)
#     idoc.export(out_path)

