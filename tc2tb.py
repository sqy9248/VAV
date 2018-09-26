# !/usr/bin/python3

import os
import xlwt
from win32com.client import Dispatch
from copy import deepcopy
from utils import logger, change_path_to_word_style

CELL_END = '\r\7'
MIN_ROWS = 5
PRESUPPOSED = '预设条件'


def _write_sheet_row(sheet, xrow, data_list):
    for xcol in range(len(data_list)):
        _write_sheet_cell(sheet, xrow, xcol, data_list[xcol])


def _write_sheet_cell(sheet, xrow, xcol, data):
    sheet.write(xrow, xcol, data)


class TCDoc(object):
    HEAD_LINE = ['Tag', 'Sources', 'Description', 'Actions', 'Expected Result']

    def __init__(self, path):
        self._path = path = change_path_to_word_style(path)
        if not os.path.isabs(path):
            logger.warning('%s不是有效路径' % path)
            raise ValueError('路径无效')
        self._base_name, self._extension = os.path.splitext(path)
        self._xls_path = self._base_name + '.xls'
        if self._extension not in ('.doc', '.docx'):
            logger.warning('%s不是有效文件类型,请选择.doc或.docx类型的文件' % self._extension)
            raise ValueError('文件类型无效')

        self._tc_list = []

    def export(self, out_path):
        logger.info('开始输出')
        ms_word = Dispatch('Word.Application')
        ms_word.Visible = 0
        ms_word.DisplayAlerts = 0
        document = ms_word.Documents.Open(self._path)
        tables = document.Tables
        self._set_tc_list(tables)
        self._to_excel(out_path)
        document.Close()
        ms_word.Quit()

    def _set_tc_list(self, tables):
        for table in tables:
            try:
                my_table = Table(table)
            except Exception as e:
                err_info = str(e)
                logger.error(err_info)
            else:
                if not my_table.is_test_case():
                    continue
                self._tc_list.append(my_table.get_test_case())

    def _to_excel(self, out_path):
        book = xlwt.Workbook()
        sheet = book.add_sheet('TC', cell_overwrite_ok=True)
        writing_xrow = 0
        _write_sheet_row(sheet, writing_xrow, self.HEAD_LINE)
        writing_xrow += 1
        tag_xcol = 0
        sources_xcol = 1
        des_xcol = 2
        pre_title_xcol = 3
        presupposed_xcol = 4
        index_xcol = 2
        actions_xcol = 3
        result_xcol = 4
        for test_case in self._tc_list:
            _write_sheet_cell(sheet, writing_xrow, tag_xcol, test_case['tag'])
            _write_sheet_cell(sheet, writing_xrow, sources_xcol, test_case['sources'])
            _write_sheet_cell(sheet, writing_xrow, des_xcol, test_case['des'])
            _write_sheet_cell(sheet, writing_xrow, pre_title_xcol, PRESUPPOSED)
            _write_sheet_cell(sheet, writing_xrow, presupposed_xcol, test_case['presupposed'])
            writing_xrow += 1
            for step in test_case['steps']:
                _write_sheet_cell(sheet, writing_xrow, index_xcol, step['index'])
                _write_sheet_cell(sheet, writing_xrow, actions_xcol, step['actions'])
                _write_sheet_cell(sheet, writing_xrow, result_xcol, step['result'])
                writing_xrow += 1
        book.save(out_path)
        logger.info('保存 %s' % out_path)


class Table(object):

    def __init__(self, table):
        self._table = []
        self._read_table(table)

    def is_test_case(self):
        if len(self._table) < MIN_ROWS:
            return False
        first_cell = self._table[0][0]
        if first_cell.find('Test Case Description') == -1 and first_cell.find('测试用例描述') == -1:
            return False
        return True

    def table(self):
        return self._table

    def get_test_case(self):
        assert self.is_test_case()
        tc_description = self._get_tc_description()
        logger.info('读取 %s 成功' % tc_description['tag'])
        test_case = deepcopy(tc_description)
        test_case['steps'] = self._get_tc_steps()
        test_case['presupposed'] = self._get_tc_presupposed()
        return test_case

    def _get_tc_description(self):
        description_str = self._table[0][1]
        description_list = description_str.split('\r')
        tag = description_list[0]
        description = description_list[1]
        sources = []
        for item in description_list[2:]:
            if item.find('Source') != -1:
                tag_begin_pos = 9
                tag_end_pos = -1
                sources.append(item[tag_begin_pos: tag_end_pos])
        description_dict = {'tag': tag, 'des': description, 'sources': sources}
        return description_dict

    def _get_tc_presupposed(self):
        presupposed_str = self._table[2][1]
        return presupposed_str

    def _get_tc_steps(self):
        headline_xrow = self._get_headline_xrow()
        headline = self._table[headline_xrow]
        head_xcol = self._get_head_xcol(headline)
        steps = []
        index = 1
        for row in self._table[headline_xrow + 1:]:
            if len(row) != len(headline):
                continue
            actions = row[head_xcol['actions']]
            result = row[head_xcol['result']]
            if actions != '' and result != '':
                steps.append({'index': index, 'actions': actions, 'result': result})
                index += 1
        return steps

    @staticmethod
    def _get_head_xcol(headline):
        actions_xcol = 0
        result_xcol = 0
        i = 0
        for head in headline:
            if head.find('描述') != -1 or head.find('Actions') != -1 or head.find('Description') != -1:
                actions_xcol = i
            elif head.find('预期结果') != -1 or head.find('Expected Result') != -1:
                result_xcol = i
            else:
                pass
            i += 1
        head_xcol = {'actions': actions_xcol, 'result': result_xcol}
        return head_xcol

    def _get_headline_xrow(self):
        headline_xrow = 0
        i = 0
        while True:
            if i >= len(self._table):
                break
            row = self._table[i]
            if row[0].find('Step') != -1 or row[0].find('步骤') != -1:
                headline_xrow = i
                break
            i += 1
        return headline_xrow

    def _read_table(self, table):
        rows = table.Rows
        for row in rows:
            cell_list = []
            cells = row.Cells
            for cell in cells:
                cell_value = str(cell)
                cell_value = cell_value.strip(CELL_END)
                cell_list.append(cell_value)
            self._table.append(cell_list)


# if __name__ == '__main__':
#     temp_path = 'f:\\python_project\\vat_tool\\RA16022_ITCS_ATP_CASCO_4130 Smart Core软件确认测试用例-（20180204).doc'
#     doc = TCDoc(temp_path)
#     doc.export()
