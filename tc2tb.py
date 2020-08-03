# !/usr/bin/python3

import os
import xlwt
# import time
import re
# from mtranslate import translate
from btrans import baidu_translate
from win32com.client import Dispatch
from copy import deepcopy
from utils import logger, change_path_to_word_style

CELL_END = '\r\7'
MIN_ROWS = 5
PRESUPPOSED = '预设条件'


def _write_sheet_row(sheet, x_row, data_list):
    for x_col in range(len(data_list)):
        _write_sheet_cell(sheet, x_row, x_col, data_list[x_col])


def _write_sheet_cell(sheet, x_row, x_col, data):
    sheet.write(x_row, x_col, data)


class TCDoc(object):
    HEAD_LINE = ['Tag', 'Sources', 'Description', 'Actions', 'Expected Result']

    def __init__(self, path):
        self._path = path = change_path_to_word_style(path)
        if not os.path.isabs(path):
            logger.warning('%s不是有效路径' % path)
            raise ValueError('路径无效')
        self._base_name, self._extension = os.path.splitext(path)
        # self._xls_path = self._base_name + '.xls'
        if self._extension not in ('.doc', '.docx'):
            logger.warning('%s不是有效文件类型,请选择.doc或.docx类型的文件' % self._extension)
            raise ValueError('文件类型无效')

    def export_to_excel(self, out_path):
        logger.info('开始输出')
        ms_word = Dispatch('Word.Application')
        ms_word.Visible = 0
        ms_word.DisplayAlerts = 0
        document = ms_word.Documents.Open(self._path)
        tables = document.Tables
        tc_list = []
        for table in tables:
            try:
                my_table = Table(table)
            except Exception as e:
                err_info = str(e)
                logger.error(err_info)
            else:
                if not my_table.is_test_case():
                    continue
                tc_list.append(my_table.get_test_case())
        self._create_excel(out_path, tc_list)
        document.Close()
        ms_word.Quit()

    def translate(self):
        logger.info('开始翻译')
        ms_word = Dispatch('Word.Application')
        ms_word.Visible = 0
        ms_word.DisplayAlerts = 0
        document = ms_word.Documents.Open(self._path)
        tables = document.Tables
        for table in tables:
            try:
                my_table = Table(table)
            except Exception as e:
                err_info = str(e)
                logger.error(err_info)
            else:
                if not my_table.is_test_case():
                    continue
                my_table.translate_table()
        # document.SaveAs2(out_path)
        document.Save()
        document.Close()
        logger.info('翻译完成')
        ms_word.Quit()

    def _create_excel(self, out_path, tc_list):
        book = xlwt.Workbook()
        sheet = book.add_sheet('TC', cell_overwrite_ok=True)
        writing_row_id = 0
        _write_sheet_row(sheet, writing_row_id, self.HEAD_LINE)
        writing_row_id += 1
        tag_col_id = 0
        sources_col_id = 1
        des_col_id = 2
        pre_title_col_id = 3
        presupposed_col_id = 4
        index_col_id = 2
        actions_col_id = 3
        result_col_id = 4
        for test_case in tc_list:
            _write_sheet_cell(sheet, writing_row_id, tag_col_id, test_case['tag'])
            _write_sheet_cell(sheet, writing_row_id, sources_col_id, test_case['sources'])
            _write_sheet_cell(sheet, writing_row_id, des_col_id, test_case['des'])
            _write_sheet_cell(sheet, writing_row_id, pre_title_col_id, PRESUPPOSED)
            _write_sheet_cell(sheet, writing_row_id, presupposed_col_id, test_case['presupposed'])
            writing_row_id += 1
            for step in test_case['steps']:
                _write_sheet_cell(sheet, writing_row_id, index_col_id, step['index'])
                _write_sheet_cell(sheet, writing_row_id, actions_col_id, step['actions'])
                _write_sheet_cell(sheet, writing_row_id, result_col_id, step['result'])
                writing_row_id += 1
        book.save(out_path)
        logger.info('保存 %s' % out_path)


class Table(object):

    def __init__(self, table):
        self._table = table
        self._reformed_table = self._reform_table(table)

    def is_test_case(self):
        if len(self._reformed_table) < MIN_ROWS:
            return False
        first_cell = self._reformed_table[0][0]
        if first_cell.find('Test Case Description') == -1 and first_cell.find('测试用例描述') == -1:
            return False
        return True

    def table(self):
        return self._reformed_table

    def get_test_case(self):
        assert self.is_test_case()
        tc_description = self._get_tc_description()
        logger.info('读取 %s 成功' % tc_description['tag'])
        test_case = deepcopy(tc_description)
        test_case['steps'] = self._get_tc_steps()
        test_case['presupposed'] = self._get_tc_presupposed()
        return test_case

    def translate_table(self):
        tc_description = self._get_tc_description()
        logger.info('读取 %s 成功' % tc_description['tag'])
        # self._translate_presupposed()
        self._tran_description()
        self._tran_presupposed()
        self._tran_actions_and_results()

    def _get_row_id(self, *args):
        row_id = -1
        for i in range(len(self._reformed_table)):
            row = self._reformed_table[i]
            for kw in args:
                if row[0].find(kw) != -1:
                    row_id = i
                    break
            else:
                continue
            break
        return row_id

    @staticmethod
    def _get_col_id(line_values, *args):
        col_id = -1
        for i, val in enumerate(line_values):
            for kw in args:
                if val.find(kw) != -1:
                    col_id = i
                    break
            else:
                continue
            break
        return col_id

    @staticmethod
    def _tran_des_cell(cell):
        text = str(cell).strip(CELL_END)
        in_list = text.split('\r')
        check_line = in_list[2]
        if not check_line.startswith('['):
            return
        ch_text = in_list[1]
        en_text = baidu_translate(ch_text)
        text = ch_text + '\r' + en_text
        in_list[1] = text
        out_text = '\r'.join(in_list)
        cell_range = cell.Range
        cell_range.Delete()
        cell_range.InsertAfter(out_text)

    @staticmethod
    def _tran_act_cell(cell):
        text = str(cell).strip(CELL_END)
        in_list = text.split('\r')
        if len(in_list) != 1:
            return
        en_text = baidu_translate(text)
        out_text = text + '\r' + en_text
        cell_range = cell.Range
        cell_range.Delete()
        cell_range.InsertAfter(out_text)

    @staticmethod
    def _tran_cell(cell):
        text = str(cell).strip(CELL_END)
        in_list = text.split('\r')
        re_ch = re.compile('[\u4e00-\u9fa5]+')
        for line in in_list:
            if not re.search(re_ch, line):
                return
        for i, line in enumerate(in_list):
            en_text = baidu_translate(line)
            in_list[i] += '\r' + en_text
        out_text = '\r'.join(in_list)
        cell_range = cell.Range
        cell_range.Delete()
        cell_range.InsertAfter(out_text)

    def _tran_description(self):
        des_row = self._get_row_id('描述', 'Description') + 1
        if des_row == -1:
            return
        des_cell = self._table.Cell(des_row, 2)
        self._tran_des_cell(des_cell)

    def _tran_presupposed(self):
        pre_row = self._get_row_id('条件', 'condition') + 1
        if pre_row == -1:
            return
        cell = self._table.Cell(pre_row, 2)
        self._tran_cell(cell)

    def _tran_actions_and_results(self):
        headline_row = self._get_row_id('Step', '步骤') + 1
        if headline_row < 0:
            return
        headline = self._reformed_table[headline_row - 1]
        input_col = self._get_col_id(headline, '输入', 'Input') + 1
        actions_col = self._get_col_id(headline, '步骤描述', 'Actions') + 1
        result_col = self._get_col_id(headline, '结果', 'Result') + 1
        for row in range(headline_row + 1, len(self._reformed_table)):
            row_name_cell = self._table.Cell(row, 1)
            row_name = str(row_name_cell)
            if row_name.find('通过') != -1 or row_name.find('Pass') != -1:
                break
            if input_col > 0:
                input_cell = self._table.Cell(row, input_col)
                self._tran_cell(input_cell)
            if actions_col > 0:
                actions_cell = self._table.Cell(row, actions_col)
                self._tran_cell(actions_cell)
            if result_col > 0:
                result_cell = self._table.Cell(row, result_col)
                self._tran_cell(result_cell)

    def _get_tc_description(self):
        description_str = self._reformed_table[0][1]
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
        presupposed_str = self._reformed_table[2][1]
        return presupposed_str

    def _get_tc_steps(self):
        headline_row_id = self._get_row_id('Step', '步骤')
        if headline_row_id < 0:
            raise LookupError('找不到步骤行')
        headline = self._reformed_table[headline_row_id]
        actions_col = self._get_col_id(headline, '步骤', 'Actions')
        result_col = self._get_col_id(headline, '结果', 'Result')
        steps = []
        index = 1
        for row in self._reformed_table[headline_row_id + 1:]:
            if len(row) != len(headline):
                continue
            actions = row[actions_col]
            result = row[result_col]
            if actions != '' and result != '':
                steps.append({'index': index, 'actions': actions, 'result': result})
                index += 1
        return steps

    @staticmethod
    def _reform_table(table):
        reformed_table = []
        rows = table.Rows
        for row in rows:
            cell_list = []
            cells = row.Cells
            for cell in cells:
                cell_value = str(cell)
                cell_value = cell_value.strip(CELL_END)
                cell_list.append(cell_value)
            reformed_table.append(cell_list)
        return reformed_table
