# from enum import IntEnum
import requests
import sys
import re
import time
import json
import xlwt


# class Operation(IntEnum):
#     EQUALS = 1
#     NOT_EQUAL = 2
#     LESS_THAN = 3
#     LESS_THAN_OR_EQUAL_TO = 4
#     GREATER_THAN = 5
#     GREATER_THAN_OR_EQUAL_TO = 6
#     CONTAINS = 7
#     NOT_CONTAIN = 8
#     BETWEEN = 9
#     NOT_BETWEEN = 10
#     IS_NULL = 11
#     IS_NOT_NULL = 12
#     IN = 13
#     NOT_IN = 14

# def _len_ch(string):
#     chs = re.findall('[\u4e00-\u9fa5]+', string)
#     total = 0
#     for ch in chs:
#         total += len(ch)
#     return total


class ClearQuest(object):
    def __init__(self, base_url, repository):
        self._base_url = base_url
        self._repository = repository
        self._cq_uid = ''
        self._database = ''
        self._session = requests.Session()

    def login(self, username, password):
        url = self._base_url + '/cqlogin.cq?action=DoLogin'
        login_data = {
            'loginId': username,
            'password': password,
            'repository': self._repository,
            'loadAllRequiredInfo': 'true'
        }
        try:
            r = self._session.post(url, data=login_data)
            # print(r.url)
            db_pattern = r'userdb:\'([^\']*)\''
            self._database = re.search(db_pattern, r.text).group(1)
            # print(self._database)
            uid_pattern = r'cqUid:\'([^\']+)\''
            self._cq_uid = re.search(uid_pattern, r.text).group(1)
            # print(self._cq_uid)
        except Exception as e:
            print(str(e))
            sys.exit(1)

    def logout(self):
        url = self._base_url + '/cqlogin.cq?action=DoLogout'
        data = {'cquid': self._cq_uid}
        self._session.post(url, data=data)

    def _find_record(self, record_id):
        url = self._base_url + '/cqfind.cq'
        payload = {
            'action': 'DoFindRecord',
            'dojo.preventCache': str(int(round(time.time(), 3) * 1000)),
            'recordId': record_id,
            'searchType': 'BY_RECORD_ID'
        }
        headers = {'cquid': self._cq_uid}
        r = self._session.get(url, params=payload, headers=headers)
        # print(r.url)
        id_pattern = r'id:\'([^\']+)\''
        resource_id = re.search(id_pattern, r.text).group(1)
        # print(resource_id)
        return resource_id

    def _get_cq_record_details(self, resource_id):
        url = self._base_url + '/cqartifactdetails.cq'
        payload = {
            'acceptAllTabsData': 'true',
            'action': 'GetCQRecordDetails',
            'dojo.preventCache': str(int(round(time.time(), 3) * 1000)),
            'resourceId': resource_id,
            'state': 'VIEW'
        }
        headers = {'cquid': self._cq_uid}
        r = self._session.get(url, params=payload, headers=headers)
        # print(r.url)
        json_to_load = r.text.replace('for(;;);', '')
        # print(json_to_load)
        cr_fields = json.loads(json_to_load)['fields']
        # print(cr_fields)
        return cr_fields

    def _execute_query(self, query_id):
        url = self._base_url + '/cqqueryresults.cq'
        payload = {
            'action': 'ExecuteQuery',
            'dojo.preventCache': str(int(round(time.time(), 3) * 1000)),
            'format': 'JSON',
            'refresh': 'false',
            'resourceId': 'cq.repo.cq-query:%s@%s' % (query_id, self._database),
            'rowCount': '1500',
            'startIndex': '1'
        }
        headers = {'cquid': self._cq_uid}
        r = self._session.get(url, params=payload, headers=headers)
        # print(r.url)
        json_to_load = r.text.replace('for(;;);', '')
        result = json.loads(json_to_load)['resultSetData']
        # result = json.loads(json_to_load)
        # print(result)
        return result

    @staticmethod
    def _format_query_result_set(result_set):
        col_data = result_set['colData']
        row_data = result_set['rowData']
        col_names = []
        col_fields = []
        for col in col_data:
            col_names.append(col['name'])
            col_fields.append(col['field'])
        report_list = [col_names]
        for row in row_data:
            line = []
            for field in col_fields:
                line.append(row[field])
            report_list.append(line)
        return report_list

    def query_report(self, query_id):
        result_set = self._execute_query(query_id)
        report_list = self._format_query_result_set(result_set)
        book = xlwt.Workbook()
        sheet = book.add_sheet('report', cell_overwrite_ok=True)
        for xcol, line in enumerate(report_list):
            for xrow, cell in enumerate(line):
                sheet.write(xcol, xrow, cell)
        book.save('QueryReport.xls')

    def search_report(self, record_list):
        result_list = []
        for record_id in record_list:
            resource_id = self._find_record(record_id)
            resource = self._get_cq_record_details(resource_id)
            # print(resource)
            resource_dict = {'id': record_id}
            for item in resource:
                if 'CurrentValue' in item.keys() and 'FieldName' in item.keys():
                    resource_dict[item['FieldName']] = item['CurrentValue']
            result_list.append(resource_dict)

        book = xlwt.Workbook()
        sheet = book.add_sheet('report', cell_overwrite_ok=True)
        fields = ['id', 'Headline', 'descriptionn', 'severity', 'CCB_Comments_long']
        for xrow, cell in enumerate(fields):
            sheet.write(0, xrow, cell)
        for i, resource_dict in enumerate(result_list):
            for xrow, field in enumerate(fields):
                cell = resource_dict.get(field, 'NotFound')
                if isinstance(cell, list):
                    cell = self._list2str(cell)
                sheet.write(i+1, xrow, str(cell))
        book.save('SearchReport.xls')

    @staticmethod
    def _list2str(i_list):
        o_string = ''
        for item in i_list:
            o_string += str(item)
        return o_string


if __name__ == '__main__':
    un = 'songqingyang'
    # un = 'yangfan'
    pw = '123456'
    rp = 'iSTP'
    # rp = 'casco_bj'
    ur = 'http://172.19.100.116/cqweb/'
    cq = ClearQuest(ur, rp)
    cq.login(un, pw)
    # cq.query_report('33566296')
    cq.search_report(['1716', '1717', '1720', '1722', '1723', '1724'])
    cq.logout()
