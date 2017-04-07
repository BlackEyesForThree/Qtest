#!/usr/bin/python3
# -*- coding:utf-8 -*-
import requests, xlrd, pymysql, time, sys
from xlutils import copy

def readExcel(file_path):
    try:
        book = xlrd.open_workbook(file_path)
    except Exception as e:
        print('路径不在或者excel不正确', e)
        return e
    else:
        sheet = book.sheet_by_index(0)
        rows = sheet.nrows
        case_list = []
        for i in range(rows):
            if i != 0:
                case_list.append(sheet.row_values(i))
        interfaceTest(case_list, file_path)


def interfaceTest(case_list, file_path):
    res_flags = []
    request_urls = []
    responses = []
    for case in case_list:
        try:
            product = case[0]
            case_id = case[1]
            interface_name = case[2]
            case_detail = case[3]
            method = case[4]
            url = case[5]
            param = case[6]
            res_check = case[7]
            tester = case[10]
        except Exception as e:
            return '测试用例格式不正确！%s' % e
        if param == '':
            new_url = url
            request_urls.append(new_url)
        else:
            new_url = url + '?' + param
            request_urls.append(new_url)
        if method.upper() == 'GET':
            print(new_url)  # 此处打印访问url
            results = requests.get(new_url).text
            # print(results)  # 此处打印返回报文,已注释
            responses.append(results)
            res = readRes(results, res_check)
        else:
            results = requests.post(new_url).text
            print(new_url)  # 此处打印访问url
            responses.append(results)
            res = readRes(results, res_check)
        if 'pass' in res:
            res_flags.append('pass')
        else:
            res_flags.append('fail')

    copy_excel(file_path, res_flags, request_urls, responses)


def readRes(res, res_check):

    for s in res_check:
        if s in res:
            pass
        else:
            return '错误，返回参数和预期结果不一致' + str(s)
    return 'pass'





def copy_excel(file_path, res_flags, request_urls, responses):
    book = xlrd.open_workbook(file_path)
    new_book = copy.copy(book)
    sheet = new_book.get_sheet(0)
    i = 1
    for request_url, response, flag in zip(request_urls, responses, res_flags):
        sheet.write(i, 8, u'%s' % request_url)
        sheet.write(i, 9, u'%s' % response)
        sheet.write(i, 11, u'%s' % flag)
        i += 1
    new_book.save('%s测试结果.xls' % time.strftime('%Y%m%d%H%M%S'))


if __name__ == '__main__':
    try:
        filename = '/Users/suxin/Downloads/MyJob/TestCase.xls'
    except IndexError as e:
        print('Please enter a correct testcase! \n e.x: python gkk.py test_case.xls')
    else:
        readExcel(filename)
    print('Done!')

