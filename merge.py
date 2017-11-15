# -*- coding: utf-8 -*-
"""
@author: Fengqian
@email: fengqian.gao@intel.com
"""

import argparse
import collections
import os
import pandas as pd


def get_headers(data):
    return data.columns.tolist()


def get_first_columns(data):
    return data.iloc[:, 0].tolist()


def get_sheet_names(filename):
    names = pd.ExcelFile(filename)
    return names.sheet_names


def get_single_sheet(filename, sheetname):
    return pd.read_excel(filename, sheetname=sheetname)


def read_excel_file(filename):
    if not os.path.isfile(filename):
        print "File {0} is not exist, please check the file path!".format(filename)
        raise Exception

    all_data = collections.OrderedDict()
    all_sheet = get_sheet_names(filename)
    for _name in all_sheet:
        all_data[_name] = get_single_sheet(filename, _name)

    return all_data


def same_headers(all_data={}):
    """
    All the sheet should have the same header line
    :param all_data:
    :return: header line
    """
    init_header = None
    for k, v in all_data.items():
        if init_header is None:
            init_header = get_headers(v)
        elif get_headers(v) == init_header:
                pass
        else:
            print "The header line of all sheets are different, error!"
            return False, None

    return True, init_header


def find_root(all_data, option):
    root_name = None
    root_first_column = None
    for k, v in all_data.items():
        first_column = get_first_columns(v)
        if option in first_column:
            root_name = k
            root_first_column = first_column
            print "Found root sheet!"
            break

    return root_name, root_first_column


def is_child(root_info, child_data):
    pass


def fuzzy_match(all_data, option):
    success, header = same_headers(all_data)
    if not success:
        raise Exception("Some sheet has different header line")
    return find_root(all_data, option)


def get_results(root_data, option_index):
    return root_data.iloc[option_index, :]


def write_to_excel(filename, all_data, option, sheet_name=None):

    root_sheet_name, columns = fuzzy_match(all_data, option)
    root_data = all_data.get(root_sheet_name)
    option_index = columns.index(option)
    result = get_results(root_data, option_index)

    index = 0
    exist_sheet_name = get_sheet_names(filename)
    while True:
        index += 1
        if sheet_name is None:
            sheet_name = "result_{0}".format(index)
            print "Use default sheet_name:{0}".format(sheet_name)
        if sheet_name in exist_sheet_name:
            print "SheetName {0} already exist, change it automatically!".format(sheet_name)
            index += 1
            sheet_name = "result_{0}".format(index)
        else:
            print "Use sheet name {0}".format(sheet_name)
            break

    all_data[sheet_name] = result
    excel_writer = pd.ExcelWriter(filename)
    for k, v in all_data.items():
        if k == sheet_name:
            v.to_excel(excel_writer, k)
        else:
            v.to_excel(excel_writer, k, index=False)
    excel_writer.save()


def merge_excel(filename, option, sheet_name=None):
    all_data = read_excel_file(filename)
    write_to_excel(filename, all_data, option)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Auto merge excel and generate results")
    parser.add_argument('-f', "--file", help="Location for the files.")
    parser.add_argument('-p', "--option", help="Options to put in final result.")
    args = parser.parse_args()
    merge_excel(args.file, args.option)
