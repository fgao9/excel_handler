# -*- coding: utf-8 -*-
"""
@author: Fengqian
@email: fengqian.gao@intel.com
"""
import collections
import pandas
import random


TEST_SIZE = 5
ATTRIBUTE_SIZE = 10


#STATIC_COLUMNS1 = {"stock ticker": ["财务打分", "技术分析", "总得分"]}
STATIC_COLUMNS1 = {"stock ticker": ["Finance Score", "Technical Analysis", "Total Score"]}


def generate_attributes_list():
    data_list = []
    for i in range(1, ATTRIBUTE_SIZE + 1):
        data_list.append("attribute_{0}".format(i))
    return {"stock ticker": data_list}


def generate_root():
    ret = collections.OrderedDict()
    ret["stock ticker"] = STATIC_COLUMNS1["stock ticker"]
    start_index = 60001
    for i in range(0, TEST_SIZE):
        random1 = random.randint(0,30)
        random2 = random.randint(0,30)
        ret[start_index] = [random1, random2, (random1 + random2)]
        start_index += 1
    return ret


def generate_child(end_attribute, value_dict):
    ret = collections.OrderedDict()
    ret["stock ticker"] = generate_attributes_list().get("stock ticker")
    ret["stock ticker"].append(end_attribute)
    start_index = 60001
    for i in range(0, TEST_SIZE):
        data_list = []
        for i in range(0, ATTRIBUTE_SIZE):
            data_list.append(random.randint(0,100))
        data_list.append(value_dict.get(start_index))
        ret[start_index] = data_list
        start_index += 1
    return ret


def split_root(dataframe, root_attribute):
    slice_data = dataframe.iloc[root_attribute, 1:]
    ret = slice_data.to_dict()
    return ret


def generate_data():
    all_data = []
    root = pandas.DataFrame(generate_root())
    all_data.append(root)
    overall_index = STATIC_COLUMNS1["stock ticker"]
    overall_index = overall_index[:-1]
    for i in overall_index:
        temp = generate_child(i, split_root(root, overall_index.index(i)))
        all_data.append(pandas.DataFrame(temp))
    return all_data


def write_to_excel(filename, data, sheetname=None):
    excel_writer = pandas.ExcelWriter(filename)
    for i in range(0, len(data)):
        data[i].to_excel(excel_writer, "Sheet{0}".format(i), index=False)
    excel_writer.save()

if __name__ == '__main__':
    all_data = generate_data()
    write_to_excel("./data/test.xls", all_data)
    

