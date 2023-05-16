#导入yaml，读取config文件
import numpy as np
import pandas as pd
import yaml
import json
import openpyxl

from public import log_output

with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.load(f, Loader=yaml.FullLoader)

# 定义函数，逐行读取config['user_input']这个Excel路经中，sheet的数据，返回一个字典。其中字段的key为'SN'列，值为该行的数据。第一行的数据为列名
def user_input_deal(file, sheet, index_input, log_print=False):
    user_input_dict = {}
    # 读取指定的 Excel 文件和指定的工作表
    df = pd.read_excel(file, sheet_name=sheet)
    # 遍历每一行数据，去除所有的空格、换行符、制表符
    for index, row in df.iterrows():
        row_dict = dict(row)
        cleaned_row_dict = {}
        for key, value in row_dict.items():
            if isinstance(value, str):
                cleaned_value = value.strip()
                cleaned_row_dict[key] = cleaned_value
            else:
                cleaned_row_dict[key] = value
        if index_input == 'Device_Port_A_index':
            cleaned_row_dict[index_input] = cleaned_row_dict['Rack_Number_A_END'] + '_' + cleaned_row_dict['Position_A_END'] + '_' + cleaned_row_dict['Port_A_END']
        user_input_dict[cleaned_row_dict[index_input]] = cleaned_row_dict

    try:
        user_input_dict.pop(key)
    except KeyError:
        pass

    if log_print:
        log_message = json.dumps(user_input_dict, indent=2, ensure_ascii=False, default=convert_int64)
        log_output('INFO', log_message)
    return user_input_dict

    # 删除key = SN的元素



def convert_int64(obj):
    """
    将对象转换为 int64 类型

    :param obj: 待转换的对象
    :return: 返回转换后的 int64 类型对象，如果无法转换则返回原对象
    """
    if isinstance(obj, np.int64):
        return int(obj)
    return obj


