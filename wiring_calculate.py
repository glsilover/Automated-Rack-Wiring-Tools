import json
import math
import pandas as pd
import yaml
import re
from public import log_output
from user_input_deal import convert_int64


with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.load(f, Loader=yaml.FullLoader)

# 将交换机端口重新组织为prot_dict的key（交换机的key是A端接口），将rack_view与switch_connections的数据进行合并。
def switch_port_collect(rack_view_dict_switch, switch_connections_dict, log_print=False):
    switch_port_dict = {}
    switch_dict = {}
    for instance in rack_view_dict_switch:
        switch_dict[rack_view_dict_switch[instance]['Device_Name']] = rack_view_dict_switch[instance]

    for switch in switch_connections_dict:
        try:
            switch_port_dict[switch_connections_dict[switch]['Device_Port_A_index']] = switch_dict[switch_connections_dict[switch]['Device_Name_A_END']].copy()


            # 将Switch Connections 这个sheet中的字段添加到以端口为key重新组装的dict中
            list = ['Used For', 'Port_A_END', 'IP_A_END', 'PP_Info_X_CONNECT_A_END', 'PP_Port_X_CONNECT_A_END', 'Rack_Number_Z_END', 'Position_Z_END', 'Device_Model_Z_END', 'Device_Name_Z_END', 'Port_Z_END', 'IP_Z_END', 'PP_Info_X_CONNECT_Z_END', 'PP_Port_X_CONNECT_Z_END', 'CABLE1', 'CABLE2', 'MODULE1', 'MODULE2', '长度计算']
            for i in list:
                switch_port_dict[switch_connections_dict[switch]['Device_Port_A_index']][i] = switch_connections_dict[switch][i]

            # TODO 机架位计算逻辑补充
            switch_port_dict[switch_connections_dict[switch]['Device_Port_A_index']]['机架位_A'] = 24
            switch_port_dict[switch_connections_dict[switch]['Device_Port_A_index']]['机架位_Z'] = 24

        except KeyError:
            # 日志红色打印
            log_message = f"注意：{switch_connections_dict[switch]['Device_Name_A_END']} ({config['sheet_switch_connections']} sheet-Device_Name_A_END)交换机名称无法在 {config['sheet_rack_view']} sheet中找到。其真实机柜为【{switch_connections_dict[switch]['Rack_Number_A_END']}】，位置为【{switch_connections_dict[switch]['Position_A_END']}】"
            log_output('ERROR', log_message)



    # josn序列化展示
    if log_print:
        log_message = json.dumps(switch_port_dict, indent=2, ensure_ascii=False, default=convert_int64)
        log_output('INFO', log_message)

    return switch_port_dict  # key名为机柜_U位_端口，value为该端口的所有信息


# 将服务器端口重新组织为prot_dict的key（服务器的key为Z端即交换机的端口）
def server_port_collect(rack_view_dict_server, log_print=False):
    server_port_dict = {}
    server_port_dict_temple = {'10G_Port1': 'eth0', '10G_Port2': 'eth1', '10G_Port3': 'eth2', '10G_Port4': 'eth3', 'IPMI': 'ilo'}

    for instance in rack_view_dict_server:

        for port in server_port_dict_temple:

            # 判断端口是否为NaN，不能用！= ‘’
            if not pd.isna(rack_view_dict_server[instance][port]):
                # 防止dict变量引用
                instance_dict = rack_view_dict_server[instance].copy()

                server_port_dict[instance_dict[port]] = instance_dict

                # 增加server端口号字段
                instance_dict['Port_A_END'] = server_port_dict_temple[port]

                # 提取机架号，拼接为Label A-Side(Cable Label)字段
                match = re.search(r'\d+', instance_dict['Device_Name'])  # \d+ 匹配一个或多个数字
                if match:
                    instance_dict['机架位'] = match.group()  # 使用 group() 方法获取匹配的字符串

                else:
                    instance_dict['机架位'] = '无法获取机架位！'
                    log_message = f"注意：无法从设备名中提取机架位：{instance_dict['Device_Name']}"
                    log_output('ERROR', log_message)

                # 拼接为Label A-Side(Cable Label)字段 = 机柜号-设备名称-U位-端口号
                server_port_dict[instance_dict[port]]['Label A-Side(Cable Label)'] = instance_dict['Rack_Number'] + '-' + instance_dict['Device_Name'] + '-' + instance_dict['Position'] + '-' + server_port_dict_temple[port]
                instance_dict['Port_Z_END'] = instance_dict[port]
            else:
                continue

    # json序列化展示
    if log_print:
        log_message = json.dumps(server_port_dict, indent=2, ensure_ascii=False, default=convert_int64)
        log_output('INFO', log_message)

    return server_port_dict  # key名为机柜_U位_端口，value为该端口的所有信息


def wiring_calculate_server(server_prot_instance, log_print=False):
    server_wiring_dict = {'布线类型': ''}

    tor_rack = server_prot_instance['Port_Z_END'].split('_')[0]
    tor_u = server_prot_instance['Port_Z_END'].split('_')[1]
    tor_port_number = server_prot_instance['Port_Z_END'].split('_')[2]

    compare_dict_a = {'机柜号_A': tor_rack, '机架位_A': '', 'U位_A': tor_u, '品牌_A': '', '设备名称_A': '', '设备型号_A': '', '设备类型_A': '', 'SN_A': ''}

    for key in compare_dict_a:
        server_wiring_dict[key] = compare_dict_a[key]
    # print(wiring_dict)
    server_wiring_dict['端口号_A'] = tor_port_number


    compare_dict_z = {'机柜号_Z': 'Rack_Number', '机架位_Z': '机架位', 'U位_Z': 'Position', '品牌_Z': '品牌', '设备名称_Z': 'Device_Name', '设备型号_Z': 'Device_Model', '设备类型_Z': '设备类型', 'SN_Z': 'SN', '端口号_Z': 'Port_A_END'}
    for key in compare_dict_z:
        server_wiring_dict[key] = server_prot_instance[compare_dict_z[key]]

    # TODO 长度如何获取？数量默认为1
    server_wiring_dict['长度'] = ''
    server_wiring_dict['数量'] = 1

    server_wiring_dict['Label A-Side(Cable Label)'] = server_wiring_dict['机柜号_Z'] + '-' + server_wiring_dict['U位_Z'] + '-' + server_wiring_dict['设备名称_Z'] + '-' + server_wiring_dict['端口号_Z']
    server_wiring_dict['Label Z-Side(Cable Label)'] = server_wiring_dict['机柜号_A'] + '-' + server_wiring_dict['U位_A'] + '-' + server_wiring_dict['设备名称_A'] + '-' + server_wiring_dict['端口号_A']

    server_wiring_dict['Label A-Side(Power1 Label)'] = server_wiring_dict['机柜号_Z'] + '-' + server_wiring_dict['U位_Z'] + '-' + server_wiring_dict['设备名称_Z'] + '-' + 'Power1'
    server_wiring_dict['Label Z-Side(Power1 Label)'] = ''
    server_wiring_dict['Label A-Side(Power2 Label)'] = ''
    server_wiring_dict['Label Z-Side(Power2 Label)'] = ''
    try:
        server_wiring_dict['布线类型'] = server_wiring_dict['设备类型_A'] + ' - TO - ' + server_wiring_dict['设备类型_Z']
    except:
        server_wiring_dict['布线类型'] = 'Spine - TO - Hybrid-core'

    # print(wiring_dict['布线类型'])

    if log_print:
        log_message = json.dumps(server_wiring_dict, indent=2, ensure_ascii=False, default=convert_int64)
        log_output('INFO', log_message)

    return server_wiring_dict  # key名为机柜_U位_端口，value为该端口的所有信息


def wiring_calculate_switch(switch_prot_instance, log_print=False):

    switch_wiring_dict = {'布线类型': ''}

    # 获取A端设备信息
    compare_dict_a = {'机柜号_A': 'Rack_Number', '机架位_A': '机架位_A', 'U位_A': 'Position', '品牌_A': '品牌', '设备名称_A': 'Device_Name', '设备型号_A': 'Device_Model', '设备类型_A': '设备类型', 'SN_A': 'SN', '端口号_A': 'Port_A_END'}
    for key in compare_dict_a:
        switch_wiring_dict[key] = switch_prot_instance[compare_dict_a[key]]

    # 获取Z端设备信息
    # TODO 需要补充出公网设备的逻辑
    if switch_prot_instance['Used For'] == 'UPLINK':
        compare_dict_z = {'机柜号_Z': '', '机架位_Z': '', 'U位_Z': '', '品牌_Z': '', '设备名称_Z': '', '设备型号_Z': '', '设备类型_Z': 'Hybrid-core', 'SN_Z': '', '端口号_Z': ''}
        for key in compare_dict_z:
            switch_wiring_dict[key] = compare_dict_z[key]

    else:
        compare_dict_a = {'机柜号_Z': 'Rack_Number', '机架位_Z': '机架位_Z', 'U位_Z': 'Position', '品牌_Z': '品牌', '设备名称_Z': 'Device_Name', '设备型号_Z': 'Device_Model', '设备类型_Z': '设备类型', 'SN_Z': 'SN', '端口号_Z': 'Port_Z_END'}
        for key in compare_dict_a:
            switch_wiring_dict[key] = switch_prot_instance[compare_dict_a[key]]

    switch_wiring_dict['长度'] = switch_prot_instance['长度计算']
    # TODO 数量默认为1
    switch_wiring_dict['数量'] = 1

    switch_wiring_dict['Label A-Side(Cable Label)'] = switch_wiring_dict['机柜号_Z'] + '-' + switch_wiring_dict['U位_Z'] + '-' + switch_wiring_dict['设备名称_Z'] + '-' + switch_wiring_dict['端口号_Z']
    switch_wiring_dict['Label Z-Side(Cable Label)'] = switch_wiring_dict['机柜号_A'] + '-' + switch_wiring_dict['U位_A'] + '-' + switch_wiring_dict['设备名称_A'] + '-' + switch_wiring_dict['端口号_A']
    switch_wiring_dict['Label A-Side(Power1 Label)'] = switch_prot_instance['Power1']
    switch_wiring_dict['Label A-Side(Power2 Label)'] = switch_prot_instance['Power2']
    switch_wiring_dict['Label Z-Side(Power1 Label)'] = ''
    switch_wiring_dict['Label Z-Side(Power2 Label)'] = ''

    try:
        switch_wiring_dict['布线类型'] = switch_wiring_dict['设备类型_A'] + ' - TO - ' + switch_wiring_dict['设备类型_Z']
    except:
        switch_wiring_dict['布线类型'] = 'Spine - TO - Hybrid-core'

    # print(wiring_dict['布线类型'])

    if log_print:
        log_message = json.dumps(switch_wiring_dict, indent=2, ensure_ascii=False, default=convert_int64)
        log_output('INFO', log_message)

    return switch_wiring_dict  # key名为机柜_U位_端口，value为该端口的所有信息


# 获取品牌
def get_brand(instance_port, log_print=False):
    instance_port['品牌'] = '待补充'
    if log_print:
        log_message = f"品牌：{instance_port['品牌']}"
        log_output('INFO', log_message)

    return instance_port


# 通过交换机机柜、U位，在【rack_view_dict_switch】中反查交换机的设备信息
def get_switch_info(server_wiring_instance, rack_view_dict_switch, log_print=False):
    instance_rack = server_wiring_instance['机柜号_A']
    instance_u = server_wiring_instance['U位_A']

    for instance in rack_view_dict_switch:
        if rack_view_dict_switch[instance]['Rack_Number'] == instance_rack and rack_view_dict_switch[instance]['Position'] == instance_u:

            # TODO 完善机架位信息获取
            server_wiring_instance['机架位_A'] = ''
            server_wiring_instance['设备名称_A'] = rack_view_dict_switch[instance]['Device_Name']
            server_wiring_instance['设备型号_A'] = rack_view_dict_switch[instance]['Device_Model']
            server_wiring_instance['设备类型_A'] = rack_view_dict_switch[instance]['设备类型']
            server_wiring_instance['SN_A'] = rack_view_dict_switch[instance]['SN']
            server_wiring_instance['布线类型'] = server_wiring_instance['设备类型_A'] + ' - TO - ' + server_wiring_instance['设备类型_Z']
            if log_print:
                log_message = json.dumps(server_wiring_instance, indent=2, ensure_ascii=False, default=convert_int64)
                log_output('INFO', log_message)
            return instance
