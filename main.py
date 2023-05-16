import yaml
import json

from public import write_to_excel, write_list_to_excel_new, sort_key
from user_input_deal import user_input_deal
from wiring_calculate import get_brand, server_port_collect, switch_port_collect, wiring_calculate_switch, wiring_calculate_server, get_switch_info

with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.load(f, Loader=yaml.FullLoader)

result_list_switch = []
result_list_server = []
result_list_all = []




# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':

    # 获取所有硬件信息(rack_view)，以SN号为key，value为完整信息
    rack_view_dict = user_input_deal(config['user_input_path'], config['sheet_rack_view'], 'SN', log_print=False)

    # 获取原始的交换机连接信息（switch_connections sheet），以拼接好的A端地址为key("T08_40U_10GE1/0/49")，value为完整信息
    switch_connections_dict = user_input_deal(config['user_input_path'], config['sheet_switch_connections'], 'Device_Port_A_index', log_print=False)

    # 重新组织rack_view_dict，将交换机和服务器分开，分别得到rack_view sheet中对应的硬件配置信息。key = SN, value = 完整信息
    rack_view_dict_switch = {}
    rack_view_dict_server = {}
    for instance_SN in rack_view_dict:
        if rack_view_dict[instance_SN]['设备类型'] != 'server':
            rack_view_dict_switch[instance_SN] = rack_view_dict[instance_SN].copy()
        else:
            rack_view_dict_server[instance_SN] = rack_view_dict[instance_SN].copy()

    # 按接口地址（server为Z端要链接的交换机端口位置，switch为A端交换机端口）为key，重新组织为新的字典。同时在switch_port_dict中合并了
    server_port_dict = server_port_collect(rack_view_dict_server, log_print=False)
    switch_port_dict = switch_port_collect(rack_view_dict_switch, switch_connections_dict, log_print=False)

    # 按最终输出的布线表格式，计算交换机的布线信息
    for switch_prot_instance in switch_port_dict:
        get_brand(switch_port_dict[switch_prot_instance])
        switch_wiring_dict = wiring_calculate_switch(switch_port_dict[switch_prot_instance], log_print=False)
        result_list_switch.append(switch_wiring_dict)

    # 按最终输出的布线表格式，计算服务器的布线信息（注意缺少A端交换机设备信息）
    for server_prot_instance in server_port_dict:
        get_brand(server_port_dict[server_prot_instance])
        server_wiring_dict = wiring_calculate_server(server_port_dict[server_prot_instance], log_print=False)
        result_list_server.append(server_wiring_dict)

    # 在【rack_view_dict_switch】中反查交换机的设备信息
    for server_wiring_instance in result_list_server:
        # print(server_wiring_instance)
        get_switch_info(server_wiring_instance, rack_view_dict_switch, log_print=False)

    # 将交换机和服务器的布线信息合并
    result_list_all = result_list_switch + result_list_server


    # 按照sheet中的顺序排序
    sorted_list = sorted(result_list_all, key=lambda x: list(x.keys())[0])



    # 输出到excel
    write_list_to_excel_new(sorted_list, config['result_file_path'], 'Result')
    print('结果输出成功！')



