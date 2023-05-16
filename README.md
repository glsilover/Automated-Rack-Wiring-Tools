


# Tips
1. `Switch Connections`sheet中的`Device_Name_A_END`字段交换机设备名称必须和`Rack_View`sheet中的`Device_Name`字段一致，否则无法匹配到交换机设备信息。
2. `Device_Name`Server的设备名称必须严格遵守server-xx的格式，其中xx为机架位数字。否则无法匹配到机架位信息。
3. `Rack_Number`无论在任何时候，Rack_Number都必须是两位数，例如`T01` `H05-MGMT`，而不能写为`T1` `H5-MGMT`。