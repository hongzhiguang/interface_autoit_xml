import os

# 工程目录
projPath = os.path.normpath(os.path.dirname(os.path.dirname(__file__)))

# 测试用例路径
testCasePath = os.path.normpath(os.path.join(projPath,"data","测试用例.xlsx"))
testCasePath1 = os.path.normpath(os.path.join(projPath,"data","测试用例1.xlsx"))

# 测试用例
api_sheet_name = u"测试用例"
api_test_case_sheet_col_no = 4
api_command_col_no = 5
api_active_col_no = 6
api_exec_time_col_no = 7
api_res_col_no = 8

# testcasesheet
test_case_item_row_no = 1
test_case_active_col_no = 3


# NBI tool
NBIToolPath = os.path.normpath(os.path.join(projPath,"tool"))
NBIToolPathJar = os.path.normpath(os.path.join(projPath,"tool","OmcaXmlNBI.jar"))

# ip.txt
ipPath = os.path.normpath(os.path.join(projPath,"data","ip.txt"))

# 网管的IP地址
nms_server_ip = "10.156.121.30"

# 鼠标操作控件的坐标
server_ip_locate = (149,43)
connect_locate = (619,43)
disconnect_locate = (865,43)
request_param_locate = (105,125)
response_data_locate = (564,121)
send_button_locate = (400,574)
clear_button_locate = (468,574)

# 循环的次数
cycles = 1


if __name__ == "__main__":
    print(projPath)
    print(testCasePath)
    print(NBIToolPath)