from utils.ToXML import *
from utils.KeyBoard import *
from action.MouseControlTool import *
from utils.XMLParse import *
import re


def main():

    # 运行restfultool
    # 第一种方式：不能在pycharm里面运行，可以打包或者在cmd里面运行Python文件
    # autoit.run("cmd.exe")
    # autoit.send("cd ..\\tool{ENTER}")
    # autoit.send("java -jar OmcaXmlNBI.jar{ENTER}")
    # 第二种方式
    autoit.run("runJar.exe")
    time.sleep(5)

    # 连接网管
    moveServerIPEntry()
    selectAll()
    copyByText(nms_server_ip)
    paste()
    setText("")
    connectNMSServer()

    # 判断是否连接成功，如果没有连接成功，退出执行
    moveToResponse()
    selectAll()
    copyByKey()
    getRes = getText()
    if getRes.decode("utf-8"):
        print("Connect NMS server fail.")
        sys.exit()

    # 读取ip.txt文件中网元的IP
    try:
        with open(ipPath,"r") as fp:
            ips = []
            for ip in fp:
                ips.append(ip.strip())
            if len(ips) > 1 or len(ips) == 0:
                print("Only support one ip.")
                sys.exit()
            else:
                testIP = ips[0]
    except:
        print("Read ip.txt fail.")
        traceback.print_exc()

    # 解析测试用例.xlsx文件
    try:
        testExcel = ParseExcel(testCasePath)
        testExcel.setSheetByName(api_sheet_name)
    except:
        print("TestExcel parse fail.")
        traceback.print_exc()
        sys.exit()


    try:
        testExcel.getRows(api_active_col_no)[1:]
    except:
        print("Get api active rows fail.")
        traceback.print_exc()
        sys.exit()

    for api_cell in testExcel.getRows(api_active_col_no)[1:]:
        if api_cell.value == "Y":
            APITestCaseSheet = testExcel.get_cell_value(api_cell.row,api_test_case_sheet_col_no)
            APIcommand = testExcel.get_cell_value(api_cell.row,api_command_col_no)
            APIexecTime = testExcel.get_cell_value(api_cell.row,api_exec_time_col_no)
            APIexecRes = testExcel.get_cell_value(api_cell.row,api_res_col_no)
            print(APITestCaseSheet,APIcommand)

            # 切换到指定的sheet，执行testcase
            testExcel.setSheetByName(APITestCaseSheet)

            # 将Python对象转成xml格式的请求
            toXML = ToXML(APIcommand)
            root_obj = toXML.getRootObj()

            # 如果APIcommand为login_command或者logout_command，那么不需要设置
            if APIcommand == "login_command" or APIcommand == "logout_command":
                pass
            else:
                toXML.setRootSubElement("address",["node_addr"],[testIP])

            # 根据第一行的字段拼接成请求的格式
            # 将所有带有Request的字段放进列表中，这些带有Request的字段就是XML请求中元素
            start_request_time = time.time()
            request_elem_cells = []
            request_elem_cells_val = []

            try:
                testExcel.getCols(test_case_item_row_no)[3:]
            except:
                print("Sheetname must same as TestCaseSheet.")
                print("Get testcase item columns fail.")
                traceback.print_exc()
            for item_cell in testExcel.getCols(test_case_item_row_no)[3:]:
                if item_cell.value and item_cell.value.split(".")[0] == "Request":
                    request_elem_cells.append(item_cell)
                    request_elem_cells_val.append(item_cell.value)

            # 获取testcase第一行字段占用的长度,决定response插入位置
            item_len = len(request_elem_cells_val) + 3

            # 判断字段中是否包含有"sequence"和"access"，如果没有则设置默认值
            seq_flag = False
            acc_flag = False
            for elem_item in request_elem_cells_val:
                if "sequence" in elem_item:
                    seq_flag = True
                    continue
                if "access" in elem_item:
                    acc_flag = True
                    continue
            if not seq_flag:
                toXML.setRootSubElement("sequence",["id"],["1"])
            if not acc_flag:
                toXML.setRootSubElement("access",["id"],["11"])

            # 提取属性的正则表达式
            pattern = re.compile(r'([\w-]+)=\"(.*?)\"', re.M)

            # 对元素进行属性赋值
            try:
                testExcel.getRows(test_case_active_col_no)[1:]
            except:
                print("Get testcase active rows fail.")
                traceback.print_exc()
            for test_cell in testExcel.getRows(test_case_active_col_no)[1:]:
                if test_cell.value == "Y":
                    caseRowNo = test_cell.row
                    for item in request_elem_cells:
                        attr_lis = []
                        attr_val_lis = []
                        try:
                            if len(item.value.split(".")[1:])==1:
                                # 根元素下的子元素不带属性的情况
                                try:
                                    tag = item.value.split(".")[1:][0]
                                    toXML.setRootSubElement(tag)
                                except:
                                    print(u"根元素下的子元素不带属性的情况")
                                    traceback.print_exc()
                            elif len(item.value.split(".")[1:])==2 and "n" in item.value.split(".")[1:]:
                                # 根元素下的子元素带多个属性的情况
                                try:
                                    tag = item.value.split(".")[1:][0]
                                    attrs_val = testExcel.get_cell_value(caseRowNo,item.column)
                                    res = re.findall(pattern,attrs_val)
                                    for t in res:
                                        attr_lis.append(t[0])
                                        attr_val_lis.append(t[1])
                                    toXML.setRootSubElement(tag,attr_lis,attr_val_lis)
                                except:
                                    print(u"根元素下的子元素带多个属性的情况")
                                    traceback.print_exc()
                            elif  len(item.value.split(".")[1:])==3 and "child" in item.value.split(".")[1:]:
                                # 根元素下的子元素，并且该子元素又带有子元素（可以有多个属性）的情况
                                try:
                                    parent_tag = item.value.split(".")[1:][0]
                                    if not toXML.isExists(root_obj,parent_tag):
                                        parent_obj = toXML.setRootSubElement(parent_tag)
                                    else:
                                        parent_obj = toXML.getElemObj(root_obj, parent_tag)
                                    child_tag = item.value.split(".")[1:][2]
                                    child_d = testExcel.get_cell_value(caseRowNo,item.column)
                                    for d in child_d.split("\n"):
                                        res = re.findall(pattern, d)
                                        for t in res:
                                            attr_lis.append(t[0])
                                            attr_val_lis.append(t[1])
                                        toXML.setSubElement(parent_obj,child_tag, attr_lis, attr_val_lis)
                                        attr_lis = []
                                        attr_val_lis = []
                                except:
                                    print(u"根元素下的子元素，并且该子元素又带有子元素的情况")
                                    traceback.print_exc()
                            else:
                                # 根元素下的子元素只带有一个属性
                                try:
                                    tag,attr = item.value.split(".")[1:]
                                    attr_val = testExcel.get_cell_value(caseRowNo, item.column)
                                    attr_lis.append(attr)
                                    attr_val_lis.append(str(attr_val))
                                    toXML.setRootSubElement(tag,attr_lis,attr_val_lis)
                                except:
                                    print(u"根元素下的子元素只带有一个属性的情况")
                                    traceback.print_exc()
                        except:
                            print("Get xml string fail.")
                            traceback.print_exc()
                    # 完成对XML对象的拼接
                    end_request_time = time.time()
                    print("拼接XML对象需要花费的时长为%d" % (end_request_time - start_request_time))

                    # 将请求的内容粘贴到输入框中，并发送请求
                    request_content = toXML.getXMLString()
                    moveToRequest()
                    selectAll()
                    copyByText(request_content)
                    getContent = getText()
                    paste()

                    # 发送请求
                    send()

                    # 获取response的内容
                    n = 30
                    moveToResponse()
                    responseContent = getText()
                    while getContent == responseContent:
                        if n == 0:
                            responseContent = b""
                            break
                        selectAll()
                        copyByKey()
                        responseContent = getText()
                        time.sleep(1)
                        n -= 1

                    if not responseContent:
                        continue
                    resp = responseContent.decode("utf-8")
                    # print(resp)
                    xmlObj = ParseXml(resp)
                    # print(xmlObj.getRootTag())

                    # 对返回的response XML格式字符串进行解析并写入表格
                    start = time.time()
                    for child in xmlObj.getRootChildObj():
                        item = "Response.%s" % child.tag
                        if xmlObj.hasChild(child):
                            item_new_val = []
                            for elem in xmlObj.getTagsAttr(child):
                                elem_str = []
                                n = len(elem[1])
                                for k,v in elem[1].items():
                                    if n == 1:
                                        elem_str.append('%s="%s"' % (k, v))
                                        break
                                    else:
                                        elem_str.append('%s="%s",' % (k,v))
                                        n -= 1
                                item_new_val.append("".join(elem_str))
                            item_new = "%s.child.%s" % (item,elem[0])
                            testExcel.write_cell(1, item_len + 1, item_new)
                            testExcel.write_cell(caseRowNo,item_len + 1,"\n".join(item_new_val))
                            item_len += 1
                        else:
                            if child.attrib:
                                if len(child.attrib) == 1:
                                    for k, v in child.attrib.items():
                                        item_new = "%s.%s" % (item, k)
                                        item_val = v
                                        testExcel.write_cell(1, item_len + 1, item_new)
                                        testExcel.write_cell(caseRowNo,item_len + 1,item_val)
                                        item_len += 1
                                else:
                                    item_val = []
                                    n = len(child.attrib)
                                    for k, v in child.attrib.items():
                                        if n == 1:
                                            item_val.append('%s="%s"'% (k, v))
                                            break
                                        else:
                                            item_val.append('%s="%s",' % (k,v))
                                            n -= 1
                                    testExcel.write_cell(1, item_len + 1, item)
                                    testExcel.write_cell(caseRowNo,item_len + 1,"\n".join(item_val))
                                    item_len += 1
                            else:
                                continue
                    testExcel.save()
                    end = time.time()
                    print(u"返回的response解析并写入Excel花费的时长为%d" % (end-start))

            # 执行完一个api，需要set sheet回到测试用例
            testExcel.setSheetByName(api_sheet_name)


if __name__ == "__main__":
    main()
