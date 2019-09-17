import xml.etree.ElementTree as ET
from config.ProjVar import *
from utils.ExcelParse import *
import sys


class ParseXml(object):

    def __init__(self,text):
        self.text = text
        print(self.text)
        try:
            self.root = ET.fromstring(self.text)
        except:
            print(u"Parse response str fail")
            traceback.print_exc()
        self.obj_tag_attr = {}

    def getRootTag(self):
        """获取根元素标签"""
        return self.root.tag

    def getRootChildObj(self):
        """获取根元素下所有子元素对象"""
        childs_obj = []
        for child in self.root:
            childs_obj.append(child)
        return childs_obj

    def getRootChildTagsAttr(self):
        """获取根元素下所有子元素的标签"""
        childs_tag_attr = {}
        for child in self.root:
            childs_tag_attr[child.tag] = child.attrib
        return childs_tag_attr

    def getTagsAttr(self,parent):
        """获取某元素下所有子元素的标签"""
        childs_tag_attr = []
        for child in parent:
            childs_tag_attr.append([child.tag,child.attrib])
        return childs_tag_attr

    def hasChild(self,elemObj):
        """判断标签是否有子元素"""
        elems = []
        for elem in elemObj:
            elems.append(elem)
        if elems:
            return True
        return False


if __name__ == "__main__":
    response = """<?xml version="1.0" ?>
<get_response_B1200_adsl_atm_vcl>
<address node_addr="10.230.9.190"/>
<sequence id="1"/>
<card index="1-1"/>
<table>
<r port="1" vpi="0" vci="35" category="2"/>
<r port="2" vpi="0" vci="35" category="2"/>
<r port="3" vpi="0" vci="35" category="2"/>
</table>
</get_response_B1200_adsl_atm_vcl>"""


    try:
        testExcel = ParseExcel(testCasePath)
        testExcel.setSheetByName("Get ATM VC")
    except:
        print("TestExcel parse fail.")
        traceback.print_exc()
        sys.exit()

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

    # 获取testcase第一行request字段占用的长度
    item_len = len(request_elem_cells_val) + 3
    print(item_len)

    xmlObj = ParseXml(response)
    # print(xmlObj.getRootTag())

    for child in xmlObj.getRootChildObj():
        item = "Response.%s" % child.tag
        if xmlObj.hasChild(child):
            item_new_val = []
            item_new = "%s.child" % (item)
            testExcel.write_cell(1, item_len + 1, item_new)
            item_len += 1
            for elem in xmlObj.getTagsAttr(child):
                item_new_val.append(str(elem[1]))
            print("\n".join(item_new_val))
        else:
            if child.attrib:
                for k,v in child.attrib.items():
                    item_new = "%s.%s" % (item,k)
                    item_val = v
                    testExcel.write_cell(1,item_len+1,item_new)
                    item_len += 1
            else:
                continue





