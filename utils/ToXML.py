import xml.etree.ElementTree as ET
import xml.dom.minidom as mindom
import traceback

class ToXML(object):

    def __init__(self,rootTag):
        if rootTag and isinstance(rootTag,str):
            self.rootTag = rootTag
            self.root = ET.Element(self.rootTag)
        else:
            print("The input root tag is illegal.")

    def getRootObj(self):
        """获取根元素对象"""
        return self.root

    def getElemObj(self,parent_obj,elem_tag):
        """输入参数parent为元素对象，根据elem_tag查找对应的element对象，注意只返回一个对象"""
        if isinstance(parent_obj,(str,list)):
            print("The sub element of parent is not a element object.")
            return False
        elem_obj = parent_obj.find(elem_tag)
        return elem_obj

    def getAllElemObj(self,parent_obj,elem_tag):
        """输入参数parent为元素对象，根据elem_tag查找对应的element对象，返回一个对象列表"""
        if isinstance(parent_obj,(str,list)):
            print("The sub element of parent is not a element object.")
            return False
        all_elem_obj = parent_obj.findall(elem_tag)
        return all_elem_obj

    def isExists(self,parent_obj,elem_tag):
        """判断元素是否存在，当不存在的时候findall返回的是None
           parent为对象
        """
        if isinstance(parent_obj,(str,list)):
            print("The sub element of parent is not a element object.")
            return False
        ret_elem = parent_obj.findall("./")
        for child in ret_elem:
            if child.tag == elem_tag:
                return True
            else:
                continue
        return False

    def setRootSubElement(self,child_tag,attr_lis=None,attr_val_lis=None,content=None):
        """设置root根元素下的子元素，以及子元素的属性、文本内容"""
        try:
            if attr_lis and attr_val_lis:
                attr_d = dict(zip(attr_lis,attr_val_lis))
                sub_elem_obj = ET.SubElement(self.root, child_tag, attrib=attr_d)
            else:
                sub_elem_obj = ET.SubElement(self.root, child_tag)
            if content:
                sub_elem_obj.text = content
            return sub_elem_obj
        except:
            traceback.print_exc()
            return

    def setSubElement(self,parent_obj,child_tag,attr_lis=None,attr_val_lis=None,content=None):
        """设置parent元素对象下的子元素，以及子元素的属性、文本内容"""
        if isinstance(parent_obj,(str,list)):
            print("The sub element of parent is not a element object.")
            return
        try:
            if attr_lis and attr_val_lis:
                attr_d = dict(zip(attr_lis, attr_val_lis))
                sub_elem_obj = ET.SubElement(parent_obj,child_tag, attrib=attr_d)
            else:
                sub_elem_obj = ET.SubElement(parent_obj, child_tag)
            if content:
                sub_elem_obj.text = content
            return sub_elem_obj
        except:
            traceback.print_exc()
            return

    def getXMLString(self):
        """根据节点返回格式化的xml字符串"""
        try:
            xml_str = ET.tostring(self.root)
            reparsed = mindom.parseString(xml_str)
            format_str = reparsed.toprettyxml()
            return format_str
        except:
            traceback.print_exc()
            return ''


if __name__ == "__main__":
    toXML = ToXML("login_command")
    root_obj = toXML.getRootObj()

    toXML.setRootSubElement("sequence",["id"],["1"],"we are family.")

    access_elem = toXML.setRootSubElement("access",["id"],["11"])
    if toXML.isExists(root_obj,"access"):
        toXML.setSubElement(access_elem, "r",['port', 'vpi', 'vci', 'category'],['1', '2', '33', '1'])
    else:
        print("fail")

    toXML.setRootSubElement("user", ["name", "password"], ["root", "public123"])

    # access_elem = toXML.getElemObj(root_obj,"access")
    # toXML.setSubElement(access_elem, "acc1")

    # acc1_elem = toXML.getElemObj(access_elem,"acc1")
    # toXML.setSubElement(acc1_elem, "acc2")

    print(toXML.getXMLString())


