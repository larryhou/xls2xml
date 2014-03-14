#encoding:gbk

import sys
import os
import re
from xlrd import open_workbook, cellname
import lxml.etree as etree
import ntpath

decoding="gb2312"
encoding="utf-8"

def test(condition, msg, errorCode = 110):
    if not condition:
        if errorCode == 0:
            stderr(msg)
        else:
            stderr("ERR" + str(errorCode) + ":", msg)
        sys.exit(errorCode)
        
def stderr(*argv):
    print >> sys.stderr, ' '.join(str(x) for x in argv)

class XMLCfg:    
    def __init__(self, node = None, name = None, root = None):
        self.map = {}
        self.node = node
        self.name = name
        self.root = root

def createNode(node, imap):
    data = etree.tostring(node, encoding=encoding)
    for key in imap.keys():
        data = data.replace("{" + key.encode(encoding) + "}", imap[key].encode(encoding))        
    node = etree.fromstring(data)
    return node

def excel2xml(sheet, data):
    if data.root == None:
        data.root = "root"
    root = etree.Element(data.root)
    for r in range(1, sheet.nrows):
        imap={}
        for c in range(sheet.ncols):
            value = sheet.cell(r,c).value
            if not isinstance(value, unicode):
                value = str(value).decode(decoding)
            
            key = data.map.get(c)
            imap[key] = value
            
        if data.node != None:
            item = createNode(data.node.__copy__(), imap)
            root.append(item)
        else:
            item = etree.SubElement(root, "item")
            for key in imap.keys():
                item.set(key, imap[key])
        
    return root

def parseXMLCfg(url):
    data = etree.parse(url)
    test(data != None, "����ģ�����ʧ��")
    
    node = data.find("exportNode").getchildren()[0]
    test(node != None, "����ģ��[exportNode]Ϊ��")
    
    name = data.find("sheetName").text
    test(name != None, "����ģ��[sheetName]Ϊ��")
    
    root = data.find("exportRoot").getchildren()[0].tag
    return XMLCfg(node, name, root)

def convert(xls_path):    
    book = open_workbook(xls_path)
    sheet = book.sheet_by_name(cfg.name)
    
    for r in range(1):
        for c in range(sheet.ncols):
            value = sheet.cell(r, c).value
            if not isinstance(value, unicode):
                value = str(value).decode(decoding)
            cfg.map[c] = value

    root = excel2xml(sheet, cfg)
    result = etree.tostring(root, encoding=encoding, pretty_print=True)
    return result

if __name__=="__main__":
    test(len(sys.argv) == 4, "usage: xls2xml xls_path cfg_path output", 0)
    
    xls_path = sys.argv[1]
    cfg_path = sys.argv[2]
    output = sys.argv[3]

    test(os.path.exists(xls_path), "EXCEL�ļ�[" + xls_path + "]������", 404)
    test(re.search(r'\.xlsx?$', xls_path.lower()) != None, "[" + xls_path + "]����EXCEL�ļ�")
    
    test(os.path.exists(cfg_path), "����ģ��[" + cfg_path + "]������", 404)
    test(cfg_path[-4:].lower() == ".xml", "����ģ��[" + cfg_path + "]����XML�ļ�")

    cfg = parseXMLCfg(cfg_path)
    
    test(output != None, "XML���Ŀ¼Ϊ��");
    if output[-4:].lower() != ".xml":
        if not os.path.exists(output):
            os.mkdir(output)
        output = output + "\\" + cfg.name + ".xml"

    result = "<?xml version='1.0' encoding='utf-8'?>\n" + convert(xls_path)
    print "output: " + output
    
    f = open(output, 'wb')
    f.write(result)
    f.close()

