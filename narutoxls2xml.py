#encoding:gbk

import subprocess
import time
import glob
import os

xls_path = "D:\\design_proj\\测试环境配置文件"
cfg_path = "D:\\naruto\\tools\\excelExport"
output = "D:\\output"

config = open(cfg_path + "\\config.ini")

elapse = time.time()

while True:
    item = config.readline()
    if not item:
        break
    item = item.strip()    
    if len(item) == 0 or item[0]=="#":
        continue
    item = item.split(" ")
    params = []
    params.append(xls_path + "\\" + item[2])
    params.append(cfg_path + "\\" + item[1])
    params.append(output + "\\" + item[3])
    result = subprocess.Popen("python xls2xml.py " + ' '.join(params), shell=True, stdout=subprocess.PIPE)
    (out, err) = result.communicate()
    if out:
        print out    
    
elapse = time.time() - elapse
print "elapse:" + "{:4.3f}".format(elapse)
