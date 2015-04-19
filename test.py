__author__ = 'use'
import os
import sys
import ConfigParser

import random
from MyTools import *
from Excel import *
from collections import OrderedDict


str1 = "  mu                            1                 2.044      0.3489E-02"
str2 = "  giv(givid,1)                  1                0.7786E-01  0.4195E-01"
str3 = "  giv(givid,1)                  3               -0.5365E-01  0.4247E-01"

slice_sln_gvid = slice(32,33)
slice_sln_var1 = slice(49,59)
slice_sln_var2 = slice(62,71)

def randomPickNGroup(size, n):
    splitSize = size/n
    modSize = size%n
    #print(splitSize)
    s = set(range(1,size+1))
    #print(s)
    l = []
    for i in range(1,n+1):
        #print("i=" + str(i))
        if i == n:
            splitSizeTmp = splitSize + modSize
        else:
            splitSizeTmp = splitSize

        subSet = []
        for j in range(1,splitSizeTmp+1):
            #print("j=" + str(j))
            selectNum = random.choice(list(s))
            s.remove(selectNum)
            subSet.append(selectNum)

        l.append(subSet)

    return l

def exportSummaryResult(excelPath, smmaryResultDict):
    print ("excelPath = " + excelPath.replace("/","\\"))
    #excel = Excel(show=self.blnShowExcel, ifFailForceRestart=self.blnForceRestartExcel)
    excel = Excel(show=True, ifFailForceRestart=False)
    excel.get_sheet(1)

    lineNum = 0

    for (mainClass_k, mainClass_v) in smmaryResultDict.items():
        lineNum+=1
        excel.set_cell(lineNum,1, mainClass_k)
        colNum = 1
        for (subClass_k, subClass_v) in mainClass_v.items():
            colNum+=1
            excel.set_cell(lineNum, colNum, subClass_k)
            #for (sln_k, sln_v) in subClass_v:

        subClass_key1 = mainClass_v.keys()[0]
        subClass_v1 = mainClass_v[subClass_key1]
        sln_keys = subClass_v1.keys()

        for sln_key in sln_keys:
            lineNum+=1
            excel.set_cell(lineNum, 1, sln_key)
            colNum = 1
            for (subClass_k, subClass_v) in mainClass_v.items():
                colNum+=1
                print(subClass_v[sln_key])
                excel.set_cell(lineNum, colNum, subClass_v[sln_key])

        lineNum+=1



    #excel.set_cell(1,1,"test")
    excel.save(excelPath)
    excel.close()
    pass

def getMainClass():
    pass

if __name__ == "__main__":
    # randomPickNGroup(513,4)
    # strA = 'test.sln'
    # print (strA[5:])
    # strB = "C:/Users/use/Favorites/Desktop/DataSource/pop1/3k/v3/GBLUP_pop14.sln"
    # arr = strB.split("/")
    # print(arr[:-2])
    # print arr
    #
    # mylist1 = [3,1,2,4]
    # mylist2 = [2,3]
    # diferenceSet = set(mylist1).difference(set(mylist2))
    # print(set(mylist1))
    # print(diferenceSet)
    # print(list2str(mylist1))
    #
    # print strA.find("a")
    #
    # d = OrderedDict()
    # d["2"] = 2
    # d["1"] = 1
    # d["3"] = 3
    #
    # print d
    # key0 = d.keys()[0]
    # print(key0)
    # print(d[key0])

    od = OrderedDict([('3k', OrderedDict([('v1', OrderedDict([('9', -0.08025899416693338), ('10', 0.06723496747784763), ('11', 0.014960169957333944), ('12', 0.21772201335278193)])),
('v2', OrderedDict([('9', 0.055305338255469565), ('10', -0.05979224123433997), ('11', -0.020168329978449622), ('12', 0.2563789965800007)])),
('v3', OrderedDict([('9', 0.01259411165849124), ('10', -0.05217982077346028), ('11', 0.01444028766346314), ('12', 0.28001933988940975)]))])),
('1k', OrderedDict([('v1', OrderedDict([('log10weight', 0.8032670525350902), ('logelength', 0.811748273250521), ('logeLD', 0.6824180985041979), ('lice', 0.7100510061801575)])),
('v2', OrderedDict([('log10weight', 0.8698842651992728), ('logelength', 0.8628378807259401), ('logeLD', 0.6435356710693607), ('lice', 0.6234182476763671)])),
('v3', OrderedDict([('log10weight', 0.8098915656862756), ('logelength', 0.8146208326164085), ('logeLD', 0.6037016771744863), ('lice', 0.6235873574522652)]))])),
('2k', OrderedDict([('v1', OrderedDict([('6', 0.16389343376431487), ('7', 0.7858055762310386), ('8', -0.05625425780991277), ('9', -0.08515789418960364)])),
('v2', OrderedDict([('6', 0.3408558394948879), ('7', 0.8265719147035854), ('8', 0.08544957231932117), ('9', -0.05612099414042342)])),
('v3', OrderedDict([('6', 0.09199531214936787), ('7', 0.8234169307541428), ('8', -0.03334585357372745), ('9', -0.08222156228547413)])),
('v4', OrderedDict([('6', 0.31731448837982096), ('7', 0.820572902292167), ('8', 0.12775620757826675), ('9', -0.009852843134926794)]))]))])

    print od

    #exportSummaryResult("C:/Users/use/Favorites/Desktop/DataSource/PearsonAnalysis_Summary.xlsx", od)

    # print sys.path[0]
    Config = ConfigParser.ConfigParser()
    Config.read(os.path.join(sys.path[0] ,"C:\\Users\\use\\Favorites\\Desktop\\DataSource3\\config.ini"))
    #ret = Config.get('DataSource', 'dataSourceFolder')
    #ret = Config.items("mainClass")
    #ret1k = ret[0]
    ret1 = Config._sections
    #print ret
    #print ret1k
    print ret1
    print ret1["mainClassDict"]["1k"]
    print ret1["general"]["blnshowexcel"]

    # print excel.get_sheetsNameList()

#     print(int(str1[slice_sln_gvid]))
#     print(float(str1[slice_sln_var1]))
#     print(float(str1[slice_sln_var2]))
#     print(int(str2[slice_sln_gvid]))
#     print(float(str2[slice_sln_var1]))
#     print(float(str2[slice_sln_var2]))
#     print(int(str3[slice_sln_gvid]))
#     print(float(str3[slice_sln_var1]))
#     print(float(str3[slice_sln_var2]))