# coding=UTF-8
__author__ = ''

import os
import sys
import random
import re
from collections import OrderedDict

from Excel import *
from CSVFile import *
from TextSliceMap import *
from ThreadPoolRunner import *
from MyTools import *
import ConfigParser

class PearsonAnalysis():

    def __init__(self, dataSourceFolder=None):
        self.dataSourceFolder = dataSourceFolder
        if self.dataSourceFolder is None:
            Config = ConfigParser.ConfigParser()
            Config.read(os.path.join(sys.path[0] ,"config.ini"))
            self.dataSourceFolder = Config.get('DataSource', 'dataSourceFolder')

        self.dataSourceFolder = self.dataSourceFolder.replace("\\", "/").replace("/","\\")
        self.dataSourceConfigPath = os.path.join(self.dataSourceFolder ,"config.txt")

        self.blnShowExcel = False             # 使用到Excel檔案時，是否顯示Excel視窗
        self.blnForceRestartExcel = False     # 如果開啟Excel失敗，是否關閉所有excel碧常事重啟
        #=============== [已棄用] ===============
        #self.isExcelFirstRowFloat = False     #避免當excel first row作為rowID時，在從excel讀取的過程中，誤將整數型別轉換為浮點數型別
        #======================================

        self.com_pheno_start_colNum = 6        # com_pheno.xlsx中可分析數值的起始欄位編號(只支援欄位編號)
        self.pheno_MappingColId = "id"         # <string型別> 或 <integer型別> 決定使用pheno_v?.csv檔案的哪個欄位，來對應Sln檔案的TextSliceMap.sln.gvid欄位值(可支援欄位名稱或欄位編號)

        self.com_pheno_NotFloatCols = [1,"ID"]    # <list型別> 決定讀取com_phone的Excel檔時哪些欄位強迫變成非float值(可支援欄位名稱或欄位編號)
        self.defaultSubClassList = ["v1", "v2", "v3", "v4"]
        self.defaultSelectProperty = [1,2,3,4] # 預設要分析的屬性ID

        # 找com_pheno.xlsx
        self.re_Com_Pheno = re.compile("com_pheno.*\.[(xlsx)(xls)(csv)]", re.IGNORECASE)
        # 找com_pheno.xlsx內的com_pheno Sheet頁
        self.re_phoneSheetName = re.compile("com_pheno.*", re.IGNORECASE)
        # 找v1~vN的規則
        self.re_SubClass = re.compile("v[1-9]+", re.IGNORECASE)
        # 找sln的規則
        self.re_SlnFile = re.compile(".*\.sln", re.IGNORECASE)

        self.initSelectPropertyMap()
        self.mainClassPathSet = set()
        self.smmaryResultDict = OrderedDict()
        self.mainClassPathSorted = []

        if os.path.exists(self.dataSourceConfigPath):
            print ("Read Config : " + self.dataSourceConfigPath)
            self.replaceSettings(self.dataSourceConfigPath)

    def replaceSettings(self, configPath):
        Config = ConfigParser.ConfigParser()
        Config.read(configPath)
        config_sections = Config._sections
        print "config_sections = ", config_sections

        if "general" in config_sections and config_sections["general"] is not None:
            if config_sections["general"].has_key("blnshowexcel"):
                self.blnShowExcel = bool(config_sections["general"]["blnshowexcel"])
            if config_sections["general"].has_key("blnforcerestartexcel"):
                self.blnForceRestartExcel = bool(config_sections["general"]["blnforcerestartexcel"])
            if config_sections["general"].has_key("com_pheno_start_colnum"):
                self.com_pheno_start_colNum = eval(config_sections["general"]["com_pheno_start_colnum"])
            if config_sections["general"].has_key("pheno_mappingcolid"):
                self.pheno_MappingColId = eval(config_sections["general"]["pheno_mappingcolid"])
            if config_sections["general"].has_key("com_pheno_notfloatcols"):
                self.com_pheno_NotFloatCols = eval(config_sections["general"]["com_pheno_notfloatcols"])
            if config_sections["general"].has_key("defaultsubclasslist"):
                self.defaultSubClassList = eval(config_sections["general"]["defaultsubclasslist"])
            if config_sections["general"].has_key("defaultselectproperty"):
                self.defaultSelectProperty = eval(config_sections["general"]["defaultselectproperty"])

        if "regexSettings" in config_sections and config_sections["regexSettings"] is not None:
            if config_sections["regexSettings"].has_key("re_com_pheno"):
                self.re_Com_Pheno = re.compile(eval(config_sections["regexSettings"]["re_com_pheno"]), re.IGNORECASE)
            if config_sections["regexSettings"].has_key("re_phonesheetname"):
                self.re_phoneSheetName = re.compile(eval(config_sections["regexSettings"]["re_phonesheetname"]), re.IGNORECASE)
            if config_sections["regexSettings"].has_key("re_subclass"):
                self.re_SubClass = re.compile(eval(config_sections["regexSettings"]["re_subclass"]), re.IGNORECASE)
            if config_sections["regexSettings"].has_key("re_slnfile"):
                self.re_SlnFile = re.compile(eval(config_sections["regexSettings"]["re_slnfile"]), re.IGNORECASE)

        if "mainClassDict" in config_sections and config_sections["mainClassDict"] is not None:
            self.mainClassDict = {}
            for (key, value) in config_sections["mainClassDict"].items():
                if key != "__name__":
                    self.mainClassDict[key] = eval(value)

        if "mainClass_subClass" in config_sections and config_sections["mainClass_subClass"] is not None:
            self.mainClass_subClassMap = {}
            for (key, value) in config_sections["mainClass_subClass"].items():
                if key != "__name__":
                    self.mainClass_subClassMap[key] = eval(value)

        print "self.blnShowExcel = ", self.blnShowExcel
        print "self.blnForceRestartExcel = ", self.blnForceRestartExcel
        print "self.com_pheno_start_colNum = ", self.com_pheno_start_colNum
        print "self.pheno_MappingColId = ", self.pheno_MappingColId
        print "self.com_pheno_NotFloatCols = ", self.com_pheno_NotFloatCols
        print "self.defaultSubClassList = ", self.defaultSubClassList
        print "self.defaultSelectProperty = ", self.defaultSelectProperty
        print "self.re_Com_Pheno.pattern = ", self.re_Com_Pheno.pattern
        print "self.re_phoneSheetName.pattern = ", self.re_phoneSheetName.pattern
        print "self.re_SubClass.pattern = ", self.re_SubClass.pattern
        print "self.re_SlnFile.pattern = ", self.re_SlnFile.pattern
        print "self.mainClassDict = ", self.mainClassDict
        print "self.mainClass_subClassMap", self.mainClass_subClassMap

    def stepOneRunConfig(self):
        self.mainClass_subClassMap = {}
        self.mainClass_subClassMap["1k"] = ["v1", "v2", "v3", "v4", "v5"]
        self.mainClass_subClassMap["33k"] = ["v1", "v2", "v3", "v4", "v5"]

    def initSelectPropertyMap(self):
        """
                設定1st, 2nd, 3rd,... 等之中需要挑出做分析的屬性
                self.mainClassDict["對應1st,, 2nd的編號"] = [對應要分析的屬性ID列表]
                """
        self.mainClassDict = {}
        self.mainClassDict["1k"] = ["log10weight", "logelength", "logeLD", "lice"]
        self.mainClassDict["4k"] = ["log10weight", "logelength", "logeLD", "lice"]
        self.mainClassDict["33k"] = [4, 5, 6, 7]

    def getMainClass(self):
        sourcePath = self.dataSourceFolder
        self.listdirToGetMainClassPaths(sourcePath)
        #self.listdirToGetMainClassPaths_old(sourcePath)
        print ("self.mainClassPathSet", self.mainClassPathSet)

    def listdirToGetMainClassPaths(self, sourcePath):
        for lists in os.listdir(sourcePath):
            path = os.path.join(sourcePath, lists).replace("\\","/")
            if os.path.isdir(path):
                self.mainClassPathSet.add(path)


    def listdirToGetMainClassPaths_old(self, sourcePath, level=1):
        if level == 1: print sourcePath
        for lists in os.listdir(sourcePath):
            path = os.path.join(sourcePath, lists).replace("\\","/")
            if path[-4:] == ".sln":
                arrPath = path.split("/")
                arrPath[0] = arrPath[0] + "/"
                newPath = os.path.join(*arrPath[:-2]).replace("\\","/")
                #print ("find sln: " + path)
                #print(newPath)


                self.mainClassPathSet.add(newPath)
            #print '│  '*(level-1)+'│--'+lists
            if os.path.isdir(path):
                self.listdirToGetMainClassPaths_old(path, level+1)

    def getPhenoResult(self, mainClassPath):
        print("getPhenoResult(" + mainClassPath + ")")
        listdir = os.listdir(mainClassPath)

        mainClass = mainClassPath.split("/")[-1]
        print("mainClass = " + mainClass)

        #subClassFolderSet = set()
        com_phenoSet = set()
        for lists in listdir:
            # if self.re_SubClass.match(lists):
            #     subClassFolderSet.add(lists)

            if self.re_Com_Pheno.match(lists):
                com_phenoSet.add(lists)

        subClassFolderList = self.mainClass_subClassMap.get(mainClass, None)

        if subClassFolderList is None:
            subClassFolderList = self.defaultSubClassList

        subClassCount = len(subClassFolderList)
        print("subClassCount = " + str(subClassCount))
        if len(com_phenoSet) != 1:
            raise Exception("com_pheno檔案位於 " + mainClassPath + " 路徑下不只一個，請移除或重新命名非必要項目(搜尋規則：com_pheno.*\.[(xlsx)(xls)(csv)])!")

        print("Get Excel File: " + os.path.join(mainClassPath, list(com_phenoSet)[0]))
        excel = Excel(os.path.join(mainClassPath, list(com_phenoSet)[0]).replace("\\","/"), self.blnShowExcel, self.blnForceRestartExcel)

        #excel.get_sheet(1)
        self.getPhoneSheet(excel)

        dataRecoderCount = excel.sheetRowCounts - 1
        print("dataRecoderCount = " + str(dataRecoderCount))

        randomPickLists = randomPickNGroup(dataRecoderCount, subClassCount)
        print ("randomPickList = ", randomPickLists)
        for randomList in randomPickLists:
            print (len(randomList))

        colNamesList = excel.get_rowData(1)
        colNamesString = list2str(colNamesList)
        #range1 = excel.get_range(1, 1, 1 , excel.sheetColCounts)
        print("list2str=", colNamesString)

        case = 0
        #for randomList in randomPickLists:
        for subClassFolder in subClassFolderList:

            phenoPath =  os.path.join(mainClassPath, "pheno_" + str(subClassFolder) + ".csv").replace("\\","/")
            phenoExceptPath = os.path.join(mainClassPath, "pheno_except_" + str(subClassFolder) + ".csv").replace("\\","/")

            csv = CSVFile(phenoPath, "utf-8")
            csv.writeLine(colNamesString)

            print("len(randomPickLists[case] = ", len(randomPickLists[case]))
            for num in randomPickLists[case]:

                rowData = excel.get_rowData(num+1)
                rowData = list(rowData)

                #if not self.isExcelFirstRowFloat:
                for colNum in self.com_pheno_NotFloatCols:
                    mappingColIdx = None
                    if is_integer(colNum):
                        #print("is_integer")
                        mappingColIdx = colNum - 1
                    else:
                        #print("is not integer")
                        m = -1
                        for colName in colNamesList:
                            #print(colName, colNum)
                            m+=1
                            if str(colNum).strip().lower() == colName.strip().lower():
                                mappingColIdx = m

                    #print("mappingColIdx = " + str(mappingColIdx))

                    try:
                        rowData[mappingColIdx] = int(rowData[mappingColIdx])
                    except Exception:
                        print("Change Excel Value Warning(int(rowData[" + str(mappingColIdx) + "])): " + str(rowData[mappingColIdx]), "You can check self.com_pheno_NotFloatCols, ignore this warning  if expected !!")
                        rowData[mappingColIdx] = str(rowData[mappingColIdx])


                csv.writeLine(list2str(rowData))

            csv = CSVFile(phenoExceptPath, "utf-8")
            csv.writeLine(colNamesString)

            differenceList = []
            i = 0
            for it in randomPickLists:
                if i != case:
                    differenceList.extend(it)
                i+=1

            print("differenceList", differenceList)
            print("differenceList len = " + str(len(differenceList)))


            for num in differenceList:
                rowData = excel.get_rowData(num+1)
                rowData = list(rowData)

                #if not self.isExcelFirstRowFloat:
                for colNum in self.com_pheno_NotFloatCols:
                    mappingColIdx = None
                    if is_integer(colNum):
                        #print("is_integer")
                        mappingColIdx = colNum - 1
                    else:
                        #print("is not integer")
                        m = -1
                        for colName in colNamesList:
                            #print(colName, colNum)
                            m+=1
                            if str(colNum).strip().lower() == colName.strip().lower():
                                mappingColIdx = m

                    #print("mappingColIdx = " + str(mappingColIdx))

                    try:
                        rowData[mappingColIdx] = int(rowData[mappingColIdx])
                    except Exception:
                        print("Change Excel Value Warning(int(rowData[" + str(mappingColIdx) + "])): " + str(rowData[mappingColIdx]), "You can check self.com_pheno_NotFloatCols, ignore this warning  if expected !!")
                        rowData[mappingColIdx] = str(rowData[mappingColIdx])


                csv.writeLine(list2str(rowData))

            case+=1

        excel.close()

    def getAnalysisResult(self, mainClassPath):

        print("getAnalysisResult(" + mainClassPath + ")")
        pearsonResultPath = os.path.join(mainClassPath, "PearsonResult.csv").replace("\\","/")
        listdir = os.listdir(mainClassPath)

        mainClass = mainClassPath.split("/")[-1]
        #mainClassNum = mainClass[0]
        print("mainClass = " + mainClass)
        #print("mainClassNum = " + mainClassNum)
        mainClassDictValue = self.mainClassDict.get(mainClass, None)
        # if mainClassDictValue is None:
        #     mainClassDictValue = self.mainClassDict.get(mainClass, None)

        if mainClassDictValue is None:
            mainClassDictValue = self.defaultSelectProperty

        print("mainClassDictValue = ", mainClassDictValue)
        if mainClassDictValue is None:
            raise Exception("getSlnFiles(" + mainClass + "): 未正確設定欲做為皮爾森分析的屬性，請正確設定 self.mainClassDict 或 self.defaultSelectProperty !!")

        subClassFolderSet = set()
        com_phenoSet = set()
        for lists in listdir:
            if self.re_SubClass.match(lists):
                subClassFolderSet.add(lists)

            if self.re_Com_Pheno.match(lists):
                com_phenoSet.add(lists)

        subClassCount = len(subClassFolderSet)
        print("subClassCount = " + str(subClassCount))
        if len(com_phenoSet) != 1:
            raise Exception("com_pheno檔案位於 " + mainClassPath + " 路徑下不只一個，請移除或重新命名非必要項目(搜尋規則：com_pheno.*\.[(xlsx)(xls)(csv)])!")

        print("Get Excel File: " + os.path.join(mainClassPath, list(com_phenoSet)[0]).replace("\\","/"))
        excel = Excel(os.path.join(mainClassPath, list(com_phenoSet)[0]).replace("\\","/"), self.blnShowExcel, self.blnForceRestartExcel)

        #excel.get_sheet(1)
        self.getPhoneSheet(excel)

        dataRecoderCount = excel.sheetRowCounts - 1
        print("dataRecoderCount = " + str(dataRecoderCount))

        #randomPickLists = randomPickNGroup(dataRecoderCount, subClassCount)

        randomPickDict = {}

        case = 0
        for subClassFolder in subClassFolderSet:
            randomPickList = []
            case+=1
            phone_v_Path = os.path.join(mainClassPath, "pheno_" + str(subClassFolder) + ".csv").replace("\\","/")
            print("Read phone_v_File(" + phone_v_Path + ")")
            csv = CSVFile(phone_v_Path, decoding="utf-8")
            #print(csv.readToString())
            list2D = csv.readTo2DList(",")
            #print list2D
            #print "-------------------------------------"

            mappingColIdx = None
            if is_integer(self.pheno_MappingColId):
                mappingColIdx = self.pheno_MappingColId - 1
            else:
                m = -1
                for colName in list2D[0]:
                    m+=1
                    if colName.strip().lower() == self.pheno_MappingColId.strip().lower():
                        mappingColIdx = m

            #print("mappingColIdx = " + str(mappingColIdx))

            list2D.pop(0)
            #print list2D[0]
            for line in list2D:
                #print(line)
                #print(line[0])'
                if line[mappingColIdx] != "":
                    data = line[mappingColIdx].lower()
                    randomPickList.append(data)

            randomPickDict[subClassFolder] = randomPickList
            pass

        print ("randomPickDict = ", randomPickDict)
        for (k, randomList) in randomPickDict.items():
            print (len(randomList))

        colNamesString = list2str(excel.get_rowData(1))
        #range1 = excel.get_range(1, 1, 1 , excel.sheetColCounts)
        print("list2str=", colNamesString)

        #case = 0

        #subClassesResult = [] #[v1, v2, v3, ...]
        subClassesResult = {} #[v1, v2, v3, ...]
        for subClass in subClassFolderSet:
            result = self.getSlnFiles(excel, mainClassDictValue, mainClassPath, subClass, randomPickDict[subClass])
            #subClassesResult.append(result)
            subClassesResult[subClass] = result
            #case+=1

        print subClassesResult
        excel.close()
        print("++++++++++++++++++++++++++++++++++++++++++++++++")

        csv = CSVFile(pearsonResultPath, "utf-8")
        classCaseNum = 0
        #for slnFilesResult  in subClassesResult:
        for (subClassName, slnFilesResult)  in subClassesResult.items():
            #classRecoderSize = randomPickLists[classCaseNum]
            classRecoderSize = len(randomPickDict[subClassName])


            for lineNum in range(1, classRecoderSize+2):
                strLine = ""
                for sectionNum in range(1, len(mainClassDictValue)+1):
                    #print(slnFilesResult[sectionNum-1][lineNum-1])
                    strLine = strLine + list2str(slnFilesResult[sectionNum-1][lineNum-1]) + ",,"

                csv.writeLine(strLine)

            csv.writeLine("")
            classCaseNum+=1

        #subClassNum=0
        subClassResultDict = OrderedDict()
        for (subClassName, slnFilesResult)  in subClassesResult.items():
            #subClassNum+=1

            slnClassNum=0
            slnResultDict = OrderedDict()
            #subClass_v1
            v_type = slnFilesResult[0][0][2][9:]
            for slnFile in slnFilesResult:
                slnClassNum+=1
                pearsonr_list1 = []
                pearsonr_list2 = []
                #type_9
                select_Type = slnFile[0][5][5:]

                slnFile.pop(0)
                for lineData in slnFile:
                    print("lineData = ", lineData)
                    pearsonr_list1.append(lineData[4])
                    pearsonr_list2.append(lineData[5])

                print("pearsonr_list1(v_type=" + str(v_type) +  ", slnClassNum=" + str(slnClassNum) + ", select_Type=" + str(select_Type) + ") : " , pearsonr_list1)
                print("pearsonr_list2(v_type=" + str(v_type) +  ", slnClassNum=" + str(slnClassNum) + ", select_Type=" + str(select_Type) + ") : " , pearsonr_list2)

                pearson_result = pearsonr(pearsonr_list1, pearsonr_list2)
                pearsonr_list1[:] = []
                pearsonr_list2[:] = []
                print("pearsonr_result(v_type=" + str(v_type) +  ", slnClassNum=" + str(slnClassNum) + ", select_Type=" + str(select_Type) + ") : " + str(pearson_result))
                if str(select_Type) in slnResultDict:
                    raise Exception("select_Type 重複，請檢查 self.mainClassDict 或 self.defaultSelectProperty 是否有重複設定!!")
                slnResultDict[str(select_Type)] = pearson_result

            subClassResultDict[str(v_type)] = slnResultDict

        return subClassResultDict

    def getSlnFiles(self, excel, mainClassDictValue, mainClassPath, subClass, randomList):

        mainClass = mainClassPath.split("/")[-1]
        colNamesList = excel.get_rowData(1)

        tmp_mainClassDictValue = []
        iter = 0
        for value in mainClassDictValue:
            if is_number(value):
                tmp_mainClassDictValue.append(mainClassDictValue[iter] + self.com_pheno_start_colNum - 1)
            else:
                tmp_mainClassDictValue.append(mainClassDictValue[iter])
            iter+=1

        print("new_mainClassDictValue = ", tmp_mainClassDictValue)

        subClassPath =  os.path.join(mainClassPath, subClass).replace("\\","/")
        listdir = os.listdir(subClassPath)

        slnFilesResult = []  #[sln1, sln2, sln3, ...]
        slnCase=0
        for lists in listdir:
            if self.re_SlnFile.match(lists):
                slnCase+=1
                print("find sln file=" + lists + ", slnCase = " +str(slnCase))
                filePath = os.path.join(subClassPath, lists).replace("\\","/")
                f = open(filePath, 'rb')
                data = f.read()
                #print data
                f.close()

                if data.find("\r\n") != -1:
                    arrLine  = data.split("\r\n")
                    arrLine.pop(0)
                elif data.find("\n") != -1:
                    arrLine  = data.split("\n")
                    arrLine.pop(0)
                else:
                    raise Exception("Sln File ( " + filePath + ") 找不到任何換行符號，該檔案可能為空或無正確換行!!")

                recoderDict = {}
                for line in arrLine:
                    if line.strip() != "":
                        givid = str(line[TextSliceMap.sln.gvid]).strip().lower()
                        slnVal1 = float(line[TextSliceMap.sln.var1])
                        recoderDict[givid] = slnVal1

                #print(arrLine)

                blnUseColName = True ; type_col_name = ""
                if is_number(mainClassDictValue[slnCase - 1]):
                    blnUseColName = False
                    type_name = excel.get_cell(1, tmp_mainClassDictValue[slnCase - 1], False, blnUseColName)
                    type_col_name = "type_" + str(tmp_mainClassDictValue[slnCase - 1]) + "_" + type_name
                    #mainClassDictValue[slnCase - 1] = mainClassDictValue[slnCase - 1] + self.com_pheno_start_colNum
                else:
                    type_col_name = "type_" + str(tmp_mainClassDictValue[slnCase - 1])

                pickList = [] #[pickLine1, pickLine2, pickLine3, ...]

                #======================================================若要修改以下內容=========================================================
                # 須同時修改 v_type = slnFilesResult[0][0][2][9:] 以及 select_Type = slnFile[0][5][5:] 重新分配倒數第二個index決定欄位和最後的切片size以取得關鍵字
                #以及pearsonr_list1.append(lineData[4])、pearsonr_list2.append(lineData[5])index決定的欄位
                pickList.append(["pickID", "mainClass_" + str(mainClass), "subClass_" + str(subClass), "slnFileName", "simulation", type_col_name])
                #================================================================================================================================

                for num in randomList:
                    # line = arrLine[num - 1]
                    # givid = int(line[TextSliceMap.sln.gvid])
                    # slnVal1 = float(line[TextSliceMap.sln.var1])

                    try:
                        slnVal1 = recoderDict[str(num)]
                    except:
                        print(" Can't find givid(" + str(num) + ") in slnFile : mainClass=" + str(mainClass) + ", subClass=" + str(subClass) + "\n" +
                              "Please check '" + lists + "' and 'pheno_" + subClass + ".csv' !")
                        raise



                    mappingColNum = None
                    if is_integer(self.pheno_MappingColId):
                        mappingColNum = self.pheno_MappingColId
                    else:
                        #print("is not integer")
                        m = 0
                        for colName in colNamesList:
                            #print(colName, colNum)
                            m+=1
                            if self.pheno_MappingColId.strip().lower() == colName.strip().lower():
                                mappingColNum = m

                    #print("mappingColNum = " + str(mappingColNum))

                    # actualVal = -1
                    # try:
                    actualVal = excel.get_cell(num, tmp_mainClassDictValue[slnCase - 1], True, blnUseColName, mappingColNum)
                    actualVal = float(actualVal)
                    # except:
                    #     print("fail")


                    pickList.append([num , mainClass, subClass, lists, slnVal1, actualVal])

                #print(randomList)
                #print pickList

                slnFilesResult.append(pickList)

        return slnFilesResult

    def getPhoneSheet(self, excel):
        sheetsName = excel.get_sheetsNameList()
        retName = None
        for name in sheetsName:
            if self.re_phoneSheetName.match(name):
                retName = excel.get_sheet(name)
            #     print("sheet name match")
            # else:
            #     print("sheet name not match: " + name)

        if retName is None:
            raise Exception("retName = " + str(retName))

    def exportSummaryResult(self, excelPath):
        print ("excelPath = " + excelPath)
        #excel = Excel(show=self.blnShowExcel, ifFailForceRestart=self.blnForceRestartExcel)
        excel = Excel(show=True, ifFailForceRestart=self.blnForceRestartExcel)
        excel.get_sheet(1)

        lineNum = 0

        for (mainClass_k, mainClass_v) in self.smmaryResultDict.items():
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

    def run(self, step=1):


        self.getMainClass()
        self.mainClassPathSorted = sorted(list(self.mainClassPathSet))

        if step==1:
            print "Start PearsonAnalysis step=1 !!"
            self.stepOneRunConfig()
            for mainClass in self.mainClassPathSet:
                self.getPhenoResult(mainClass)

            # resultFilePath = os.path.join(self.dataSourceFolder, "PearsonAnalysis_Summary.xlsx").replace("\\","/")
            # resultFilePath = os.path.join(self.dataSourceFolder, "PearsonAnalysis_Summary.xlsx").replace("/","\\")
            # print("resultFilePath = " + resultFilePath)
            #
            # for mainClassPath in self.mainClassPathSorted:
            #     mainClass = mainClassPath.split("/")[-1]
            #     self.smmaryResultDict[str(mainClass)] = self.getAnalysisResult(mainClassPath)
            #
            # print self.smmaryResultDict
            #
            # self.exportSummaryResult(resultFilePath)
        elif step ==2:
            print "Start PearsonAnalysis step=2 !!"
            resultFilePath = os.path.join(self.dataSourceFolder, "PearsonAnalysis_Summary.xlsx").replace("\\","/")
            resultFilePath = os.path.join(self.dataSourceFolder, "PearsonAnalysis_Summary.xlsx").replace("/","\\")
            print("resultFilePath = " + resultFilePath)

            for mainClassPath in self.mainClassPathSorted:
                mainClass = mainClassPath.split("/")[-1]
                self.smmaryResultDict[str(mainClass)] = self.getAnalysisResult(mainClassPath)

            print self.smmaryResultDict

            self.exportSummaryResult(resultFilePath)
        else:
            pass


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sourceFolder = sys.argv[1]
    else:
        sourceFolder = None
    print("sourceFolder=", sourceFolder)
    print("TextSliceMap.sln.gvid = ", TextSliceMap.sln.gvid)
    print("TextSliceMap.sln.var1 = ", TextSliceMap.sln.var1)
    pearsonRunner = PearsonAnalysis(sourceFolder)
    pearsonRunner.run(1)

