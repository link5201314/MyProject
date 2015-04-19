# coding=UTF-8
__author__ = 'Isaac1'
import win32com.client
#import win32con
#import win32gui
import pythoncom
import string
import types

from MyTools import *

class Excel:
    def __init__(self, filename=None, show=True, ifFailForceRestart = False):
        pythoncom.CoInitialize()
        self.xlApp = win32com.client.dynamic.Dispatch('Excel.Application')
        self.xlApp.DisplayAlerts = False
        self.xlApp.SheetsInNewWorkbook = 1
        self.sheetsCount = -1
        self.sheetRowCounts = -1
        self.sheetColCounts = -1
        self.ifFailForceRestart = ifFailForceRestart

        if filename:
            self.filename = filename
            try:
                self.xlBook = self.xlApp.Workbooks.Open(filename)
            except:

                if self.ifFailForceRestart:
                    print("Excel檔案開啟失敗(" + filename + ")：嘗試重啟excel進程")
                else:
                    print("Excel檔案開啟失敗(" + filename + ")：\r\n請確保有安裝Microsoft Excel應用程式，以及確保欲讀取的檔案並未正在使用中!!")
                    raise

                execCommand("taskkill /F /IM excel.exe")
                sleep(5)
                try:
                    self.xlBook = self.xlApp.Workbooks.Open(filename)
                except:
                    print("Excel檔案開啟失敗(" + filename + ")：\r\n請確保有安裝Microsoft Excel應用程式，以及確保欲讀取的檔案並未正在使用中!!")
                    raise

        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

        if show:
            #print("show")
            self.show()
        else:
            #print("not show")
            self.hide()
        self.sheetsCount = self.xlBook.Sheets.Count

    def __del__(self):
        self.sheetsCount = self.sheetRowCounts = self.sheetColCounts = self.filename = self.xlBook = self.xlSheet = None
        del self.xlApp
        pythoncom.CoUninitialize()

    def get_sheet(self,ws_identification):
        self.xlSheet = self.xlBook.Worksheets(ws_identification)
        self.setSheetColRow()
        return self.xlSheet.Name

    def get_sheetsCount(self):
        return self.xlBook.Sheets.Count

    def get_sheetsNameList(self):
        l = []
        for idx in range(1, self.get_sheetsCount()+1):
            l.append(self.xlBook.Sheets.Item(idx).Name)

        return l

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)

    def quit(self):
        self.xlApp.Quit()

    def show(self):
        self.xlApp.Visible = True

    def hide(self):
        #sleep(5)
        self.xlApp.Visible = False
        #hwnd = win32gui.FindWindow("XLMAIN", None)
        #XLMAIN
        #EXCEL
        #print("hwnd=" + str(hwnd))
        #win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

    def col2num(self,col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num

    def num2col(self,n):
        s = ""
        while n != 0:
            s += chr((n - 1) % 26 + 65)
            n //= 27
        return s[::-1]

    def setSheetColRow(self):
        Cells = self.xlSheet.Columns.Item(1).Cells
        lastCell = Cells.Find("*", Cells.Item(1), -4163, 1, 1, 2)
        if lastCell is not None:
            self.sheetRowCounts = lastCell.Row
        else:
            self.sheetRowCounts = 0

        #print(self.sheetRowCounts)


        self.sheetColCounts = 1
        if self.sheetRowCounts > 0:
            while self.xlSheet.UsedRange.Cells(1, self.sheetColCounts).Value is not None:
                self.sheetColCounts += 1

        self.sheetColCounts -= 1

        #print(self.sheetColCounts)

    def getColNum(self, colId):
        if is_number(colId):
            return colId
        else:
            return self.col2num(colId)

    def getRowByName(self, rowName, rowNamdColNum = 1):

        range = self.get_range(1, rowNamdColNum, self.sheetRowCounts, rowNamdColNum)
        i = 1
        #print(range)
        for rowId in range:
            if rowId[0] is None:
                print("rowId[0] is None !!")
                raise
            if is_number(rowId[0]) and is_number(rowName):
                if (float(rowId[0]) == float(rowName)):
                    return i
            else:
                if (str(rowId[0]).strip().lower() == str(rowName).strip().lower()):
                    return i

            i+=1

    def getColByName(self, colName):
        cols = self.get_range(1, 1, 1, self.sheetColCounts)[0]
        i = 1
        #print(cols)
        for colId in cols:
            if is_number(colId) and is_number(colName):
                if (float(colId) == float(colName)):
                    return i
            else:
                if (str(colId).strip().lower() == str(colName).strip().lower()):
                    return i

            i+=1

    def readTo2DList(self):
        data2D = []

        for i in range(1, self.sheetRowCounts+1):
            dataRow = []
            for j in range(1, self.sheetColCounts+1):
                dataRow.append(self.get_cell(i,j))

            data2D.append(dataRow)

        return data2D

    def get_cell(self, row, col, useRowName=False, useColName=False, rowNamdColNum=1):
        "get value of one cell"
        #sht = self.xlBook.Worksheets(sheet)
        #print(self.xlSheet)
        if useRowName:
            rowNum = self.getRowByName(row, rowNamdColNum)
        else:
            rowNum = row

        if useColName:
            colNum = self.getColByName(col)
        else:
            colNum = self.getColNum(col)

        #print("rowNum=",rowNum)
        #print("colNum=",colNum)
        try:
            return self.xlSheet.Cells(rowNum, colNum).Value
        except Exception:
            print("fail")

    def set_cell(self, row, col, value, useRowName=False, useColName=False):
        "set value of one cell"
        #sht = self.xlBook.Worksheets(sheet)
        if useRowName:
            rowNum = self.getRowByName(row)
        else:
            rowNum = row

        if useColName:
            colNum = self.getColByName(col)
        else:
            colNum = self.getColNum(col)

        self.xlSheet.Cells(row, col).Value = value

    def get_range(self, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        #sht = self.xlBook.Worksheets(sheet)
        range = self.xlSheet.Range(self.xlSheet.Cells(row1, col1), self.xlSheet.Cells(row2,col2))
        #range.NumberFormat = '@'
        return range.Value

    def set_range(self, leftCol, topRow, data):
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        #sht = self.xlBook.Worksheets(sheet)
        self.xlSheet.Range(self.xlSheet.Cells(topRow, leftCol), self.xlSheet.Cells(bottomRow,rightCol)).Value = data

    def get_rowData(self, row, useRowName=False, rowNamdColNum=1):
        if useRowName:
            rowNum = self.getRowByName(row, rowNamdColNum)
        else:
            rowNum = row

        return self.get_range(rowNum, 1, rowNum, self.sheetColCounts)[0]

