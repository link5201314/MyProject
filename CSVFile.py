# coding=UTF-8
__author__ = 'Isaac1'
import io

class CSVFile():
    def __init__(self, filepath, encoding=None, decoding=None, append = False):
        self.path = filepath
        self.file = None
        self.decoding = decoding
        self.encoding = encoding
        self.append = append

    def read(self):
        self.file = open(self.path, 'rb')
        data = self.file.read()
        self.file.close()
        return data

    def readToString(self):
        return self.read().decode(self.decoding)

    def readTo2DList(self, delimiter='\t', rmLastEmpty = True):
        udata = self.readToString()
        data2D = []
        lines = udata.split('\n')

        if rmLastEmpty and lines[len(lines)-1].strip() == "":
            lines.pop()

        for line in lines:
            rows = line.split(delimiter)
            dataRow = []
            for field in rows:
                dataRow.append(field)

            data2D.append(dataRow)

        return data2D

    def write(self, strBuffer):
        if self.append:
            self.file = io.open(self.path, 'a', encoding=self.encoding)
        else:
            self.file = io.open(self.path, 'w', encoding=self.encoding)
        self.file.write(unicode(strBuffer))
        self.file.close()
        self.append = True

    def writeLine(self, strBuffer):
        if self.append:
            self.file = io.open(self.path, 'a', encoding=self.encoding)
        else:
            self.file = io.open(self.path, 'w', encoding=self.encoding)
        self.file.write(unicode(strBuffer))
        self.file.write(u"\n")
        self.file.close()
        self.append = True

    def writeFrom2DList(self, list, delimiter=','):
        if self.append:
            self.file = io.open(self.path, 'a', encoding=self.encoding)
        else:
            self.file = io.open(self.path, 'w', encoding=self.encoding)
        for line in list:
            self.file.write(unicode(line))
            self.file.write(u"\n")

        self.file.close()
        self.append = True

