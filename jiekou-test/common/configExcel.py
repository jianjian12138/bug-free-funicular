#!/usr/bin/env python
#-*- coding:utf-8 -*-
import xlrd


class ExcelUtil(object):

    def __init__(self, excelPath, sheetName):
        self.data = xlrd.open_workbook(excelPath)           #打开excle表格，参数是文件路径
        self.table = self.data.sheet_by_name(sheetName)     #通过名字获取
        # get titles
        self.row = self.table.row_values(0)                  #获取第一行值 参数是第几行
        self.col = self.table.col_values(0)                  #获取第一列值 参数是第几列
        # get rows number
        self.rowNum = self.table.nrows                       #获取总行数
        # get columns number
        self.colNum = self.table.ncols                       #获取总列数
        # the current column
        self.curRowNo = 1                                    #当前行数为1

    def hasNext(self):
        #判断excle是否为空
        if self.rowNum == 0 or self.rowNum <= self.curRowNo :
            return False
        else:
            return True

    def next(self):
        r = []         #创建空列表r
        while self.hasNext():      #excel不为空时
            s = {}                 #创建空字典s
            col = self.table.row_values(self.curRowNo)     #获取第二行数的列的值
            i = self.colNum            #总列数赋值为i
            for x in range(i):         #x遍历总列数i
                s[self.row[x]] = col[x]       #对应行列的值
            r.append(s)                       #把字典的值加入列表r中
            self.curRowNo += 1                #行数加1
        return r

    def rowlist(self, i):
        # 按行读取存为list,去除空字符
        rowlist = self.table.row_values(i)
        n = rowlist.count("")
        for i in range(n):
            rowlist.remove(u'')
        return rowlist

    def readasdict(self):
        d = {}
        col = self.table.col_values(0)
        nrows = self.table.nrows
        for i in range(nrows):
            val = self.rowlist(i)[1:]
            if len(val) == 1:
                 d[col[i]] = val[0]
            else:
                d[col[i]] = val
        return d

    def readaslitbyrow(self, i, j):
        l = []
        s = self.rowlist(i)
        e = self.rowlist(j)
        for i in range(1, len(s)):
            d = {}
            d[s[0]] = s[i]
            d[e[0]] = e[i]
            l.append(d)
        return l

if __name__=="__main__":
    import config
    import os
    excelPath = os.path.join(config.data_path, "testdata")
    excel = ExcelUtil(excelPath, title_line=True)
    #print excel.readaslitbyrow(2,3)
    #print excel.rowlist(3)
    print excel.data
    # u = excel.readasdict()
    # s = u["locator"]
    # for i in s:
    #     print i
