# -*- coding: cp936 -*-
import os
import sys
import time
import datetime
import win32com.client

# for Py2Exe Support 
##from win32com.client import gencache
##gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 6)
#gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 7)

import xml.etree.ElementTree
import xml.etree.cElementTree as et

from wsbase import *

currPath = os.getcwd()
month  = time.strftime("%m",time.localtime()) 
day  = time.strftime("%d",time.localtime())
year = time.strftime("%Y",time.localtime())

tree = et.ElementTree(file="config.xml")

projectName = tree.find('*/projectName').text
projectVersion = tree.find('*/version').text

covTools = tree.find('*/analyzeVer').text
contact = tree.find('*/contact').text

company = tree.find('*/testCompany').text
testContact = tree.find('*/testContact').text

label = unicode("分析报告", "gb2312")

def calc_total_summary():
    totalLOC = 0
    outstandingCount  = 0
    with open('componentMetrics.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            #The second field is codeLineCount, so the array index should be 1
            #
            totalLOC = totalLOC + int(data[1].strip())
            outstandingCount  = outstandingCount  + int(data[5].strip())
    #print totalLOC
    return totalLOC,outstandingCount

def calc_level_num():
    totalHigh = 0
    totalMedium = 0
    totalLow = 0
    with open('defectList.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            if (data[8].strip() != "Dismissed"):
                if (data[6].strip() == "High"):
                    totalHigh = totalHigh + 1
                elif (data[6].strip() == "Medium"):
                    totalMedium = totalMedium + 1
                elif (data[6].strip() == "Low"):
                    totalLow = totalLow + 1
    return totalHigh, totalMedium, totalLow

def calc_acceptance_num():
    totalFP = 0
    totalInten = 0
    totalBug = 0
    totalUnclassified = 0
    with open('defectList.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            if (data[9].strip() == "Unclassified"):
                    totalUnclassified = totalUnclassified + 1
            elif (data[9].strip() == "Bug") :
                    totalBug = totalBug + 1
            elif (data[9].strip() == "Intentional") :
                    totalInten = totalInten + 1
            elif (data[9].strip() == "False Positive") :
                    totalFP = totalFP + 1
    return totalUnclassified, totalBug, totalInten, totalFP


def calc_category_num():
    high = dict()
    medium = dict()
    low = dict()
    with open('defectList.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            if (data[8].strip() != "Dismissed"):
                if (data[6].strip() == "High"):
                    if(data[5].strip() in high):
                        high[data[5].strip()] = high[data[5].strip()] + 1
                    else :
                        high[data[5].strip()] = 1
                elif (data[6].strip() == "Medium"):
                    if (data[5].strip() in medium):
                        medium[data[5].strip()] = medium[data[5].strip()] + 1
                    else :
                        medium[data[5].strip()] = 1
                elif (data[6].strip() == "Low"):
                    if (data[5].strip() in low):
                        low[data[5].strip()] = low[data[5].strip()] + 1
                    else :
                        low[data[5].strip()] = 1
    return high, medium, low

def calc_component_metrics():
    all = dict()
    #all[component]=[high,medium,low,loc]

    with open('componentMetrics.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            all[data[0].strip()] = [0,0,0,data[1].strip(),0]
            
    with open('defectList.csv') as fn:
        for each_line in fn:
            data = each_line.split(',')
            if(data[8].strip() != "Dismissed"):
                if(data[6].strip() == "High"):
                    all[data[7].strip()][0] = all[data[7].strip()][0]+1
                elif (data[6].strip() == "Medium"):
                    all[data[7].strip()][1] = all[data[7].strip()][1]+1
                elif (data[6].strip() == "Low"):
                    all[data[7].strip()][2] = all[data[7].strip()][2]+1

    for key, value in all.items():
        all_defect_num = value[0] + value[1] + value[2]
        if(all_defect_num != 0):
            value [4] = "%.3f" % float(float(all_defect_num * 1000) / float(value[3])) 
    return all

##def RGB( Red, Green, Blue ):
##    HEX = (hex( Red )[-2:],hex( Green )[-2:],hex( Blue )[-2:] )
##    return HEX

def write_component_density(doc):
    all = calc_component_metrics()
    print all

    tmpList = []
    for key, value in all.items():
        (cmap, cname ) = key.split('.')
        tmpList.append([cname, value[4], value[0], value[1], value[2], value[3]])

    sortHighList = sorted([(t[1],t[0],t[2],t[3],t[4], t[5]) for t in tmpList], reverse = True) 
    print sortHighList
    rows = len(sortHighList)
    print "rows = " + str(rows)
    
    #aRange = doc.Range(doc.Bookmarks("componentDensity").Range.Start, doc.Bookmarks("componentDensity").Range.End)
    #aRange = doc.Range(0,0)
    aRange = doc.Bookmarks("componentDensity").Range
    tabDensity = doc.Tables.Add(Range = doc.Bookmarks("componentDensity").Range, NumRows= rows + 1, NumColumns= 3)
    #i = app.CentimetersToPoints(2.5) equals to  i = 71
    tabDensity.Columns.PreferredWidth = 60
    print str(tabDensity)
    tabDensity.Style = "普通表格"

    print tabDensity.Columns.PreferredWidth
    #print tabDensity.Style

    
    tabDensity.Cell(Row=1,Column=1).Range.InsertAfter(Text='模块名称')
    tabDensity.Cell(Row=1,Column=2).Range.InsertAfter(Text='代码行')
    tabDensity.Cell(Row=1,Column=3).Range.InsertAfter(Text="缺陷密度")
    tabDensity.Rows(1).Height = 40
    tabDensity.Rows(1).HeightRule = 1 
    
    intCount = 1
    for item in sortHighList :
        intCount = intCount + 1
        tabDensity.Cell(Row=intCount, Column=1).Range.InsertAfter(Text=item[1])
        tabDensity.Cell(Row=intCount, Column=2).Range.InsertAfter(Text=str(item[5]))
        tabDensity.Cell(Row=intCount, Column=3).Range.InsertAfter(Text=str(item[0]))
        tabDensity.Rows(intCount).Height = 50
        tabDensity.Rows(intCount).HeightRule = 1 


    sp =  doc.Bookmarks("componentMetric").Range.Select()
    
    objShape = doc.Shapes.AddChart(Type = 58).Chart

    objShape.ChartData.Workbook.Application.WindowState = -4140

    chartWorkSheet = objShape.ChartData.Workbook.Worksheets(1)

    chartWorkSheet.ListObjects("表1").Resize(chartWorkSheet.Range("A1:D"+str(rows+1)))
    
    #objShape.SetSourceData(Source = chartWorkSheet.Range("A1:D"+str(rows+1)))
    chartWorkSheet.Range("B1").FormulaR1C1 = "高风险"
    chartWorkSheet.Range("C1").FormulaR1C1 = "中风险"
    chartWorkSheet.Range("D1").FormulaR1C1 = "低风险"
    
    intCount = rows + 1
    for item in sortHighList :
        chartWorkSheet.Range("A"+ str(intCount)).FormulaR1C1 = item[1]
        chartWorkSheet.Range("B"+ str(intCount)).FormulaR1C1 = item[2]
        chartWorkSheet.Range("C"+ str(intCount)).FormulaR1C1 = item[3]
        chartWorkSheet.Range("D"+ str(intCount)).FormulaR1C1 = item[4]
        intCount = intCount - 1

    objShape.SeriesCollection(1).Format.Fill.Visible = True
    objShape.SeriesCollection(1).Format.Fill.ForeColor.RGB = -16777024
    objShape.SeriesCollection(1).Format.Fill.Transparency = 0
    objShape.SeriesCollection(1).Format.Fill.Solid

    objShape.SeriesCollection(2).Format.Fill.Visible = True
    objShape.SeriesCollection(2).Format.Fill.ForeColor.RGB = -16727809
    objShape.SeriesCollection(2).Format.Fill.Transparency = 0
    objShape.SeriesCollection(2).Format.Fill.Solid

    objShape.SeriesCollection(3).Format.Fill.Visible = True
    objShape.SeriesCollection(3).Format.Fill.ForeColor.RGB = -2894893
    objShape.SeriesCollection(3).Format.Fill.Transparency = 0
    objShape.SeriesCollection(3).Format.Fill.Solid

    objShape.Legend.Position = -4107
    
    objShape.HasTitle = True
    objShape.ChartTitle.Delete()
    objShape.Axes(1).Delete()
   
    objShape.Parent.Line.Visible = False 
    objShape.Parent.Top = 110
    objShape.Parent.Left = 160
    objShape.Parent.Width = 360
    objShape.Parent.Height = 50 * (rows+1)
    objShape.ChartData.Workbook.Application.Quit()
        
def write_level_chart(doc, level, data):
    needSort = True
    if (level == "High"):
        c = "高风险缺陷统计"
        color = -16777024
        position = 0
        sp = doc.Bookmarks("highLevelChart").Range.Select()
    elif (level == "Medium") :
        c = "中风险缺陷统计"
        color = -16727809
        position = 1
        sp = doc.Bookmarks("mediumLevelChart").Range.Select()
    elif (level == "Low") :
        c = "低风险缺陷统计"
        color = -2894893
        position = 0
        sp = doc.Bookmarks("lowLevelChart").Range.Select()
    else :
        total = 0
        for i in data :
            total = total + i
        
        c = "检测到总共 " + str(total) + " 个缺陷"
        color = -16777024
        position = 1
        sp = doc.Bookmarks("acceptance").Range.Select()
        needSort = False
    if(len(data) == 0):
        print "No " + level + " Impact Defects. \n"
        return       
    LevelChart = doc.Shapes.AddChart(Type=57).Chart
    LevelChart.ChartData.Workbook.Application.WindowState = -4140

    chartWorkSheet = LevelChart.ChartData.Workbook.Worksheets(1)
  
    chartWorkSheet.ListObjects("表1").Resize(chartWorkSheet.Range("A1:B"+str(len(data) + 1)))
    chartWorkSheet.Range("B1").FormulaR1C1 = c

    #
    if (needSort == True):
 
##        print data
##        print level
        highList = []
        for key, value in data.items():
            highList.append([key, value])

        sortHighList = sorted([(t[1],t[0]) for t in highList]) 
##    print data[0]
##    print sortHighList
        i = 1
        for item in sortHighList:
            i = i + 1
            chartWorkSheet.Range("A"+str(i)).FormulaR1C1 = item[1]
            chartWorkSheet.Range("B"+str(i)).FormulaR1C1 = item[0]
    else :
        #chartWorkSheet.ListObjects("表1").Resize(chartWorkSheet.Range("A1:B"+str(len(data))))
        LevelChart.Axes(2).MaximumScale = 1  # xlValue 2  坐标轴显示值。 
 
        chartWorkSheet.Range("B2:B"+str(len(data)+1)).NumberFormatLocal = "0.0%"
        #print data
        chartWorkSheet.Range("A5").FormulaR1C1 = "已查看的报告"
        chartWorkSheet.Range("B5").FormulaR1C1 = float(data[1]+data[2]+data[3])/float(total)

	chartWorkSheet.Range("A4").FormulaR1C1 = "BUG"
	chartWorkSheet.Range("B4").FormulaR1C1 = float(data[1])/float(total)

	chartWorkSheet.Range("A3").FormulaR1C1 = "bug, 但已知不会造成严重后果"
	chartWorkSheet.Range("B3").FormulaR1C1 = float(data[2])/float(total)
	
	chartWorkSheet.Range("A2").FormulaR1C1 = "误报"
	chartWorkSheet.Range("B2").FormulaR1C1 = float(data[3])/float(total)

    

    LevelChart.ChartGroups(1).VaryByCategories = False

    
    LevelChart.SeriesCollection(1).Format.Fill.Visible = True
    LevelChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = color
    LevelChart.SeriesCollection(1).Format.Fill.Transparency = 0
    LevelChart.SeriesCollection(1).Format.Fill.Solid

    #highLevelChart.ChartTitle.Characters.Font.Italic = True
    LevelChart.ChartTitle.Characters.Font.Size = 18

    ##highLevelChart.Axes(1).Format.TextFrame2.TextRange.Font.Size = 12 #xlCategory = 1, can not work
    LevelChart.ChartArea.Format.Line.Visible = False
    
    LevelChart.Legend.Delete()
    LevelChart.ApplyDataLabels()

    LevelChart.Parent.Top = 30 + position * 300
    LevelChart.Parent.Left = 50
    LevelChart.Parent.Width = 450
    LevelChart.Parent.Height = 50 * len(data)
    LevelChart.ChartData.Workbook.Application.Quit()
    time.sleep(2)

def write_defect_level_chart(doc):
    defectLevelChart = doc.Shapes.AddChart().Chart
    
    defectLevelChart.ChartData.Workbook.Application.WindowState = -4140
    
    chartWorkSheet = defectLevelChart.ChartData.Workbook.Worksheets(1)
    
    #chartWorkSheet.Visible = 2 # error xlSheetVeryHidden  2
    
    chartWorkSheet.ListObjects("表1").Resize(chartWorkSheet.Range("A1:B4"))
    
    chartWorkSheet.Range("B1").FormulaR1C1 = "缺陷风险等级统计"
    chartWorkSheet.Range("A2").FormulaR1C1 = "高风险"
    chartWorkSheet.Range("A3").FormulaR1C1 = "中风险"
    chartWorkSheet.Range("A4").FormulaR1C1 = "低风险"
    chartWorkSheet.Range("B2").FormulaR1C1 = calc_level_num()[0]
    chartWorkSheet.Range("B3").FormulaR1C1 = calc_level_num()[1]
    chartWorkSheet.Range("B4").FormulaR1C1 = calc_level_num()[2]

    defectLevelChart.ChartGroups(1).VaryByCategories = True

    #High Impact 修改为红色
    defectLevelChart.SeriesCollection(1).Points(1).Format.Fill.Visible = True
    defectLevelChart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = -16777024
    defectLevelChart.SeriesCollection(1).Points(1).Format.Fill.Transparency = 0
    defectLevelChart.SeriesCollection(1).Points(1).Format.Fill.Solid

    #Medium Impact 修改为黄色
    defectLevelChart.SeriesCollection(1).Points(2).Format.Fill.Visible = True
    defectLevelChart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = -16727809
    defectLevelChart.SeriesCollection(1).Points(2).Format.Fill.Transparency = 0
    defectLevelChart.SeriesCollection(1).Points(2).Format.Fill.Solid

    #Low Impact 取默认，为灰色
    
    defectLevelChart.Legend.Delete()
    defectLevelChart.ApplyDataLabels()

    defectLevelChart.Parent.Top = 75
    defectLevelChart.Parent.Left = 50

    defectLevelChart.ChartData.Workbook.Application.Quit()
    time.sleep(1)

# main entry
if (not(os.path.isfile('test'))):
    print "没有缺陷数据生成" + "\n"
    #exit(1)
else:
    os.remove('test')
    
wordApp = win32com.client.Dispatch("Word.Application")
#wordApp = gencache.EnsureDispatch("Word.Application")

# For the sake of ensuring the correct module is used...
mod = sys.modules[wordApp.__module__]
print "The module hosting the object is", mod

wordApp.Visible = False

pwd = "autosoft217liu"
noPwd = ''
doc = wordApp.Documents.Open(FileName=currPath +"\\TemplateNew.jasper", ReadOnly = True, PasswordDocument = pwd, WritePasswordDocument = pwd, Visible=False) #不知为何密码不起作用

try:
    
    #doc = wordApp.Documents.Add()
    doc.SaveAs(currPath + "\\" + projectName + label + ".docx")

    total_summary = calc_total_summary()
    #print total_summary
    #封面
    doc.Bookmarks("coverProjectName").Range.Text = projectName
    #第2页--项目概述
    doc.Bookmarks("summary_prjName").Range.Text = projectName
    doc.Bookmarks("summary_prjVer").Range.Text = projectVersion
    doc.Bookmarks("summary_totalLOC").Range.Text = total_summary[0]
    doc.Bookmarks("summary_Density").Range.Text = "%.3f" % float(float(total_summary[1])*1000/float(total_summary[0]))
    doc.Bookmarks("summary_HMdefectNum").Range.Text = total_summary[1]
    doc.Bookmarks("summary_COVtools").Range.Text = covTools
    doc.Bookmarks("summary_prjPerson").Range.Text = contact
    doc.Bookmarks("summary_testCompany").Range.Text = company
    doc.Bookmarks("summary_testPerson").Range.Text = testContact
    doc.Bookmarks("summary_Date").Range.Text = year + '年 '+ month+ '月 '+day+'日'
    
    #print  "InlineShape = " + str(doc.InlineShapes.Count) + " Shapes = " + str(doc.Shapes.Count)
    
    sp = doc.Bookmarks("defectLevelChart").Range.Select()
    write_defect_level_chart(doc)

    data = calc_category_num()
    
    write_level_chart(doc, "High", data[0])
    write_level_chart(doc, "Medium", data[1])
    write_level_chart(doc, "Low", data[2])
    
    data = calc_acceptance_num()
    write_level_chart(doc, "acceptance", data)


    sp = doc.Bookmarks("componentDensity").Range.Select()
    write_component_density(doc)

    #好像只能显示活动窗口的属性，不能修改 word选项设置的项
    #doc.ActiveWindow.View.ShowParagraphs = False
    doc.Save()
##except:
##    print "Exception Raised."
finally:

    doc.ExportAsFixedFormat(currPath + "\\" + projectName + label + ".pdf", ExportFormat = 17) #wdExportFormatPDF  17  将文档导出为 PDF 格式。
    doc.Close(-1)
    wordApp.Quit()
    print "The Report is generated at " + currPath + "\\" + projectName + label + ".pdf"
    #os.remove(currPath + "\\" + projectName + label + ".docx")
