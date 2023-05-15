#coding=utf-8
import sys
import os
from   os.path import abspath, dirname
sys.path.append(abspath(dirname(__file__)))
import tkinter
import tkinter.filedialog
from   tkinter import *
import Fun
ElementBGArray={}  
ElementBGArray_Resize={} 
ElementBGArray_IM={} 

import newCUSRP
import warnings
import requests
import re
import urllib
import datetime
import xlrd
import calendar
import runreport as rp
import autodownload as at
warnings.filterwarnings('ignore')
data=newCUSRP.initialize()
def Button_11_onCommand(uiName,widgetName):
    startdate= Fun.GetText(uiName,'Entry_39')
    enddate = Fun.GetText(uiName,'Entry_8')
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "查询时间为："+startdate+"--"+enddate)
    try:
        if int(enddate)<int(startdate):
            Fun.MessageBox("截至日期在起始日期之前，请重新输入正确的时间！")
        else:
            if len(enddate) | len(startdate) !=6:
                Fun.MessageBox("日期输入错误，请重新输入正确格式，如202201")
            else:
                gender = Fun.GetTKVariable(uiName,'Group_1')
                placebox = Fun.GetElement(uiName,'ComboBox_13')
                place1=placebox.get()
                querrybox = Fun.GetElement(uiName,'ComboBox_18')
                querry=querrybox.get()
                whatbox = Fun.GetElement(uiName,'ComboBox_14')
                what=whatbox.get()
                whitebox= Fun.GetElement(uiName,'ComboBox_34')
                white=whitebox.get()
                dict_what={'所有':'all','开票额':'sale','纳税额':'tax','返税额-所有':'taxback_detail','返税额-应返':'taxback_apply','返税额-实返':'taxback_real'}
                dict_white={'否':True,'是':False}
                if querry=='每个月总和':
                    result=data.report(start=int(startdate),end=int(enddate),place=place1,method=dict_what.get(what),whole=dict_white.get(white))
                    if what=="所有":
                        a=str(int(result.sum()['价税合计']))
                        listBox.insert(tkinter.END, place1+'开票总额'+a)
                        b=str(int(result.sum()['合计']))
                        listBox.insert(tkinter.END, place1+'纳税总额'+b)
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                        d=str(int(result.sum()['实返合计']))
                        listBox.insert(tkinter.END, place1+'实返合计'+d)
                    if what=="开票额":
                        a=str(int(result.sum()['价税合计']))
                        listBox.insert(tkinter.END, place1+'开票总额'+a)
                    if what=="纳税额":
                        b=str(int(result.sum()['合计']))
                        listBox.insert(tkinter.END, place1+'纳税总额'+b)
                    if what=="返税额-所有":
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                        d=str(int(result.sum()['实返合计']))
                        listBox.insert(tkinter.END, place1+'实返合计'+d)
                    if what=="返税额-应返":
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                    if what=="返税额-实返":
                         d=str(int(result.sum()['实返合计']))
                         listBox.insert(tkinter.END, place1+'实返合计'+d)
                    listBox.insert(tkinter.END, '详细结果在文件中查看')
                    Fname='d:/analyze program/查询匹配/'+place1+startdate+"--"+enddate+what+'按月.xlsx'
                    result.to_excel(Fname)
                if querry=='每个公司总和':
                    if what == '所有':
                        result=data.all_calculate(start=int(startdate),end=int(enddate),place=place1,whole=dict_white.get(white))
                        a=str(int(result.sum()['价税合计']))
                        listBox.insert(tkinter.END, place1+'开票总额'+a)
                        b=str(int(result.sum()['合计']))
                        listBox.insert(tkinter.END, place1+'纳税总额'+b)
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                        d=str(int(result.sum()['实返合计']))
                        listBox.insert(tkinter.END, place1+'实返合计'+d)
                    if what == '开票额':
                        result=data.Sale_calculate(start=int(startdate),end=int(enddate),place=place1,whole=dict_white.get(white))
                        a=str(int(result.sum()['价税合计']))
                        listBox.insert(tkinter.END, place1+'开票总额'+a)
                    if what == '纳税额':
                        result=data.Tax_calculate(start=int(startdate),end=int(enddate),place=place1,method='type',whole=dict_white.get(white))
                        b=str(int(result.sum()['合计']))
                        listBox.insert(tkinter.END, place1+'纳税总额'+b)
                    if what == '返税额-所有':
                        result=data.TaxBack_calculate(start=int(startdate),end=int(enddate),place=place1,method='all',whole=dict_white.get(white))
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                        d=str(int(result.sum()['实返合计']))
                        listBox.insert(tkinter.END, place1+'实返合计'+d)
                    if what == '返税额-应返':
                        result=data.TaxBack_calculate(start=int(startdate),end=int(enddate),place=place1,method='apply',whole=dict_white.get(white))
                        c=str(int(result.sum()['申请合计']))
                        listBox.insert(tkinter.END, place1+'应返总额'+c)
                    if what == '返税额-实返':
                        result=data.TaxBack_calculate(start=int(startdate),end=int(enddate),place=place1,method='real',whole=dict_white.get(white))   
                        d=str(int(result.sum()['实返合计']))
                        listBox.insert(tkinter.END, place1+'实返合计'+d)
                    listBox.insert(tkinter.END, '详细结果在文件中查看')
                    Fname='d:/analyze program/查询匹配/'+place1+startdate+"--"+enddate+what+'按公司.xlsx'
                    result.to_excel(Fname)
                print(result)  
    except:Fun.MessageBox("时间输入格式错误，请重新输入正确的时间！")
    #combobox.current(0)
def Button_21_onCommand(uiName,widgetName):
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "正在下载数据，预计时间2分钟...")
    at.autoDownload()
    listBox.insert(tkinter.END,'所有数据下载完成')
def Button_27_onCommand(uiName,widgetName):
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "请输入需要查询的月份，并将资料放在待匹配企业文件夹中")
    
    
    
def Button_28_onCommand(uiName,widgetName):
    startdate= Fun.GetText(uiName,'Entry_39')
    enddate= Fun.GetText(uiName,'Entry_8')
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "分析报表的时间区间为："+startdate+"--"+enddate)
    try:
        if int(enddate)<int(startdate):
            Fun.MessageBox("截至日期在起始日期之前，请重新输入正确的时间！")
        else:
            if len(enddate) | len(startdate) !=6:
                Fun.MessageBox("日期输入错误，请重新输入正确格式，如202201")
            else:
                rp.makereport(start=startdate,end=enddate)
                listBox.insert(tkinter.END, '报告文件生成成功！')
    except:Fun.MessageBox("生成报告失败！")
def Button_29_onCommand(uiName,widgetName):
    pass
def Button_36_onCommand(uiName,widgetName):
    startdate1= Fun.GetText(uiName,'Entry_39')
    enddate1= Fun.GetText(uiName,'Entry_8')
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "分析报表的时间区间为："+startdate1+"--"+enddate1)
    if int(enddate1)<int(startdate1):
        Fun.MessageBox("截至日期在起始日期之前，请重新输入正确的时间！")
    else:
        if len(enddate1) | len(startdate1) !=6:
            Fun.MessageBox("日期输入错误，请重新输入正确格式，如202201")
        else:
            TaxDownload(stardate=int(startdate1),enddate=int(enddate1))
            listBox.insert(tkinter.END, '海口完税证明导出成功！')
            listBox.insert(tkinter.END, '入库时间区间为：'+startdate1+'--'+enddate1)
def Button_37_onCommand(uiName,widgetName):
    startdate= Fun.GetText(uiName,'Entry_39')
    enddate = Fun.GetText(uiName,'Entry_8')
    placebox = Fun.GetElement(uiName,'ComboBox_13')
    place1=placebox.get()
    listBox = Fun.GetElement(uiName, "ListBox_43")
    listBox.insert(tkinter.END, "统计时间为："+startdate+"--"+enddate)
    whitebox= Fun.GetElement(uiName,'ComboBox_34')
    white=whitebox.get()
    dict_white={'否':True,'是':False}
    result=data.Sale_calculate(start=int(startdate),end=int(enddate),place=place1,method='goods',whole=dict_white.get(white))   
    Fname='d:/analyze program/查询匹配/'+place1+startdate+"--"+enddate+'开票种类.xlsx'
    result.to_excel(Fname)
    listBox.insert(tkinter.END, "结果已生成在查询匹配文件中")
def Button_70_onCommand(uiName,widgetName):
    topLevel = tkinter.Toplevel()
    topLevel.attributes("-toolwindow", 1)
    topLevel.wm_attributes("-topmost", 1)
    import add
    add.add(topLevel)
    tkinter.Tk.wait_window(topLevel)
