#coding=utf-8
#import libs 
import sys
import add_cmd
import add_sty
import Fun
import os
import tkinter
from   tkinter import *
import tkinter.ttk
import tkinter.font
#Add your Varial Here: (Keep This Line of comments)
#Define UI Class
class  add:
    def __init__(self,root,isTKroot = True):
        uiName = self.__class__.__name__
        Fun.Register(uiName,'UIClass',self)
        self.root = root
        Fun.Register(uiName,'root',root)
        style = add_sty.SetupStyle()
        if isTKroot == True:
            root.title("Form1")
            Fun.CenterDlg(uiName,root,778,406)
            root['background'] = '#efefef'
        Form_1= tkinter.Canvas(root,width = 10,height = 4)
        Form_1.place(x = 0,y = 0,width = 778,height = 406)
        Form_1.configure(bg = "#efefef")
        Form_1.configure(highlightthickness = 0)
        Fun.Register(uiName,'Form_1',Form_1)
        #Create the elements of root 
        Label_2 = tkinter.Label(Form_1,text="企业名称")
        Fun.Register(uiName,'Label_2',Label_2)
        Fun.SetControlPlace(uiName,'Label_2',53,27,77,34)
        Label_2.configure(relief = "flat")
        Label_2_Ft=tkinter.font.Font(family='System', size=12,weight='bold',slant='roman',underline=0,overstrike=0)
        Label_2.configure(font = Label_2_Ft)
        Label_3 = tkinter.Label(Form_1,text="统一社会信用代码")
        Fun.Register(uiName,'Label_3',Label_3)
        Fun.SetControlPlace(uiName,'Label_3',344,26,156,35)
        Label_3.configure(relief = "flat")
        Label_3_Ft=tkinter.font.Font(family='System', size=12,weight='bold',slant='roman',underline=0,overstrike=0)
        Label_3.configure(font = Label_3_Ft)
        Entry_5_Variable = Fun.AddTKVariable(uiName,'Entry_5','')
        Entry_5 = tkinter.Entry(Form_1,textvariable=Entry_5_Variable)
        Fun.Register(uiName,'Entry_5',Entry_5)
        Fun.SetControlPlace(uiName,'Entry_5',136,25,138,36)
        Entry_5.configure(relief = "sunken")
        Entry_6_Variable = Fun.AddTKVariable(uiName,'Entry_6','')
        Entry_6 = tkinter.Entry(Form_1,textvariable=Entry_6_Variable)
        Fun.Register(uiName,'Entry_6',Entry_6)
        Fun.SetControlPlace(uiName,'Entry_6',528,25,187,36)
        Entry_6.configure(relief = "sunken")
        Button_8 = tkinter.Button(Form_1,text="查询")
        Fun.Register(uiName,'Button_8',Button_8)
        Fun.SetControlPlace(uiName,'Button_8',307,333,160,48)
        Button_8.configure(command=lambda:add_cmd.Button_8_onCommand(uiName,"Button_8"))
        Button_8_Ft=tkinter.font.Font(family='System', size=12,weight='bold',slant='roman',underline=0,overstrike=0)
        Button_8.configure(font = Button_8_Ft)
        ListBox_9 = tkinter.Listbox(Form_1)
        Fun.Register(uiName,'ListBox_9',ListBox_9)
        Fun.SetControlPlace(uiName,'ListBox_9',60,175,669,141)
        Label_11 = tkinter.Label(Form_1,text="资料起始日期")
        Fun.Register(uiName,'Label_11',Label_11)
        Fun.SetControlPlace(uiName,'Label_11',37,80,100,20)
        Label_11.configure(relief = "flat")
        Label_11_Ft=tkinter.font.Font(family='System', size=12,weight='bold',slant='roman',underline=0,overstrike=0)
        Label_11.configure(font = Label_11_Ft)
        Entry_12_Variable = Fun.AddTKVariable(uiName,'Entry_12','')
        Entry_12 = tkinter.Entry(Form_1,textvariable=Entry_12_Variable)
        Fun.Register(uiName,'Entry_12',Entry_12)
        Fun.SetControlPlace(uiName,'Entry_12',137,80,133,23)
        Entry_12.configure(relief = "sunken")
        Label_14 = tkinter.Label(Form_1,text="资料截至日期")
        Fun.Register(uiName,'Label_14',Label_14)
        Fun.SetControlPlace(uiName,'Label_14',371,80,100,20)
        Label_14.configure(relief = "flat")
        Label_14_Ft=tkinter.font.Font(family='System', size=12,weight='bold',slant='roman',underline=0,overstrike=0)
        Label_14.configure(font = Label_14_Ft)
        Entry_15_Variable = Fun.AddTKVariable(uiName,'Entry_15','')
        Entry_15 = tkinter.Entry(Form_1,textvariable=Entry_15_Variable)
        Fun.Register(uiName,'Entry_15',Entry_15)
        Fun.SetControlPlace(uiName,'Entry_15',486,80,120,20)
        Entry_15.configure(relief = "sunken")
        #Inital all element's Data 
        Fun.InitElementData(uiName)
        #Add Some Logic Code Here: (Keep This Line of comments)


#Create the root of Kinter 
if  __name__ == '__main__':
    root = tkinter.Tk()
    MyDlg = add(root)
    root.mainloop()
