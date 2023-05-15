#coding=utf-8
#import libs 
import sys
import GUI_cmd
import GUI_sty
import Fun
import os
import tkinter
from   tkinter import *
import tkinter.ttk
import tkinter.font
#Add your Varial Here: (Keep This Line of comments)
#Define UI Class
class  GUI:
    def __init__(self,root,isTKroot = True):
        uiName = self.__class__.__name__
        Fun.Register(uiName,'UIClass',self)
        self.root = root
        Fun.Register(uiName,'root',root)
        style = GUI_sty.SetupStyle()
        if isTKroot == True:
            root.title("Form1")
            Fun.CenterDlg(uiName,root,1100,660)
            root['background'] = '#333333'
        Form_1= tkinter.Canvas(root,width = 10,height = 4)
        Form_1.place(x = 0,y = 0,width = 1100,height = 660)
        Form_1.configure(bg = "#333333")
        Form_1.configure(highlightthickness = 0)
        Fun.Register(uiName,'Form_1',Form_1)
        #Create the elements of root 
        #Inital all element's Data 
        Fun.InitElementData(uiName)
        #Add Some Logic Code Here: (Keep This Line of comments)


#Create the root of Kinter 
if  __name__ == '__main__':
    root = tkinter.Tk()
    MyDlg = GUI(root)
    root.mainloop()
