#coding=utf-8
#import libs 
import sys
import Embedding_cmd
import Embedding_sty
import Fun
import os
import Page1
import tkinter
from   tkinter import *
import tkinter.ttk
import tkinter.font
#Add your Varial Here: (Keep This Line of comments)
#Define UI Class
class  Embedding:
    def __init__(self,root,isTKroot = True):
        uiName = self.__class__.__name__
        Fun.Register(uiName,'UIClass',self)
        self.root = root
        Fun.Register(uiName,'root',root)
        style = Embedding_sty.SetupStyle()
        if isTKroot == True:
            root.title("Form1")
            Fun.CenterDlg(uiName,root,629,510)
            root['background'] = '#efefef'
        Form_1= tkinter.Canvas(root,width = 10,height = 4)
        Form_1.place(x = 0,y = 0,width = 629,height = 510)
        Form_1.configure(bg = "#efefef")
        Form_1.configure(highlightthickness = 0)
        Fun.Register(uiName,'Form_1',Form_1)
        #Create the elements of root 
        Frame_2 = tkinter.Frame(Form_1)
        Fun.Register(uiName,'Frame_2',Frame_2)
        Fun.SetControlPlace(uiName,'Frame_2',27,19,573,423)
        Frame_2.configure(bg = "#888888")
        Frame_2.configure(relief = "flat")
        Page1_2 = Page1.Page1(Frame_2,False)
        Fun.Register(uiName,'Page1_2',Page1_2)
        #Inital all element's Data 
        Fun.InitElementData(uiName)
        #Add Some Logic Code Here: (Keep This Line of comments)


#Create the root of Kinter 
if  __name__ == '__main__':
    root = tkinter.Tk()
    MyDlg = Embedding(root)
    root.mainloop()
