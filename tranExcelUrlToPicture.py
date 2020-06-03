#!/usr/bin/env python3
# encoding: utf-8
#Author :dhj
#Date:2020-06-01

import wx
import os
import re
import pandas as pd
import threading
import requests
import xlsxwriter
from PIL import Image
import time
import win32api,win32con
from pathlib import Path


'''
    主要思路：
        1、先打开文件，取得文件路径等信息
        2、读取Excel，取得所有url地址
        3、根据url地址，下载所有图片到本地
        4、打开Excel，将图片插入的对应url地址后面一列
'''

class Frame(wx.Frame):                  # 定义GUI框架类
    # 框架初始化方法     
    def __init__(self, parent=None, id=-1, pos=wx.DefaultPosition,title='转换Excel中url为图片'):
        wx.Frame.__init__(self, parent, id, title,pos, size=(1200, 450))
        self.panel = wx.Panel(self)
        self.fileName = wx.TextCtrl(self.panel,pos=(5,5),size=(310,25))
        self.loadPic = wx.TextCtrl(self.panel,pos=(5,45),size=(1150,325),style=wx.TE_MULTILINE|wx.TE_RICH2)

        self.wildcard='表格文件(*.xlsx)|*.xlsx'
        self.openBtn = wx.Button(self.panel, -1, '打开', pos=(320, 5))
        self.openBtn.Bind(wx.EVT_BUTTON, self.OnOpen)

        self.saveAsBtn = wx.Button(self.panel, -1, '下载图片', pos=(400, 5))
        self.saveAsBtn.Bind(wx.EVT_BUTTON, self.DownloadPic)

        self.saveAsBtn = wx.Button(self.panel, -1, '导入图片', pos=(480, 5))
        self.saveAsBtn.Bind(wx.EVT_BUTTON, self.ImportPicToExcel)
        
        self.df = None #存储读取到的excel信息
        self.picDir = os.path.join(os.getcwd(),"downloadpics")   # 创建存储图片的文件夹
        self.urlList = []
        self.column = -1
        self.pattern = '^http[\w,\W]*'  #url地址匹配
    
    #获取url地址所在列
    def FindUrlColumn(self):
        if self.df.shape[0] > 0: #至少需要有一行记录
            for col in range(self.df.shape[1]):
                if re.search(self.pattern,str(self.df.iloc[0,col])):
                    self.column = col
            if self.column == -1:
                win32api.MessageBox(0, "请在第一行展现Url地址，以判断Url地址所在列！", "提醒",win32con.MB_OK)  
        else:
            win32api.MessageBox(0, "请确保表格中至少有一行记录！", "提醒",win32con.MB_OK)  

    # 取得表格中所有图片的url地址
    def GetUrlsFromFile(self):
        for index,row in self.df.iterrows():
            if row[self.column] and re.search(self.pattern,row[self.column]):
                self.urlList.append((str(index),row[self.column]))  

    # 从url链接下载图片命名为 fileName
    def SinglePicDownload(self,url, fileName): 
        try:
            r = requests.get(url)
            data = r.content
        except Exception as err:
            data = str(err)     #如果下载报错，将错误信息存入文件
        with open(fileName, 'wb') as f:                  # 将数据存储在指定位置
            f.write(data)

    def OnOpen(self, event):
        dlg = wx.FileDialog(self, message='打开文件',
                            defaultDir='',
                            defaultFile='', 
                            wildcard=self.wildcard, 
                            style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.fileName.SetValue(dlg.GetPath())
            dlg.Destroy()  
        
        #读取Excel内容
        self.df = pd.read_excel(self.fileName.GetValue())
        self.df.fillna("填充",inplace= True)
        # print(self.df)
        
        #定位Url地址所在列
        self.FindUrlColumn()
        
        #获取Excel中的Url地址
        self.GetUrlsFromFile()

        #在可执行文件所在位置创建目录：用于存放下载的图片
        if not os.path.exists(self.picDir):
            os.mkdir(self.picDir)
        # print(self.urlList)
        self.loadPic.SetLabel(self.loadPic.Value + "文件已打开，请开始转换！\n")

    #下载图片
    def DownloadPic(self, event):
        self.loadPic.SetLabel(self.loadPic.Value + "开始下载图片！\n")
        threadList = []

        #多线程下载图片
        for picName,picUrl in self.urlList:                    
            try:
                picName = os.path.join(self.picDir,picName + '.' + picUrl.split('.')[-1])
                t = threading.Thread(target=self.SinglePicDownload,args=(picUrl,picName))
                threadList.append(t) 
            except Exception as err:
                print(picName,picUrl,err)
        for t in threadList: 
            t.setDaemon(True)
            t.start()
            #当活动子线程数大于500时，阻塞
            while threading.activeCount()>500:
                time.sleep(1)
        for t in threadList:
            t.join
        
        #当活动子线程数大于1时，等待。目的是防止有子线程未执行完（即可能图片没有全部下载完），影响后面一步的Excel图片插入.
        while threading.activeCount() > 1:
            self.loadPic.SetLabel(self.loadPic.Value + "剩余线程数：" + str(threading.activeCount()) + "\n")
            time.sleep(1)
        self.loadPic.SetLabel(self.loadPic.Value + "下载图片完毕！\n")
        win32api.MessageBox(0, "下载图片完毕！", "提醒",win32con.MB_OK)  

    def ImportPicToExcel(self, event):
        self.loadPic.SetLabel(self.loadPic.Value + "开始导入图片！\n")
        newFileName = os.path.join(os.path.split(self.fileName.GetValue())[0],"新" + os.path.split(self.fileName.GetValue())[1]) #创建新文件
        with xlsxwriter.Workbook(newFileName) as book:
            sheet = book.add_worksheet('Sheet1')
            sheet.set_column(self.column + 1,self.column + 1,10) #地址所在列加1列用于放置图片，故增加宽度
            #处理列名
            for i in range(self.column + 1):
                sheet.write(0,i,self.df.columns.values.tolist()[i])
            sheet.write(0,self.column+1,"插入图片")
            for i in range(self.df.shape[1] - self.column - 1):
                sheet.write(0,i + self.column + 2,self.df.iloc[0,i + self.column + 1])     

            for row in range(1,self.df.shape[0]+1):  #按行依次插入：图片前self.column+1列，图片后df.shape[1]-self.column-1列，图片列
                picPath = os.path.join(self.picDir,str(row-1) + '.' + self.df.iloc[row-1,self.column].split('.')[-1])
                for col1 in range(self.column + 1):  #插入url所在位置前的列（含url列）
                    sheet.write(row,col1,self.df.iloc[row-1,col1])                
                for col2 in range(self.df.shape[1] - self.column - 1):  #插入url列所在位置后的列（不含url列）
                    sheet.write(row,col2 + self.column + 2,self.df.iloc[row-1,col2 + self.column + 1])                
                if Path(picPath).is_file():   #如果文件存在则插入图片
                    try:                    
                        with Image.open(picPath) as img:
                            sheet.insert_image(row,self.column + 1,picPath,{'y_offset': 3,'x_scale': 75/img.width, 'y_scale': 75/img.height,'url': self.df.iloc[row-1,self.column]}) #插入图片，同时设置纵向偏移及缩放比例                                                        
                            self.loadPic.SetLabel(self.loadPic.Value + "图片" + picPath + ",导入完毕！\n")
                            sheet.set_row(row,60) #设置行高
                    except Exception as err:                    
                        sheet.write(row,self.column + 1,"图片未下载成功：" + str(err))                        
        self.loadPic.SetLabel(self.loadPic.Value + "导入图片完毕！\n")
        win32api.MessageBox(0, "导入图片完毕！", "提醒",win32con.MB_OK)

class App(wx.App):                      # 定义应用类
    def OnInit(self):
        self.frame = Frame()
        self.frame.Show(True)
        self.SetTopWindow(self.frame)   # 设置窗口
        return True

if __name__ == '__main__':
    app = App()
    app.MainLoop()    
