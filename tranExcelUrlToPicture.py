#!/usr/bin/env python3
# encoding: utf-8

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
        
        self.df = None
        self.picDir = os.path.join(os.getcwd(),"downloadpics")   # 创建存储图片的文件夹

    # 从url链接下载图片命名为 fileName
    def GetWeb(self,url, fileName):  
        r = requests.get(url)                             # 返回url请求的数据
        data = r.content
        with open(fileName, 'wb') as f:                  # 将数据存储在指定位置
            f.write(data)

    # 取得表格中所有图片的url地址
    def GetPicUrls(self):
        pat = r'http'    # 利用正则匹配出图片url
        rec = re.compile(pat)
        urlList = []  
        self.df = pd.read_excel(self.fileName.GetValue())
        for index,row in self.df.iterrows():
            # print(type(row[0]),type(row[1]),row[1])
            if row[1] and len(rec.match(row[1]).groups() ) > 0:
                urlList.append((str(row[0]) + '_' + str(index),row[1]))         
        return urlList
    
    #下载图片
    def DownloadPic(self, event):
        urlList = self.GetPicUrls()
        self.loadPic.SetLabel(self.loadPic.Value + "开始下载图片！")
        threadList = []
        for picName,picUrl in urlList:                    
            try:
                picName = os.path.join(self.picDir,picName + '.' + picUrl.split('.')[-1])
                t = threading.Thread(target=self.GetWeb,args=(picUrl,picName))
                threadList.append(t) 
            except Exception as err:
                print(err)
                print(picName,picUrl)
        for t in threadList: 
            # self.loadPic.SetLabel(self.loadPic.Value + "\n开启线程：" + str(t))
            t.setDaemon(True)
            t.start()
            while threading.activeCount()>500:
                time.sleep(5)
        for t in threadList:
            t.join
        
        while threading.activeCount() > 1:
            self.loadPic.SetLabel(self.loadPic.Value + "\n剩余线程数：" + str(threading.activeCount()))
            time.sleep(5)
        self.loadPic.SetLabel(self.loadPic.Value + "\n下载图片完毕！")
        win32api.MessageBox(0, "下载图片完毕！", "提醒",win32con.MB_OK)   

    def OnOpen(self, event):
        dlg = wx.FileDialog(self, message='打开文件',
                            defaultDir='',
                            defaultFile='', 
                            wildcard=self.wildcard, 
                            style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.fileName.SetValue(dlg.GetPath())
            dlg.Destroy()  
        
        if not os.path.exists(self.picDir):
            os.mkdir(self.picDir)

    def ImportPicToExcel(self, event):
        self.loadPic.SetLabel(self.loadPic.Value + "开始导入图片！")
        with xlsxwriter.Workbook(self.fileName.GetValue()) as book:
            sheet = book.add_worksheet('Sheet1')
            sheet.set_column("C:C",10) #设置列宽
            for index,row in self.df.iterrows():
                sheet.set_row(index,60) #设置行高
                picPath = os.path.join(self.picDir,row[0] + '_' + str(index) + '.' + row[1].split('.')[-1])
                self.loadPic.SetLabel(self.loadPic.Value + "\n正在导入图片：" + picPath + "……")
                try:                    
                    with Image.open(picPath) as img:
                        sheet.write('A' + str(index+1),row[0])
                        sheet.write('B' + str(index+1),row[1])
                        sheet.insert_image('C' + str(index+1),picPath,{'y_offset': 3,'x_scale': 75/img.width, 'y_scale': 75/img.height,'url': row[1]}) #插入图片，同时设置纵向偏移及缩放比例
                except Exception as err:                    
                    sheet.write('A' + str(index+1),row[0])
                    sheet.write('B' + str(index+1),row[1])
                    sheet.write('C' + str(index+1),"图片未下载成功：" + str(err))
        self.loadPic.SetLabel(self.loadPic.Value + "\n导入图片完毕！")
        win32api.MessageBox(0, "导入图片完毕！", "提醒",win32con.MB_OK)

class App(wx.App):                      # 定义应用程序类
    def OnInit(self):                   # 类初始化方法
        self.frame = Frame()
        self.frame.Show(True)
        self.SetTopWindow(self.frame)   # 设置顶层框架
        return True

if __name__ == '__main__':              # 使用__name__检测当前模块
    app = App()
    app.MainLoop()    