#encoding:utf-8
import requests
import re

import xlrd
import xlwt

import datetime

from matplotlib import pyplot as plt
import mpl_finance as mpf
from matplotlib.pylab import date2num
from matplotlib.font_manager import FontProperties

import wx
from wx import adv

import time
import os


TitleColour = (219,62,62) #red
BackgroundColour = (249,251,252) #white
SeperationColour = (233,233,233) #gray
LeftBakgroundColour =(255,255,255) #white
LeftButEnterColour = (255,238,238) #pink

OrangeColour = (255,163,11) #Orange
LightOrangeColour = (255,189,80) #LightOrange

GrayColour = (220,220,220) #gray
DarkGrayColour = (180,180,180) #Dark gray
BrownColour = (172,151,151) #Brown

GreenColour = (132,209,79)

LigthPinkColour = (230,184,183) #pink


MidButEnterColour = (0,205,205) #blue
MidButClickColour = (255,47,47) #red

SpaceButEnterColour = (235,235,235) #gray

class STOCK():
    def __init__(self):
        self.StockHtmlToInfoLinkTable = \
        {
            "f1": "f1",
            "f2": "close",
            "f3": "change",
            "f4": "f4",
            "f5": "volume",
            "f6": "turnover ",
            "f7": "amplitude",
            "f8": "turnoverRate",
            "f9": "per",
            "f10": "volumeRatio",
            "f11": "f11",
            "f12": "code",
            "f13": "type",
            "f14": "name",
            "f15": "high",
            "f16": "low",
            "f17": "open",
            "f18": "preClose",
            "f20": "f20",
            "f21": "f21",
            "f22": "f22",
            "f23": "pbr",
            "f24": "f24",
            "f25": "f25",
            "f62": "f62",
            "f115": "f115",
            "f128": "f128",
            "f136": "f136",
            "f140": "f140",
            "f141": "f141",
            "f152": "f152"
        }
    
        self.AllStockList = []
        self.Workbook = xlwt.Workbook(encoding = 'ascii')
        
        # bmp = wx.Image("haha.gif", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        # adv.SplashScreen(bmp, adv.SPLASH_CENTRE_ON_SCREEN | adv.SPLASH_TIMEOUT, 2000, None, -1)
        # wx.Yield()
        # self.MainFrame = wx.Frame(None, -1, 'Ben', pos=(200, 200), size=(800,500), style=wx.SIMPLE_BORDER)
        self.MainFrame = wx.Frame(None, -1, 'Ben', pos=(200, 200), size=(800,500))
        self.MainFrame.Bind(wx.EVT_MOUSE_EVENTS, lambda event:self.moveFrame(event, self.MainFrame)) 
        # image1 = wx.StaticBitmap(self.MainFrame, -1, wx.Image('yueting11.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap(), (0, 0))
        # self.showFrame()
        # time.sleep(1)
        # image2 = wx.StaticBitmap(self.MainFrame, -1, wx.Image('yueting12.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap(), (0, 0))
        # self.showFrame()
        '''
        for timer in range(0, 255):
            self.MainFrame.SetTransparent(timer)
            time.sleep(0.031)
            self.showFrame()
            print(timer)
        '''
        # time.sleep(1)
        # image1.Hide()
        # image2.Hide()
        self.showFrame()
        
        wx.StaticText(self.MainFrame, -1, '月亭股票系统',pos=(100, 20), size=(120, 25),style = wx.ALIGN_LEFT)
        
        wx.StaticText(self.MainFrame, -1, 'Code',pos=(100, 70), size=(120, 25),style = wx.ALIGN_LEFT)
        self.CodeNumTextCtrl = wx.TextCtrl(self.MainFrame, -1, '', pos=(100, 100), size=(200, 25))
        
        wx.StaticText(self.MainFrame, -1, 'Date',pos=(100, 150), size=(120, 25),style = wx.ALIGN_LEFT)
        self.DateTextCtrl = wx.TextCtrl(self.MainFrame, -1, '', pos=(100, 180), size=(200, 25))
        
        wx.StaticText(self.MainFrame, -1, 'Days',pos=(100, 230), size=(120, 25),style = wx.ALIGN_LEFT)
        self.DaysTextCtrl = wx.TextCtrl(self.MainFrame, -1, '', pos=(100, 260), size=(200, 25))
        
        getButton = wx.Button(self.MainFrame, -1, 'Get', pos=(100, 340), size=(200, 25))
        getButton.Bind(wx.EVT_MOUSE_EVENTS, lambda event:self.clickButton(event, getButton, LightOrangeColour, OrangeColour, '', self.getInputInfo))

        startButton = wx.Button(self.MainFrame, -1, 'Start', pos=(100, 420), size=(200, 25))
        startButton.Bind(wx.EVT_MOUSE_EVENTS, lambda event:self.clickButton(event, startButton, LightOrangeColour, OrangeColour, '', self.testClick))

        self.getAllBasicStockInfo()
        print("It is OK")

    def saveExcel(self):
        self.Workbook.save('Test.xls') 
        
    def getOnePageBasicStockInfo(self, url):
        res = requests.get(url)
        # The html source is xxxxx({xxxxxxxx:{"total":4285,"diff":[{DataOne}, {DataTwo}, {DataThree}, {DataN}] 
        # The pattern is to get all Data, without use the symbol '?'
        pattern = re.compile(r'\[(.*)\]')
        result = pattern.findall(res.text)
        # The pattern is to get each Data, use the symbol '?'
        pattern = re.compile(r'{(.*?)}')
        
        result = pattern.findall(result[0])
        for i in range(0, len(result)):
            stockHtmlInfoStr = result[i]
            stockHtmlInfoList = stockHtmlInfoStr.split(',')

            # Init stockBasicInfoDict default value to None for all key in self.StockHtmlToInfoLinkTable
            stockBasicInfoDict = {}
            for key in self.StockHtmlToInfoLinkTable:
                stockBasicInfoDict[self.StockHtmlToInfoLinkTable[key]] = None

            # Assign the value to key
            # key = self.StockHtmlToInfoLinkTable[stockHtmlInfo.split(":")[0].strip()[1:-1]]
            # value = stockHtmlInfo.split(":")[1].strip()
            for stockHtmlInfo in stockHtmlInfoList:
                stockBasicInfoDict[(self.StockHtmlToInfoLinkTable[stockHtmlInfo.split(":")[0].strip()[1:-1]])] = stockHtmlInfo.split(":")[1].strip()

            self.AllStockList.append(stockBasicInfoDict)
       
    def getAllBasicStockInfo(self):
        # Initial the network page to 1
        netPageNum = 1
        # This URL is Shanghai and Shenzhen stock list, refer to eastmoney.com
        url = "http://56.push2.eastmoney.com/api/qt/clist/get?cb=jQuery1124024185585108590257_1609056571066&pn=" + str(netPageNum) + "&pz=20&po=1&np=1&ut=bd1d9ddb04089700cf9c27f6f7426281&fltt=2&invt=2&fid=f3&fs=m:0+t:6,m:0+t:13,m:0+t:80,m:1+t:2,m:1+t:23&fields=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f12,f13,f14,f15,f16,f17,f18,f20,f21,f23,f24,f25,f22,f11,f62,f128,f136,f115,f152&_=1609056583163"
        res = requests.get(url)
        # Get total number stocks from html, "total":4285,"
        pattern = re.compile(r'"total":(.*?),"')
        stockTotalNum = int(pattern.findall(res.text)[0])
        # Get all stock information from network page based on the total stock number(This is represent Shanghai and Shenzhen stock list)
        while(stockTotalNum > len(self.AllStockList)):
            url = "http://56.push2.eastmoney.com/api/qt/clist/get?cb=jQuery1124024185585108590257_1609056571066&pn=" + str(netPageNum) + "&pz=20&po=1&np=1&ut=bd1d9ddb04089700cf9c27f6f7426281&fltt=2&invt=2&fid=f3&fs=m:0+t:6,m:0+t:13,m:0+t:80,m:1+t:2,m:1+t:23&fields=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f12,f13,f14,f15,f16,f17,f18,f20,f21,f23,f24,f25,f22,f11,f62,f128,f136,f115,f152&_=1609056583163"
            self.getOnePageBasicStockInfo(url)
            netPageNum = netPageNum + 1
        
    def saveAllBasicStockInfo(self):
        # Get the excel title from self.StockHtmlToInfoLinkTable
        excelBasicInfoTitleList = []
        for key in self.StockHtmlToInfoLinkTable:
            excelBasicInfoTitleList.append(self.StockHtmlToInfoLinkTable[key])
            
        # Save the all basic stock information to excel sheet"Basic" 
        basicStockInfoSheet = self.Workbook.add_sheet('Basic')
        # Initial the sheet row and column number to 0
        rowNum = 0
        columnNum = 0
        # Write the title
        for i in range(0, len(excelBasicInfoTitleList)):
            columnNum = i
            basicStockInfoSheet.write(rowNum, columnNum, excelBasicInfoTitleList[i])
        rowNum = rowNum + 1
        # Write the stock data
        for stockBasicInfoDict in self.AllStockList:
            for i in range(0, len(excelBasicInfoTitleList)):
                columnNum = i
                basicStockInfoSheet.write(rowNum, columnNum, stockBasicInfoDict[excelBasicInfoTitleList[i]])
            rowNum = rowNum + 1
        self.saveExcel()
        
    def getAllStockDataList(self, attribute):
        retValueList = []
        attributeIfValid = False
        for key in self.StockHtmlToInfoLinkTable:
            if(attribute == self.StockHtmlToInfoLinkTable[key]):
                attributeIfValid = True
                
        if(True == attributeIfValid):
            for stockDict in self.AllStockList:
                retValueList.append(stockDict[attribute])
        else:
            # retValueList still []
            pass
        return(retValueList)
          
    def getDailyData(self):
        url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?fields1=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12,f13&fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61&beg=0&end=20500101&ut=fa5fd1943c7b386f172d6893dbfba10b&rtntype=6&secid=0.002475&klt=101&fqt=1&cb=jsonp1609487549370"       
        res = requests.get(url)
        pattern = re.compile(r'\[(.*)\]')
        result = pattern.findall(res.text)
        stockHtmlDataInfoList = result[0].split("\",\"")[1:-1]
        
        dailyStockInfoSheet = self.Workbook.add_sheet('Daily')
        excelTitleList = ["date", "open", "close", "high", "low", "code", "turnover", "amplitude", "change", "changeAmount", "turnoverRate"]
        rowNum = 0
        columnNum = 0
        # Write the title
        for i in range(0, len(excelTitleList)):
            columnNum = i
            dailyStockInfoSheet.write(rowNum, columnNum, excelTitleList[i])
        rowNum = rowNum + 1
        
        for stockHtmlDataInfo in stockHtmlDataInfoList:
            stockHtmlDataList = stockHtmlDataInfo.split(',')
            for i in range(0, len(stockHtmlDataList)):
                columnNum = i
                dailyStockInfoSheet.write(rowNum, columnNum, stockHtmlDataList[i])
            rowNum = rowNum + 1
            
    def formatDataforKLine(self, codeNum, kLineType):
        # self.CodeNum = '600000'
        codeList = self.getAllStockDataList('code')
        typeList = self.getAllStockDataList('type')
        
        kLineDateList = []
        kLineOpenDataList = []
        kLineCloseDataList = []
        kLineHighDataList = []
        kLineLowDataList = []
        kLineTurnoverList = []
        # K-line type
        klt = '101'
        if('day' == kLineType):
            klt = '101'
        elif('week' == kLineType):
            klt = '102' 
        else:
            pass
            
        if(("\"" + codeNum + "\"") in codeList):
            stockType = typeList[codeList.index("\"" + codeNum + "\"")]
            
            url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?fields1=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12,f13&fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61&beg=0&end=20500101&ut=fa5fd1943c7b386f172d6893dbfba10b&rtntype=6&secid=" + stockType + "." + codeNum + "&klt=" + klt + "&fqt=1&cb=jsonp1609487549370"       
            
            res = requests.get(url)
            pattern = re.compile(r'\[(.*)\]')
            result = pattern.findall(res.text)
            
            stockHtmlDataInfoList = result[0][1:-1].split("\",\"")

            excelTitleList = ["date", "open", "close", "high", "low", "code", "turnover", "amplitude", "change", "changeAmount", "turnoverRate"]
            
            # K-Line data sequence to follow function mpf.candlestick_ohlc
            # date, open, high, low, close
            j = 0
            for i in range(0, len(stockHtmlDataInfoList)):
                stockHtmlDataInfo = stockHtmlDataInfoList[i].split(',')
                kLineDateList.append(stockHtmlDataInfo[excelTitleList.index('date')])
                # kLineDataList.append((j, float(stockHtmlDataInfo[excelTitleList.index('open')]), float(stockHtmlDataInfo[excelTitleList.index('high')]), float(stockHtmlDataInfo[excelTitleList.index('low')]), float(stockHtmlDataInfo[excelTitleList.index('close')])))
                kLineOpenDataList.append(float(stockHtmlDataInfo[excelTitleList.index('open')]))
                kLineCloseDataList.append(float(stockHtmlDataInfo[excelTitleList.index('close')]))
                kLineHighDataList.append(float(stockHtmlDataInfo[excelTitleList.index('high')]))
                kLineLowDataList.append(float(stockHtmlDataInfo[excelTitleList.index('low')]))
                kLineTurnoverList.append(float(stockHtmlDataInfo[excelTitleList.index('turnover')]))
                j = j + 1
        return(kLineDateList, kLineOpenDataList, kLineCloseDataList, kLineHighDataList, kLineLowDataList, kLineTurnoverList)
           
    def displayKLine(self, codeNum, path, kLineDateList, kLineOpenDataList, kLineCloseDataList, kLineHighDataList, kLineLowDataList, kLineTurnoverList):
        codeList = self.getAllStockDataList('code')
        typeList = self.getAllStockDataList('name')
        if(("\"" + codeNum + "\"") in codeList):
            stockName = typeList[codeList.index("\"" + codeNum + "\"")]
    
        fig = plt.figure(figsize = (30,15))

        rect1 = [0.1,0.4,0.8,0.5]
        rect2 = [0.1,0.2,0.8,0.2]
        
        axDayKline = plt.axes(rect1)
        axDayVolume = plt.axes(rect2)
        font = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=14) 
        axDayKline.set_title((str(u'股票代码：'.encode('utf-8').decode('utf-8')) + codeNum + '   股票名称：' +  stockName), fontproperties = font, fontsize = 20)

        averageData10List = self.movingAverage(kLineCloseDataList, 10)
        averageData30List = self.movingAverage(kLineCloseDataList, 30)
        
        # Set K-Line drawing
        kLineDataList = []
        for i in range(0, len(kLineDateList)):
            kLineDataList.append((i, kLineOpenDataList[i], kLineHighDataList[i], kLineLowDataList[i], kLineCloseDataList[i]))
        mpf.candlestick_ohlc(axDayKline, kLineDataList, width=0.5, colorup='r', colordown='g', alpha=0.6)
        axDayKline.plot(averageData10List, label='10day Average')
        axDayKline.plot(averageData30List, label='30day Average')
        axDayKline.legend(loc = 'upper left')   
        axDayKline.set_xticks(range(0, len(kLineDateList)))

        mpf.volume_overlay(axDayVolume, kLineOpenDataList, kLineCloseDataList, kLineTurnoverList, colorup='r', colordown='g', width=0.5, alpha=0.8)
        maxDisplayLables = 50
        count = len(kLineDateList)

        kLineDateSeqList = []
        kLineDateLableList = []
        gap = int(count/maxDisplayLables)
        index = 0
        for i in range(0, maxDisplayLables):   
            kLineDateSeqList.append(index)
            kLineDateLableList.append(kLineDateList[index])
            index = index + gap
            
        axDayVolume.set_xticks(kLineDateSeqList)
        # axDayVolume.set_xlim(20, 5)
        axDayVolume.set_xticklabels(kLineDateLableList, rotation = 90)
        axDayVolume.grid(color='gray', linestyle='dashed')
        
        plt.subplots_adjust(hspace=0)
        # plt.show()
        plt.savefig(path + codeNum + ".png")
        
    def movingAverage(self, dataList, number):
        averageDataList = []
        if(number > len(dataList)):
            averageDataList = []
        else:
            for i in range(0, number):
                averageData = None
                averageDataList.append(averageData)
            for i in range(number, len(dataList)):
                sum = 0
                for j in range(i - number, i):
                    sum = sum + dataList[j]
                averageData = sum/number 
                averageDataList.append(averageData)
            
        return(averageDataList)    

    def getAllStockName(self):
        self.getAllBasicStockInfo()
        nameList = self.getAllStockDataList('name')
        print(nameList)

    def getInputInfo(self):
        localPath = os.getcwd()
        path = localPath + '\output\\' + str(time.strftime("%Y%m%d_%H%M%S", time.localtime())) + '\\'
        if not (os.path.exists(path)):
            os.makedirs(path) 
        else:
            pass
    
        self.CodeNum = self.CodeNumTextCtrl.GetValue()
        self.DateNum = self.DateTextCtrl.GetValue()
        self.DaysNum = self.DaysTextCtrl.GetValue()
        
        (kLineDateList, kLineOpenDataList, kLineCloseDataList, kLineHighDataList, kLineLowDataList, kLineTurnoverList) = self.formatDataforKLine(self.CodeNum, kLineType = 'day')

        for j in range(0, (len(kLineDateList) - 60)):
            if(float(kLineCloseDataList[j]) * 2 < float(kLineCloseDataList[j + 60])):
                if(100 < (len(kLineDateList) - j)):
                    kLineDateList = kLineDateList[(j - 200) : (j + 100)]
                    kLineOpenDataList = kLineOpenDataList[(j - 200) : (j + 100)]
                    kLineCloseDataList = kLineCloseDataList[(j - 200) : (j + 100)]
                    kLineHighDataList = kLineHighDataList[(j - 200) : (j + 100)]
                    kLineLowDataList = kLineLowDataList[(j - 200) : (j + 100)]
                    kLineTurnoverList = kLineTurnoverList[(j - 200) : (j + 100)]
                else:
                    kLineDateList = kLineDateList[(j - 300) : len(kLineDateList)]
                    kLineOpenDataList = kLineOpenDataList[(j - 300) : len(kLineOpenDataList)]
                    kLineCloseDataList = kLineCloseDataList[(j - 300) : len(kLineCloseDataList)]
                    kLineHighDataList = kLineHighDataList[(j - 300) : len(kLineHighDataList)]
                    kLineLowDataList = kLineLowDataList[(j - 300) : len(kLineLowDataList)]
                    kLineTurnoverList = kLineTurnoverList[(j - 300) : len(kLineTurnoverList)]
                self.displayKLine(self.CodeNum, path, kLineDateList, kLineOpenDataList, kLineCloseDataList, kLineHighDataList, kLineLowDataList, kLineTurnoverList)
                break
        
        
    
    def clickButton(self, event, button, enterColour, leaveColour, clickColour, Func, *args):
        if event.Entering():
            if(enterColour != ''):
                button.SetBackgroundColour(enterColour)
                button.Refresh()
        elif event.ButtonDown():
            if(clickColour != ''):
                button.SetBackgroundColour(clickColour)
                button.Refresh()
        elif event.ButtonUp():
            Func(*args)
            if(enterColour != ''):
                button.SetBackgroundColour(enterColour)    
                button.Refresh()
            #self.WorkSpacePanel(button)
        elif event.Leaving():
            if(leaveColour != ''):
                button.SetBackgroundColour(leaveColour)
                button.Refresh()
                
    def moveFrame(self, event, Frame):
        """ The basic frame title move event """   
        if event.ButtonDown():
            self.LastCurPointerPos = event.GetPosition()
            self.Moveflag = True

        elif event.Dragging():
            ThisCurPointerPos = event.GetPosition()
            DertaPosition = event.GetPosition()
            NewFramePos = Frame.GetPosition()
            detar_X = int(ThisCurPointerPos[0]) - int(self.LastCurPointerPos[0])
            detar_Y = int(ThisCurPointerPos[1]) - int(self.LastCurPointerPos[1])
            DertaPosition[0] = detar_X
            DertaPosition[1] = detar_Y
            if(True == self.Moveflag):
                Frame.SetPosition(NewFramePos + DertaPosition)
        elif event.ButtonUp():
            self.Moveflag = False
        else:
            self.Moveflag = False
            
    def showFrame(self):
        self.MainFrame.Refresh()
        self.MainFrame.Show()
    
    def testClick(self):
        localPath = os.getcwd()
        path = localPath + '\output\\' + str(time.strftime("%Y%m%d_%H%M%S", time.localtime())) + '\\'
        if not (os.path.exists(path)):
            os.makedirs(path) 
        else:
            pass
        self.catchStock(path)
    
    def catchStock(self, path):
        codeList = self.getAllStockDataList('code')
        for i in range(0, len(codeList)):
            codeNum = codeList[i][1 : -1]
            print(str(i) + "  " + codeNum)
            try:
                (kLineDateList, kLineDataList, kLineOpenDataList, kLineCloseDataList, kLineTurnoverList) = self.formatDataforKLine(codeNum, kLineType = 'day')
                if(len(kLineDateList) > 60):
                    for j in range(0, (len(kLineDateList) - 60)):
                        if(float(kLineCloseDataList[j]) * 2 < float(kLineCloseDataList[j + 60])):
                            if(100 < (len(kLineDateList) - j)):
                                kLineDateList = kLineDateList[(j - 200) : (j + 100)]
                                kLineDataList = kLineDataList[(j - 200) : (j + 100)]
                                kLineOpenDataList = kLineOpenDataList[(j - 200) : (j + 100)]
                                kLineCloseDataList = kLineCloseDataList[(j - 200) : (j + 100)]
                                kLineTurnoverList = kLineTurnoverList[(j - 200) : (j + 100)]
                            else:
                                kLineDateList = kLineDateList[(j - 300) : len(kLineDateList)]
                                kLineDataList = kLineDataList[(j - 300) : len(kLineDataList)]
                                kLineOpenDataList = kLineOpenDataList[(j - 300) : len(kLineOpenDataList)]
                                kLineCloseDataList = kLineCloseDataList[(j - 300) : len(kLineCloseDataList)]
                                kLineTurnoverList = kLineTurnoverList[(j - 300) : len(kLineTurnoverList)]
                            self.displayKLine(codeNum, path, kLineDateList, kLineDataList, kLineOpenDataList, kLineCloseDataList, kLineTurnoverList)
                            break
            except:
                print(codeNum + "get fail")
    
def Main():
    app = wx.App()
    app.locale  = wx.Locale(wx.LANGUAGE_ENGLISH)
    Interface = STOCK();
    # Interface.showFrame()
    app.MainLoop()
    input("OK")
    
if __name__ == '__main__':    
    Main()