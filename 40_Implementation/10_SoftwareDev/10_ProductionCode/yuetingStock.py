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
from mpl_toolkits.axisartist.parasite_axes import HostAxes, ParasiteAxes
from matplotlib.ticker import MultipleLocator, FormatStrFormatter

import wx
from wx import adv

import time
import os
import gc


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
        self.getMarketIndexStockInfo()
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
       
    def getMarketIndexStockInfo(self):
        # Get ShangHai market index information
        self.ShKLineDateList = []
        self.ShKLineOpenDataList = []
        self.ShKLineCloseDataList = []
        self.ShKLineHighDataList = []
        self.ShKLineLowDataList = []
        self.ShKLineTurnoverList = []
        
        kLineDataListDict = self.formatDataforKLine(marketIndex = 'ShangHai', codeNum = 'NONE', kLineType = 'NONE')
        kLineDateList = kLineDataListDict['DateList']
        kLineOpenDataList = kLineDataListDict['OpenDataList']
        kLineCloseDataList = kLineDataListDict['CloseDataList']
        kLineHighDataList = kLineDataListDict['HighDataList']
        kLineLowDataList = kLineDataListDict['LowDataList']
        kLineTurnoverList = kLineDataListDict['TurnoverList']
        
        startNum = 0
        for i in range(0, len(kLineDateList)):
            if('2016' in kLineDateList[i]):
                startNum = i
                break
        for i in range(startNum, len(kLineDateList)):
            self.ShKLineDateList.append(kLineDateList[i])
            self.ShKLineOpenDataList.append(kLineOpenDataList[i])
            self.ShKLineCloseDataList.append(kLineCloseDataList[i])
            self.ShKLineHighDataList.append(kLineHighDataList[i])
            self.ShKLineLowDataList.append(kLineLowDataList[i])
            self.ShKLineTurnoverList.append(kLineTurnoverList[i])
        
        # Get Shenzhen market index information    
        self.SzKLineDateList = []
        self.SzKLineOpenDataList = []
        self.SzKLineCloseDataList = []
        self.SzKLineHighDataList = []
        self.SzKLineLowDataList = []
        self.SzKLineTurnoverList = []
        kLineDataListDict = self.formatDataforKLine(marketIndex = 'ShenZhen', codeNum = 'NONE', kLineType = 'NONE')
        kLineDateList = kLineDataListDict['DateList']
        kLineOpenDataList = kLineDataListDict['OpenDataList']
        kLineCloseDataList = kLineDataListDict['CloseDataList']
        kLineHighDataList = kLineDataListDict['HighDataList']
        kLineLowDataList = kLineDataListDict['LowDataList']
        kLineTurnoverList = kLineDataListDict['TurnoverList']
        
        startNum = 0
        for i in range(0, len(kLineDateList)):
            if('2016' in kLineDateList[i]):
                startNum = i
                break
        for i in range(startNum, len(kLineDateList)):
            self.SzKLineDateList.append(kLineDateList[i])
            self.SzKLineOpenDataList.append(kLineOpenDataList[i])
            self.SzKLineCloseDataList.append(kLineCloseDataList[i])
            self.SzKLineHighDataList.append(kLineHighDataList[i])
            self.SzKLineLowDataList.append(kLineLowDataList[i])
            self.SzKLineTurnoverList.append(kLineTurnoverList[i])         
         
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
            
    def formatDataforKLine(self, codeNum, kLineType, marketIndex):
        kLineDataListDict = {'DateList':[], 'OpenDataList':[], 'CloseDataList':[], 'HighDataList':[], 'LowDataList':[], 'TurnoverList':[], 'TurnoverRateList':[]}
        codeList = self.getAllStockDataList('code')
        typeList = self.getAllStockDataList('type')
        
        kLineDateList = []
        kLineOpenDataList = []
        kLineCloseDataList = []
        kLineHighDataList = []
        kLineLowDataList = []
        kLineTurnoverList = []
        kLineTurnoverRateList = []
        
        if("ShangHai" == marketIndex):
            url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112405928137549796721_1610794224551&secid=1.000001&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58&klt=101&fqt=0&beg=19900101&end=20220101&_=1610794224553"
        elif("ShenZhen" == marketIndex):
            url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112408148579023741576_1610794023579&secid=0.399001&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58&klt=101&fqt=0&beg=19900101&end=20220101&_=1610794023636"
        else:
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
            else:
                return None
            url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?fields1=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12,f13&fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61&beg=0&end=20500101&ut=fa5fd1943c7b386f172d6893dbfba10b&rtntype=6&secid=" + stockType + "." + codeNum + "&klt=" + klt + "&fqt=1&cb=jsonp1609487549370"       

        res = requests.get(url)
        pattern = re.compile(r'\[(.*)\]')
        result = pattern.findall(res.text)
        
        stockHtmlDataInfoList = result[0][1:-1].split("\",\"")

        excelTitleList = ["DateList", "OpenDataList", "CloseDataList", "HighDataList", "LowDataList", "TurnoverList", "businessVolume", "amplitude", "change", "changeAmount", "TurnoverRateList"]
        # K-Line data sequence to follow function mpf.candlestick_ohlc
        # Shanghai Market Example
        # date,           open,    close,   high,     low    turnover  
        # ['1990-12-19', '96.05', '99.98', '99.98', '95.79', '1260', '494000.00', '0.00']
        # Storck Example
        # CodeNum: 000411
        # date,           open,    close,   high,     low    turnover  businessVolume amplitude  change  changeAmount  turnoverRate
        # ['2021-01-29', '13.12', '12.61', '13.24', '12.41', '44524', '56633160.00', '6.34', '-3.67', '-0.48', '2.15']
        # ['2021-02-01', '12.63', '13.49', '13.60', '12.63', '67912', '90141395.00', '7.69', '6.98', '0.88', '3.28']
        # ['2021-02-02', '13.49', '13.28', '13.50', '13.08', '40722', '53915286.00', '3.11', '-1.56', '-0.21', '1.96']
        # ['2021-02-03', '13.20', '12.86', '13.26', '12.83', '30547', '39684134.00', '3.24', '-3.16', '-0.42', '1.47']
        # ['2021-02-04', '12.87', '12.59', '13.10', '12.40', '32246', '41010209.00', '5.44', '-2.10', '-0.27', '1.56']
        # ['2021-02-05', '12.77', '13.04', '13.39', '12.65', '54481', '71395356.00', '5.88', '3.57', '0.45', '2.63']
        j = 0
        kLineDataListDict = {}
        for i in range(0, len(excelTitleList)):
            kLineDataListDict[excelTitleList[i]] = []
            
        for i in range(0, len(stockHtmlDataInfoList)):
            stockHtmlDataInfo = stockHtmlDataInfoList[i].split(',')
            for j in range(0, len(excelTitleList)):
                if("DateList" == excelTitleList[j]):
                    kLineDataListDict[excelTitleList[j]].append(stockHtmlDataInfo[j])
                else:
                    try:
                        kLineDataListDict[excelTitleList[j]].append(float(stockHtmlDataInfo[j]))
                    except:
                        kLineDataListDict[excelTitleList[j]].append(0)
        
        return(kLineDataListDict)
           
    def saveKLineExcel(self, codeNum, path, excelKLineDataListDict, point1, point2):
        kLineDataListTypeList = ['DateList', 'OpenDataList', 'CloseDataList', 'HighDataList', 'LowDataList', 'TurnoverList', 'TurnoverRateList']
        kLineDateListDict = {}
        for kLineDataListType in kLineDataListTypeList:
            kLineDateListDict[kLineDataListType] = excelKLineDataListDict[kLineDataListType]     
        
        kLineWorkbook = xlwt.Workbook(encoding = 'ascii')
        dataSheet = kLineWorkbook.add_sheet('data')
        
        rowNum = 0
        columnNum = 0
        # Write title
        for i in range(0, len(kLineDataListTypeList)):
            columnNum = i
            dataSheet.write(rowNum, columnNum, kLineDataListTypeList[i])
        rowNum = rowNum + 1
            
        for i in range(0, len(kLineDateListDict["DateList"])):
            for j in range(0, len(kLineDataListTypeList)):
                columnNum = j
                dataSheet.write(rowNum, columnNum, kLineDateListDict[kLineDataListTypeList[j]][i])
            rowNum = rowNum + 1
            
        kLineWorkbook.save(path + codeNum + '_' + kLineDateListDict["DateList"][0] + ".xls")
        print('---------------excel saved')
        
        # Release memory
        del kLineWorkbook 
    
    def displayKLine(self, codeNum, path, displayKLineDataListDict, point1, point2):
        kLineDateList = displayKLineDataListDict['DateList']
        kLineOpenDataList = displayKLineDataListDict['OpenDataList']
        kLineCloseDataList = displayKLineDataListDict['CloseDataList']
        kLineHighDataList = displayKLineDataListDict['HighDataList']
        kLineLowDataList = displayKLineDataListDict['LowDataList'] 
        kLineTurnoverList = displayKLineDataListDict['TurnoverList']
        kLineTurnoverRateList = displayKLineDataListDict['TurnoverRateList']
        
        codeList = self.getAllStockDataList('code')
        nameList = self.getAllStockDataList('name')
        typeList = self.getAllStockDataList('type')
        if(("\"" + codeNum + "\"") in codeList):
            stockName = nameList[codeList.index("\"" + codeNum + "\"")]
            stockType = typeList[codeList.index("\"" + codeNum + "\"")]
        
        if('0' == stockType):
            stockTypeName = u'深证成指   '
            marketKLineDateList = self.SzKLineDateList
            marketKLineOpenDataList = self.SzKLineOpenDataList
            marketKLineCloseDataList = self.SzKLineCloseDataList
            marketKLineHighDataList = self.SzKLineHighDataList
            marketKLineLowDataList = self.SzKLineLowDataList
            marketKLineTurnoverList = self.SzKLineTurnoverList
        else:
            stockTypeName = u'上证指数   '
            marketKLineDateList = self.ShKLineDateList
            marketKLineOpenDataList = self.ShKLineOpenDataList
            marketKLineCloseDataList = self.ShKLineCloseDataList
            marketKLineHighDataList = self.ShKLineHighDataList
            marketKLineLowDataList = self.ShKLineLowDataList
            marketKLineTurnoverList = self.ShKLineTurnoverList

          
        fig = plt.figure(figsize = (30,15))
  
        #['LeftPos', 'BotownPos', 'Length', 'Heigth']
        rectMarketKline = [0.02,0.65,0.96,0.25]
        rectDayKline = [0.02,0.4,0.96,0.25]
        rectDayVolume = [0.02,0.25,0.96,0.15]
        rectDayVolumeRate = [0.02,0.1,0.96,0.15]
        
        axMaketKline = plt.axes(rectMarketKline)
        axDayKline = plt.axes(rectDayKline)
        axDayVolume = plt.axes(rectDayVolume)
        axDayVolumeRate = plt.axes(rectDayVolumeRate)
        
        font = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=14) 
        axMaketKline.set_title(stockTypeName + (str(u'股票代码：'.encode('utf-8').decode('utf-8')) + codeNum + '   股票名称：' +  stockName), fontproperties = font, fontsize = 20)

        
        # Set Market K-Line drawing
        marketKLineDataList = []
        # Set K-Line drawing
        kLineDataList = []
        for i in range(0, len(kLineDateList)):
            kLineDataList.append((i, kLineOpenDataList[i], kLineHighDataList[i], kLineLowDataList[i], kLineCloseDataList[i]))
            j = marketKLineDateList.index(kLineDateList[i])
            marketKLineDataList.append((j, marketKLineOpenDataList[j], marketKLineHighDataList[j], marketKLineLowDataList[j], marketKLineCloseDataList[j]))
        
        mpf.candlestick_ohlc(axMaketKline, marketKLineDataList, width=0.5, colorup='r', colordown='g', alpha=0.6)
        
        mpf.candlestick_ohlc(axDayKline, kLineDataList, width=0.5, colorup='r', colordown='g', alpha=0.6)
        averageData10List = self.movingAverage(kLineCloseDataList, 10)
        averageData30List = self.movingAverage(kLineCloseDataList, 30)
        axDayKline.plot(averageData10List, label='10day Average')
        axDayKline.plot(averageData30List, label='30day Average')
        axDayKline.plot([point1, point2],[kLineCloseDataList[point1], kLineCloseDataList[point2]], color = 'black', linewidth=2)
        axDayKline.legend(loc = 'upper left')   
        axDayKline.set_xticks(range(0, len(kLineDateList)))

        mpf.volume_overlay(axDayVolume, kLineOpenDataList, kLineCloseDataList, kLineTurnoverList, colorup='r', colordown='g', width=0.5, alpha=0.8)
        
        axDayVolumeRate.plot(list(range(0, len(kLineTurnoverRateList))), kLineTurnoverRateList)
        
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
            
        axDayVolumeRate.set_xticks(kLineDateSeqList)
        # axDayVolume.set_xlim(20, 5)
        axDayVolumeRate.set_xticklabels(kLineDateLableList, rotation = 90)
        axDayVolumeRate.grid(color='gray', linestyle='dashed')
        
        plt.subplots_adjust(hspace=0)
        # plt.show()
        plt.savefig(path + codeNum + '_' + kLineDateList[0] + ".png")
        # Release memory
        fig.clf()
        plt.close()
        del fig
        gc.collect()
        print('---------------image saved')
        
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
    
        codeNum = self.CodeNumTextCtrl.GetValue()
        self.DateNum = self.DateTextCtrl.GetValue()
        self.DaysNum = self.DaysTextCtrl.GetValue()
        
        self.catchStock(codeNum, path)
        
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
        self.catchAllStock(path)
    
    def catchAllStock(self, path):
        codeList = self.getAllStockDataList('code')
        for i in range(0, len(codeList)):
            codeNum = codeList[i][1 : -1]
            print(str(i) + "  " + codeNum)
            try:
                self.catchStock(codeNum, path)
            except:
                print(codeNum + "get fail")
    
    def catchStock(self, codeNum, path):
        # Generate image folder
        imagePath = path + '\Image\\'
        if not (os.path.exists(imagePath)):
            os.makedirs(imagePath) 
        else:
            pass
        # Generate excel data folder
        excelPath = path + '\Excel\\'
        if not (os.path.exists(excelPath)):
            os.makedirs(excelPath) 
        else:
            pass
        totalDataNumber = 350
        preDataNumber = 200
        deltaDataNumber = 15
        deltaDataMultiple = 2
        point1 = preDataNumber
        point2 = point1 + deltaDataNumber
        
        kLineDataListDict = self.formatDataforKLine(codeNum, kLineType = 'day', marketIndex = 'NONE')
        kLineDateList = kLineDataListDict['DateList']
        kLineOpenDataList = kLineDataListDict['OpenDataList']
        kLineCloseDataList = kLineDataListDict['CloseDataList']
        kLineHighDataList = kLineDataListDict['HighDataList']
        kLineLowDataList = kLineDataListDict['LowDataList']
        kLineTurnoverList = kLineDataListDict['TurnoverList']
        kLineTurnoverRateList = kLineDataListDict['TurnoverRateList']
        
        startNum = 0
        for i in range(0, len(kLineDateList)):
            if('2017' in kLineDateList[i]):
                startNum = i
                break 

        while(startNum + totalDataNumber < len(kLineDateList)):
            for j in range(startNum, (len(kLineDateList) - deltaDataNumber)):
                if(float(kLineCloseDataList[j]) * deltaDataMultiple < float(kLineCloseDataList[j + deltaDataNumber])):
                    if((totalDataNumber - preDataNumber) < (len(kLineDateList) - j)):
                        displayKLineDateList = kLineDateList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineOpenDataList = kLineOpenDataList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineCloseDataList = kLineCloseDataList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineHighDataList = kLineHighDataList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineLowDataList = kLineLowDataList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineTurnoverList = kLineTurnoverList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]
                        displayKLineTurnoverRateList = kLineTurnoverRateList[(j - preDataNumber) : (j + totalDataNumber - preDataNumber)]

                    else:
                        displayKLineDateList = kLineDateList[(len(kLineDateList) - totalDataNumber) : len(kLineDateList)]
                        displayKLineOpenDataList = kLineOpenDataList[(len(kLineDateList) - totalDataNumber) : len(kLineOpenDataList)]
                        displayKLineCloseDataList = kLineCloseDataList[(len(kLineDateList) - totalDataNumber) : len(kLineCloseDataList)]
                        displayKLineHighDataList = kLineHighDataList[(len(kLineDateList) - totalDataNumber) : len(kLineHighDataList)]
                        displayKLineLowDataList = kLineLowDataList[(len(kLineDateList) - totalDataNumber) : len(kLineLowDataList)]
                        displayKLineTurnoverList = kLineTurnoverList[(len(kLineDateList) - totalDataNumber) : len(kLineTurnoverList)]
                        displayKLineTurnoverRateList = kLineTurnoverRateList[(len(kLineDateList) - totalDataNumber) : len(kLineTurnoverRateList)]
                        
                        point1 = totalDataNumber - (len(kLineDateList) - j)
                        point2 = point1 + deltaDataNumber
                        
                    startNum = j + totalDataNumber - preDataNumber    
                    displayKLineDataListDict = {}
                    displayKLineDataListDict['DateList'] = displayKLineDateList
                    displayKLineDataListDict['OpenDataList'] = displayKLineOpenDataList
                    displayKLineDataListDict['CloseDataList'] = displayKLineCloseDataList
                    displayKLineDataListDict['HighDataList'] = displayKLineHighDataList
                    displayKLineDataListDict['LowDataList'] = displayKLineLowDataList
                    displayKLineDataListDict['TurnoverList'] = displayKLineTurnoverList
                    displayKLineDataListDict['TurnoverRateList'] = displayKLineTurnoverRateList

                    self.displayKLine(codeNum, imagePath, displayKLineDataListDict, point1, point2)
                    self.saveKLineExcel(codeNum, excelPath, displayKLineDataListDict, point1, point2)
                    
                    del displayKLineDataListDict
                    gc.collect()
                    break
            startNum = j + totalDataNumber - preDataNumber
        # Release memory
        del kLineDataListDict
        del kLineDateList
        del kLineOpenDataList
        del kLineCloseDataList
        del kLineHighDataList
        del kLineLowDataList
        del kLineTurnoverList
        del kLineTurnoverRateList
    
def Main():
    app = wx.App()
    app.locale  = wx.Locale(wx.LANGUAGE_ENGLISH)
    Interface = STOCK();
    # Interface.showFrame()
    app.MainLoop()
    input("OK")
    
def testMemory():
    a = list(range(10000*10000))
    del a
    gc.collect()  
    input("testMemory")
    
def test():
    fig = plt.figure(figsize = (30,15))
    
    #['LeftPos', 'BotownPos', 'Length', 'Heigth']
    rectMarketKline = [0.02,0.60,0.96,0.3]
    rectDayKline = [0.02,0.1,0.96,0.3]

    
    axMaketKline = plt.axes(rectMarketKline)
    axDayKline = plt.axes(rectDayKline)
    
    xmajorLocator = MultipleLocator(10)
    xmajorFormatter = FormatStrFormatter('%.0f')
    xminorLocator   = MultipleLocator(5) 
    ymajorLocator = MultipleLocator(0.5)
    ymajorFormatter = FormatStrFormatter('%.1f')
    yminorLocator   = MultipleLocator(0.1)
    

    axDayKline.xaxis.set_major_locator(xmajorLocator)  # 设置x轴主刻度
    axDayKline.xaxis.set_major_formatter(xmajorFormatter)  # 设置x轴标签文本格式
    axDayKline.xaxis.set_minor_locator(xminorLocator)  # 设置x轴次刻度

    axDayKline.yaxis.set_major_locator(ymajorLocator)  # 设置y轴主刻度
    axDayKline.yaxis.set_major_formatter(ymajorFormatter)  # 设置y轴标签文本格式
    axDayKline.yaxis.set_minor_locator(yminorLocator)  # 设置y轴次刻度
    axDayKline.xaxis.grid(True, linestyle = "-",which='both') #x坐标轴的网格使用主刻度
    axDayKline.yaxis.grid(True, linestyle = "-",which='minor') #y坐标轴的网格使用次刻度
    plt.show()
    input("OK")
    
if __name__ == '__main__':    
    Main()
