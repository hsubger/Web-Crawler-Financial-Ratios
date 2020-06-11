import requests
import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup as bs
from pprint import pprint

'''
目的: 讀取資產負債表及損益表，並計算相關財務分析比率
目標公司: [1101, 1102, 1103, 1104, 1108, 1109]
資料來源: 永豐金證券財務分析(公開資訊觀測站報表科目較簡略，所以改用永豐金證券)
程式執行後會顯示個比率計算結果，同時也會輸出excel及圖表
(財務分析比率涵蓋 1.短期償債能力 2.長期償債能力 3.資產使用效率 4.獲利能力 5.投資報酬率)
'''

class statements():
    def __init__(self, co_ids):
        self.__co_ids = co_ids
        self.__urls = {
            'incomeStatement':'https://stockchannelnew.sinotrade.com.tw/z/zc/zcq/zcqa/zcqa_',
            'financialPosition':'https://stockchannelnew.sinotrade.com.tw/z/zc/zcp/zcpb/zcpb_',
            #'cashFlow':'https://stockchannelnew.sinotrade.com.tw/z/zc/zc3/zc3a_'
            }
        self.__result = {}
    def getInfo(self):
        for co_id in self.__co_ids:
            temp = {}
            for statement, url in self.__urls.items(): 
                temp[statement] = self.__process(url + str(co_id) + '.djhtm')
            self.__result[co_id] = temp
        return self.__result

    def __process(self, url, retry = 0):
        clean = {}
        if retry == 100:
            clean['error'] = '無法取得資料'
            return clean
        re = requests.get(url)
        if re.status_code != requests.codes.ok: return self.__process(url, retry = retry + 1)
        if '查無' in re.text: 
            soup = bs(re.text, 'html5lib')
            return soup.select('.t3n0')[0].text
        df = pd.read_html(re.text)[2]
        for i in range(df.shape[0]):
            tempRowName = list( df.iloc[i, :] )[0]
            tempRowData = list( df.iloc[i, :] )[1:7]
            if tempRowName == "種類": continue
            for i in range(len(tempRowData)): tempRowData[i] = float(tempRowData[i])
            clean[tempRowName] = tempRowData
        return clean

class FSA():
    def __init__(self, co_ids, statements):
        self.__co_ids = co_ids
        self.__statements = statements

    def analysis(self):
        with pd.ExcelWriter('ratio.xlsx') as writer:
            temp, dfData = self.__solvency_short(writer)
            print('短期償債能力')
            print(dfData)
            dfData.to_excel(writer, sheet_name = '短期償債能力', header=False, index=False)
            self.__fig(dfData, '短期償債能力')

            temp, dfData = self.__solvency_long(writer)
            print('長期償債能力')
            print(dfData)
            dfData.to_excel(writer, sheet_name = '長期償債能力', header=False, index=False)
            self.__fig(dfData, '長期償債能力')

            temp, dfData = self.__assetEfficiency(writer)
            print('資產使用效率')
            print(dfData)
            dfData.to_excel(writer, sheet_name = '資產使用效率', header=False, index=False)
            self.__fig(dfData, '資產使用效率')

            temp, dfData = self.__earning(writer)
            print('獲利能力')
            print(dfData)
            dfData.to_excel(writer, sheet_name = '獲利能力', header=False, index=False)
            self.__fig(dfData, '獲利能力')

            temp, dfData = self.__returns(writer)
            print('投資報酬率')
            print(dfData)
            dfData.to_excel(writer, sheet_name = '投資報酬率', header=False, index=False)
            self.__fig(dfData, '投資報酬率')

    def __solvency_short(self, writer):
        temp = {}
        for co_id in self.__co_ids:
            temFP = self.__statements[co_id]['financialPosition']
            temIS = self.__statements[co_id]['incomeStatement']
            if type(temFP) == str or type(temIS) == str: 
                temp[co_id] = [temFP, temIS]
            else:
                finanAsset = [ a + b + c + d for a, b, c, d in zip(
                                        temFP['透過損益按公允價值衡量之金融資產－流動'],
                                        temFP['透過其他綜合損益按公允價值衡量之金融資產－流動'],
                                        temFP['按攤銷後成本衡量之金融資產－流動'],
                                        temFP['避險之金融資產－流動'],)]
                quickAsset = [ a + b + c for a, b, c in zip(temFP['現金及約當現金'], finanAsset, temFP['應收帳款及票據']) ]
                CostGoodsSold = temIS['營業成本'][:-1]
                aveInventory = [ (temFP['存貨'][i]+temFP['存貨'][i+1])/2 for i in range(len(temFP['存貨'])-1) ]
                temp[co_id] = {
                    'Current Ratio': [ round(a/b, 3) for a, b in zip(temFP['流動資產'], temFP['流動負債']) ][:-1],
                    'Quick Ratio':  [ round(a/b, 3) for a, b in zip(quickAsset, temFP['流動負債']) ][:-1],
                    'Days in Inventory': [ round(365/(a/b), 3) for a, b in zip(CostGoodsSold, aveInventory)]
                }

        ratioList = ['Current Ratio', 'Quick Ratio', 'Days in Inventory']
        dfData = self.__table(temp, ratioList)
        
        return temp, dfData

    def __solvency_long(self, writer):
        temp = {}
        for co_id in self.__co_ids:
            temFP = self.__statements[co_id]['financialPosition']
            temIS = self.__statements[co_id]['incomeStatement']
            if type(temFP) == str or type(temIS) == str: 
                temp[co_id] = [temFP, temIS]
            else:
                earnBeforeIT = [ a + b for a, b in zip(temIS['稅前淨利'], temIS['利息支出']) ]
                longTermFunds = [ a + b for a, b in zip(temFP['股東權益總額'], temFP['非流動負債']) ]
                temp[co_id] = {
                    'Debt to Total Assets Ratio': [ round(a/b, 3) for a, b in zip(temFP['負債總額'], temFP['資產總額']) ][:-1],
                    'Equity Ratio': [ round(a/b, 3) for a, b in zip(temFP['股東權益總額'], temFP['資產總額']) ][:-1],
                    "Debt To Shareholder's Equity": [ round(a/b, 3) for a, b in zip(temFP['負債總額'], temFP['股東權益總額']) ][:-1],
                    "Shareholder's Equity To Fixed Assets": [ round(a/b, 3) for a, b in zip(temFP['股東權益總額'], temFP['不動產廠房及設備']) ][:-1],
                    'Long Term Funds to Fixed Assets': [ round(a/b, 3) for a, b in zip(longTermFunds, temFP['不動產廠房及設備']) ][:-1],
                    'Times Interest Earned': [ round(a/b, 3) for a, b in zip(earnBeforeIT, temIS['利息支出']) ][:-1],
                }

        ratioList = ['Debt to Total Assets Ratio', 'Equity Ratio', "Debt To Shareholder's Equity",
                    "Shareholder's Equity To Fixed Assets", 'Long Term Funds to Fixed Assets', 'Times Interest Earned' ]
        dfData = self.__table(temp, ratioList)
        
        return temp, dfData

    def __assetEfficiency(self, writer):
        temp = {}
        for co_id in self.__co_ids:
            temFP = self.__statements[co_id]['financialPosition']
            temIS = self.__statements[co_id]['incomeStatement']
            if type(temFP) == str or type(temIS) == str: 
                temp[co_id] = [temFP, temIS]
            else:
                aveInventory = [ (temFP['存貨'][i]+temFP['存貨'][i+1])/2 for i in range(len(temFP['存貨'])-1) ]
                aveConstantAsset = [ (temFP['不動產廠房及設備'][i]+temFP['不動產廠房及設備'][i+1])/2 for i in range(len(temFP['不動產廠房及設備'])-1) ]
                aveAsset =  [ (temFP['資產總額'][i]+temFP['資產總額'][i+1])/2 for i in range(len(temFP['資產總額'])-1) ]
                temp[co_id] = {
                    'Sales to Cash': [ round(a/b, 3) for a, b in zip(temIS['營業收入淨額'], temFP['現金及約當現金']) ][:-1],
                    'Sales to Accounts Receivable': [ round(a/b, 3) for a, b in zip(temIS['營業收入淨額'], temFP['應收帳款及票據']) ][:-1],
                    'Sales to Inventory': [ round(a/b, 3) for a, b in zip(temIS['營業成本'][:-1], aveInventory) ],
                    'Salses to Fixed Assets': [ round(a/b, 3) for a, b in zip(temIS['營業收入淨額'][:-1], aveConstantAsset) ],
                    'Sales to Total Asset': [ round(a/b, 3) for a, b in zip(temIS['營業收入淨額'][:-1], aveAsset) ],
                }

        ratioList = ['Sales to Cash', 'Sales to Accounts Receivable', 'Sales to Inventory',
                      'Salses to Fixed Assets', 'Sales to Total Asset' ]
        dfData = self.__table(temp, ratioList)
        
        return temp, dfData
    
    def __earning(self, writer):
        temp = {}
        for co_id in self.__co_ids:
            temFP = self.__statements[co_id]['financialPosition']
            temIS = self.__statements[co_id]['incomeStatement']
            if type(temFP) == str or type(temIS) == str: 
                temp[co_id] = [temFP, temIS]
            else:
                afterTax = [ a-b for a, b in zip(temIS['稅前淨利'], temIS['所得稅費用']) ]
                temp[co_id] = {
                    'Gross Profit Margin': [ round(a/b, 3) for a, b in zip(temIS['營業毛利'], temIS['營業收入淨額']) ][:-1],
                    'Operating Net Profit Margin':  [ round(a/b, 3) for a, b in zip(temIS['營業利益'], temIS['營業收入淨額']) ][:-1],
                    'Pre-Tax Income Margin':[ round(a/b, 3) for a, b in zip(temIS['稅前淨利'], temIS['營業收入淨額']) ][:-1],
                    'Net Operating Profit After Tax': [ round(a/b, 3) for a, b in zip(afterTax, temIS['營業收入淨額']) ][:-1],
                }

        ratioList = ['Gross Profit Margin', 'Operating Net Profit Margin', 'Pre-Tax Income Margin', 'Net Operating Profit After Tax']
        dfData = self.__table(temp, ratioList)
        
        return temp, dfData

    def __returns(self, writer):
        temp = {}
        for co_id in self.__co_ids:
            temFP = self.__statements[co_id]['financialPosition']
            temIS = self.__statements[co_id]['incomeStatement']
            if type(temFP) == str or type(temIS) == str: 
                temp[co_id] = [temFP, temIS]
            else:
                afterTax = [ a-b for a, b in zip(temIS['稅前淨利'], temIS['所得稅費用']) ]
                temp[co_id] = {
                    'ROA': [ round(a/b, 3) for a, b in zip(afterTax, temFP['資產總額']) ][:-1],
                    'ROE': [ round(a/b, 3) for a, b in zip(afterTax, temFP['股東權益總額']) ][:-1],
                }

        ratioList = ['ROA', 'ROE']
        dfData = self.__table(temp, ratioList)
        
        return temp, dfData
        
    def __table(self, data, ratioList):
        matrixData = []
        for i in ratioList:
            matrixData.append([i, 2019, 2018, 2017, 2016, 2015])
            for j in data.keys():
                tem = [j]
                if type(data[j]) == list:
                    tem.extend([math.nan]*5)
                else:
                    tem.extend(data[j][i])
                matrixData.append(tem)
        return pd.DataFrame( np.array(matrixData) )        

    def __fig(self, dfData, sn):
        nfig = dfData.shape[0]//7
        plt.figure( figsize=( 20, 5*((nfig-1)//3 + 1) ) )
        for i in range(1, nfig+1):
            df = dfData.iloc[(i-1)*7:i*7,]

            x = list(df.iloc[0,])[1:][::-1]

            plt.subplot( (nfig-1)//3 + 1, 3, i)
            minn, maxx = 0, 0
            for j in range(1, df.shape[0]):
                label = df.iloc[j, ][0]
                y = [float(k) for k in df.iloc[j, ][1:]][::-1]

                plt.plot(x, y, label = label)

                if j == 1: minn, maxx = min(y), max(y)
                else:
                    if minn > min(y): minn = min(y)
                    if maxx < max(y): maxx = max(y)

            ytick = [ round(minn + (maxx-minn)*(i/9), 3) for i in range(10)]
            plt.yticks(ytick)
            plt.title(list(df.iloc[0,])[0], fontsize = 15)
        plt.legend(loc = 'upper right', bbox_to_anchor = (1.3, 1), fontsize = 'x-large')
        plt.savefig(sn + '.jpg')
        


co_ids = [1101, 1102, 1103, 1104, 1108, 1109]
statements = statements(co_ids)
#pprint(statements.getInfo())
FSARatio = FSA(co_ids, statements.getInfo())
FSARatio.analysis()
