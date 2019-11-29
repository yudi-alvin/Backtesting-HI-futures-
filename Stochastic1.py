# -*- coding: utf-8 -*-
"""
Created on Tue Nov 12 17:18:11 2019
@author: Yudi Alvin
Required:
pip install xlwt
pip install mplfinance

"""

#import relevant modules
import pandas as pd
import numpy as np
import xlwt
import math
#from pandas_datareader import data
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os as os
import ta as ta
import warnings
warnings.filterwarnings('ignore')
from mpl_finance import candlestick_ohlc

from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()


def analyze_stochastic(q,nperiod,dperiod=3, T=80, U = 20):
    dfl=[]
    datadir = 'HI'
    lof = os.listdir(datadir)
    n = len(lof)
    commission =3.47 # please fill in the commission 3.47
    
    for j in range(n):
        print("day " + str(j+1))
        
        if lof[j].find('.csv') == -1: # just in case there are other non-data files
            continue
        fn = datadir + '//' + lof[j]
        df = pd.read_csv(fn)
        
        ##########qmin data##########
        dataLength = len(df.index)
        tp = df['Price'].array     #take price Price  
        tv = df['Volume'].array     #Volume
        tt = df['DateTime'].array       
        bid = df['Bbid'].array # best bid
        ask = df['Bask'].array #best ask
        bidsize = df['Bbs'].array
        asksize = df['Bas'].array
        intraday_time = pd.to_datetime(tt, format = '%Y-%m-%d %H:%M:%S')
    
        df1 = pd.DataFrame(data = {'Price': tp, 'Volume': tv, 'Bid': bid, 'Ask': ask, \
                                   'BidS': bidsize, 'AskS': asksize}, \
                                   index = intraday_time)
                                   
        dfqmin = df1.resample(q).last()        
        dfqmin.dropna(axis = 0, how = 'all', inplace = True)
        #print(dfqmin)
        S = pd.Series(tp, index = intraday_time)
        ohlc = S.resample(q).ohlc()
        print("ohcl")
        print(ohlc)
        
        #############Stochastic############
        for m in range (len(nperiod)):
            df=ohlc
            print(nperiod[m])
            k = ta.momentum.stoch(df["high"], df["low"], df["close"], n=nperiod[m], fillna=False)
            d = ta.stoch_signal(df["high"], df["low"], df["close"], n=nperiod[m],d_n= dperiod, fillna=False)
            df['%K'] = k
            df['%D'] = d
            #k(fast) cross d(slow)
            #pd.set_option('display.max_rows', None)
            pd.set_option('display.max_columns', None)
            df['Sell Entry'] = ((df['%K'] < df['%D']) & (df['%K'].shift(1) > df['%D'].shift(1))) & (df['%D'] > T)
            df['Buy Entry'] = ((df['%K'] > df['%D']) & (df['%K'].shift(1) < df['%D'].shift(1))) & (df['%D'] < U)
      
            
            #remove duplicate signal no need because we use ffill
#            sellStatus =False
#            buyStatus =False
#            for row in df.index:
#                if df.loc[row]['Sell Entry'] == True and sellStatus == False:
#                    sellStatus = True
#                    buyStatus = False
#                elif df.loc[row]['Sell Entry'] == True and sellStatus == True:
#                    df.set_value(row, 'Sell Entry', False) 
#                if df.loc[row]['Buy Entry'] == True and buyStatus == False:
#                    sellStatus = False
#                    buyStatus = True
#                elif df.loc[row]['Buy Entry'] == True and buyStatus == True:
#                    df.set_value(row, 'Buy Entry', False) 
            #print("modified")
            

            #Create empty "Position" column
            df['Position'] = np.nan 
            #Set position to -1 for sell signals
            df.loc[df['Sell Entry'],'Position'] = -1 
            #Set position to -1 for buy signals
            df.loc[df['Buy Entry'],'Position'] = 1 
            #Set starting position to flat (i.e. 0)
            df['Position'].iloc[0] = 0 
			
			#limit to one position open at a time
            #Forward fill the position column to show holding of positions through time (remove buy buy/ sell sell to buy/sell only)
            df['Position'] = df['Position'].fillna(method='ffill')
            df["diff"] =df['Position'].diff()
            df["diff"].iloc[0] =0
			
			#drop the zero region
            futures_signals = df.drop(df[df["diff"] == 0].index)
            print(futures_signals)
            
            if futures_signals.count()[1] > 2:
                futures_signals.sort_index(inplace = True)
                
                ###############long position######################
                orig=futures_signals
                i = futures_signals.index
                if futures_signals.Position[0] == -1.0:
                   futures_signals = futures_signals.drop([i[0], i[-1]])
        
                i = futures_signals.index
                if futures_signals.Position[-1] == 1.0:
                   futures_signals = futures_signals.drop(i[-1])

                
                # See the profitability of long trades
                futures_long_profits=pd.DataFrame()
                futures_long_profits["Price"] =  futures_signals.loc[(futures_signals["Position"] == 1.0),"close"]
                dfpf= futures_signals["close"] -futures_signals["close"].shift(1)
                s = pd.Series(dfpf.loc[(futures_signals["Position"] == -1.0)].values)
                s.index = futures_long_profits.index[:len(s)]
                futures_long_profits.loc[:,'Profit'] = s
                futures_long_profits.Profit= futures_long_profits.Profit.fillna(0)
                s =  pd.Series(futures_signals.loc[(futures_signals["Position"] == -1.0)].index)
                s.index = futures_long_profits.index[:len(s)]
                futures_long_profits.loc[:,'End Date'] = s
               
                futures_long_profits['Profit_After_Commission']= futures_long_profits['Profit']-commission
                print("long positions")
                print(futures_long_profits)
                print('\n')
                
                

                ###############Short position######################
				futures_signals=orig
                i = futures_signals.index
                if futures_signals.Position[0] == 1.0:
                   futures_signals = futures_signals.drop([i[0], i[-1]])
        
                i = futures_signals.index
                if futures_signals.Position[-1] == -1.0:
                   futures_signals = futures_signals.drop(i[-1])
				   
                futures_short_profits=pd.DataFrame()
                futures_short_profits["Price"] =  futures_signals.loc[(futures_signals["Position"] == -1.0),"close"]
                dfpf= futures_signals["close"].shift(1)- futures_signals["close"]
                s = pd.Series(dfpf.loc[(futures_signals["Position"] == 1.0)].values)
                s.index = futures_short_profits.index[:len(s)]
                futures_short_profits.loc[:,'Profit'] = s
                futures_short_profits.Profit= futures_short_profits.Profit.fillna(0)
                s =  pd.Series(futures_signals.loc[(futures_signals["Position"] == 1.0)].index)
                s.index = futures_short_profits.index[:len(s)]
                futures_short_profits.loc[:,'End Date'] = s
                futures_short_profits['Profit_After_Commission']= futures_short_profits['Profit']-commission
                print(futures_short_profits)
                
                total = futures_long_profits.Profit_After_Commission.sum() + futures_short_profits.Profit_After_Commission.sum()
                print('Total Profit After Commission= %0.2f' % ( total))
                
				#concatenate the long and short dataframe together
                if(not futures_long_profits.empty and not futures_short_profits.empty):
                    df=pd.concat([futures_long_profits,futures_short_profits])
                elif(not futures_long_profits.empty and futures_short_profits.empty):
                    df = futures_long_profits
                else:
                    df = futures_short_profits
                print("j is here:" + str(j))
			#concatenate each period to the list of df frame
            else:
                df= pd.DataFrame()
            if j == 0:
                if (not df.empty):
                    dfl.append(df)
                else:
                    print("Append empty df")
                    dfl.append(pd.DataFrame(columns=['Price', 'Profit', 'End Date', 'Profit_After_Commission']))
            else:
                if (not df.empty):

                    dfl[m] = pd.concat([dfl[m],df])
            
   
    return dfl
	
#generate a list of n periods
#Note: dont generate long list of nperiod as the pd.concat has limited memory
nperiod=[]
for i in range(5,20):
    nperiod.append(i)

##########################INPUT######################
q='5min'                                            #   
nperiod=[16] #run with one period                    #
dperiod=3                                             #
result = analyze_stochastic(q,nperiod,dperiod)      #
filename = '5minStochastic.xls'                      #
#####################################################    
#print(result)
#e1 = result[0].Profit.cumsum()
#plt.plot(e1)
#write to excel
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')

ws.write(0, 1, "nperiod")
ws.write(0, 2, "Pwinning")
ws.write(0, 3, "Risk to reward")
ws.write(0, 4, "t-test")
ws.write(0, 5, "Profit")

ws.write(0, 7, "nperiod")
ws.write(0, 8, "Pwinning")
ws.write(0, 9, "Risk to reward")
ws.write(0, 10, "t-test")
ws.write(0, 11, "Profit After Commission")

ws.write(0, 12, "dperiod")
ws.write(1, 12, dperiod)

for i in range(len(result)):
    df = result[i]
    n =df.count()[1]
    nwin = df[df['Profit']>0].count()[1]
    nloss = df[df['Profit']<0].count()[1]
    pwin= round((nwin/n),2)
    
    #Risk to reward
    if nwin==0:
        rtr=0
    else:
        profitSum=df[df['Profit']>0].sum()
        lossSum= df[df['Profit']<0].sum()
        avgProfit=profitSum[1]/nwin
        avgLoss= lossSum[1]/nloss
        rtr=round((avgProfit/-avgLoss),5)
        
    #t-statistics
    
    avgPnL = df['Profit'].mean()
    std= np.std(df['Profit'])
    if std ==0:
        t =0
    else:
        t=round((math.sqrt(n) * avgPnL / std),5)
    
    
    profit=df['Profit'].sum()
    
    #writing to excel
    #use pip install xlwt

        #row,column,value
    ws.write(i+1, 1, str(i+10))
    ws.write(i+1, 2, str(pwin))
    ws.write(i+1, 3, str(rtr))
    ws.write(i+1, 4, str(t))
    ws.write(i+1, 5, str(profit))
    
    df = result[i]
    n =df.count()[1]
    nwin = df[df['Profit_After_Commission']>0].count()[1]
    nloss = df[df['Profit_After_Commission']<0].count()[1]
    pwin= round((nwin/n),2)
    
    #Risk to reward
    if nwin==0:
        rtr=0
    else:
        profitSum=df[df['Profit_After_Commission']>0].sum()
        lossSum= df[df['Profit_After_Commission']<0].sum()
        avgProfit=profitSum[1]/nwin
        avgLoss= lossSum[1]/nloss
        rtr=round((avgProfit/-avgLoss),5)
        
    #t-statistics
    
    avgPnL = df['Profit_After_Commission'].mean()
    std= np.std(df['Profit_After_Commission'])
    if std ==0:
        t =0
    else:
        t=round((math.sqrt(n) * avgPnL / std),5)
    
    profit=df['Profit_After_Commission'].sum()
    
    #writing to excel
    #use pip install xlwt

        #row,column,value
    ws.write(i+1, 7, str(i+10))
    ws.write(i+1, 8, str(pwin))
    ws.write(i+1, 9, str(rtr))
    ws.write(i+1, 10, str(t))
    ws.write(i+1, 11, str(profit))
 
wb.save(filename) 

