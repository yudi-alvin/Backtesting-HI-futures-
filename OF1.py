# -*- coding: utf-8 -*-
"""
Created on Tue Nov 12 11:23:16 2019
@author: Yudi Alvin
Required:
pip install xlwt
pip install mplfinance

"""



import numpy as np
import pandas as pd
import xlwt
import math
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os as os
import ta as ta
import warnings
warnings.filterwarnings('ignore')
from mpl_finance import candlestick_ohlc

from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()


def analyze_OF(q,T1=50,T2=500, U1=-50, U2=-500):
    dfl=[]
    datadir = 'HI'
    lof = os.listdir(datadir)
    n = len(lof)
    commission =3.47 # please fill in the commission 3.47
    
    for k in range(n):
        print("Day " + str(k))
        if lof[k].find('.csv') == -1: # just in case there are other non-data files
            continue
        fn = datadir + '//' + lof[k]
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
        print(dfqmin)
        
        ######ORDER FLOW#######
        of = np.zeros(dataLength, dtype=int)
        for i in range(dataLength):
            mp = 0.5*(bid[i] + ask[i])
            if tp[i] > mp:
                of[i] = tv[i]
            elif tp[i] < mp:
                of[i] = -tv[i]
    
        df_of = pd.DataFrame(data = np.transpose(of), columns = ['Order Flow'], \
                            index = intraday_time) 
    
        ofqmin = df_of.resample(q).sum()
        ofqmin.dropna(axis = 0, how = 'all', inplace = True)
        #print(ofqmin)
        dfqmin['OF'] = ofqmin['Order Flow']
        dfqmin["Regime"] = np.where(dfqmin['OF'] > T1, 1, 0) 
        dfqmin["Regime"] = np.where(dfqmin['OF'] > T2, -1, dfqmin["Regime"]) 
        dfqmin["Regime"] = np.where(dfqmin['OF'] < U1, -1, dfqmin["Regime"])
        dfqmin["Regime"] = np.where(dfqmin['OF'] < U2, 1, dfqmin["Regime"])

		#drop the 0 region 
        futures = dfqmin
        futures = futures.drop(futures[futures["Regime"] == 0 ].index)
        print(futures)
		
        regime_orig=futures.iloc[-1]["Regime"]
        futures.Regime[-1] = 0#to take out the last period assuming we dont trade in the last period,to ensure that all trades close out,
        signal = futures["Regime"] - futures["Regime"].shift(1) #shift one row down so its basically become 2nd -1st
        signal[0] = 0
        futures["Signal"]= np.sign(signal)                    #return the sign
        futures_signals = futures.drop(futures[futures["Signal"] == 0 ].index)
        print(futures_signals)
        if futures_signals.count()[1] > 2:
            futures_signals.sort_index(inplace = True)
            
            ###############long position######################
            orig=futures_signals
            i = futures_signals.index
            if futures_signals.Signal[0] == -1.0:
               futures_signals = futures_signals.drop([i[0], i[-1]])
        
            i = futures_signals.index
            if futures_signals.Signal[-1] == 1.0:
               futures_signals = futures_signals.drop(i[-1])

            # See the profitability of long trades
            futures_long_profits=pd.DataFrame()
            futures_long_profits["Price"] =  futures_signals.loc[(futures_signals["Signal"] == 1.0),"Price"]
            print(len(futures_long_profits))
            print(futures_long_profits)
            dfpf= futures_signals["Price"] -futures_signals["Price"].shift(1)
            s = pd.Series(dfpf.loc[(futures_signals["Signal"] == -1.0)].values)
            print(s)
            s.index = futures_long_profits.index[:len(s)]
            futures_long_profits.loc[:,'Profit'] = s
            futures_long_profits.Profit= futures_long_profits.Profit.fillna(0)
            s =  pd.Series(futures_signals.loc[(futures_signals["Signal"] == -1.0)].index)
            s.index = futures_long_profits.index[:len(s)]
            futures_long_profits.loc[:,'End Date'] = s
           
            futures_long_profits['Profit_After_Commission']= futures_long_profits['Profit']-commission
            print(futures_long_profits)
            print('\n')
            futures_signals=orig
            #print(futures_signals)
			
            ###############Short position######################
            i = futures_signals.index
            if futures_signals.Signal[0] == 1.0:
               futures_signals = futures_signals.drop([i[0], i[-1]])
        
            i = futures_signals.index
            if futures_signals.Signal[-1] == -1.0:
               futures_signals = futures_signals.drop(i[-1])
        
            futures_short_profits=pd.DataFrame()
            futures_short_profits["Price"] =  futures_signals.loc[(futures_signals["Signal"] == -1.0),"Price"]
            print(len(futures_short_profits))
            dfpf= futures_signals["Price"].shift(1)- futures_signals["Price"]
            s = pd.Series(dfpf.loc[(futures_signals["Signal"] == 1.0)].values)
            s.index = futures_short_profits.index[:len(s)]
            futures_short_profits.loc[:,'Profit'] = s
            futures_short_profits.Profit= futures_short_profits.Profit.fillna(0)
            s =  pd.Series(futures_signals.loc[(futures_signals["Signal"] == 1.0)].index)
            s.index = futures_short_profits.index[:len(s)]
            futures_short_profits.loc[:,'End Date'] = s
            futures_short_profits['Profit_After_Commission']= futures_short_profits['Profit']-commission
            print(futures_short_profits)
            
            total = futures_long_profits.Profit_After_Commission.sum() + futures_short_profits.Profit_After_Commission.sum()
            print('Total Profit After Commission= %0.2f' % ( total))
			
			#concatenate the long and short dataframe togethe
            if(not futures_long_profits.empty and not futures_short_profits.empty):
                df=pd.concat([futures_long_profits,futures_short_profits])
            elif(not futures_long_profits.empty and futures_short_profits.empty):
                df = futures_long_profits
            else:
                df = futures_short_profits
            print(df)
		#concatenate each period to the list of df frame
		else:
                df= pd.DataFrame()
		if k == 0:
			if (not df.empty):
				dfl.append(df)
			else:
				dfl.append(pd.DataFrame(columns=['Price', 'Profit', 'End Date', 'Profit_After_Commission']))
		else:
			if (not df.empty):
				dfl[0] = pd.concat([dfl[0],df])
				#only one at a time
				
    return dfl

##########################INPUT######################
"Optimize q, T1,T2,U1,U2""" 
q='8min'                                            #                                        #                                         #
result = analyze_OF(q,T1=70,T2=500, U1=-70, U2=-500)#
filename = '8minq.xls'                              #
#####################################################    
print(result)
#e1 = result[0].Profit.cumsum()
#plt.plot(e1)
        

#write to excel
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')

ws.write(0, 1, "qmin")
ws.write(0, 2, "Pwinning")
ws.write(0, 3, "Risk to reward")
ws.write(0, 4, "t-test")
ws.write(0, 5, "Profit")

ws.write(0, 7, "qmin")
ws.write(0, 8, "Pwinning")
ws.write(0, 9, "Risk to reward")
ws.write(0, 10, "t-test")
ws.write(0, 11, "Profit After Commission")


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
    ws.write(i+1, 1, q)
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
    t=round((math.sqrt(n) * avgPnL / std),5)
    
    profit=df['Profit_After_Commission'].sum()
    
    #writing to excel
    #use pip install xlwt

        #row,column,value
    ws.write(i+1, 7, q)
    ws.write(i+1, 8, str(pwin))
    ws.write(i+1, 9, str(rtr))
    ws.write(i+1, 10, str(t))
    ws.write(i+1, 11, str(profit))
 
wb.save(filename) 




