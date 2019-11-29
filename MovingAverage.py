"""
pip install xlwt
pip install mplfinance
"""

from __future__ import (absolute_import, division, print_function,
                        unicode_literals)
import xlwt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')
import os as os
import math
from datetime import datetime
import ta as ta
def moving_average(x, n, type='simple'):
   # compute an n period moving average.
   # type is 'simple' | 'exponential'
   x = np.asarray(x)
   if type == 'simple':
      weights = np.ones(n)
   else:
      weights = np.exp(np.linspace(-1., 0., n))
   weights /= weights.sum()

   a = np.convolve(x, weights, mode='full')[:len(x)]
   print(a)
   a[:n] = a[n]
   return a

######################################################################################### 
equity1=[]
equity2=[]

def analyze_moving_average(l,q):
    
    dfl=[]
    datadir = 'HI'
    lof = os.listdir(datadir)
    n = len(lof)
    commission =3.47 # please fill in the commission 3.47
    for k in range(n):
#        if 1<k<200:
#            continue

        ###################Reading the file###################
        if lof[k].find('.csv') == -1: # just in case there are other non-data files
            continue
        fn = datadir + '//' + lof[k]
        df = pd.read_csv(fn)
        
        #################Resampling############################
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
      
        pd.set_option('display.max_columns', None)
        
        ################Finding optimal moving average############
        for j in range(len(l)):
            print(l[j])
            futures = dfqmin
            print( "day " + str(k+1))
            fast, slow = l[j]
            futures['fast'] = moving_average(futures.Price, fast, type='simple')
            futures['slow'] = moving_average(futures.Price, slow, type='simple')
#            futures['fast'] = ta.ema_indicator(futures.Price,n= fast, fillna=False)
#            futures['slow'] = ta.ema_indicator(futures.Price,n= slow, fillna=False)
            
            futures['fast-slow'] = futures['fast'] - futures['slow']
            futures["Regime"] = np.where(futures['fast-slow'] > 0, 1, 0) 
            futures["Regime"] = np.where(futures['fast-slow'] <= 0, -1, futures["Regime"])
            
            #Generating Signal
            print(futures)
            futures = dfqmin
            futures = futures.drop(futures[futures["Regime"] == 0 ].index)
            regime_orig=futures.iloc[-1]["Regime"]
            futures.Regime[-1] = 0#to take out the last period assuming we dont trade in the last period,to ensure that all trades close out,
            signal = futures["Regime"] - futures["Regime"].shift(1) #shift one row down so its basically become 2nd -1st
            signal[0] = 0
            futures["Signal"]= np.sign(signal)                    #return the sign
            futures_signals = futures.drop(futures[futures["Signal"] == 0 ].index)
            print(futures_signals)
			
            #if futures_signals.count()[1] <=2:
            #    continue
            #############Taking long/Short position##############3
            elif futures_signals.count()[1] > 2:
                futures_signals.sort_index(inplace = True)
                ###############long position######################
                orig=futures_signals
                i = futures_signals.index
                if futures_signals.Signal[0] == -1.0:
                   futures_signals = futures_signals.drop([i[0], i[-1]])
            
                i = futures_signals.index
                if futures_signals.Signal[-1] == 1.0:
                   futures_signals = futures_signals.drop(i[-1])
                
                #print(futures_signals)
                
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
                print("Long position")
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
                print("Short")
                print(futures_short_profits)
                total= futures_long_profits.Profit.sum() + futures_short_profits.Profit.sum()
                equity1.append(sum(equity1) + total)
                total = futures_long_profits.Profit_After_Commission.sum() + futures_short_profits.Profit_After_Commission.sum()
                print('Total Profit After Commission= %0.2f' % ( total))
                equity2.append(sum(equity2) + total)


                ####### populating profits from long and short position
                if(not futures_long_profits.empty and not futures_short_profits.empty):
                    df=pd.concat([futures_long_profits,futures_short_profits])
                elif(not futures_long_profits.empty and futures_short_profits.empty):
                    df = futures_long_profits
                else:
                    df = futures_short_profits
                    
            else:
                df= pd.DataFrame()   
            if k ==0:
                if (not df.empty):
                    dfl.append(df)
                else:
                    dfl.append(pd.DataFrame(columns=['Price', 'Profit', 'End Date', 'Profit_After_Commission']))
            else:
                if (not df.empty):
                    print(l[j])
                    dfl[j] = pd.concat([dfl[j],df])
                
            

    return dfl
    
########################################Change parameter q (optimize it)####################################33
######################################## Remember to change the file name in line 243##########
#generate a list of n periods
#Note: dont generate long list of nperiod as the pd.concat has limited memory
#Note: large periods with large qmin shorten the trading period which might make the moving average function failed- havent add in exception handler
l=[]
for i in range(5,31,2):
    for j in range(9,41,2):
        if ((j,i) not in l) and (i !=j):
            l.append((i,j))

#l=[(5,9),(5,11),(5,13),(5,15),(5,17),(5,19),(5,21),(7,12),(7,15),(7,19),(7,26),(9,13),(9,17),(9,21),(9,26),(11,15),(11,19),(11,27)] #optimize parameter for professor to run


##########################INPUT######################
l=[(11,15)] #to run only one fast-slow 
q = '1min'                                          #
result= analyze_moving_average(l,q)                 #
print(result)
#result.to_csv('9minMAresult.csv')                               #
filename="1minMOVAVGequity.xls"                          #
####################################################3
#write to excel
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')

ws.write(0, 1, "fast-slow")
ws.write(0, 2, "Pwinning")
ws.write(0, 3, "Risk to reward")
ws.write(0, 4, "t-test")
ws.write(0, 5, "Profit")

ws.write(0, 7, "fast-slow")
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
    elif nloss == 0:
        rtr =1
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
    ws.write(i+1, 1, str(l[i]))
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
    elif nloss == 0:
        rtr =1
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
    ws.write(i+1, 7, str(l[i]))
    ws.write(i+1, 8, str(pwin))
    ws.write(i+1, 9, str(rtr))
    ws.write(i+1, 10, str(t))
    ws.write(i+1, 11, str(profit))
 
wb.save(filename) 

