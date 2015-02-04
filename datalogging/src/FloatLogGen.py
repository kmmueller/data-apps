'''
Created on Oct 5, 2014

@author: mueller
'''
# This script will attempt to replicate what the "Generate" function did in Excel:
# set up the excel spreadsheet with the proper references to TOS to grab the data
# Since we're trying to deal with the future, I will use RTD style, rather than DDE
# This is an example of what needs to be generated:
# =RTD("tos.rtd", , "BID", ".KMP141018P92.5")  I chose the example with a .5 included
# to be sure I would generate the cells properly.

import re
import numpy as np
import pandas as pd
from openpyxl import Workbook

# from FloatPyModules.FloatLogGen import dataItemFrame, dataItemColumns
# import openpyxl

# from FloatPyModules.FloatDBGen import callName

# First, generate a list containing the different strikes and types (call, put, bid, ask
# delta, theta, gamma, vega, rho)
#
# Since this routine gets called only once to set up the spreadsheet, I will prompt
# for the same data as in FloatDBGen. In the future there may be some things I can 
# reuse from there, but for now let's see if I can get anything to work.

ticker = 'INTC'
#input('Enter stock ticker: ')
expDate = '150220'
# input('Enter the expiration date: ')
lowStrike = float(30)
#float(input('Enter the low strike: '))
step = float(1)
#float(input('Enter the step size between strikes: '))
numStrikes = float(6)
#float(input('Enter the number of strikes: '))

# Build an array of the strike prices based on low, step and number of strikes
# Get the max price based on steps, low, number of strikes
maxStrike=(lowStrike+(numStrikes)*step)
print('Max Strike=',maxStrike)
strikes=np.arange(lowStrike,maxStrike,step) #numpy arange - allows for float values for min, max, and step

callName="."+ticker+expDate+'C'
putName="."+ticker+expDate+'P'
wb=Workbook()
ws=wb.active
dataItemNames=["=RTD(\"tos.rtd\", ,\"LAST\", \""+ticker+"\")"]
# playing with writing to openpyxl. Here I have setup the workbook object in memory, set the first sheet to active, and
# now am writing the first row as the ticker LAST value.  I think I can use this format to write out multiple stocks to 
# the same page
ws.append(dataItemNames)

xx=[]
for x in (strikes):
    m=re.compile('\d+')
    n=m.findall(str(x))
    if (n[1]=='0'):
        x=int(x)
        print('This is the new x:', x)
    else:
        x=n[0]+"."+n[1]
        print('This is new x with a decimal')
    #print(m.group(0))
    callAsk="=RTD(\"tos.rtd\", , \"ASK\", \""+callName+str(x)+"\")"
    callBid="=RTD(\"tos.rtd\", , \"BID\", \""+callName+str(x)+"\")"
    callDelta="=RTD(\"tos.rtd\", , \"DELTA\", \""+callName+str(x)+"\")"
    callTheta="=RTD(\"tos.rtd\", , \"THETA\", \""+callName+str(x)+"\")"
    callGamma="=RTD(\"tos.rtd\", , \"GAMMA\", \""+callName+str(x)+"\")"
    callVega="=RTD(\"tos.rtd\", , \"VEGA\", \""+callName+str(x)+"\")"
    callRho="=RTD(\"tos.rtd\", , \"RHO\", \""+callName+str(x)+"\")"
    putAsk="=RTD(\"tos.rtd\", , \"ASK\", \""+putName+str(x)+"\")"
    putBid="=RTD(\"tos.rtd\", , \"BID\", \""+putName+str(x)+"\")"
    putDelta="=RTD(\"tos.rtd\", , \"DELTA\", \""+putName+str(x)+"\")"
    putTheta="=RTD(\"tos.rtd\", , \"THETA\", \""+putName+str(x)+"\")"
    putGamma="=RTD(\"tos.rtd\", , \"GAMMA\", \""+putName+str(x)+"\")"
    putVega="=RTD(\"tos.rtd\", , \"VEGA\", \""+putName+str(x)+"\")"
    putRho="=RTD(\"tos.rtd\", , \"RHO\", \""+putName+str(x)+"\")"
    xx.append([callAsk,callBid,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putBid,putDelta,putTheta,putGamma,putVega,putRho])
    ws.append([callAsk,callBid,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putBid,putDelta,putTheta,putGamma,putVega,putRho])
    dataItemColumns=(callAsk,callBid,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putBid,putDelta,putTheta,putGamma,putVega,putRho)
    for y in (dataItemColumns):
        dataItemNames.append(y)
xx=np.asarray(xx)
print("This is xx:")
print(xx)
print(xx.shape)
xx=pd.DataFrame(xx)
xx.index=strikes
dfColumnNames=("callAsk","callBid","callDelta","callTheta","callGamma","callVEga","callRho","putAsk","putBid","putDelta","putTheta","putVega","putGamma","putRho")

xx.columns=dfColumnNames

print("This is xx as a dataframe:")
#print(xx)
#yy=xx
#print(yy)
#zz=[xx,yy]
#zz=yy.append(xx)

#xx=pd.DataFrame(xx,index=strikes,columns=dataItemNames)

#print(xx)
#print("All call & put related rtd expressions:",dataItemNames)
#print(dataItemColumns)
#dIColumns=np.array(dataItemColumns)
#dataItemArray=np.array(dataItemNames)
#dataItemArray.columns=dIColumns


#print(dataItemArray)
#dataItemFrame=pd.DataFrame(dataItemArray)
#print(dataItemFrame)

#testFrame=pd.DataFrame(dataItemNames,index=strikes,columns=dataItemColumns)
#print(testFrame)



# Now for the magic - can I export this dataframe right to excel? And if so, what happens?
#dataItemFrame.to_excel('C:\Users\mueller\Documents\test1.xlsx', sheet_name='Sheet1')
# xls=pd.ExcelWriter(r'C:\Users\mueller\Documents\test1.xlsx')
#zz.to_excel(r'C:\Users\mueller\Documents\zz.xlsx', startrow=1,startcol=1,sheet_name='Sheet1')
wb.save("Example.xlsx")
