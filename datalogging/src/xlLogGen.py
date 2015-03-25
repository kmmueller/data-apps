'''
Created on Mar 8, 2015
Extracting the function from FloatLogGen2 to make it standalone and callable from the flow module.
@author: dev-machine
'''
import re
import numpy as np
from openpyxl import Workbook

def xlLogGen(ticker, expDate, lowStrike, step, numStrikes,tableName,wb):
    
    # Build an array of the strike prices based on low, step and number of strikes
    # Get the max price based on steps, low, number of strikes
    maxStrike=(lowStrike+(numStrikes)*step)
    print('Max Strike=',maxStrike)
    strikes=np.arange(lowStrike,maxStrike,step) #numpy arange - allows for float values for min, max, and step

    callName="."+ticker+expDate+'C'
    putName="."+ticker+expDate+'P'
    #wb=Workbook()
    ws=wb.active
    dataItemNames=["=RTD(\"tos.rtd\", ,\"LAST\", \""+ticker+"\")"]
    nameTable=tableName
# now write each set of data with a header incluuding the floatbook table name and number of strikes.  This way I 
# can simplify what I need to write in Excel to log to mysql.  Almost there.
# 3/12/2015
    
    ws.append([nameTable, numStrikes])
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
        callAskSize="=RTD(\"tos.rtd\", , \"ASK_SIZE\", \""+callName+str(x)+"\")"
        callBid="=RTD(\"tos.rtd\", , \"BID\", \""+callName+str(x)+"\")"
        callBidSize="=RTD(\"tos.rtd\", , \"BID_SIZE\", \""+callName+str(x)+"\")"
        callVolume="=RTD(\"tos.rtd\", , \"VOLUME\", \""+callName+str(x)+"\")"
        callOpenInt="=RTD(\"tos.rtd\", , \"OPEN_INT\", \""+callName+str(x)+"\")"
        callDelta="=RTD(\"tos.rtd\", , \"DELTA\", \""+callName+str(x)+"\")"
        callTheta="=RTD(\"tos.rtd\", , \"THETA\", \""+callName+str(x)+"\")"
        callGamma="=RTD(\"tos.rtd\", , \"GAMMA\", \""+callName+str(x)+"\")"
        callVega="=RTD(\"tos.rtd\", , \"VEGA\", \""+callName+str(x)+"\")"
        callRho="=RTD(\"tos.rtd\", , \"RHO\", \""+callName+str(x)+"\")"
        putAsk="=RTD(\"tos.rtd\", , \"ASK\", \""+putName+str(x)+"\")"
        putAskSize="=RTD(\"tos.rtd\", , \"ASK_SIZE\", \""+putName+str(x)+"\")"
        putBid="=RTD(\"tos.rtd\", , \"BID\", \""+putName+str(x)+"\")"
        putBidSize="=RTD(\"tos.rtd\", , \"BID_SIZE\", \""+putName+str(x)+"\")"
        putVolume="=RTD(\"tos.rtd\", , \"VOLUME\", \""+putName+str(x)+"\")"
        putOpenInt="=RTD(\"tos.rtd\", , \"OPEN_INT\", \""+putName+str(x)+"\")"
        putBid="=RTD(\"tos.rtd\", , \"BID\", \""+putName+str(x)+"\")"
        putDelta="=RTD(\"tos.rtd\", , \"DELTA\", \""+putName+str(x)+"\")"
        putTheta="=RTD(\"tos.rtd\", , \"THETA\", \""+putName+str(x)+"\")"
        putGamma="=RTD(\"tos.rtd\", , \"GAMMA\", \""+putName+str(x)+"\")"
        putVega="=RTD(\"tos.rtd\", , \"VEGA\", \""+putName+str(x)+"\")"
        putRho="=RTD(\"tos.rtd\", , \"RHO\", \""+putName+str(x)+"\")"
        xx.append([callAsk,callAskSize,callBid,callBidSize,callVolume,callOpenInt,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putAskSize,putBid,putBidSize,putVolume,putOpenInt,putDelta,putTheta,putGamma,putVega,putRho])
        ws.append([callAsk,callAskSize,callBid,callBidSize,callVolume,callOpenInt,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putAskSize,putBid,putBidSize,putVolume,putOpenInt,putDelta,putTheta,putGamma,putVega,putRho])
        dataItemColumns=(callAsk,callAskSize,callBid,callBidSize,callVolume,callOpenInt,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putAskSize,putBid,putBidSize,putVolume,putOpenInt,putDelta,putTheta,putGamma,putVega,putRho)
        for y in (dataItemColumns):
            dataItemNames.append(y)
    return(wb)
