'''
Created on Oct 4, 2014
Modified 2/6/2015 to include new quantities to track - ask and bid sizes, volumes
and open interest for calls and puts.

@author: mueller
'''
#import pandas as pd
from sqlalchemy import create_engine
import re
import numpy as np
from matplotlib.delaunay.testfuncs import steep
#from scipy.signal.ltisys import step2
#
def logTableGen(ticker, expDate, lowStrike, step, numStrikes):
    engine = create_engine("mysql+mysqlconnector://root:admin@localhost/floatbook")
    
    
    # Taking the variables to generate the database table for logging as input.  We need the following:
    # ticker - ticker for the stock
    # expDate - expiration date for the options
    # lowStrike - lowest strike to begin logging data
    # step - step size between strikes
    # numStrikes - number of strikes
    
    # Build an array of the strike prices based on low, step and number of strikes
    # Get the max price based on steps, low, number of strikes
    maxStrike=(lowStrike+(numStrikes)*step)
    print('Max Strike=',maxStrike)
    strikes=np.arange(lowStrike,maxStrike,step) #numpy arange - allows for float values for min, max, and step
    
    
    print('Strikes are: ',strikes)
    
    # Now build the names of the options as a single list to build database table
    # we need the strike price, and 
    # we need the following values for each strike:
    # Calls: ask, bid, delta, theta, gamma, vega, rho
    # Puts: ask, bid, delta, theta, gamma, vega, rho
    dataBaseColumns=[]
    callName=ticker+expDate+'C'
    putName=ticker+expDate+'P'
    
    print('Call name=',callName)
    print('Put name=',putName)
    
    
    dataTypeFLOAT=' float DEFAULT NULL, ' # need this one as it forms past of the query to define the mysql table
    dataBaseColumnNames=[ticker, dataTypeFLOAT]
    for x in (strikes):
        m=re.compile('\d+')
        n=m.findall(str(x))
        if (n[1]=='0'):
            x=int(x)
            print('This is the new x:', x)
        else:
            x=n[0]+"p"+n[1]
            print('This is new x with a decimal')
        #print(m.group(0))
        callAsk=callName+str(x)+"_ask"+dataTypeFLOAT
        callAskSize=callName+str(x)+"_asksize"+dataTypeFLOAT
        callBid=callName+str(x)+"_bid"+dataTypeFLOAT
        callBidSize=callName+str(x)+"_bidsize"+dataTypeFLOAT
        callVolume=callName+str(x)+"_volume"+dataTypeFLOAT
        callOpenInt=callName+str(x)+"_openint"+dataTypeFLOAT
        callDelta=callName+str(x)+"_delta"+dataTypeFLOAT
        callTheta=callName+str(x)+"_theta"+dataTypeFLOAT
        callGamma=callName+str(x)+"_gamma"+dataTypeFLOAT
        callVega=callName+str(x)+"_vega"+dataTypeFLOAT
        callRho=callName+str(x)+"_rho"+dataTypeFLOAT
        putAsk=putName+str(x)+"_ask"+dataTypeFLOAT
        putAskSize=putName+str(x)+"_asksize"+dataTypeFLOAT
        putBid=putName+str(x)+"_bid"+dataTypeFLOAT
        putBidSize=putName+str(x)+"_bidsize"+dataTypeFLOAT
        putVolume=putName+str(x)+"_volume"+dataTypeFLOAT
        putOpenInt=putName+str(x)+"_openint"+dataTypeFLOAT
        putDelta=putName+str(x)+"_delta"+dataTypeFLOAT
        putTheta=putName+str(x)+"_theta"+dataTypeFLOAT
        putGamma=putName+str(x)+"_gamma"+dataTypeFLOAT
        putVega=putName+str(x)+"_vega"+dataTypeFLOAT
        putRho=putName+str(x)+"_rho"+dataTypeFLOAT
    
        dataBaseColumns=(callAsk,callAskSize,callBid,callBidSize,callVolume,callOpenInt,callDelta,callTheta,callGamma,callVega,callRho,putAsk,putAskSize,putBid,putBidSize,putVolume,putOpenInt,putDelta,putTheta,putGamma,putVega,putRho)
        for y in (dataBaseColumns):
            dataBaseColumnNames.append(y)
    print("All call & put related names:",dataBaseColumnNames)
    
    # Now, we need to make the SQL query to build the database table.  
    # We already have the full set of column names, but we want to add two special columns:
    # id - defined auto incrementing unique identifier
    # ts - timestamp (SQL timestamp)
    # 
    # Since we're allowing more flexibility in our database, we need to change the naming convention
    # for out table names.  This is now Ticker+Expiration+Lowstrike+step+#of steps
    
    engine=create_engine("mysql+mysqlconnector://root:admin@localhost/floatbook")  # +mysqlconnector = use the mysql ODBC connector instead of the native one
    # set name of data table
    m=re.compile('\d+')
    n=m.findall(str(lowStrike))
    if (n[1]=='0'):    # this loop gets rid of the decimal point in the name for data table
        nlowStrike=int(lowStrike)  #if I don't get rid of it, mysql barfs on the table name
        print(nlowStrike)
    else:
        nlowStrike=n[0]+"p"+n[1]
        print(lowStrike) 
        
    stepTest=m.findall(str(step))
    if (stepTest[1]=='0'):
        nstep=int(step)
        print(nstep)
    else:
        nstep=stepTest[0]+"p"+stepTest[1]
        print(nstep)
    dataTableName=ticker+expDate+str(nlowStrike)+str(nstep)+str(int(numStrikes))    
    print('Data table name is:',dataTableName)
    
    idAutoIncrement = 'id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY, '
    timeStamp = 'ts timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP)'
    
    tableCreate='create table '+dataTableName+' ('+idAutoIncrement
    for z in dataBaseColumnNames:
        tableCreate=tableCreate+z
    tableCreate=tableCreate+timeStamp
    print(tableCreate)
    
    engine.execute(tableCreate)
    return(dataTableName)
# code to test that the function works as expected:
#===============================================================================
#ticker = 'GPRO'
#expDate = '150220'
#lowStrike = float(57.5)
#step = float(2.5)
#numStrikes = float(7)
#print("Calling function...")
#logTableGen(ticker, expDate, lowStrike,step,numStrikes)
#===============================================================================


