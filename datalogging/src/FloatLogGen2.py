'''
Created on Jan 2, 2015
This example uses a function call to generate the float logging excel workbook.
For this example I've hardcoded 2 stocks - INTC and GPRO.  I have commented
out the interactive input promtps; those can be used later.  I can theoretically
do any number (up to limits of excel # of rows) in a single sheet.  

Pretty cool!  Go 2015!!
@author: mueller

Updated files and repository on 2/6/2015
'''# The idea here is to turn the working FloatLogGen.py into a function
# and then test whether I can call multiple times to append to a single
# openpyxl workbook object and thereby write out n different stocks.
# Going to start hacking things together using pieces from the old FloatLogGen
# and some unique identifier information from FloatDBGen.py as well
import re
import numpy as np
from openpyxl import Workbook
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
wb=Workbook()
def xlLogGen(ticker, exDate, lowStrike, step, numStrikes,wb):
	
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
		putAskSize="=RTD(\"tos.rtd\", , \"ASK_SIZE\", \""+callName+str(x)+"\")"
		putBid="=RTD(\"tos.rtd\", , \"BID\", \""+callName+str(x)+"\")"
		putBidSize="=RTD(\"tos.rtd\", , \"BID_SIZE\", \""+callName+str(x)+"\")"
		putVolume="=RTD(\"tos.rtd\", , \"VOLUME\", \""+callName+str(x)+"\")"
		putOpenInt="=RTD(\"tos.rtd\", , \"OPEN_INT\", \""+callName+str(x)+"\")"
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


ww=xlLogGen(ticker,expDate,lowStrike,step,numStrikes,wb)
#xx=xlLogGen(ticker,expDate,lowStrike,step,numStrikes,ww)
#xx.save("xx.xlsx")
ticker = 'GPRO'
#input('Enter stock ticker: ')
expDate = '150220'
# input('Enter the expiration date: ')
lowStrike = float(57.5)
#float(input('Enter the low strike: '))
step = float(2.5)
#float(input('Enter the step size between strikes: '))
numStrikes = float(7)
ww=xlLogGen(ticker,expDate,lowStrike,step,numStrikes,ww)
ww.save("Multiple_stock_Example2.xlsx")



