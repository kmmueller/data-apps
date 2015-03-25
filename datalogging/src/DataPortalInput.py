'''
Making a main program to call the routines and build the excel workbook.
3/24/2015
'''
from openpyxl.workbook import Workbook
'''
Take existing FLOW.py module and wrap with loop so as to build a single stock excel with multiple expiration dates...at the same conditions
'''

import logTableGen
import xlLogGen
from InputConfigs import *

ww=Workbook()
for expDate in ExpList:
    tableName=logTableGen.logTableGen(ticker, expDate, lowStrike, step, numStrikes)
    ww=xlLogGen.xlLogGen(ticker, expDate, lowStrike, step, numStrikes,tableName,ww)
    
xlFileName= '%s%d%d%d.xlsx' % (ticker,lowStrike,numStrikes,step)
print(xlFileName)
ww.save(xlFileName)
print('Congratulations - we are done.  Time for beer!')