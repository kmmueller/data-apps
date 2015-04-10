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

#ConfigFile = input('Enter name for the config file ')

def getVarFromFile(filename):
    import imp
    f = open(filename,'r')
    global data
    data = imp.load_source('data', '', f)
    f.close()

getVarFromFile('c:\InputValues.txt')

ww=Workbook()
for expDate in data.ExpList:
    tableName=logTableGen.logTableGen(data.ticker, expDate, data.lowStrike, data.step, data.numStrikes)
    ww=xlLogGen.xlLogGen(data.ticker, expDate, data.lowStrike, data.step, data.numStrikes,tableName,ww)
    
xlFileName= '%s%d%d%d.xlsx' % (data.ticker,data.lowStrike,data.numStrikes,data.step)
print(xlFileName)
ww.save(xlFileName)
print('Congratulations - we are done.  Time for beer!')
