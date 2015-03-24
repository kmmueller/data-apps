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
loopDone="no"
ww=Workbook()
while (loopDone=="no"):
    ticker = input('Enter stock ticker: ')
    expDate = input('Enter the expiration date: ')
    lowStrike = float(input('Enter the low strike: '))
    step = float(input('Enter the step size between strikes: '))
    numStrikes = float(input('Enter the number of strikes: '))
    tableName=logTableGen.logTableGen(ticker, expDate, lowStrike, step, numStrikes)
    ww=xlLogGen.xlLogGen(ticker, expDate, lowStrike, step, numStrikes,tableName,ww)
    #print('The table created is ', tableName)
    loopDone= input('Are you finished? Please enter yes or no. ')
    
xlFileName= input('Enter name for your new excel logging file: ')

print(xlFileName)

ww.save(xlFileName)
print('Congratulations - we are done. ')