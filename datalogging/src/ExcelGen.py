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
#from InputConfigs import *

ww=Workbook()
ExpList = ['150417','150424','150501','150508','150515','150522','150529','150619','150630','150717','150918','150930','151219','151231','160115','160318','160617','160916','161216','170120','170317','171215']
#ExpList = ['150327','150331','150402','150410','150417','150424','150501','150515','150619','150630','150717','150918','150930','151219','151231','160115']
ticker = 'SPY'
lowStrike = float(200)
step = float(0.5)
numStrikes = float(40)  # 22 columns generated.  40 numStrikes is OK
suffix = '_ww16'

for expDate in ExpList:
    tableName=logTableGen.logTableGen(ticker, expDate, lowStrike, step, numStrikes)
    ww=xlLogGen.xlLogGen(ticker, expDate, lowStrike, step, numStrikes,tableName,ww)
    
xlFileName= '%s%d%d%d%s.xlsx' % (ticker,lowStrike,numStrikes,step,suffix)
print('\n....Excellent work Smithers... Excel file generated')
print(xlFileName)
ww.save(xlFileName)