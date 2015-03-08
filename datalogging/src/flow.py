'''
Making a main program to call the routines and build the excel workbook.
2/7/2015
'''
from openpyxl.workbook import Workbook
'''
Idea is to have a way to set up an arbitrary number of stocks adn generate
the excel workbook.  I will try first just looping and building one by one
and see if this strategy works.
first, set up the variables for one of the stocks
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
    ww=xlLogGen.xlLogGen(ticker, expDate, lowStrike, step, numStrikes,ww)
    print('The table created is ', tableName)
    loopDone= input('Are you finished? Please enter yes or no. ')
xlFileName= input('Enter name for your new excel logging file: ')
ww.save(xlFileName)
print('Congratulations - we are done. ')