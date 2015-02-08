'''
Making a main program to call the routines and build the excel workbook.
2/7/2015
'''
'''
Idea is to have a way to set up an arbitrary number of stocks adn generate
the excel workbook.  I will try first just looping and building one by one
and see if this strategy works.
first, set up the variables for one of the stocks
'''

ticker = input('Enter stock ticker: ')
expDate = input('Enter the expiration date: ')
lowStrike = float(input('Enter the low strike: '))
step = float(input('Enter the step size between strikes: '))
numStrikes = float(input('Enter the number of strikes: '))