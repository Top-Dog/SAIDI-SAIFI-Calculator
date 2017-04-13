'''
Created on 13/04/2017

@author: Sean D. O'Connor
'''
import time, sys
from SAIDISAIFI import Parser, Output
from MSOffice import Excel


if __name__ == "__main__":    
	# Get the currently active MS Excel Instace
	xl = Excel.Launch.Excel(visible=True, runninginstance=True)
	print xl.xlBook.Name
	
	# Instsiate the parser
	p = Parser.ParseORS(xl)

	# Expects input as a interger referencing the fiscal year that 
	# is being worked on e.g. 2017 for the year ending 31/3/2018
	p.Set_Year(int(sys.argv[1]))

	print "Success!"
	#time.sleep(5)