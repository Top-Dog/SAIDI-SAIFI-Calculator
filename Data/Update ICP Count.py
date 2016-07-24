'''
Created on 31/05/2016

@author: Sean D. O'Connor
'''
import time
from SAIDISAIFI import Parser, Output
from MSOffice import Excel


if __name__ == "__main__":    
	# Get the currently active MS Excel Instace
	xl = Excel.Launch.Excel(visible=True, runninginstance=True)
	print xl.xlBook.Name
	
	# Instsiate the parser, restore the defaults to the Input worksheet
	p = Parser.ParseORS(xl)
	p.Set_Mean_ICPs()
	p.Restore_Table_2()

	print "Success!"
	time.sleep(5)
