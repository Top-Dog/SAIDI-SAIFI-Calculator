'''
Created on 31/05/2016

@author: Sean D. O'Connor
'''
import time, sys
from SAIDISAIFI import Parser, Output
from MSOffice import Excel


if __name__ == "__main__":    
	# Get the currently active MS Excel Instace
	xl = Excel.Launch.Excel(visible=True, runninginstance=True)
	print xl.xlBook.Name
	
	# Instsiate the parser, restore the defaults to the Input worksheet
	p = Parser.ParseORS(xl)
	p.Set_Mean_ICPs(int(sys.argv[1]))
	#p.Restore_Table_2(int(sys.argv[1])+1)

	print "Success!"
	time.sleep(5)
