'''
Created on 15/04/2016

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
	p.Restore_Input_Default() # Restores the annual input table
	p.Restore_Table_2() # Restores the monthly data input table
	p.Set_Mean_ICPs() # Automtically calculate and apply the average number of ICPs

	# Clear all the worksheets we don't need
	xlPlotter = Output.ORSPlots(None, xl)
	xlPlotter.Clean_Workbook()

	print "Success!"
	time.sleep(5)