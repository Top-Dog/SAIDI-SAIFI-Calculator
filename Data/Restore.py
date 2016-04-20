'''
Created on 15/04/2016

@author: Sean D. O'Connor
'''

from SAIDISAIFI import Parser, Output
from MSOffice import Excel


if __name__ == "__main__":    
	# Get the currently active MS Excel Instace
	xl = Excel.Launch.Excel(visible=True, runninginstance=True)
	print xl.xlBook.Name
	
	# Instsiate the parser, restore the defaults to the Input worksheet
	p = Parser.ParseORS(xl)
	p.Restore_Input_Default()

	# Clear all the worksheets we don't need
	xlPlotter = Output.ORSPlots(None, xl)
	xlPlotter.Clean_Workbook()
