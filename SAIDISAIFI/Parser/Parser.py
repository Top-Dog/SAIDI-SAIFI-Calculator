'''
Created on 13/4/2016

@author: Sean D. O'Connor
'''

import datetime
from MSOffice.Excel.Worksheets import Sheet
from .. import pos, CUST_NUMS

class ParseORS(object):
	def __init__(self, xlInstance):
		self.xlInst = xlInstance
		self.Sheet = Sheet(xlInstance)
		self.InputSheet = "Input"
		
		# Create some dictionaries to hold customer numbers
		self.NetworkNames = ["ELIN", "TPCO", "OTPO", "LLNW"] # The order the network names appear in the input spreadsheet
		self.Customer_Nums = {}
		for name in self.NetworkNames:
			self.Customer_Nums[name] = {}
		
		# Create a ordered list of network charts/tables to display
		self.StartDates = []

		
	def Read_Num_Cust(self):
		"""Read the number of customers"""
		o = pos(row=3, col=1)
		Table = self.Sheet.getRange(self.InputSheet, o.row, o.col,
					self.Sheet.getMaxRow(self.InputSheet, o.col, o.row), 
					self.Sheet.getMaxCol(self.InputSheet, o.col, o. row))
		# Get the number of customers - fill up the dictionaries
		for row in Table:
			cell = row.__iter__()
			year = cell.next().year
			cell.next() # Skip over a column
			for name in self.NetworkNames:
				self.Customer_Nums[name][year] = cell.next()
		return self.Customer_Nums
		
	def Read_Dates_To_Publish(self):
		"""Read the dates that require results to be outputted.
		Returns the date range expected by the ORS calcualtor."""
		o = pos(row=3, col=1)
		Table = self.Sheet.getRange(self.InputSheet, o.row, o.col,
					self.Sheet.getMaxRow(self.InputSheet, o.col, o.row), 
					self.Sheet.getMaxCol(self.InputSheet, o.col, o. row))
		# Get the number of customers - fill up the dictionaries
		for row in Table:
			if row[-1] == "Y" or row[-1] == "y":
				self.StartDates.append(datetime.datetime(row[0].year-1, 4, 1)) # Substract 1 from the end of fiscal year to get the start of the year
		if len(self.StartDates):
			return min(self.StartDates), datetime.datetime(max(self.StartDates).year+1, 3, 31) # Add 1 back to the year, so we get the end of fiscal year
		else:
			return None, None
		
	def Restore_Input_Default(self):
		"""Restore the default state of the table that has all the ICP counts for each network"""
		o = pos(row=3, col=1)
		NetworkNames = ["ELIN", "TPCO", "OTPO", "LLNW"] # [name for name in CUST_NUMS] --> gives a random order
		OutputTable = []
		for year in CUST_NUMS.get(NetworkNames[0]):
			rowData = [datetime.datetime(year, 3, 31), year-1]
			for name in NetworkNames:
				rowData.append(CUST_NUMS.get(name).get(year))
			rowData.append("Y") # The default option is to display the sheet
			OutputTable.append(rowData)
		# Sort the rows in the table so that we have the dates in cronological order
		OutputTable = sorted(OutputTable, key=lambda e: e[0], reverse=False)
		# Set the output range in excel
		self.Sheet.setRange(self.InputSheet, o.row-1, o.col+1, [NetworkNames])
		self.Sheet.setRange(self.InputSheet, o.row, o.col, OutputTable)