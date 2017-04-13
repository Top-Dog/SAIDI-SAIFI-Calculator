'''
Created on 13/4/2016

@author: Sean D. O'Connor
'''

import datetime
from MSOffice.Excel.Worksheets import Sheet
from .. import pos, CUST_NUMS

class ParseORS(object):
	NetworkNames = ["ELIN", "TPCO", "OTPO", "LLNW"] # The order the network names appear in the input spreadsheet
	InputSheet = "Input" # The name of the Excel sheet used for user input
	Months = ["March (yr 0)", "April", "May", "June", "July", "August", "September", "October",
			"November", "December", "January", "February", "March (yr 1)"] # The months in a fiscal year

	def __init__(self, xlInstance):
		self.xlInst = xlInstance
		self.Sheet = Sheet(xlInstance)
		
		
		# Create some dictionaries to hold customer numbers
		
		self.Customer_Nums = {}
		for name in self.NetworkNames:
			self.Customer_Nums[name] = {}
		
		# Create a ordered list of network charts/tables to display
		self.StartDates = []

	def Read_Num_Cust(self):
		"""Read the average number of customers"""
		o = pos(row=3, col=1)
		Table = self.Sheet.getRange(self.InputSheet, o.row, o.col,
					self.Sheet.getMaxRow(self.InputSheet, o.col, o.row), 
					self.Sheet.getMaxCol(self.InputSheet, o.col, o. row))
		# Get the number of customers - fill up the dictionaries
		for row in Table:
			cell = row.__iter__()
			year = cell.next().year # Skip to column B
			cell.next() # Skip over column C
			for name in self.NetworkNames:
				self.Customer_Nums[name][year] = cell.next() # Get the value of the current cell and incriment the address pointer
		return self.Customer_Nums

	def _Read_Last_Date(self):
		"""Read the last date the user has input data for"""
		o = pos(row=3, col=1)
		# Read the last date in the yearly table
		FinalDate = self.Sheet.getCell(self.InputSheet, o.col, 
								 self.Sheet.getMaxCol(self.InputSheet, o.col, o. row))
		return self.Sheet.getDateTime(FinalDate)

	def Read_Num_Cust_Final(self):
		"""Read the number of customers in the final year only.
		This is data from the by month table."""
		o = pos(row=3, col=1)
		MonthlyICPs = {}
		# Read the last date in the yearly table
		FinalDate = self._Read_Last_Date()
		
		Table = self.Sheet.getRange(self.InputSheet, o.row, o.col,
					self.Sheet.getMaxRow(self.InputSheet, o.col, o.row), 
					self.Sheet.getMaxCol(self.InputSheet, o.col, o. row))
		
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
	
	def Read_Last_Date(self):
		"""Read (get) the last date field"""
		o = pos(row=17, col=12)
		lastdate = self.Sheet.getCell(self.InputSheet, o.row, o.col)
		try:
			return self.Sheet.getDateTime(lastdate)
		except AttributeError:
			# The user entered an invalid date 
			return datetime.datetime.now()

	def Restore_Input_Default(self):
		"""Restore the default state of the table that has all the ICP counts for each network"""
		o = pos(row=3, col=1)
		OutputTable = []
		for year in CUST_NUMS.get(self.NetworkNames[0]):
			RowData = [datetime.datetime(year, 3, 31), year-1]
			for name in self.NetworkNames:
				RowData.append(CUST_NUMS.get(name).get(year))
			#RowData.append("Y") # The default option is to display the sheet
			OutputTable.append(RowData)
		# Sort the rows in the table so that we have the dates in cronological order
		OutputTable = sorted(OutputTable, key=lambda e: e[0], reverse=False)
		# Set the output range in excel
		self.Sheet.setRange(self.InputSheet, o.row-1, o.col+2, [self.NetworkNames])
		self.Sheet.setRange(self.InputSheet, o.row, o.col, OutputTable)

	def Set_Mean_ICPs(self, lastyear):
		"""Sets the mean (average) number of ICPS in the annual table from 
		records in the monthly table. The date 31/03/XXXX must appear in the 
		main annual data table, otherwise no updates will occur."""
		o_table1 = pos(row=3, col=1) # Top left cnr of annual table
		o_table2 = pos(row=3, col=11) # Top left cnr of the monthly table
		# Table Rows: [Month name, Year of months occurance, self.NetworkNames[]]

		# Find the average number of ICPs
		coloffset = 2 # Offset of ICP data from the left most column in the table
		maxrow = 15 # Last row in the table
		for network in self.NetworkNames:
			# Find the number of records in the monthly table
			lastrow = self.Sheet.getMaxRow(self.InputSheet, coloffset + o_table2.col, o_table2.row)
			if lastrow > maxrow:
				lastrow = o_table2.row
			# Compute the average number of ICPs (using the first and last row in the monthly table,
			# in the case of only a single month, then the first and last row are the same).
			avrg = (self.Sheet.getCell(self.InputSheet, lastrow, coloffset + o_table2.col) + \
			    self.Sheet.getCell(self.InputSheet, o_table2.row, coloffset + o_table2.col)) / 2

			# Place the average in the specified record
			# Input param replcaed the line below
			#lastyear = int(self.Sheet.getCell(self.InputSheet, o_table2.row, o_table2.col + 1)) + 1
			try:
				lastrow = self.Sheet.brief_search(self.InputSheet, "31/03/"+str(lastyear+1)).Row
				self.Sheet.setCell(self.InputSheet, lastrow, coloffset + o_table1.col, avrg)
			except:
				pass

			coloffset += 1

	def Restore_Table_2(self, lastyear=None):
		"""Builds the table for gathering ICP data by month"""
		o_table1 = pos(row=3, col=1) # Top left cnr of annual table
		o_table2 = pos(row=3, col=11) # Top left cnr of the monthly table

		# Get the row for the last date in the annual data table (table1)
		lastrow = self.Sheet.getMaxRow(self.InputSheet, o_table1.col, o_table1.row)
		if lastrow > 10000:
			lastrow = o_table1.row

		# If lastyear is None, then the last row in year in the last row of the
		# annual data table (table1) is used.
		if not lastyear:
			lastyear = self.Sheet.getDateTime(
					self.Sheet.getCell(self.InputSheet, lastrow, o_table1.col)).year

		# Build the two left most column of table 2 (Months names, Fiscal year)
		rowindex = 0
		# Iterate over an array of months (arranged in order for a fiscal year)
		for month in self.Months:
			fiscalyear = lastyear
			if self.Months.index("December") >= self.Months.index(month):
				# We are in first 3/4 of the year
				fiscalyear -= 1
			
			# Write each row in table2 (first two columns only), for example, "March (yr0), 2016"
			self.Sheet.setRange(self.InputSheet, o_table2.row+rowindex, o_table2.col, [(month, fiscalyear)])
			rowindex += 1

		# DON'T DO THIS! This will copy the last years average... we want the number of ICPs
		# as of the 31st March on the previous fiscal year instead...
		# Automatically copy the previous years data to record 0
		##lastrow -= 1 # Get the previous years data
		##if lastrow <= o_table1.row:
		##	lastrow = o_table1.row
		##previousrange= self.Sheet.getRange(self.InputSheet, lastrow, o_table1.col+2, lastrow, o_table1.col+5)
		##self.Sheet.setRange(self.InputSheet, o_table2.row, o_table2.col+2, previousrange)

	def Set_Year(self, year):
		"""Used to update the table labels at the end/start of a new fiscal year"""
		o_table1 = pos(row=3, col=1) # Top left cnr of annual table
		o_table2 = pos(row=3, col=11) # Top left cnr of the monthly table

		# Try and find the year in table1, otherwise create a new row(s)
		lastrow = self.Sheet.getMaxRow(self.InputSheet, o_table1.col, o_table1.row)
		if lastrow > 10000:
			lastrow = o_table1.row

		try:
			lastyear = self.Sheet.getDateTime(
				self.Sheet.getCell(self.InputSheet, lastrow, o_table1.col)).year
		except:
			lastyear = datetime.datetime.strptime(
					self.Sheet.getCell(self.InputSheet, lastrow, o_table1.col),
					"%d/%m/%Y").year
		while lastyear <= year:
			# We need to create some new rows in table1
			lastrow += 1
			lastyear += 1
			self.Sheet.setRange(self.InputSheet, lastrow, o_table1.col, [(datetime.datetime(lastyear, 3, 31), lastyear-1)])

		# Change the labels on table 2 to match the selected year
		self.Restore_Table_2(lastyear)