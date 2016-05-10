import datetime, numpy as np
from MSOffice.Excel.Charts import XlGraphs
from MSOffice.Excel.Worksheets import Sheet, Template
from MSOffice.Excel.Launch import c
from ..Constants import * 
from .. import pos


class ORSPlots(): # ORSCalculator
	"""A class to create the SAIDI/SAIFI charts 
	from the ORS data."""

	Timestamp = 'Date'
	Cap = 'Cap'
	Target = 'Target'
	Collar = 'Collar'
	Planned = 'Planned'
	Unplanned = 'Unplanned'
	CapUnplanned = 'Capped Unplanned (Excess)'
		
	# Set the order of the headings and data in the sheet
	NetworkHeadings = ["ELIN", "OTPO", "TPCO"]
	IndexHeadings = ["SAIDI", "SAIFI"]
	DataHeadings = [Cap, Target, Collar, 
					Planned, Unplanned, CapUnplanned]

	def __init__(self, ORSCalculatorInstance, xlInstance):
		self.ORS = ORSCalculatorInstance
		self.xlInst = xlInstance
		self.Sheet = Sheet(self.xlInst)
		self.InputSheet = "Input"
		self.CalculationSheet = "Calculation"
		self.StatsSheet = "Statistics"
		
		#self.DefaultXAxis = self.Generate_X_Values()
		self.Graphs = XlGraphs(xlInstance, self.Sheet)
		#self.NumberOfRecords = self.get_num_records()
		
		self.srcRowOffset = 2 # The number of rows taken up by headings (before the data begins)
		self.srcColumnOffset = 1 # The number of columns that the data is offset by
		self.Timestamp = 'Date'
		self.Cap = 'Cap'
		self.Target = 'Target'
		self.Collar = 'Collar'
		self.Planned = 'Planned'
		self.Unplanned = 'Unplanned'
		self.CapUnplanned = 'Capped Unplanned (Excess)'
		
		# Set the order of the headings and data in the sheet
		self.NetworkHeadings = ["ELIN", "OTPO", "TPCO"]
		self.IndexHeadings = ["SAIDI", "SAIFI"]
		self.DataHeadings = [self.Cap, self.Target, self.Collar, 
								self.Planned, self.Unplanned, self.CapUnplanned]
								
		self.DateOrigin = pos(row=4, col=1) # The postion to place the fist date
		
		# Graph offsets and dimentions
		self.ChartDimentions = pos(x=700, y=480)
		self.ChartOrigins = []
		x = 85
		y = 100
		for network in self.NetworkHeadings:
			self.ChartOrigins.append(pos(x=x, y=y))
			x += self.ChartDimentions.x # Shift the charts for the next network to the right
		
	def __repr__(self):
		"""Return an identifying instance of the network being handled"""
		return "in the __repr__ for ORSPlots" + self.ORS.networknames[0]
		
	def Annual_Sheets(self, dates):
		"""Create a new sheet for every indvidual year in the ORS results"""
		for date in dates:
			year = str(date.year)
			self.Create_Sheet(year)
			self.Fill_Dates(date, year)
			self.Populate_Fixed_Stats(year) # Com Com table values scaled linearly
			self.Populate_Daily_Stats(year) # Daily real world SAIDI/SAIDI
			
			self.Create_Graphs(year)
	
	def Clean_Workbook(self):
		self.Sheet.rmvSheet(keepList=[self.InputSheet, self.StatsSheet])
	
	def Create_Sheets(self, dates):
		"""creates all the sheets needed in the workbook.
		Autofills the dates. All other data is network spesfic"""
		for date in dates:
			year = str(date.year)
			self.Create_Sheet(year)
			self.Fill_Dates(date, year)
	
	def Create_Sheet(self, suffix):
		"""Create the "Calculation" sheet in excel, delete it if it already exisits"""
		suffix = " " + suffix
		# Remove the sheet, if it exists, then re-add it -- Don't do this when adding multiple networks
		self.Sheet.rmvSheet(removeList=[self.CalculationSheet+suffix])
		self.Sheet.addSheet(self.CalculationSheet+suffix)
		
		DataOrigin = pos(row=4, col=1)
		LeftHeadings = [self.Timestamp]
		TableHeadings = LeftHeadings + \
			len(self.NetworkHeadings) * len(self.IndexHeadings) * self.DataHeadings
		
		# Assign row 3 data
		self.Sheet.setRange(self.CalculationSheet+suffix, DataOrigin.row-1, DataOrigin.col, [TableHeadings]) # Write the row headings into the table
		self.Sheet.mergeCells(self.CalculationSheet+suffix, 1, DataOrigin.col, DataOrigin.row-1, DataOrigin.col) # Date Cells
		
		# Assign row 2 data
		col = DataOrigin.col + len(LeftHeadings)
		columns = list(np.linspace(col, 
			col + (len(self.NetworkHeadings) * len(self.IndexHeadings) - 1) * len(self.DataHeadings), 
			len(self.NetworkHeadings) * len(self.IndexHeadings)))
		index = 0
		for col in columns:
			self.Sheet.setRange(self.CalculationSheet+suffix, DataOrigin.row-2, int(col), [[self.IndexHeadings[index % len(self.IndexHeadings)]]]) # self.IndexHeadings[int((col-len(LeftHeadings)) % len(self.IndexHeadings))]
			index += 1
			self.Sheet.mergeCells(self.CalculationSheet+suffix, 2, col, 2, col - 1 +
				len(self.NetworkHeadings) * len(self.IndexHeadings) * len(self.DataHeadings) / (len(self.NetworkHeadings) * len(self.IndexHeadings))) # Row 2
		
		# Assign row 1 data
		col = DataOrigin.col + len(LeftHeadings)
		columns = list(np.linspace(col, 
			col + (len(self.NetworkHeadings) * len(self.IndexHeadings) - 2) * len(self.DataHeadings), 
			len(self.NetworkHeadings)))
		index = 0
		for col in columns:
			self.Sheet.setRange(self.CalculationSheet+suffix, DataOrigin.row-3, int(col), [[self.NetworkHeadings[index % len(self.NetworkHeadings)]]])
			index += 1
			self.Sheet.mergeCells(self.CalculationSheet+suffix, 1, col, 1, col - 1 +
				len(self.NetworkHeadings) * len(self.IndexHeadings) * len(self.DataHeadings) / len(self.NetworkHeadings)) # Row 1
			
		# Fit cells (remove this)
		#for column in range(len(TableHeadings)):
		#    self.Sheet.autofit(self.CalculationSheet+suffix, column+1)

	def Generate_Dates(self, startdate, enddate=None):
		"""Generate an array of all the days for 1 fincial year"""
		if type(enddate) != datetime.datetime:
			enddate = datetime.datetime(startdate.year + 1, 3, 31)
		dt = datetime.timedelta(1) # A time delta of 1 day
		TimeStamps = []
		while startdate <= enddate:
			TimeStamps.append([startdate])
			startdate = startdate + dt
		return TimeStamps

	def Fill_Dates(self, date, suffix):
		"""Fill the date column in hte Excel sheet with the date values read from the parser"""
		suffix = " " + suffix
		row, col = \
			self.Sheet.setRange(self.CalculationSheet+suffix, self.DateOrigin.row, self.DateOrigin.col, self.Generate_Dates(date))

	def Populate_Fixed_Stats(self, suffix):
		"""Create series values for the Limit, Cap, and Collar. 
		Populate the excel sheet with these values."""
		suffix = " " + suffix
		network = self.ORS.networknames[0]
		if network == "ELIN":
			ColOffset = 0
		elif network == "OTPO":
			ColOffset = len(self.IndexHeadings) * len(self.DataHeadings)
		elif network == "TPCO":
			ColOffset = len(self.IndexHeadings) * len(self.DataHeadings) * (len(self.NetworkHeadings) - 1)
		RowHeadings = ["CAP", "TARGET", "COLLAR"] # The order the rows appear in the Excel spreadsheet
		
		self.Sheet.set_calculation_mode("manual")
		# Assumes that the data will start in col 2 i.e. col 1 is for the dates
		Column = 2 + ColOffset
		OriginCol = Column
		for param in self.IndexHeadings:
			for heading in RowHeadings:
				LinearRange = np.linspace(0, self.ORS._get_CC_stats(heading)[self.IndexHeadings.index(param)], 
					self.Sheet.getMaxRow(self.CalculationSheet+suffix, 1, 4) - 3 + 1)[1:]
				self.Sheet.setRange(self.CalculationSheet+suffix, 4, Column, [[i] for i in LinearRange])
				Column += 1
			Column += len(self.DataHeadings) - len(RowHeadings)
		self.Sheet.set_calculation_mode("automatic")
		
	def Populate_Daily_Stats(self, suffix):
		"""Create series values for the Planned and Unplanned SAIDI/SAIFI. 
		Populate the excel sheet with these values."""
		sheetname = self.CalculationSheet + " " + suffix
		network = self.ORS.networknames[0]
		ColOffset = 5 # Magic number: offset of the column in the data table
		if network == "ELIN":
			ColOffset += 0
		elif network == "OTPO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings)
		elif network == "TPCO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings) * (len(self.NetworkHeadings) - 1)
		RowHeadings = ["Unplanned", "Capped Unplanned", "Planned"] # The order the rows appear in the Excel spreadsheet
		
		self.Sheet.set_calculation_mode("manual")
		FiscalyearDays = self.Sheet.getRange(sheetname, 4, 1, self.Sheet.getMaxRow(sheetname, 1, 4), 1)
		FiscalyearDays = [self.Sheet.getDateTime(date[0]) for date in FiscalyearDays]
		SAIDIcol = []
		SAIFIcol = []
		row = 4
		for day in FiscalyearDays:
			SAIDIrow = []
			SAIFIrow = []
			x, y = self.ORS._get_indicies(day, "planned", applyBoundary=True)
			SAIDIrow.append(x)
			SAIFIrow.append(y)
			x, y = self.ORS._get_indicies(day, "unplanned", applyBoundary=True)
			SAIDIrow.append(x)
			SAIFIrow.append(y)
			x, y = self.ORS._get_indicies(day, "unplanned", applyBoundary=False)
			SAIDIrow.append(x-SAIDIrow[1])
			SAIFIrow.append(y-SAIFIrow[1])
			
			# Add the new rows to the table stored in memmory
			SAIDIcol.append(SAIDIrow)
			SAIFIcol.append(SAIFIrow)
			
			# here for debugging only
			#self.Sheet.setRange(sheetname, row, ColOffset, [SAIDIrow])
			#self.Sheet.setRange(sheetname, row, ColOffset+len(self.DataHeadings), [SAIFIrow])
			#row += 1
			
		
		# The table columns need to be cummulative
		SAIDIsums = [0 for i in RowHeadings]
		SAIFIsums = [0 for i in RowHeadings]
		SAIDITable = []
		SAIFITable = []
		row = 4
		# Loop through every row
		for SAIDIrow, SAIFIrow in zip(SAIDIcol, SAIFIcol):
			ColumnIndex = 0
			# Loop through every column
			for SAIDIval, SAIFIval in zip(SAIDIrow, SAIFIrow):
				SAIDIsums[ColumnIndex] += SAIDIval
				SAIFIsums[ColumnIndex] += SAIFIval
				ColumnIndex += 1
			#self.Sheet.setRange(sheetname, row, ColOffset, [SAIDIsums])
			#self.Sheet.setRange(sheetname, row, ColOffset+len(self.DataHeadings), [SAIFIsums])
			SAIDITable.append(SAIDIsums[:]) # This copys by value, not by reference
			SAIFITable.append(SAIFIsums[:]) # This copys by value, not by reference
			row += 1
			
		self.Sheet.setRange(sheetname, 4, ColOffset, SAIDITable)
		self.Sheet.setRange(sheetname, 4, ColOffset+len(self.DataHeadings), SAIFITable)
		self.Sheet.set_calculation_mode("automatic")
	
	def Create_Graphs(self, suffix):
		"""Create the SAIDI/SAIFI chart"""
		sheetname = self.CalculationSheet + " " + suffix
		network = self.ORS.networknames[0]
		ColOffset = 2 # Magic number: where the data starts in the table (column 2)
		if network == "ELIN":
			ColOffset += 0
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartEIL.crtx')
		elif network == "OTPO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings)
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartOJV.crtx')
		elif network == "TPCO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings) * (len(self.NetworkHeadings) - 1)
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartTPC.crtx')
		
		ylables = ["Average Outage Duration (Minutes/ICP)", "Average No. Outages (Interruptions/ICP)"]
		for stat in self.IndexHeadings:
			ChartName = network + " " + stat + " " + suffix # e.g. ELIN SAIDI 2015
			self.Graphs.Create_Chart(ChartName, self.Generate_Date_Range(suffix),
					sheetname=sheetname)
			# Add the indvidual series to the chart
			i = self.IndexHeadings.index(stat) * (len(self.DataHeadings) * (len(self.IndexHeadings) - 1)) # Add a sub-offset
			for col in range(i + ColOffset, i + ColOffset + len(self.DataHeadings)):
				self.Graphs.Add_Series(ChartName, self.Generate_Range(suffix, col), serieslabels=True)
			# Apply the templated style, reapply attributes like ylabel
			self.Graphs.Apply_Template(ChartName, chartpath,
					ylabel=ylables[self.IndexHeadings.index(stat)])
			# Make the chart bigger
			self.Graphs.Set_Dimentions(ChartName, self.ChartDimentions.x, self.ChartDimentions.y)
			# Stack charts from the same network vetically
			origin = self.ChartOrigins[self.NetworkHeadings.index(network)]
			self.Graphs.Set_Position(ChartName, origin.x, origin.y + self.IndexHeadings.index(stat) * self.ChartDimentions.y)

	def Generate_Date_Range(self, suffix, **kwargs):
		"""@return: String of a formatted excel range.
		Create an excel formulae for a range"""
		sheetname = kwargs.get("sheetname", self.CalculationSheet + " " + suffix)
		return "='%s'!%s%d:%s%d" % (sheetname, self.Sheet.num_to_let(self.DateOrigin.col), self.DateOrigin.row, 
			self.Sheet.num_to_let(self.DateOrigin.col), self.Sheet.getMaxRow(sheetname, self.DateOrigin.col, self.DateOrigin.row))

	def Generate_Range(self, suffix, col, **kwargs):
		"""@return: String of a formatted excel range.
		Create an excel formulae for a range i nthe supplied sheets"""
		sheetname = kwargs.get("sheetname", self.CalculationSheet + " " + suffix)
		if type(col) is int:
			col = self.Sheet.num_to_let(col) # Convert the column number to a corosponding letter to match an excel table 
		return "='%s'!%s%d:%s%d" % (sheetname, col, self.DateOrigin.row-1, 
			col, self.Sheet.getMaxRow(sheetname, self.DateOrigin.col, self.DateOrigin.row))


class ORSOutput(ORSPlots):
	"""A module to publish results tables, and restore default values to
	user input fields."""
	def __init__(self, ORSCalculatorInstance, xlInstance):
		self.ORS = ORSCalculatorInstance
		#self.xlInst = xlInstance
		self.Sheet = Sheet(xlInstance)
		self.InputSheet = "Input"
		self.CalculationSheet = "Calculation"
		self.StatsSheet = "Statistics"
		self.NetworkName = self.ORS.networknames[0]
		
		self.DateOrigin = pos(row=4, col=1) # The postion to place the fist date
	
	def Populate_Reliability_Stats(self):
		"""Write the 2015 determination statistics to a prebuilt table in excel.
		Used on the 'Statistics' sheet."""
		# Table origin is the top left
		ComComOrigin = pos(row=4, col=1)
		CalcOrigin = pos(row=4, col=9)
		
		# The row order of the values being calcualted by the ORSCalculator
		# TODO: The row text must match the values _get_stats() is expecting, we may need to introduce some aliases to make it more readable
		RowHeadings = [["UBV"], ["LIMIT"], ["TARGET"], ["COLLAR"], ["CAP"]]
		CalcRowValues = [self.ORS._get_stats(heading[0]) for heading in RowHeadings] # ORS calcualted values
		CCRowValues = [self.ORS._get_CC_stats(heading[0]) for heading in RowHeadings] # Com Com hard coded values
		
		# Work out how many columns we need to offset these SAIDI, SAIFI results
		NetworkOrder = ["ELIN", "OTPO", "TPCO"] # The order the networks are presented going across the table in excel
		try:
			Offset = NetworkOrder.index(self.NetworkName)
		except ValueError:
			print "%s is not a recognised network" % self.NetworkName
			Offset = 0 # Default option
		Offset *= 2
		
		# Table 1
		self.Sheet.setRange(self.StatsSheet, ComComOrigin.row, ComComOrigin.col, RowHeadings) # Write the row headings into the table
		ComComOrigin.col += 1 # Incriment the next column (i.e. don't overwrite row headings)
		row = ComComOrigin.row
		for rowValue in CalcRowValues:
			# The SAIDI/SAIFI values are written to excel in the order that they are returned by the ORSCalculator function
			self.Sheet.setRange(self.StatsSheet, row, ComComOrigin.col+Offset, [rowValue])
			row += 1

		# Table 2
		self.Sheet.setRange(self.StatsSheet, CalcOrigin.row, CalcOrigin.col, RowHeadings)
		CalcOrigin.col += 1 # Incriment the next column (i.e. don't overwrite row headings)
		row = CalcOrigin.row
		for rowValue in CCRowValues:
			# The SAIDI/SAIFI values are written to excel in the order that they are returned by the ORSCalculator function
			self.Sheet.setRange(self.StatsSheet, row, CalcOrigin.col+Offset, [rowValue])
			row += 1

	def Create_Summary_Table(self):
		"""Create a summary sheet in excel, delete it if it already exisits"""
		# Remove the sheet, if it exists, then re-add it -- Don't do this when adding multiple networks
		self.Sheet.rmvSheet(removeList=["YTD Monthly Breakdown"])
		self.Sheet.addSheet("YTD Monthly Breakdown")

	def Summary_Table(self, suffix):
		"""Publish a summary of the year-to-date stats at the bottom of every sheet"""
		suffix = " " + suffix
		network = self.ORS.networknames[0]
		currentDate = datetime.datetime.now()
		RowHeadings = ["CAP", "TARGET", "COLLAR"] # The order the rows appear in the Excel spreadsheet
		TableHeadings = ["YTD Cap", "YTD Target", "YTD Collar", "YTD Total", "YTD Planned", "YTD Unplanned", "Projected Incentive/Penalty"]
		columns = [1, 2, 3, 4, 5]
		if network == "ELIN":
			RowOffset = 2
			ColOffset = 1
		elif network == "OTPO":
			RowOffset = 2 + 12
			ColOffset = len(self.IndexHeadings) * len(self.DataHeadings) + 1
		elif network == "TPCO":
			RowOffset = 2 + 2*12
			ColOffset = len(self.IndexHeadings) * len(self.DataHeadings) * (len(self.NetworkHeadings) - 1) + 1

		maxrow = self.Sheet.getMaxRow(self.CalculationSheet+suffix, 1, 4)
		self.Sheet.setRange("Summary", maxrow + RowOffset, 1, [[network]+TableHeadings]) # Write the heading data
		
		# Find the row that corrosponds to the current date
		Dates = self.Sheet.getRange(self.CalculationSheet+suffix, 4, 1, maxrow, 1)
		Dates = [self.Sheet.getDateTime(Date[0]) for Date in Dates] # Convert a 2D list of tuples to a 1D list
		try:
			index = Dates.index( datetime.datetime(currentDate.year, currentDate.month, currentDate.day) )
		except ValueError:
			index = len(Dates) - 1
				
		for param in self.IndexHeadings:
			# Read the entire row of data
			YTD_row = self.Sheet.getRange(self.CalculationSheet+suffix, index+4, 1, index+4, 
				self.Sheet.getMaxCol(self.CalculationSheet+suffix, 2, 3))[0]
			# Convert the row data to: CAP, TARGET, COLLAR, YTD Total, YTD Planned, YTD Unplanned
			#YTD_row[ColOffset : len(DataHeadings)+ColOffset+1]
			i = self.IndexHeadings.index(param)
			TableRow = [YTD_row[ColOffset], YTD_row[ColOffset+1], YTD_row[ColOffset+2], 
			   YTD_row[ColOffset+3] + YTD_row[ColOffset+4], YTD_row[ColOffset+3], 
					[0.5*CC_Revenue_At_Risk.get(network, 0)/(self.ORS._get_stats("CAP")[i] - self.ORS._get_stats("TARGET")[i])]]

			RowOffset += 1
			self.Sheet.setRange("Summary", maxrow + RowOffset, 1, [[param]+TableRow]) # Write the heading data
			ColOffset += len(self.DataHeadings)
		
		Table = []
		Table.append(["Revenue at risk", CC_Revenue_At_Risk.get(network, "No Revenue Found")]) 		# Revenue at Risk
		Table.append(["Total Number of ICPs", self.ORS._get_total_customers(Dates[index])]) 		# Total Number of ICPs
		Table.append(["Year to date figures as of", Dates[index]]) 		# Date
		self.Sheet.setRange("Summary", maxrow + RowOffset+1, 1, Table)

	def YTD_Stats(self, suffix):
		"""Use a template table to represent the results.
		There will be tables for every month ending since the start of the new fiscal year."""
		# Find the row that corrosponds to the current date; column values read in appear as a list of tuples
		suffix = " " + suffix
		Dates = self.Sheet.getRange(self.CalculationSheet+suffix, 4, 1, maxrow, 1)
		Dates = [self.Sheet.getDateTime(Date[0]) for Date in Dates]
		try:
			currentDate = datetime.datetime.now()
			index = Dates.index(datetime.datetime(currentDate.year, currentDate.month, currentDate.day))
		except ValueError:
			# We want the last date of the year, since  this date is not in the searched range
			index = len(Dates) - 1

		template = Template(self.Sheet, r"C:\Users\sdo\Documents\Research and Learning\Git Repos\SAIDI-SAIFI-Calculator\Data\Templates.xlsx")
		template.Place_Template("Summary", self.Sheet._getCell("A1"))
		params = {"DATE": Dates[index], "CUST_NO": self.ORS._get_total_customers(Dates[index]), "REV_RISK": CC_Revenue_At_Risk.get(network, "No Revenue Found"),
			"SAIFI_YTD_CAP": 0, "SAIFI_YTD_TARGET": 0}
		template.Set_Values(params)
		template.Auto_Fit()