import datetime, numpy as np
from MSOffice.Excel.Charts import XlGraphs
from MSOffice.Excel.Worksheets import Sheet, Template, shtRange
from MSOffice.Excel.Launch import c
from ..Constants import * 
from .. import pos


class ORSPlots(object): # ORSCalculator
	"""A class to create the SAIDI/SAIFI charts 
	from the ORS data."""

	Timestamp = 'Date'
	Cap = 'Cap/Limit'
	Target = 'Target'
	Collar = 'Collar'
	Planned = 'Planned'
	Unplanned = 'Unplanned Normalised'
	CapUnplanned = 'Unplanned (Normalised Out)'
		
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
		
		self.Graphs = XlGraphs(xlInstance, self.Sheet)
		
		self.srcRowOffset = 2 # The number of rows taken up by headings (before the data begins)
		self.srcColumnOffset = 1 # The number of columns that the data is offset by
		
		# Set the order of the headings and data in the sheet
		self.NetworkHeadings = ["ELIN", "OTPO", "TPCO"]
		self.IndexHeadings = ["SAIDI", "SAIFI"]			
		self.DateOrigin = pos(row=4, col=1) # The postion to place the fist date on the Excel sheet
		
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
			self.Populate_Daily_Stats(date) # Daily real world SAIDI/SAIDI
			
			self.Create_Graphs(date)
	
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
		"""Fill the date column in the Excel sheet with the date values read from the parser"""
		suffix = " " + suffix
		row, col = \
			self.Sheet.setRange(self.CalculationSheet+suffix, self.DateOrigin.row, self.DateOrigin.col, self.Generate_Dates(date))

	def _Correct_Graph_Slope(self, suffix, enddate=datetime.datetime.now()):
		"""When the data is truncated for stacked area charts we need two of the same end dates
		in a row to create a vertical drop on the graph, otherwise it slopes with delta x = one time interval"""
		suffix = " " + suffix
		searchterm = datetime.datetime(enddate.year, enddate.month, enddate.day).date().__str__().split('-')
		searchterm = searchterm[2] + '/' + searchterm[1] + '/' + searchterm[0]

		maxrow = self.Sheet.getMaxRow(self.CalculationSheet+suffix, 1, 4)
		results = self.Sheet.search(shtRange(self.CalculationSheet+suffix, None, 4, 1, maxrow, 1), 
						searchterm)
		if len(results) == 1:
			self.Sheet.setCell(self.CalculationSheet+suffix, results[0].Row+1, 1, results[0].Value)

	def _Correct_Graph_Axis(self, ChartName, enddate=datetime.datetime.now()):
		"""Adds the final end of year date to the x axis e.g. 31/3/xxxx"""
		#ChartName = FullName + " " + stat + " YTD as at " + enddate.date().__str__()
		self.Graphs.Set_Max_X_Value(ChartName, enddate) 

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

	def _Calc_Rows(self, dates, ORS):
		SAIDIcol = []
		SAIFIcol = []
		for day in dates:
			SAIDIrow = []
			SAIFIrow = []
			x, y = ORS._get_indicies(day, "planned", applyBoundary=True)
			SAIDIrow.append(x)
			SAIFIrow.append(y)
			x, y = ORS._get_indicies(day, "unplanned", applyBoundary=True)
			SAIDIrow.append(x)
			SAIFIrow.append(y)
			x, y = ORS._get_indicies(day, "unplanned", applyBoundary=False)
			SAIDIrow.append(x-SAIDIrow[1])
			SAIFIrow.append(y-SAIFIrow[1])

			SAIDIcol.append(SAIDIrow)
			SAIFIcol.append(SAIFIrow)
			
			# Here for debugging only
			#row=4
			#self.Sheet.setRange(sheetname, row, ColOffset, [SAIDIrow])
			#self.Sheet.setRange(sheetname, row, ColOffset+len(self.DataHeadings), [SAIFIrow])
			#row += 1
		return SAIDIcol, SAIFIcol
		
	def Populate_Daily_Stats(self, enddate=datetime.datetime.now()):
		"""Create series values for the Planned and Unplanned SAIDI/SAIFI. 
		Populate the excel sheet with these values."""
		FiscalYear = str(self._get_fiscal_year(enddate))
		sheetname = self.CalculationSheet + " " + FiscalYear
		if enddate > datetime.datetime.now():
			enddate = datetime.datetime.now()
		print "Debug, the new enddate:", enddate

		network = self.ORS.networknames[0]
		ColOffset = 5 # Magic number: offset of the column in the data table
		if network == "ELIN":
			ColOffset += 0
		elif network == "OTPO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings)
		elif network == "TPCO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings) * (len(self.NetworkHeadings) - 1)
		
		self.Sheet.set_calculation_mode("manual")
		startdate = self.Sheet.getDateTime(self.Sheet.getCell(sheetname, 4, 1))
		lastdate = self.Sheet.getDateTime(self.Sheet.getCell(sheetname, self.Sheet.getMaxRow(sheetname, 1, 4), 1))
		delta_days = (lastdate - startdate).days +  1
		FiscalyearDays = [startdate + datetime.timedelta(days=i) for i in range(delta_days)]
		
		# Truncate the data if it is for the present year
		maxrow = self.Sheet.getMaxRow(sheetname, 1, 4)
		#enddate = datetime.datetime.now() # Use this variable to force the graphs to only display a limited set of information
		searchterm = datetime.datetime(enddate.year, enddate.month, enddate.day).date().__str__().split('-')
		searchterm = searchterm[2] + '/' + searchterm[1] + '/' + searchterm[0]
		searchresult = self.Sheet.search(shtRange(sheetname, None, 4, 1, maxrow, 1), 
							  searchterm)
		if len(searchresult) == 1 or len(searchresult) == 2: # There could be two dates, if we are duplicating the end date to achieve a vertical truncation on our stacked area chart
			StopTime = self.Sheet.getDateTime(searchresult[0].Value)
		else:
			StopTime = FiscalyearDays[-1] # update this to be FiscalyearDays[0] now?

		SAIDIcol, SAIFIcol = self._Calc_Rows(FiscalyearDays, self.ORS)
		
		# The table columns need to be cummulative
		SAIDIsums = [0 for i in self.DataHeadings[3:]]
		SAIFIsums = [0 for i in self.DataHeadings[3:]]
		SAIDITable = []
		SAIFITable = []
		row = 4
		# Loop through every row
		for SAIDIrow, SAIFIrow, day in zip(SAIDIcol, SAIFIcol, FiscalyearDays):
			ColumnIndex = 0
			# Loop through every column
			for SAIDIval, SAIFIval in zip(SAIDIrow, SAIFIrow):
				# Add the new rows to the table stored in memmory
				if day <= StopTime: # means we will stop on the current day, but then fixing graph slopes break
					SAIDIsums[ColumnIndex] += SAIDIval
					SAIFIsums[ColumnIndex] += SAIFIval
				else:
					SAIDIsums[ColumnIndex] = None
					SAIFIsums[ColumnIndex] = None
				ColumnIndex += 1
			#self.Sheet.setRange(sheetname, row, ColOffset, [SAIDIsums])
			#self.Sheet.setRange(sheetname, row, ColOffset+len(self.DataHeadings), [SAIFIsums])
			SAIDITable.append(SAIDIsums[:]) # This copys by value, not by reference
			SAIFITable.append(SAIFIsums[:]) # This copys by value, not by reference
			row += 1
			
		self.Sheet.setRange(sheetname, 4, ColOffset, SAIDITable)
		self.Sheet.setRange(sheetname, 4, ColOffset+len(self.DataHeadings), SAIFITable)
		self._Correct_Graph_Slope(FiscalYear) # Makes the area plot look a bit better, but mutates the source data, so must be run last
		self.Sheet.set_calculation_mode("automatic")

	def _get_fiscal_year(self, enddate):
		"""Get the fiscal year as defined in NZ. Returns the year as a int"""
		year = enddate.year
		if enddate.month - 4 < 0:
			year -= 1
		return year
	
	def Create_Graphs(self, enddate):
		"""Create the SAIDI/SAIFI chart"""
		FiscalYear = str(self._get_fiscal_year(enddate))
		graphenddate = None
		if enddate == datetime.datetime(int(FiscalYear)+1, 3, 31):
			graphenddate = datetime.datetime(int(FiscalYear)+1, 4, 1)

		if enddate > datetime.datetime.now():
			enddate = datetime.datetime.now()

		network = self.ORS.networknames[0]
		ColOffset = 2 # Magic number: where the data starts in the table (column 2)
		FullName = ""
		if network == "ELIN":
			ColOffset += 0
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartEIL.crtx')
			FullName = "Electricty Invercargill"
		elif network == "OTPO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings)
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartOJV.crtx')
			FullName = "OtagoNet+ESL"
		elif network == "TPCO":
			ColOffset += len(self.DataHeadings) * len(self.IndexHeadings) * (len(self.NetworkHeadings) - 1)
			chartpath = os.path.expanduser('~/Documents/SAIDI and SAIFI/Templates/ORSChartTPC.crtx')
			FullName = "The Power Company"
		
		ylables = ["Average Outage Duration (Minutes/ICP)", "Average No. Outages (Interruptions/ICP)"]
		for stat in self.IndexHeadings:
			ChartName = FullName + " " + stat + " YTD as at " + enddate.date().__str__()
			self.Graphs.Create_Chart(ChartName, self.Generate_Date_Range(FiscalYear),
					sheetname=self.CalculationSheet + " " + FiscalYear)
			# Add the indvidual series to the chart
			i = self.IndexHeadings.index(stat) * (len(self.DataHeadings) * (len(self.IndexHeadings) - 1)) # Add a sub-offset
			for col in range(i + ColOffset, i + ColOffset + len(self.DataHeadings)):
				self.Graphs.Add_Series(ChartName, self.Generate_Range(FiscalYear, col), serieslabels=True)
			# Apply the templated style, reapply attributes like ylabel
			self.Graphs.Apply_Template(ChartName, chartpath,
					ylabel=ylables[self.IndexHeadings.index(stat)])
			# Apply a fix to the final dates on the graph
			if graphenddate:
				self._Correct_Graph_Axis(ChartName, graphenddate)
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

# No coupling with the ORS calculator
class ORSSheets(ORSPlots):
	InputSheet = "Input"
	CalculationSheet = "Calculation"
	StatsSheet = "Statistics"
	OutputFileName = "Current Month SAIDI and SAIFI"
	DateOrigin = pos(row=4, col=1) # The postion to place the fist date

	def __init__(self, xlInstance):
		self.Sheet = Sheet(xlInstance)

	def YTD_Table(self, *params):
		"""Use a template table to represent the results.
		Create a YTD results table for a given date (datetime object)."""
		# The sheet with the (year) date may not have been created, so we can't rely on reading data from it
		#year = self.ORS._get_fiscal_year(date)
		#startdate = datetime.datetime(year, 4, 1)
		#enddate = datetime.datetime(year+1, 3, 31)
		#delta_days = (enddate - startdate).days +  1
		#Dates = [startdate + datetime.timedelta(days=i) for i in range(delta_days)]
		#BasicDate = datetime.datetime(date.year, date.month, date.day)
		#index = Dates.index(BasicDate)

		params = iter(params)
		out = params.next()
		for p in params:
			out = self.Merge_Dictionaries(out, p)

		template = Template(self.Sheet, r"C:\Users\sdo\Documents\Research and Learning\Git Repos\SAIDI-SAIFI-Calculator\Data\Templates.xlsx")
		template.Place_Template("Rob", self.Sheet._getCell(self.OutputFileName, 1 ,1))
		template.Set_Values(out)
		template.Auto_Fit() # Handles the closing of the template file

	def YTD_Book(self, *params):
		"""Create a new workbook with PNL Commercial Summary table"""
		params = iter(params)
		out = params.next()
		for p in params:
			out = self.Merge_Dictionaries(out, p)

		
		template = Template(Sheet(xlInstance), r"C:\Users\sdo\Documents\Research and Learning\Git Repos\SAIDI-SAIFI-Calculator\Data\Templates.xlsx")
		template.Place_Template("Rob", self.Sheet._getCell(self.OutputFileName, 1 ,1))
		template.Set_Values(out)
		template.Auto_Fit() # Handles the closing of the template file

	def Rename_Network(self, networkname):
		name = ""
		if networkname == "TPCO":
			name = "TPC"
		elif networkname == "OTPO":
			name = "OJV"
		elif networkname == "ELIN":
			name = "EIL"
		return name

	def Merge_Dictionaries(self, x1, x2):
		"""Returns the ersults of merging two dictionaries.
		x1 is master so, any duplicate keys in x2 will be overwritten"""
		cpy = x2.copy()
		cpy.update(x1)
		return cpy

# Coupling with the ors calculator
class ORSOutput(ORSSheets):
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
		self.Sheet.rmvSheet(removeList=[self.OutputFileName])
		self.Sheet.addSheet(self.OutputFileName)

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

	def Generate_Values(self, enddate, startdate=None):
		if not startdate:
			year = self.ORS._get_fiscal_year(enddate) - 1 # Function returns the year that the the 31st of march occurs on
			startdate = datetime.datetime(year, 4, 1)
		else:
			year = self.ORS._get_fiscal_year(startdate) - 1

		Dates = self.Generate_Dates(startdate, enddate)
		Dates = [Date[0] for Date in Dates]
		# Calculate the cumulative sum of SAIDI and SAIFI 
		# (planned, unplanned, unplanned normed) for the given dates
		SAIDI_, SAIFI_ = self._Calc_Rows(Dates, self.ORS)
		netname = self.Rename_Network(self.ORS.networknames[0])
		#params = [netname + "_" + param for param in params]
		
		name = self.Rename_Network(self.NetworkName) + "_"
		params = {}
		params[name+"DATE_END"] = enddate
		params[name+"CUST_NUM"] = self.ORS._get_total_customers(enddate)
		# Sum the columns in this matrix
		params[name+"SAIDI_NORMED_OUT"] = np.sum(SAIDI_, 0)[2]
		params[name+"SAIFI_NORMED_OUT"] = np.sum(SAIFI_, 0)[2]
		params[name+"SAIDI_UNPLANNED"] = np.sum(SAIDI_, 0)[1]
		params[name+"SAIFI_UNPLANNED"] = np.sum(SAIFI_, 0)[1]
		params[name+"SAIDI_PLANNED"] = np.sum(SAIDI_, 0)[0]
		params[name+"SAIFI_PLANNED"] = np.sum(SAIFI_, 0)[0]

		# Com Com Interpolations (could use np.linspace)
		SAIDI_TARGET, SAIFI_TARGET = self.ORS._get_CC_stats("TARGET")
		num_days = (datetime.datetime(year+1, 3, 31) - datetime.datetime(year, 3, 31)).days
		x_days = (enddate - datetime.datetime(year, 4, 1)).days
		SAIDI_M = SAIDI_TARGET/num_days
		SAIFI_M = SAIFI_TARGET/num_days
		params[name+"CC_SAIDI_YTD"] = SAIDI_M * (1 + x_days)
		params[name+"CC_SAIFI_YTD"] = SAIFI_M * (1 + x_days)

		for key, val in params.iteritems():
			print key, "\t", val
		return params