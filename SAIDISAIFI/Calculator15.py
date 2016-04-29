'''
Created on 30/3/2016

@author: Sean D. O'Connor
'''
import csv, os, datetime, calendar, operator, numpy as np
from tabulate import tabulate

# Aux class for reading the ORS via ODBC
from CalculatorAux import ODBC_ORS
from Constants import *

class ORSCalculator(object):
	"""Calculator for the 2015 - 2020 period. After this period updates from the new determination
	will need to be applied"""
	def __init__(self, avrgCustomerNum, networkname,
				 startDate=None, endDate=None):
		self.FeederCol = 0
		self.ORSNumCol = 1
		self.LinkedORSCol = 2
		self.DateCol = 3
		self.TimeCol = 4
		self.NetworkCol = 5
		self.FaultTypeCol = 6
		self.CusmMinCol = 7
		self.UniqueICPCol = 8
		self.AutoReclose = 9
		
		# Boundary values should be calculated as the 23rd largest DAILY SAIDI (i.e. grouped days) for UNPLANNED outages
		#self.BoundryCalcPeriod = (2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014) # Year ending, e.g. 31/03/2005 for 2005
		self.CCStartDate = datetime.datetime(2004, 4, 1) # start of determination period 
		self.CCEndDate = datetime.datetime(2014, 3, 31) # end of the determination period
		self.BoundryCalcPeriod = [year for year in range(self._get_fiscal_year(self.CCStartDate), self._get_fiscal_year(self.CCEndDate) + 1)]
		
		if not startDate and not endDate:
			self.startDate = datetime.datetime(2004, 4, 1) # We need this start date or the boundary values will be incorrect
			self.endDate = datetime.datetime(2016, 3, 31) #datetime.datetime(2015, 3, 31) # 2015
		else:
			self.startDate = startDate #datetime.datetime(startDate.year-1, 4, 1)
			self.endDate = endDate #datetime.datetime(endDate.year, 3, 31)
		self.deltaDay = datetime.timedelta(1)
		
		self.networknames = [network.strip() for network in networkname.split(',')]
		self.outFolder = FILE_DIRS.get(self.networknames[0], r"C:\Users\sdo\Documents\SAIDI and SAIFI\Fed 2016 SAIDI & SAIFI\tests")
		self.ORSout = "\ORS - CC format.csv"
		self.ORSin = "\ORS - Raw.csv"
		self.CCin = "ComCom - Raw.csv"
		
		# The point from which all 3 ICP connections were counted
		# Orginally this was set at 1/4/2010, but two outages in EIL on 3/4/2009 and 8/8/2009 made me rethink the date
		if "OTPO" in self.networknames:
			self.threeICPStartDate = datetime.datetime(2010, 4, 1) # OJV
		else:
			self.threeICPStartDate = datetime.datetime(2009, 4, 1) # EIL, TPC
		
		# How different SAIDI figures must be before triggering a flag (seen in STDIO)
		self.SAIDItolerance = 0.000001
		self.SAIFItolerance = 0.000001
		self.Count3OrLessICPs = False # Choose not to discard <= 3 ICPs prior to 2010 (year ending 2011 and on = must be counted)
		# False = Don't count faults with 3 or less ICPs
		self.LinkOutages = True # (True) means that outages with same linked ORS# are grouped - Default, (False) no linking so the outage occurs on the day it was recorded
		
		try: 
			os.remove(self.outFolder + self.ORSout) # Remove the previously generated files
		except:
			pass
			
		# Initialise some class variables
		self.boundarySAIDIValue = 0
		self.boundarySAIFIValue = 0
		
		self.avrgCustomerNum = avrgCustomerNum

		# We need to know individual ICPs, so we can exldue outages affecting 3 or less (prior to a a given date)
		self.AllFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		self.GroupedAllFaults = {} # key = date, values = [SAIDI, SAIFI]
		
		self.PlannedFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		self.GroupedPlannedFaults = {} # key = date, values = [SAIDI, SAIFI]
		
		self.UnplannedFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count. Feeder]
		self.GroupedUnplannedFaults = {} # key = date, values = [SAIDI, SAIFI]
		
		self.table = [] # Table used for tabulating results
		
		print "\nRunning: %s network calculations" % str(self.networknames)
		
	def remove_file(self, filename):
		"""removes a file.
		@return: The name of the file that was removed (even if it was not removed)"""
		if not os.path.isfile(filename):
			filename = os.path.join(self.outFolder, filename)
		try:
			os.remove(filename)
		except:
			print 'File (%s) does not exist' % filename
		return filename
		
	def _set_boundary_values(self):
		'''Recalls the boundary SAIDI and SAIFI values (23 largest values in the unplanned outages)
		from a dictionary and returns the boundary values for SAIDI and SAIFI for all UNPLANNED outages'''
		dic = {}
		# We only calculate boundary values in the 5 year period (1/04/2004 - 31/03/2014)
		for date in self.GroupedUnplannedFaults:
			#if self._get_fiscal_year(self._date_to_str(date)) in self.BoundryCalcPeriod:
			if self._get_fiscal_year(date) in self.BoundryCalcPeriod:
				dic[date] = self.GroupedUnplannedFaults.get(date)
		
		self.boundarySAIDIValue = self.nth_largest(23, dic, 0) # Boundary SAIDI
		self.boundarySAIFIValue = self.nth_largest(23, dic, 1) # Boundary SAIFI
	
		return self.boundarySAIDIValue, self.boundarySAIFIValue
		
	def generate_fault_dict_csv(self, FaultType, IndividualFaults, DailyFaults):
		"""Generates three dictionaries of faults: planned, unplanned, all.
		Using a csv file as the data source"""
		with open(self.outFolder + self.ORSin, 'rb') as orscsvfile:
			ORS = csv.reader(orscsvfile)
			ORS.next() # Incriment the file pointer past the header
			for record in ORS:
				self._process_fault_record(record, FaultType, IndividualFaults)
				
		self._group_same_day(IndividualFaults, DailyFaults)
		
	def generate_fault_dict(self, FaultType, IndividualFaults, DailyFaults):
			"""Generates three dictionaries of faults: planned, unplanned, all.
			Using the ORS via an ODBC connection as the datasource"""
			# Add handlers for criteria in the SQL
			ORS = ODBC_ORS()
			qryrows = ORS.get_query_results()
			
			# There is no header row when reading from the DB
			# The record and CSV file MUST follow the expected layout!
			for record in qryrows:
				self._process_fault_record(record, FaultType, IndividualFaults)
			
			# Warning, a hack: a way to read in the transpower outages for the purchased OJV/transpower assets
			try:
				with open(os.path.join(FILE_DIRS.get("GENERAL"), "EXTRA RECORDS.csv"), 'rb') as orscsvfile:
					ORS = csv.reader(orscsvfile)
					ORS.next() # Incriment the file pointer past the header
					for record in ORS:
						# As long as the extra fault is a PNL class outage it will be counted (also autoreclose='N")
						self._process_fault_record(record, FaultType, IndividualFaults)
			except IOError:
				# The network doesn't have any extra records
				pass
					
			self._group_same_day(IndividualFaults, DailyFaults)
		
	def _process_fault_record(self, record, FaultType, IndividualFaults):
		"""Process an individual record read from the database. Decide if it
		is a new fault, linked to another fault and what date a group collection
		of faults should fall on. Handles the linking of outages and creating
		new records."""
		# Only count faults for the network (specfied by name) we are testing
		if record[self.NetworkCol] in self.networknames:
			# Get the stats for a new record in the ORS
			Date, SAIDI, SAIFI = self.get_fault(record, FaultType)
			ICPcount = record[self.UniqueICPCol]
			Feeder = record[self.FeederCol] # TODO: fix this line, make sure that empty string or NONE is handeled
			if type(ICPcount) is not int:
				try:
					ICPcount = float(record[self.UniqueICPCol].replace(',', ''))
				except:
					ICPcount = 0
					#print "Bad data. No Unique ICPs found in", record
			
			# Handle the linking of faults
			if self.LinkOutages:
				# Get the stats for an old record in the ORS, or 0 if nothing exists
				Date0, SAIDI0, SAIFI0, ICPcount0, Feeder0 = IndividualFaults.get(record[self.LinkedORSCol], [Date, 0, 0, 0, Feeder])
				if record[self.ORSNumCol] == record[self.LinkedORSCol]:
					# The outage is always brought back to the date where the linked ORS# and ORS# are the same (not the max or min. date) - should it be the min. date?
					Date0 = Date
				if SAIDI0 + SAIDI != 0 and SAIFI0 + SAIFI != 0:
					# Only add non-zero records to the dictionaries, i.e. only faults are added (no clear days with 0 SAIDI/SAIFI)
					IndividualFaults[record[self.LinkedORSCol]] = [Date0, SAIDI0 + SAIDI, SAIFI0 + SAIFI, ICPcount0 + ICPcount, Feeder0]
			else:
				if SAIDI != 0 and SAIFI != 0:
					IndividualFaults[record[self.ORSNumCol]] = [Date, SAIDI, SAIFI, ICPcount, Feeder0]
	
	def _group_same_day(self, dic, sortedDic):
		# sortedDic: key = date, Values = [SAIDI, SAIFI]
		'''Look through each unique Linked ORS number dictionary and group same dates (days) together'''
		#currentDate = self.startDate
		currentDate = min(self.CCStartDate, self.startDate)
		
		while currentDate != max(self.CCEndDate, self.endDate) + self.deltaDay:
			for faultkey in dic: # faultkey is the LinkedORS#
				if dic.get(faultkey)[0] == currentDate: # There is a matching fault date and current date i.e. the fault occured on this date                    
					buff = sortedDic.get(currentDate)
					if not buff: # Test if 'buff' is nonetype. This happens when the dictionary returns null 
						buff = [0, 0]
					if self.Count3OrLessICPs or currentDate >= self.threeICPStartDate:
						sortedDic[currentDate] = [buff[0] + dic.get(faultkey)[1], buff[1] + dic.get(faultkey)[2]]
					else:
						try:
							if dic.get(faultkey)[3] > 3.0 and currentDate < self.threeICPStartDate:
								# Don't count faults that defined by linked ID affect less than or equal to 3 ICPs
								sortedDic[currentDate] = [buff[0] + dic.get(faultkey)[1], buff[1] + dic.get(faultkey)[2]]                     
						except:
							print "String to float cast failed. No valid data for: ", currentDate
							print dic.get(faultkey)
			currentDate += self.deltaDay

	def _group_same_feeder(self, dic, sortedDic, startdate, enddate):
		"""Group outages by day and by feeder, so the sorted dictinary will contain 
		dates of individual feeder outages per day."""
		# {date : {"feeder 1" : [SAIDI, SAIFI], "feeder 2" : [SAIDI, SAIFI], ...}}
		# These are raw SAIDI/SAIFIs i.e. there are no boundary value cappings applied

		# Loop through the domain [startdate, enddate], so if start and end date re the same you will get one days worth of data back
		while startdate < (self.deltaDay + enddate):
			for linkedORSNum in dic:
				Date, SAIDI, SAIFI, UniqueICPs, Feeder = dic.get(linkedORSNum)
				if Date == startdate:
					Feeders = sortedDic.get(startdate, {})
					SAIDI0, SAIFI0 = Feeders.get(Feeder, [0, 0])
					Feeders[Feeder] = [SAIDI0 + SAIDI, SAIFI0 + SAIFI]
					sortedDic[startdate] = Feeders
			startdate += self.deltaDay

	
	def _get_indicies(self, currentDate, faultType, applyBoundary=True):
		'''Recalls the SAIDI and SAIFI values from a dictionary. 
		Handles unplanned outages and returns 
		min(Unplanned Boundary Value, Daily Value) for unplanned 
		faults only.
		The dictionaries being used expect a date as the key i.e. results are stored on a per day resolution'''
		SAIDIp = 0 # planned
		SAIFIp = 0
		SAIDIup = 0 # unplanned
		SAIFIup = 0
		DefaultValues = [0, 0]

		if faultType == "planned" or faultType == "all":
			dic = self.GroupedPlannedFaults
			#for key in dic:
			#    if key == currentDate:
			#        SAIDIp = dic.get(key)[0]
			#        SAIFIp = dic.get(key)[1]
			#        break
			SAIDIp = dic.get(currentDate, DefaultValues)[0]
			SAIFIp = dic.get(currentDate, DefaultValues)[1]
		
		# Boundary values only apply to unplanned outages
		if faultType == "unplanned" or faultType == "all":
			dic = self.GroupedUnplannedFaults
			#for key in dic:
				#if key == currentDate:
			if applyBoundary:
				SAIDIup = min([dic.get(currentDate, DefaultValues)[0], self.boundarySAIDIValue])
				SAIFIup = min([dic.get(currentDate, DefaultValues)[1], self.boundarySAIFIValue])
			else:
				# Retrive figures with no capping (or "normalising")
				SAIDIup = dic.get(currentDate, DefaultValues)[0]
				SAIFIup = dic.get(currentDate, DefaultValues)[1]
				#    break
				
		return SAIDIp+SAIDIup, SAIFIp+SAIFIup
					
	def get_fault(self, record, faultType):
		"""This method will only count PNL related outages, and discard all others -- this is not quite working yet"""
		date = record[self.DateCol]
		if type(date) is str: 
			date = self._str_to_date(record[self.DateCol]) # date is a string object (from csv), covert it --> datetime object
		assert type(date) is datetime.datetime, "Error: the date for record %s is not valid" % str(record)
		SAIDI, SAIFI = 0, 0
		weight = 0
		if record[self.FaultTypeCol] == "Planned - PowerNet":
			weight = 0.5
		elif record[self.FaultTypeCol] == "Unplanned - PowerNet":
			weight = 1
		
		try:
			# Handle strings from the CSV source
			SAIDI = weight * float(record[self.CusmMinCol].replace(',', '')) / self._get_total_customers(record[self.DateCol])
			SAIFI = weight * float(record[self.UniqueICPCol].replace(',', '')) / self._get_total_customers(record[self.DateCol])
		except AttributeError:
			# Handle ints/floats from the ODBC source
			try:
				SAIDI = weight * float(record[self.CusmMinCol]) / self._get_total_customers(record[self.DateCol])
				SAIFI = weight * float(record[self.UniqueICPCol]) / self._get_total_customers(record[self.DateCol])
			except TypeError:
				if faultType == "all":
					print "There is data missing for this record: %s. No SAIDI or SAIFI value could be determined." % (record)
		except ValueError:
			if faultType == "all":
				print "There is data missing for this record: %s. No SAIDI or SAIFI value could be determined." % (record)
		
		# Zero out any values that don't belong in the planned/unplanned set respectively,
		# do this so we can have planned, unplanned, and all outage dictionaries 
		if faultType == "planned" and record[self.FaultTypeCol] == "Unplanned - PowerNet":
			SAIDI = 0
			SAIFI = 0
		elif faultType == "unplanned" and record[self.FaultTypeCol] == "Planned - PowerNet":
			SAIDI = 0
			SAIFI = 0
			
		# Ignore autoreclose events (i.e. events under 1 minute)
		if record[self.AutoReclose] == "Y":
			SAIDI = 0
			SAIFI = 0
		
		return date, SAIDI, SAIFI
	
	def _str_to_date(self, date):
		date = date.split('/')
		return datetime.datetime(int(date[2]), int(date[1]), int(date[0]))
	
	def _date_to_str(self, date):
		month = date.month
		if month <= 9:
			month = '0' + str(month)
		else:
			month = str(month)
		return str(date.day) + '/' + month + '/' + str(date.year)
	
	def _get_total_customers(self, date):
		"""Takes the date formatted as either a datetime object or string.
		Returns -1 if no ICPs are specfied for a given year."""
		try:
			fiscalYear = self._get_fiscal_year(date)
		except:
			fiscalYear = self._get_fiscal_year(self._date_to_str(date))
		
		return self.avrgCustomerNum.get(fiscalYear, -1)
	
	def nth_largest(self, n, iterobject, valueIndex):
		#return sorted(iterobject.iteritems(), key=operator.itemgetter(1), reverse=False)[:n][n-1][1][valueIndex]
		''' valueindex=0 for SAIDI, or =1 for SAIFI'''
		if valueIndex == 1: # SAIFI
			return sorted(iterobject.iteritems(), key=lambda e: e[1][1], reverse=True)[n-1][1][1]
		elif valueIndex == 0: # SAIDI
			return sorted(iterobject.iteritems(), key=operator.itemgetter(1), reverse=True)[n-1][1][0]
	
	def _month_ends(self, years):
		'''Create a list of datetime objects that are the final day of every month for all the
		years found in the input "years" list object'''
		endOfMonthDates = []
		for year in years:
			for month in range(1, 12 + 1):
				endOfMonthDates.append(datetime.datetime(year, month, calendar.monthrange(year, month)[1]))
		return endOfMonthDates

	def _get_fiscal_year(self, date):
		'''Returns the fiscal year end (31/03/XXXX) for a particular date'''
		try:
			year = int(date.split('/')[2])
			if int(date.split('/')[1]) - 3 > 0:
				year += 1
		except:
			year = date.year
			if date.month - 3 > 0:
				year += 1
			#date = self._date_to_str(date)
			#year = int(self._date_to_str(date).split('/')[2])

		return year
				
	def generate_fiscal_year_ends(self):
		'''Generate the last day of each fiscal year'''
		endDates = []
		#for year in range(2005, 2005 + 10):
		LastYear = self.endDate.year
		if self.endDate.month > 3:
			# We have crossed into the next fiscal year period
			LastYear += 1
		for year in range(self.startDate.year, LastYear + 1):
			endDates.append(datetime.datetime(year, 3, 31))
		return endDates
	
	def generate_calender_year_ends(self):
		'''Generate the last day of each calendar year'''
		endDates = []
		#for year in range(2005, 2005 + 10):
		for year in range(self.startDate.year, self.endDate.year + 1):
			endDates.append(datetime.datetime(year, 12, 31))
		return endDates
	
	def last_day_of_month(self, any_day):
		"""
		@param any_day: A datetime object
		@return: A datetime object that is the last date of month for the supplied date"""
		next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
		return next_month - datetime.timedelta(days=next_month.day)
	
	def generate_month_ends(self):
		'''Generate the last day of each month'''
		endDates = []
		StartMonth = self.startDate.month
		EndMonth = 12
		EndYear = self.endDate.year
		for year in range(self.startDate.year, self.endDate.year + 1):
			if year == EndYear:
				EndMonth = self.endDate.month
			for month in range(StartMonth, EndMonth + 1):
				endDates.append(self.last_day_of_month(
					datetime.datetime(year, month, 1)))
				StartMonth = 1
		return endDates
	
	def generate_day_endsOLD(self, startDate, endDate):
		'''Generate each day in the study period'''
		endDates = []
		StartDay = startDate.day
		EndDay = endDate.day
		StartMonth = startDate.month
		EndMonth = 12
		EndYear = endDate.year
		
		for year in range(startDate.year, endDate.year + 1):
			if year == EndYear:
				EndMonth = endDate.month
			for month in range(StartMonth, EndMonth + 1):
				EndDay = self.last_day_of_month(datetime.datetime(year, month, 1)).day
				if year == EndYear and month == EndMonth:
					EndDay = endDate.day
				for day in range(StartDay, EndDay + 1):
					endDates.append(
						datetime.datetime(year, month, day))
				StartMonth = 1
		return endDates
		
	def generate_day_ends(self, startDate, endDate):
		"""A poosible replacment function?? Seems like it's a lot more simple"""
		dt = datetime.timedelta(days=1)
		dayends = []
		while startDate <= endDate:
			dayends.append(startDate)
			startDate += dt
		return dayends
		
	def generate_stats(self, check=True):
		'''Populates all the dictionaries with information on SAIDI and SAIFI for each 
		fault type'''
		self.generate_fault_dict('planned', self.PlannedFaults, self.GroupedPlannedFaults)
		self.generate_fault_dict('unplanned', self.UnplannedFaults, self.GroupedUnplannedFaults)
		self.generate_fault_dict('all', self.AllFaults, self.GroupedAllFaults)
		assert len(self.AllFaults) != 0, "No faults found! The supplied network name is proably incorrect."
		self._set_boundary_values()
		
		# A trivial check to make sure that all faults are either planned or unplanned (class B or C type outages)
		# and placed into the right dictionaries
		if check:
			#currentDate = self.startDate
			currentDate = min(self.startDate, self.CCStartDate)
			while currentDate <= max(self.endDate, self.CCEndDate):                
				SAIDI, SAIFI = self._get_indicies(currentDate, "all") # Planned and unplanned
				plannedSAIDI, plannedSAIFI = self._get_indicies(currentDate, "planned") # Planned only
				unplannedSAIDI, unplannedSAIFI = self._get_indicies(currentDate, "unplanned") # Unplanned only
				assert plannedSAIDI + unplannedSAIDI == SAIDI, "ERROR: SAIDI sum does not match" 
				assert plannedSAIFI + unplannedSAIFI == SAIFI, "ERROR: SAIFI sum does not match"   
				currentDate += self.deltaDay

		print "Network statistics successfully generated!"
	
	def period_endings(self, resolution):
		resolutions = ("fiscal year", "calendar year", "month", "day") # Always display planned, unplanned, total
		#dMonth = datetime.timedelta()
		if resolution == resolutions[0]:
			PeriodEndings = self.generate_fiscal_year_ends()
		elif resolution == resolutions[1]:
			PeriodEndings = self.generate_calender_year_ends()
		elif resolution == resolutions[2]:
			PeriodEndings = self.generate_month_ends()
		elif resolution == resolutions[3]:
			PeriodEndings = self.generate_day_ends(self.startDate, self.endDate)
		else:
			PeriodEndings = ()
		return PeriodEndings
		
	def _get_CC_stats(self, arg):
		"""Returns the hard coded CC cacluated stats"""
		CCSAIDIVal = CC_Vals.get(self.networknames[0]).get("SAIDI_" + arg)
		CCSAIFIVal = CC_Vals.get(self.networknames[0]).get("SAIFI_" + arg)
		return CCSAIDIVal, CCSAIFIVal
	
	def _get_stats(self, arg):
		"""Generate statistics for the tables.
		@param arg: A string defining a known calculation methodoligy e.g. UBV, LIMIT, CAP...
		@return: A tuple of the SAIDI and SAIFI statistic calculated from the arg"""
		ArraySAIDI = []
		ArraySAIFI = []   
		#currentDate = datetime.datetime(2004, 4, 1) # start of determination period
		currentDate = self.CCStartDate
		#while currentDate <= datetime.datetime(2014, 3, 31): # end of the determination period
		while currentDate <= self.CCEndDate: # end of the determination period
			SAIDI, SAIFI = self._get_indicies(currentDate, "all")
			ArraySAIDI.append(SAIDI)
			ArraySAIFI.append(SAIFI)
			currentDate += self.deltaDay
		
		# Target
		SAIDI_target = sum(ArraySAIDI)/len(self.BoundryCalcPeriod)
		SAIFI_target = sum(ArraySAIFI)/len(self.BoundryCalcPeriod)
		
		# Collar - sample variance
		SAIDI_collar = SAIDI_target - 365**0.5 * np.std(ArraySAIDI, dtype=np.float64, ddof=1)
		SAIFI_collar = SAIFI_target - 365**0.5 * np.std(ArraySAIFI, dtype=np.float64, ddof=1)
		
		# Limit - sample variance
		SAIDI_limit = SAIDI_target + 365**0.5 * np.std(ArraySAIDI, dtype=np.float64, ddof=1)
		SAIFI_limit = SAIFI_target + 365**0.5 * np.std(ArraySAIFI, dtype=np.float64, ddof=1)  
		
		# Cap
		SAIDI_cap = SAIDI_limit
		SAIFI_cap = SAIFI_limit
		
		if arg in ("UBV"):
			#return self.boundarySAIDIValue, self.boundarySAIFIValue
			return self.Verify_CC_Stats(arg, self.boundarySAIDIValue, self.boundarySAIFIValue)
		elif arg in ("LIMIT"):
			#return SAIDI_limit, SAIFI_limit
			return self.Verify_CC_Stats(arg, SAIDI_limit, SAIFI_limit)
		elif arg in ("TARGET"):
			#return SAIDI_target, SAIFI_target
			return self.Verify_CC_Stats(arg, SAIDI_target, SAIFI_target)
		elif arg in ("CAP"):
			#return SAIDI_cap, SAIFI_cap
			return self.Verify_CC_Stats(arg, SAIDI_cap, SAIFI_cap)
		elif arg in ("COLLAR"):
			#return SAIDI_collar, SAIFI_collar
			return self.Verify_CC_Stats(arg, SAIDI_collar, SAIFI_collar)
	
	def Verify_CC_Stats(self, arg, SAIDIval, SAIFIval):
		"""Compare the generated stats to the offcial (saved) CC stats"""
		# order is important when specifying network names. Make sure the first network name 
		# in the string matches the compare network in the CC documents.
		
		# Get the hard coded CC stats
		CCSAIDIVal, CCSAIFIVal = self._get_CC_stats(arg)
		
		SAIDIi = 0
		SAIFIi = 0
		if self.ABS_Diff(1e-2, SAIDIval, CCSAIDIVal):
			SAIDIi = SAIDIval
		else:
			print "WARNING! The SAIDI (%s) values calculated differ from the CC values." % arg
			#SAIDIi = CCSAIDIVal
			SAIDIi = SAIDIval
		if self.ABS_Diff(1e-2, SAIFIval, CCSAIFIVal):
			SAIFIi = SAIFIval
		else:
			print "WARNING! The SAIFI (%s) values calculated differ from the CC values." % arg
			#SAIFIi = CCSAIFIVal
			SAIFIi = SAIFIval
		return SAIDIi, SAIFIi
	
	def ABS_Diff(self, tolerance, *args):
		"""Absolute differece.
		@return: True if the max. differce in the supplied arguments is less
		than the defiend tolerance, False otherwise."""
		if abs(max(args) - min(args)) > tolerance:
			return False
		return True
		
	def DA_Table(self, fileName, startdate, enddate):
		"""Create the Distrobution Automation (DA) anaylsis table
		based on unplanned outages only."""
		FilePath = self.remove_file(fileName)
		if not os.path.exists(os.path.dirname(FilePath)):
			try:
				os.makedirs(os.path.dirname(FilePath))
			except OSError as exc: # Guard against race condition
				if exc.errno != errno.EEXIST:
					raise

		# Group faults by day, then by feeder
		Feeder_Grouped_Outages = {} # These are percentages, not actual absolute values
		Date_Grouped_Outages = {} # This includes outage affecting 3 or less ICPs, so self.GroupedUnplannedFaults.get(Date) can return None
		self._group_same_feeder(self.UnplannedFaults, Date_Grouped_Outages, startdate, enddate)
		Table = []
		Headings = ["Feeder Name", "SAIDI %", "SAIFI %"]
		SAIDItot, SAIFItot = 0, 0
		# Group all feeders with the same name
		for Date, Feeders in Date_Grouped_Outages.iteritems():
			SAIDIDay, SAIFIDay = self.GroupedUnplannedFaults.get(Date, [0, 0]) # Return 0 for records we can't find (3 or less ICPs)
			SAIDItot += SAIDIDay
			SAIFItot += SAIFIDay
			for Feeder, Stats in Feeders.iteritems():
				SAIDI0, SAIFI0 = Feeder_Grouped_Outages.get(Feeder, [0, 0])
				#SAIDI, SAIFI = Stats[0]/SAIDIDay, Stats[1]/SAIFIDay
				SAIDI, SAIFI = Stats[0], Stats[1] # The currently selected feeders contribution
				Feeder_Grouped_Outages[Feeder] = [SAIDI0+SAIDI, SAIFI0+SAIFI]

		# Now that the feeders are grouped by name, print their combined SAIDI and SAIFI
		for Feeder, Stats in Feeder_Grouped_Outages.iteritems():
			Table.append([Feeder, Stats[0]/SAIDItot*100, Stats[1]/SAIFItot*100])

		with open(FilePath, "a") as results_file:
			results_file.write(tabulate(Table, Headings,  tablefmt="orgtbl", floatfmt=".5f", 
				numalign="right")) 

	def display_stats(self, resolution, fileName):
		"""Method that creates tables for displaying data
		Replacement method for annual_stats"""
		self.table = []
		headers = ["Date Ending", "|", 
				   "Unique ICPs", "|", 
				   "Linked ORS #", "|", 
				   "Planned SAIDI", "Unplanned SAIDI", "Total SAIDI", "|", 
				   "Planned SAIFI", "Unplanned SAIFI", "Total SAIFI"]
		FilePath = self.remove_file(fileName)
		if not os.path.exists(os.path.dirname(FilePath)):
			try:
				os.makedirs(os.path.dirname(FilePath))
			except OSError as exc: # Guard against race condition
				if exc.errno != errno.EEXIST:
					raise
		
		PeriodEndings = self.period_endings(resolution)
		# Totals for the period resolution (year, month...)
		Total_UP_SAIDI_Period = 0
		Total_P_SAIDI_Period = 0
		Total_UP_SAIFI_Period = 0
		Total_P_SAIFI_Period = 0
		
		# currentDate = self.startDateDisplay
		# self.endDateDisplay
		currentDate = self.startDate
		while currentDate <= self.endDate:
			plannedSAIDI, plannedSAIFI = self._get_indicies(currentDate, "planned")
			unplannedSAIDI, unplannedSAIFI = self._get_indicies(currentDate, "unplanned")
			Total_P_SAIDI_Period += plannedSAIDI
			Total_UP_SAIDI_Period += unplannedSAIDI
			Total_P_SAIFI_Period += plannedSAIFI
			Total_UP_SAIFI_Period += unplannedSAIFI
			
			if resolution == "outage":                
				# Dictionary values: [Date0, SAIDI0, SAIFI0, ICPcount0]
				# Planned events first
				for LinkedORSNum in self.PlannedFaults:
					Record = self.PlannedFaults.get(LinkedORSNum)
					if Record[0] == currentDate:
						self._add_table_record(currentDate, self._get_total_customers(currentDate),
						   LinkedORSNum,
						   Record[1], 0,
						   Record[2], 0)
				# Unplanned events second
				for LinkedORSNum in self.UnplannedFaults:
					Record = self.UnplannedFaults.get(LinkedORSNum)
					if Record[0] == currentDate:
						self._add_table_record(currentDate, self._get_total_customers(currentDate),
						   LinkedORSNum,
						   0, Record[1],
						   0, Record[2])
						
			elif currentDate in PeriodEndings:
				self._add_table_record(currentDate, self._get_total_customers(currentDate),
					   "N/A",
					   Total_P_SAIDI_Period, Total_UP_SAIDI_Period,
					   Total_P_SAIFI_Period, Total_UP_SAIFI_Period)
				Total_UP_SAIDI_Period = 0
				Total_P_SAIDI_Period = 0
				Total_UP_SAIFI_Period = 0
				Total_P_SAIFI_Period = 0
			
			currentDate += self.deltaDay

		# Write the ORS records to the file
		with open(FilePath, "a") as results_file:
			results_file.write("Network: %s\n" % str(self.networknames))
			results_file.write(tabulate(self.table, headers, floatfmt=".5f", numalign="right"))
		
		# Write the (CC) parameters table to the bottom of the file
		if resolution == "fiscal year" or resolution == "calendar year":
			UBV = self._get_stats("UBV")
			LIMIT = self._get_stats("LIMIT")
			TARGET = self._get_stats("TARGET")
			COLLAR = self._get_stats("COLLAR")
			CAP = self._get_stats("CAP")
			headers = ["Parameter", "SAIDI", "SAIFI"]
			table = [["Unplanned Boundary", UBV[0], UBV[1]], 
					 ["Limit",  LIMIT[0],   LIMIT[1]], 
					 ["Target", TARGET[0],  TARGET[1]], 
					 ["Collar", COLLAR[0],  COLLAR[1]], 
					 ["Cap",    CAP[0],     CAP[1]]]
			with open(FilePath, "a") as results_file:
				results_file.write("\n\nDaily Parameters for the period 1/4/2015 - 31/3/2020:\n")
				results_file.write("(Based on outages 1/04/2004 - 31/03/2013)\n")
				results_file.write(tabulate(table, headers,  tablefmt="orgtbl", floatfmt=".5f", 
											numalign="right")) # add a style
			
	def _add_table_record(self, date, NumICPs, LinkedORSNo, pSAIDI, upSAIDI, 
												pSAIFI, upSAIFI):
		"""Adds a row to the table"""
		# E.g. formatting floats: "{0:.5f}".format(monthlySAIDI)
		monthAbbr = {month: abbr for month,abbr in enumerate(calendar.month_abbr)}
		self.table.append([self._date_to_str(date) + ' (' + monthAbbr.get(date.month) + ')', "|", str(round(NumICPs, 1)),
						   "|", LinkedORSNo,
						   "|", pSAIDI, upSAIDI, pSAIDI+upSAIDI, 
						   "|", pSAIFI, upSAIFI, pSAIFI+upSAIFI])

