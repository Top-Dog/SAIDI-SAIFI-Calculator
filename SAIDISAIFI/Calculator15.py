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
				 startDate=None, endDate=None, **kwargs):
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
		self.boundarySAIDIValue = kwargs.get("boundarySAIDIValue", None)
		self.boundarySAIFIValue = kwargs.get("boundarySAIFIValue", None)
		
		self.avrgCustomerNum = avrgCustomerNum

		# We need to know individual ICPs, so we can exldue outages affecting 3 or less (prior to a a given date)
		self.AllFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		self.GroupedAllFaults = {} # key = date, values = [SAIDI, SAIFI, NumberOfOutageRecords]
		
		self.PlannedFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		self.GroupedPlannedFaults = {} # key = date, values = [SAIDI, SAIFI, NumberOfOutageRecords]
		
		self.UnplannedFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		self.GroupedUnplannedFaults = {} # key = date, values = [SAIDI, SAIFI, NumberOfOutageRecords]
		
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
		
		# Only update the boundary values if they're not set
		if self.boundarySAIDIValue is None:
			self.boundarySAIDIValue = self.nth_largest(23, dic, 0) # Boundary SAIDI
		if self.boundarySAIFIValue is None:
			self.boundarySAIFIValue = self.nth_largest(23, dic, 1) # Boundary SAIFI
	
		return self.boundarySAIDIValue, self.boundarySAIFIValue
		
	def generate_fault_dict_csv(self, FaultType, IndividualFaults, DailyFaults):
		"""Generates three dictionaries of faults: planned, unplanned, all.
		Using a csv file as the data source"""
		with open(self.outFolder + self.ORSin, 'rb') as orscsvfile:
			records = csv.reader(orscsvfile)
			records.next() # Incriment the file pointer past the header
			for record in records:
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
					records = csv.reader(orscsvfile)
					records.next() # Incriment the file pointer past the header
					for record in records:
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
			Feeder = record[self.FeederCol]
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
					# The outage is always brought back to the date where the linked ORS# and ORS# are the same (not the max. or min. date) - should it be the min. date?
					Date0 = Date
				if SAIDI0 + SAIDI != 0 and SAIFI0 + SAIFI != 0:
					# Only add non-zero records to the dictionaries, i.e. only faults are added (no clear days with 0 SAIDI/SAIFI)
					IndividualFaults[record[self.LinkedORSCol]] = [Date0, SAIDI0 + SAIDI, SAIFI0 + SAIFI, ICPcount0 + ICPcount, Feeder0]
			else:
				if SAIDI != 0 and SAIFI != 0:
					IndividualFaults[record[self.ORSNumCol]] = [Date, SAIDI, SAIFI, ICPcount, Feeder]
	
	def _group_same_dayOLD(self, dic, sortedDic):
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

	def _group_same_day(self, dic, sortedDic):
		# sortedDic: key = date, Values = [SAIDI, SAIFI, NumberOfOutageRecords]
		'''Look through each unique Linked ORS number dictionary and group same dates (days) together'''
		for LinkedORSNum, Buff in dic.iteritems():
			if Buff[0] >= min(self.CCStartDate, self.startDate) and Buff[0] <= max(self.CCEndDate, self.endDate): # The DB query already enforces a date criteria, so this line is probably not needed
				GroupedDays = sortedDic.get(Buff[0], [0, 0, 0])
				if self.Count3OrLessICPs or Buff[0] >= self.threeICPStartDate:
					sortedDic[Buff[0]] = [GroupedDays[0] + Buff[1], GroupedDays[1] + Buff[2], GroupedDays[2] + 1]
				else:
					try:
						if Buff[3] > 3.0 and Buff[0] < self.threeICPStartDate:
							# Don't count faults that defined by linked ID affect less than or equal to 3 ICPs
							sortedDic[Buff[0]] = [GroupedDays[0] + Buff[1], GroupedDays[1] + Buff[2], GroupedDays[2] + 1]                     
					except:
						print "String to float cast failed. No valid data for: ", Buff[0]
						print Buff

	def _group_same_feeder(self, dic, sortedDic, startdate, enddate):
		"""Group outages by day and by feeder, so the sorted dictionary will contain 
		dates of individual feeder outages per day."""
		# {date : {"feeder 1" : [SAIDI, SAIFI, ICP count, [List of ORS numbers]], "feeder 2" : [SAIDI, SAIFI, ICP count, [List of ORS numbers]], ...}}
		# These are raw SAIDI/SAIFIs i.e. there are no boundary value cappings applied

		# Loop through the domain [startdate, enddate], so if start and end date re the same you will get one days worth of data back
		while startdate < (self.deltaDay + enddate):
			for linkedORSNum in dic:
				Date, SAIDI, SAIFI, UniqueICPs, Feeder = dic.get(linkedORSNum)
				# Loop through every linked ORS No. so we don't need to use a "Grouped Unplanned Faults" dictionary.
				if Date == startdate:
					Feeders = sortedDic.get(startdate, {})
					SAIDI0, SAIFI0, ICPs0, FaultIDs = Feeders.get(Feeder, [0, 0, 0, []])
					Feeders[Feeder] = [SAIDI0 + SAIDI, SAIFI0 + SAIFI, ICPs0 + UniqueICPs, FaultIDs + [linkedORSNum]]
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
		DefaultValues = [0, 0, 0]

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

	def _get_num_faults(self, currentDate, faultType):
		"""Gets the raw data values for a particualr fault type
		@return: tuple (number of ICPs out by planned, number of ICPs out by unplanned) on a given day
		"""
		numfaults = 0
		# key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
		if faultType == "unplanned" or faultType == "all":
			for linkedORS, data in self.UnplannedFaults.iteritems():
				if data[0] == currentDate:
					numfaults += 1 #data[3]
		
		if faultType == "planned" or faultType == "all":
			for linkedORS, data in self.PlannedFaults.iteritems():
				if data[0] == currentDate:
					numfaults += 1 #data[3]
		return numfaults

	def get_capped_days(self, day, end):
		"""Return a list days that unplanned SAIDI or SAIFI exceedded UBVs"""
		UBVSAIDI = []
		UBVSAIFI = []
		while day <= end:
			SAIDI1, SAIFI1 = self._get_indicies(day, "unplanned", applyBoundary=False)
			SAIDI2, SAIFI2 = self._get_indicies(day, "unplanned", applyBoundary=True)
			outage_numbers = self._get_ubv_outages(day)
			if SAIDI1 != SAIDI2:
				UBVSAIDI.append([day.date(), SAIDI1, SAIDI2, str(outage_numbers)])
			if SAIFI1 != SAIFI2:
				UBVSAIFI.append([day.date(), SAIFI1, SAIFI2, str(outage_numbers)])
			day += self.deltaDay
		return UBVSAIDI, UBVSAIFI
		
					
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
		# OJV is a special case, deal to it here
		#if self.networknames[0] == "OTPO":
		#	self.boundarySAIDIValue = 13.2414436340332
		#	self.boundarySAIFIValue = 0.176474571228027
		
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
		
		# Collar - sample variance (Delta Degrees of Freedom should be set to 1 to eliminate bias from a sample std. deviation i.e. divide by N-1)
		SAIDI_collar = SAIDI_target - 365**0.5 * np.std(ArraySAIDI, dtype=np.float64, ddof=1)
		SAIFI_collar = SAIFI_target - 365**0.5 * np.std(ArraySAIFI, dtype=np.float64, ddof=1)
		
		# Limit - sample variance (Delta Degrees of Freedom should be set to 1 to eliminate bias from a sample std. deviation i.e. divide by N-1)
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
			print "WARNING! %s Diff(SAIDI (%s)) > tolerance" % (self.networknames, arg)
			if "OTPO" in self.networknames:
				SAIDIi = CCSAIDIVal
			else:
				SAIDIi = SAIDIval
		if self.ABS_Diff(1e-2, SAIFIval, CCSAIFIVal):
			SAIFIi = SAIFIval
		else:
			print "WARNING! %s Diff(SAIFI (%s)) > tolerance" % (self.networknames, arg)
			if "OTPO" in self.networknames:
				SAIFIi = CCSAIFIVal
			else:
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
		FilePath = self._rm_file(fileName)

		# Group faults by day, then by feeder
		Feeder_Grouped_Outages = {} # These are percentages, not actual absolute values
		Date_Grouped_Outages = {} # This includes outage affecting 3 or less ICPs, so self.GroupedUnplannedFaults.get(Date) can return None
		self._group_same_feeder(self.UnplannedFaults, Date_Grouped_Outages, startdate, enddate)
		Table = []
		Headings = ["Feeder Name", "SAIDI %", "SAIFI %"]
		SAIDItot, SAIFItot = 0, 0
		# Group all feeders with the same name
		for Date, Feeders in Date_Grouped_Outages.iteritems():
			SAIDIDay, SAIFIDay, NumberOfOutages = self.GroupedUnplannedFaults.get(Date, [0, 0, 0]) # Return 0 for records we can't find (3 or less ICPs)
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
			results_file.write("%s to %s\n" % (startdate, enddate))
			results_file.write(tabulate(Table, Headings,  tablefmt="orgtbl", floatfmt=".5f", 
				numalign="right"))

	def Outages_Feeder_Year(self, fileName, startdate, enddate):
		"""Create a table of outages by feeder, for the time period specified"""
		FilePath = self._rm_file(fileName)

		# Group faults by day, then by feeder
		Date_Grouped_Outages = {} # This includes outage affecting 3 or less ICPs, looks like: {Date : {"feeder 1" : [SAIDI, SAIFI, ICP count], "feeder 2" : [SAIDI, SAIFI, ICP count], ...}, Date : {...}, ...}
		Aggregate_Data = {} # Actual Values (Absolute)
		self._group_same_feeder(self.UnplannedFaults, Date_Grouped_Outages, startdate, enddate) # Modifies the Date_Grouped_Outages dictionary
		#self._group_same_feeder(self.PlannedFaults, Date_Grouped_Outages, startdate, enddate) # Modifies the Date_Grouped_Outages dictionary
		Table = [[]]
		Heading = []
		Heading1 = ["Year", "Feeder Name", "SAIDI", "SAIFI", "Cumulative No. ICPs"]
		SAIDISubtotal, SAIFISubtotal = 0, 0
		SAIDITotals, SAIFITotals = [], []
		# Aggregate all faults for a year (combines all the faults for a given feeder over a year)
		for Date, Feeders in Date_Grouped_Outages.iteritems():
			Year = self._get_fiscal_year(Date)
			if Year in Aggregate_Data:
				self._add_feeder_dicts(Aggregate_Data[Year], Feeders)
			else:
				Aggregate_Data[Year] = Feeders
				Heading += Heading1

		# Aggregate all the feeder names
		CombinedFeederNames = {} # How many faults records against each feeder
		CombinedYears = []
		for Year, Feeders in Aggregate_Data.iteritems():
			CombinedYears.append(Year)
			for Feeder in Feeders:
				# Overwrite any existing keys...
				CombinedFeederNames[Feeder] = CombinedFeederNames.get(Feeder, 0) + 1
		FeederNames = sorted([key for key in CombinedFeederNames])
		CombinedYears = sorted(CombinedYears)

		# Table type 1
		# Create blocks of stats for each year 
		# Indvidual feeder names for every column
		#blocks = []
		#for Year, Feeders in Aggregate_Data.iteritems():
		#	cols = []
		#	for Feeder, Stats in Feeders.iteritems():
		#		cols.append([Year, Feeder, Stats[0], Stats[1], Stats[2]])
		#	blocks.append(cols)
		#	#Table = self._add_table_cols(Table, block)

		# Table Type 2
		# Create blocks of stats for each feeder
		# The same feeder name on every row
		blocks = []
		Heading = ["Feeder Name"]
		Table = [[Feeder] for Feeder in FeederNames]
		for Year in CombinedYears:
			TableCol = []
			Heading += [str(Year-1) + " SAIDI", str(Year-1) + " SAIFI", str(Year-1) + " ORS No."]
			#HeadingDate = str(Year-1)
			Feeders = Aggregate_Data.get(Year, {})
			for Feeder in FeederNames:
				Stats = Feeders.get(Feeder, [0, 0, 0, []]) # SAIDI, SAIFI, ICP Nums, ORS Nums.
				TableCol.append([Stats[0], Stats[1], Stats[3]])
			blocks.append(TableCol)
		
		# Group all the column blocks together
		for block in blocks:
			Table = self._add_table_cols(Table, block)

		# Write the headings and table to the output file
		with open(FilePath, "a") as results_file:
			results_file.write("%s to %s\n" % (startdate, enddate))
			results_file.write(tabulate(Table, Heading,  
							   tablefmt="plain", floatfmt=".5f", numalign="right"))

		# Write the table to a csv file
		with open(os.path.splitext(FilePath)[0] + ".csv", "w") as csvfile:
			csvfile.write("%s to %s\n" % (startdate, enddate))
			for col in Heading:
				csvfile.write(col + ",")
			csvfile.write("\n")
			for row in Table:
				for col in row:
					if col is not None:
						csvfile.write('"' + str(col) + '"')
					else:
						csvfile.write('"' + '"')
					csvfile.write(",")

				csvfile.write("\n")


	def _add_feeder_dicts(self, dict1, dict2):
		"""Add to feeder dicts together, replaces dict1"""
		# {"feeder 1" : [SAIDI, SAIFI, ICP count, [ORS #s]], "feeder 2" : [SAIDI, SAIFI, ICP count, [ORS #s]], ...}
		# {"feeder 2" : [SAIDI, SAIFI, ICP count, [ORS #s]], "feeder 3" : [SAIDI, SAIFI, ICP count, [ORS #s]], ...}
		# Update existing keys in dict1
		for key, value in dict1.iteritems():
			SAIDI, SAIFI, ICPcount, FaultIDs = dict2.get(key, [0, 0, 0, []])
			dict1[key] = [value[0]+SAIDI, value[1]+SAIFI, value[2]+ICPcount, value[3]+FaultIDs]
		# Find the keys not in dict1, but that are in dict2, and add them to dict1
		diffKeys = set(dict2.keys()) - set(dict1.keys())
		for key in diffKeys:
			SAIDI, SAIFI, ICPcount, FaultIDs = dict2.get(key, [0, 0, 0, []])
			dict1[key] = [SAIDI, SAIFI, ICPcount, FaultIDs]

	def _add_table_cols(self, leftBlock, rightBlock):
		"""Adds two blocks of columns to each other"""
		Table = []
		leftBlockNull = len(leftBlock[0])*[""] # How many cols wide is the left block
		rightBlockNull = len(rightBlock[0])*[""] # How many cols wide is the right block
		rightRowsDim = len(rightBlock)
		leftRowsDim = len(leftBlock)

		if leftRowsDim - rightRowsDim >= 0:
			# Left block has more rows
			for i in range(rightRowsDim):
				Table.append(leftBlock[i] + rightBlock[i])
			for i in range(i+1, leftRowsDim):
				Table.append(leftBlock[i] + rightBlockNull)
		else:
			# Right part has more rows
			for i in range(leftRowsDim):
				Table.append(leftBlock[i] + rightBlock[i])
			for i in range(i+1, rightRowsDim):
				Table.append(leftBlockNull + rightBlock[i])
		return Table
			

	def _get_ubv_outages(self, date):
		"""Get all the unplanned outages that occur on the date of a UBV outage"""
		orsnums = []
		for LinkedORSNum, Dataset in self.UnplannedFaults.iteritems(): # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count. Feeder]
			if Dataset[0] == date:
				orsnums.append(LinkedORSNum)
		return orsnums

	def _rm_file(self, filename):
		# Remove any existing file in the same directory
		FilePath = self.remove_file(filename)
		if not os.path.exists(os.path.dirname(FilePath)):
			try:
				os.makedirs(os.path.dirname(FilePath))
			except OSError as exc: # Guard against race condition
				if exc.errno != errno.EEXIST:
					raise
		return FilePath
	
	def _table(self, filename, headings, table, newlines=0):
		# Create a new file
		with open(filename, "a") as results_file:
			results_file.write(tabulate(table, headings,  tablefmt="orgtbl", floatfmt=".5f", 
				numalign="right")) 
			for i in range(newlines+1):
				results_file.write("\n")

	def Capped_Outages_Table(self, filename, startdate, enddate):
		# Remove any existing file in the same directory
		filename = self._rm_file(filename)

		SAIDIDays, SAIFIDays = self.get_capped_days(startdate, enddate)

		# Build the SAIDI table
		Headings = ["Date", "Pre-Normalised (SAIDI)", "Normalised (SAIDI)", "ORS Number(s)"]
		self._table(filename, Headings, SAIDIDays, 2)

		# Build the SAIFI table
		Headings = ["Date", "Pre-Normalised (SAIFI)", "Normalised (SAIFI)", "ORS Number(s)"]
		self._table(filename, Headings, SAIFIDays)

	def display_stats(self, resolution, fileName):
		# Very slow method.. needs some tlc
		"""Method that creates tables for displaying data
		Replacement method for annual_stats"""
		self.table = []
		headers = ["Date Ending", "|", 
				   "Unique ICPs", "|", 
				   "Linked ORS #", "|", 
				   "Planned SAIDI", "Unplanned SAIDI", "Total SAIDI", "|", 
				   "Planned SAIFI", "Unplanned SAIFI", "Total SAIFI"]
		FilePath = self._rm_file(fileName)
		
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
			NumCustomers = self._get_total_customers(currentDate)
			
			if resolution == "outage":                
				# Dictionary values: [Date0, SAIDI0, SAIFI0, ICPcount0]
				# Planned events first
				#for LinkedORSNum in self.PlannedFaults:
				#	Record = self.PlannedFaults.get(LinkedORSNum)
				#	if Record[0] == currentDate:
				#		self._add_table_record(currentDate, self._get_total_customers(currentDate),
				#		   LinkedORSNum,
				#		   Record[1], 0,
				#		   Record[2], 0)
				for LinkedORSNum, Record in self.PlannedFaults.iteritems():
					if Record[0] == currentDate:
						self._add_table_record(currentDate, NumCustomers,
						   LinkedORSNum,
						   Record[1], 0,
						   Record[2], 0)

				# Unplanned events second
				#for LinkedORSNum in self.UnplannedFaults:
				#	Record = self.UnplannedFaults.get(LinkedORSNum)
				#	if Record[0] == currentDate:
				#		self._add_table_record(currentDate, self._get_total_customers(currentDate),
				#		   LinkedORSNum,
				#		   0, Record[1],
				#		   0, Record[2])
				for LinkedORSNum, Record in self.UnplannedFaults.iteritems():
					if Record[0] == currentDate:
						self._add_table_record(currentDate, NumCustomers,
						   LinkedORSNum,
						   0, Record[1],
						   0, Record[2])
						
			elif currentDate in PeriodEndings:
				self._add_table_record(currentDate, NumCustomers,
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
		monthAbbr = {month: abbr for month, abbr in enumerate(calendar.month_abbr)}
		self.table.append([self._date_to_str(date) + ' (' + monthAbbr.get(date.month) + ')', "|", str(round(NumICPs, 1)),
						   "|", LinkedORSNo,
						   "|", pSAIDI, upSAIDI, pSAIDI+upSAIDI, 
						   "|", pSAIFI, upSAIFI, pSAIFI+upSAIFI])


from DataStructures import OutageRecord

class ORSCalculatorEnhanced(ORSCalculator):
	def __init__(self):
		super(ORSCalculatorEnhanced, self)
		self.outageRecords = {}
		self.networknames = []
		self.OutageClasses = ["Unplanned - PowerNet", "Planned - PowerNet"]
		
		# Run the SQL Query and collect the outage records
		ORS = ODBC_ORS()
		query = ORS.load_sql("OutageDetail.sql")
		records = ORS.run_query(query)
		columnNames = ORS.get_column_names()
		for record in records:
			outageRecord = {}
			for columnName, value in zip(columnNames, record):
				outageRecord[columnName] = value
			
			# Only select the fault records we are interested in
			if outageRecord["Network"] in self.networknames and outageRecord["ClassDescription"] in self.OutageClasses:
				self.outageRecords[outageRecord.get("Out_Linked_Num")] = OutageRecord(outageRecord)

	def group_linked_outages(self, outages):
		"""Group all the outages by their linked ORS number."""
		# {linkedorsnumber : class<linkedoutages>}
		linkedOutages = {}
		for outage in outages:
			record = linkedOutages.get(outage.LinkedID)
			if record:
				linkedOutages[outage.LinkedID].append(outage)
			else:
				linkedOutages[outage.LinkedID] = [outage]

		return linkedOutages


	def group_outages_by_ors_num(self, outages):
		Outages = {}
		for outage.ID in outages:
			pass


	def group_outages_by_date(self, ghtei):
		pass


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
			Feeder = record[self.FeederCol]
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
					# The outage is always brought back to the date where the linked ORS# and ORS# are the same (not the max. or min. date) - should it be the min. date?
					Date0 = Date
				if SAIDI0 + SAIDI != 0 and SAIFI0 + SAIFI != 0:
					# Only add non-zero records to the dictionaries, i.e. only faults are added (no clear days with 0 SAIDI/SAIFI)
					IndividualFaults[record[self.LinkedORSCol]] = [Date0, SAIDI0 + SAIDI, SAIFI0 + SAIFI, ICPcount0 + ICPcount, Feeder0]
			else:
				if SAIDI != 0 and SAIFI != 0:
					IndividualFaults[record[self.ORSNumCol]] = [Date, SAIDI, SAIFI, ICPcount, Feeder]