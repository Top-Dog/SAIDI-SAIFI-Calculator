'''
Created on 30/03/2016
Last Modfied: 18/4/2016

@author: Sean D. O'Connor

Run the SAIDI SAIFI calculator using 
my Excel libriary, spesfic run file
for the SAIDI SAIFI Calulator py file.
'''
import threading, time, multiprocessing, datetime, sys
from SAIDISAIFI import ORSCalculator, Output, Parser, Constants
from MSOffice import Excel
from progbar import ProgressBar
import SAIDISAIFI

def sum_like_keys2(*args):
	"""Sums all the values that have the same
	keys in a set of dictionarys. For numeric types only."""
	result = {}
	for dic in args:
		for key in dic:
			result[key] = result.get(key, 0) + dic.get(key, 0)
	return result

def sum_like_keys(dicts):
	"""Sums all the values that have the same
	keys in a set of dictionarys. For numeric types only."""
	result = {}
	for dic in dicts:
		for key in dic:
			result[key] = result.get(key, 0) + dic.get(key, 0)
	return result

def update_prog_bar(progbar):
	while progbar.current_iter < progbar.max_value:
		progbar.update_thread()
		time.sleep(0.5)
	progbar.update_thread()
	print

def worker_networks(startdate, enddate, threadID, NetworkInQueue, NetworkOutQueue, ICPNums):
	print "Process %d started" % threadID
	while True:
		try:
			NetworkName = NetworkInQueue.get(True, 0.1)
		except:
			print "Process %d timed out" % threadID
			return
		ICPs = []
		names = [x.strip(' ') for x in NetworkName.split(',')]
		for name in names:
			ICPs.append(ICPNums.get(name))

		# Deal with OJV's boundary values differently - if no UBV is provided, then the calculated one is used
		#if "OTPO" in NetworkName:
		#	Network = ORSCalculator(sum_like_keys(ICPs), NetworkName, startdate, enddate, boundarySAIDIValue=13.2414436340332, boundarySAIFIValue=0.176474571228027)
		#else:
		#	Network = ORSCalculator(sum_like_keys(ICPs), NetworkName, startdate, enddate)

		# Bevan said to use the abbreviated figures... so here we go
		# If more than one network name is in NetworkName use the first four letters of the name e.g. "OTPO, LLNW"
		Network = ORSCalculator(sum_like_keys(ICPs), NetworkName, startdate, enddate,
			boundarySAIDIValue=Constants.CC_Vals.get(NetworkName[:4]).get("SAIDI_UBV"), 
			boundarySAIFIValue=Constants.CC_Vals.get(NetworkName[:4]).get("SAIFI_UBV"))

		Network.generate_stats()
		Network.display_stats("outage", "Individual Outages.txt")
		Network.display_stats("month", "Results Table - Monthly.txt")
		Network.display_stats("fiscal year", "Results Table.txt")
		Network.display_stats("day", "Results Table - Daily.txt")
		
		# Debugging stuff - produce the complete fault record in the database, excludes extra outages
		DBG = SAIDISAIFI.CalculatorAux.ORSDebug(Network)
		DBG.create_csv()

		# Distrobution Automation calculation over the display period (same interval as the output tables)
		_Start_Time = datetime.datetime(2002, 4, 1)
		_End_Time = datetime.datetime.now() #datetime.datetime(2016, 3, 31)
		Network.DA_Table("DA Table.txt", datetime.datetime(2015, 4, 1), datetime.datetime(2016, 3, 31))
		Network.Outages_Feeder_Year("DA Profiles.txt", datetime.datetime(2002, 4, 1), datetime.datetime(2016, 9, 16))
		Network.Capped_Outages_Table("UBV Outages.txt", _Start_Time, _End_Time)

		# Put the completed network into an output queue
		NetworkOutQueue.put(Network)


if __name__ == "__main__":
	# Start clock for timing the app's execution time
	starttime = datetime.datetime.now()

	# Get the currently active MS Excel Instace
	try:
		xl = Excel.Launch.Excel(visible=True, runninginstance=True)
		print xl.xlBook.Name
	except:
		print "You need to have the Excel sheet open and set as the active window"
		time.sleep(5)
		sys.exit()
	xl.xlApp.ScreenUpdating = False 

	# Handles reading all the data from the UI (Excel)
	p = Parser.ParseORS(xl)
	ICPNums = p.Read_Num_Cust() # The average number of unique ICPs to be used in the calcs
	startdate, enddate = p.Read_Dates_To_Publish() # Determine the minimum date range to run the calculator
	Last_Pub_Date = min(p.Read_Last_Date(), datetime.datetime.now()) # This will be used in "Rob's" table used for commercial
	selected_date_sheet = "User Defined" # The name to be appened to "Calculation" for custom date range sheet (uses Last_Pub_Date)

	# Setup the output handlers
	xlDocument = Output.ORSSheets(xl)

	# Unique names for each (group of) network(s)
	Networks = ["OTPO, LLNW", "ELIN", "TPCO"]
	NetworkInQueue = multiprocessing.Queue(maxsize=len(Networks))
	NetworkOutQueue = multiprocessing.Queue(maxsize=len(Networks))
	for n in Networks:
		NetworkInQueue.put(n)

	# Start the worker processes; produce dictionaries of SAIDI and SAIFI
	MAX_NUM_OF_WORKER_PROCESSES = 3
	processes = []
	for process_i in range(MAX_NUM_OF_WORKER_PROCESSES):
		process = multiprocessing.Process(target=worker_networks, args=(startdate, enddate, process_i, NetworkInQueue, NetworkOutQueue, ICPNums))
		processes.append(process)
		process.start()
	
	# Work with Excel COM to produce graphs for one network at a time (avoid COM threading/asyncronous behaviour)
	num_networks = 0
	ReportValues = {}
	while num_networks < len(Networks):
		# Get a network that has completed its calculations
		Network = NetworkOutQueue.get(True) # Blocks indefinetly if nothing is in the queue
		
		# Create a new instance of the plot and table generator (for Excel output)
		xlPlotter = Output.ORSPlots(Network, xl)
		xlTables = Output.ORSOutput(Network, xl)

		# Populate a dictionary that contains the keys for populating pre-built Excel template sheets
		ReportValues = xlDocument.Merge_Dictionaries(ReportValues, xlTables.Generate_Values(Last_Pub_Date))
		
		# Only do this once, for the very first network being run - setup the excel book
		if num_networks == 0:
			xlPlotter.Clean_Workbook()
			for yrstart in p.StartDates:
				year = str(yrstart.year)
				# Create a new sheet, delete any existing sheets with the same name
				xlPlotter.Create_Sheet(year)
				# Populate Dates
				xlPlotter.Fill_Dates(yrstart, year)
				# Create the summary tables in Excel
				xlTables.Create_Summary_Table()
			# Extra Graph
			xlPlotter.Create_Sheet(selected_date_sheet)
			xlPlotter.Fill_Dates(datetime.datetime(xlPlotter._get_fiscal_year(Last_Pub_Date), 4, 1), selected_date_sheet)
		
		# Update the ComCom comparison table in excel
		xlTables.Populate_Reliability_Stats()
		
		# Create a new progress bar
		pb = ProgressBar(len(p.StartDates)+1, "SAIDI/SAIFI graph(s)")
		pb_thread = threading.Thread(target=update_prog_bar, args=(pb,))
		pb_thread.start()
		
		# Fill the series columns, create the graphs
		for yrstart in p.StartDates:       
			year = str(yrstart.year)
			xlPlotter.Populate_Fixed_Stats(year) # Com Com table values scaled linearly
			xlPlotter.Populate_Daily_Stats(datetime.datetime(yrstart.year+1, 3, 31), year) # Daily real world SAIDI/SAIDI
			#xlTables.Summary_Table(year)
			xlPlotter.Create_Graphs(datetime.datetime(yrstart.year+1, 3, 31), year)
			pb.update_paced()
		# Extra Graph
		xlPlotter.Populate_Fixed_Stats(selected_date_sheet)
		xlPlotter.Populate_Daily_Stats(Last_Pub_Date, selected_date_sheet)
		xlPlotter.Create_Graphs(Last_Pub_Date, selected_date_sheet)
		pb.update_paced()
		
		# Wait for the progress bar to complete to 100%
		pb_thread.join()
		num_networks += 1

	# Wait for all workers to finish
	for process in processes:
		process.join()

	# Populate Excel template sheets - Any future dates will be set to todays date
	xlDocument.YTD_Sheet(ReportValues)
	#xlDocument.YTD_Book(SAIDISAIFI.Constants.FILE_DIRS.get("GENERAL")+r"\Test Template Document.xlsx", ReportValues) # Creates a new file with the templates filled in, just for testing
	
	xl.xlApp.ScreenUpdating = True 

	# Let the user know that we are done - show the execution time
	print "Task completed in %d seconds" % (datetime.datetime.now() - starttime).seconds
	time.sleep(8)