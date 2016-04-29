'''
Created on 30/03/2016
Last Modfied: 18/4/2016

@author: Sean D. O'Connor

Run the SAIDI SAIFI calculator using 
my Excel libriary, spesfic run file
for the SAIDI SAIFI Calulator py file.
'''
import threading, time, multiprocessing, datetime, sys
from SAIDISAIFI import ORSCalculator, Output, Parser
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

		Network = ORSCalculator(sum_like_keys(ICPs), NetworkName, startdate, enddate)
		Network.generate_stats()
		Network.display_stats("outage", "Individual Outages.txt")
		Network.display_stats("month", "Results Table - Monthly.txt")
		Network.display_stats("fiscal year", "Results Table.txt")
		Network.display_stats("day", "Results Table - Daily.txt")
		
		# Debugging stuff - produce the complete fault record in the database, excludes extra outages
		DBG = SAIDISAIFI.CalculatorAux.ORSDebug(Network)
		DBG.create_csv()

		# Distrobution Automation stuff
		Network.DA_Table("DA Table.txt", datetime.datetime(2014,4,1), datetime.datetime(2016,4,1))

		# Put the completed network into an output queue
		NetworkOutQueue.put(Network)


if __name__ == "__main__":    
	starttime = datetime.datetime.now()
	# Get the currently active MS Excel Instace
	try:
		xl = Excel.Launch.Excel(visible=True, runninginstance=True)
		print xl.xlBook.Name
	except:
		print "You need to have the Excel sheet open and set as the active window"
		time.sleep(5)
		sys.exit()

	p = Parser.ParseORS(xl)
	# All ICP counts are averages as of the 31 March i.e. fincial year ending
	ICPNums = p.Read_Num_Cust()
	
	startdate, enddate = p.Read_Dates_To_Publish()

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
	while num_networks < len(Networks):
		Network = NetworkOutQueue.get(True) # Blocks indefinetly if nothing is in the queue
		
		xlPlotter = Output.ORSPlots(Network, xl)
		xlTables = Output.ORSOutput(Network, xl)
		
		# Only do this once, for the very first network being run - setup the excel book
		if num_networks == 0:
			xlPlotter.Clean_Workbook()
			for yrstart in p.StartDates:
				year = str(yrstart.year)
				# Create a new sheet, delete any existing sheets with the same name
				xlPlotter.Create_Sheet(year)
				# Populate Dates
				xlPlotter.Fill_Dates(yrstart, year)
		
		# Update the ComCom comparison table in excel
		xlTables.Populate_Reliability_Stats()
		
		# Create a new progress bar
		pb = ProgressBar(len(p.StartDates), "SAIDI/SAIFI graph(s)")
		pb_thread = threading.Thread(target=update_prog_bar, args=(pb,))
		pb_thread.start()
		
		# Fill the series columns, create the graphs
		for yrstart in p.StartDates:       
			year = str(yrstart.year)
			xlPlotter.Populate_Fixed_Stats(year) # Com Com table values scaled linearly
			xlPlotter.Populate_Daily_Stats(year) # Daily real world SAIDI/SAIDI
			xlTables.Summary_Table(year)
			xlPlotter.Create_Graphs(year)
			pb.update_paced()
		
		# Wait for the progress bar to complete to 100%
		pb_thread.join()
		num_networks += 1

	# Wait for all workers to finish
	for process in processes:
		process.join()
	
	# Let the user know that we are done
	print "Task completed in %d seconds" % (datetime.datetime.now() - starttime).seconds
	raw_input("Done. Press the return key to exit.")