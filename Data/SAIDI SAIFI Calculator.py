'''
Created on 30/03/2016
Last Modfied: 18/4/2016

@author: Sean D. O'Connor

Run the SAIDI SAIFI calculator using 
my Excel libriary, spesfic run file
for the SAIDI SAIFI Calulator py file.
'''
import threading, time, multiprocessing, datetime
from SAIDISAIFI import ORSCalculator, Output, Parser
from MSOffice import Excel
from progbar import ProgressBar

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
	print "Thread %d started" % threadID
	while True:
		try:
			NetworkName = NetworkInQueue.get(True, 0.1)
		except:
			print "Thread %d timed out" % threadID
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

		# Put the completed network into an output queue
		NetworkOutQueue.put(Network)


if __name__ == "__main__":    
	starttime = datetime.datetime.now()
	# Get the currently active MS Excel Instace
	xl = Excel.Launch.Excel(visible=True, runninginstance=True)
	print xl.xlBook.Name
	
	# Example: Produce a csv output of the faults read via the ODBC
	#OJV = ORSCalculator(ORSpathOJV, avrgCustomerNumOJV, "OTPO, LLNW")
	#DBG = ORSDebug(OJV)
	#DBG.create_csv()

	p = Parser.ParseORS(xl)
	# All ICP counts are averages as of the 31 March i.e. fincial year ending
	ICPNums = p.Read_Num_Cust()
	
	startdate, enddate = p.Read_Dates_To_Publish()

	Networks = ["OTPO, LLNW", "ELIN", "TPCO"]
	NetworkInQueue = multiprocessing.Queue(maxsize=len(Networks))
	NetworkOutQueue = multiprocessing.Queue(maxsize=len(Networks))
	for n in Networks:
		NetworkInQueue.put(n)


	MAX_NUM_OF_WORKER_THREADS = 3
	threads = []
	for iThread in range(MAX_NUM_OF_WORKER_THREADS):
		thread = multiprocessing.Process(target=worker_networks, args=(startdate, enddate, iThread, NetworkInQueue, NetworkOutQueue, ICPNums))
		# These threads don't need to be daemon, as they'll terminate propertly with a timeout
		threads.append(thread)
		thread.start()
	
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
		
		pb_thread.join()
		num_networks += 1

	for t in threads:
		t.join()
	print "Task completed in %d seconds" % (datetime.datetime.now() - starttime).seconds
	raw_input("Done. Press the return key to exit.")