import datetime

headers = [i[0] for i in cursor.description]
def data_to_dict(headers, datarow):
	dict = {}
	assert len(headers) == len(data), "The supplied header length does not match the length of the data."
	for header, value in zip(headers, data):
		dict[header] = value
	return dict

class OutageRecord(object):
	"""A container class for outages in the ORS"""
	def __init__(self, dbrecord):
		LinkedID = dbrecord.get("Out_Linked_Num")
		ID = dbrecord.get("Out_Num")
		
		FeederName = dbrecord.get("Out_Feeder_Name")
		FeederAbbreviation = dbrecord.get("StageFeederAbbrv")
		OutageStageIsolationPoint = dbrecord.get("OutS_Isol_Pnt")
		OutageStageIsolationPointName = dbrecord.get("Stage_Isolation_Pnt")

		Network = dbrecord.get("Out_Network")
		Class = dbrecord.get("Class_Desc")
		AutoReclose = dbrecord.get("Out_AutoReclose")
		Cause = dbrecord.get("Out_Cause_Desc")
		Note = dbrecord.get("Out_Note")
		Voltage = dbrecord.get("Voltage")
		
		OffDateTime = dbrecord.get("OutS_OffDate") + datetime.timedelta(
			hours=dbrecord.get("OutS_OffTime").hour, minutes=dbrecord.get("OutS_OffTime").minute)
		OnDateTime = dbrecord.get("OutS_OnDate") + datetime.timedelta(
			hours=dbrecord.get("OutS_OnTime").hour, minutes=dbrecord.get("OutS_OnTime").minute)
		
		NumUniqueICPs = dbrecord.get("UniqueICPs")
		NumStageICPs = dbrecord.get("OutS_Calc_Cnt")
		StageDuration = dbrecord.get("OutS_Calc_Dur")
		StageCustMin = dbrecord.get("OutS_Calc_CustMin")

		OutageCause = dbrecord.get("Out_Cause_Desc")
		OuttageNote = dbrecord.get("Out_Note")

class Calculator(object):
	self.PlannedFaults = {} # key = linked ORS, values = [date, SAIDI, SAIFI, unique ICP count, Feeder]
	self.GroupedPlannedFaults = {} # key = date, values = [SAIDI, SAIFI, NumberOfOutageRecords]

class Day(object):
	# All the linked outages for the day
	# {key=LinkedORSNumber, value=class<LinkedOutage>}
	LinkedOutages = {-258 : class<LinkedOutage>}
	TotalSAIDI = 0
	TotalSAIFI = 0
	NumberOutages = 0

	# Example Tallies (to be calculated)
	OutagesCause = {'Defective Equipment' : (frequencyTally, SAIDI, SAIFI)}
	OutagesNote = {'Broken Crossarm' : (frequencyTally, SAIDI, SAIFI)}
	OutagesIsolation = {'Q10258' : (frequencyTally, SAIDI, SAIFI)}

class LinkedOutage(object):
	# All the link outages for the day
	# {key=ORSNumber, value=class<Outage>}
	Outages = {-259 : class<Outage>}
	AutoReclose = False
	CauseDescription = ""
	OutageClass = "" # Make this a model class?



	LinkedID = dbrecord.get("Out_Linked_Num")
	ID = dbrecord.get("Out_Num")
		
	FeederName = dbrecord.get("Out_Feeder_Name")
	FeederAbbreviation = dbrecord.get("StageFeederAbbrv")
	Network = dbrecord.get("Out_Network")
	Class = dbrecord.get("Class_Desc")
	AutoReclose = dbrecord.get("Out_AutoReclose")
	Cause = dbrecord.get("Out_Cause_Desc")
	Note = dbrecord.get("Out_Note")
		
	OffDateTime = dbrecord.get("OutS_OffDate") + datetime.timedelta(
		hours=dbrecord.get("OutS_OffTime").hour, minutes=dbrecord.get("OutS_OffTime").minute)
	OnDateTime = dbrecord.get("OutS_OnDate") + datetime.timedelta(
		hours=dbrecord.get("OutS_OnTime").hour, minutes=dbrecord.get("OutS_OnTime").minute)
		
	NumUniqueICPs = dbrecord.get("UniqueICPs")
	NumStageICPs = dbrecord.get("OutS_Calc_Cnt")
	StageDuration = dbrecord.get("OutS_Calc_Dur")
	StageCustMin = dbrecord.get("OutS_Calc_CustMin")