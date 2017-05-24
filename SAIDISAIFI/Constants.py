import os

# Working directories
FILE_DIRS = {
	#"OTPO" : "C:\Users\sdo\Documents\SAIDI and SAIFI\Fed 2016 SAIDI & SAIFI\OJV", 
	#"ELIN" : "C:\Users\sdo\Documents\SAIDI and SAIFI\Fed 2016 SAIDI & SAIFI\EIL",
	#"TPCO" : "C:\Users\sdo\Documents\SAIDI and SAIFI\Fed 2016 SAIDI & SAIFI\TPC"
	"OTPO" : os.path.expanduser('~/Documents/SAIDI and SAIFI/Stats/OJV'),
	"ELIN" : os.path.expanduser('~/Documents/SAIDI and SAIFI/Stats/EIL'),
	"TPCO" : os.path.expanduser('~/Documents/SAIDI and SAIFI/Stats/TPC'),
	"GENERAL" : os.path.expanduser('~/Documents/SAIDI and SAIFI'),
	}

# These values are used for verification only. They are offical data from the CC, they should 
# not be edited. TPC was generated using the same method as EIL and TPC, when the cacluated values
# matched the determination values. In CC data limit = cap.
CC_Vals = {
	# Change to abbreviated figures to match CC determination - undone
	"OTPO" : {"SAIDI_LIMIT" : 254.915313720703, "SAIDI_UBV" : 13.2414436340332, "SAIFI_LIMIT" : 2.92730903625488, "SAIFI_UBV" : 0.176474571228027,
				"SAIDI_TARGET" : 224.577346801757, "SAIDI_COLLAR" : 194.239379882812, "SAIDI_CAP" : 254.915313720703,
				"SAIFI_TARGET" : 2.52385830879211, "SAIFI_COLLAR" : 2.12040758132934, "SAIFI_CAP" : 2.92730903625488},
	#"OTPO" : {"SAIDI_LIMIT" : 254.915, "SAIDI_UBV" : 13.241, "SAIFI_LIMIT" : 2.927, "SAIFI_UBV" : 0.176,
	#		"SAIDI_TARGET" : 224.5773, "SAIDI_COLLAR" : 194.2394, "SAIDI_CAP" : 254.9153,
	#		"SAIFI_TARGET" : 2.5239, "SAIFI_COLLAR" : 2.1204, "SAIFI_CAP" : 2.9273},
	
	# Change to abbreviated figures to match CC determination - undone
	"ELIN" : {"SAIDI_LIMIT" : 31.1267299652099, "SAIDI_UBV" : 3.24435400962829, "SAIFI_LIMIT" : 0.771682739257812, "SAIFI_UBV" : 0.0798516571521759,
				"SAIDI_TARGET" : 24.075885772705, "SAIDI_COLLAR" : 17.0250415802001, "SAIDI_CAP" : 31.1267299652099,
				"SAIFI_TARGET" : 0.59350174665451, "SAIFI_COLLAR" : 0.415320783853531, "SAIFI_CAP" : 0.771682739257812},    
	#"ELIN" : {"SAIDI_LIMIT" : 31.127, "SAIDI_UBV" : 3.244, "SAIFI_LIMIT" : 0.772, "SAIFI_UBV" : 0.080,
	#			"SAIDI_TARGET" : 24.0759, "SAIDI_COLLAR" : 17.0250, "SAIDI_CAP" : 31.1267,
	#			"SAIFI_TARGET" : 0.5935, "SAIFI_COLLAR" : 0.4153, "SAIFI_CAP" : 0.7717},
			
	# Figures before the date change affecting 3 ICP or less.
	#"TPCO" : {"SAIDI_LIMIT" : 165.45900, "SAIDI_UBV" : 6.20432, "SAIFI_LIMIT" : 3.15665, "SAIFI_UBV" : 0.11105,
	#            "SAIDI_TARGET" : 149.82174, "SAIDI_COLLAR" : 134.18448, "SAIDI_CAP" : 165.45900,
	#            "SAIFI_TARGET" : 2.84444, "SAIFI_COLLAR" : 2.53222, "SAIFI_CAP" : 3.15665}
	"TPCO" : {"SAIDI_LIMIT" : 165.66626, "SAIDI_UBV" : 6.20432, "SAIFI_LIMIT" : 3.15824, "SAIFI_UBV" : 0.11105,
				"SAIDI_TARGET" : 150.02806, "SAIDI_COLLAR" : 134.38985, "SAIDI_CAP" : 165.66626,
				"SAIFI_TARGET" : 2.84602, "SAIFI_COLLAR" : 2.53380, "SAIFI_CAP" : 3.15824}
}

# The maximum allowable revenue for non-exempt EDB * 1%  = revenue at risk (1/04/2015 - 31/3/2020)
# Added TPCO, which is an exempt EDB
CC_Revenue_At_Risk = {
	"OTPO" : 24780000,
	"ELIN" : 13565000,
	"TPCO" : 43897700
	}

# Average customers on each network as of the end of the fincial year
# For example, the year 2016 infers the average number of customers as of 31/03/2016
CUST_NUMS = {
	"OTPO" : {
		2003: 14596,
		2004: 14596,
		2005: 14596,
		2006: 14463,
		2007: 14463,
		2008: 14721,
		2009: 14754,
		2010: 14767,
		2011: 14785,
		2012: 14813,
		2013: 14818,
		2014: 14781,
		2015: 14806,
		2016: 14836,
		2017: 14912,
		},
	"LLNW" : {
		2003: 0,
		2004: 19,    
		2005: 21.5,     
		2006: 44,     
		2007: 76,     
		2008: 111,     
		2009: 134,     
		2010: 141,     
		2011: 149.5,    
		2012: 159.5,    
		2013: 167.5,    
		2014: 192,    
		2015: 268,    
		2016: 448,
		2017: 877,
		}, 
	"ELIN" : {
		2003: 16759,
		2004: 16759,
		2005: 16759, 
		2006: 16871, 
		2007: 16871, 
		2008: 16974, 
		2009: 17069, 
		2010: 17180, 
		2011: 17232, 
		2012: 17255, 
		2013: 17240, 
		2014: 17257,
		2015: 17333,
		2016: 17326,
		2017: 17377,
		},
	"TPCO" : {
		2003: 31967,
		2004: 31967,
		2005: 31967, 
		2006: 32102, 
		2007: 32373, 
		2008: 32892, 
		2009: 33345, 
		2010: 33871, 
		2011: 34241, 
		2012: 34495, 
		2013: 34581, 
		2014: 34789,
		2015: 35002,
		2016: 35245,
		2017: 35608,
		},
}