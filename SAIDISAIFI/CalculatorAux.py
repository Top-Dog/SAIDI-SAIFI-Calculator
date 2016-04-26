'''
Created on 30/3/2016

@author: Sean D. O'Connor

This is an extension of the Calculator15 module. 
The classes in this module allow for debugging,
comparission with the ComCom records, and direct
access to the ORS database using ODBC.
'''

import pyodbc, os, csv

class ODBC_ORS_ACCESS(object):
    """Directly connect to the PNL Outage Recording System (ORS)
    and pull live data, rather than manually doing the export and 
    formatting the data in a excel/csv file.
    Requires the pyodbc module."""
    def __init__(self):
        # Connection Parameters (Const)
        self.connStr = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=H:\MSA_Prog\ORS_Prog_a2k3.mde;"
            )
        # The SQL query to perform on the DB.
        # We are using my existing query "Query_Sean_Outage_Class",
        # which can be treated as a table.
        # Remember to use single quotes around strings, else an error is generated.
        # TODO: move all "WHERE" criteria into dictionary generator method.                    
        self.sql = """SELECT 
                    Query_Sean_Outage_Class.Out_AutoReclose, 
                    Query_Sean_Outage_Class.Out_Num, 
                    Query_Sean_Outage_Class.Out_Linked_Num, 
                    Query_Sean_Outage_Class.Out_OffDate, 
                    Query_Sean_Outage_Class.Out_OffTime, 
                    Query_Sean_Outage_Class.Out_Network, 
                    Query_Sean_Outage_Class.Class_Desc, 
                    Query_Sean_Outage_Class.Out_Calc_CustMin, 
                    Query_Sean_Outage_Class.Out_Calc_ICP
                    FROM Query_Sean_Outage_Class
                    WHERE (((Query_Sean_Outage_Class.Out_AutoReclose)='N') AND 
                    ((Query_Sean_Outage_Class.Class_Desc)='Planned - PowerNet' Or 
                    (Query_Sean_Outage_Class.Class_Desc)='Unplanned - PowerNet'));"""
        
        self._connect()
        self._run_query()
    
    def _connect(self):
        """Handle a connetion to the ODBC"""
        self.connection = pyodbc.connect(self.connStr)
        
    def _run_query(self):
        """Run the SQl quer code"""
        self.queryresults = self.connection.execute(self.sql)
        
    def get_query_results(self):
        """Return the query results,
        formatted ready for wrting to excel"""
        QueryData = []
        for row in self.queryresults:
            # The row is already a tuple (itterable object), so there is no need to enclose it in a list 
            QueryData.append(row)
        return QueryData
        
    def close(self):
        """Close the file pointer and DB connection"""
        self.queryresults.close()
        self.connection.close()

# This uses a SQL query as opposed to a MS access one
class ODBC_ORS(object):
    """Directly connect to the PNL Outage Recording System (ORS)
    and pull live data, rather than manually doing the export and 
    formatting the data in a excel/csv file.
    Requires the pyodbc module."""
    def __init__(self):
        # Connection Parameters (Const)
        self.connStr = (
            "DRIVER={SQL Server};SERVER=PNLICP1"
            )
        # The SQL query to perform on the DB.
        # Remember to use single quotes around strings, else an error is generated.
        # TODO: move all "WHERE" criteria into dictionary generator method.                    
        self.sql = """
            SELECT Out_Num, Out_Num, Out_Linked_Num, Out_OffDate, Out_OffTime, Out_Network, Class_Desc, Out_Calc_CustMin, Out_Calc_ICP, Out_AutoReclose
            FROM (SELECT Out_Num, Out_Linked_Num, Out_OffDate, Out_OffTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, Out_Calc_Cust, Out_Calc_CustMin, Out_AutoReclose, Out_Class, t.Class_Desc
                        FROM dbo.tbl_Outage 
                              LEFT JOIN (SELECT CONVERT(int,VLC_Code) AS Class_Code, VLC_Desc AS Class_Desc
                                                FROM tbl_Valid_Lookup_Code
                                                WHERE VLC_Family=24)t ON Out_Class = t.Class_Code
                        GROUP BY Out_Num, Out_Linked_Num, Out_OffDate, Out_OffTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, Out_Calc_Cust, Out_Calc_CustMin, Out_AutoReclose, Out_Class, t.Class_Desc
                        HAVING Out_OffDate >='4/1/2002') d
            GROUP BY Out_Num, Out_Linked_Num, Out_OffDate, Out_OffTime, Out_Network, Class_Desc, Out_Calc_CustMin, Out_Calc_ICP, Out_AutoReclose, Out_Num
            ORDER BY Out_OffDate;
            """
        
        self._connect()
        self._run_query()
    
    def _connect(self):
        """Handle a connetion to the ODBC"""
        self.connection = pyodbc.connect(self.connStr)
        self.cursor = self.connection.cursor()
        
    def _run_query(self):
        """Run the SQl quer code"""
        self.queryresults = self.connection.execute(self.sql)
        #self.queryresults = self.cursor(self.sql)
        
    def get_query_results(self):
        """Return the query results,
        formatted ready for wrting to excel"""
        QueryData = []
        for row in self.queryresults:
            # The row is already a tuple (itterable object), so there is no need to enclose it in a list 
            QueryData.append(row)
        return QueryData
        
    def close(self):
        """Close the file pointer and DB connection"""
        self.queryresults.close()
        self.connection.close()

def ORSComCom(object):
    def __init__(self):
        pass
    
    def ors_output(self):
        '''Write a formatted record to an output CSV file suitable for comparison with COMCOM'''
        try:
            currentDate = self.startDate
            with open(self.outFolder + self.ORSout, 'ab') as genfile:
                writer = csv.writer(genfile)
                #while currentDate != self.endDate + self.deltaDay:
                while currentDate < self.endDate + self.deltaDay:
                    strDate = self._date_to_str(currentDate)
                    SAIDI, SAIFI = self._get_indicies(currentDate, "all")
                    writer.writerow([strDate, SAIDI, SAIFI])
                    currentDate += self.deltaDay
        except:
            print "The output file is open somewhere, probably in MS Excel. Close it."
            print "No output CSV file %s was created." % self.ORSout
            
    def compare(self):
        '''Compare the ComCom data with the data from the ORS that has been converted to ComCom format'''
        rowIndex = 0
        count = 1
        with open(self.outFolder + self.ORSout, 'rb') as ORSfile:
            ORSreader = csv.reader(ORSfile)
            for row in ORSreader:
                CCSAIDI, CCSAIFI = self._comcom_record_recall(rowIndex) # Get the ComCom SAIDI figure
                diffSAIDI = abs(CCSAIDI - float(row[1]))
                if diffSAIDI > self.SAIDItolerance:
                    print "%d. SAIDI tolerance triggered: %.5f. Date: %s. Row Index: %d" % (count, CCSAIDI - float(row[1]), row[0], rowIndex+1)
                diffSAIFI = abs(CCSAIFI - float(row[2]))
                if diffSAIFI > self.SAIFItolerance:
                    print "%d. SAIFI tolerance triggered: %.5f. Date: %s. Row Index: %d" % (count, CCSAIFI - float(row[2]), row[0], rowIndex+1)
                rowIndex += 1
                if diffSAIDI > self.SAIDItolerance or diffSAIFI > self.SAIFItolerance:
                    count += 1
    
    def _comcom_record_recall(self, rowIndex):
        '''Read a record from the ComCom csv file - which was copy&paste from the offical data'''
        localRowIndex = -1 # Offset for the header row
        with open(os.path.join(self.outFolder, self.CCin), 'rb') as CCfile:
            CCreader = csv.reader(CCfile)
            for row in CCreader:
                if localRowIndex == rowIndex:
                    try:
                        return float(row[3]), float(row[4]) # return the ComCom SAIDI figure
                    except:
                        return float(0), float(0) # The filed is blank, so there was 0 SAIDI that day
                localRowIndex += 1
    

class ORSDebug(object):
    """A debug class for the ORS Calculator. This allows for verbose 
    debugging options, writing to files etc."""
    def __init__(self, orsCalc):
        self.orsCalc = orsCalc
    
    def Change_Network(self, newnetwork):
        self.orsCalc = newnetwork
        
    def create_csv(self):
        """Creates a CSV file for the requested neteork"""
        # Number of ICPs as opposed to number of customers
        # Outage duration should be used instead of "Customer min"
        print "Debug: ", self.orsCalc.networknames
        Headings = ["ORS Number", "ORS Number", "Linked_ORS #", 
                    "Date Off", "Time Off", "Network", 
                    "Interruption Class", "Customer min", 
                    "Number of Unique ICP's", "Auto Reclose/Under 1 minute?"]
        ors = ODBC_ORS()
        qryrows = ors.get_query_results()
        with open(os.path.join(self.orsCalc.outFolder, "ORS Dump.csv"), 'wb') as ffaults:
            f = csv.writer(ffaults)
            f.writerow(Headings)
            for row in qryrows:
                if row[self.orsCalc.NetworkCol] in self.orsCalc.networknames:
                    f.writerow(row)
    
    def debug_ranking_values(self, dic):
        dic = {}
        # We only calculate boundary values every 5 years (1/04/2004 - 31/03/2014)
        for date in self.GroupedUnplannedFaults:
            if self._get_fiscal_year(self._date_to_str(date)) in self.BoundryCalcPeriod:
                dic[date] = self.GroupedUnplannedFaults.get(date)
                
        print 'SAIDI/SAIFI - ranked in decreasing order'
        print "SAIDI", "SAIFI"
        for i in range(1, 26):
            Date1 =  sorted(dic.iteritems(), key=lambda e: e[1][0], reverse=True)[i-1][0]
            Date2 =  sorted(dic.iteritems(), key=lambda e: e[1][1], reverse=True)[i-1][0]
            SAIDI = self.nth_largest(i, dic, 0) # Boundary SAIDI
            SAIFI = self.nth_largest(i, dic, 1) # Boundary SAIFI
            print str(i)+".", Date1, SAIDI, Date2, SAIFI
            
    def group_like_events(self):
        """A new UNTESTED function that attempts to group records from the the ORS
        if the MS access hasn't already summed customer minutes and outages"""
        faults = {}
        with open(self.outFolder + self.ORSin, 'rb') as orscsvfile:
            ORSreader = csv.reader(orscsvfile)
            rowNum = 0
            for row in ORSreader:
                key = row[self.ORSNumCol]
                if rowNum >= 1: # Skip the header
                    if key not in faults:
                        # Create a new fault record
                        faults[key] = row
                    else:
                        # Add the new row to an existing record
                        CustomerMins = row[self.CusmMinCol]
                        NumUniqueCustomers = row[self.UniqueICPCol]
                        row[self.CusmMinCol] += CustomerMins
                        row[self.UniqueICPCol] += NumUniqueCustomers
                        faults[key] = row
                else:
                    header = row                
                rowNum += 1
        try:
            os.remove(self.outFolder + "\grouped faults.csv")
        except:
            pass
        
        with open(self.outFolder + "\grouped faults.csv", 'ab') as genfile:
                writer = csv.writer(genfile)
                writer.writerow(header)
                for key in faults:
                    writer.writerow(faults.get(key))