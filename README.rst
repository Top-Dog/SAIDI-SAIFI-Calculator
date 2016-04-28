This is a very purpose specific module that is designed to read data from the Outage Recoding System (ORS) and create dictionaries of organised data suitable for analysis. The module is also able to produce .txt files that summarise the accumulated SAIDI and SAIFI for a particular network at predefined time step/date resolutions.

This module calculates SAIDI and SAIFI in accordance with the New Zealand Commerce Commission "Electricity Distribution Services Default Price-Quality Path". It was designed to give suitable information for the annual Asset Management Plans (AMPs) as well as the information disclosures.

It is only suitable for calculating SAIDI and SAIFI between 2015 and 2020.

**Usage Examples**
------------------------------
*Produce a csv output of the faults read from the database*
-----------------------------------------------------
>>> from SAIDISAIFI import ORSCalculator, ORSDebug
>>> import datetime
>>> ICPCount = {2000: 13456, 2001: 13987, 2002: 14012, 2003: 14000}
>>> NetworkName = "OTPO, LLNW"
>>> startdate = datetime.datetime(1999, 4, 1)
>>> enddate = datetime.datetime(2003, 3, 31)
>>> OJV = ORSCalculator(ICPCount, NetworkName, startdate, enddate)
>>> DBG = ORSDebug(OJV)
>>> DBG.create_csv()

*Produce the stats files for a particular network*
-----------------------------------------------------
>>> from SAIDISAIFI import ORSCalculator, Output, Parser
>>> import datetime
>>> ICPCount = {2000: 13456, 2001: 13987, 2002: 14012, 2003: 14000}
>>> NetworkName = "OTPO, LLNW"
>>> startdate = datetime.datetime(1999, 4, 1)
>>> enddate = datetime.datetime(2003, 3, 31)
>>> Network = ORSCalculator(ICPCount, NetworkName, startdate, enddate)
>>> Network.generate_stats()
>>> Network.display_stats("outage", "Individual Outages.txt")
>>> Network.display_stats("month", "Results Table - Monthly.txt")
>>> Network.display_stats("fiscal year", "Results Table.txt")
>>> Network.display_stats("day", "Results Table - Daily.txt")
