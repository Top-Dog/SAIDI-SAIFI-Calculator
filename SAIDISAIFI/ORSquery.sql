SELECT Out_Feeder_Name, 
			Out_Num, 
			Out_Linked_Num, 
			Out_OffDate, 
			Out_OffTime, 
			Out_Network, 
			Class_Desc, 
			Out_Calc_CustMin, 
			Out_Calc_ICP, 
			Out_AutoReclose 
			
	FROM (
		SELECT y.Out_Feeder_Name, 
		Out_Num, 
		Out_Linked_Num, 
		Out_OffDate, 
		Out_OffTime, 
		Out_CC_YearEnd, 
		Out_Network, 
		Out_Calc_ICP, 
		Out_Calc_Cust,
		Out_Calc_CustMin, 
		Out_AutoReclose, 
		Out_Class, 
		t.Class_Desc FROM dbo.tbl_Outage 
			LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Class_Code, VLC_Desc AS Class_Desc
				FROM tbl_Valid_Lookup_Code
			WHERE VLC_Family=24)t ON Out_Class = t.Class_Code
		
			LEFT JOIN (SELECT NMIP_Code, NMIP_Name AS Out_Feeder_Name FROM tbl_NM_IsolationPoint)y ON Out_Feeder = y.NMIP_Code
									
			GROUP BY y.Out_Feeder_Name, 
			Out_Num, 
			Out_Linked_Num, 
			Out_OffDate, 
			Out_OffTime, 
			Out_CC_YearEnd, 
			Out_Network, 
			Out_Calc_ICP, 
			Out_Calc_Cust, 
			Out_Calc_CustMin, 
			Out_AutoReclose, 
			Out_Class, 
			t.Class_Desc
	HAVING Out_OffDate >='4/1/2002') d

	GROUP BY Out_Feeder_Name, 
	Out_Num, 
	Out_Linked_Num, 
	Out_OffDate, 
	Out_OffTime, 
	Out_Network, 
	Class_Desc, 
	Out_Calc_CustMin, 
	Out_Calc_ICP, 
	Out_AutoReclose, 
	Out_Num
			
	ORDER BY Out_OffDate;