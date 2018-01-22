SELECT Feeder_Name,ISNULL(StageFeeder,'') AS 'StageFeeder',IsolationPoint, Out_Num, Out_Linked_Num, CONVERT(varchar(10),OutS_OffDate,103) AS OffDate,		
	CONVERT(varchar,OutS_OffTime,108) AS OffTime, Out_Network, Class_Desc, Out_Calc_CustMin AS 'TotalCustMin',OutS_Calc_CustMin AS 'CustMinsForStage', Out_Calc_ICP AS 'UniqueICPs',	
	OutS_Calc_Cnt AS 'StageICPs',Out_AutoReclose 	
FROM (	
	SELECT y.Feeder_Name,i.StageFeeder AS StageFeeder,IsolationPoint, Out_Num, Out_Linked_Num, OutS_OffDate, OutS_OffTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, Out_Calc_Cust,		
			Out_Calc_CustMin,OutS_Calc_CustMin,OutS_Calc_Cnt,Out_AutoReclose, Out_Class, t.Class_Desc 
	FROM dbo.tbl_Outage	
		JOIN tbl_Outage_Stage ON OutS_Out_Id = Out_Id
		LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Class_Code, VLC_Desc AS Class_Desc FROM tbl_Valid_Lookup_Code WHERE VLC_Family=24)t ON Out_Class = t.Class_Code		
		LEFT JOIN (SELECT NMIP_Code, NMIP_Name AS Feeder_Name FROM tbl_NM_IsolationPoint)y ON Out_Feeder = y.NMIP_Code		
		LEFT JOIN (SELECT NMS_NMIP_Code,NMS_Feeder AS StageFeeder,NMS_NMIP_Name AS IsolationPoint FROM tbl_NM_Structure)i ON OutS_Isol_Pnt = i.NMS_NMIP_Code
	GROUP BY y.Feeder_Name,i.StageFeeder,IsolationPoint, Out_Num, Out_Linked_Num, OutS_OffDate, OutS_OffTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, Out_Calc_Cust, Out_Calc_CustMin,OutS_Calc_CustMin,		
			OutS_Calc_Cnt,Out_AutoReclose, Out_Class, t.Class_Desc
HAVING OutS_OffDate >='4/1/2002') d		
GROUP BY Feeder_Name,StageFeeder,IsolationPoint,Out_Num,Out_Linked_Num,OutS_OffDate, OutS_OffTime, Out_Network, Class_Desc, Out_Calc_CustMin,OutS_Calc_CustMin, Out_Calc_ICP,OutS_Calc_Cnt,		
	 Out_AutoReclose, Out_Num	
ORDER BY Out_Num,OutS_OffDate,OffTime