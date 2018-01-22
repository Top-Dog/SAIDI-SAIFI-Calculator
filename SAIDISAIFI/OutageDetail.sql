SELECT Out_Feeder_Name, StageFeederAbbrv, OutS_Isol_Pnt, Stage_Isolation_Pnt, Out_Num, Out_Linked_Num, OutS_OffDate, OutS_OffTime, OutS_OnDate, OutS_OnTime, Out_Network, Class_Desc, Out_Calc_ICP AS UniqueICPs, OutS_Calc_Cnt, OutS_Calc_Dur, OutS_Calc_CustMin, Voltage, Out_AutoReclose, Out_Cause_Desc, Out_Note, Out_Reason
FROM (SELECT y.Out_Feeder_Name, z.StageFeederAbbrv, a.OutS_Isol_Pnt, z.Stage_Isolation_Pnt, Out_Num, Out_Linked_Num, a.OutS_OffDate, a.OutS_OffTime, a.OutS_OnDate, a.OutS_OnTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, a.OutS_Calc_Cnt, a.OutS_Calc_Dur, a.OutS_Calc_CustMin, v.Voltage, Out_AutoReclose, t.Class_Desc, o.Out_Cause_Desc, Out_Note, Out_Reason
	FROM dbo.tbl_Outage
		LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Class_Code, VLC_Desc AS Class_Desc FROM tbl_Valid_Lookup_Code WHERE VLC_Family=24)t ON Out_Class = t.Class_Code
		LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Cause_Code, VLC_Desc AS Out_Cause_Desc FROM tbl_Valid_Lookup_Code WHERE VLC_Family=21)o ON Out_Cause = o.Cause_Code
		LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Voltage_Code, VLC_Desc AS Voltage FROM tbl_Valid_Lookup_Code WHERE VLC_Family=22)v ON Out_Voltage = v.Voltage_Code
		LEFT JOIN (SELECT NMIP_Code, NMIP_Name AS Out_Feeder_Name FROM tbl_NM_IsolationPoint)y ON Out_Feeder = y.NMIP_Code
		LEFT JOIN (SELECT * FROM dbo.tbl_Outage_Stage)a ON OutS_Out_Id = Out_Id
		LEFT JOIN (SELECT NMS_NMIP_Code, NMS_Feeder AS StageFeederAbbrv, NMS_NMIP_Name AS Stage_Isolation_Pnt FROM tbl_NM_Structure)z ON OutS_Isol_Pnt = z.NMS_NMIP_Code
	GROUP BY y.Out_Feeder_Name, z.StageFeederAbbrv, a.OutS_Isol_Pnt, Out_Num, z.Stage_Isolation_Pnt, Out_Linked_Num, a.OutS_OffDate, a.OutS_OffTime, a.OutS_OnDate, a.OutS_OnTime, Out_CC_YearEnd, Out_Network, Out_Calc_ICP, a.OutS_Calc_Cnt, a.OutS_Calc_Dur, a.OutS_Calc_CustMin, v.Voltage, Out_AutoReclose, t.Class_Desc, o.Out_Cause_Desc, Out_Note, Out_Reason
	HAVING a.OutS_OffDate >='4/1/2002') d
GROUP BY Out_Feeder_Name, StageFeederAbbrv, OutS_Isol_Pnt, Out_Num, Stage_Isolation_Pnt, Out_Linked_Num, OutS_OffDate, OutS_OffTime, OutS_OnDate, OutS_OnTime, Out_Network, Class_Desc, Out_Calc_ICP, OutS_Calc_Cnt, OutS_Calc_Dur, OutS_Calc_CustMin, Voltage, Out_AutoReclose, Out_Num, Out_Cause_Desc, Out_Note, Out_Reason
ORDER BY OutS_OffDate;