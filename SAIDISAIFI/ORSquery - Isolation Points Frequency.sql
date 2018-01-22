SELECT y.Feeder_Name,i.StageFeeder AS StageFeeder,IsolationPoint, Out_Num, Out_Linked_Num, CONVERT(varchar(10),OutS_OffDate,103) AS 'OffDate', --CONVERT(varchar(8),OutS_OffTime,108) AS 'OffTime', 
              CONVERT(varchar(10),Out_CC_YearEnd,103) AS 'YearEnd', Out_Network AS 'Network', Out_Calc_ICP AS 'UniqueICP', 
              Out_Calc_Cust AS 'TotalICP',Out_Calc_CustMin,Out_AutoReclose AS 'AR', Out_Class, t.Class_Desc,OutSI_MRP AS 'ICP',
              COUNT(OutSI_MRP) AS 'NumOccurrences'
FROM dbo.tbl_Outage
              JOIN tbl_Outage_Stage ON OutS_Out_Id = Out_Id
              JOIN tbl_Outage_Stage_ICP ON OutSI_OutS_Id = OutS_Id
        LEFT JOIN (SELECT CONVERT(int, VLC_Code) AS Class_Code, VLC_Desc AS Class_Desc FROM tbl_Valid_Lookup_Code WHERE VLC_Family=24)t ON Out_Class = t.Class_Code
        LEFT JOIN (SELECT NMIP_Code, NMIP_Name AS Feeder_Name FROM tbl_NM_IsolationPoint)y ON Out_Feeder = y.NMIP_Code
       LEFT JOIN (SELECT NMS_NMIP_Code,NMS_Feeder AS StageFeeder,NMS_NMIP_Name AS IsolationPoint FROM tbl_NM_Structure)i ON OutS_Isol_Pnt = i.NMS_NMIP_Code

WHERE OutS_OffDate >='4/1/2002' 
AND Out_Num = 4175 AND IsolationPoint = 'MONUMENT CNR E'
GROUP BY y.Feeder_Name,i.StageFeeder,IsolationPoint, Out_Num, Out_Linked_Num, OutS_OffDate, 
       Out_CC_YearEnd, Out_Network, Out_Calc_ICP, Out_Calc_Cust,       Out_Calc_CustMin,Out_AutoReclose, Out_Class, t.Class_Desc,OutSI_MRP
ORDER BY Out_Num
