<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->

<%checktimeout()
  'checkguest()
  
Dim sValText, iLastPers, lOrgNo, iCt 
Dim dtRptDate, bCurrPers, sKey
Dim bNotify, iCtr, sValTemp, conn, RS, sHref

Dim sTemp, iTemp, bTemp, sTemp2
Dim iCst1, iCst2, iCst3, iCst4, bnr, iQPID
Dim ACLDefined,MsgID


	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")

	Set conn = GetNewCN()
	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", conn)
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, conn)

	Set RS = Server.CreateObject("ADODB.Recordset")


	'Validation PHASE		
	sValText = ""
	bNotify = False

	'Validate DriverO
	sTemp = Request.Form("optDriverO")

	If Len(Trim(sTemp)) = 0 then
		bNotify = True
		sValText = sValText & "It must be indicated whether the driver was alone.<BR>"
	End if

	'Validate Units
	sTemp = Request.Form("optVeh")
	If Len(Trim(sTemp)) = 0 then
		bNotify = True
		sValText = sValText & "The ownership status of the vehicle involved must be indicated.<BR>"
	End if

	'Validate Material Name
	sTemp = Request.Form("optCoBus")

	If Len(Trim(sTemp)) = 0 then
		bNotify = True
		sValText = sValText & "It must be specified if the vehicle was on company business.<BR>"
	End if

	'Validate Weather Cond - Rd Cond
	iTemp = 0
	sTemp = Request.Form("chkWCDry")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkWCWet")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkWCIce")
	If sTemp = "on" then iTemp = iTemp + 1
		
	If iTemp > 1 Then
		bNotify = True
		sValText = sValText & "The weather conditions selections Dry, Wet/Slick, and Snow/Icy " &_
		 "are mutually exclusive.<BR>"
	End if

	'Validate Weather Cond - Visibility
	iTemp = 0
	sTemp = Request.Form("chkWCClear")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkWCDust")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkWCFog")
	If sTemp = "on" then iTemp = iTemp + 1

	If iTemp > 1 Then
		bNotify = True
		sValText = sValText & "The weather conditions selections Clear, Dust/Sandstorm, and Fog are mutually exclusive.<BR>"
	End if

	'Validate Road Type - Surface
	iTemp = 0
	sTemp = Request.Form("chkRTPaved")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkRTUnpaved")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkRTOffRd")
	If sTemp = "on" then iTemp = iTemp + 1
		
	If iTemp > 1 Then
		bNotify = True
		sValText = sValText & "The road type selections Paved, Unpaved, and Off Road are mutually exclusive.<BR>"
	End if


	'Validate Road Type - Grade
	iTemp = 0
	sTemp = Request.Form("chkRTUp")
	If sTemp = "on" then iTemp = iTemp + 1
	sTemp = Request.Form("chkRTDown")
	If sTemp = "on" then iTemp = iTemp + 1

	If iTemp > 1 Then
		bNotify = True
		sValText = sValText & "The road type selections Up a Grade and Down a Grade " &_
		 "are mutually exclusive.<BR>"
	End if
	


'WRITE PHASE		
	If bNotify = false Then	
		RS.LockType = 3		
		sTemp = "SELECT * FROM tblRIRauto WHERE QPID="&SafeNum(iQPID)
		RS.Open sTemp, conn
		
		If RS.EOF Then 
			RS.AddNew							
			 RS("QPID")		= iQPID
			MsgID=127
		else
			MsgID=106
		End if
		RS("RevDate")= Date()
		RS("DriverName") = left(Request.Form("DriverName"),50)
		iTemp = Request.Form("optConvoy")
		If iTemp = 1 then
			RS("Convoy")= True
		Else
			RS("Convoy")= False
		End if
			
		iTemp = Request.Form("optDriverO")
		If iTemp = 1 then
			RS("DriverO")= True
		Else
			RS("DriverO")= False
		End if
		
		iTemp = Request.Form("optVeh")
		RS("VehicleStatus")= iTemp
						
		iTemp = Request.Form("optCoBus")
		If iTemp = 1 then
			RS("CoBus")= True
		Else
			RS("CoBus")= False
		End if
		
		'Weather Cond
		sTemp = Null
		sTemp2 = Null
		sTemp = Request.Form("chkWCDry")
		If sTemp = "on" then sTemp2 = "D"
		sTemp = Request.Form("chkWCWet")
		If sTemp = "on" then sTemp2 = "W"
		sTemp = Request.Form("chkWCIce")
		If sTemp = "on" then sTemp2 = "I"
		RS("RoadCond")= sTemp2
				
	
		sTemp = Null
		sTemp2 = Null
		sTemp = Request.Form("chkWCClear")
		If sTemp = "on" then sTemp2 = "C"
		sTemp = Request.Form("chkWCDust")
		If sTemp = "on" then sTemp2 = "D"
		sTemp = Request.Form("chkWCFog")
		If sTemp = "on" then sTemp2 = "F"
		RS("Visibility")= sTemp2
		
		sTemp = Request.Form("chkWCHot")
		If sTemp = "on" then
			RS("Heat")= True
		else
			RS("Heat")= False
		End If

		'Road Type
		sTemp = Null
		sTemp2 = Null
		sTemp = Request.Form("chkRTPaved")
		If sTemp = "on" then sTemp2 = "P"
		sTemp = Request.Form("chkRTUnpaved")   
		If sTemp = "on" then sTemp2 = "U"
		sTemp = Request.Form("chkRTOffRd")
		If sTemp = "on" then sTemp2 = "O"
		RS("RoadSurface")= sTemp2
		
		sTemp = Null
		sTemp2 = Null
		sTemp = Request.Form("chkRTUp")
		If sTemp = "on" then sTemp2 = "U"
		sTemp = Request.Form("chkRTDown")
		If sTemp = "on" then sTemp2 = "D"
		RS("RoadGrade")= sTemp2
		
		sTemp = Request.Form("chkRTCurve")
		If sTemp = "on" then
			RS("RoadCurve")= True
		else
			RS("RoadCurve")= False
		End If		
		
		sTemp = Request.Form("chkRTNarrow")
		If sTemp = "on" then
			RS("RoadNarrow")= True
		else
			RS("RoadNarrow")= False
		End If	
		
		sTemp = Request.Form("chkRTPoor")
		If sTemp = "on" then
			RS("PoorSurf")= True
		else
			RS("PoorSurf")= False
		End If	
		
		'Accident Type
		sTemp = Request.Form("chkATHitF")
		If sTemp = "on" then
			RS("ATHitF")= True
		else
			RS("ATHitF")= False
		End If	
		
		sTemp = Request.Form("chkATHitB")
		If sTemp = "on" then
			RS("ATHitB")= True
		else
			RS("ATHitB")= False
		End If
		
		sTemp = Request.Form("chkATBack")
		If sTemp = "on" then
			RS("ATBack")= True
		else
			RS("ATBack")= False
		End If
		
		sTemp = Request.Form("chkATHitSO")
		If sTemp = "on" then
			RS("ATHitSO")= True
		else
			RS("ATHitSO")= False
		End If
		
		sTemp = Request.Form("chkATHitP")
		If sTemp = "on" then
			RS("ATHitPed")= True
		else
			RS("ATHitPed")= False
		End If
		
		sTemp = Request.Form("chkATRoll")
		If sTemp = "on" then
			RS("ATRO")= True
		else
			RS("ATRO")= False
		End If
		
		sTemp = Request.Form("chkATSS")
		If sTemp = "on" then
			RS("ATSS")= True
		else
			RS("ATSS")= False
		End If	
		
		sTemp = Request.Form("chkATPass")
		If sTemp = "on" then
			RS("ATPass")= True
		else
			RS("ATPass")= False
		End If
		
		sTemp = Request.Form("chkATPassed")
		If sTemp = "on" then
			RS("ATPassed")= True
		else
			RS("ATPassed")= False
		End If
		
		sTemp = Request.Form("chkATHitR")
		If sTemp = "on" then
			RS("ATHitRun")= True
		else
			RS("ATHitRun")= False
		End If
		
		sTemp = Request.Form("chkATHitA")
		If sTemp = "on" then
			RS("ATHitA")= True
		else
			RS("ATHitA")= False
		End If
		
		sTemp = Request.Form("chkATRanOR")
		If sTemp = "on" then
			RS("ATRanOR")= True
		else
			RS("ATRanOR")= False
		End If
		
		sTemp = Request.Form("chkHeadOC")
		If sTemp = "on" then
			RS("HeadOC")= True
		else
			RS("HeadOC")= False
		End If
		
				
		'Bottom
		iTemp = Request.Form("optDrug")
		If iTemp = 1 then
			RS("Drugs")= True
		else
			RS("Drugs")= False
		End if
		
		
		iTemp = Request.Form("txtSpeed")
		If IsNumeric(iTemp)then 
			RS("Speed")= iTemp
		Else
			RS("Speed")= Null
		End if
		
		iTemp = Request.Form("optSpeedU")
		If iTemp = 1 then 
			RS("SpeedUnit")= "M"
		Else
			RS("SpeedUnit")= "K"
		End if
		
		iTemp = Request.Form("optMonitor")
		If iTemp = 1 then
			RS("Monitor")= True
		else
			RS("Monitor")= False
		End if
		
		iTemp = Request.Form("optSeatbelt")
		If iTemp = 1 then
			RS("Seatbelts")= True
		else
			RS("Seatbelts")= False
		End if
		
		iTemp = Request.Form("optCert")
		If iTemp = 1 then
			RS("Certificate")= True
		else
			RS("Certificate")= False
		End if
		
		iTemp = Request.Form("optCitation")
		If iTemp = 1 then
			RS("Citation")= True
		else
			RS("Citation")= False
		End if
		
		iTemp = Request.Form("optDD")
		If iTemp = 1 then
			RS("DD")= True
		else
			RS("DD")= False
		End if
		
		iTemp = Request.Form("optCD")
		If iTemp = 1 then
			RS("CD")= True
		else
			RS("CD")= False
		End if
		
		'Update
		RS.Update
		RS.Close
		
		' Costs
		UpdateCost conn, 2, iQPID, lOrgNo, dtRptDate
		UpdateUserInfo iQPID,conn
		Set RS = Nothing
		Set conn = Nothing
			
		LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossAuto2.asp",MsgID,""
		Response.Redirect("LOSSauto.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
	Else
		conn.Close
		Set conn = Nothing
	End if
%>
				
<HTML>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
	<TR>
		<TD valign=top align=center>
			<p class=title id=styleMedium>Processing RIR - Loss Report (2)</p>
		</TD>
	</TR>
						
	<TR>
		<TD><HR>
		</TD>
	</TR>
			
	<TR>
		<TD><span class=urgent id=styleMedium>		
			<%Response.Write sValText%>	<BR><BR></span>
			<I><B>Hit the back button on your browser to correct these problems...</B></I>		
		</TD>
	</TR>
</TABLE>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSauto2.asp;1 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSauto2.asp;1 %>
