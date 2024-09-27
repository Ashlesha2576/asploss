<%@ Language=VBScript %>
<%option explicit
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'     7-May-2014            Varun Sharma                 Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
'   05-Nov-2014                Varun Sharma                ENH044752  HSE locking of lagging indicators - safety net - key to unlock
%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  
Dim sValText, iLastPers, lOrgNo, iCt 
Dim dtRptDate, bCurrPers, sKey, bDisplayInjury, bFatality
Dim bNotify, iCtr, sValTemp, conn, RS, dtTemp, sHref, iTotalDays,txtComments,rsoutcomevalue,frmdbrsoutcomevalue,dbrsoutcomevalue
Dim sTemp, iTemp, bTemp, sTemp2, sTempFatal,rshiddentxtRDays,rshiddentxtDays,textoutcomevalue, OrgLDays, OrgRDays
Dim iCst1, iCst2, iCst3, iCst4, bnr, iQPID
DIM ACLDefined,MsgID,DelMsg
Dim PType,Fat,SLBFlag,lockval
lOrgNo = Request.QueryString("OrgNo")
dtRptDate = Request.QueryString("rptDate")
iQPID = Request.QueryString("QPID")
iCt = Request.Form("txtCt")

Set conn = GetNewCN()

sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", conn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, conn)

bDisplayInjury = False
bFatality = False
If trim(lcase(Request.Form("bDisplayInjury"))) = "true" then bDisplayInjury = True
'Check for delete
If Request.QueryString("Delete")=1 then 
	Set RS = conn.Execute("Select * from tblRIRPers Where QPID="& SafeNum(iQPID) & " and Seq=" & SafeNum(trim(Request.QueryString("ID"))))
	If Not RS.EOF Then
		LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LOSSpers2.asp",137,""	
		conn.Execute "DELETE FROM tblRIRPers WHERE QPID="& SafeNum(iQPID) & " AND Seq=" & SafeNum(trim(Request.QueryString("ID")))
		DelMsg=server.URLEncode("Entry Successfully Deleted")
		conn.execute("sp_UpdateDateTime " & SafeNum(iQPID) & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
	else
		DelMsg=server.URLEncode("Entry is Not Found")
	End If
	Set RS = Nothing
	conn.close
    Set conn = Nothing
	Response.Redirect("LOSSpers.asp" & sKey & "&msg=" &Delmsg )
end if
	
'Validation PHASE		
sValText = ""
iLastPers = 0
bNotify = False
	
For iCtr = 1 to iCt
		
	'Line validation variables - reset each loop
	bTemp = False
	bCurrPers = False
	sValTemp = ""
					
	'Reporter filled in?
	sTemp = Request.Form("txtName" & iCtr)
	If sTemp <> "Medically Confidential"	Then				
		If Len(Trim(sTemp)) = 0  then
			bTemp = True
			sValTemp = sValTemp & "No name entered for person " & iCtr & ".<BR>"
		Else
			bCurrPers = True
		End if
	End If
					
	'Validate GIN, Co, 3rd Pty
	sTemp = Request.Form("txtType" & iCtr)
					
	If Len(Trim(sTemp)) = 0 then
		bTemp = True
		sValTemp = sValTemp & "No Employee, Company or 3rd Pty Status entered for person " & iCtr & ".<BR>"
	Else
		bCurrPers = True
	End if

		
	If bDisplayInjury then 	
			'Validate Injuries
			sTemp = UCase(Request.Form("txtInj" & iCtr))
			sTempFatal = Request.Form("txtFatal" & iCtr)
							
			If Len(Trim(sTemp)) = 0 AND UCase(sTempFatal) <> "Y" then
				bTemp = True
				sValTemp = sValTemp & "No injury information entered for person " & iCtr & ".<BR>"
			Else
				bCurrPers = True
			End if
				
			'Validate Injury Location
			sTemp = UCase(Request.Form("txtInjLoc" & iCtr))
							
			If Len(Trim(sTemp)) = 0 AND UCase(sTempFatal) <> "Y" then
				bTemp = True
				sValTemp = sValTemp & "No injury location (body parts affected) entered for person " & iCtr & ".<BR>"			
			Else
				bCurrPers = True
			End if
				

			'Validate Outcome
			sTemp = Request.Form("Outcome" & iCtr)
			lockval = Request.Form("txtlock")
			
            if lockval ="" then	
			If Len(Trim(sTemp)) = 0 AND UCase(sTempFatal) <> "Y" then
				bTemp = True
				sValTemp = sValTemp & "No Outcome entered for person " & iCtr & ".<BR>"
			Else
				bCurrPers = True
			End if
            End if
			'Validate Days Lost / Permanent Disability
			sTemp = Request.Form("txtDays" & iCtr)
			If Len(Trim(sTemp)) = 0 AND UCase(Request.Form("Outcome" & iCtr)) = "P" then
				bTemp = True
				sValTemp = sValTemp & "No days lost entered for person " & iCtr & ".<BR>"
			End if
	End If
		
	'Person entered this line?
	If bCurrPers = True Then
		If iLastPers <> iCtr - 1 Then
			'Person entered this line, but a previous line incomplete
			bNotify = True
			sValText = sValText & "Incomplete entry for person on line " & iLastPers + 1 & ".<BR>"
		End if
			
		iLastPers = iCtr
			
		'Check line validation
		If bTemp = True Then
			bNotify = True
			sValText = sValText & sValTemp & "<BR>"
		End If
	End if		
Next
	

'WRITE PHASE		
If bNotify = false Then	
	conn.execute "DELETE FROM tblRIRPers WHERE QPID=" & SafeNum(iQPID) 
	MsgID=113	
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.LockType = 3
						
	For iCtr = 1 To iLastPers
		sTemp = "SELECT * FROM tblRIRpers WHERE QPID=" & SafeNum(iQPID) & " AND Seq=" & SafeNum(iCtr) 
		RS.Open sTemp, conn
			
		If RS.EOF then bNR = True
			
		If bNR = True Then 
			RS.AddNew
			rs("QPID") = iQPID
			RS("Seq") = iCtr
		End if
		rsoutcomevalue=trim(Request.Form("hiddoutcomestylesmall" & iCtr))
	
			rshiddentxtRDays=Request.Form("hiddentxtRDays" & iCtr)
			
			rshiddentxtDays=Request.Form("hiddentxtDays" & iCtr)
			dim rsoutcomers,sqlrs
	set rsoutcomers = Server.CreateObject("ADODB.Recordset")
	sqlrs = "select * from tlkpInjuryOutcome where InjuryOutcomeID = '"&rsoutcomevalue&"'"
	rsoutcomers.open sqlrs, conn
	dbrsoutcomevalue = ""
	if not rsoutcomers.eof then
	dbrsoutcomevalue = rsoutcomers("InjuryOutcomeDesc")
	end if
	rsoutcomers.close
	set rsoutcomers = nothing
	
	
						
		RS("RevDate")= Date()
			
		sTemp = Request.Form("txtName" & iCtr)
		RS("Name")= left(trim(sTemp),60)
				
		RS("Age")= null 
			
		sTemp = Request.Form("txtSrDate" & iCtr)
		If isDate(sTemp) Then RS("Seniority")= CDate(sTemp)
			
		sTemp = Request.Form("txtReturnWorkDate" & iCtr)
		If isDate(sTemp) Then RS("ReturnWorkDate")= CDate(sTemp)

        PType = trim(Request.Form("txtPosExp" & iCtr))
		If Len(PType) > 0 Then RS("PosExp")= PType

		PType = trim(Request.Form("txtType" & iCtr))
		If Len(PType) > 0 Then RS("InjuredPartyType")= PType
	
        sTemp = Request.Form("txtJobfcn" & iCtr)
		If Len(sTemp) > 0 Then RS("JobFunction")= sTemp

        sTemp = Request.Form("txtJob" & iCtr)
		If Len(sTemp) > 0 Then RS("JobFID")= sTemp 
    
			
		sTemp = Request.Form("txtSince" & iCtr)
		If isNumeric(sTemp) Then RS("HrsAwake")= sTemp
			
		sTemp = Request.Form("txtSleep" & iCtr)
		If isNumeric(sTemp) Then RS("HrsSlept")= sTemp
			
		sTemp = Request.Form("txtOnDuty" & iCtr)
		If isNumeric(sTemp) Then RS("HrsOnDuty")= sTemp
			
		sTemp = Request.Form("txtInj" & iCtr)
		If Len(sTemp) > 0 Then RS("InjuryType")= UCase(sTemp)
			
		sTemp = Request.Form("txtInjLoc" & iCtr)
		If Len(sTemp) > 0 Then RS("InjuryPart")= UCase(sTemp)
			
		sTemp = Request.Form("Outcome" & iCtr)
		FAT=False
		if sTemp = "" Then
			RS("Outcome")= NULL
			RS("Fatality") = False
		Else
			RS("Outcome")= sTemp
			if sTemp = "X" Then
				RS("Fatality") = True
				FAT=True
				
				conn.Execute "Update tblRIRRisk Set FailSafe = 0, FailLucky = 0 Where QPID = " & SafeNum(iQPID)
			Else
				RS("Fatality") = False
			End If		
		End if
		
		sTemp = Request.Form("MedVac" & iCtr)
		If sTemp <> "" Then 
			RS("MedVac") = sTemp
		End If
		
		
		'New SLBInv/Con Flag
		SLBFlag = Request.Form("SlbCon" & iCtr)
		If SLBFlag <> "1" Then SLBFlag=0
		If Not (Trim(PType)="3" and FAT) Then SLBFlag = 0
		RS("SLBConFlag") = SLBFlag
				
		iTotalDays = 0
		
		OrgLDays = RS("DaysLost")
		sTemp = Request.Form("txtDays" & iCtr)
		If sTemp= "" Then 
			RS("DaysLost")= 0
		Elseif isNumeric(sTemp) Then 
			RS("DaysLost")= sTemp
			iTotalDays = iTotalDays + sTemp
		Else
			RS("DaysLost")= 0
		End If
		
		OrgRDays = RS("ReducedWorkDays")
		sTemp = Request.Form("txtRDays" & iCtr)
		If sTemp = "" Then
			RS("ReducedWorkDays")= 0
		elseif isNumeric(sTemp) Then 
			RS("ReducedWorkDays")= sTemp
			iTotalDays = iTotalDays + sTemp
		Else
			RS("ReducedWorkDays")= 0
		End If
		
		if iTotalDays >=100 Then
			CheckMinSeverity conn,3
		elseif iTotalDays >=1 Then
			CheckMinSeverity conn,2
		end if

		If NOT bDisplayInjury Then 
			RS("Fatality")= False
			RS("InjuryType")=" " 
			RS("InjuryPart")= " " 
			RS("DaysLost")= 0
			RS("ReducedWorkDays")= 0
			RS("Outcome")= NULL
		End If
			
	
	dim rsoutcome,sql
	set rsoutcome = Server.CreateObject("ADODB.Recordset")
	sql= "select * from tlkpInjuryOutcome where InjuryOutcomeID = '"&trim(Request.Form("Outcome" & iCtr))&"'"
	rsoutcome.open sql, conn
	frmdbrsoutcomevalue = ""
	if not rsoutcome.eof then
	
	frmdbrsoutcomevalue = rsoutcome("InjuryOutcomeDesc")
	end if
	rsoutcome.close
	set rsoutcome = nothing
		
		If RS("Outcome") = "X" then		
			bFatality = True			
			conn.Execute "Update tblRIRRisk Set FailSafe = 0, FailLucky = 0 Where QPID = " & SafeNum(iQPID)
		End If
		
		if Request.Form("Outcome" & iCtr) <> "" then
			if (Request.form("txtRDays"& iCtr) <> rshiddentxtRDays) or (Request.form("txtDays"& iCtr) <> rshiddentxtDays) or (dbrsoutcomevalue <> frmdbrsoutcomevalue) then
				If txtComments <>"" then txtComments=txtComments & " <BR> " end if
				txtComments=txtComments&" RIR Pers Loss modified. Person "&iCtr&". Outcome from "& dbrsoutcomevalue & " - LWD ("&rshiddentxtDays&"), RWD("&rshiddentxtRDays&") to Outcome "&frmdbrsoutcomevalue&" - LWD("&Request.form("txtDays"& iCtr)&"), RWD("&Request.form("txtRDays"& iCtr)&")"
			end if
			
		end if 
		RS.Update
		RS.Close
	Next
	'Pers Costs
	If bDisplayInjury Then			
		UpdateCost conn, 6, iQPID, lOrgNo, dtRptDate
	End If
	Set RS = Nothing
		
	If bFatality then CheckClass(conn)
	UpdateUserInfo iQPID,conn
	conn.Close
	Set conn = Nothing
	LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossPers2.asp",MsgID,txtComments								
	Response.Redirect("LOSSpers.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
Else
		conn.Close
		Set conn = Nothing						
End if
	

Sub CheckMinSeverity (cn, Severity)
	Dim RS, UpdateMCD, WhereClause
	UpdateMCD = False
	WhereClause = "QID=" & SafeNum(iQPID)
	RemoveMilliSecondDataGeneric "tblRIRP1","RevDate",WhereClause
	Set RS = Server.CreateObject("ADODB.Recordset")
	rs.locktype = 3
	RS.Open "SELECT * FROM tblRIRP1 WHERE QID=" & SafeNum(iQPID), cn
	
	If Not RS.EOF Then
		If IsNull(RS("HSESeverity")) Then
			RS("HSESeverity") = Severity
			UpdateMCD = True
		ElseIf RS("HSESeverity") < Severity Then
			RS("HSESeverity") = Severity
			UpdateMCD = True
		End If		
	End If
	rs.update
	RS.Close
	cn.execute "sp_UpdateRIRSeverity @QID=" & SafeNum(iQPID)
	Set rs = Nothing
	
	if UpdateMCD Then
		MsgID=115
		UpdateMajorChangeDate "tblRIRp1", dtRptDate, lOrgNo
	End If
End Sub



Sub CheckClass(cn)
	Dim RS, strSQL,rQID
	'Sreedhar - Changes this Procedure to change Severity,Protect and pre-poplate the AccessList
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT * FROM tblRIRP1 with (NOLOCK) WHERE QID=" & SafeNum(iQPID)  , cn
	If Not RS.EOF Then
		If RS("SLBInv") AND RS("HSE") AND RS("Class")=1 AND RS("HSESeverity")<4 Then
			MsgID=114			
		End IF
	END IF
	RS.Close
	Set rs = Nothing
	If MsgID=114 Then cn.execute "spR_AutoFatalityProtect " & SafeNum(iQPID)	
End Sub	
	
	
	
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
			<p class=title id=styleMedium>Processing RIR - Loss Report (1)</p>
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


</BODY>
</HTML>


<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSpers2.asp;7 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1652495] 29-AUG-2012 08:14:36 (GMT) MPatil2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '       3*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '       4*[1837481] 20-MAY-2014 10:40:11 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture." %>
<% '       5*[1838252] 23-MAY-2014 11:14:16 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '       6*[1867842] 13-NOV-2014 11:40:26 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '       7*[1915045] 18-AUG-2015 15:01:07 (GMT) VSharma16 %>
<% '         "ENH077171-SAXON - HSE Personnel Loss Tab addition" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSpers2.asp;7 %>
