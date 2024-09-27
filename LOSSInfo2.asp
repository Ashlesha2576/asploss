<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  
Dim sValText, lOrgNo, iCt 
Dim dtRptDate, bCurrPers, sKey
Dim bNotify, iCtr, sValTemp, conn, RS, dtTemp, sHref

Dim sTemp, iTemp, bTemp, sTemp2,iQPID
Dim ACLDefined,MsgID,DelMsg


	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")
	iCt = Request.Form("txtRows")
	Set conn = GetNewCN()
	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", conn)
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, conn)

	'Check for delete
	If Request.QueryString("Delete")=1 then 
		Set RS = conn.Execute("Select * from tblRIRInfo Where QPID="& SafeNum(iQPID) & " and Seq=" & SafeNum(trim(Request.QueryString("ID"))))
		If Not RS.EOF Then
			LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossInfo2.asp",131,""	
			conn.Execute "DELETE FROM tblRIRInfo WHERE QPID="& SafeNum(iQPID) & " AND Seq=" & SafeNum(trim(Request.QueryString("ID")))
			DelMsg=server.URLEncode("Entry Successfully Deleted")
			conn.execute("sp_UpdateDateTime " & SafeNum(iQPID) & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
		else
			DelMsg=server.URLEncode("Entry is Not Found")
		End If
		Set RS = Nothing
		conn.close
	    Set conn = Nothing
		Response.Redirect("LossInfo.asp" & sKey & "&msg=" &Delmsg )
	end if

	'Validation PHASE		
	sValText = ""
	bNotify = False
	Redim bRowData(iCt)
	
	For iCtr = 1 to iCt
		bRowData(iCtr) = False
		
		'Validate...
		sTemp = Request.Form("cmbType" & iCtr)
		If sTemp = "" Then
			iTemp = Len(Trim(Request.Form("txtDesc" & iCtr)))
			iTemp = iTemp + Len(Trim(Request.Form("txtRef" & iCtr)))
			iTemp = iTemp + Len(Trim(Request.Form("txtQty" & iCtr)))
			iTemp = iTemp + Len(Trim(Request.Form("txtUnit" & iCtr)))
			If iTemp > 0 Then		
				bNotify = True
				sValText = "No type was selected in row " & iCtr & ".<BR>"
			Else 
				bRowData(iCtr) = False
			End If
		Else
			If Len(Trim(Request.Form("txtDesc" & iCtr)))=0 then
				bNotify = True
				sValText = "No Description entered in row " & iCtr & ".<BR>"				
			Else
				bRowData(iCtr) = True
			End If
		End If
		
		'Validate Units (Numeric)
		sTemp = Trim(Request.Form("txtQty" & iCtr))
		If Len(sTemp) > 0 Then
			If Not IsNumeric(sTemp) Then
				bNotify = True
				sValText = "Qty in row " & iCtr & " must be numeric or blank.<BR>"
			End If
		End If
				
	Next
'WRITE PHASE		
	If bNotify = false Then	
		conn.execute "DELETE FROM tblRIRInfo WHERE QPID=" & SafeNum(iQPID) 		
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.LockType = 3	
		MsgID=108
		For iCtr = 1 to iCt
			If bRowData(iCtr) Then
				iTemp = iCtr
				sTemp = "SELECT * FROM tblRIRInfo WHERE QPID=" & SafeNum(iQPID)  & " AND Seq=" & SafeNum(iTemp)
				RS.Open sTemp, conn
			
				If RS.EOF Then 
					RS.AddNew							
					RS("QPID")		= iQPID
					RS("Seq")		= iTemp
				End if
					
				RS("RevDate")= Date()
			
				RS("Type") = Request.Form("cmbType" & iCtr)
				RS("Description") = left(Request.Form("txtDesc" & iCtr),50)
				RS("RefNo") = left(Request.Form("txtRef" & iCtr),30)
				sTemp = Request.Form("txtQty" & iCtr)
				If trim(sTemp) = "" Then sTemp = Null
				RS("Qty") = sTemp
				RS("Unit") = Request.Form("txtUnit" & iCtr)
				RS.Update
				RS.Close
			End If
		Next
		UpdateCost conn, 4, iQPID, lOrgNo, dtRptDate
		UpdateUserInfo iQPID,conn
		conn.close
		Set conn = Nothing
				
		LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossInfo2.asp",MsgID,""								
		Response.Redirect("LossInfo.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
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
			<p class=title id=styleMedium>Processing Loss Report
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

<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSInfo2.asp;2 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1652495] 29-AUG-2012 08:14:36 (GMT) MPatil2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSInfo2.asp;2 %>
