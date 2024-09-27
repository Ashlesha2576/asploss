<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  
Dim sValText, lOrgNo, iCt 
Dim dtRptDate, bCurrPers, sKey,SCATSQL,rsn
Dim bNotify, iCtr, sValTemp, conn, RS, dtTemp, sHref

Dim sTemp, iTemp, bTemp, sTemp2,iQPID,sSQL1
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
		Set RS = conn.Execute("Select * from tblRIRAssets Where QPID=" & SafeNum(iQPID) & " AND Seq=" & SafeNum(trim(Request.QueryString("ID"))))
		If Not RS.EOF Then
			conn.Execute "DELETE FROM tblRIRAssets WHERE  QPID=" & SafeNum(iQPID) & " AND Seq=" & SafeNum(trim(Request.QueryString("ID")))
			LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossAssets2.asp",133,""	
			DelMsg=server.URLEncode("Entry Successfully Deleted")
			conn.execute("sp_UpdateDateTime " & SafeNum(iQPID) & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
		else
			DelMsg=server.URLEncode("Entry is Not Found")
		End If
		Set RS = Nothing
		conn.close
	    Set conn = Nothing
		Response.Redirect("Lossassets.asp" & sKey & "&msg=" &Delmsg)
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
		conn.execute "DELETE FROM tblRIRAssets WHERE QPID=" & SafeNum(iQPID) 
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.LockType = 3	
		MsgID=105
		For iCtr = 1 to iCt
			If bRowData(iCtr) Then
				iTemp = iCtr
				sTemp = "SELECT * FROM tblRIRAssets WHERE QPID=" & SafeNum(iQPID) & " AND Seq=" & SafeNum(iTemp)
				RS.Open sTemp, conn
			
				If RS.EOF Then 
					RS.AddNew							
					RS("QPID")		= iQPID
					RS("Seq")		= iTemp
				End if
					
				RS("RevDate")= Date()
			
				RS("Type") = Request.Form("cmbType" & iCtr)
				RS("Description") = left(Request.Form("txtDesc" & iCtr),50)
                RS("AssetNo") = left(Request.Form("txtAssetNo" & iCtr),50)
                RS("OEMValue") = Request.Form("txtContractor" & iCtr)
				RS("RefNo") = left(Request.Form("txtRef" & iCtr),30)
				RS("SN")=left(Request.Form("txtSN" & iCtr),30)
				sTemp = Request.Form("txtQty" & iCtr)
				If trim(sTemp) = "" Then sTemp = Null
				RS("Qty") = sTemp
				RS("Unit") = Request.Form("txtUnit" & iCtr)
				RS("Status")=Request.Form("txtStatus" & iCtr)
				If Request.Form("txtComputer" & iCtr)="1" Then 
					RS("Computer")=1 					
					Set rsn = Server.CreateObject("ADODB.RecordSet")
					SCATSQL ="SELECT PARENTID,Rationale,RptDate FROM tblSCAT_Parentchild LEFT JOIN tblrirp1 ON PARENTID=QID WHERE CHILDID ="&iQPID		
					rsn.open SCATSQL, conn
					if not rsn.eof then
					    sSQL1 = "DELETE FROM tblSCAT_Parentchild WHERE CHILDID="&SafeNum(iQPID)
					    conn.execute sSQL1 
					    sSQL1="DELETE FROM tblririnvdetails WHERE QPID="&SafeNum(iQPID)
					    conn.execute sSQL1	
					end if
					rsn.close
					set rsn=nothing						
					sTemp=Request.Form("txtProtected" & iCtr)
					If sTemp<>"1" Then sTemp="0"
					RS("Preventable")=sTemp 
					sTemp=Request.Form("txtAInv" & iCtr)
					If sTemp<>"1" Then sTemp="0"
					RS("AInvestigation")=sTemp
				Else 
					RS("Computer")=0
					RS("Preventable")=0
					RS("AInvestigation")=0
				End If
				RS.Update
					
				RS.Close
			End If
		Next
			
		'Assets Costs
		UpdateCost conn, 1, iQPID, lOrgNo, dtRptDate
		UpdateUserInfo iQPID,conn
		Set RS = Nothing
		Set conn = Nothing		
		LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossAssets2.asp",MsgID,""
		Response.Redirect("Lossassets.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
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
			<p class=title id=styleMedium>Processing RIR - Loss Report (4)</p>
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

<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSassets2.asp;3 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1652495] 29-AUG-2012 08:14:36 (GMT) MPatil2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '       3*[1915048] 18-AUG-2015 15:01:51 (GMT) VSharma16 %>
<% '         "ENH077303-Saxon - Asset Loss Tab" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSassets2.asp;3 %>
