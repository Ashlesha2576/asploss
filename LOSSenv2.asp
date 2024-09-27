<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  'checkguest()
  
Dim lOrgNo, dtRptDate, iQPID, sKey, PV
Dim cn, RS, sSQL,Errors, sHref, iRow
Dim dMats, dMat, key,ACLDefined,MsgID,DelMsg
Set pv=GetVariables()

lOrgNo		= PV("OrgNo")
dtRptDate	= PV("rptDate")
iQPID		= PV("QPID")

Set cn = GetNewCN()
sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)

'Check for DELETE
'==================================================================
If pv("delete") = "1" Then
		Set RS = cn.Execute("Select * from tblRIREnvDetail Where QPID= " & SafeNum(iQPID) & " AND Seq = " & SafeNum(pv("seq")))
		If Not RS.EOF Then
			LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossEnv2.asp",129,""
			cn.Execute "DELETE FROM tblRIREnvDetail WHERE QPID= " & SafeNum(iQPID) & " AND Seq = " & SafeNum(pv("seq"))
			DelMsg=server.URLEncode("Entry Successfully Deleted")
			cn.execute("sp_UpdateDateTime " & SafeNum(iQPID) & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
		else
			DelMsg=server.URLEncode("Entry is Not Found")
		End If
		Set RS = Nothing
		cn.close
	    Set cn = Nothing
		Response.Redirect("LossEnv.asp" & sKey & "&msg=" &Delmsg )
End If



'Load up the Material pollution records into a dict of dicts.
'==================================================================
Set dMats = server.CreateObject("Scripting.dictionary")
iRow = 1
dim x,zzz
do while pv.exists("seq_" & irow)
	Set dMat = server.CreateObject("Scripting.dictionary")
	for each x in split("seq,targettype,containment,material,qty,unit",",")
		if x="material" or x="unit" then
			zzz=pv(x & "_" & irow)
			If instr(1,zzz,":")>0 then zzz=split(zzz,":")(0)
			dMat.Add x,zzz
		else
			dMat.Add x, pv(x & "_" & irow)
		End IF
		'print x & irow & "= " & pv(x & "_" & irow)
	next
	dMat.Add "description", left(pv("description_" & irow),50)	
	dMats.Add iRow,dMat
	iRow = iRow +1
loop


	
'Validation PHASE		
'==================================================================
errors = ""

For each key in dMats
	Set dMat = dMats(key)

	if dMat("targettype") <> "" OR dMat("material") <> "" OR dMat("description") <> "" OR dMat("qty") <> "" Then
		if dMat("targettype") = "" Then
			errors = errors & "Item " & key & ":  No Target specified.<BR>"
		End If

		if dMat("material") = "" Then
			errors = errors & "Item " & key & ":  No Material specified.<BR>"
		End If

		if dMat("qty") = "" Then
			errors = errors & "Item " & key & ":  No Quantity specified.<BR>"
		Elseif not IsNumeric(dMat("qty")) Then
			errors = errors & "Item " & key & ":  Quantity should be an integer number.<BR>"
		End If

		if dMat("unit") = "" Then
			errors = errors & "Item " & key & ":  No Unit specified.<BR>"
		End If
	End If
Next

If PV("Description") = "" then
	errors = errors & "Please enter a description for the event.<BR>"
End if
					
If errors <> "" then 
	ShowErrors(errors)
	Response.End
End If

'Write PHASE		

Set RS = Server.CreateObject("ADODB.Recordset")
RS.LockType = 3		

sSQL = "SELECT * FROM tblRIRenv WHERE QPID=" & SafeNum(iQPID)
RS.Open sSQL, cn
			
If RS.EOF Then 
	RS.AddNew							
	 RS("QPID")= iQPID
	MsgID=128
Else
	MsgID=107
End if
			
RS("RevDate")= GetUTC()
RS("PDHabitat")=GetChkVal(PV("PDHabitat"))
RS("PDFlora")=GetChkVal(PV("PDFlora"))
RS("SSLetters")=GetChkVal(PV("SSLetters"))
RS("SSReprimands")=GetChkVal(PV("SSReprimands"))
RS("SSComplaints")=GetChkVal(PV("SSComplaints"))
RS("SSNotices")=GetChkVal(PV("SSNotices"))
RS("SSContractLoss")=GetChkVal(PV("SSContractLoss"))
RS("SSNegMedia")=GetChkVal(PV("SSNegMedia"))
RS("SSFines")=GetChkVal(PV("SSFines"))
RS("SSJobRemoval")=GetChkVal(PV("SSJobRemoval"))
RS("SSLawsuits")=GetChkVal(PV("SSLawsuits"))
RS("SSPermitActions")=GetChkVal(PV("SSPermitActions"))
RS("SSLBI")=GetChkVal(PV("SSLBI"))
RS("SSCCSanctions")=GetChkVal(PV("SSCCSanctions"))

If PV("Description") = "" Then
	RS("Description") = NULL
Else
	RS("Description") = left(PV("Description"),3500)
End If
							
RS.Update
RS.Close
			
'Update the Detail records.
For each key in dMats
	Set dMat = dMats(key)

	if dMat("targettype") <> "" OR dMat("material") <> "" OR dMat("description") <> "" OR dMat("qty") <> "" Then
		'We have data, we already validated it, so assume it's good.
		
		sSQL = "SELECT * FROM tblRIRenvDetail WHERE QPID=" & iQPID & " AND Seq =" & dMat("seq")
		RS.Open sSQL, cn

		if RS.EOF Then
			RS.AddNew			
			rs("QPID")		= iQPID
		End If

		for each x in split("seq,targettype,containment,material,description,qty,unit",",")
			rs(x)	= dMat(x)
		next
		rs("RevDate") = GetUTC()

		RS.Update
		RS.Close
	ElseIf dMat("targettype") = "" And dMat("material") = "" And dMat("description") = "" And dMat("qty") = "" And dMat("seq") <> 0 Then
		'Delete blank records
		sSQL = "DELETE FROM tblRIRenvDetail WHERE QPID=" & SafeNum(iQPID) & " AND Seq =" & SafeNum(dMat("seq"))
		cn.execute sSql
	End If
Next
Set rs = nothing

'Costs
UpdateCost cn, 3, iQPID, lOrgNo, dtRptDate
UpdateUserInfo iQPID,Cn

cn.CLose		
Set cn = Nothing

LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossEnv2.asp",MsgID,""								
Response.Redirect("LOSSenv.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))


Function GetChkVal(val)
	If Val<>"" Then GetChkVal="1" else GetChkVal="0"
End Function
		
Sub ShowErrors(txt)
	%>
	<html>
	<head>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	</head>
	<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>
		<span class=urgent>
			<%=txt%>
		</span>
		<br>
		<I><B>Hit the back button on your browser to correct these problems...</B></I>		
	</body>
	</html>
	<%
End Sub
%>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSenv2.asp;3 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1373763] 09-JUL-2010 12:21:53 (GMT) PMakhija %>
<% '         "Swift# 2471356-Modification to Env. Loss Tab" %>
<% '       3*[1652495] 29-AUG-2012 08:14:36 (GMT) MPatil2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSenv2.asp;3 %>
