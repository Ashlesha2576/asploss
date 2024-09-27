<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%
checktimeout()
If IsIE then Response.Expires = -1

Dim sOrg, dtRptDate, sRptNo, cn, RS, lOrgNo, sKey
Dim bPers, bAuto, bEnv, bOth, RS1 , iQPID

Dim sTemp, sTemp2, sTemp3, iTemp, RSlkp, sSel

Dim  dtTemp, sHref
Dim iRows, iCntr, bNew, conn
DIM ACLDefined




lOrgNo = Request.QueryString("OrgNo")
dtRptDate = Request.QueryString("rptDate")
iQPID = Request.QueryString("QPID")
Set cn = GetNewCn()
Set RS = Server.CreateObject("ADODB.Recordset")	

sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
sTemp = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID) 
Set RS1 = cn.Execute(sTemp)


	Function GetUnits(Name,Sel,cn)
	Dim Str,SQL,RS,val,sTemp
		SQL = "SELECT ID, Unit FROM tblUOM With (NOLOCK) WHERE QuestDisplay = 1 ORDER BY Unit"
		Set RS=cn.execute(SQL)
		Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
		Str=Str&"	<option value=''>"&vbCRLF
		Sel=Ucase(Sel)
		While Not RS.EOF 									
			Val = UCase(Trim(RS("ID")))
			if Val=Sel then sTemp="SELECTED" else sTemp=""
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& RS("Unit") &vbCRLF
			RS.MoveNext
		Wend
		Str=Str&"</select>"	&vbCRLF
		RS.Close
		Set RS=Nothing
		GetUnits=Str
	End Function
	
	
	Function getLossType(Name,Sel,cn)
	Dim Str,SQL,RS,val,sTemp
		SQL = "SELECT ID, Description FROM tlkpLossSubCategories With (NOLOCK) WHERE LossCatID = 4 ORDER BY Description"
		Set RS=cn.execute(SQL)
		Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
		Str=Str&"	<option value=''>(Select One)"&vbCRLF
		
		While Not RS.EOF 									
			Val = Trim(RS("ID"))
			if Val=Sel then sTemp="SELECTED" else sTemp=""
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& RS("Description") &vbCRLF
			RS.MoveNext
		Wend
		Str=Str&"</select>"	&vbCRLF
		RS.Close
		Set RS=Nothing
		getLossType=Str
	End Function

%>

<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	<script ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	function cmdDelete_onclick() {
		var bConfirm = window.confirm('Are you sure you wish to DELETE this record');
		return (bConfirm) 
	}
	//-->
	</script>
</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>

<%
displaymenubar(RS1)
RS1.Close

RS1.Open "SELECT Count(*) AS RecCount FROM tblRIRInfo with (NOLOCK) WHERE QPID=" & SafeNum(iQPID), cn
iRows = RS1("RecCount")+1
if iRows<2 then iRows=2 
RS1.Close
Set RS1=Nothing
if ACLDefined Then DisplayConfidential() 
%>

<form name="frmInfo" method="post"  action="LossInfo2.asp<%=sKey%>">

<table border=0 cellPadding=0 cellSpacing=0 width=100%>
	<tr class=reportheading>
		<TD align=left colspan=3>
			Report Date:&nbsp;<%=FmtDate(dtRptDate) & " (UTC)"%>
		</TD>
						
		<TD align=right colspan=3 class=field>
				Report Number:&nbsp;
				<span class=urgent >
				<%
				Response.Write "<A href='" & "RIRview.asp" & sKey & "'>" & getreportnumber(dtRptDate)& "</A>"%>													
		</TD>
	</tr>
</table>

<input type="hidden" name="txtRows" value="<%=iRows%>">

				
<table width="100%" border="1" cellPadding="2" cellSpacing="0">
	<tr>
		<td align = center colspan=7 class=boxednote id=styleSmall>
			To add more items fill all the lines displayed and click save.<br>
			To delete an entry click on the "Item Number".
		</td>
	</tr>		
	<tr>
		<td align=center>Item</td>				
		<td align=center>Type</td>
		<td align=center>Description of loss</td>
		<td align=center>Reference Number&nbsp;
			<span id=styleSmall >(Asset/Product #)</span></td>
		<td align=center>Quantity</td>
		<td align=center>Unit</td></tr>
	<%				
	sTemp = ""
	sTemp = "SELECT * FROM tblRIRInfo with (NOLOCK) WHERE QPID=" & SafeNum(iQPID) & " ORDER BY Seq"
	RS.Open sTemp, cn
			
	For iCntr = 1 to iRows
		bNew = False
		If RS.EOF or RS.BOF Then bNew = True%>
	<tr>
		<td align=center>
			<%If not bnew Then 
				Response.Write "<A class=em href=LossInfo2.asp" & sKey & "&Delete=1&ID=" & RS("SEQ") & " onclick='return cmdDelete_onclick()'>" & iCntr & "</A>" 
			else
				Response.Write iCntr
			End If
		%>		
		</td>
		
		
		<%If bNew Then sTemp = "" Else sTemp=RS("Type")%>
		<td><%=getLossType("cmbType"&iCntr,sTemp,Cn)%></td>
		
		<td align='center'>					
			<%sTemp = ""
			If bNew = False Then sTemp=RS("Description")%>
			<input type="text" name="txtDesc<%=iCntr%>" size="25" value="<%=DisplayQuotes(sTemp)%>">
			</td>
				
		<td align='center'>
			<%sTemp = ""
			If bNew = False Then sTemp=RS("RefNo")%>
			<input type="text" name="txtRef<%=iCntr%>" size="10" value="<%=DisplayQuotes(sTemp)%>">
			</td>
					
		<td align='center'>
			<%sTemp = ""
			If bNew = False Then sTemp=RS("Qty")%>
			<input type="text" name="txtQty<%=iCntr%>" size="2" value='<%=sTemp%>' maxlength=5>
			</td>
				
		<%If bNew Then sTemp = "" Else sTemp=RS("Unit")%>
		<td><%=getUnits("txtUnit"&iCntr,sTemp,Cn)%></td>				
	</tr>
					
	<%
	If not bNew then RS.MoveNext 
	Next
	RS.Close%>
			
</table>
		
<%	
	DisplayCost cn, 4, iQPID
	Cn.Close
	Set Cn=Nothing	
%>

				
		

<table width=100%>
	<tr>
		<td align="right" valign="top">
			<input type="submit" name="cmdSubmit" value="Save Data">
		</td>
	</tr>			
</table>

</form>
</body>
</html>



<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSInfo.asp;1 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSInfo.asp;1 %>
