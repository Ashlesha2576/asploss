<%@ Language=VBScript %>
<%'option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  If IsIE then Response.Expires = -1


Dim RS,cn, sRptNo, bNR, RS1, sSel, sSelN
Dim bPers, bAuto, bEnv, bOth
Dim lOrgNo, dtRptDate, sKey, iQPID
Dim iKcost1
Dim sTemp, sTemp2, sTemp3, iTemp
DIM ACLDefined
	
	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")
	Set cn = GetNewCn()
	
	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	sTemp = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID) 
	RS1.Open sTemp, cn
	
	'SQL String
	sTemp = "SELECT * FROM tblRIRauto WHERE QPID=" & SafeNum(iQPID) 
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open sTemp, cn
		
	bNR = False
	If RS.EOF Then bNR = True

%>
<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	<SCRIPT LANGUAGE="javascript">

	function prepSubmit()
	{
		//Validate routine
			
		var msg;
		var SubmitOK;

		msg = '';
		SubmitOK=true;			
		
		
		if ((document.frmLOSSauto.optConvoy[0].checked || document.frmLOSSauto.optConvoy[1].checked) != true)
		{
				msg = msg + 'Please specify if vehicle was in a convoy.\n';
				SubmitOK=false;
		}	

		if ((document.frmLOSSauto.optDriverO[0].checked || document.frmLOSSauto.optDriverO[1].checked) != true)
		{
				msg = msg + 'Please specify if driver was only occupant.\n';
				SubmitOK=false;
		}
		if ((document.frmLOSSauto.optVeh[0].checked || document.frmLOSSauto.optVeh[1].checked || document.frmLOSSauto.optVeh[2].checked) != true)			
		{
				msg = msg + 'Please specify vehicle status.\n';
				SubmitOK=false;
		}	
		if ((document.frmLOSSauto.optCoBus[0].checked || document.frmLOSSauto.optCoBus[1].checked) != true)
		{
				msg = msg + 'Please specify if vehicle was on company business.\n';
				SubmitOK=false;
		}	


		if (!SubmitOK)
		{
			alert(msg);
		}
		return SubmitOK;
	}

</SCRIPT>
</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>
<%
displaymenubar(RS1)
RS1.Close
Set RS1=Nothing
%>
<%if ACLDefined Then DisplayConfidential() %>
<form name="frmLOSSauto" method="post" onsubmit="return(prepSubmit())" action="LOSSauto2.asp<%=sKey%>">

<table border=0 cellPadding=0 cellSpacing=0 width=100%>
    <tr><td class=boxednote id=styleTiny colspan=6><%=mSymbol%> - Mandatory fields.</td></tr> 
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


<%
	sTemp = ""
	sSel = ""	
	sSelN = ""	
	iTemp = 0
%>

<table width="100%" border="1" cellPadding="2" cellSpacing="0">		
	<tr>
		<td colspan=10 class=field>Driver Name:&nbsp;
			<%If bNR Then sTemp="" Else sTemp=RS("DriverName")%>
			<input type="textbox" name="DriverName" value="<%=DisplayQuotes(sTemp)%>" maxlength="50">
		</td>
	</tr>		
	<tr>				
		<td colspan="10" width="100%">
			<table width="100%" border="0" cellPadding="2" cellSpacing="2">
				<tr>
					<td>
						<%If bNR=False Then
							iTemp = RS("Convoy")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<span id=styleSmall >Was vehicle travelling in convoy?</span><%=mSymbol%>
						<input <%=sSel%>type="radio" name="optConvoy" value="1">Yes
						<input <%=sSelN%>type="radio" name="optConvoy" value="0">No
					</td>
															
					<td>
						<%If bNR=False Then 
							iTemp = RS("DriverO")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<span id=styleSmall >Was driver the only occupant?</span><%=mSymbol%>
						<input <%=sSel%>type="radio" name="optDriverO" value="1">Yes
						<input <%=sSelN%>type="radio" name="optDriverO" value="0">No
					</td>
				</tr>
														
				<tr>
					<td>
						<%If bNR=False Then iTemp = RS("VehicleStatus")%>
						<span id=styleSmall >Was vehicle:</span><%=mSymbol%>
						<%sSel = ""
						If iTemp = 1 then sSel= "checked "%>
						<input <%=sSel%>type="radio" name="optVeh" value="1">Company Owned
						<%sSel = ""
						If iTemp = 2 then sSel= "checked "%>
						<input <%=sSel%>type="radio" name="optVeh" value="2">Rented/Leased
						<%sSel = ""
						If iTemp = 3 then sSel= "checked "%>
						<input <%=sSel%>type="radio" name="optVeh" value="3">Personal Vehicle
					</td>
							
					<td>
						<%
						If bNR=False Then 
							iTemp = RS("CoBus")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<span id=styleSmall >On Company Business?</span><%=mSymbol%>
						<input <%=sSel%>type="radio" name="optCoBus" value="1">Yes
						<input <%=sSelN%>type="radio" name="optCoBus" value="0">No
					</td>
				</tr>
			</table>
		</td>
	</tr>
			
	<tr>
		<td width="20%" colspan="2"><span id=styleSmall ><b>Weather Conditions</b></span></td>
		<td width="40%" colspan="4"><span id=styleSmall ><b>Road Type</b></span></td>
		<td width="40%" colspan="4"><span id=styleSmall ><b>Accident Type</b></span></td>
	</tr>
			
	<tr>
		<td width="20%" colspan="2">
			<%
			sTemp=""
			sTemp2=""
			If bNR=False Then sTemp = RS("RoadCond")
			If bNR=False Then sTemp2 = RS("Visibility")%>
			<span id=styleSmall >
			<%sSel = ""
			If sTemp = "D" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCDry">Dry<br>
			<%sSel = ""
			If sTemp = "W" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCWet">Wet/Slick<br>
			<%sSel = ""
			If sTemp2 = "C" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCClear">Clear<br>
			<%sSel = ""
			If sTemp2 = "D" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCDust">Dust/Sandstorm<br>
			<%sSel = ""
			If bNR = False then 						
				If RS("Heat") <> 0 Then sSel= "checked "
			End If%>
			<input <%=sSel%>type="checkbox" name="chkWCHot">Extreme Heat<br>
			<%sSel = ""
			If sTemp2 = "F" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCFog">Fog<br>
			<%sSel = ""
			If sTemp = "I" then sSel= "checked "%>
			<input <%=sSel%>type="checkbox" name="chkWCIce">Snow/Icy
			</span>
		</td>
					
		<td width="40%" colspan="4" valign="top">
			<%
			If bNR=False Then sTemp = RS("RoadSurface")
			If bNR=False Then sTemp2 = RS("RoadGrade")%>
			<table width="100%" border="0" cellPadding="0" cellSpacing="0">
				<tr>
					<td>
						<%sSel = ""
						If sTemp = "P" then sSel= "checked "%>
						<input <%=sSel%>type="checkbox" name="chkRTPaved">Paved
					</td>
					<td>
						<%sSel = ""
						If sTemp = "U" then sSel= "checked "%>
						<input <%=sSel%>type="checkbox" name="chkRTUnpaved">Unpaved
					</td>
				</tr>
				<tr>
					<td>
						<%sSel = ""
						If sTemp = "O" then sSel= "checked "%>
						<input <%=sSel%>type="checkbox" name="chkRTOffRd">Off Road
					</td>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("RoadCurve") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkRTCurve">Curve
					</td>
				</tr>
				<tr>
					<td>
						<%sSel = ""
						If sTemp2 = "U" then sSel= "checked "%>
						<input <%=sSel%>type="checkbox" name="chkRTUp">Up a grade
					</td>
					<td>
						<%sSel = ""
						If sTemp2 = "D" then sSel= "checked "%>
						<input <%=sSel%>type="checkbox" name="chkRTDown">Down a grade
					</td>
				</tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then						
							If RS("RoadNarrow") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkRTNarrow">Narrow
					</td>
					<td>
						<%sSel = ""
						If bNR = False then					
							If RS("PoorSurf") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkRTPoor">Poor surface
					</td>
				</tr>
			</table>
		</td>
					
		<td width="40%" colspan="4" valign="top">				
			<table width="100%" border="0" cellPadding="0" cellSpacing="0">
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then															
							If RS("ATHitF") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitF">Hit vehicle in front</td>
					<td>
						<%sSel = ""
						If bNR = False then															
							If RS("ATSS") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATSS">Sideswipe</td></tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATHitB") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitB">Hit from behind</td>
					<td>
						<%sSel = ""
						If bNR = False then 	
							If RS("ATPass") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATPass">Passing</td></tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATBack") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATBack">Backed into</td>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATPassed") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATPassed">Being passed</td></tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATHitSO") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitSO">Hit stationary object</td>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATHitRun") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitR">Hit &amp; Run</td></tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATHitPed") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitP">Hit pedestrian</td>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATHitA") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATHitA">Hit animal</td></tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATRO") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATRoll">Rollover</td>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("ATRanOR") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkATRanOR">Ran off road
					</td>
				</tr>
				<tr>
					<td>
						<%sSel = ""
						If bNR = False then 
							If RS("HeadOC") <> 0 Then sSel= "checked "
						End If%>
						<input <%=sSel%>type="checkbox" name="chkHeadOC">Head-on Collision</td>
					<td>
					<td>&nbsp;</td>
				</tr>				
			</table>
		</td>
	</tr>
		
	<tr>
		<td width="50%" colspan="5" valign="top">				
			<table width="100%" border="0" cellPadding="0" cellSpacing="0">
				<tr>
					<td>
						<span id=styleSmall >Was alcohol/drugs involved?</span></td>
					<td>
						<%
						If bNR=False Then 
							iTemp = RS("Drugs")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optDrug" value="1">Yes
						<input <%=sSelN%>type="radio" name="optDrug" value="0">No</td></tr>
				<tr>
					<td>
						<%sTemp = ""
						If bNR=False Then sTemp = RS("Speed")%>
						<span id=styleSmall >Speed when accident occurred</span>
						<input type="text" name="txtSpeed" size="3" value='<%=sTemp%>' maxlength='4'></td>
					<td>
						<%sTemp = ""
						If bNR = False then sTemp = RS("SpeedUnit")
						sSel = ""
						IF sTemp = "M" Then sSel = "checked "%>
						<input <%=sSel%>type="radio" name="optSpeedU" value="1">mph
						<%sSel = ""
						IF sTemp = "K" Then sSel = "checked "%>
						<input <%=sSel%>type="radio" name="optSpeedU" value="2">km/h</td></tr>
				<tr>
					<td>
						<span id=styleSmall >Driving monitor present and working?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("Monitor")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optMonitor" value="1">Yes
						<input <%=sSelN%>type="radio" name="optMonitor" value="0">No</td></tr>
				<tr>
					<td>
						<span id=styleSmall >All persons wearing seatbelts?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("Seatbelts")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optSeatbelt" value="1">Yes
						<input <%=sSelN%>type="radio" name="optSeatbelt" value="0">No</td></tr>
			</table>
		</td>
				
		<td width="50%" colspan="5" valign="top">				
			<table width="100%" border="0" cellPadding="0" cellSpacing="0">
				<tr>
					<td>
						<span id=styleSmall >Driving certificate held for this unit?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("Certificate")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optCert" value="1">Yes
						<input <%=sSelN%>type="radio" name="optCert" value="0">No</td></tr>
				<tr>
					<td>
						<span id=styleSmall >Charged by police?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("Citation")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optCitation" value="1">Yes
						<input <%=sSelN%>type="radio" name="optCitation" value="0">No</td></tr>
				<tr>
					<td>
						<span id=styleSmall >Defensive driving training up to date?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("DD")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optDD" value="1">Yes
						<input <%=sSelN%>type="radio" name="optDD" value="0">No</td></tr>
				<tr>
					<td>
						<span id=styleSmall >Commentary drive up to date?</span></td>
					<td>
						<%
						If bNR=False Then
							iTemp = RS("CD")
							If iTemp <> 0 Then 
								sSel = "checked "
								sSelN = ""
							Else
								sSel = ""
								sSelN = "checked "
							End if
						End if%>
						<input <%=sSel%>type="radio" name="optCD" value="1">Yes
						<input <%=sSelN%>type="radio" name="optCD" value="0">No</td></tr>
			</table></td></tr>
				
</table>
		
<%
	DisplayCost cn, 2, iQPID
	RS.Close
	Set RS=Nothing
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



<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSauto.asp;1 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSauto.asp;1 %>
