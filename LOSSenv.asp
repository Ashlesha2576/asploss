<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="../Inc_Java_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  If IsIE then Response.Expires = -1

Dim dtRptDate, sRptNo, lOrgNo, iQPID,sHref, tmp
Dim cn, rs, rsP1, sKey, sSQL,PV,Solid,Liquid,allunit
Dim dMaterials, dTargets, dUnits
Dim PDHabitat, PDFlora, Description,ACLDefined
Dim SSLetters, SSNotices, SSFines, SSPermitActions, SSReprimands, SSContractLoss
Dim SSJobRemoval, SSLBI, SSComplaints, SSNegMedia, SSLawsuits, SSCCSanctions,matval

Set pv = GetVariables()

lOrgNo = pv("OrgNo")
dtRptDate = pv("rptDate")
iQPID = Request.QueryString("QPID")
Set cn = GetNewCn()
	
sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
sSQL = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID) 
Set rsP1 = cn.Execute(sSQL)

	
sSQL = "SELECT * FROM tblRIRenv With (NOLOCK) WHERE QPID=" & SafeNum(iQPID) 
Set RS = cn.Execute(sSQL)

InitVars(rs)
rs.close
set rs = nothing
EnvMaterial Cn
%>

<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	<SCRIPT LANGUAGE="javascript">
	 var solid='<%=Solid%>'
	 var liquid='<%=Liquid%>'	
	 var allunit='<%=allunit%>'
		function On_Change(i){
			var mat,unit,mat1
			var f = window.document.frmEnv;
			var selmat = trimString(eval('f.material_' + i + '.options[f.material_' + i + '.selectedIndex].value'));

			var mTy=selmat.split(":")

			if  (mTy[0]==24)
			{

			 mat=allunit.split(";")

			}
			else
			{

		    if(mTy[1]=='S') mat=solid.split(";")
			if(mTy[1]=='L') mat=liquid.split(";")
			}
			unit=eval('f.unit_' + i)
			unit.options.length=1
			for(var j=0;j<mat.length;j++){
				mat1=mat[j].split(":")
				unit.options[unit.options.length] = new Option(mat1[2], mat1[0]+":"+mat1[1]);
			}
		}
		function prepSubmit() {
			//Validate routine
				
			var msg,iRow;
			var SubmitOK;
			var f = window.document.frmEnv;
			var MUTy,UTy;
			var seq,targettype,material,description,qty,unit;
			var errorsheader  = ''



			errorsheader += '_______________________________________________\n\n';
			errorsheader += 'The form was not saved because of the following error(s).\n';
			errorsheader += '_______________________________________________\n\n';
			
			msg = '';
			SubmitOK=true;			

			if (trimString(f.Description.value) == '') {
					msg = msg + 'Please enter a description for the event.\n';
					SubmitOK=false;
			}	

			iRow = 1;
			
			while (typeof eval('f.seq_' + iRow) != 'undefined') {
				targettype	 = trimString(eval('f.targettype_' + iRow + '.value'));
				material	 = trimString(eval('f.material_' + iRow + '.options[f.material_' + iRow + '.selectedIndex].value'));
				description	 = trimString(eval('f.description_' + iRow + '.value'));
				qty			 = trimString(eval('f.qty_' + iRow + '.value'));
				unit		 = trimString(eval('f.unit_' + iRow + '.options[f.unit_' + iRow + '.selectedIndex].value'));
				
				if ((targettype != "") | (material != "") | (description != "") | (qty != "")) {
					//They've entered some data, let's validate it.  Ingore UOM if it exists alone.
					
					//if they have data, I want, at a minimum, a target, a material, a qty, and a uom.  Description is optional.
					if (targettype == "") {
						msg += "Item " + iRow + ":  No Target specified.\n";
					}

					if (material == "") {
						msg += "Item " + iRow + ":  No Material specified.\n";
					}

					if (qty == "") {
						msg += "Item " + iRow + ":  No Quantity specified.\n";
					} else {
						if (isNaN(parseInt(qty))) {
							msg += "Item " + iRow + ":  Quantity should be an integer number.\n";
						} else {
							if (parseInt(qty) != qty) {
								msg += "Item " + iRow + ":  Quantity should be an integer number.\n";
							}
						}
					}
					
					if (unit == "") {
						msg += "Item " + iRow + ":  No Unit specified.\n";
					}
					
					if((material !="") && (unit!="") ){
					var mTyval=material.split(":")
						MUTy=material.split(":")[1];
						UTy=unit.split(":")[1];
						
						if(MUTy!=UTy && mTyval[0]!=24){
							if(MUTy=="L") 
								msg += "Item " + iRow + " unit should be Liquid Unit.\n"; 
							else 
								msg += "Item " + iRow + " unit should be Solid Unit.\n"; 
						}
					}
				}
				
				iRow++;
			}
			if (msg!='') {
				alert(errorsheader + msg);
				SubmitOK=false;
			}
			
			return SubmitOK;
		}
		
		function trimString(str) {
			str = this != window? this : str;
			return str.replace(/^\s+/g, '').replace(/\s+$/g, '');
		}
		
		function deleteItem(i) {
			return confirm('Are you sure you wish to DELETE record ' + i + '?');
		}

	</SCRIPT>
</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>

<%
displaymenubar(rsP1)
rsP1.Close
Set rsP1=Nothing	
if ACLDefined Then DisplayConfidential() 
%>

	<form name="frmEnv" method="post" onsubmit="return(prepSubmit())" action="LOSSenv2.asp<%=sKey%>">

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
	<%=EnvDetails(cn,iQpid)%>
	<table width="100%" border="1" cellPadding="0" cellSpacing="0">
		<tr>				
			<td class=reportheading align="left" colspan=5>
				<b><span id=styleSmall>Physical Damage Incident</span></b>
			</td>		
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellPadding="0" cellSpacing="0">
					<tr>
						<td>
							<input type="checkbox" name="PDHabitat" <%=Checked(PDHabitat)%>>
							<span id=styleSmall>Habitat: Damage to natural environment - ruts, tracks, wetlands disturbances, etc...</span>
						</td>
					</tr>
					<tr>
						<td>
							<input type="checkbox" name="PDFlora" <%=Checked(PDFlora)%>>
							<span id=styleSmall>Floral/Fauna: Loss of plant or animals to direct impact, contamination events, etc...</span>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</table>
	<!-- Sanctions and Scrutiny Incident -->
	<table width="100%" border="1" cellPadding="0" cellSpacing="0">
						
		<tr>				
			<td class=reportheading align="left" colspan=5>
				<b><span id=styleSmall>Sanctions and Scrutiny</span></b>
			</td>
			
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellPadding="0" cellSpacing="0">
					<tr>
						<td width="33%"><span id=styleSmall><b>Regulatory</b></span></td>
						<td width="33%"><span id=styleSmall><b>Client</b></span></td>
						<td width="33%"><span id=styleSmall><b>Legal/Public</b></span></td>
					</tr>
					<tr>
						<td><input type="checkbox" value="1" name="SSLetters" <%=Checked(SSLetters)%>><span id=styleSmall>Letters/Notices</span></td>
						<td><input type="checkbox" value="1" name="SSReprimands" <%=Checked(SSReprimands)%>><span id=styleSmall>Reprimands</span></td>
						<td><input type="checkbox" value="1" name="SSComplaints" <%=Checked(SSComplaints)%>><span id=styleSmall>Complaints</span></td>
					</tr>
					<tr>
						<td><input type="checkbox" value="1" name="SSNotices" <%=Checked(SSNotices)%>><span id=styleSmall>Regulatory Report Required</span></td>
						<td><input type="checkbox" value="1" name="SSContractLoss" <%=Checked(SSContractLoss)%>><span id=styleSmall>Contract Loss</span></td>
						<td><input type="checkbox" value="1" name="SSNegMedia" <%=Checked(SSNegMedia)%>><span id=styleSmall>Negative Media</span></td>
					</tr>
					<tr>
						<td><input type="checkbox" value="1" name="SSFines" <%=Checked(SSFines)%>><span id=styleSmall>Fines</span></td>
						<td><input type="checkbox" value="1" name="SSJobRemoval" <%=Checked(SSJobRemoval)%>><span id=styleSmall>Job Removal</span></td>
						<td><input type="checkbox" value="1" name="SSLawsuits" <%=Checked(SSLawsuits)%>><span id=styleSmall>Lawsuits</span></td>
					</tr>
					<tr>
						<td><input type="checkbox" value="1" name="SSPermitActions" <%=Checked(SSPermitActions)%>><span id=styleSmall>Permit Actions</span></td>
						<td><input type="checkbox" value="1" name="SSLBI" <%=Checked(SSLBI)%>><span id=styleSmall>Loss of Bonus or Incentive</span></td>
						<td><input type="checkbox" value="1" name="SSCCSanctions" <%=Checked(SSCCSanctions)%>><span id=styleSmall>Civil/Criminal Sanctions</span></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<!-- Description -->	
	<table width="100%" border="1" cellPadding="0" cellSpacing="0">
		<tr>
			<td class=reportheading align="left" colspan=1>
				<table cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td valign=top>
							<b><span id=styleSmall>Description</span></b>
						</td>
						<td>
							&nbsp;
						</td>
						<td>
							<span id=styleSmall>
								For any box checked in Sanctions and Scrutiny Incident,
								describe the character and extent of the incident:
								Species and number, acres of damage, type of sanctions, extent of lawsuit, etc...
							</span>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align=center>
				<table cellPadding=0 cellSpacing=0><tr>
					<td><TEXTAREA name=Description rows=7 wrap=virtual cols=60 onKeyUp='return validlength(document.frmEnv.Description, msgDescription,3500)'><%=Safedisplay(Description)%></TEXTAREA></td>
					<%=PutWordCounter("n3500","msgDescription")%>
				</tr></table>	
			</td>	
		</tr>
	</table>

	<%DisplayCost cn, 3, iQPID%>

	<table width=100%>
		<tr>
			<td align="right" valign="top">
				<input type="hidden" name="OrgNo" value="<%=lOrgNo%>">
				<input type="hidden" name="RptDate" value="<%=dtRptDate%>">
				<input type="hidden" name="QPID" value="<%=iQPID%>">
				<input type="submit" name="cmdSubmit" value="Save Data">
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
<%
	Cn.Close
	Set Cn=Nothing
%>

<%
Sub InitVars(rs)
	If rs.eof Then	
		PDHabitat		=	False
		PDFlora			=	False
		SSLetters		=	False
		SSNotices		=	False
		SSFines			=	False
		SSPermitActions	=	False
		SSReprimands	=	False
		SSContractLoss	=	False
		SSJobRemoval	=	False
		SSLBI			=	False
		SSComplaints	=	False
		SSNegMedia		=	False
		SSLawsuits		=	False
		SSCCSanctions	=	False
		Description		=	""
	Else
		PDHabitat		=	rs("PDHabitat")
		PDFlora			=	rs("PDFlora")
		SSLetters		=	rs("SSLetters")
		SSNotices		=	rs("SSNotices")
		SSFines			=	rs("SSFines")
		SSPermitActions	=	rs("SSPermitActions")
		SSReprimands	=	rs("SSReprimands")
		SSContractLoss	=	rs("SSContractLoss")
		SSJobRemoval	=	rs("SSJobRemoval")
		SSLBI			=	rs("SSLBI")
		SSComplaints	=	rs("SSComplaints")
		SSNegMedia		=	rs("SSNegMedia")
		SSLawsuits		=	rs("SSLawsuits")
		SSCCSanctions	=	rs("SSCCSanctions")
		Description		=	rs("Description")
	End If
End Sub


Function Checked(condition)
	If Condition Then
		Checked = " checked "
	Else 
		Checked = ""
	End IF
End Function

Sub EnvMaterial(cn)
	dim rs
	Set dMaterials = server.CreateObject("Scripting.dictionary")
	Set dTargets = server.CreateObject("Scripting.dictionary")
	Set dUnits = server.CreateObject("Scripting.dictionary")
	
	Set rs = cn.execute("SELECT * FROM tlkpEnvMaterials With (NOLOCK) ORDER BY Description")
	
	do while not rs.eof
		dMaterials.Add clng(rs("ID"))&":"&rs("Type"),cstr(rs("Description"))
		rs.movenext
	loop
	
	Set rs = cn.execute("SELECT * FROM tlkpEnvTargets With (NOLOCK) ORDER BY Description")
	do while not rs.eof
		dTargets.Add clng(rs("ID")),cstr(rs("Description")) 		
		rs.movenext
	loop

	Set rs = cn.execute("SELECT ID, Unit,Type FROM tlkpEnvUOM With (NOLOCK) ORDER BY Type,Unit")
	Solid=""
	Liquid=""
	allunit=""
	do while not rs.eof
		dUnits.Add clng(rs("ID"))&":"&rs("Type"),cstr(rs("Unit"))
		If rs("Type")="S" then Solid=Solid&";"& clng(rs("ID"))&":"&rs("Type")&":"&cstr(rs("Unit"))
		If rs("Type")="L" then Liquid=Liquid&";"& clng(rs("ID"))&":"&rs("Type")&":"&cstr(rs("Unit"))
		allunit=allunit&";"& clng(rs("ID"))&":"&rs("Type")&":"&cstr(rs("Unit"))
		rs.movenext
	loop
	Solid  = mid(Solid,2)
	Liquid = mid(Liquid,2)
	allunit = mid(allunit,2)
End Sub

Function EnvDetails(cn,iQPID)
	Dim tmp, rownumber, maxSeq	
	
	tmp = ""
	tmp = tmp & "<table width=""100%"" border=""1"" cellPadding=""0"" cellSpacing=""0"">" & vbCRLF
	tmp = tmp & "<tr><td class=reportheading align=""Left"" colspan=7><span id=styleSmall>Material(s) Involved In Discharge/Disposal Incident</span></td></tr>" & vbCRLF
	tmp = tmp & "<tr><td align = center colspan=7 class=boxednote id=styleSmall>To add more items fill all the lines displayed and click save.<br>To delete an entry click on the ""Item Number"".</td></tr>"
	tmp = tmp & "<tr>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Item</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Target Of Pollution</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Contained in Secondary Containment</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Material</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Description</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Quantity</span></td>" & vbCRLF
	tmp = tmp & "	<td align=Center><span id=styleSmall>Unit</span></td>" & vbCRLF
	tmp = tmp & "</tr>" & vbCRLF

	Set rs = cn.execute("SELECT * FROM tblRIREnvDetail With (NOLOCK) WHERE QPID = " & SafeNum(iQPID) & " ORDER BY Seq")
	rownumber = 0
	maxSeq = 0
	do while not rs.eof
		rownumber = rownumber + 1
		tmp = tmp & DetailRow(rownumber, dTargets, dMaterials, dUnits, rs("Seq"), rs("TargetType"),rs("Containment"), rs("Material"), rs("QTY"), rs("UNIT"), rs("Description"),iQPID,true)
		if rs("seq") > maxSeq Then maxSeq = rs("seq")
		rs.movenext
	Loop
	RS.Close
	Set RS=Nothing
	'Add one or two blank rows
	tmp = tmp & DetailRow(rownumber+1, dTargets, dMaterials, dUnits, maxSeq+1,0,0,0,"",0,"",iQPID,false)
	If RowNumber = 0 Then
		tmp = tmp & DetailRow(rownumber+2, dTargets, dMaterials, dUnits, maxSeq+2,0,0,0,"",0,"",iQPID,false)
	End If

	tmp = tmp & "<tr>" & vbCRLF
	tmp = tmp & "	<td align=Center colspan=7 style=""color:red;"">Note:  The Cost of the <b>materials</b> lost in spills or other accidental releases (entered above) are recorded in the <b>Assets Loss</b> section of the event report. <b>Lost Products</b> in the Assets section of the Main page should be checked as well, if any material was lost.</td>" & vbCRLF
	tmp = tmp & "</tr>" & vbCRLF

	tmp = tmp & "</table>"
	
	EnvDetails = tmp
End Function



Function DetailRow(rn,dT,dM,dU,seq,target,Containment,material,qty,unit,desc,QID,existingRow)
	dim tmp,val,Mty,sSQLU,rsP1U,testval,Mtyid
	dim k
	
	tmp =		"<tr>" & vbCRLF
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<span id=styleSmall><b>" & vbCRLF
	If existingRow Then
		tmp = tmp & "			<a href='LossEnv2.asp?OrgNo=" & lOrgNo & "&RptDate=" & server.URLEncode(dtRptDate) & "&QPID=" & QID & "&seq=" & seq & "&Delete=1' onclick='return deleteItem(" & rn & ")'>" & rn & "</a>" & vbCRLF
	Else
		tmp = tmp & "			" & rn & vbCRLF
	End if
	tmp = tmp & "		</b></span>" & vbCRLF
	tmp = tmp & "		<input type=hidden name=seq_" & rn & " value=" & seq & ">" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF	
	'Targets
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<select name=targettype_" & rn & ">" & vbCRLF
	tmp = tmp & "			<option value="""">" & vbCRLF
	for each k in dT.keys
		tmp = tmp & "			<option value=" & k 
		
		if trim(k) = trim(target) then
			tmp = tmp & " selected "
		end if
		tmp = tmp & ">" & dT(k) & vbCRLF
	next
	tmp = tmp & "		</select>" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF
	
	'containment
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<input type=""checkbox"" name=containment_" & rn & " value=""1"" " & IIF(Containment=1,"checked","") & ">" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF
	
	'Materials
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<select name=material_" & rn & " onchange=""On_Change('"&rn&"')"">" & vbCRLF
	tmp = tmp & "			<option value=''>" & vbCRLF
	
	for each k in dM.keys
		tmp = tmp & "			<option value=" & k 
		val=split(K,":")
		if trim(val(0)) = trim(material) then
			tmp = tmp & " selected "
			Mty=val(1)
			Mtyid=val(0)
		end if
		tmp = tmp & ">" & dM(k) & vbCRLF
	next
	tmp = tmp & "		</select>" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF

	'Description
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<input type=""text"" name=description_" & rn & " value=""" & replace(desc,"""","&quot;") & """ maxlength='50'>" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF

	'QTY
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<input type=""text"" name=qty_" & rn & " size=3  maxlength=5 value=" & qty & " >" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF

	'Unit
	tmp = tmp & "	<td align=Center>" & vbCRLF
	tmp = tmp & "		<select name=unit_" & rn & ">" & vbCRLF
	tmp = tmp & "			<option value=''>Select Unit" & vbCRLF

	for each k in dU.keys
		val=split(K,":")
		
		if Mtyid=24 then
		tmp = tmp & "			<option value=" & k 
			if trim(val(0)) = trim(unit) then
				tmp = tmp & " selected "
			end if
			tmp = tmp & " " & unit & " "
			tmp = tmp & ">" & dU(k) & vbCRLF
		else
		if Mty=val(1)  Then
			tmp = tmp & "			<option value=" & k 
			if trim(val(0)) = trim(unit) then
				tmp = tmp & " selected "
			end if
			tmp = tmp & " " & unit & " "
			tmp = tmp & ">" & dU(k) & vbCRLF
           end if
		   end if
	next
	tmp = tmp & "		</select>" & vbCRLF
	tmp = tmp & "	</td>" & vbCRLF

	tmp = tmp & "</tr>" & vbCRLF
	
	DetailRow = tmp
End Function
%>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSenv.asp;3 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1373763] 09-JUL-2010 12:21:53 (GMT) PMakhija %>
<% '         "Swift# 2471356-Modification to Env. Loss Tab" %>
<% '       3*[1619376] 22-JUN-2012 11:59:59 (GMT) ATuscano %>
<% '         "SWIFT #2638011 - Feature: Added Character Limitation Indicator and Character Counter for all Multiline Textboxes" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSenv.asp;3 %>
