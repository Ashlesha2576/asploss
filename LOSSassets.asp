<%@ Language=VBScript %>
<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()
  If IsIE then Response.Expires = -1
%>

<%
Dim sOrg, dtRptDate, sRptNo, cn, RS, lOrgNo, sKey,WidthFactor
Dim bPers, bAuto, bEnv, bOth, RS1

Dim sTemp, sTemp2, sTemp3, iTemp, RSlkp, sSel,bSQ

Dim  dtTemp, sHref
Dim iRows, iCntr, bNew, conn, iQPID
DIM ACLDefined,VarHideLegacyInvestigation
	
	VarHideLegacyInvestigation = fncD_Configuration("HideLegacyInvestigation")
	
	WidthFactor = 1
	if GetBrowserType() = "MSIE" Then
		WidthFactor = 1.8
	End iF

	Set cn = GetNewCn()
	Set RS = Server.CreateObject("ADODB.Recordset")
	
	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")
	
	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	sTemp = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID) 
	RS1.Open sTemp, cn
	If not RS1.EOF then	
		bSQ = Rs1("ServiceQuality")
	end if

	Function GetAssetStatus(Name,Sel,cn)
	Dim Str,SQL,RS,val,sTemp
		SQL = "SELECT ID, Description FROM tlkpRIRAssetStatus With (NOLOCK) Order By ID"
		Set RS=cn.execute(SQL)
		Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
		Str=Str&"	<option value='0'>"&vbCRLF
		Sel=Ucase(Sel)
		While Not RS.EOF 									
			Val = UCase(Trim(RS("ID")))
			if Val=Sel then sTemp="SELECTED" else sTemp=""
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& SafeDisplay(RS("Description")) &vbCRLF
			RS.MoveNext
		Wend
		Str=Str&"</select>"	&vbCRLF
		RS.Close
		Set RS=Nothing
		GetAssetStatus=Str
	End Function
	
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
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& SafeDisplay(RS("Unit")) &vbCRLF
			RS.MoveNext
		Wend
		Str=Str&"</select>"	&vbCRLF
		RS.Close
		Set RS=Nothing
		GetUnits=Str
	End Function
	
	
	Function getLossType(Name,Sel,cn)
	Dim Str,SQL,RS,val,sTemp
		SQL = "SELECT ID, Description FROM tlkpLossSubCategories With (NOLOCK) WHERE LossCatID = 1 ORDER BY Description"
		Set RS=cn.execute(SQL)
		Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
		Str=Str&"	<option value=''>(Select One)"&vbCRLF
		
		While Not RS.EOF 									
			Val = Trim(RS("ID"))
			if Val=Sel then sTemp="SELECTED" else sTemp=""
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& SafeDisplay(RS("Description")) &vbCRLF
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
	
	<script language="JavaScript">
		<!--
	    function trimString(str) 
	    {
	        str = this != window? this : str;
	        return str.replace(/^\s+/g, '').replace(/\s+$/g, '');
	    }
	    function prepSubmit() 
	    {
	        var ctr = 0;
	        var LD, RD, OC;
	        var errorsheader = '';
	        var errorsmsg = '';
	        var frmdoc = document.frmRIR;
			
			
	        errorsheader += '_______________________________________________\n\n';
	        errorsheader += 'The form was not saved because of the following error(s).\n';
	        errorsheader += '_______________________________________________\n\n';
			
	        for (ctr = 1;typeof eval('document.frmAssets.cmbType' + ctr) != 'undefined';ctr++) 
	        {
	            var CmpChk = eval('document.frmAssets.txtComputer' + ctr + '.checked');
	            if (eval('document.frmAssets.cmbType' + ctr + '.options[document.frmAssets.cmbType' + ctr + '.selectedIndex].value') != '') 
	            {
	                LD = eval('document.frmAssets.txtQty' + ctr + '.value') ? eval('document.frmAssets.txtQty' + ctr + '.value') : 0;
				
	                if(trimString(eval('document.frmAssets.txtDesc' + ctr + '.value')) == '')
	                {
					
	                    errorsmsg += 'Item ' + ctr + ': Please enter a description.\n';
	                }
					
	                if((trimString(eval('document.frmAssets.txtRef' + ctr + '.value'))) == '' && (CmpChk== true))
	                {
					   
	                    errorsmsg += 'Item ' + ctr + ': Please enter a Ref/Part Number.\n';
	                }

	                if((trimString(eval('document.frmAssets.txtSN' + ctr + '.value')) == '') && (CmpChk== true))
	                {
					   
	                    errorsmsg += 'Item ' + ctr + ': Please enter a Serial Number.\n';
	                }	

	                if (LD <= 0 || isNaN(LD))	
	                {
	                    errorsmsg += 'Item ' + ctr + ': Please enter a valid quantity (greater than zero).\n';
	                }
										
					
	            }
	            else
	            {
	                var temp=trimString(eval('document.frmAssets.txtDesc' + ctr + '.value'));
	                temp+=trimString(eval('document.frmAssets.txtRef' + ctr + '.value'));
	                temp+=trimString(eval('document.frmAssets.txtQty' + ctr + '.value'));
	                temp+=trimString(eval('document.frmAssets.txtAssetNO' + ctr + '.value'));
	                temp+=trimString(eval('document.frmAssets.txtUnit' + ctr + '.options[document.frmAssets.txtUnit' + ctr + '.selectedIndex].value'));
	                if(temp.length > 0)
	                {
	                    errorsmsg += 'Item ' + ctr + ': Please select a type.\n';
	                }
	                if (eval('document.frmAssets.txtContractor' + ctr + '.options[document.frmAssets.txtContractor' + ctr + '.selectedIndex].value') != '0') 
	                {
	                    errorsmsg += 'Item ' + ctr + ': Please select a type.\n';
	                }
	            }
	        }
			
	        if (errorsmsg != '')
	        {
	            alert(errorsheader += errorsmsg);
	            return false;
	        }
	        else
	        {
	            return true;
	        }
	    }
		

	    function Computer_OnClick(ctr) {
		
	        var D1 = eval('document.frmAssets.txtComputer' + ctr + '.checked');
	        var Msg
	        if(D1==true)
	        {
	            Msg='               COMPUTER or LAPTOP LOSS \n\n' 
	            Msg=Msg + 'Check this box ONLY if either a computer or laptop was lost or stolen.\n\n' 
	            Msg=Msg + 'ONLY check Adequately Protected flag if ALL recommended risk'
	            Msg=Msg + 'control measures according to IT Security	Standards '
	            Msg=Msg + 'were in place at the time of the event.\n\n'
	            Msg=Msg + 'You must also enter a number in the quantity field.\n\n'
				<%If bSQ and (Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideLegacyInvestigation))) THEN%>
					Msg=Msg + 'Selecting Computer Loss will disable ability to link a Parent Investigation.'
					Msg=Msg + 'If you have linked a parent investigation to this RIR, It will be deleted after selecting Computer Loss.'
				<%End IF%>
	            window.alert(Msg);				
	        }
	        else
	        {
	            var D2 = eval('document.frmAssets.txtProtected' + ctr);
	            var D3 = eval('document.frmAssets.txtAInv' + ctr);
	            D2.checked=false;
	            D3.checked=false;
	        }
	        return true 
	    }
		
	    function Protected_OnClick(ctr) {
		
	        var D1 = eval('document.frmAssets.txtComputer' + ctr + '.checked');
	        var D2 = eval('document.frmAssets.txtProtected' + ctr );
	        var D3 = eval('document.frmAssets.txtAInv' + ctr );
	        if(D1==false){
	            if(D2.checked) return false; 
	            if(D3.checked) return false;
	        }
	        return true 
	    }

	    function extTrim(txt) {
	        //Like trim, but trims CR, LF, TAB, and SPACE
			
	        var trimChars, startPos, stopPos, foundText, idx
			
	        trimChars = " \t\n\r";
	        startPos = -1;
	        stopPos = -1;
			
	        for (idx = 0;idx<txt.length;idx++) {
	            if (trimChars.indexOf(txt.charAt(idx)) == -1) {
	                if (!foundText)	startPos = idx
	                foundText = true
	                stopPos = idx
	            }
	        }

	        if (startPos != -1)
	            return  txt.substr(startPos,stopPos-startPos+1);
	        return ""
	    }		
		
	    function cmdDelete_onclick() {
	        var bConfirm = window.confirm('Are you sure you wish to DELETE this record');
	        return (bConfirm) 
	    }
            //-->	

	    function ContractorName_onchange(Obj) {
	        var UrlStr = "../Utils/SearchTPSupplier.asp?source=0&optname=" + Obj.name		    		    
	        if (Obj.options[Obj.selectedIndex].text == "(Search OEM)") {		        
	            open(UrlStr, 'searchTPSupplier', 'height=450,width=900,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
	            Obj.selectedIndex = Obj.options.length - 1;
	        }

	    }

	    function addOptionTPSupplier(sText, sValue, optName) {
	        var opt, seq, temp
	        sValue = sValue.replace(/'/i, "''");
	        opt = eval("document.frmAssets." + optName);
	        fValue = sValue.split(":")
	        opt.options[opt.options.length] = new Option(sText, fValue[0]);
	        opt.selectedIndex = opt.options.length - 1;

	    } 

	</script>	
	
</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>

<%
displaymenubar(RS1)
RS1.Close

RS1.Open "SELECT Count(*) AS RecCount FROM tblRIRAssets With (NOLOCK) where QPID='" & SafeNum(iQPID) & "'", cn 'WHERE OrgNo=" & lOrgNo & " AND RptDate='" & dtRptDate & "' ", cn
iRows = RS1("RecCount")+1
if iRows<2 then iRows=2 
RS1.Close
Set RS1=Nothing
if ACLDefined Then DisplayConfidential()
%>

<form name="frmAssets" method="post"  action="Lossassets2.asp<%=sKey%>"  onsubmit="return prepSubmit()">

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
		<td align = center colspan=13 class=boxednote id=styleSmall>
			To add more items fill all the lines displayed and click save.<br>
			To delete an entry click on the "Item Number".
		</td>
	</tr>		
	<tr>
		<td align=center rowspan=2>Item</td>				
		<td align=center rowspan=2>Type</td>
		<td align=center rowspan=2>Description of loss</td>
        <td align=center rowspan=2>Asset Number</td>
        <td align=center rowspan=2>OEM</td>
		<td align=center Colspan=3>Computer Related</td>
		
		<td align=center rowspan=2>Ref/Part</br> Number <%=mSymbol%></td>
		<td align=center rowspan=2>Serial<br> Number <%=mSymbol%></td>
		<td align=center rowspan=2>Qty <%=mSymbol%></td>
		<td align=center rowspan=2>Unit</td>
		<td align=center rowspan=2>Status</td>
	</tr>
	<tr>
		<td align=center>Computer Loss</td>
		<td align=center>Adequately<BR>Protected*</td>
		<td align=center>Asset<BR>Investigation**</td>
	</tr>
	<%				
	sTemp = ""
	sTemp = "SELECT * FROM tblRIRAssets With (NOLOCK) WHERE QPID=" & iQPID & " ORDER BY Seq"
	RS.Open sTemp, cn
			
	For iCntr = 1 to iRows
		bNew = False
		If RS.EOF or RS.BOF Then bNew = True%>
	<tr>
		<td align=center>
			<%If not bnew Then 
				Response.Write "<A class=em href=Lossassets2.asp" & sKey & "&Delete=1&ID=" & RS("SEQ") & " onclick='return cmdDelete_onclick()'>" & iCntr & "</A>" 
			else
				Response.Write iCntr
			End If
		%>		
		</td>
		
		<%If bNew Then sTemp = "" Else sTemp=RS("Type")%>
		<td><%=getLossType("cmbType"&iCntr,sTemp,Cn)%></td>
							
		<td align='center'>					
			<%If bNew Then sTemp = "" Else sTemp=RS("Description")%>
			<input type="text" name="txtDesc<%=iCntr%>" size="<%=9 * WidthFactor %>" value="<%=displayQuotes(sTemp)%>">
		</td>

		<td align='center'>					
			<%If bNew Then sTemp = "" Else sTemp=RS("AssetNO")%>
			<input type="text" name="txtAssetNO<%=iCntr%>" style="width: 60px;" value="<%=displayQuotes(sTemp)%>">
		</td>

       <td>
							<select name="txtContractor<%=iCntr%>" style="width: 100px;" LANGUAGE="javascript" onchange="return ContractorName_onchange(this)">
							 <%
								sTemp = 0
								If Not bNew Then sTemp = RS("OEMValue")
                                
								If (Not IsNumeric(sTemp) )Then
									response.write "<option selected value='0'>No OEM Involved</option>"
								elseif(sTemp=0) then
									response.write "<option selected value='0'>No OEM Involved</option>"
								else
                                    response.write "<option selected value='0'>No OEM Involved</option>"
									response.write "<option selected value="&sTemp&">"&getContractorName(sTemp) & "</option>"
								End If
								%>
							<option value="0">(Search OEM)
							</select>
						</td>
        	
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("Computer")%>
			<input type="checkbox" name="txtComputer<%=iCntr%>" value='1'  <%if sTemp="1" Then response.write " checked "%> onclick='return Computer_OnClick(<%=iCntr%>)'>
		</td>
		
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("Preventable")%>
			<input type="checkbox" name="txtProtected<%=iCntr%>" value='1'  <%if sTemp="1" Then response.write " checked "%> onclick='return Protected_OnClick(<%=iCntr%>)'>
		</td>
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("AInvestigation")%>
			<input type="checkbox" name="txtAInv<%=iCntr%>" value='1'  <%if sTemp="1" Then response.write " checked "%> onclick='return Protected_OnClick(<%=iCntr%>)'>
		</td>
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("RefNo")%>
			<input type="text" name="txtRef<%=iCntr%>" size="5" value="<%=displayQuotes(sTemp)%>">
		</td>
					
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("SN")%>
			<input type="text" name="txtSN<%=iCntr%>" size="5" value="<%=displayQuotes(sTemp)%>">
		</td>
			
		<td align='center'>
			<%If bNew Then sTemp = "" Else sTemp=RS("QTy")%>
			<input type="text" name="txtQty<%=iCntr%>" size="2" value='<%=sTemp%>' maxlength='5'>
		</td>
			
		<%If bNew Then sTemp = "" Else sTemp=RS("Unit")%>
		<td><%=getUnits("txtUnit"&iCntr,sTemp,Cn)%></td>	
		
		<%If bNew Then sTemp = "" Else sTemp=Trim(RS("Status"))%>
		<td><%=getAssetStatus("txtStatus"&iCntr,sTemp,Cn)%></td>		
	</tr>
					
	<%
	If not bNew then RS.MoveNext 
	Next
	RS.Close%>
	<tr><td colspan=13 id=styleTiny>
	* - ONLY check this flag if ALL recommended risk control measures according to IT Security 
	Standards were in place at the time of the event.<BR>
	** - Check box to be used if Asset is determined to be lost ONLY as a result of an IT investigation into 
	assets that were not connected to the SLB Network for an inexplicable period of time.
	</td></tr>		
</table>
		
<% DisplayCost cn, 1, iQPID	
   Cn.Close()
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




<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSassets.asp;4 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1851145] 01-AUG-2014 07:03:11 (GMT) Rbhalave %>
<% '         "ENH034183 - IT Laptop Loss report enhancements" %>
<% '       3*[1852515] 11-AUG-2014 09:58:16 (GMT) Rbhalave %>
<% '         "ENH034183 - IT Laptop Loss report enhancements" %>
<% '       4*[1915048] 18-AUG-2015 15:01:51 (GMT) VSharma16 %>
<% '         "ENH077303-Saxon - Asset Loss Tab" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSassets.asp;4 %>
