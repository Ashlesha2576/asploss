<%@ Language=VBScript %>

<%
'*********************************************************************************************
'1. File Name		              :  Losstime.asp
'2. Description           	      :  Time loss data entry page
'3. Calling Forms   	          : 
'4. Stored Procedures Used        : 
'5. Views Used	   	              : 
'6. Module	   	                  : RIR (HSE/SQ)				
'7. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'    5-Aug-2009			     Nilesh Naik        	 	Modified - changed for NPT SWIFT # 2401608
'   22-Sep-2009				 Shailesh Kadam				Modified - Change the NPT Images.              
'   24-Sep-2009              Nilesh Naik                Modified - changed for fixing NPT bugs SWIFT # 2401608
'    2-Oct-2009              Nilesh Naik                Modified - changed for fixing NPT bugs SWIFT # 2401608                                          
'	06-Oct-2009				 Micheal Anthony			Modified - changed for NPT SWIFT # 240160
'   07-Oct-2009				 Shailesh Kadam             Modified - To Display correct Segment ID & NPT Image
'   11-Oct-2009              Shailesh Kadam             Modified - Change condition to accept NPT values 0.12
'	21-Oct-2009				 Micheal Anthony			Modified - SWIFT # 2434444 - NPT Issues - post go live
'   22-Oct-2009			     Micheal Anthony				NPT SWIFT #2434444 
'   30-Oct-2009			     Nilesh Naik				Modified - NPT SWIFT #2434444 .To fix NPT value should be allowed upto 2 decimals
'   30-Oct-2009				Shailesh					Modified - To Exclude some segment for severity check Swift# 2438856
'	03-Nov-2009				Micheal Anthony				Modified - SWIF # 2434444 - Code Correction
'   7-May-2014              Varun Sharma                 Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
'   21-Jun-2016              Varun Sharma                 NFT101068- RIR locking to prevent data integrity issues
'************************************************************************************************
%>

<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<!-- #INCLUDE FILE="RIRHelpText.asp"-->
<%checktimeout()
  If IsIE then Response.Expires = -1
%>

<%
Dim sOrg, dtRptDate, sRptNo, cn, RS, lOrgNo, sKey
Dim bPers, bAuto, bEnv, bOth, RS1, iQPID, strCrtieria
Dim sTemp, sTemp2, sTemp3, iTemp, RSlkp, sSel
Dim  dtTemp, sHref, strSelectedLossType 'Micheal SWIFT # 2401608 Added variable strSelectedLossType
Dim iRows, iCntr, bNew, intNPT,intAgreenpt,txtCausevalue,NoAgreement,NPTConfirmed,SQPQNPT,NoAgreementval,NPTConfirmedval
DIM ACLDefined
Dim msegment,AgreeNpt,txtnptval
Dim LossTypeCtr,LossDescID,TLExists,DefSeg
Dim mdblloss_g1 ,mdblloss_g2,mdblloss_g3 ,RS2,mintseverity, strExcellencePage, StrNPTImage 'Micheal SWIFT # 2401608 Added variable strExcellencePage
sHSEWarningText="*** Warning this report will be locked from editing Key HSE Fields in " & LockCountSQ & " day(s) ***"
'Micheal SWIFT # 2401608 Start
StrNPTImage = GetQuestServer() & "RIR/NPTImage.asp"
strExcellencePage = GetQuestServer() & "RIR/ExcellenceInExec.asp"
'Micheal SWIFT # 2401608 END
lOrgNo = Request.QueryString("OrgNo")
dtRptDate = Request.QueryString("rptDate")
iQPID = Request.QueryString("QPID")

Set cn = GetNewCN()
intNPT=getNPT(iQPID,cn)
intAgreenpt=0
SQPQNPT=0
' Start Add by Micheal for SWIFT # 2401608
cn.CursorLocation=3
' End Add

Set RS = Server.CreateObject("ADODB.Recordset")

'Shailesh 30-Oct-2009 Swift# 2438856
Dim blnNPT_Exempt
blnNPT_Exempt = ChkNPT_Exemption()
'Shailesh 30-Oct-2009

sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
sTemp = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID)
Set RS1 = cn.Execute(sTemp)


sTemp = "SELECT losscat_g1 , losscat_g2 , losscat_g3,SQSeverity,isnull(DisagreeNPT,0) as DisagreeNPT,CauseReview,isnull(SQPQNPT,0) as SQPQNPT ,NoAgreement,isnull(LegacyNpt,0) as LegacyNpt ,NPTConfirmed FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID)
Set RS2 = cn.Execute(sTemp)
mdblloss_g1 =  RS2("losscat_g1")
mdblloss_g2 =  RS2("losscat_g2")
mdblloss_g3 =  RS2("losscat_g3")
mintseverity = RS2("SQSeverity")

intAgreenpt=RS2("DisagreeNPT")
txtCausevalue=RS2("CauseReview")
SQPQNPT=RS2("SQPQNPT")
NoAgreementval=RS2("NoAgreement")
NPTConfirmedval=RS2("NPTConfirmed")
If NoAgreementval=1 Then NoAgreement = " checked=""checked""" else NoAgreement=""
If NPTConfirmedval=1 Then NPTConfirmed = " checked=""checked""" else NPTConfirmed=""
if intAgreenpt="0" then SQPQNPT=intNPT


RS2.Close
Set RS2=Nothing

Sub GetMatrixData()
	Dim cmdMatrix, rsMatrix, Cn

	Set Cn = GetNewCn()
	Cn.CursorLocation = 3
	Set rsMatrix = Server.CreateObject("ADODB.RecordSet")
	Set cmdMatrix = Server.CreateObject("ADODB.Command")
	
	With cmdMatrix
		.ActiveConnection = Cn
		.CommandType = adCmdStoredProc
		.CommandText = "SPRIR_GetTimeLossMatrix"
		
		'.Parameters.Append .CreateParameter("@MatrixId", adInteger, adParamInput, ,IIF(int_MatrixId=0,Null,int_MatrixId))

		Set rsMatrix = .Execute()
	End With
	
    Dim intCtr	
    intCtr = 4 
    Response.Write "<Script Language=JavaScript>" & vbCrLf
    Response.Write "var matrix = new Array(5);" & vbCrLf

    Response.Write "for(i=0;i<=4;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        matrix[i] = new Array(6);" & vbCrLf
    Response.Write "    }" & vbCrLf
	Do While Not rsMatrix.EOF	    
        Response.Write "    matrix[" & intCtr & "] = [" & rsMatrix("MatrixId") & "," & rsMatrix("CustSLBLoss1") & "," & rsMatrix("CustSLBLoss2") & "," & rsMatrix("NPTHours1") & "," & rsMatrix("NPTHours2") & "," & rsMatrix("QuestHours") & "];" & vbCrLf
        intCtr = intCtr - 1	
	    rsMatrix.MoveNext
	Loop
    Response.Write "</Script>" & vbCrLf
End Sub	

Function getCausereview(sel)
Dim str,sql,rs

    SQL = "select ID,Name from tlkpSQ_CauseReview  Order by ID"
    Set rs = cn.execute(SQL)
      Str="<select name='selCausereview' id='selCausereview'  onChange='Onchangecause()'>"
   ' Str=Str&"<option value='0'>(Please Select)"&vbCRLF
    if not rs.EOF then
        While not rs.EOF 
		    Str=Str&"<option id=styleTiny value='"&rs("ID")&"' "&iif(rs("ID")=sel," selected","")&">" & rs("Name") &vbCRLF				
        rs.MoveNext
        Wend
    end if
    Str=Str&"</select>"
    getCausereview = Str

End Function

'' Start Add by Micheal for SWIFT # 2401608
Function getLossType(Name,Sel,cn,iCntr1,isScript)
Dim Str,SQL,RS,val,sTemp,cmdLossType

        cn.CursorLocation = 3
		Set RS = Server.CreateObject("ADODB.RecordSet")
		Set cmdLossType = Server.CreateObject("ADODB.Command")
	
		With cmdLossType
			.ActiveConnection = cn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_fetchtabledata"
			
			.Parameters.Append .CreateParameter("@Name", adVarChar,1,100  ,"tlkpLossSubCategories")
			.Parameters.Append .CreateParameter("@Where", adVarChar,1,100 ,"LossCatID = 7")
			.Parameters.Append .CreateParameter("@Orderby", adVarChar,1,100 ,"Description")
			
			Set RS = .Execute()
		End With
		
	''SQL = "SELECT ID, Description FROM tlkpLossSubCategories With (NOLOCK) WHERE LossCatID = 7 ORDER BY Description"
	''Set RS=cn.execute(SQL)
	' *****************************************************************
    ' Code changed for NPT <<2401608>>
    ' *****************************************************************

	'' Start Add by Micheal for SWIFT # 2401608
	If isScript Then
		'Response.Write "<Script Language=JavaScript>"&vbCRLF
		Response.Write "var arrLossTypeId = new Array(" & RS.RecordCount & ");"&vbCRLF
		Response.Write "var arrLossTypeDesc = new Array(" & RS.RecordCount & ");"&vbCRLF
		Response.Write "arrLossTypeId[0]=0;"&vbCRLF
		Response.Write "arrLossTypeDesc[0]='(Select One)';"&vbCRLF
	Else	
	' End Add
		'Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
		'Micheal SWIFT # 2401608 Start (Chage Request)
		Str="<select name='"&Name&"' id='"&Name&"' onblur='calulatemoney("&iCntr1&")' onChange='checkLossType(" & iCntr1 & ");'>"&vbCRLF
		'Micheal SWIFT # 2401608 END (Chage Request)
		Str=Str&"	<option value=''>(Select One)"&vbCRLF
	'' Start Add by Micheal for SWIFT # 2401608
	End If
	' End Add
	strSelectedLossType="" 'Micheal SWIFT # 2401608 
	While Not RS.EOF
		Val = Trim(RS("ID"))
		' Start Add by Micheal for SWIFT # 2401608
		If isScript Then
			Response.Write "arrLossTypeId[" & RS.AbsolutePosition & "]=" & val & ";"&vbCRLF
			Response.Write "arrLossTypeDesc[" & RS.AbsolutePosition & "]='" & RS("Description") & "';"&vbCRLF
		Else
		'End Add
			If Val=Sel Then 'Micheal SWIFT # 2401608 Start
				sTemp="SELECTED" 
				strSelectedLossType=RS("Description")
			Else 
				sTemp=""
				strSelectedLossType=""
			End If 'Micheal SWIFT # 2401608 END
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& RS("Description") &vbCRLF
		' Start Add by Micheal for SWIFT # 2401608
		End If
		'End Add
		RS.MoveNext
	Wend
	' Start Add by Micheal for SWIFT # 2401608
	If Not isScript Then
	' End Add
		Str=Str&"</select>"	&vbCRLF
	' Start Add by Micheal for SWIFT # 2401608
	End If	
	' End Add
	RS.Close
	Set RS=Nothing
	getLossType=Str
End Function
Function getSegment(Name,Sel,cn, isScript)
'****************************************************************************************
'1. Function/Procedure Name          : getSegment
'2. Description           	         : Fetch business segments , this drop box is added for NPT
'3. Calling Forms:   	             : LOSSTIME.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   5-Aug-2009			    Nilesh Naik	        	   Added for SWIFT # 2401608

'****************************************************************************************

Dim Str,SQL,RS,val,sTemp,cmdProductline

        cn.CursorLocation = 3
		Set RS = Server.CreateObject("ADODB.RecordSet")
		Set cmdProductline = Server.CreateObject("ADODB.Command")
	
		With cmdProductline
			.ActiveConnection = cn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_fetchtabledata"
			
			.Parameters.Append .CreateParameter("@Name", adVarChar,1,100  ,"tblproductlines")
			.Parameters.Append .CreateParameter("@Where", adVarChar,1,100 ,"")
			.Parameters.Append .CreateParameter("@Orderby", adVarChar,1,100 ,"ProductLine")
			
			Set RS = .Execute()
		End With
			
	'''SQL = "select PLID,ProductLine  from tblproductlines ORDER BY ProductLine"
	'''Set RS=cn.execute(SQL)
	' Start Add by Micheal for SWIFT # 2401608
	If isScript Then
		'Response.Write "<Script Language=JavaScript>"&vbCRLF
		Response.Write "var arrSegmentId = new Array(" & RS.RecordCount & ");"&vbCRLF
		Response.Write "var arrSegmentDesc = new Array(" & RS.RecordCount & ");"&vbCRLF
		Response.Write "arrSegmentId[0]=0;"&vbCRLF
		'Micheal SWIFT # 2401608 Start [Changed (Select One) to (Not Applicable)]
		Response.Write "arrSegmentDesc[0]='(Not Applicable)';"&vbCRLF
		'Micheal SWIFT # 2401608 End
	Else	
	' End Add
		Str="<select Id='" &Name&"' name='"&Name&"' " & IIF((strSelectedLossType<>"SLB"), "Disabled" , "") & " >"&vbCRLF
		'Response.Write "arrSegmentDesc[0]='(Select One)';"&vbCRLF
		'Micheal SWIFT # 2401608 Start - 06-Oct-2009
		Str=Str&"	<option value=''>(Not Applicable)"&vbCRLF
		'Micheal SWIFT # 2401608 End - 06-Oct-2009
	' Start Add by Micheal for SWIFT # 2401608
	End If
	' End Add
		
	While Not RS.EOF 									
		Val = parseint((RS("PLID")))
		' Start Add by Micheal for SWIFT # 2401608
		If isScript Then
			Response.Write "arrSegmentId[" & RS.AbsolutePosition & "]=" & val & ";"&vbCRLF
			Response.Write "arrSegmentDesc[" & RS.AbsolutePosition & "]='" & SafeDisplay(RS("ProductLine")) & "';"&vbCRLF
		Else
		'End Add
			if Val=Sel then sTemp="SELECTED" else sTemp=""
			Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& SafeDisplay(RS("ProductLine")) &vbCRLF
		'Start Add by Micheal for SWIFT # 2401608
		End If
		' End Add
		RS.MoveNext
	Wend
	' Start Add by Micheal for SWIFT # 2401608
	If Not isScript Then
	'End Add
		Str=Str&"</select>"	&vbCRLF
	' Start Add by Micheal for SWIFT # 2401608
	End If
	'End Add
	RS.Close
	Set RS=Nothing
	getSegment=Str
End Function

Function defaultsegment()
'****************************************************************************************
'1. Function/Procedure Name          : defaultsegment
'2. Description           	         : Fetch Defualt business segments 
'3. Calling Forms:   	             : LOSSTIME.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   5-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT # 2401608

'****************************************************************************************
   Dim strSeg,rsSeg,plid,cmdLocation
    On Error Resume Next

	    cn.CursorLocation = 3
		Set rsSeg = Server.CreateObject("ADODB.RecordSet")
		Set cmdLocation= Server.CreateObject("ADODB.Command")
	
		With cmdLocation
			.ActiveConnection = cn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_fetchtabledata"
			
			.Parameters.Append .CreateParameter("@Name", adVarChar,1,100  ,"tblqt_questtree")
			.Parameters.Append .CreateParameter("@Where", adVarChar,1,100 ,"id  = "& SafeNum(lOrgNo)& "")
			.Parameters.Append .CreateParameter("@Orderby", adVarChar,1,100 ,"name")
			
			Set rsSeg = .Execute()
		End With
		
	If (rsSeg.EOF or rsSeg.BOF) then 
		plid = 0 
    else
		plid =   IIF(rsSeg("plid")="",0,rsSeg("plid"))''rsSeg("plid")
    End if
	
	rsSeg.Close
	Set rsSeg=Nothing
	defaultsegment=plid
	
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime.asp",err.Description
End If
End function

Function getNPT(QPID,cn)
'****************************************************************************************
'1. Function/Procedure Name          : getNPT
'2. Description           	         : Fetch NPT value, new field added for NPT 
'3. Calling Forms:   	             : LOSSTIME.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   5-Aug-2009			    Nilesh Naik	        	   Added for SWIFT # 2401608

'****************************************************************************************
Dim strSQL,rsnpt,intval,sTemp,cmdNPT
On Error Resume Next
	
	   cn.CursorLocation = 3
		Set rsnpt = Server.CreateObject("ADODB.RecordSet")
		Set cmdNPT= Server.CreateObject("ADODB.Command")
	
		With cmdNPT
			.ActiveConnection = cn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_fetchNPTvalues"
			
			.Parameters.Append .CreateParameter("@QID", adInteger, adParamInput, ,SafeNum(QPID))
			.Parameters.Append .CreateParameter("@case", adInteger, adParamInput, ,1)
						
			Set rsnpt = .Execute()
		End With
		
		
	If (rsnpt.EOF or rsnpt.BOF) then 
		intval = 0 
    else
		intval = IIF(Trim(rsnpt("npt")) & "" ="", 0, Trim(rsnpt("npt")))
    End if
	
	rsnpt.Close
	Set rsnpt=Nothing
	getNPT=intval
	
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime.asp",err.Description
End If
End Function

Function fetchcostid(strLoss)
'****************************************************************************************
'1. Function/Procedure Name          : fetchcostid
'2. Description           	         : Fetch cost id ,name of text boxes conatins cost id 
'3. Calling Forms:   	             : LOSSTIME.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   5-Aug-2009			    Nilesh Naik	        	   Added for SWIFT # 2401608 

'****************************************************************************************
Dim strSQL,rsCostid,getID,cmdCostid
On Error Resume Next
	strSQL = " SELECT C.ID as ID FROM tlkpLossCostCategories C with (NOLOCK) INNER JOIN tlkpLossCategories L with (NOLOCK) "
	strSQL = strSQL + "  ON C.LossCatID = L.ID WHERE  (C.LossCatID = 7 AND C.Description like '%"&strLoss& "%')"
	
	 Set rsCostid=cn.execute(strSQL)
		
	If (rsCostid.EOF or rsCostid.BOF) then 
		getID = 0 
    else
		getID = Trim(rsCostid("ID"))
    End if
	
	rsCostid.Close
	Set rsCostid=Nothing
	fetchcostid=getID
	
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime.asp",err.Description
End If
 

End Function

%>

<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	<%Call GetMatrixData %>
	<script ID="clientEventHandlersJS" LANGUAGE="javascript">
	
	

<%=getLossType("cmbType","",cn,0,true)%>
<%=getSegment("cmbsegmentType","",cn,true)%>
<%=fncGetLossDesc("cmbLossDesc","","",cn,true,"TimeLoss")%>
//Micheal SWIFT # 2401608 06-Oct-2009 - Start (Function for checking which Loss Type is selected in each row)
function checkLossType(lossTypeIndex)
{
	var objLossType = document.getElementById('cmbType' + lossTypeIndex);
	var objSegment = document.getElementById('cmbsegmentType' + lossTypeIndex);

	if(objLossType.options[objLossType.selectedIndex].text=='SLB')
	{
		var DefaultSegment = parseInt(document.getElementById('txtDefaultSegment').value,10);

		//Code Added by Shailesh 10/07/2009
		var i=0;
		for(var i=0; i<=objSegment.length;i++)
		{
			if (objSegment[i].value == DefaultSegment)
			{
			    objSegment.selectedIndex = i;
				break;
			}	
		}
		//Code Added by Shailesh 10/07/2009
		
		//objSegment.selectedIndex = DefaultSegment; --Commented by Shailesh 10/07/2009
		objSegment.disabled=false;
	}
	else
	{
		objSegment.selectedIndex = 0
		objSegment.disabled=true;
	}
}
//Micheal SWIFT # 2401608 06-Oct-2009 - End
function addrow()
    {
		var DefaultSegment = parseInt(document.getElementById('txtDefaultSegment').value,10);

        var allRowCount = parseInt(document.getElementById('txtRows').value,10);
        allRowCount++;
        document.getElementById('txtRows').value = allRowCount;
        
        var newRowNum = document.getElementById('lossgrid').rows.length;
		newRowNum;

		var newRow = document.getElementById('lossgrid').insertRow(newRowNum);
		
		newRow.vAlign='top';
		newRow.align='center'
		//arrLossTypeId
		var col0HTML ="&nbsp;";  //"<IMG SRC ='../images/DeleteImage.gif' BORDER=0 >";
		var col1HTML = allRowCount;

		var col2HTML = "<Select name='cmbType" + allRowCount + "' id='cmbType" + allRowCount + "' onblur='calulatemoney(" + allRowCount + ");' onChange='checkLossType(" + allRowCount +");'>" //Micheal SWIFT # 2401608 06-Oct-2009

		for(var i=0; i<=arrLossTypeId.length-1;i++)
		{
			col2HTML = col2HTML + "<OPTION Value=" + arrLossTypeId[i] + ">" + arrLossTypeDesc[i];
		}
		var col2HTML = col2HTML + "</Select>";
		

		var col3HTML = "<Select name='cmbsegmentType" + allRowCount + "' id='cmbsegmentType" + allRowCount + "' Disabled>"; //Micheal SWIFT # 2401608 06-Oct-2009
		for(var i=0; i<=arrSegmentId.length-1;i++)
		{
			/*if(parseInt(arrSegmentId[i],10)==DefaultSegment) //Micheal SWIFT # 2401608 06-Oct-2009
			{
				col3HTML = col3HTML + "<OPTION Value=" + arrSegmentId[i] + " SELECTED>" + arrSegmentDesc[i];
			}
			else
			{*/ //Micheal SWIFT # 2401608 06-Oct-2009
				col3HTML = col3HTML + "<OPTION Value=" + arrSegmentId[i] + ">" + arrSegmentDesc[i];
			//}//Micheal SWIFT # 2401608 06-Oct-2009
		}
		var col3HTML = col3HTML + "</Select>";
		
		var col4HTML = "<Select name='cmbLossDesc" + allRowCount + "' id='cmbLossDesc" + allRowCount + "'>"
		for(var i=0; i<=arrLossDescId.length-1;i++)
		{
			col4HTML = col4HTML + "<OPTION Value=" + arrLossDescId[i] + ">" + arrLossDesc[i];
		}
		var col4HTML = col4HTML + "</Select>";
		
		var col5HTML = "<Input Name='txtQty" + allRowCount + "' Size=2 maxlength='5' onchange='calculateRedmoney(" + allRowCount + ")'>";
		var col6HTML = "hours"
		var col7HTML = "<Input Name='txtunitcost" + allRowCount + "' Size=2  maxlength=6 onchange='calculateRedmoney(" + allRowCount + ")'>";
		var col8HTML = "K$";
		var col9HTML = "<Input Name='txtRedMoney" + allRowCount + "' Size=5  maxlength=8 onchange='calulatemoney(" + allRowCount + ")'>";
		var col10HTML = "K$";
		
		var col0=newRow.insertCell(0);
		var col1=newRow.insertCell(1);
		var col2=newRow.insertCell(2);
		var col3=newRow.insertCell(3);
		var col4=newRow.insertCell(4);
		var col5=newRow.insertCell(5);
		var col6=newRow.insertCell(6);
		var col7=newRow.insertCell(7);
		var col8=newRow.insertCell(8);	
		var col9=newRow.insertCell(9);
		var col10=newRow.insertCell(10);	
        
        col1.align='right';		
        col0.innerHTML=col0HTML;
		col1.innerHTML=col1HTML;
		col2.innerHTML=col2HTML;
		col3.innerHTML=col3HTML;
		col4.innerHTML=col4HTML;
		col5.innerHTML=col5HTML;
		col6.innerHTML=col6HTML;
		col6.align='center';
		col7.innerHTML=col7HTML;
		col8.innerHTML=col8HTML;
		col8.align='center';
		col9.innerHTML=col9HTML;
		col10.innerHTML=col10HTML;
		col10.align='center';
    }
    //Micheal Add
    //**********************************************************************************************************************************
    // Code added for NPT <<2401608>>,all required validation for NPT as well as validation already exists in LossTime2.asp
    // **********************************************************************************************************************************

    function validation()
    {

        var valueoNPT = eval ('window.document.frmTime.hdnNPT.value');
		
        var AgreeNPTval =	eval ('window.document.frmTime.txtAgreenpt.value');
	   
	    var valueNPT = eval ('window.document.frmTime.hdnAgreeNPT.value');
		//var valueNPT = document.getElementById("hdnAgreeNPT").value ;
		
		var valueNPTspan = document.getElementById('calcNPT').innerText 
	
		var valueNPTSQPQ = eval ('window.document.frmTime.txtnpt.value');
		
		var	intSvrty=  eval ('window.document.frmTime.severity.value') ;
	   	
		if(intSvrty==1 && isNaN(parseFloat(valueNPTSQPQ)))
		{
			alert('Value of SQ/PQ NPT must be greater than 0 or Equal to 0 ');
			return false ;
		}
		
	   	if (parseFloat(valueNPTSQPQ) <= 0  && intSvrty != 1) //Shailesh  11-Oct-2009 'Added IsNumericVal Check for 0.12 value NPT Entry
		{
			alert('Value of SQ/PQ NPT must be greater than 0 ');
			return false ;
		}
		
		if ((parseFloat(AgreeNPTval,10)!=0) && (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1))
		{
			alert('Under Review is non-zero then Cause of Review can not be "No Review Required');
			return false ;
		}
	   
        if (parseFloat(valueoNPT) == 0 )
        {
            nptvalidation();
            return ;
        }
     
        
		var errormessage ='' ;
		var errorheader ='' ;
				       
		errorheader += '___________________________________________________\n\n';
		errorheader += 'The Time loss form was not saved because of the following error(s).\n';
		errorheader += '____________________________________________________\n\n';
							
		    
		if (!(IsNumericval(valueNPTSQPQ,1)))
		{	 			            
			errormessage += 'NPT must be numeric value \n'		           
		}
		
		//if (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value!=1)
		//{
		
		if((document.getElementById("optagreeNPTconfirmed").checked==true) && (AgreeNPTval >0))
		{
	
		errormessage += 'Under Review NPT should be Zero when NPT is confirmed \n'	
		}
		//}
	
		
		//if ((valueNPT < 1))  //Shailesh 11-Oct-2009 'Commented
		/*
		if (parseFloat(valueNPT) <= 0) //Shailesh 11-Oct-2009 'Added IsNumericVal Check for 0.12 value NPT Entry
		{
			errormessage += 'Value of NPT must be greater than 0 \n'
		}
		*/
		//start added for SWIFT # 2434444 -to check NPT value upto 2 decimal
		if (errormessage == '')
		{
		    if (!(checkDecimals(valueNPTSQPQ)))
		    {	 			            
			    errormessage += 'SQ/PQ NPT must not be more than 2 decimal places \n'		           
		    }
		}
		
		
			//if (!(IsNumericval(AgreeNPTval,1)))
		//{	 			            
		//	errormessage += 'Under Review NPT must be numeric value \n'		           
		//}
		
		if ((AgreeNPTval!=0) && (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1))
		{
		errormessage += 'Under Review is non-zero then Cause of Review can not be "No Review Required"\n'
		}
		
		if (AgreeNPTval=="")
		{
		errormessage += 'Under review NPT  must be numeric value\n'
		}
		
		
		if (parseFloat(valueNPTSQPQ,10) < parseFloat(AgreeNPTval,10))
		{

		errormessage += 'Under review NPT  should be less then SQ/PQ NPT\n'
		
		}
		
		
		//end added for SWIFT # 2434444 -to check NPT value upto 2 decimal
		
        //changes done as per new requirements of NPT , if NPT value is changed then delete 
          if (parseFloat(valueoNPT,10) != parseFloat(valueNPT,10))
        {
            if (errormessage !='')
            {
                alert(errorheader += errormessage);
                return false;
            }
		
            var strMes = 'You have elected to update the NPT for this SQ event. \n'
                strMes +='Press OK to Continue. \n' 
                
            
            if (confirm(strMes))
            {
                var	intSvrty=  eval ('window.document.frmTime.severity.value') ;
                var intNPT = ScanNPT(valueNPT,matrix)
                
                if(intNPT > parseInt(intSvrty,10))
                {
					<%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
						{
							answer = severityAlert(valueNPTspan,0,intSvrty,3);
                			return;
                		}
                	//Shailesh 30-Oct-2009 Swift# 2438856 	
                	<%else%>
                		{
							window.document.frmTime.action = 'LOSSTime2.asp?<%=sKey%>'
							window.document.frmTime.submit();
							return;                	
						}	
                	<%end if%>	
                	//Shailesh 30-Oct-2009 Swift# 2438856 	                
                }
                else
                {   
                    window.document.frmTime.action = 'LOSSTime2.asp?<%=sKey%>'
                    window.document.frmTime.submit();
                    return;
                }
    
                return;
            }
            else
            {
                return;
            }

        }
        
        //
		var tempSum, tempTotalSum ,tempMax;
		tempSum =0;
		tempTotalSum =0 ;
		tempMax=0 ;
		var iRows1 = '<%=iRows%>' ;
		var iRows1 = eval ('window.document.frmTime.txtRows.value') ;
		/*	 	 
		for(var i=1; i<=iRows1; i++) 
		{		        
			var tempSum = eval ('window.document.frmTime.txtQty'+i+'.value') ;
					      
			if (tempSum != "")
			{
				tempTotalSum = tempTotalSum + parseFloat(tempSum);
				if (tempMax < (parseFloat(tempSum)))  //changed from parseInt  to parseFloat to fix NPT bugs
				{
					tempMax =  parseFloat(tempSum) ;
				} 
			} 
		}    
      
		if (valueNPT < tempMax )
		{
			errormessage += 'You have entered a Loss Time that is greater than your NPT. Please correct the data to continue. \n'
		}

		if (valueNPT >  tempTotalSum )
		{
			errormessage += 'The sum of all the Loss Times entered is less than the NPT. Please correct the data to continue. \n'
		}
       */
		for(var i=1; i<=iRows1; i++) 
		{
			var tempLen =0 ; 
			var tempDesc = eval ('window.document.frmTime.cmbLossDesc'+i+'.value') ;
			if (tempDesc != '0')
			{   
				tempLen = tempLen  +tempDesc.length; 
			} 
			            	
			var tempSum = eval ('window.document.frmTime.txtQty'+i+'.value') ;
			if (tempSum !='')
			{ 
				if (!(IsNumericval(tempSum,1)))
				{
					errormessage += 'Qty in Item '+i+' must be numeric \n'
				}
				else
				{
					tempLen = tempLen + tempSum.length; 
				}
			} 
			var tempSum = eval ('window.document.frmTime.txtunitcost'+i+'.value') ;
			if (tempSum !='')
			{ 
				if (!(IsNumericval(tempSum,1)))
				{
					errormessage += 'Unit Cost/Hour in  Item '+i+' must be numeric \n'
				}
				else
				{
					tempLen = tempLen + tempSum.length; 
				}
			} 
			var tempSum = eval('window.document.frmTime.txtRedMoney'+i+'.value') ;
			if (tempSum !='')
			{ 
				if (!(IsNumericval(tempSum,1)))
				{
					errormessage += 'Red Money in Item '+i+' must be numeric \n'
				}
				else
				{
					tempLen = tempLen  +tempSum.length; 
				}			                
			} 
			var tempSum = eval ('window.document.frmTime.cmbType'+i+'.value') ;
			if ((tempLen > 0) && (tempSum ==''))
			{      
				errormessage += 'No type was selected in Item '+i+' \n'
				
			} 
			//if ((tempLen > 0) && (tempDesc ==''))
			if (((tempLen > 0) && (tempDesc =='0')) ||( (tempLen == 0) && (tempSum.length > 1 )))
			{
				errormessage += 'No Description was selected in Item '+i+' \n'
			} 
			if (tempDesc =='-1')
			{
				errormessage += '"Other (Legacy)" is not a Valid Description. Please change the Description in Item '+i+' \n'
			} 
			//new validation red money must be completed
			if ( (tempSum != '') &&(tempDesc != '0') )
			{
			    var tmpRedm =  eval('window.document.frmTime.txtRedMoney'+i+'.value') ;
				if (tmpRedm =='')
				{
				    errormessage += 'Red Money must be completed in Item '+i+'. Zero is an acceptable value. \n'
				}
			}
		}
        
		var strtempid =0;
		var myArray = new Array();
		var boolflag = new Array();

		myArray[0] = "SLB";
		myArray[1] = "client";
		myArray[2] = "3rdparty";
		          
		boolflag[0]= false;
		boolflag[1]= false;
		boolflag[2]= false;
          
		for(var i=1; i<=iRows1; i++) 
		{
			strname = "cmbType"+i ;
			var df = eval ('window.document.frmTime.' + strname + '') ;
			var txt=df.options[df.selectedIndex].text;
			if (txt !='')
			{
				if (txt.toLowerCase()=="slb")
				{
					boolflag[0]= true;
				}           
				else if (txt.toLowerCase()=="client")
				{
					boolflag[1]= true;
				}                    
				else if (txt.toLowerCase()=="3rd party")
				{
					boolflag[2]= true;
				}
			}
		} 
		               
		if (!( boolflag[0]))
		{
			var blnlosscond = eval('window.document.frmTime.lossg2.value') ;
			if(parseInt(blnlosscond,10) == 1)
			{
				errormessage += 'Data not entered for Loss Type SLB \n'
			}
		}
		if (!( boolflag[1]))
		{
		   
			var blnlosscond = eval ('window.document.frmTime.lossg1.value') ;

			if(parseInt(blnlosscond,10) == 1)
			{
				errormessage += 'Data not entered for Loss Type Client \n'
			}
		}
		if (!( boolflag[2]))
		{
			var blnlosscond = eval ('window.document.frmTime.lossg3.value') ;
			if(parseInt(blnlosscond,10) == 1)
			{             
				errormessage += 'Data not entered for Loss Type 3rd Party \n'
			}
		}
                              
		
           
		var intNPT = valueNPT; 
		var intLoss  ; 
		var intSvrty = 1;
		var strtemp ;
		var id1 ;
		var totalcost; 
		var answer ='' ;
		    
		intSvrty=  eval ('window.document.frmTime.severity.value') ;
		id1='<%=fetchcostid("client")%>';  //we can hard code value here , for client it is 30
		strtemp = "window.document.frmTime.CC_" +id1 +  "" ;
			       
		var txtclientcost= eval(strtemp);
		           
		id1='<%=fetchcostid("slb")%>';  //we can hard code value here , for slb it is 33
		strtemp = "window.document.frmTime.CC_" +id1 +  "" ;
		var txtslbcost= eval(strtemp);
		            
		id1='<%=fetchcostid("Remediation")%>';  //we can hard code value here , for Remediation it is 32
		strtemp = "window.document.frmTime.CC_" +id1 +  "" ;
		var txtremedcost= eval(strtemp);
		            
		id1='<%=fetchcostid("Litigation")%>';  //we can hard code value here , for Litigation it is 31
		strtemp = "window.document.frmTime.CC_" +id1 +  "" ;
		var txtlitcost= eval(strtemp);
				
        if (!(IsNumericval(txtremedcost.value,1)))
        {
            // to fix npt bugs 
            if ((txtremedcost.value).length != 0)
            {
                errormessage += 'Remediation value must be numeric \n'
            }   
        }
        if (!(IsNumericval(txtlitcost.value,1)))
        {
             // to fix npt bugs
            if ((txtlitcost.value).length != 0)
            {
                errormessage += 'Litigation/Other value must be numeric \n'
            }    
        }
        
		if (errormessage !='')
		{
			alert(errorheader += errormessage);
			return false;
		}
		if(txtclientcost.value==""){txtclientcost.value=0;}
		if(txtslbcost.value==""){txtslbcost.value=0;}
		if(txtremedcost.value==""){txtremedcost.value=0;}
		if(txtlitcost.value==""){txtlitcost.value=0;}
		
		totalcost = parseFloat(txtclientcost.value) + parseFloat(txtslbcost.value) + parseFloat(txtremedcost.value) + parseFloat(txtlitcost.value) ;
		intLoss = totalcost ;
		         		
		intNPT = ScanNPT(intNPT,matrix)
		intLoss = ScanLoss(intLoss,matrix)
		    		
		if(intNPT > parseInt(intSvrty,10) || intLoss > parseInt(intSvrty,10))
		{
			<%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
			
				answer = severityAlert(valueNPTspan,totalcost,intSvrty,1);
			<%else%>				
				frmTime.submit ();
			<%end if%> //Shailesh 30-Oct-2009 Swift# 2438856			
		}
		else
		{
			frmTime.submit ();
		}
	}
    
    //*****new functions for npt 
    
    //***************************************************************************************************
    // Javascript function added for NPT <<2401608>>,to open a new severity Window
    //***************************************************************************************************	
    function severityAlert(intNPT,totalcost,intSvrty,intCond) 
	{
        var ParmA ;
        var ParmB ;
        var returnA ;
        var returnB;
          
        var MyArgs = new Array(ParmA);
		window.document.frmTime.intmess.value = '1';
        var WinSettings = "center:yes;resizable:yes;dialogHeight:700px;dialogWidth:600px;resizable:0;scrollbars=1"
		
        var MyArgs = window.open("RIRTimeLossMatrix.asp?cond="+intCond+"&NPTValue= "+intNPT+"&DolLoss="+totalcost+"&RIRSeverity="+intSvrty+"",MyArgs, "modal,toolbar=false,center:yes;resizable:yes,dialogHeight:700px,dialogWidth:600px,resizable:1,scrollbars=1");
        MyArgs.focus();
	}
	
    //***************************************************************************************************
    // Javascript function added for NPT <<2401608>>,to to check npt value as per severity matrix
    //***************************************************************************************************
    
    function ScanNPT(NPT,matrix)
	{
		<%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
		{
			for(i=0;i<=4;i++)
			{
			    if(matrix[i][4]==0)matrix[i][4]=9999999999
			    if(NPT >= parseFloat(matrix[i][3]) && NPT < parseFloat(matrix[i][4]))
			    {
			        if (matrix[i][0] ==5)
			        {
						var valueoNPT = eval ('window.document.frmTime.hdnNPT.value');
						var valueoLoss = eval ('window.document.frmTime.HdnLoss.value');
						if(parseFloat(valueoNPT) >= parseFloat(matrix[i][3]) || parseFloat(valueoLoss) >= parseFloat(matrix[i][1])) 
						{
							return 4 ; //catastrpohic and multi catastrpohic should be same 
						}
					}
			        return matrix[i][0];
			    }
			}
		}
		//Shailesh 30-Oct-2009 Swift# 2438856 
		<%else%> 
			return NPT;
		<%end if%>	
		//Shailesh 30-Oct-2009 Swift# 2438856 
	}
	
    //***************************************************************************************************
    // Javascript function "ScanLoss" added to to check Losses as per severity matrix for NPT SWIFT # 2401608
    //***************************************************************************************************	
	function ScanLoss(Loss,matrix)
	{
		<%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
		{	
			for(i=0;i<=4;i++)
			{
			    if(matrix[i][2]==0)matrix[i][2]=9999999999
			    if(Loss >= parseFloat(matrix[i][1]) && Loss < parseFloat(matrix[i][2]))
			    {
			        if (matrix[i][0] ==5) 
			        {
						var valueoLoss = eval ('window.document.frmTime.HdnLoss.value');
						var valueoNPT = eval ('window.document.frmTime.hdnNPT.value');
						if(parseFloat(valueoNPT) >= parseFloat(matrix[i][3]) || parseFloat(valueoLoss) >= parseFloat(matrix[i][1])) 
						{
							return 4 ; //catastrpohic and multi catastrpohic should be same 
						}
			        }
			        return matrix[i][0];
			    }
			}
		}	
		//Shailesh 30-Oct-2009 Swift# 2438856 
		<%else%>
			return Loss;
		<%end if%>		   
		//Shailesh 30-Oct-2009 Swift# 2438856 
	}
    //*****************************************************************************************************
	function cmdDelete_onclick()
	{
		var bConfirm = window.confirm('Are you sure you wish to DELETE this record');
		return (bConfirm) 
	}
	
    //***************************************************************************
    // Javascript function "IsNumericalval" to check number validation  for NPT SWIFT # 2401608
    //***************************************************************************
	function IsNumericval(strString,strcase)
	{
		var strValidChars ;
		var strChar;
		var blnResult = true;
		if (strcase == 1 )
		{
			strValidChars = "0123456789." ;
		}
		else
		{
			strValidChars = "0123456789" ;
		}
		if (strString.length == 0) return false;
		//start add to fix "."  char bug
		if (strcase =1)
		{
		    if (strString =="."  ||  (strString.lastIndexOf(".") !=strString.indexOf("." )))return false;       
		}
		       
		for (i = 0; i < strString.length && blnResult == true; i++)
		{
			strChar = strString.charAt(i);
			if (strValidChars.indexOf(strChar) == -1)
			{
				blnResult = false;
			}
		}
		return blnResult ;
	}
	//***************************************************************************************************
    // Added Javascript function "keylock" to disable slb and client cost for NPT SWIFT # 2401608
    //***************************************************************************************************	
	function keylock(ev)
	{
	alert(ev.keyCode);
	    if(window.event) // IE
        {
            ev.keyCode = 0;
        }
        else if(ev.which) // Netscape/Firefox/Opera
        {
            return false ;
        }
	}
	 //***************************************************************************************************
    // Added Javascript function "nptvalidation" to  check severity validation
    //***************************************************************************************************	
    function nptvalidation() 
    {
        
		var valueNPT = eval ('window.document.frmTime.hdnAgreeNPT.value');
		var valueNPTspan = document.getElementById('calcNPT').innerText 
		//var valueSQPQ = eval ('window.document.frmTime.txtnpt.value');
		//var valueNPTAgree = eval ('window.document.frmTime.txtAgreenpt.value');
		var	intSvrty=  eval ('window.document.frmTime.severity.value') ;	        
		//start added for SWIFT # 2434444 -to check NPT value upto 2 decimal & chnage the sequence of two functions
		
		//if ((valueNPT < 1)) 'Shailesh 11-Oct-2009 'Commented
		
		/*
		if (parseFloat(valueNPT) <= 0) //Shailesh  11-Oct-2009 'Added IsNumericVal Check for 0.12 value NPT Entry
		{
			alert('Value of NPT must be greater than 0 ');
			return false ;
		}
		*/
		

		
		
		/*
		if (!(IsNumericval(valueNPT,1)))
		{				            
			alert('NPT must be numeric value');	
			return false;	           
		}
	*/
	
	
	
	
		if (!(checkDecimals(valueNPT)))
		{	 			            
			alert('NPT must not be more than 2 decimal places');	
			return false ;	           
		}
		//end added for SWIFT # 2434444 -to check NPT value upto 2 decimal
		
		var intNPT = ScanNPT(valueNPT,matrix)
		
		if(intNPT > parseInt(intSvrty,10))
		{
			<%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
			
				answer = severityAlert(valueNPTspan,0,intSvrty,1);
			<%else%>	
				frmTime.submit ();
			<%end if%>	
		}
		else
		{
			frmTime.submit ();
		}
    
    }
     //***************************************************************************************************
    // Added Javascript function "calculateRedmoney" to caluculate  red money automatically for NPT SWIFT # 2401608
    //***************************************************************************************************	
     	
    function calculateRedmoney (iCntr1) 
	{  
        var strname
        var iRows1 = eval ('window.document.frmTime.txtRows.value') ;
        var intUnitCost  = eval ('window.document.frmTime.txtunitcost'+iCntr1+'.value') ;
        var intQty  = eval ('window.document.frmTime.txtQty'+iCntr1+'.value') ;
        var strcost= 'window.document.frmTime.txtRedMoney'+iCntr1+'.value' ;
       
        if ((intQty == ".") || (intUnitCost ==".") )
        {
         return ;
        }
        if (intQty.lastIndexOf(".") != intQty.indexOf("." ))
            return ;
        if (intUnitCost.lastIndexOf(".") !=intUnitCost.indexOf("." ))
            return ;
         	            
        if (IsNumericval(intUnitCost,1) && IsNumericval(intQty,1))
        {   
            var txtcost1 = eval(strcost);
            //eval ('window.document.frmTime.txtRedMoney'+iCntr1).value =Math.round((( parseFloat(intQty)*parseFloat(intUnitCost)))*Math.pow(10,2))/Math.pow(10,2);	
              eval ('window.document.frmTime.txtRedMoney'+iCntr1+'').value =Math.round((( parseFloat(intQty)*parseFloat(intUnitCost)))*Math.pow(10,2))/Math.pow(10,2);	
        }
            
        calulatemoney(0); 
        return ; 	
	}
    //***************************************************************************************************
    // Added Javascript function "calulatemoney" to caluculate client and slb red money automatically for NPT SWIFT # 2401608
    //***************************************************************************************************	
	function calulatemoney(iCntr10) 
	{
	    
		var strname
		var dbltotalClient
		var dbltotalthirdparty
		var dbltotalSchlumberger
		var iRows1 = eval ('window.document.frmTime.txtRows.value') ;
			 
		dbltotalClient =0.00;
		dbltotalthirdparty =0.00;
		dbltotalSchlumberger =0.00;
			 
		var id2='<%=fetchcostid("client")%>';  //we can hard code value here , for client it is 30
		var strtemp2 = "window.document.frmTime.CC_" +id2 +  "" ;
		var txtclientcost2= eval(strtemp2);
		txtclientcost2.value = '0' ;
			  
			        
		var id3='<%=fetchcostid("slb")%>';  //we can hard code value here , for client it is 30
		var strtemp3 = "window.document.frmTime.CC_" +id3 +  "" ;
		var txtclientcost3= eval(strtemp3);
		txtclientcost3.value = '0' ;
		        
		for(var i=1; i<=iRows1; i++) 
		{
			strname = "cmbType"+i ;
			      
			var df = eval ('window.document.frmTime.' + strname + '') ;
			var val=df.options[df.selectedIndex].value; // to fetch the value 
			var txt=df.options[df.selectedIndex].text;  // to fetch the text i.e. SLB, 3 rd party or client
			         
			var val1= eval ('window.document.frmTime.txtRedMoney'+i+'.value');
			if (txt.toLowerCase() =='slb')
			{
				if (val1 != "" )
				if (IsNumericval(val1,1))
				{
					dbltotalSchlumberger = parseFloat(dbltotalSchlumberger) + parseFloat(val1) ;
				}
				else
				{
					alert('Red Money must be numeric value');
				}
		           
				var id='<%=fetchcostid("slb")%>'; //we can hard code value here , for slb it is 33
					       
				var strtemp = "window.document.frmTime.CC_" +id +  "" ;
				var txtslbcost= eval(strtemp);
				txtslbcost.value = '' ;
				txtslbcost.value =   Math.round(dbltotalSchlumberger*Math.pow(10,3))/Math.pow(10,3);
				
		    }
		    else if (txt.toLowerCase() =='client')
		    {
				if (val1 != "" )
				if (IsNumericval(val1,1))
				{
				  dbltotalClient = parseFloat(dbltotalClient) + parseFloat(val1) ;
				}
				else
				{
				  alert('Red Money must be numeric value');
				}
				        
				var id1='<%=fetchcostid("client")%>';  //we can hard code value here , for client it is 30
				var strtemp = "window.document.frmTime.CC_" +id1 +  "" ;
				var txtclientcost= eval(strtemp);
				txtclientcost.value = '' ;
				txtclientcost.value =   Math.round(dbltotalClient*Math.pow(10,3))/Math.pow(10,3);//to round the number for 2 decimal
		    }
		} 
		return ;
	}
	//Micheal SWIFT # 2401608 06-Oct-2009
	function showExcellence() 
	{
			window.open('<%=strExcellencePage%>','Excellence_In_Execution','height=570,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
	}
	
	function showNPTImage() 
	{
			window.open('<%=StrNPTImage%>','NPT_Image','height=550,width=700,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
	}	
	//Micheal SWIFT # 2401608 06-Oct-2009
	//Micheal SWIFT # 2434444 22-Oct-2009 - Start
	function OnkeypressDisagreeNPT(e)
	{
		var AgreeNPT =	eval ('window.document.frmTime.txtAgreenpt.value');
	
		var a =	eval ('window.document.frmTime.txtnpt.value');
		return CalcNPTHours(AgreeNPT,a);
	}
	
function onfrmLoad()
		{
		var AgreeNPT =	eval ('window.document.frmTime.txtAgreenpt.value');


				if (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1)
		{

				document.getElementById('HSEBLOCK24').style.display = 'none';
				document.getElementById('HSEBLOCK25').style.display = 'none';
				document.getElementById('HSEBLOCK26').style.display = 'none';
				
				}
				else
				{

				document.getElementById('HSEBLOCK24').style.display = '';
				document.getElementById('HSEBLOCK25').style.display = '';
				}
				
			 if(document.getElementById("optagreeNPTconfirmed").checked==true)
			   {
			   document.getElementById('HSEBLOCK25').style.display = 'none';
               }			   
			
				
				
				 if(document.getElementById("optagreeNPT").checked==true)
			   {
			   document.getElementById('HSEBLOCK24').style.display = 'none';
               }			   
				
				
				
	   	}
	
	
	function CheckdataNPTconfirmed()
			{
			
			  if(document.getElementById("optagreeNPTconfirmed").checked==true)
			   {
			   document.getElementById("AnyHours").innerHTML ="Any Hours currently Under Review with External/Internal Client (NPT)";
			   document.getElementById("AnyHours").style.color = "black";
			   document.getElementById("AnyHourshh").style.color = "black";
			   document.getElementById('HSEBLOCK25').style.display = 'none';
			   }
			   else
			   {
			   document.getElementById('HSEBLOCK25').style.display = '';
			   }
			
			}
			
		function CheckdataagreeNPT()
			{
			
			
			  if(document.getElementById("optagreeNPT").checked==true)
			   {
			   document.getElementById('HSEBLOCK26').style.display = '';
			   document.getElementById("AnyHours").innerHTML ="Agree to Disagree (NPT)";
			   document.getElementById("AnyHours").style.color = "red";
			   document.getElementById("AnyHourshh").style.color = "red";
			   document.getElementById("NoteID").style.color = "red";
			   document.getElementById("NoAgreement").style.color = "red";
			   document.getElementById('HSEBLOCK24').style.display = 'none';
			   }
			   else
			   {
			   document.getElementById("AnyHours").innerHTML ="Any Hours currently Under Review with External/Internal Client (NPT)";
			   document.getElementById("AnyHours").style.color = "black";
			   document.getElementById("AnyHourshh").style.color = "black";
			   document.getElementById("NoAgreement").style.color = "black";
			   document.getElementById('HSEBLOCK24').style.display = '';
			   document.getElementById('HSEBLOCK26').style.display = 'none';
			   }
			
			}	
			
			
	
	function Onchangecause()
		{
		
		if (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1)
		{
		//document.getElementById('optagreeNPTconfirmed').disabled = true;
		//document.getElementById('optagreeNPT').disabled = true;
		//document.getElementById("optagreeNPT").checked==false
		//document.getElementById("optagreeNPTconfirmed").checked==false
		document.getElementById('HSEBLOCK24').style.display = 'none';
		document.getElementById('HSEBLOCK25').style.display = 'none';
		document.getElementById('HSEBLOCK26').style.display = 'none';
		 document.getElementById("AnyHours").innerHTML ="Any Hours currently Under Review with External/Internal Client (NPT)";
			   document.getElementById("AnyHours").style.color = "black";
			   document.getElementById("AnyHourshh").style.color = "black";
			   document.getElementById("NoAgreement").style.color = "black";
		
		}
		else
		{
		//document.getElementById('optagreeNPTconfirmed').disabled = false;
		//document.getElementById('optagreeNPT').disabled = false;
		
		document.getElementById('HSEBLOCK24').style.display = '';
		document.getElementById('HSEBLOCK25').style.display = '';
		
		}
		
		}
	
	function CalcNPTHours(aNPT,cNPT)
	{
	

		
		//if (aNPT > cNPT)
		//{
		//alert('Original NPT should be gretar then Under review NPT');
		//document.getElementById("calcNPT").innerHTML =0;
		//}
		document.getElementById("calcNPT").innerHTML =cNPT - aNPT;
		document.getElementById("hdnAgreeNPT").value =cNPT - aNPT;
		document.getElementById("NPTH").innerHTML =aNPT;
		

	}
	function disableEnterKey(e)
	{
		var key;

		if(window.event)
		{
			key = window.event.keyCode; //IE
		}else
		{     
			key = e.which; //firefox      
		}		
		return (key != 13);
	}
	//Micheal SWIFT # 2434444 22-Oct-2009 - End
	//start SWIFT # 2434444 30-0ct -NPT value should be upto 2 decimals
    function checkDecimals(fieldValue)
    {   var blnResult =true;
        decallowed = 2; 
        if (!(isNaN(fieldValue)) && (fieldValue != ""))
        {
            if (fieldValue.indexOf('.') == -1) fieldValue += ".";
            dectext = fieldValue.substring(fieldValue.indexOf('.')+1, fieldValue.length);
            if (dectext.length > decallowed)
            {  
               var start = fieldValue.indexOf('.')+3;
               for (var i =1 ; i <= (dectext.length-decallowed) ;i++)
               {   
                    var number =fieldValue.substring(start, start+1);
                    if (number != "0" )
                    {
                        blnResult= false;
                    }
                   start = start +1;
               }
            }
        }    
        return blnResult ;
    }
    //end SWIFT # 2434444 30-0ct -NPT value should be upto 2 decimals
    
    
    function fncToggleQty(rowQty)
    {
 	    var DescField = document.getElementById('cmbLossDesc' + rowQty);
	    var QtyField = document.getElementById('txtQty' + rowQty);
	    var UnitCostField = document.getElementById('txtunitcost' + rowQty);

        var DescVal = DescField.options[DescField.selectedIndex].value;
           
        if (DescVal == 6 || DescVal == 7) //NPT or Rig Time
        {
            QtyField.disabled = false;
            QtyField.className = '';
            UnitCostField.disabled = false;
            UnitCostField.className = '';
        }
        else
        {
            QtyField.value = "";
		    QtyField.disabled = true;
            QtyField.className = 'MinorHeading';
            UnitCostField.value = "";
		    UnitCostField.disabled = true;
            UnitCostField.className = 'MinorHeading';
        }
    }
    </script>
</head>
<%if  LockCountSQ = 0  and  not chkSQLockingMgmt() and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate)) then%>
<body MARGINWIDTH="0" MARGINHEIGHT="0" LEFTMARGIN="0" TOPMARGIN="0" RIGHTMARGIN="0" >
<%else%>
<body MARGINWIDTH="0" MARGINHEIGHT="0" LEFTMARGIN="0" TOPMARGIN="0" RIGHTMARGIN="0" onLoad="onfrmLoad()">
<%end if%>


<script Language="JavaScript" src="../inc/wz_tooltip.js"></script>
<%
displaymenubar(RS1)
RS1.Close

RS1.Open "SELECT Count(*) AS RecCount FROM tblRIRTime with (NOLOCK) WHERE QPID=" & SafeNum(iQPID), cn
''RS1.Open "SELECT t.*, r.npt FROM tblRIRTime  t with (NOLOCK)INNER JOIN tblrirp1 r WITH (NOLOCK)on t.qpid = r.qid WHERE  QPID=" & SafeNum(iQPID), cn
iRows = RS1("RecCount")+1
if iRows<4 then iRows=4 ''changed for NPT <<2401608>> , default row = 4 now instead of 2
RS1.Close
Set RS1=Nothing
if ACLDefined Then DisplayConfidential()



%>
<form name="frmTime" method="post" action="LossTime2.asp<%=sKey%>">
<style>

.LockWarning
{
display:inline-block;
color:#cc0000;
background-color:#ffff99;
font-weight:900;
font-size:14px;
text-align:center;
}

</style>
<%If  not chkSQLockingMgmt() and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate))   Then%>
<table border=0 align=center cellPadding=0 cellSpacing=0 width=100%>
<TR>
		<TD align = center>
		<span id='Warning22' class='LockWarning'><%Response.Write sHSEWarningTextSQ%></span>			
		</TD>
	</TR>
</table>
<%END IF%>
<table border="0" cellPadding="0" cellSpacing="0" width="100%">
	<tr class="reportheading">
		<td align="left" colspan="3">
			Report Date:&nbsp;<%=FmtDate(dtRptDate) & " (UTC)"%>
		</td>
						
		<td align="right" colspan="3" class="field">
				Report Number:&nbsp;
				<span class="urgent">
				<%
				Response.Write "<A href='" & "RIRview.asp" & sKey & "'>" & getreportnumber(dtRptDate)& "</A>"%>													
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
</table>
<input type="hidden" name="intmess" value="<%=Request("intmess")%>">
<input type="hidden" name="txtRows" id="txtRows" value="<%=iRows%>">
<input type="hidden" name="txtReturnVal" id="txtReturnVal">
<input type="hidden" name="txtQPIID" value="<%=iQPID%>">
<input type="hidden" name="txtCriteria">
<input type="hidden" name="hdnNPT" value="<%=displayQuotes(intNPT)%>" />
<input type="hidden" name="hdnAgreeNPT" id="hdnAgreeNPT" maxlength="5"  value="<%=intNPT%>" />
<input type="hidden" name="lossg1" value="<%=IIF(mdblloss_g1,"1","0")%>">
<input type="hidden" name="lossg2" value=" <%=IIF(mdblloss_g2,"1","0")%>">
<input type="hidden" name="lossg3" value="<%=IIF(mdblloss_g3,"1","0")%>">
<input type="hidden" name="severity" value="<%=mintseverity%>">

<%' Start Add by Micheal for SWIFT # 2401608%>
<input Type="Hidden" Value="<%=defaultsegment()%>" Name="txtDefaultSegment" Id="txtDefaultSegment">
<%'End Add%>

<table border="1" align="center" cellPadding="0" cellSpacing="0" width="100%">
	<tr class="reportheading">
	    <td align="center" colspan="8" class="field">
			Non-Productive Time
			
		</td>		
	</tr> 
    <tr>
		<td align="center">
		<a HREF="<%=StrNPTImage%>" Title="NPT" onClick="showNPTImage();;return false;"><b>NPT</b></a><BR>				
		<a HREF="<%=strExcellencePage%>" Title="Display Excellence in Execution" onClick="showExcellence();;return false;"><b>Excellence in Execution</b></a>
		</td>
	</tr>
    <tr id="HSEBLOCK1">
	    <td  align="right"><b>SQ/PQ Non Conformance START Time to SQ/PQ Non Conformance END Time - through consultation with External/Internal Client (NPT) </b>
	       <input type="text" name="txtnpt" size="2" value="<%=displayQuotes(SQPQNPT)%>" maxlength="5"  onchange ="return OnkeypressDisagreeNPT(event);">   <!--onchange ="return OnkeypressClientNPT(event);"-->
	       <%=mSymbol%> Hours<a href="javascript:void(0)" onmouseover="Tip('This is the value of NPT determined through consultation with the Client.</br>Only where there is a disagreement  requiring a Review </br> is it required to enter a &lsquo;Under Review&rsquo; value')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    </td>                
    </tr>  
</table>
	<table border="0" align="center" cellPadding="0" cellSpacing="0" width="100%">
  <tr id="HSEBLOCK21">
	    <td align="right">
		<b><span id="AnyHours">Any Hours currently Under Review with External/Internal Client (NPT) </span></b>
	       <input type="text" name="txtAgreenpt" size="2" value="<%=displayQuotes(intAgreenpt)%>" maxlength="5" onchange ="return OnkeypressDisagreeNPT(event);"> 
	       <%=mSymbol%> <span id="AnyHourshh">Hours</span><a href="javascript:void(0)" onmouseover="Tip('If there is not an agreement between SLB and the Client on the quantity of NPT. Specify the number of hours that is currently &lsquo;Under Review&rsquo;.</br></br>Once an agreement has been reached the Agreed NPT value (above) should be amended and the &lsquo;Under Review&rsquo; value zeroed.</br></br>If an &lsquo;Agree to Disagree&rsquo; situation arises the &lsquo;Under Review&rsquo; value remains.')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    </td>                
  </tr>  

  <tr id="HSEBLOCK22">
	    <td align="right">
		<b>Cause of any Review : </b>
	       <%=getCausereview(txtCausevalue)%><a href="javascript:void(0)" onmouseover="Tip('Select the Cause of any Review.</br>Please do not amend the selection once the agreement</br> has been resolved.')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    </td>                
  </tr>  

  <tr id="HSEBLOCK23" align="center">
	    <td  nowrap><b>NPT Hours in agreement with External/Internal Client = </b><span id="calcNPT"><%=intNPT%></span>&nbsp;Hours<a href="javascript:void(0)" onmouseover="Tip('This is the value exported from this report that</br> contributes to Segment and Corporate NPTr  metrics.')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a></td>                
  </tr> 

  
    <tr id="HSEBLOCK24" >
	    <td align="center"><font ><b>Agreement has been reached with External/Internal Client - NPT is confirmed </b></font> <span> <input type="Checkbox" onClick="CheckdataNPTconfirmed()" name="optagreeNPTconfirmed" id="optagreeNPTconfirmed"  value="1" <%=NPTConfirmed%>  > </span><a href="javascript:void(0)" onmouseover="Tip('If a Review has occurred it is required to confirm that the Review has been completed to close the report.</br>The &lsquo;Consultation NPT&rsquo; value should be amended to the agreed value and the &lsquo;Under Review&rsquo; valued zeroed.')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a> </td>    
   </tr>
  
    <tr id="HSEBLOCK25" >
	    <td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><span id="NoAgreement">No Agreement reached with External/Internal Client (agree to disagree)</span> </b> <span> <input type="Checkbox" onClick="CheckdataagreeNPT()" name="optagreeNPT" id="optagreeNPT"  value="1"  <%=NoAgreement%>> </span> <a href="javascript:void(0)" onmouseover="Tip('If no agreement can be reached &lsquo;Agree to Disagree&rsquo; should be selected. The &lsquo;Consultation NPT&rsquo; and &lsquo;Under  Review&rsquo; values remain unchanged.</br></br>It is now required to create a secondary RIR report to capture the &lsquo;Under Review&rsquo; NPT value. The child RIR should be linked to  this parent Report.')" onmouseout="UnTip()"><img src="../images/qmark.gif" border=0 height=16 width=16 ></a></td>    
   </tr> 
   
     <tr id="HSEBLOCK26" >
	    <td align="center"><b><span id="NoteID">Note:&nbsp;A secondary RIR needs to be created which is linked to this RIR Report</br>&nbsp;&nbsp;&nbsp; that has <span id="NPTH"><%=intAgreenpt%></span> hours of Non-SLB Related NPT attributed to it,to ensure that NPT</br>&nbsp; Totals remains correct.</span> </b>
		
		</td>    
    </tr> 
   


  
</table>
<%If Cdbl(intNPT) = 0 Then%>

    <table width="100%">
    <tr id="HSEBLOCK2">
        <td align="right" valign="top">
    	    <input type="button" name="cmdnptSubmit" value="Save Data" onclick="validation()">
    	    <!--SWIFT # 2434444 Hidden Textbox to check if Total Loss changed from Catastrophic to Multiple Catastrophic-->
    	    <input Type="Hidden" Name="HdnLoss" Id="HdnLoss" Value="0"> 
    	    <!--SWIFT # 2434444 -->
        </td> 	
       </tr>
	   <%	
				Dim valueNPTSQPQ,sTemp1,checkvalue,rsch
			    sTemp1 = ""
				set rsch = Server.CreateObject("ADODB.recordset")
				valueNPTSQPQ = "SELECT sqpqnpt FROM tblrirp1 with (NOLOCK) WHERE QID=" & SafeNum(iQPID) & ""
				rsch.open valueNPTSQPQ,cn
				if not rsch.eof then
					checkvalue = rsch("sqpqnpt")
				end if
				rsch.close
				if(parseFloat(checkvalue)=0.00) then 
				sTemp1 = "Delete from tblRIRTime WHERE QPID=" & SafeNum(iQPID) & ""
				sTemp2 = "Delete from tblRIRCosts WHERE QPID=" & SafeNum(iQPID) & ""
				cn.execute(sTemp1)
				cn.execute(sTemp2)
				cn.close()
				Set cn = nothing
				end if
				set rsch=nothing
		%>
    </table>
<% End if %>

<%If Cdbl(intNPT) <> 0   Then 'Shailesh 11-Oct-2009 'Change Clng To CDbl for 0.12 NPT Entry%>
    <table border="1" cellPadding="0" cellSpacing="0" width="100%">	
	<tr class="reportheading">
	    <td align="center" colspan="8" class="field">Red Money</td>		
	</tr> 
		<tr>
				    <td align="center" colspan="11" class="boxednote"><font color=red><b>Red Money Calculation Sheet</font></b> (ALL COST ESTIMATES ARE IN <font color=red><B>THOUSANDS</B></font> U.S. DOLLARS)</td>
			    </tr>		
			    <tr>
				    <td align="center" colspan="11" class="boxednote" id="styleSmall">
					    To add more items click on &quot;Add New Row&quot; Button.
					    To delete an entry click on the Delete icon.
				    </td>
			    </tr>
	            
    <tr id="HSEBLOCK3">    	
	    <td align="left" valign="top" style="height: 108px">    				
		    <table width="100%" id="lossgrid" border="1" cellPadding="2" cellSpacing="0">
				
			    <tr>
			        <td align="right" style="height: 19px">&nbsp;</td>	
				    <td align="right" style="height: 20px">Item</td>				
				    <td align="center" style="height: 20px">Party incurring the Loss  <span> <%=mSymbol%> </span></td><!-- changed the label to fix NPT bugs -->
				    <td align="center" style="height: 20px">Segment Sustaining the Loss</td>   <!-- changed the label to fix NPT bugs -->
				    <td align="center" style="height: 18px">Description of loss <span> <%=mSymbol%> </span></td>
				    <td align="center" style="height: 20px">Quantity</td>
				    <td align="center" style="height: 20px">Unit</td>
				    <td align="center" style="height: 20px">Unit Cost/Hour</td>
				    <td align="center" style="height: 20px">Unit</td>
				        				
				    <td align="center" style="height: 20px">Red Money $ &nbsp;<b> <%=getHelpLink("Red Money")%> </b>&nbsp; <span> <%=mSymbol%> </span></td>
				    <td align="center" style="height: 20px">Unit</td>
				    </tr>
			    <%	
			
			    sTemp = ""
			     sTemp = "SELECT * FROM tblRIRTime with (NOLOCK) WHERE QPID=" & SafeNum(iQPID) & " ORDER BY Seq"
			    ''sTemp = "SELECT t.*, r.npt FROM tblRIRTime  t with (NOLOCK)  INNER JOIN tblrirp1 r WITH (NOLOCK)on t.qpid = r.qid  WHERE QPID=" & SafeNum(iQPID) & " ORDER BY Seq"
    			LossTypeCtr = 1
			    RS.Open sTemp, cn
    			If RS.EOF or RS.BOF then TLExists = False else TLExists = True
			    For iCntr = 1 to iRows
    			    DefSeg=""
				    bNew = False
				    If RS.EOF or RS.BOF Then bNew = True
					
					%>
			    <tr >
			    <td align="right">
					    <%If not bnew Then 
						   
						    Response.Write "<A href=LossTime2.asp" & sKey & "&Delete=1&Typeid="& RS("type")&"&redmoney="&RS("RedMoney")&"&row="& iCntr &"&ID=" & RS("SEQ") & " onclick='return cmdDelete_onclick()' >  <IMG SRC ='../images/DeleteImage.gif' ALT='Delete' TITLE='Delete' BORDER=0 ></A>"
					    else
						    Response.Write "&nbsp;" 
					    End If
				    %>		
				    </td>
				    <td align="right">
					    <%If not bnew Then 
						
						    Response.Write iCntr
					    else
						    Response.Write iCntr
					    End If
				    %>		
				    </td>
					
				    <%If bNew  Then
				        if Not TLExists then
				            if LossTypeCtr = 1 then 'First Row
				                if mdblloss_g2 = true then  'SLB i.e. 20
				            	    sTemp = "20"
				            	    DefSeg = defaultsegment()
				                elseif mdblloss_g1 = true then 'Client i.e. 19
				                    sTemp = "19"
				                    LossTypeCtr = LossTypeCtr + 1
				                elseif mdblloss_g3 = true then '3rd Party i.e. 21
				                    sTemp = "21"
				                    LossTypeCtr = LossTypeCtr + 2
				                else
				                    sTemp="" 
				                end if
				            elseif LossTypeCtr = 2 then 'Second Row
				                if mdblloss_g1 = true then 'Client
				                    sTemp = "19"
				                elseif mdblloss_g3 = true then '3rd Party
				                    sTemp = "21"
				                    LossTypeCtr = LossTypeCtr + 1
				                else
				                    sTemp="" 
				                end if
				            elseif LossTypeCtr = 3 then 'Third Row
				                if mdblloss_g3 = true then '3rd Party
				                    sTemp = "21"
				                else
				                    sTemp="" 
				                end if
				            end if
				        else
				            sTemp="" 
				        end if
				    else 
				        sTemp=RS("Type")
				    end if%>
				    <td align="center"><%=getLossType("cmbType"&iCntr,sTemp,cn,iCntr,false)%></td>
				    <!-- Code changed  for NPT <<2401608>> , new drop down-->
				     <%If bNew  Then sTemp=DefSeg else sTemp=RS("plid") 'Micheal SWIFT # 2401608 06-Oct-2009%>
				     <td align="center"><%=getSegment("cmbsegmentType"&iCntr,sTemp,cn,false)%></td>
    				 
				     <%If bNew  Then sTemp="" else sTemp=iif(trim(RS("LossDescID"))="" or trim(RS("Description"))<>"",-1,RS("LossDescID"))%>
				     <td align="center"><%=fncGetLossDesc("cmbLossDesc"&iCntr,iCntr,sTemp,cn,false,"TimeLoss")%>
				     <%
				     'sTemp = "<br>" & RS("Description")
				     'if trim(RS("Description")<>"" then response.Write sTemp
				     %>
				     </td>
				    <td align="center">
					    <%
					    LossDescID = iif(sTemp="",0,sTemp)
					    sTemp = ""
					    If bNew = False Then sTemp=RS("Qty")
					    if LossDescID = 6 or LossDescID = 7 then'Rig Time or NPT%>
					        <input type="text" name="txtQty<%=iCntr%>" id="txtQty<%=iCntr%>" size="2" value="<%=sTemp%>" maxlength="5" onchange="calculateRedmoney(<%=iCntr%>)">
					    <%else %>
					        <input type="text" name="txtQty<%=iCntr%>" id="txtQty<%=iCntr%>" size="2" value="" maxlength="5" onchange="calculateRedmoney(<%=iCntr%>)" class="MinorHeading" disabled>
					    <%end if %>
				    </td>
				    <td align="center">  hours</td>	
				    <td align="center">
					    <%sTemp = ""
					    If bNew = False Then sTemp=RS("UnitCost")
					    if LossDescID = 6 or LossDescID = 7 then'Rig Time or NPT%>
					        <input type="text" name="txtunitcost<%=iCntr%>" id="txtunitcost<%=iCntr%>" size="2" value="<%=sTemp%>" maxlength="6" onchange="calculateRedmoney(<%=iCntr%>)">
                        <%else %>
					        <input type="text" name="txtunitcost<%=iCntr%>" id="txtunitcost<%=iCntr%>" size="2" value="" maxlength="6" onchange="calculateRedmoney(<%=iCntr%>)" class="MinorHeading" disabled>
                        <%end if %>
				    </td>
				    <td align="center">K$</td>	
				    <td align="center">
					    <%sTemp = ""
					    If bNew = False Then sTemp=RS("RedMoney")%>
                        <input type="Hidden" name="hiddentxtRedMoney<%=iCntr%>" value="<%=sTemp%>">
					    <input type="text" name="txtRedMoney<%=iCntr%>" size="5" value="<%=displayQuotes(sTemp)%>" maxlength="8" onchange="calulatemoney(<%=iCntr%>)">
				    </td>
				     <td align="center">K$</td>
			    </tr>					
			    <%
				    If not bNew then RS.MoveNext 
				    LossTypeCtr = LossTypeCtr + 1
			    Next
			    RS.Close%>			
		    </table>
		</td>
	</tr>
    <tr id="HSEBLOCK4">
     <td style="width: 22px" align="left"> <input type="button" name="cmdAdd" value="Add New Row" onclick="addrow()"></td>
    </tr>
    <%	
        'start added to fix NPT <<2401608>> Bugs
		 if  LockCountSQ =0   and  not chkSQLockingMgmt() and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate))   Then
    else
    if (Request("intmess") <> "") then 
            DisplayCosttime cn, 7, iQPID ,Request("intmess")  
        else
            DisplayCosttime cn, 7, iQPID,0  
        end if
        'end add
	    ' DisplayCosttime cn, 7, iQPID  ''changed for NPT <<2401608>>
	    cn.Close
	    Set cn=Nothing	
       end if		
    %>
    </table>
	
    <table id="HSEBLOCK5" width="100%">
	    <tr>
		    <td align="right" valign="top">
			   
			    <input type="button" name="cmdSubmit" value="Save Data" onclick="validation()">
		    </td>
	    </tr>			
    </table>
	
<%If LockCountSQ = 0 and   not chkSQLockingMgmt() and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate))   Then%>
<table border=0 align=center cellPadding=0 cellSpacing=0 width=100%>
    <TR>
		<TD align = center>
		<span id="Warning23" class="LockWarning">Use view tab above to display values</span>			
		</TD>
	</TR>
</table>
<%END IF%>
	
<%End If%>

	<%
	 if  LockCountSQ = 0  and  not chkSQLockingMgmt() and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate))   Then

		Response.Write "<script type='text/javascript'>document.getElementById('Warning22').innerHTML = '*** Key SQ Fields in this Report are now Locked ***';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK1').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK2').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK3').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK4').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK5').style.display = 'none';</script>"
		
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK21').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK22').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK23').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK24').style.display = 'none';</script>"
		Response.Write "<script type='text/javascript'>if (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1) {document.getElementById('HSEBLOCK25').style.display = 'none';}</script>"
		Response.Write "<script type='text/javascript'>if (frmTime.selCausereview[frmTime.selCausereview.selectedIndex].value==1) {document.getElementById('HSEBLOCK26').style.display = 'none';}</script>"
    else
	if (CDate(dtRptDatetmp) >= CDate(comparedateEventdate)) then
	    Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('Warning22')) != 'undefined' && document.getElementById('Warning22') != null) {document.getElementById('Warning22').innerHTML = '*** Warning this report will be locked from editing Key SQ Fields in  " & LockCountSQ &" day(s) ***';}</script>"
	end if
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK1')) != 'undefined' && document.getElementById('HSEBLOCK1') != null) {document.getElementById('HSEBLOCK1').style.display = ''; }</script>"
		'Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK10').style.display = '';</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK3')) != 'undefined' && document.getElementById('HSEBLOCK3') != null) {document.getElementById('HSEBLOCK3').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK4')) != 'undefined' && document.getElementById('HSEBLOCK4') != null) {document.getElementById('HSEBLOCK4').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK5')) != 'undefined' && document.getElementById('HSEBLOCK5') != null) {document.getElementById('HSEBLOCK5').style.display = ''; }</script>"

		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK21')) != 'undefined' && document.getElementById('HSEBLOCK21') != null) {document.getElementById('HSEBLOCK21').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK22')) != 'undefined' && document.getElementById('HSEBLOCK22') != null) {document.getElementById('HSEBLOCK22').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK23')) != 'undefined' && document.getElementById('HSEBLOCK23') != null) {document.getElementById('HSEBLOCK23').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK24')) != 'undefined' && document.getElementById('HSEBLOCK24') != null) {document.getElementById('HSEBLOCK24').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>if (typeof(document.getElementById('HSEBLOCK25')) != 'undefined' && document.getElementById('HSEBLOCK25') != null) {document.getElementById('HSEBLOCK25').style.display = ''; }</script>"
		Response.Write "<script type='text/javascript'>var element =  document.getElementById('HSEBLOCK26'); if (typeof(element) == 'undefined' && element == null) {document.getElementById('HSEBLOCK26').style.display = '';}</script>"
		
		
	if NoAgreementval=1  then
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK24').style.display = 'none';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK26').style.display = '';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHours').innerHTML = 'Agree to Disagree (NPT)';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHours').style.color= 'red';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHourshh').style.color= 'red';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('NoAgreement').style.color= 'red';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('NoteID').style.color= 'red';</script>"
	else
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK24').style.display = '';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK26').style.display = 'none';</script>"
	end if
	
	
	if NPTConfirmedval=1  then
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK25').style.display = 'none';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK26').style.display = 'none';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHours').innerHTML = 'Any Hours currently Under Review with External/Internal Client (NPT)';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHours').style.color= 'black';</script>"
	Response.Write "<script type='text/javascript'>document.getElementById('AnyHourshh').style.color= 'black';</script>"
	else
	Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK25').style.display = '';</script>"
	end if
	
	end if
	

	
%>
</form>
</body>
</html>


<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSTime.asp;29 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1255593] 17-AUG-2009 16:18:31 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02" %>
<% '       3*[1260199] 19-AUG-2009 12:59:35 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 - update script" %>
<% '       4*[1261359] 20-AUG-2009 15:24:29 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation" %>
<% '       5*[1262427] 26-AUG-2009 15:56:05 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Formatting issues" %>
<% '       6*[1264780] 28-AUG-2009 15:16:12 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 Changes as per new requirement" %>
<% '       7*[1274615] 22-SEP-2009 21:22:34 (GMT) SKadam3 %>
<% '         "SWIFT #2401608 - NPT - Changes to conform to new SQ Std02" %>
<% '       8*[1274713] 24-SEP-2009 08:13:19 (GMT) SKadam3 %>
<% '         "SWIFT #2401608 - NPT - Changes to conform to new SQ Std02" %>
<% '       9*[1275699] 29-SEP-2009 15:48:26 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      10*[1277927] 02-OCT-2009 17:50:40 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      11*[1277363] 07-OCT-2009 14:50:18 (GMT) MAnthony2 %>
<% '         "SWIFT# 2401608 NPT Change Request" %>
<% '      12*[1280118] 08-OCT-2009 11:30:38 (GMT) SKadam3 %>
<% '         "SWIFT# 2401608 : NPT - Change NPT Image" %>
<% '      13*[1280834] 09-OCT-2009 15:42:30 (GMT) MAnthony2 %>
<% '         "Swift # 2401608 - NPT Changes" %>
<% '      14*[1280948] 12-OCT-2009 13:00:09 (GMT) SKadam3 %>
<% '         "NPT - Hot Fixes" %>
<% '      15*[1282819] 15-OCT-2009 17:07:25 (GMT) MAnthony2 %>
<% '         "SWIFT # 2434444 - NPT Issues - post go live" %>
<% '      16*[1285260] 23-OCT-2009 17:24:27 (GMT) MAnthony2 %>
<% '         "SWIFT # 2401608 - NPT Bugs" %>
<% '      17*[1289068] 30-OCT-2009 11:10:31 (GMT) NNaik %>
<% '         "SWIFT # 2434444 : NPT Issues - post go live" %>
<% '      18*[1289354] 03-NOV-2009 05:23:47 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      19*[1290024] 03-NOV-2009 13:54:13 (GMT) MAnthony2 %>
<% '         "SWIFT # 2434444 - NPT Code Correction" %>
<% '      20*[1290343] 04-NOV-2009 13:41:57 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      21*[1330387] 04-MAR-2010 12:43:40 (GMT) SKadam3 %>
<% '         "SWIFT #2463151 - Remove NPT 'image' from TIME LOSS tab" %>
<% '      22*[1398286] 16-NOV-2010 12:29:24 (GMT) PMakhija %>
<% '         "Swift#2502389-Cross Scripting issue in RIR module except Reports" %>
<% '      23*[1633354] 07-AUG-2012 16:27:13 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      24*[1653769] 07-AUG-2012 22:49:53 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      25*[1654828] 21-AUG-2012 15:53:12 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      26*[1821515] 26-FEB-2014 16:53:18 (GMT) APrakash6 %>
<% '         "Dummy" %>
<% '      27*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      28*[1863826] 09-OCT-2014 14:30:58 (GMT) Rbhalave %>
<% '         "dummy build test" %>
<% '      29*[1878140] 24-JUN-2016 06:56:22 (GMT) VSharma16 %>
<% '         "NFT101068- RIR locking to prevent data integrity issues" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSTime.asp;29 %>
