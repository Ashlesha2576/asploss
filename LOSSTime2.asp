<%@ Language=VBScript %>
<%
'*********************************************************************************************
'1. File Name		              :  Losstime2.asp
'2. Description           	      :  Save Time loss data entry page
'3. Calling Forms   	          : 
'4. Stored Procedures Used        : 
'5. Views Used	   	              : 
'6. Module	   	                  : RIR (HSE/SQ)				
'7. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'    5-Aug-2009			     Nilesh Naik        	 	Modified - changed for NPT SWIFT #2401608
'    25-Sep-2009			 Nilesh Naik        	 	Modified - changed for NPT Bugs SWIFT #2401608
'    30-Oct-2009			 Shailesh					Modified - To Exclude some segment for severity check Swift# 2438856
'    09-Nov-2009			 Shailesh					Modified - Add safenum for NPT field. Swift #2440954
'    11-Nov-2009			 Shailesh					Modified - Add Trim & New Variable fNPT for NPT field. Swift #2440954
'     7-May-2014            Varun Sharma                 Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
'************************************************************************************************
%>

<%option explicit%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<%checktimeout()

Dim sValText, lOrgNo, iCt 
Dim dtRptDate, bCurrPers, sKey
Dim bNotify, iCtr, sValTemp, conn, RS, dtTemp, sHref


Dim sTemp, iTemp, bTemp, sTemp2,iQPID
Dim ACLDefined,MsgID,DelMsg

Dim mdblclient,mdblslb,mdblthirdparty ,lngNpt,lngonpt,blnSave,intMess,txtComments,RSNPTval,FrmNPTval,RSredmoney,frmredmoney  ''added for NPT
dim DisagreeNPT,CauseReview,SQPQNPT,NPTConfirmed,NoAgreement,LegacyNpt,NPTvalue

DisagreeNPT= cDbl(Request.Form("txtAgreenpt"))
CauseReview= Request.Form("selCausereview")
SQPQNPT= cDbl(Request.Form("txtnpt"))  
'NPTvalue= cDbl(Request("hdnAgreeNPT"))
NPTConfirmed=  Request.Form("optagreeNPTconfirmed")
NoAgreement= Request.Form("optagreeNPT")
LegacyNpt= cDbl(Request.Form("txtAgreenpt"))


	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")
	iCt = Request.Form("txtRows")
	intMess = Request("intMess")
	If intMess = "" Then intMess = "0"
	Set conn = GetNewCN()

	'Shailesh 30-Oct-2009 Swift# 2438856
	Dim blnNPT_Exempt
	blnNPT_Exempt = ChkNPT_Exemption()
	'Shailesh 30-Oct-2009

	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", conn)
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, conn)
	If IsNumeric(Request("hdnAgreeNPT")) Then
		lngNpt = cDbl(Request("hdnAgreeNPT"))
	Else
		lngNpt = 0
	End If
	'Check for delete
	If Request("deleteAll")="true" then 
	    'dim intNpt
	    deleteallTime iQPID,lngNpt
	    LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossTime2.asp",139,""
	    DelMsg=server.URLEncode("Entry Successfully Deleted")
	    'Response.Redirect("LossTime.asp" & sKey & "&msg=" &Delmsg )
	    conn.execute("sp_UpdateDateTime " & iQPID & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
	    Response.Redirect("LossTime.asp" & sKey & "&msg=" &Delmsg & "&intmess=" & intMess)
	end if 
	
	If Request.QueryString("Delete")=1 then 
		Set RS = conn.Execute("Select * from tblRIRTime Where QPID="& SafeNum(iQPID) & " and Seq=" & SafeNum(trim(Request.QueryString("ID"))))
		If Not RS.EOF Then
			LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossTime2.asp",139,""	
			''conn.Execute "DELETE FROM tblRIRTime WHERE QPID="& SafeNum(iQPID) & " AND Seq=" & SafeNum(trim(Request.QueryString("ID")))
			'Call deletecost(SafeNum(iQPID),Request("Typeid"),Request("redmoney"),SafeNum(trim(Request.QueryString("ID"))))
			Call deletecost(SafeNum(iQPID),Request("Typeid"), SafeNum(Request("redmoney")),SafeNum(trim(Request.QueryString("ID"))))
			
			DelMsg=server.URLEncode("Entry Successfully Deleted")
			conn.execute("sp_UpdateDateTime " & iQPID & "," & "'" & Trim(session("UserName")) & "','" & Session("UID") & "','R'")
		else
			DelMsg=server.URLEncode("Entry is Not Found")
		End If
		Set RS = Nothing
		conn.close
	    Set conn = Nothing
		'Response.Redirect("LossTime.asp" & sKey & "&msg=" &Delmsg )
		 Response.Redirect("LossTime.asp" & sKey & "&msg=" &Delmsg & "&intmess=" & intMess)
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
			iTemp = Len(Trim(Request.Form("txtRef" & iCtr)))
			iTemp = iTemp + Len(Trim(Request.Form("txtQty" & iCtr)))
			iTemp = iTemp + Len(Trim(Request.Form("txtUnit" & iCtr)))
			
			If iTemp > 0 Then		
				'bNotify = True
				sValText = "No type was selected in row " & iCtr & ".<BR>"
			Else 
				bRowData(iCtr) = False
			End If
		Else
			If Request.Form("cmbLossDesc" & iCtr)=0 then
				'bNotify = True
				sValText = "No Description selected in row " & iCtr & ".<BR>"				
			Else
				bRowData(iCtr) = True
			End If
		End If
		
		'Validate Units (Numeric)
		sTemp = Trim(Request.Form("txtQty" & iCtr))
		If Len(sTemp) > 0 Then
			If Not IsNumeric(sTemp) Then
				'bNotify = True
				sValText = "Qty in row " & iCtr & " must be numeric or blank.<BR>"
			End If
		End If			
	Next
	
''***************************************************************************************************
If IsNumeric(Request.Form("hdnNPT")) Then
	lngonpt = cDbl(Request.Form("hdnNPT"))
Else
	lngonpt = 0
End If

if lngonpt = 0 then 

    UpdateNpt1
    'Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
     Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved") & "&intmess=" & intMess)
end if
if (Request("txtCriteria")="3") then 

    deleteallTime iQPID, lngNpt
	UpdateNpt1    
    'Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
     
    Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved") & "&intmess=" & intMess)
end if 

'WRITE PHASE
		
	If  bNotify = false  Then	
		conn.execute "DELETE FROM tblRIRTime WHERE  QPID="& SafeNum(iQPID) 		
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.LockType = 3	
		MsgID=110
		
		mdblclient = 0 
	    mdblslb = 0 
	    
		For iCtr = 1 to iCt
			
			If bRowData(iCtr) Then
				iTemp = iCtr
				sTemp = "SELECT * FROM tblRIRTime WHERE  QPID="& SafeNum(iQPID) & " AND Seq=" & SafeNum(iTemp)
				RS.Open sTemp, conn			
				If RS.EOF Then 
					RS.AddNew							
					RS("QPID")		= iQPID
					RS("Seq")		= iTemp
				End if					
				RS("RevDate")= Date()			
				RS("Type") = Request.Form("cmbType" & iCtr)
				RS("Description") = ""
				RS("LossDescID") = Request.Form("cmbLossDesc" & iCtr)
				RS("RefNo") = left(Request.Form("txtRef" & iCtr),30)
				'SWIFT - 2434444 Start
				If IsNumeric(Request.Form("txtQty" & iCtr)) Then
					sTemp = cDbl(Request.Form("txtQty" & iCtr))
				Else
					sTemp = 0
				End If
				'SWIFT - 2434444 End
				If trim(sTemp) = "" Then sTemp = Null
				RS("Qty") = sTemp
				RS("Unit") = Request.Form("txtUnit" & iCtr)	
				'start added for fixing the NPT  bugs
				RS("UnitCost") = IIF(Request.Form("txtunitcost" & iCtr)="",0,Request.Form("txtunitcost" & iCtr))
				'end added for fixing the NPT  bugs
				' *****************************************************************
                ' Code changed for NPT <<2401608>>
                ' *****************************************************************
				RS("RedMoney") = IIF(Request.Form("txtRedMoney" & iCtr) & ""="",0,Request.Form("txtRedMoney" & iCtr))
 RSredmoney=Request.Form("hiddentxtRedMoney" & iCtr)
 frmredmoney=Request.Form("txtRedMoney" & iCtr)
 
	   if (RSredmoney <> frmredmoney) and (frmredmoney <> "") and (RSredmoney <> "") and (MsgID=110)  then
		if txtComments = "" then
	txtComments = typename(Request.Form("cmbType" & iCtr))&" Red Money from"&"  "& RSredmoney &"  "&"K$" &"  "& "to" &"  "& frmredmoney&"  "&"K$" 
	else 
	
		txtComments=txtComments&"  "& "<BR>"&typename(Request.Form("cmbType" & iCtr))&" Red Money from"&"  "& RSredmoney &"  "&"K$" &"  "& "to" &"  "& frmredmoney&"  "&"K$"
	end if 
	end if 
			    RS("plid") = IIF(Request.Form("cmbsegmentType" & iCtr)="",0,Request.Form("cmbsegmentType" & iCtr))
			        		      	  
			    if ((instr((typename(Request.Form("cmbType" & iCtr))),"Schlumberger") > 0) and (isnumeric(Request.Form("txtRedMoney" & iCtr))))then
  						mdblslb =  mdblslb + Request.Form("txtRedMoney" & iCtr)
  		        end if    
			    if ((instr((typename(Request.Form("cmbType" & iCtr))),"Client") >0) and (isnumeric(Request.Form("txtRedMoney" & iCtr)))) then    
			            mdblclient= mdblclient + Request.Form("txtRedMoney" & iCtr)	
			    end if 
			  ' *****************************************************************
    
    
				RS.Update					
				RS.Close
			End If
		Next
			
		' Costs
		call UpdateNpt(mdblclient,mdblslb) ''new function added for npt <<2401608>>
		''UpdateCost conn, 7, iQPID, lOrgNo, dtRptDate
		UpdateCostTime  conn, 7, iQPID, lOrgNo, dtRptDate,Formatnumber(mdblslb,2),  Formatnumber(mdblclient ,2) ''function changed for npt <<2401608>>
		UpdateUserInfo iQPID,conn
		
		Set RS = Nothing
		conn.Close
		Set conn = Nothing		
	if (FrmNPTval <> "") and (RSNPTval <> "") then
		if (RSNPTval <> FrmNPTval) and (MsgID=110) then
		if txtComments = "" then
		txtComments = "NPT from"&"  "& RSNPTval &"  "&"to"&"  "& FrmNPTval&"  "&"hours" 
		else
		txtComments=txtComments&"  "& "|| NPT from"&"  "& RSNPTval &"  "&"to"&"  "& FrmNPTval&"  "&"hours" 
        end if 		
		end if 
		end if 
		LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","LossTime2.asp",MsgID,txtComments
		'Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved"))
		 Response.Redirect("LossTime.asp" & sKey & "&msg=" & server.URLEncode("Data Successfully Saved") & "&intmess=" & intMess)
	Else
		conn.Close
		Set conn = Nothing
	End if
	
Sub UpdateNpt (mdblclient,mdblslb)
'****************************************************************************************
'1. Function/Procedure Name          : UpdateNpt
'2. Description           	         : Update NPT,severity if it is changed 
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   12-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608 

'****************************************************************************************
On Error Resume Next
	Dim RSnpt, iRecCount, strSQL,intsev,RSCostCat,intnewsev,inttotalcost,strMessage,conn1
	Set conn1 = GetNewCN()
	inttotalcost = 0 	
	Set RSCostCat=Server.CreateObject("ADODB.Recordset")
	RSCostCat.LockType = 3 
	RSCostCat.Open "SELECT * FROM tlkpLossCostCategories with (NOLOCK) WHERE LossCatID=7",conn 

	Do While Not RSCostCat.EOF
		If Request.Form("CC_" & RSCostCat("ID"))<>"" AND isnumeric(Request.Form("CC_" & RSCostCat("ID")))  Then
			inttotalcost = inttotalcost + Request.Form("CC_" & RSCostCat("ID"))     
		End If
		RSCostCat.MoveNext
	Loop	
	RSCostCat.Close ()
	Set RSCostCat = nothing     
	
	inttotalcost = inttotalcost '+ mdblclient + mdblslb   
	'response.write inttotalcost        
    'response.end
	Set RSnpt=Server.CreateObject("ADODB.Recordset")
	RSnpt.LockType = 3	
	RSnpt.Open "SELECT QID,RptDate,sqseverity,Severity,HSESeverity,SQPQNPT,DisagreeNPT,NoAgreement,NPT,LegacyNpt,CauseReview,NPTConfirmed,OriginalNpt FROM tblrirp1  WHERE qid=" &  SafeNum(iQPID), conn1 
	
	If not RSnpt.EOF then 
		intsev = RSnpt("sqseverity")
	End If
	
	'''call severity and it is true 
	
	If blnNPT_Exempt = 0 Then 'Shailesh 30-Oct-2009 Swift# 2438856 		
		intnewsev = NewSev(Request.Form("hdnAgreeNPT"),inttotalcost,intsev)
	'Shailesh 30-Oct-2009 Swift# 2438856	
	else
	    intnewsev = ""
	end if	
	'Shailesh 30-Oct-2009 Swift# 2438856
	 RSNPTval=cDbl(RSnpt("NPT"))
	 FrmNPTval=cDbl(Request.Form("hdnAgreeNPT"))
	if 	(intnewsev <>"" and intnewsev <>  RSnpt("sqseverity"))  then 
	    If not RSnpt.EOF then 
	         intMess =1 

	         RSnpt("sqseverity")=intnewsev
			 RSnpt("Severity")=Max(intnewsev,RSnpt("HSESeverity"))
		     RSnpt("NPT") = cDbl(Request.Form("hdnAgreeNPT"))
		     'RSnpt("majorchangedate") =Now ()
			 
			 conn1.execute("Update tblRIRp1 set MajorChangeDate=" &"'"& Now() & "'" & "where QID="& SafeNum(iQPID))    'To fix the issue caused by driver update TLS.
		     RSnpt.Update
	    End If
        strMessage = "Severity modified From " & cstr(GetSevClass(intsev)) & " to " & cstr(GetSevClass(intnewsev))
        LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",103,strMessage 
	else 

		If not RSnpt.EOF then 
			'Shaielsh 11-Nov-09 - Swift #2440954

			dim fNpt,LegacyNptvalue,disagreeNPTval,NoAgreementval,SQPQNPTval,NoAgreementvaltext,NoAgreementvaltextval
			
			
			if IsNull(RSnpt("SQPQNPT")) = false then
				SQPQNPTval = cdbl(RSnpt("SQPQNPT"))
			else
				SQPQNPTval = 0
			end if
			
			
			if IsNull(RSnpt("NoAgreement")) = false then
				NoAgreementval = cdbl(RSnpt("NoAgreement"))
				else
				NoAgreementval=0
			end if
			
			
			if IsNull(RSnpt("DisagreeNPT")) = false then
				disagreeNPTval = cdbl(RSnpt("DisagreeNPT"))
			else
				disagreeNPTval = 0
			end if
			
			if IsNull(RSnpt("NPT")) = false then
				fNPT = cdbl(RSnpt("NPT"))
			else
				fNPT = 0
			end if	
			
			if IsNull(RSnpt("LegacyNpt")) = false then
				LegacyNptvalue = cdbl(RSnpt("LegacyNpt"))
			else
				LegacyNptvalue = 0
			end if
			
			
					
			'Shaielsh 11-Nov-09 - Swift #2440954
			If fNPT <> cDbl(trim(Request.Form("hdnAgreeNPT"))) Then 'Shaielsh 09-Nov-09 add fNPT variable for checking & use trim - Swift #2440954
				RSnpt("NPT") = cDbl(Request.Form("hdnAgreeNPT"))

				RSnpt("DisagreeNPT") = DisagreeNPT
				RSnpt("CauseReview") =CauseReview
				RSnpt("SQPQNPT") = SQPQNPT
				
				if CauseReview=1 then 
				RSnpt("NPTConfirmed") =NULL
				RSnpt("NoAgreement") = NULL
				else
				RSnpt("NPTConfirmed") =NPTConfirmed
				RSnpt("NoAgreement") = NoAgreement
				end if
				
				if ((LegacyNptvalue)> (cDbl(Request.Form("txtAgreenpt"))) ) Then
				RSnpt("LegacyNpt") = LegacyNptvalue
				else
				RSnpt("LegacyNpt") = LegacyNpt
				end if

				
				if SQPQNPTval=0 then 
				RSnpt("OriginalNpt") =SQPQNPT
				end if
				RSnpt.Update
				
				If disagreeNPTval <> cDbl(Request.Form("txtAgreenpt")) Then
				strMessage = "Under Review NPT modified From " & disagreeNPTval & " to " & cDbl(Request.Form("txtAgreenpt"))
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",205,strMessage 
				end if
				
				if NoAgreementval=1 then
				NoAgreementvaltext="Yes"
				else
				NoAgreementvaltext="NO"
				end if
				
				if  Request.Form("optagreeNPT")=1 then
				NoAgreementvaltextval="Yes"
				else
				NoAgreementvaltextval="NO"
				end if
				
				
				If NoAgreementval <> Request.Form("optagreeNPT") Then
				strMessage = "Agree to disagree checkbox is modified From " & NoAgreementvaltext & " to " & NoAgreementvaltextval
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",206,strMessage 
				end if
			else

				RSnpt("DisagreeNPT") = DisagreeNPT
				RSnpt("CauseReview") =CauseReview
				RSnpt("SQPQNPT") = SQPQNPT
				if CauseReview=1 then 
				RSnpt("NPTConfirmed") =NULL
				RSnpt("NoAgreement") = NULL
				else
				RSnpt("NPTConfirmed") =NPTConfirmed
				RSnpt("NoAgreement") = NoAgreement
				end if
				if ((LegacyNptvalue)> (cDbl(Request.Form("txtAgreenpt"))) ) Then
				RSnpt("LegacyNpt") = LegacyNptvalue
				else
				RSnpt("LegacyNpt") = LegacyNpt
				end if
				
				if SQPQNPTval=0 then 
				RSnpt("OriginalNpt") =SQPQNPT
				end if
				
				RSnpt.Update
				
				If disagreeNPTval <> cDbl(Request.Form("txtAgreenpt")) Then
				strMessage = "Under Review NPT modified From " & disagreeNPTval & " to " & cDbl(Request.Form("txtAgreenpt"))
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",205,strMessage 
				end if
				
				if NoAgreementval=1 then
				NoAgreementvaltext="Yes"
				else
				NoAgreementvaltext="NO"
				end if
				
				if  Request.Form("optagreeNPT")=1 then
				NoAgreementvaltextval="Yes"
				else
				NoAgreementvaltextval="NO"
				end if
				
				
				If NoAgreementval <> Request.Form("optagreeNPT") Then
				strMessage = "Agree to disagree checkbox is modified From " & NoAgreementvaltext & " to " & NoAgreementvaltextval
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",206,strMessage 
				end if 
				
			End If
		End If
	end if 
	RSnpt.Close()
	set RSnpt = nothing 
	
	conn1.close
	    Set conn1 = Nothing
	
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime2.asp",err.Description & " QID=" & SafeNum(iQPID) &" Desc:UpdateNpt"
End If	
End Sub


Function NewSev (intNPT,intLoss,intSeverity)
'****************************************************************************************
'1. Function/Procedure Name          : NewSev
'2. Description           	         : To fetch new severity
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   12-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608 

'****************************************************************************************
On Error Resume Next
Dim Cn, rs_Severity, cmdErrorMsg
			
		conn.CursorLocation = 3
		Set rs_Severity = Server.CreateObject("ADODB.RecordSet")
		Set cmdErrorMsg = Server.CreateObject("ADODB.Command")
	
		With cmdErrorMsg
			.ActiveConnection = conn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_SeverityErrorMsg"
			
			.Parameters.Append .CreateParameter("@NPTValue", adDouble, adParamInput, 2,SafeNum(intNPT))
			.Parameters.Append .CreateParameter("@DolLoss", adDouble, adParamInput, 2,SafeNum(intLoss))
			.Parameters.Append .CreateParameter("@RIRSeverity", adInteger, adParamInput, ,SafeNum(intSeverity))
			
			Set rs_Severity = .Execute()
		End With
		if (rs_Severity.state = 1) then 
			If not rs_Severity.eof then
	             NewSev = rs_Severity("severity")
	        end if
		end if
		
		Set rs_Severity = Nothing
		'Cn.Close
		'Set Cn = Nothing
				 
If Err.Number <> 0 Then
' Log the ERROR
 ' LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTIME2.asp",,err.Description 
   LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID)  &" Desc:NewSev"
End If
		
End Function



Function typename (inttype)
'****************************************************************************************
'1. Function/Procedure Name          : NewSev
'2. Description           	         : To fetch Type
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   12-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608

'****************************************************************************************
On Error Resume Next
Dim RSType,strTypename,cmdLosscat,intTyp
intTyp = inttype
	
        conn.CursorLocation = 3
		Set RSType = Server.CreateObject("ADODB.RecordSet")
		Set cmdLosscat = Server.CreateObject("ADODB.Command")
	
		With cmdLosscat
			.ActiveConnection = conn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_fetchNPTvalues"
			.Parameters.Append .CreateParameter("@QID", adInteger, adParamInput, ,0)
			.Parameters.Append .CreateParameter("@case", adInteger, adParamInput, ,3)
			.Parameters.Append .CreateParameter("@strLoss", adVarChar,1,100," ")
			.Parameters.Append .CreateParameter("@intLossid", adInteger, adParamInput, ,cint(trim(intTyp)))
        Set RSType = .Execute()
		End With
		
	 if not RSType.EOF then
	     strTypename = trim(RSType("description"))
	     
	 end if 
     
     
     RSType.close()
     set RSType = nothing 
     
     typename = strTypename
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID) &" Desc:Typename"
End If
		
End Function

Function deletecost (Qpid,Typeid,RedMoney,seq)
'****************************************************************************************
'1. Function/Procedure Name          : NewSev
'2. Description           	         : To fetch new severity
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   12-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608

'****************************************************************************************
On Error Resume Next
Dim Cn, cmdDelete
		 
		conn.CursorLocation = 3
		'Set rs_Severity = Server.CreateObject("ADODB.RecordSet")
		Set cmdDelete = Server.CreateObject("ADODB.Command")
	 
		With cmdDelete
			.ActiveConnection = conn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_Deletetimeloss"
			.Parameters.Append .CreateParameter("@QPID", adInteger, adParamInput, ,Qpid)
			.Parameters.Append .CreateParameter("@Typeid", adInteger, adParamInput, ,Typeid)
			.Parameters.Append .CreateParameter("@RedMoney", adDouble, adParamInput,2 ,RedMoney)
			.Parameters.Append .CreateParameter("@seq", adInteger, adParamInput, ,seq)
			.Execute()
		End With
	   				 
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID) &" Desc: deletecost"
End If
		
End Function

'function added for NPT SWIFT #2401608 to show severity
Function GetSevClass(Id)
On Error Resume Next
Dim SQL,RSSev
	SQL=" Select 'S' as Type,SeverityID,SeverityDesc from tlkpRIRSeverity where SeverityID ="&Id &""
	Set RSSev=conn.execute(SQL)
	If not RSSev.eof then 
	    GetSevClass = RSSev("SeverityDesc")
	 end if 
	RSSev.close
	Set RSSev=nothing
	
If Err.Number <> 0 Then
' Log the ERROR
   LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID) &" Desc:GetSevClass"
End If
End Function

Sub UpdateNpt1 ()

'****************************************************************************************
'1. Function/Procedure Name          : UpdateNpt1
'2. Description           	         : Update NPT,severity if it is changed 
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   12-Aug-2009			    Nilesh Naik	        		Added for NPT SWIFT #2401608 
'   22-Oct-2009			    Micheal Anthony				NPT SWIFT #2434444 
'****************************************************************************************
On Error Resume Next
	Dim RSnpt, iRecCount, strSQL,intsev,RSCostCat,intnewsev,inttotalcost,strMessage
	inttotalcost = 0 	
	       
	intsev = cint(Request.Form("severity"))
	'''call severity and it is true 
    
	If blnNPT_Exempt = 0 Then 'Shailesh 30-Oct-2009 Swift# 2438856 		
		intnewsev = NewSev(Request.Form("hdnAgreeNPT"),inttotalcost,intsev)
	'Shailesh 30-Oct-2009 Swift# 2438856	
	else
	    intnewsev = ""
	end if	
	'Shailesh 30-Oct-2009 Swift# 2438856  
	 
    Set RSnpt=Server.CreateObject("ADODB.Recordset")
    RSnpt.LockType = 3	
	RSnpt.Open "SELECT QID,RptDate,sqseverity,Severity,HSESeverity,SQPQNPT,DisagreeNPT,NoAgreement,NPT,LegacyNpt,CauseReview,NPTConfirmed,OriginalNpt FROM tblrirp1  WHERE qid=" &  SafeNum(iQPID) & " ",conn 
	if 	(intnewsev <>"" and intnewsev <>  intsev)  then 
	    If not RSnpt.EOF then 
	         intMess =1
	         RSnpt("sqseverity")=intnewsev
			 RSnpt("Severity")=Max(intnewsev,RSnpt("HSESeverity"))
		     RSnpt("NPT") = cDbl(Request.Form("hdnAgreeNPT"))
		     'RSnpt("majorchangedate") =Now ()
			 
			 conn.execute("Update tblRIRp1 set MajorChangeDate=" &"'"& Now() & "'" & "where QID="& SafeNum(iQPID))    'To fix the issue caused by driver update TLS.
		     RSnpt.Update
	    End If
        strMessage = "Severity modified From " & cstr(GetSevClass(intsev)) & " to " & cstr(GetSevClass(intnewsev))
        LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",103,strMessage 
	else 
		If not RSnpt.EOF then 
			'Shaielsh 11-Nov-09 - Swift #2440954
			dim fNpt,LegacyNptvalue,disagreeNPTval,NoAgreementval,SQPQNPTval,NoAgreementvaltext,NoAgreementvaltextval
			if IsNull(RSnpt("SQPQNPT")) = false then
				SQPQNPTval = cdbl(RSnpt("SQPQNPT"))
			else
				SQPQNPTval = 0
			end if
			
			
			if IsNull(RSnpt("NoAgreement")) = false then
				NoAgreementval = cdbl(RSnpt("NoAgreement"))
				else
				NoAgreementval=0
			end if
			
			
			if IsNull(RSnpt("DisagreeNPT")) = false then
				disagreeNPTval = cdbl(RSnpt("DisagreeNPT"))
			else
				disagreeNPTval = 0
			end if
			
			if IsNull(RSnpt("NPT")) = false then
				fNPT = cdbl(RSnpt("NPT"))
			else
				fNPT = 0
			end if	
			
			if IsNull(RSnpt("LegacyNpt")) = false then
				LegacyNptvalue = cdbl(RSnpt("LegacyNpt"))
			else
				LegacyNptvalue = 0
			end if

			'Shaielsh 11-Nov-09 - Swift #2440954		
			If fNPT <> cDbl(trim(Request.Form("hdnAgreeNPT"))) Then 'Shaielsh 09-Nov-09 add fNPT variable for checking & use trim - Swift #2440954
				RSnpt("NPT") = cDbl(Request.Form("hdnAgreeNPT"))

				RSnpt("DisagreeNPT") = DisagreeNPT
				RSnpt("CauseReview") =CauseReview
			
				RSnpt("SQPQNPT") = SQPQNPT
				
				if CauseReview=1 then 
				RSnpt("NPTConfirmed") =NULL
				RSnpt("NoAgreement") = NULL
				else
				RSnpt("NPTConfirmed") =NPTConfirmed
				RSnpt("NoAgreement") = NoAgreement
				end if
				
				if ((LegacyNptvalue)> (cDbl(Request.Form("txtAgreenpt"))) ) Then
				RSnpt("LegacyNpt") = LegacyNptvalue
				else
				RSnpt("LegacyNpt") = LegacyNpt
				end if

				
				if SQPQNPTval=0 then 
				RSnpt("OriginalNpt") =SQPQNPT
				end if
				RSnpt.Update
				
				If disagreeNPTval <> cDbl(Request.Form("txtAgreenpt")) Then
				strMessage = "Under Review NPT modified From " & disagreeNPTval & " to " & cDbl(Request.Form("txtAgreenpt"))
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",205,strMessage 
				end if
				
				if NoAgreementval=1 then
				NoAgreementvaltext="Yes"
				else
				NoAgreementvaltext="NO"
				end if
				
				if  Request.Form("optagreeNPT")=1 then
				NoAgreementvaltextval="Yes"
				else
				NoAgreementvaltextval="NO"
				end if
				
				
				If NoAgreementval <> Request.Form("optagreeNPT") Then
				strMessage = "Agree to disagree checkbox is modified From " & NoAgreementvaltext & " to " & NoAgreementvaltextval
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",206,strMessage 
				end if
				
				else

				RSnpt("DisagreeNPT") = DisagreeNPT
				RSnpt("CauseReview") =CauseReview
				
				RSnpt("SQPQNPT") = SQPQNPT
				if CauseReview=1 then 
				RSnpt("NPTConfirmed") =NULL
				RSnpt("NoAgreement") = NULL
				else
				RSnpt("NPTConfirmed") =NPTConfirmed
				RSnpt("NoAgreement") = NoAgreement
				end if
				if ((LegacyNptvalue)> (cDbl(Request.Form("txtAgreenpt"))) ) Then
				RSnpt("LegacyNpt") = LegacyNptvalue
				else
				RSnpt("LegacyNpt") = LegacyNpt
				end if
				
				if SQPQNPTval=0 then 
				RSnpt("OriginalNpt") =SQPQNPT
				end if
				
				RSnpt.Update
				
				If disagreeNPTval <> cDbl(Request.Form("txtAgreenpt")) Then
				strMessage = "Under Review NPT modified From " & disagreeNPTval & " to " & cDbl(Request.Form("txtAgreenpt"))
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",205,strMessage 
				end if
				
				if NoAgreementval=1 then
				NoAgreementvaltext="Yes"
				else
				NoAgreementvaltext="NO"
				end if
				
				if  Request.Form("optagreeNPT")=1 then
				NoAgreementvaltextval="Yes"
				else
				NoAgreementvaltextval="NO"
				end if
				
				
				If NoAgreementval <> Request.Form("optagreeNPT") Then
				strMessage = "Agree to disagree checkbox is modified From " & NoAgreementvaltext & " to " & NoAgreementvaltextval
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","LOSSTime2.asp",206,strMessage 
				end if 
				
				
			End If
		End If
	end if 

RSnpt.Close()
set RSnpt = nothing 			

If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID) &".Desc : UpdateNpt1"
End If	
End Sub


Sub deleteallTime (Qpid,intNPT)
'****************************************************************************************
'1. Function/Procedure Name          : deleteall
'2. Description           	         : To delete the time loss data 
'3. Calling Forms:   	             : LOSSTIME2.asp
'4. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   27-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608

'****************************************************************************************
On Error Resume Next
Dim Cn, cmdDelete
		 
		conn.CursorLocation = 3
		
		Set cmdDelete = Server.CreateObject("ADODB.Command")
	   	    
		With cmdDelete
			.ActiveConnection = conn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_DeleteAlltimeloss"
			.Parameters.Append .CreateParameter("@QPID", adInteger, adParamInput, ,Qpid)
			.Parameters.Append .CreateParameter("@NPTValue", adDouble, adParamInput, 2,intNPT)
			
			.Execute()
		End With
	   				 
If Err.Number <> 0 Then
' Log the ERROR
  LogEntry 2,"LossTime2.asp",err.Description  & " QID=" & SafeNum(iQPID) &" Desc:deleteallTime"
End If
		
End sub
	
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
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSTime2.asp;21 %>
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
<% '       7*[1275699] 29-SEP-2009 15:48:26 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '       8*[1277927] 02-OCT-2009 17:50:40 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '       9*[1280834] 09-OCT-2009 15:42:30 (GMT) MAnthony2 %>
<% '         "Swift # 2401608 - NPT Changes" %>
<% '      10*[1282666] 15-OCT-2009 01:40:28 (GMT) SVadla %>
<% '         "NPT Hot fix for Overall Severity and Mail Header" %>
<% '      11*[1282819] 15-OCT-2009 17:07:25 (GMT) MAnthony2 %>
<% '         "SWIFT # 2434444 - NPT Issues - post go live" %>
<% '      12*[1285260] 23-OCT-2009 17:24:27 (GMT) MAnthony2 %>
<% '         "SWIFT # 2401608 - NPT Bugs" %>
<% '      13*[1289354] 03-NOV-2009 05:23:47 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      14*[1290343] 04-NOV-2009 13:41:57 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      15*[1292487] 10-NOV-2009 07:38:10 (GMT) SKadam3 %>
<% '         "SWIFT #2440954 - NPT - HotFix - Invalid Use of Null" %>
<% '      16*[1293409] 13-NOV-2009 16:09:13 (GMT) MAnthony2 %>
<% '         "LossTime2.asp - NPT Error correction" %>
<% '      17*[1650037] 03-AUG-2012 07:07:01 (GMT) MAnthony2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '      18*[1633354] 07-AUG-2012 16:27:13 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      19*[1666353] 14-SEP-2012 12:13:23 (GMT) MPatil2 %>
<% '         "SWIFT #2657957 - Fix:RIR update date reflects edits made to Contractor, Investigation & Time Loss" %>
<% '      20*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      21*[1837481] 20-MAY-2014 10:40:11 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture." %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSTime2.asp;21 %>
