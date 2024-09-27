<%@ Language=VBScript %>
<%option explicit%>

<%
'****************************************************************************************
'1. File Name		              :  RIRdisp2.asp
'2. Description           	      :  Main RIR Page for saving the RIR entry in database
'3. Calling Forms   	          : 
'4. Stored Procedures Used        : 
'5. Views Used	   	              : 
'6. Module	   	                 : RIR (HSE/SQ)				
'7. History
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   5-Aug-2009			     Nilesh Naik        	 	Modified - changed for NPT SWIFT #2401608
'   29-Sep-2009			 	 Nilesh Naik        	 	Modified - changed for NPT SWIFT #2401608 to fix the bugs
'   7-May-2014            	 Varun Sharma               Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
'   06-Nov-2014              Varun Sharma               ENH044752  HSE locking of lagging indicators - safety net - key to unlock
'	17-Feb-2015				 Sagar Chaudhari			Modified - ENH053415-Addition of PTEC project	 acknowledgement on HSE and SQ RIR and OI reports
'   18-March-2015	  		 Rohan Bhalave				NFT056368  FEATURE QUEST Upgrades for Facilities
'    17-Apr-2015			 Rohan Bhalave				ENH059068 P&AM Process Safety (Updates in tab and the report)
' 10-July-2015				Sagar Chaudhari				ENH077170 - SAXON - Operations at Time of Event (Category and Sub-Category)
' 	20-OCT-2015			    Sagar Chaudhari 			ENH086389 - Saxon - SQ Categories
'   23-Nov-2015				Rohan Bhalave          		NFT087279 - TS segment becoming a forced segment - Rig related flag at sub sub segment level
' 	17-FEB-2016			    Varun Sharma				    ENH096140 - Integrated Projects
'****************************************************************************************
%>

<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<!-- #INCLUDE FILE="Inc_SQWLNotify.asp"-->
<!-- #INCLUDE FILE="Inc_SQWISNotify.asp"-->
<% 

	checktimeout()


	Dim dtRptDate, sRptNo, sKey,lPL 'Added by Deepak for D&M
	Dim bNotify, cn, sSQL, rs2
	Dim iClass, iSev, bNew,delopfdetail
	Dim iTemp, sTemp, bTemp,HideSQ,HideWSSQ
	Dim iQPID,lOrgNo, dtTemp,bHSEDisplayed,bSQDisplayed,bHSESelected,bSQSelected,dteventTemp,dtfrmTemp,frmSLBInvment,RSSLBInvment,StrEvent,frmStrEvent
	Dim RS, sErrors, pTemp, currentUTCDateTime, RS1, UpdateMCD, RSAls, aSQL,RSsqlnvInvmen,frmRSsqlnvInvment,rsrelated,rsexternal,frmrelated,frmexternal, OrgWIBEvent, OrgNPT
	Dim RIRStatus, RIRStatusDestEmail, IsSPRequired, IsDamageRequired,RSsqlnvInvment,frmCAffect,rsClientAffected, inttxtNPT,intOrgNPT 
	Dim CategoryList 'used to assist in sending email notifications of closures/reviews
	Dim Reporter, ReporterUID, BSegID,rstrClientAffected,frmstrClientAffected,RSsqClientAffect
	DIM ACLDefined, iRAOpen,iRAClosed
	Dim iHSESev, iSQSev,bAutoProtect,sUID,sUName
	Dim SLBInv, IndRec, SLBCon , SQStd,RWPQ,QID,MsgID,txtComments,arrSevClass
	Dim NotifyIndex,External,SLBRel,ClientAffect,WellSite,confMsg,ProtectDoc,DelRS
    Dim StrClientAffected '---added by ANIL for NPT <<2401608>>
    Dim sBig,RSCont,Slist,sContVal,sSerID,sRisk,sMode,sContID,iCtr,spsl4ID,rsspsl4ID,SPS_ID,RigOprVal
	Dim SwiQn1, SwiQn2A, SwiQn2B, SwiQn3, RSswi,sSWI,lockval, spsl2,spsl3,spsl4,LocBSID, EnforceSel, SQMappingID, spsb2
	Dim SlimCntrl, SlimTech, SlimProc, SlimCompet, SlimBehav, sSlim, RSslim,ISubBSIDs,SQPQNPTval,EventSubClassSafety,iBSID
	Dim Asset_SN_ID, FN_ID, IsFailureCreated, WO_ID, hasEquipmentResult, ErrorMessages 
	Dim iogpRs,iogpTier,varHideLegacyIogp
    Dim strWL,rsWL,strWLSQ,rsWLSQ,hdChkInput,WlRo
	On Error Resume Next
	
	NotifyIndex=False
	WellSite=0
	SLBInv = False
	IndRec = False
	SLBCon = False
	SLBRel = False
	External = False
	ProtectDoc = False
	sUID=lcase(trim(Session("UID")))
	sUName=trim(Session("UserName"))
	If sUID="" Then
		sUID=trim(Request.Form("CreatorID"))
		sUName=trim(Request.Form("CreatorNm"))
	End IF
lockval = Request.Form("txtlock")
'IOGP legacy
varHideLegacyIogp = fncD_Configuration("LegacyIogp")
'TS Change
LocBSID = Request.Form("LocBSID")   
EnforceSel= Request.Form("EnforceSelection")
hdChkInput=Request.Form("hdnchkinput")

Dim opfepcc,opfonm,isOPFval,hdClass,iclass1,optSQ,optClass,iClassn
		opfepcc=Request.Form("opfepcc")
		opfonm=Request.Form("opfonm")
		isOPFval=Request.Form("isOPFval")
		hdClass=Request.Form("hdClass")
		iclass1=Request.Form("iclass")
		optSQ=Request.Form("optSQ")
		optClass=Request.Form("optClass")
		iClassn=Request.Form("iClassn")
		if isOPFval="True" and optSQ="on" then
		if cint(iClassn)<>cint(optClass) and isOPFval="True" and optSQ="on" and (opfepcc=1 or opfonm=1) and optClass<>""  and (not(cint(iClassn)= 2 and cint(optClass)=3)) and not((cint(iClassn)= 3 and cint(optClass)=2)) then
			delopfdetail=1
		end if 
		end if
	


spsl2 =  iif (request("SQL2_0")="",0,Request("SQL2_0"))
spsl3 =  iif (request("SQL3_0")="",0,Request("SQL3_0"))
spsl4 =  iif (request("SQL4_0")="",0,Request("SQL4_0"))

    spsb2 =  iif (request("SQB2_0")="",0,Request("SQB2_0"))
	SwiQn1  =  iif (request("swiqn")="",0,Request("swiqn"))
	SwiQn2A =  iif (request("swiqntwo")="",0,Request("swiqntwo"))
	SwiQn2B =  iif (request("swiqntwob")="",0,Request("swiqntwob"))
	SwiQn3  =  iif (request("swiqnthree")="",0,Request("swiqnthree"))

 ' SLIM CHANGES
	'SlimCntrl  =  iif (request("sCtrl")="",0,Request("sCtrl"))
	'SlimTech  =  iif (request("sTech")="",0,Request("sTech"))
	'SlimProc  =  iif (request("sProc")="",0,Request("sProc"))
	'SlimCompet  =  iif (request("sComp")="",0,Request("sComp"))
	'SlimBehav  =  iif (request("sBehav")="",0,Request("sBehav"))

	
	frmSLBInvment=Request("optSLBInvment")
	RSSLBInvment=Request("slbin")
	
	frmRSsqlnvInvment=Request("optSQInvment")
	RSsqlnvInvment=Request("sqlnv")
	
	RSsqClientAffect=Request("sqClientAffect")
	
	frmCAffect=Request.Form("OptCAffect")
	If Request("optSLBInvment") <> "" Then
		Select Case Request("optSLBInvment")
			case 1
				SLBInv = True
				IndRec = true
				SLBCon = False
			case 2
				SLBInv = True
				IndRec = False
				SLBCon = False
			case 3
				SLBInv = False
				IndRec = False
				SLBCon = true
		End Select
	End If
	
	If Request("optSQInvment") <> "" Then
		Select Case Request("optSQInvment")
			case 1
				SLBRel = True
				External = False
			case 2
				SLBRel = True
				External = True
		End Select
	End If
	
	txtComments=""
	'NBL 08/02/02
	'First Param - Server Variables on the form
	'Second Param - Referer Page
	Call ResetServerVariables(Request.Form("sServerVariables"),"RIRdsp.asp")

	UpdateMCD = False
	IsSPRequired = false
	IsDamageRequired = False
	
	lOrgNo = Request.QueryString("OrgNo")
	dtRptDate = Request.QueryString("rptDate")
	iQPID = Request.QueryString("QPID")
	If iQPID="" Then iQPID=0
	Set cn = GetNewCN()

'WL update RO in WL SQ Details tab
	strWL = "select distinct(isroinvolved) from tblRIR_SQWLIncidents where qpid="& SafeNum(iQPID)&" and isroinvolved=1"	
	SET rsWL = cn.execute (strWL)
	 If NOT rsWL.eof Then			
		WlRo = 1
	 Else
		WlRo = 0
	 End If	
	rsWL.close
	Set rsWL = Nothing

	if hdChkInput="True" and WlRo = "0" Then
		cn.Execute "Update tblRIR_SQWLIncidents set IsRoInvolved=1 Where QPID = " & SafeNum(iQPID)
	 Else if hdChkInput="False" Then
		cn.Execute "Update tblRIR_SQWLIncidents set IsRoInvolved=2 Where QPID = " & SafeNum(iQPID)
	 End IF
	End IF
  If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"RIRDsp2.asp",err.Description  &"-At Line188 QID="& iQPID
	End If	
	On Error Resume Next
	If delopfdetail="1" and delopfdetail<>"" and Request.Form("chkClosed")="" Then 
		fncDeleteOPFDetails()
	End If
	
		spsl4ID = "SELECT ID FROM tblRIR_SPSData WHERE Suffix ='" & spsl4 &"'"
		Set rsspsl4ID = cn.execute(spsl4ID)
			 if not rsspsl4ID.EOF then
			SPS_ID=rsspsl4ID("ID")
			 else
			SPS_ID=0
			end if
	
	'Check for user access.
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
	bHSEDisplayed = False
	bSQDisplayed = False
	If Request.Form("HSEDisplayed")=1 then bHSEDisplayed = True
	If Request.Form("SQDisplayed")=1 then bSQDisplayed = True

	'HSE / SQ Selection
	bHSESelected = False
	bSQSelected = False
	If Request.Form("optHSE")="on" then bHSESelected = True
	If Request.Form("optSQ")="on" then bSQSelected = True

	'Added by Deepak for D&M Tab Development
	lPL=GetProductLine(lOrgNo)
	iBSID=GetSubBusinessSegID(lOrgNo)
	ISubBSIDs=  GetSubSubBusinessSegID(lOrgNo)
	SQMappingID = GetSQMappingID(ISubBSIDs)
		
	if not (isREWSQMapping(SQMappingID) or isSPWL(LocBSID) or (isOFS(lPL) and not (isMNSIT(LocBSID))) or isWTSSQMapping(SQMappingID) or isEMS(SQMappingID) or isWSSQMapping(SQMappingID) or isIPMSeg(SQMappingID) or isSWACO(SQMappingID) or (isOne(SQMappingID) and not (isOneCPL(LocBSID)))) then HideSQ = 1 else HideSQ=0    'isWTS(lPL)  --isRew(lPL)  isWS(lPL)
    if isWSSQMapping(SQMappingID) then HideWSSQ = 1   'isWS(lPL)
	'Delete Illumina Data - start
		If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"RIRDsp2.asp",err.Description  &"-At Line228 QID="& iQPID
	End If	
	If Request.Form("ALSDelete") = "1" Then
		On Error Resume Next
		aSQL = "SELECT * FROM tblRIR_SQALSDetailsGeneral WHERE QPID="& SafeNum(iQPID) 
		Set RSAls = Server.CreateObject("ADODB.Recordset")		
		RSAls.LockType = 3
		RSAls.Open aSQL, cn
		If Not RSAls.BOF Or Not RSAls.EOF Then 
			RSAls("Illumina_Job_AID") = ""
			RSAls("ActivityType") = 0
			RSAls("WellName") = Null
			RSAls("FieldName") = Null
			RSAls("Pull_Reason") = 0
			RSAls("Pull_Reason_Specific") = 0
			RSAls("LiftWatcher") = 0 
			RSAls("GOR") = Null
			RSAls("GOR_UOM") = 0
			RSAls("Oil_Gravity") = Null
			RSAls("Water_Gravity") = Null
			RSAls("Water_Cut") = Null
			RSAls("CO2") = Null
			RSAls("H2S") = Null
			RSAls("H2") = Null
			RSAls("CO") = Null
			RSAls("N2") = Null

			RSAls("Corrosive") = 0 
			RSAls("NORM") = 0
			RSAls("Abrasive") = 0
			RSAls("Paraffin") = 0
			RSAls("Scale") = 0
			RSAls("Other") = 0

			RSAls("Environment") = 0
			RSAls("BHT") = Null
			RSAls("BHT_UOM") = 0
			RSAls("Installation_Type") = 0
			RSAls("Well_Geometry") = 0
			RSAls("Total_Depth_MD") = Null
			RSAls("Total_Depth_UOM") = 0
			RSAls("ESP_Bottom_Depth") = Null
			RSAls("ESP_Bottom_Depth_UOM") = Null
			RSAls.Update		
			RSAls.Close
		End If	
				
		Set RSAls = Nothing
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp2.asp",err.Description  &"-At Line271 QID="& iQPID
		End If		
	End If	
	'Delete Illumina Data - end
	'Check for delete
	If Request.QueryString("Delete")=1 and iQPID>0 then 		
		sSQL = "SELECT * FROM tblRIRp1 WHERE QID=" & iQPID
		Set RS = Server.CreateObject("ADODB.Recordset")		
		RS.Open sSQL, cn
		If rs.EOF or rs.BOF then Response.Redirect("../Utils/RecNotFound.htm")
		Set DelRS = Server.CreateObject("ADODB.Recordset") 
		Set DelRS = cn.Execute("sp_updatedeleted "  & SafeNum(lOrgNo) & ",'" & dtRptDate & "', RIR , '" & formatquotes(sUName)& "','" & sUID & "'")
		If DelRS("RetVal") = 0 Then 
			LogAuditTrail  RS("QID"),lorgNo,dtRptdate,"R","RIRDsp2.asp",100,""
		
	    	Set RS = Nothing
			Set DelRS = Nothing
	    	%> 
			<SCRIPT LANGUAGE=javascript>
			<!--
			    alert("Record Successfully Deleted");
			    location.href = 'RIRcrit.asp?RptType=RIR&RptName=RIRlst.asp&Criteria=yes';
			//-->
			</SCRIPT>
			<%
		Else
			%> 
			<SCRIPT LANGUAGE=javascript>
			<!--
			    alert("Deletion Failed");
			//-->
			</SCRIPT>
			<%
			sKey ="?OrgNo=" & lOrgNo &"&RptDate=" & server.URLEncode(dtrptdate)
			pTemp="RIRdsp.asp" & sKey & "&Msg=" & Server.URLEncode("Data Successfully Saved")
			Response.Redirect(pTemp)
		End If
	end if
 
	'RIR Validation
	'------------------------------------------------------------------

	bNotify=False
	sErrors=""
	
	'HSE/SQ
	If Not bHSESelected and Not bSQSelected Then
		bNotify=True
		sErrors = sErrors & "HSE and/or Service Quality box(es) must be checked " & "<BR>"
	End If

	'use bNotify to validate, sErrors to hold error string
	dtTemp = Request.Form("txtEvDate")
dtfrmTemp = Request.Form("txtEvDate")
	If Not IsDate(dtTemp) Then
		bNotify = true
		sErrors = sErrors & "Invalid Event Date - '" & dtTemp & "'<BR>"
	Else
		dtTemp = CDate(dtTemp)
		if not validateEventDate(dtRptDate,dtTemp) then
   			bNotify = True
			sErrors = sErrors & "Invalid Event Date - '" & dtTemp & "'. It is either in the future or too far in the past.<BR>"
		end if

	
	End if
	
	dtTemp = Request.Form("txtEvTime")
	if not IsDate(dtTemp) then
		bNotify = true
		sErrors = sErrors & "Invalid Event Time - '" & dtTemp & "'<BR>"
	end if
	
	'Classification
	'--------------------------------------------------------------
	iClass = Request.Form("optClass")
					
	If iClass = 0 Then
		bNotify = True
		sErrors = sErrors & "No Incident Classification Selected.<BR>"
	End If
			
	'Severity
	'--------------------------------------------------------------
	
	If iClass=1 then
		iHSESev = Request("cmbHSESeverity")
		iSQSev = Request("cmbSQSeverity")

		If iHSESev = "" Then iHSESev = 0
		If iSQSev = "" Then iSQSev = 0
		
		If (iHSESev + iSQSev = 0) and (lockval <> 0) Then
			bNotify = True
			sErrors = sErrors & "No Severity Specified for Accident.<BR>"
		End If

		'If HSE is not checked, HSESev must be 0
		If Not bHSESelected then iHSESev = 0
		'If SQ is not checked, SQSev must be 0
		If Not bSQSelected then iSQSev = 0
	Else
		iHSESev = 0
		iSQSev = 0
	End if
	iSev = MAX(iHSESev,iSQSev)

	'Injury / Illness
	'--------------------------------------------------------------
	If (Request("LossCat_A2") = "1" OR Request("LossCat_A3") = "1") AND Request("LossCat_A1") = "1" then
		bNotify = True
		sErrors = sErrors & "Cannot select Illness and Injury.<BR>"
	End If

	'Reporter
	'--------------------------------------------------------------
	sTemp = Request.Form("txtReporter")
	If Len(Trim(sTemp)) = 0 Then
		bNotify = True
		sErrors = sErrors & "No Reporter Entered.<BR>"
	End If

	'Description
	'--------------------------------------------------------------
	sTemp = Request.Form("txtShortDesc")
	If Len(Trim(sTemp)) = 0 Then
		bNotify = True
		sErrors = sErrors & "No Short Description Entered.<BR>"
	End If
	
	'Location
	'--------------------------------------------------------------
	iTemp = Request.Form("txtLocation")
	
	If iTemp = "0" Then
		bNotify = True
		sErrors = sErrors & "No Incident Location Selected.<BR>"
	ElseIf instr(iTemp,"RIG") Then
	       if Request.Form("txtCRMRigID")="" Then
		    bNotify = True
		    sErrors = sErrors & "No CRM RIG Name Selected.<BR>"
	       Elseif ((Request.Form("txtCRMRigID")="NO-CRM-RIG") AND Trim(Request.Form("txtLoc"))="") Then 
		   'If Trim(Request.Form("txtLoc"))="" Then
			 bNotify = True
		    	 sErrors = sErrors & "Please Select Either CRM RIG Name OR Enter Site Name.<BR>"
		   'End IF	
	       END IF
	ElseIf instr(iTemp,"OPT") Then
	       if ((Request.Form("txtCRMRigID")="" OR Request.Form("txtCRMRigID")="NO-CRM-RIG") AND Trim(Request.Form("txtLoc"))="") Then
		    bNotify = True
		    sErrors = sErrors & "Please Select Either CRM RIG Name OR Enter Site Name.<BR>"
	       END IF
	'Else
		 'If Trim(Request.Form("txtLoc"))="" Then
			 'bNotify = True
		    	 'sErrors = sErrors & "No Site Name Entered.<BR>"
		 'End IF
	End If

	'Business Segment
	'--------------------------------------------------------------
	sTemp = Request.Form("txtBSegment")
	
	If Len(Trim(sTemp)) = 0 and not EnforceSel  Then   'TS changes
		bNotify = True
		sErrors = sErrors & "No Business Segment Selected.<BR>"
	End If

	'NEW Category @@@peter 1/31/2001 
	Dim row, bValid

	bValid=false
	For row = 0 To UBound(Application("a_LossCategories"), 2)
		If Request.Form(Application("a_LossCategories")(0,row)) = "1" Then bValid=true
	Next
    
	' *******************************************************************************************
    ' Code changed for NPT <<2401608>> disabled checkbox values are not being  got properly 
    ' *******************************************************************************************
    For row = 0 To UBound(Application("a_LossCategories"), 2)
        If Request.Form(Application("a_LossCategories")(0,row)) = "" and (instr(1,Application("a_LossCategories")(0,row),"_G")) Then
        'Start add for SWIFT # 2401609	

            If Request("optSQInvment") <> "" Then
                if ((Request("optSQInvment") ="1") or (Request("optSQInvment") ="2")) then
                    bValid=true
                end If
            end if   
          ' End add
        end if 
	Next
    ' ****************************************************************************************
	If not bValid then
		bNotify = True
		sErrors = sErrors & "No Category Selections Made.<BR>"
	End If
   
	'Check for open RIRs when closing
	GetRACount iQPID,"R",iRAOpen,iRAClosed
	If Request.Form("chkClosed")="on" AND iRAOpen > 0  Then
		bNotify = True
		sErrors = sErrors & "Report cannot be closed.<BR>There are " &   iRAOpen & " Remedial Actions Open. <br>  Please close all remedial actions and try again."
	End If

	'HSE Related
	If bHSESelected and bHSEDisplayed Then	
		'Hazard Cat
		iTemp = Request.Form("txtHazard")
		if Request.Form("optHSE")="on" and (LockCount<>0 or chkHSELockingMgmt()) then 
		If iTemp = 0 Then
			bNotify = True
			sErrors = sErrors & "No Hazard Category Selected.<BR>"
		End If
		End If
	End If

	'SQ Related
	If bSQSelected and bSQDisplayed Then

		IsSPRequired = cbool(trim(Request.Form("IsSPRequired")))
		IsDamageRequired = cbool(trim(Request.Form("IsDamageRequired")))
		HideSQ = cbool(trim(Request.Form("HideSQ")))
		HideWSSQ = cbool(trim(Request.Form("HideWSSQ")))
        'Changed by Deepak For D&M tab Development
		'Service/Product
		If ((IsSPRequired and (not isDMSQMapping(SQMappingID) or not ShowDM())) AND (IsSPRequired and (not isPF(lPL) or not ShowPF()))) Then   'isDM(lPL)
			'#2594250 - GSS ML Data Tab
			If request.Form("txtIsGSSML") = "YES" Then
				If trim(Request.Form("txtSPCategory")) = "" Then bNotify = True : sErrors = sErrors & "No Service/Product Category Selected.<BR>"
				If trim(Request.Form("txtSubSPCategory")) = "" OR trim(Request.Form("txtSubSPCategory")) = "0" Then bNotify = True : sErrors = sErrors & "No Service/Product Sub Category Selected.<BR>"
			End If
			'#2594250 - GSS ML Data Tab
		End If
		 'Damage
		If ((isDamageRequired and (not isDMSQMapping(SQMappingID)  or not ShowDM())) AND (isDamageRequired and (not isPF(lPL) or not ShowPF()))) Then   'isDM(lPL)
			'#2594250 - GSS ML Data Tab
			'If request.Form("txtIsGSSML") = "YES" and not issaxon(lPL)  Then
			If request.Form("txtIsGSSML") = "YES" and not issaxon(GetSubSubBusinessSegID(lOrgNo))  Then
				If trim(Request.Form("txtDamage")) = "" Then bNotify = True : sErrors = sErrors & "No Damage Category Selected.<BR>"
				If trim(Request.Form("txtSubDamage")) ="" OR trim(Request.Form("txtSubDamage")) ="0" Then bNotify = True : sErrors = sErrors & "No Damage Sub Category Selected.<BR>"
			End If
			'#2594250 - GSS ML Data Tab
		End IF
		
		'Failure -- @Visali 02/15/2006 
		if (((HideSQ and not HideWSSQ) and (not isDMSQMapping(SQMappingID) or not ShowDM())) AND ((HideSQ and not HideWSSQ) and (not isPF(lPL) or not ShowPF()))) then		    'isDM(lPL)
			
			'#2594250 - GSS ML Data Tab
			If request.Form("txtIsGSSML") = "YES" Then
				If trim(Request.Form("txtFailure")) = "" Then bNotify = True : sErrors = sErrors & "No Failure Category Selected.<BR>"
				If trim(Request.Form("txtSubFailure")) = "" OR trim(Request.Form("txtSubFailure")) = "0" Then bNotify = True : sErrors = sErrors & "No Failure Sub Category Selected.<BR>"
			End If
			'#2594250 - GSS ML Data Tab
			
		end if
	
	End If

'6/13/2002 peter Check for fatality
	If iSev <> 4 and iclass=1 and SLBInv and IsFatality() Then
		 bNotify = True 
		 sErrors = sErrors & "You have a fatality reported on the personnel loss screen.<BR>"
		 sErrors = sErrors & "SLB involved fatality incidents must be classified as Catastrophic<BR>"
		 sErrors = sErrors & "If you wish to change the classification then remove the fatality first.<BR>"
		 sErrors = sErrors & "YOUR RECORD HAS NOT BEEN SAVED<BR>"
	End If

	'Open RS
	'------------------------------------------------------------------
	If bNotify = false then	
		'Extract all severity and class types
		InitSeverityAndClass
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		sTemp = "SELECT * FROM tblRIRp1 WHERE OrgNo=" & SafeNum(lOrgNo) & " AND RptDate='" & dtRptDate & "'"	  
	    'Response.Write stemp
	    RS.LockType = 3
		RS.Open sTemp, cn

		'Update data
		'------------------------------------------------------------------
		'Header
		bNew=false
		currentUTCDateTime = getUTC()
		If rs.EOF then 
			RS.AddNew
			RS("OrgNo") = lOrgNo
			'@@@ 5/5/2000 start using gmt for creation date
			dtRptDate = currentUTCDateTime
			RS("RptDate") = dtRptDate
			RS("RevDate") = dtRptDate
			RS("MajorChangeDate") = Now()
			RS("Source") = 1	'From DirectEntry
			UpdateMCD = True
			bNew=true
			If lcase(trim(Request.form("txtReporter"))) <> "anonymous" Then
				RS("CreateName")= left(sUName,60)
				RS("CreateUID") = left(sUID,30)
			Else
				RS("CreateName")= "Anonymous"
				RS("CreateUID") = NULL
			End if
			iQPID = 0
			MsgID=101'New Record
		Else
			iQPID = RS("QID")	
			MsgID=102'Exist Record Updation		
		End if
		
		sKey ="?OrgNo=" & lOrgNo &"&RptDate=" & server.URLEncode(dtrptdate)

		'RS("RevDate") = currentUTCDateTime
		dteventTemp=RS("EventDateTime")
		RS("EventDateTime") = Request.Form("txtEvDate") & " " & Request.Form("txtEvTime")
		
		If len(Request.Form("txtClient"))>0 then 
			RS("CustID") = clng(Request.Form("txtClient"))
		else
			RS("CustID") = 0
		end if
		
		SaveSearched_Object("SearchedCRMClient")
		If len(Request.Form("txtCRMClient"))>0 then 
			RS("CRMClient") = trim(Request.Form("txtCRMClient"))
		else
			RS("CRMClient") = ""
		end if
		
		SaveSearched_Object("SearchedCRMRig")
		If len(Request.Form("txtCRMRigID"))>0 then 
			RS("CRMRigID") = trim(Request.Form("txtCRMRigID"))
		else
			RS("CRMRigID") = ""
		end if
		
		'***** (MS HIDDEN) - Uncommented complete If loop  ***** 
		'If len(Request.Form("txtContractor"))>0 then 
		'	RS("ContractorID") = clng(Request.Form("txtContractor"))
		'else
		'	RS("ContractorID") = 0
		'end if
		
		'Reporter
		'------------------------------------------------------------------
		RetrieveFQNValues Request.Form("txtReporter"), Reporter, ReporterUID, "", ""
		
		If ReporterUID = "" Then
			RS("ReporterUID") = NULL
		else
			RS("ReporterUID") = left(ReporterUID,30)
		End if
		RS("Reporter") = left(Reporter,60)
		RIRStatusDestEmail = RS("ReporterUID")
		
		If iHSESev=4 and (RS("SLBInv") <> SLBInv or  RS("SLBConcerned") <> SLBCon) Then
			UpdateMCD=True
			If txtComments <>"" then txtComments=txtComments & "<BR>"	
			txtComments=txtComments & "SLB Involved/SLB Concerned Flag Modified."
		End IF
        
	   ' *****************************************************************
       ' Code to DELETE Time -loss data if all NPT losses are unticked.issue <<2401608>>
       ' *****************************************************************
		If trim(RS("NPT")) <> "" Then 
		OrgNPT = trim(RS("NPT"))
		Else 
		OrgNPT = 0
		End If
		If Request.Form("OptCAffect") = "1" Then StrClientAffected = True Else StrClientAffected = False
		if (Request.Form("LossCat_G3")="" AND Request.Form("LossCat_G2")="" AND Request.Form("LossCat_G1") ="" AND (Request("optSQInvment")=3)) then 
            call initTime(cn) 
            RS("NPT") = 0
	    end if 
	    'START CODE ADDED BY NILESH AT 29-SEP TO FIX NPT BUGS, clear the time loss when classification is  Hazardous  
	    if (bSQSelected and iclass ="3") then
	        call initTime(cn) 
            RS("NPT") = 0
	    end if 
	   'START CODE ADDED BY NILESH AT 2-oct TO FIX NPT BUGS, clear NPT value  when classification is not Accident 
	   if Not isTime(RS)	or iClass <> 1	then 
	        RS("NPT") = 0
	    end if 
	   ' *****************************************************************
		if iClass = 1 and Request.Form("optSQ")="on" then
		if (LockCountSQ <> 0  or  chkSQLockingMgmt()) then
			if not trim(request.form("txtNPT"))="" and Request("optSQInvment") = 2 and iSQSev >= 1 then 
			RS("NPT") = formatquotes(request.form("txtNPT"))
			
					if IsNull(RS("SQPQNPT")) = false then
				SQPQNPTval = cdbl(RS("SQPQNPT"))
			else
				SQPQNPTval = 0
			end if
			if SQPQNPTval=0 then 
			RS("OriginalNpt") = formatquotes(request.form("txtNPT"))
			end if
			
			RS("SQPQNPT") = formatquotes(request.form("txtNPT"))
			
	
			end if
            if trim(request.form("txtNPT"))="" and request.Form("isTimeLossEntered") = "0" then 
			RS("NPT") = NULL
			RS("SQPQNPT") = NULL
			RS("OriginalNpt") = NULL
			end if
			end if
		else
			RS("NPT") = NULL
		end if
        if Request.Form("rdNPT")="1" then
            RS("NPTFlag")=1 
        elseif Request.Form("rdNPT")="0" and iSQSev > 1 then
            RS("NPTFlag")=0
        else
            if request.Form("isTimeLossEntered") = "0" or not iClass = 1 then RS("NPTFlag")=NULL
        end if   
		OrgWIBEvent =  RS("WIBEvent")       
        if Request.Form("rdWIBEventSQ")="1" or Request.Form("rdWIBEventHSE")="1" or trim(Request.Form("rdEventSubCat"))="1" or trim(Request.Form("rdEventSubCat"))="3" then RS("WIBEvent")=1 else RS("WIBEvent")=0
        if Request.Form("rdAccDischarge")="1" then RS("AccDischarge")=1 else RS("AccDischarge")=0
        if Request.Form("rdFireExplosion")="1" then RS("FireExplosion")=1 else RS("FireExplosion")=0
        
		RS("IndRec") = IndRec
		RS("SLBInv") = SLBInv
		RS("SLBConcerned") = SLBCon
		RS("SLBRelated") = SLBRel
		RS("RIRExternal") =External
		
        rsClientAffected=RS("ClientAffected")
		'@@@peter 07/18/2001  default slbinv for SQ
		If Request.Form("optSQ")="on" AND Request.Form("optHSE")<>"on" then RS("SLBInv") = True
		If Request.Form("RegRec") = "1" Then RS("Daylight") = True Else RS("Daylight") = False
		If Request.Form("OptCAffect") = "1" Then RS("ClientAffected") = True Else RS("ClientAffected") = False
		'Class/Severity
		'(Values set above in validation section)
		'5/23/2000 peter
		If not isnull (RS("Class")) and not isnull (RS("Severity")) then
			If ((cint(RS("Class")) <> cint(iClass) ) or (cint(RS("Severity")) <> cint(iSev))) Then 
				UpdateMCD = True
				If MsgID<>101 Then
					MsgID=103
					If(cint(RS("Class")) <> cint(iClass)) Then txtComments=txtComments & "Classification modified From " & rtrim(GetSevClass("C",RS("Class"))) & " to " & rtrim(GetSevClass("C",iClass)) 
					
					If(cint(RS("Severity")) <> cint(iSev)) Then 
						If txtComments <>"" then txtComments=txtComments & "<BR>"	
						txtComments=txtComments & "Severity modified From " & rtrim(GetSevClass("S",RS("Severity"))) & " to " & rtrim(GetSevClass("S",iSev)) 
						If iSev<4 Then ClearApproval iQPID,cn
					end if
				End IF
				'8/28/2001 clear the pers fields where applicable
				
				If cint(iSev) = 0 Then 
					ClearPers cn, "NonAccident"					
				ElseIf (iSev) < 4 Then
					ClearPers cn, "ClearFatality"
				End If
			end if		
		End If
		
		'SWI changes
			RS("Qn1") = SwiQn1 
			RS("Qn2A") = SwiQn2A 
			RS("Qn2B") = SwiQn2B
			RS("Qn3") = SwiQn3

					
	 RS("SPS_L2")= spsl2
	 RS("SPS_L3") = spsl3
	 RS("SPS_L4") = SPS_ID
     RS("SPS_B2") = spsb2
    
	'Saxon Changes 	
		RS("OperationCat") = iif (request.form("cat_sq")=" ", null , request.form("cat_sq"))
		RS("OperationSubCat") = iif (request.form("subcat_sq")=" ", null , request.form("subcat_sq"))
	
	'D&M Change - Clearing D&M Incidents - @Visali 03/31/2010
	    If ShowDM() then
		    If not isnull (RS("Class")) then
			    if (RS("Class")=1 and iClass <> 1) or (RS("Class")=2 and iClass <> 2) then 
			        cn.Execute "DELETE FROM tblRIR_SQDM_Incidents WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "DELETE FROM tblRIR_SQDM_StuckPipedetails WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "Update tblRIR_SQDM_Main Set IWellSiteDesc = '' Where QPID = " & SafeNum(iQPID)
			        cn.Execute "Update tblRIR_SQDM_Main Set ICorrectiveAct = '' Where QPID = " & SafeNum(iQPID)
		        End If
		        if RS("Class")=3 and iClass <> 3 then 
			        cn.Execute "DELETE FROM tblRIR_SQDM_Incidents WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "Update tblRIR_SQDM_Main Set IWellSiteDesc = '' Where QPID = " & SafeNum(iQPID)
			        cn.Execute "Update tblRIR_SQDM_Main Set ICorrectiveAct = '' Where QPID = " & SafeNum(iQPID)
		        End If
		    End if
		End if
		
	    If ShowPF() then
		    If not isnull (RS("Class")) then
			    if (RS("Class")=1 and iClass <> 1) or (RS("Class")=2 and iClass <> 2) then 
			        cn.Execute "DELETE FROM tblRIR_SQPF_Incidents WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "DELETE FROM tblRIR_SQPF_StuckPipedetails WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "Update tblRIR_SQPF_Main Set IWellSiteDesc = '' Where QPID = " & SafeNum(iQPID)
			        cn.Execute "Update tblRIR_SQPF_Main Set ICorrectiveAct = '' Where QPID = " & SafeNum(iQPID)
		        End If
		        if RS("Class")=3 and iClass <> 3 then 
			        cn.Execute "DELETE FROM tblRIR_SQPF_Incidents WHERE QPID = " & SafeNum(iQPID) 
			        cn.Execute "Update tblRIR_SQPF_Main Set IWellSiteDesc = '' Where QPID = " & SafeNum(iQPID)
			        cn.Execute "Update tblRIR_SQPF_Main Set ICorrectiveAct = '' Where QPID = " & SafeNum(iQPID)
		        End If
		    End if
		End if
		
		RS("Class") = iClass
		RS("Exposure")=0
		RS("Severity") = iSev
		RS("HSESeverity") = iHSESev
		RS("SQSeverity") = iSQSev
		
		'Location
		RS("LocType") = split(Request.Form("txtLocation"),":")(0)
		RS("Location") = left(Request.Form("txtLoc"),50)
		
		'Desc
		RS("ShortDesc") = left(Request.Form("txtShortDesc"),50)
		RS("FullDesc") = left (Request.Form("txtFullDesc"),4000)
		
		'HazCat
		RS("HazardCat") = Request.Form("txtHazard")
		If Request.Form("rdEventSafety") = "1" Then RS("EventClassSafety") = True Else RS("EventClassSafety") = False
		If Request.Form("rdEventChoice") = "1" Then RS("EventClassChoice") = True Else RS("EventClassChoice") = False
		If Request.Form("rdPLSSInv") = "1" Then RS("PLSSInv") = True Else RS("PLSSInv") = False	
		If trim(Request.Form("rdEventSubCat")) = "" Then EventSubClassSafety = NULL Else EventSubClassSafety = trim(Request.Form("rdEventSubCat")) 
		RS("EventSubClassSafety") = EventSubClassSafety
		'TS Changes
		IF EnforceSel THEN
		RS("BusinessSegment") = LocBSID
		ELSE
		RS("BusinessSegment") = trim(Request.Form("txtBSegment"))	
		END IF

		'SERVICE QUALITY
		'------------------
		'Service/Product Category

		If trim(Request.Form("txtSPCategory")) <> "" Then RS("SQSPcatID") = trim(Request.Form("txtSPCategory"))
		If trim(Request.Form("txtSubSPCategory")) <> "" Then RS("SQSPSubcatID") = trim(Request.Form("txtSubSPCategory"))
		
		'Failure
		If trim(Request.Form("txtFailure")) <> "" Then RS("SQFcatID") = trim(Request.Form("txtFailure"))
		If trim(Request.Form("txtSubFailure")) <> "" Then RS("SQFSubcatID") = trim(Request.Form("txtSubFailure"))

		'Damage
		'if issaxon(lPL) and request.form ("ShowDamage")="N" Then
		if issaxon(GetSubSubBusinessSegID(lOrgNo)) and request.form ("ShowDamage")="N" Then
			RS("SQDcatID") =0
			RS("SQDSubcatID") =  0
		else 
			If trim(Request.Form("txtDamage")) <> "" Then RS("SQDcatID") = trim(Request.Form("txtDamage"))
			If trim(Request.Form("txtSubDamage")) <>"" Then RS("SQDSubcatID") = trim(Request.Form("txtSubDamage"))
		End if 
		'Cause
		If trim(Request.Form("txtCause")) <>"" Then RS("SQCcatID") = trim(Request.Form("txtCause"))
		If trim(Request.Form("txtSubCause")) <> "" Then RS("SQCSubcatID") = trim(Request.Form("txtSubCause"))
		'Job ID
		IF (trim(Request.Form("txtJobID")) <> "" OR NOT isnull(trim(Request.Form("txtJobID")))) THEN 
			RS("JobID") = trim(Request.Form("txtJobID"))
		END IF	
		'Well Site
		'If trim(Request.Form("optWellSite")) <> "" Then RS("WellSite") = trim(Request.Form("optWellSite"))
		If (bSQDisplayed and (HideSQ=0 and HideWSSQ=0)) and (not isEMS(SQMappingID) or WellSite>0) Then 
		    If Isnull(Request.Form("WellSiteValue")) Then
		        RS("WellSite") = trim(Request.Form("WellSiteValue"))
		     Else
		        If not Isnull(rs("CRMRigID")) and rs("CRMRigID")<> "" Then
                    RigOprVal = getRigOprEnvVal(rs("CRMRigID"))	    
                    RS("WellSite") = RigOprVal
                else
                    RS("WellSite") = 0
                End if
            End if
		Else
		    RS("WellSite") = 0
		End if
		'response.end
		'#2594250-GSS ML Data tab
		'If request.Form("txtIsGSSML")="NO" Then
		'Misc
		If trim(Request.Form("optSQStandard")) <>"" Then 
			If IsNumeric(trim(Request.Form("optSQStandard"))) Then RS("SQDelayHrs") = trim(Request.Form("optSQStandard"))
		Else
			RS("SQDelayHrs") = "0"
		End If

		
		If trim(Request.Form("txtSQNRedone")) <> "" Then
			If isnumeric ( trim(Request.Form("txtSQNRedone"))) Then RS("SQNRedone") = trim(Request.Form("txtSQNRedone"))
		Else
			RS("SQNRedone") = "0"
		End If
		
		If trim(Request.Form("txtPFailure")) <> "" Then 
			If isnumeric ( trim(Request.Form("txtPFailure"))) Then RS("PFailure") = trim(Request.Form("txtPFailure"))
		Else
			RS("PFailure") = "0"
		End If

		'#2594250-GSS ML Data tab
	
		
		If trim(Request.Form("chkReviewed"))="on" then 
			'JES 20010608 We don't know if we want to notify on Reviewed just yet, or not...
			if rs("Reviewed") <> 1 And rs("Reviewed") <> True Then 
				RIRStatus = "THIS REPORT HAS BEEN ACKNOWLEDGED."
				If msgID>101 Then MsgID=122 
				If txtComments <>"" then txtComments=txtComments & "<BR>"
				txtComments=txtComments & "Acknowledged checkbox clicked."
			End IF
			RS("Reviewed") = 1
		Else
			If RS("Reviewed") then
				If msgID>101 Then MsgID=122 
				If txtComments <>"" then txtComments=txtComments & "<BR>"
				txtComments=txtComments & "Acknowledged checkbox un-clicked."
		    End IF
		    RS("Reviewed") = 0
		End if

		If Request.Form("chkClosed")="on" then 
			if rs("closed") <> 1 AND rs("closed") <> True Then 
				If rs("Source")=8 then NotifyIndex=True
				RIRStatus = "THIS REPORT HAS BEEN CLOSED."
				If msgID>101 Then MsgID=121
				If txtComments <>"" then txtComments=txtComments & "<BR>"
				txtComments=txtComments & "Close checkbox clicked."				
			End IF
			RS("Closed") = 1
			RS("Reviewed") = 1
		Else
		    If RS("Closed") then
				If msgID>101 Then MsgID=121 
				If txtComments <>"" then txtComments=txtComments & "<BR>"
				txtComments=txtComments & "Close checkbox un-clicked."
		    End IF
			RS("Closed") =0			
		End if

		'If bNew And lcase(trim(Request.form("txtReporter"))) = "anonymous" Then
		If lcase(trim(Request.form("txtReporter"))) = "anonymous" or instr(1,lcase(trim(Request.form("txtReporter"))),"anonymous") Then
			RS("UpdatedBy")= "Anonymous"
			RS("UpdateUID") = NULL
		Else
	        RS("UpdatedBy") = left(sUName,60)
		    RS("UpdateUID") = left(sUID,30)
		End if				

		'20030124 JES I don't see any other way, but to hardcode this...
		If rs("LossCat_A1") Then
			'Injury was checked before
			If request("LossCat_A1") = "" Then
				'it's unchecked now
				ClearPers cn, "ClearInjury"
			End If
		End If
		
		If rs("LossCat_A2") Then
			'Heatlh was checked before
			If request("LossCat_A2") = "" Then
				'it's unchecked now
				ClearPers cn, "ClearHealth"
			End If
		End If
		
		
		If MsgID=102 then 
			
			dim js_tmpSQLossCats, js_tmpHSELossCats, js_tmpCategory, js_tmpHLoss, HL_js_tmpCategory1,HL_js_tmpCategory2, sqlADTLog
			dim RSADTLog, js_tmpELoss, HL_js_tmpCategory3, HL_js_tmpCategory4, js_tmpSLoss, ReqSafety, RSSafety, RSWIBEvent, ReqWIBEvent
			
			'Health Loss Audit Start
			js_tmpHLoss = split("LossCat_A2 LossCat_A3")
			js_tmpCategory = ""
			HL_js_tmpCategory1 = ""
			HL_js_tmpCategory2 = ""
			for each js_tmpCategory in js_tmpHLoss
				sqlADTLog = "select Description,SectionDesc from tlkpLossSubCategories where ColumnName = '"&js_tmpCategory&"'"
				Set RSADTLog = Server.CreateObject("ADODB.Recordset")		
				RSADTLog.Open sqlADTLog, cn
				If Not RSADTLog.EOF  Then 
					while Not RSADTLog.EOF
						if Rs(js_tmpCategory) = True then 
							if HL_js_tmpCategory1 <> "" then HL_js_tmpCategory1 = HL_js_tmpCategory1&" , " end if
							HL_js_tmpCategory1 = HL_js_tmpCategory1& " "&RSADTLog("Description")
						end if
							 
						if Request(js_tmpCategory) = 1 then 
							if HL_js_tmpCategory2 <> "" then HL_js_tmpCategory2 = HL_js_tmpCategory2&" , " end if
							HL_js_tmpCategory2 = HL_js_tmpCategory2& " "&RSADTLog("Description")
						end if
					RSADTLog.MoveNext
					Wend
				End If	
				RSADTLog.Close
				Set RSADTLog = Nothing	
			Next
			if Trim(HL_js_tmpCategory1) = "" then HL_js_tmpCategory1 = "-" end if
			if Trim(HL_js_tmpCategory2) = "" then HL_js_tmpCategory2 = "-" end if
			if trim(HL_js_tmpCategory1) <> trim(HL_js_tmpCategory2) then
				If txtComments <>"" then txtComments=txtComments & " <br> " end if
				txtComments = txtComments & "RIR Health loss Modified from  " &HL_js_tmpCategory1& " to " &HL_js_tmpCategory2
			end if
			'Health Loss Audit End
			
			
			'Environment Loss Audit Start
			js_tmpELoss = split("LossCat_C1 LossCat_C2 LossCat_C3 LossCat_C4")
			js_tmpCategory = ""
			HL_js_tmpCategory3 = ""
			HL_js_tmpCategory4 = ""
			for each js_tmpCategory in js_tmpELoss
				sqlADTLog = "select Description,SectionDesc from tlkpLossSubCategories where ColumnName = '"&js_tmpCategory&"'"
				Set RSADTLog = Server.CreateObject("ADODB.Recordset")		
				RSADTLog.Open sqlADTLog, cn
				If Not RSADTLog.EOF  Then 
					while Not RSADTLog.EOF
						if Rs(js_tmpCategory) = True then 
							if HL_js_tmpCategory3 <> "" then HL_js_tmpCategory3 = HL_js_tmpCategory3&" , " end if
							HL_js_tmpCategory3 = HL_js_tmpCategory3& " "&RSADTLog("Description")
						end if 
						if Request(js_tmpCategory) = 1 then 
							if HL_js_tmpCategory4 <> "" then HL_js_tmpCategory4 = HL_js_tmpCategory4&" , " end if
							HL_js_tmpCategory4 = HL_js_tmpCategory4& " "&RSADTLog("Description") 
						end if
					RSADTLog.MoveNext
					Wend
				End If	
				RSADTLog.Close
				Set RSADTLog = Nothing		
			Next
			if Trim(HL_js_tmpCategory3) = "" then HL_js_tmpCategory3 = "-" end if
			if Trim(HL_js_tmpCategory4) = "" then HL_js_tmpCategory4 = "-" end if
			if trim(HL_js_tmpCategory3) <> trim(HL_js_tmpCategory4) then
				If txtComments <>"" then txtComments=txtComments & " <BR> " end if
				txtComments = txtComments & "RIR Environment loss Modified From  " &HL_js_tmpCategory3& " to " &HL_js_tmpCategory4
			end if
			'Environment Loss Audit End
			
			
			'Safety Loss Audit Start
			js_tmpSLoss = split("LossCat_A1 LossCat_B1 LossCat_B2 LossCat_E1 LossCat_E2 LossCat_E3 LossCat_E4 LossCat_E5 LossCat_F1 LossCat_F2 LossCat_F3")
			js_tmpCategory = ""
			RSSafety = ""
			ReqSafety = ""
			for each js_tmpCategory in js_tmpSLoss
				sqlADTLog = "select LSC.Description,SectionDesc,LC.Description from tlkpLossSubCategories LSC, tlkpLossCategories LC where LSC.LossCatID = LC.ID and LSC.ColumnName = '"&js_tmpCategory&"'"
				Set RSADTLog = Server.CreateObject("ADODB.Recordset")		
				RSADTLog.Open sqlADTLog, cn
				If Not RSADTLog.EOF  Then 
					while Not RSADTLog.EOF
						if Rs(js_tmpCategory) = True then 
							if RSSafety <> "" then RSSafety = RSSafety&" , " end if
							RSSafety = RSSafety& " "&RSADTLog(2)&"-"&RSADTLog(0)
						end if
							if Request(js_tmpCategory) = 1 then 
								if ReqSafety <> "" then ReqSafety = ReqSafety&" , " end if
								ReqSafety = ReqSafety& " "&RSADTLog(2)&"-"&RSADTLog(0)
							end if
					RSADTLog.MoveNext
					Wend
				End If	
				RSADTLog.Close
				Set RSADTLog = Nothing					
			Next
			if Trim(RSSafety) = "" then RSSafety = "-" end if
			if Trim(ReqSafety) = "" then ReqSafety = "-" end if
			if trim(RSSafety) <> trim(ReqSafety) then
				If txtComments <>"" then txtComments=txtComments & " <BR> " end if
				txtComments = txtComments & "RIR Safety loss Modified From  " &RSSafety& " to " &ReqSafety
			end if	
			'Safety Loss Audit End
			
			
			
			'Well Barrier Integrity Audit START
			If (request("rdWIBEventHSE") <> "" or request("rdWIBEventSQ") <> "") then
			If OrgWIBEvent = 0 then RSWIBEvent = "No" Else RSWIBEvent = "Yes" End if
			If request("rdWIBEventHSE") = "" then 
				If request("rdWIBEventSQ") = 0 then ReqWIBEvent = "No" Else ReqWIBEvent = "Yes" End if
			else
				If request("rdWIBEventHSE") = 0 then ReqWIBEvent = "No" Else ReqWIBEvent = "Yes" End if
			end if
			If (RSWIBEvent  <> ReqWIBEvent) then
			If txtComments <>"" then txtComments=txtComments & " <BR> " end if
				if Trim(ReqWIBEvent) = "" then ReqWIBEvent = "-" end if
				if Trim(RSWIBEvent) = "" then RSWIBEvent = "-" end if
				txtComments = txtComments & "RIR Well Integrity Barrier modified from Well Barrier Element involved "&"  "& RSWIBEvent &"  "& "to" &"  "&ReqWIBEvent
				end if
			end if
			'Well Barrier Integrity Audit End
' NPT Loss Audit Start
                                                inttxtNPT = iif (request("txtNPT")=NULL,0,Request("txtNPT"))
                                                intOrgNPT = OrgNPT
						
						if inttxtNPT <> "" then
                                                If (inttxtNPT) <> (intOrgNPT) Then
                                                                If txtComments <>"" then txtComments=txtComments & " <BR> " end if
                                                                txtComments=txtComments&" NPT from"&"  "& intOrgNPT &"  "&" hours to"&"  "& inttxtNPT&"  "&"hours" 
                                                End if
						End if
                                                ' NPT Loss Audit End

			
		end if	

		
'NEW Category @@@peter 1/31/2001 
		'--------------------------------------------------------------
		CategoryList="''"
		For row = 0 To UBound(Application("a_LossCategories"), 2)
			If Request.Form(Application("a_LossCategories")(0,row)) = "1" Then 
			if (LockCountSQ <> 0  or  chkSQLockingMgmt() or Application("a_LossCategories")(0,row)<>"LossCat_G1") then
				RS(Application("a_LossCategories")(0,row)) = True
			End If	
				CategoryList = CategoryList & ",'" & Application("a_LossCategories")(0,row) & "'"
			Else
			if (LockCountSQ <> 0  or  chkSQLockingMgmt() or Application("a_LossCategories")(0,row)<>"LossCat_G1") then
				RS(Application("a_LossCategories")(0,row)) = false
			End If
			End If
		Next

		If Request.Form("optSQ")="on" then
			RS("ServiceQuality")= true
		else
			RS("ServiceQuality")= false
			If iQPID>0 Then cn.Execute "DELETE FROM tblRIRinvSQ WHERE QPID = " & iQPID 
			If iQPID>0 Then cn.Execute "DELETE FROM tblRIRWellDataSQ WHERE QPID = " & iQPID 
		End If

		If Request.Form("optHSE")="on" then
			RS("HSE")= true
		else
			RS("HSE")= false
		End If

		If Request.Form("projectNO")= "on"  then RS("projectNO")= 1 else RS("projectNO") =  0 
		If Request.Form("projectIDS")="on" then RS("projectIDS")= 2 else RS("projectIDS") =  0 
		'If Request.Form("projectIPS")="on" then RS("projectIPS")= 3 else RS("projectIPS") =  0
		If Request.Form("projectIFS")="on" then RS("projectIFS")= 3 else RS("projectIFS") =  0		
		'If Request.Form("projectISM")="on" then RS("projectISM")= 4 else RS("projectISM") =  0 
		If Request.Form("projectSPM")="on" then RS("projectSPM")= 5 else RS("projectSPM") =  0
		
		if isIPMWCSS(iBSID) and Request.Form("projectNO")<> "on" then RS("projectIDS")= 2 
		'if lPL=130 and Request.Form("projectNO")<> "on" then RS("projectIPS")= 3
		if isIPMIFS(iBSID) and Request.Form("projectNO")<> "on" then RS("projectIFS")= 3 
		'if isIPMPRSS(iBSID) and Request.Form("projectNO")<> "on" then RS("projectISM")= 4
		if isIPMAPS(iBSID) and Request.Form("projectNO")<> "on" then RS("projectSPM")= 5

		If Request.Form("projectNO")= "on" then
				If iQPID>0 Then  
						'if not isIPMSeg(GetProductLine(lOrgNo)) then '- Added for deleting the data from the tblRIRipm
						if not isIPMSeg(SQMappingID) then '- Added for deleting the data from the tblRIRipm
							cn.Execute "DELETE FROM tblRIRipm WHERE QPID = " & iQPID 
						end if
				end if 
		end if
     
		'Removed below as per: ENH157930
		'Change for P&AM report
		'If bSQSelected Then
		'	if not (iClass <3 and frmRSsqlnvInvment=2 and Request.Form("rdWIBEventSQ")="1" ) then
		'		cn.Execute "DELETE FROM tblIPM_PSEdata WHERE QPID = " & iQPID 
		'	end if
		'End If
		
		'Change for P&AM report
		'If bHSESelected Then
		'	if not (iClass <3 and frmSLBInvment=1 and Request.Form("LossCat_C1")= 1) then
		'		cn.Execute "DELETE FROM tblIPM_PSEdata WHERE QPID = " & iQPID 
		'	end if
		'End If
		
		
		 If Request.Form("optPTECInv")="1" then
			RS("PTEC")= true
		elseif Request.Form("optPTECInv")="0" then 
			RS("PTEC")= false
		End If
	 
		If Request.Form("optROPInv")="1" then
			RS("ROP")= true
		elseif Request.Form("optROPInv")="0" then 
			RS("ROP")= false
		End If
	 
       If Request.Form("optSegInv")="1" then '***** (MS HIDDEN) - Commented complete If loop section  ***** 
          RS("IsSegmentInv")= True
       Else 
          RS("IsSegmentInv")=False
       End If    
       
	   if Request.Form("optgot")=1 THEN
	     RS("GRCInv") = 1
		 Else
		 RS("GRCInv") = 0
	   End If
	   
	   if Request.Form("optgot")=1 THEN
	   RS("TCCInvolved") = 1
	   End If
	   
        '--Removed--2713511ST(Start)
		'If Request.Form("optHSE")="on" then '***** (MS HIDDEN) - Commented If loop  ***** 
		'	'Contractors Inv			
		'	If Request.Form("optContractorInv")="1" then
		'		RS("ContractorID")= 1
		'	else
		'		RS("ContractorID")= 0
		'		If iQPID>0 Then cn.Execute "DELETE FROM tblRIRContractors WHERE QPID = " & iQPID 
		'	End If
		'End If
		'--Removed--2713511ST(End)
		
        'TCCInvolved '***** (MS HIDDEN) - Uncommented complete If loop  ***** 
		'If Request.Form("optTCCInv")="1" then
		'	RS("TCCInvolved")= true
		'else
		'	RS("TCCInvolved")= false
		'	If iQPID>0 Then 
		'	    cn.Execute "DELETE FROM tblRIR_SQTCCDetails WHERE QPID = " & iQPID 
		'	    cn.Execute "DELETE FROM tblRIR_SQTCCCatDetails WHERE QPID = " & iQPID 
		'	    cn.Execute "DELETE FROM tblRIR_SQTCCLocation WHERE QPID = " & iQPID 
		'	end if    
		'End If
		
		
		'HiPo Changes 	
		if RS("Class") = 3 or (RS("Class") <> 3 and RS("ServiceQuality") and not RS("HSE")) then 
			cn.Execute "Update tblRIRRisk Set FailSafe = 0, FailLucky = 0 Where QPID = " & SafeNum(iQPID)
		End If	
		
		RS("Department") = cint(Request.Form("lstDepartment"))
		RS("AccountUnit")=trim(Request.Form("txtAccountUnit"))
		'Oct 30, 2007 - Sreedhar Vadla
		'If UpdateMCD Then RS("MajorChangeDate") = Now()	
		RS.Update
		cn.execute("Update tblRIRp1 set RevDate=" &"'"& currentUTCDateTime & "'" & "where QID="& iQPID)   'Code to update RevDate using update script and solve driver error.
		
		If not UpdateMCD Then
		cn.execute("Update tblRIRp1 set MajorChangeDate=" &"'"& Now() & "'" & "where QID="& iQPID)
		End IF
		
		If iQPID=0 Then
			Set rs2=cn.execute("Select QID from tblRIRP1 with (NOLOCK) Where OrgNo="&lOrgNo&" and RptDate='"&dtRptDate&"'")
			If Not rs2.EOF Then iQPID=rs2("QID")
	    	rs2.Close
		End IF

		Set RSCont = Server.CreateObject("ADODB.Recordset")
        RSCont.LockType = 3	       
        cn.execute("Delete from tblRIRContractors Where QPID='"&SafeNum(iQPID)&"'")
        cn.execute("Update tblRIRp1 set ContractorID= 0 where QID="& iQPID) 
        sBig = Trim(Request.Form("chkTPSupplier"))
        'alert(sBig)
        sSerID=0
        sRisk=""
        sMode=""
       
        If instr(sBig,",")	Then           
            Slist = split(sBig,",")             
            For iCtr = lbound(Slist) to ubound(Slist)	                        
	            If instr(Slist(iCtr),":")	Then				
		            sContVal=split(Slist(iCtr),":")
		            sContID=sContVal(0)		            		            			            
		            sSerID=sContVal(1)
		            sRisk=sContVal(2)
		            sMode=sContVal(3)
		            
		            If sContID>0 Then
			            sTemp = "SELECT * FROM tblRIRContractors WHERE QPID='"& SafeNum(iQPID) & "' AND SeqID='"&SafeNum(iCtr+1)&"'" 
			            RSCont.Open sTemp, cn
			            If RSCont.EOF then 
				            RSCont.AddNew
				            RSCont("OrgNo")= lOrgNo
				            RSCont("RptDate")= dtRptDate
				            RSCont("QPID") = iQPID
				            RSCont("SeqID") = iCtr+1			
			            End if
			            RSCont("ContractorID")=sContID
			            if sSerID <> "" Then RSCont("ServiceID")=sSerID Else RSCont("ServiceID") = 0 
			            RSCont("PreQualified")=0
			            RSCont("Contract")=0			            
			            If trim(sRisk)="" Or isNull(sRisk) Then 
                            sRisk=" "
                        Else
                            if len(trim(sRisk))>0 then sRisk=Left(trim(sRisk),1)
                        End If
                        RSCont("RiskRate")=sRisk
                        If trim(sMode)="" Or isNull(sMode) Then 
                            sMode=" "
                        Else
                            if len(trim(sMode))>0 then sMode=Right(trim(sMode),1)
                        End If			            			          			            
			            RSCont("Mode")=sMode
			            RSCont("RevDate")= Date()
			            RSCont("RevBy")=Session("UID")
			            
			            RSCont.Update
			          
			            RSCont.Close			            
			            cn.execute("Update tblRIRp1 set ContractorID= 1 where QID="& iQPID)
		            End IF
	            End If	
            Next
        Else
            If instr(sBig,":")	Then
                sContVal=split(sBig,":")
                sContID=sContVal(0)		            		            			            
                sSerID=sContVal(1)
                sRisk=sContVal(2)
                sMode=sContVal(3)
                
                If sContID>0 Then
	                sTemp = "SELECT * FROM tblRIRContractors WHERE QPID='"& SafeNum(iQPID) & "' AND SeqID='"&SafeNum(iCtr+1)&"'"
    	            
	                RSCont.Open sTemp, cn
	               
	                If RSCont.EOF then 
		                RSCont.AddNew
		                RSCont("OrgNo")= lOrgNo
		                RSCont("RptDate")= dtRptDate
		                RSCont("QPID") = iQPID
		                RSCont("SeqID") = iCtr+1				
	                End if
	                RSCont("ContractorID")=sContID
	                if sSerID <> "" Then RSCont("ServiceID")=sSerID Else RSCont("ServiceID") = 0 
	                RSCont("PreQualified")=0
	                RSCont("Contract")=0	                
		            If trim(sRisk)="" Or isNull(sRisk) Then 
                        sRisk=" "
                    Else
                        if len(trim(sRisk))>0 then sRisk=Left(trim(sRisk),1)
                    End If
                    RSCont("RiskRate")=sRisk
                    If trim(sMode)="" Or isNull(sMode) Then 
                        sMode=" "
                    Else
                        if len(trim(sMode))>0 then sMode=Right(trim(sMode),1)
                    End If			            			          			            
		            RSCont("Mode")=sMode
	                RSCont("RevDate")= Date()
	                RSCont("RevBy")=Session("UID")
    	            
	                RSCont.Update
	               
	                RSCont.Close			            
	                cn.execute("Update tblRIRp1 set ContractorID= 1 where QID="& iQPID)
                End IF
            End If
            Set RSCont = Nothing
                                   
        End If    
        						

		If Request.Form("optSegInv")="1" then '***** (MS HIDDEN) - Commented complete If loop section  ***** 
		Else
			' for Multi Segment : No
     		fncRemoveMultiSeg(iQPID)
		End If   
		
		
		IF Request.Form("rdEventSafety")="1" AND ((trim(Request.Form("rdEventSubCat"))="4" OR trim(Request.Form("rdEventSubCat"))="5" OR trim(Request.Form("rdEventSubCat"))="6" OR trim(Request.Form("rdEventSubCat"))="8" OR trim(Request.Form("rdEventSubCat"))="9")) then 
			fncRemoveIOGPTabDetails(iQPID)
		End IF 
		IF iClass <> 1 and (Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(varHideLegacyIogp))) then 
			fncRemoveIOGPUpdatedTabDetails(iQPID)
		End IF 
		
		'11/15/98 Delete in case they unclicked an option or not an accident
		if Not isPers(RS)	or iClass <> 1	then initPers(cn)
		if Not isAuto(RS)	or iClass <> 1	then initAuto(cn)
		if Not isEnv(RS)	or iClass <> 1	then initEnv(cn)
		if Not isOther(RS)	or iClass <> 1	then initOth(cn)
		if Not ((isAssets(RS) OR bSQSelected)	AND iClass = 1)	then initAssets(cn)
		if Not isInfo(RS)	or iClass <> 1	then initInfo(cn)
		if Not isTime(RS)	or iClass <> 1	then initTime(cn)
		if iClass<>3 or not bHSEselected then initStop(cn)
		if iClass<>3 or not bHSEselected then initHOC(cn)
		If Not bSQSelected Then initQStop(cn)
		If Not bSQSelected Then initInvSeg(cn)	' #2588673 '***** (MS HIDDEN) - Commented line ***** 
        
    if (not Request.Form("rdWIBEventSQ")="1" and not Request.Form("rdWIBEventHSE")="1") and (not trim(Request.Form("rdEventSubCat"))="1" and not trim(Request.Form("rdEventSubCat"))="3") then
                cn.execute "DELETE FROM tblRIRWellBarriers WHERE QPID="& SafeNum(iQPID)
        end if
				
		' SLIM Changes
		'Set RSslim = Server.CreateObject("ADODB.Recordset")	
		
		'set RSslim = cn.execute ("SELECT * FROM tblRIR_SLIMdata WHERE QPID = " & iQPID )
		'if RSslim.eof then
		'	cn.execute ("INSERT INTO tblRIR_SLIMdata values (" & iQPID & ",'SLIM-INIT'," & SlimCntrl & "," &SlimTech& "," &SlimProc& "," &SlimCompet& "," &SlimBehav& "," & DATE() &",'" & Session("UID") & "')")
		'	LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","RIRdsp.asp",1254,"Entered"
		'else
		'	cn.execute ("update tblRIR_SLIMdata set  ControlDomain=" &SlimCntrl& ", TechDomain =" & SlimTech& ",ProcDomain=" &SlimProc& ",CompetDomain=" & SlimCompet& ",BehaviorDomain="& SlimBehav& " where QPID = " & iQPID & " and SlimType = 'SLIM-INIT' ")
		'	LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","RIRdsp.asp",1254,"updated"
		'end if
		
		'RSslim.close
		'set RSslim = nothing

		'IF iClass = 3 or not bSQSelected  then initSlim(cn)

		
		' SLIM Changes end 
		
		
		'Sreedhar Clear Investigation - Data Gathering
		If iQPID>0 Then DeleteDGTreeInfo iQPID,cn
		
		
		'20010606 JES Notification of closure/review
		'-------------------------------------------
		if RIRStatus <> ""  Then
			'We have something to send, and someone to send it to.
			if (not IsNull(RIRStatusDestEmail)) and RIRStatusDestEmail <> "" Then 
				RIRStatusDestEmail = RIRStatusDestEmail & "@slb.com"
				NotifyUser RIRStatusDestEmail,RIRStatus & vbCRLF & vbCRLF,""
			End IF
			If NotifyIndex then NotifyIndexUser			
		End If
		
		RS.Close
		Set RS = Nothing
		
		'Oct 30, 2007 - Sreedhar Vadla
		'This cause the RIR to consildiate Safetynet data twise by running Update Trigger.
		'So we disabling this event.
		
		'If UpdateMCD Then
		'	UpdateMajorChange iQPID,"R",cn
		'End if
		
		if (RSsqlnvInvment =2  ) then
	     rsrelated="Related Yes"
		 rsexternal="External Yes"
		End if
		
			if (RSsqlnvInvment =1  ) then
	     rsrelated="Related Yes"
		 rsexternal="External No"
		End if
		
		if (RSsqlnvInvment =3  ) then
	     rsrelated="Related No"
		 rsexternal="External -"
		End if
		
		
		if (frmRSsqlnvInvment =2  ) then
	     frmrelated="Related Yes"
		 frmexternal="External Yes"
		End if
		
			if (frmRSsqlnvInvment =1  ) then
	     frmrelated="Related Yes"
		 frmexternal="External No"
		End if
		
		if (frmRSsqlnvInvment =3  ) then
	     frmrelated="Related No"
		 frmexternal="External -"
		End if
		
		 
		 if (frmRSsqlnvInvment  <> RSsqlnvInvment  ) and (MsgID=102) and (RSsqlnvInvment <> "" and frmRSsqlnvInvment <> "" ) then
		 if txtComments = "" then
		txtComments = " Activity/Process/Service modified from"&"  "& rsrelated &"  "& "to" &"  "& frmrelated &"  "& "and"  &"  "& rsexternal &"  "& "to" &"  "& frmexternal
		else
		txtComments=txtComments&"  "& "<BR> Activity/Process/Service modified from"&"  "& rsrelated &"  "& "to" &"  "& frmrelated &"  "& "and"  &"  "& rsexternal &"  "& "to" &"  "& frmexternal
		end if
	     end if 

		 if (RSSLBInvment =1  ) then
	      StrEvent="SLB Involved/Industry Recognized"
		 End if
		
		 if (RSSLBInvment =2  ) then
	      StrEvent="SLB Involved/Non Industry Recognized"
		 End if
		
		 if (RSSLBInvment =3  ) then
	      StrEvent="SLB Non Involved/Advisory"
		 End if
		
		 if (RSSLBInvment =4  ) then
	      StrEvent="SLB Non Involved/Informative"
		 End if
		 
		 if (frmSLBInvment =1  ) then
	      frmStrEvent="SLB Involved/Industry Recognized"
		 End if
		 if (frmSLBInvment =2  ) then
	      frmStrEvent="SLB Involved/Non Industry Recognized"
		 End if
		 if (frmSLBInvment =3  ) then
	      frmStrEvent="SLB Non Involved/Advisory"
		 End if
		 if (frmSLBInvment =4  ) then
	      frmStrEvent="SLB Non Involved/Informative"
		 End if
		
		  
		  if (frmSLBInvment  <> RSSLBInvment  ) and (MsgID=102) and (RSSLBInvment <> "" and  frmSLBInvment <> "") then
		  if txtComments = "" then
		 txtComments = "RIR Activity Type modified from"&"  "& StrEvent &"  "& "to" &"  "& frmStrEvent
		 else
		 txtComments=txtComments&"  "& "<BR> RIR Activity Type modified from"&"  "& StrEvent &"  "& "to" &"  "& frmStrEvent
		 end if
	     end if 
		 
		
		if RSsqClientAffect ="True" then
		rstrClientAffected="Yes"
		else
		rstrClientAffected="No"
		end if
		
		if frmCAffect ="1" then
		frmCAffect="True"
		frmstrClientAffected="Yes"
		else
		frmCAffect="False"
		frmstrClientAffected="No"
		end if
	
		
		 
		if ( frmCAffect <> RSsqClientAffect) and (MsgID=102) and (rsClientAffected <> "" )  then
		if txtComments = "" then
		txtComments = "Client Affected modified from"&"  "& rstrClientAffected &"  "& "to" &"  "& frmstrClientAffected 
       else
txtComments=txtComments&"  "& "<BR> Client Affected modified from"&"  "& rstrClientAffected &"  "& "to" &"  "& frmstrClientAffected 
     end if		
	end if 

if (MsgID=102) then
	dim evmonth,evdays,evyear,frmevmonth,frmevdays,frmevyear,esvmonth,frmesvmonth


	evmonth=MonthName(Month(dteventTemp))
	evdays=Day(dteventTemp)
	evyear=Year(dteventTemp)

	

   frmevmonth=MonthName(Month(dtfrmTemp))
	frmevdays=Day(dtfrmTemp)
	frmevyear=Year(dtfrmTemp)
	esvmonth=evmonth&"  "&evdays&"th"&"  "&evyear
	frmesvmonth=frmevmonth&"  "&frmevdays&"th"&"  "&frmevyear


	if  (dteventTemp <> "" and dtfrmTemp <> "" ) then
		if (FormatDateTime( dteventTemp, 2) <> FormatDateTime( dtfrmTemp, 2)) and (MsgID=102)  then
		if txtComments = "" then
		txtComments = "RIR Event date modified from"&"  "& esvmonth &"  "& "to" &"  "& frmesvmonth  
else
txtComments=txtComments&"  "& "<BR> RIR Event date modified from"&"  "& esvmonth &"  "& "to" &"  "& frmesvmonth  
end if		
	end if 
	end if 
end if 
		LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","RIRDsp2.asp",MsgID,txtComments
		
		if iClass = 1 and Request.Form("optSQ")="on" and iQPID > 0 and Request("optSQInvment") = 2 and iSQSev >= 1 and request.Form("isTimeLossEntered") = "0" then
			if (LockCountSQ <> 0  or  chkSQLockingMgmt()) then
			cn.execute "delete from tblRIRCosts where LossCatID = 7 and QPID = " & iQPID & " and CostCatID in (30,33)"
			end if
			if not trim(request.form("txtNPT_LossCat_G1"))="" then
				set rs = cn.execute ("select * from tblRIRCosts where QPID = " & iQPID & " and CostCatID = 30")
				if rs.eof then
					cn.execute ("insert into tblRIRCosts values (" & iQPID & ",7,30," & formatquotes(request.form("txtNPT_LossCat_G1")) & ")")
				else
					cn.execute ("update tblRIRCosts set Cost = " & formatquotes(request.form("txtNPT_LossCat_G1")) & " where QPID = " & iQPID & " and CostCatID = 30")
				end if
			end if
			if not trim(request.form("txtNPT_LossCat_G2"))="" then
				set rs = cn.execute ("select * from tblRIRCosts where QPID = " & iQPID & " and CostCatID = 33")
				if rs.eof then
					cn.execute ("insert into tblRIRCosts values (" & iQPID & ",7,33," & formatquotes(request.form("txtNPT_LossCat_G2")) & ")")
				else
					cn.execute ("update tblRIRCosts set Cost = " & formatquotes(request.form("txtNPT_LossCat_G2")) & " where QPID = " & iQPID & " and CostCatID = 33")
				end if
			end if
		end if

		'2/21/2001 peter
		if iClass=1 then 
			cn.Execute "DELETE FROM tblRIRStop WHERE QPID = " & SafeNum(iQPID) 
		Else
			cn.Execute "DELETE FROM tblRIRCosts WHERE QPID = " & SafeNum(iQPID) 
		End If
		
		IF EnforceSel THEN
		BSegID = LocBSID
		ELSE
		BSegID = trim(Request.Form("txtBSegment"))	
		END IF
		
		
		IF MSgID=121 AND (BSegID=252 OR BSegID=9129 OR BSegID=9174 OR BSegID=9183) THEN 
		    ConfMsg=NotifyWISRiteNet(iQPID,"Closed")
		    LogAuditTrail iQPID,lorgNo,dtRptdate,"R","RIRDsp2.asp",163,"Closed :" & ConfMsg
		END IF    
		If MSgID=121 and trim(Request.Form("txtAccountUnit"))<>"" Then
		    ConfMsg=NotifyRiteNet(iQPID,"Closed")
		    LogAuditTrail iQPID,lorgNo,dtRptdate,"R","RIRDsp2.asp",163,"Closed :" & ConfMsg
		End if
		
		pTemp="RIRdsp.asp" & sKey & "&Msg=" & Server.URLEncode("Data Successfully Saved")

		ProtectDoc=False
		If iHSESev=4 Then
			If SLBInv or SLBCon Then
				ProtectDoc=True
			ElseIf Request("ProtectDoc")<>"" Then
				cn.execute("DELETE FROM tblAccessList where QID="&SafeNum(iQPID)&" and RptType='R'")
				LogAuditTrail  iQPID,lorgNo,dtRptdate,"R","RIRDsp2.asp",709,"User Selected Un-Protect Document"
			End IF
		End IF
		 
		'Auto protection is Right now Disabled.
		if (bNew or UpdateMCD) and ProtectDoc then
			cn.execute "spR_AutoProtect " & SafeNum(lOrgNo) & ",'" & dtRptDate & "'" 	
			pTemp = pTemp & "&APMsg=1" 
		End IF
		
		'Redirect...
	
		Response.Redirect (pTemp)
	End if
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-At LineNo1573 QID="&SafeNum(iQPID)
		End If

'Display Errors....
'------------------------------------------------------------------%>
<HTML>
	<head>
		<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	</head>
	<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>
		<TABLE align=left border=0 cellPadding=0 cellSpacing=0 width=100%>
			<TR>
				<TD width=5>&nbsp;</TD>
				<TD valign=top>
					<p class=title id=styleMedium><div align=center>Updating RIR</div></p>
					<HR>
					<span class=urgent id=styleMedium><%=sErrors%></span>
					<P>&nbsp;</P>
					<I><B>Hit the back button on your browser to correct these problems...</B></I>
				</TD>
			</TR>
		</TABLE>
	</BODY>
</HTML>
<%
Function CheckMRC()
	Dim msSQL,iCnt,mRS,iHold,sTemp,mConn
	On Error Resume Next
	
	CheckMRC = true
	msSQL = "SELECT * FROM tblRIRInvSQ WHERE QPID = " & SafeNum(iQPID) 
	SET mConn = getNewCn()
	SET mRS = Server.CreateObject("ADODB.Recordset")
	mRS.Open msSQL,mConn
	iHold = 0
	If mRS.EOF AND mRS.BOF Then Exit Function
	
	For iCnt = 1 To 10		
		iHold = iHold + cint(mRS("MRC" & iCnt))		
	Next
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun CheckMRC QID="&SafeNum(iQPID)
		End If
	If iHold = -1 Then
		CheckMRC = false
	End If
	mRS.Close
	Set mRS = nothing
	mConn.Close
	Set mConn = nothing	
End Function

' slim 
Function GetCFlag()
	On Error Resume Next
	Dim rs,sql
	GetCFlag = 0
	sql = "Select Completed from tblRIRInvDetails With (NOLOCK) where QPID = "&SafeNum(QPID)
	Set rs = cn.execute(sql)
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun GetCFlag QID="&SafeNum(iQPID)
		End If
	if not rs.eof then
		GetCFlag = rs("Completed")
	end if	
End Function


Sub InitPers(cn)
On Error Resume Next
	'6/30/2000 peter removed when I opened up Personnel
	InitCosts cn, 6
		If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitPers QID="&SafeNum(QPID)
	End If
end sub

Sub InitAuto(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRAuto WHERE QPID = " & SafeNum(iQPID) )
	InitCosts cn, 2 
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitAuto QID="&SafeNum(QPID)
	End If
end sub

Sub InitEnv(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIREnv WHERE QPID = " & SafeNum(iQPID) )
	InitCosts cn, 3
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitEnv QID="&SafeNum(QPID)
	End If
end sub

Sub InitOth(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIROth WHERE QPID = " & SafeNum(iQPID) )
	InitCosts cn, 5 
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitOth QID="&SafeNum(QPID)
	End If
end sub

Sub InitAssets(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRAssets WHERE QPID = " & SafeNum(iQPID) )
	InitCosts cn, 1
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitAssets QID="&SafeNum(QPID)
	End If
end sub

Sub InitInfo(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRInfo WHERE QPID = " & SafeNum(iQPID) )
	InitCosts cn, 4 
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitInfo QID="&SafeNum(QPID)
	End If
end sub

Sub InitTime(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRTime WHERE QPID = " & SafeNum(iQPID) )
	
	InitCosts cn, 7
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitTime QID="&SafeNum(QPID)
	End If
end sub

Sub InitCosts(cn, iType)
On Error Resume Next
if (LockCountSQ <> 0  or  chkSQLockingMgmt()) then
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRCosts  WHERE QPID = " & SafeNum(iQPID) &" AND LossCatID=" & SafeNum(iType))
	end if
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitCosts QID="&SafeNum(QPID)
	End If
End Sub

Sub InitStop(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRSTOP  WHERE QPID = " & SafeNum(iQPID) )
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitStop QID="&SafeNum(QPID)
	End If
End Sub

Sub InitHOC(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRSwacoHOCMain  WHERE QPID = " & SafeNum(iQPID) )
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIRSwacoHOCDetails  WHERE QPID = " & SafeNum(iQPID) )
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub InitHOC QID="&SafeNum(QPID)
	End If
End Sub

Sub initQStop(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIR_sqQStop  WHERE QPID = " & SafeNum(iQPID) )
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub initQStop QID="&SafeNum(QPID)
	End If
End Sub

Sub initInvSeg(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIR_SQInvSegment WHERE QPID = " & SafeNum(iQPID) )
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub initInvSeg QID="&SafeNum(QPID)
	End If
End Sub

Sub initSlim(cn)
On Error Resume Next
	If iQPID>0 Then cn.Execute ("DELETE FROM tblRIR_SLIMdata  WHERE QPID = " & SafeNum(iQPID) )
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub initSlim QID="&SafeNum(QPID)
	End If
End Sub

Sub ClearApproval(QID,cn)
	dim SQL
	On Error Resume Next
	SQL="Delete From tblRIRStatus Where sType=1 and QPID='"&SafeNum(QID)&"'"
	cn.execute(SQL)
	If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub ClearApproval QID="&SafeNum(QPID)
	End If
End Sub

Function getRigOprEnvVal(CRMID)
	Dim sql,crmRS,crmCn,Temp
	On Error Resume Next
	'getCRMClientName="NO CLIENT"
	If CRMID<>"" Then
	    'If Left(CRMID,2)="P:" Then Temp=Mid(CRMID,3) else Temp=CRMID
	    SQL="Select Case (Rtrim(LTrim(R.RigOprEnv))) WHEN 'Offshore - Shallow' then 1 WHEN 'Offshore - Deepwater' then 1 WHEN 'Land' then 2 "
        SQL= SQL & "WHEN 'Swamp/Inland Waters' then 2 ELSE 0 END as RigOprEnvVal from tblCRMRigs R where RigId ='"&CRMID&"'"
        Set crmCn=getNewCn()
		set crmRS=crmcn.execute(sql)
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun getRigOprEnvVal QID="&SafeNum(QPID)
		End If
		If not crmRS.EOF Then 
		    getRigOprEnvVal=SafeDisplay(crmRS("RigOprEnvVal"))
		Else
		    getRigOprEnvVal=0
		End if
		crmRS.close
		Set crmRS=nothing
		crmCn.close
		Set crmCn=nothing
	End IF	
End Function

Sub ClearPers(cn,clearMode)
	Dim SQL
	On Error Resume Next
	
	Select Case lcase(clearMode)
		Case "nonaccident"
			sql = "UPDATE tblRIRPers SET Fatality=0, outcome = NULL, ReducedWorkDays=0, DaysLost=0, InjuryPart=NULL, InjuryType=NULL  WHERE QPID = " & SafeNum(iQPID) 
		Case "clearfatality"
			sql = "UPDATE tblRIRPers SET Fatality=0  WHERE QPID = " & iQPID 
		Case "clearhealth"
			sql = "DELETE FROM tblRIRPers WHERE InjuryType IN (SELECT CODE FROM tlkpInjuryCategories WHERE Type = 'H') AND QPID = " & SafeNum(iQPID) 
		Case "clearinjury"
			sql = "DELETE FROM tblRIRPers WHERE InjuryType IN (SELECT CODE FROM tlkpInjuryCategories WHERE Type = 'I') AND QPID = " & SafeNum(iQPID) 
	End Select
	
	If Sql = "" Then Err.Raise 1, "Invalid ClearMode value ('" & clearMode & "')"
	If iQPID>0 Then cn.Execute (sql)
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun ClearPers QID="&SafeNum(QPID)
		End If
end sub

Sub NotifyIndexUser()
dim Email,URL,EmailAddress,EmailSubj,RptNo,rs2
	On Error Resume Next
	'Email Address Extraction is changed on 08042005 @svadla
	' Start,changed by Nilesh for Swift 2417393 
	EmailAddress= fncD_Configuration("IndexEmail_Test")
	If ISProd() then EmailAddress= fncD_Configuration("IndexEmail_Prod")
	If EmailAddress="" Then EmailAddress = "svadla@slb.com"
		
	RptNo=GetReportNumber(RS("RptDate"))
	EmailSubj="QUEST Close Notification INDEX:"&RS("ForeignID")&" Report Closing Notification"
	URL= GetQUESTServer() & "SR.asp?Q=" & TRIM(rs("QID")) & vbCRLF & vbCRLF
	Email="<XML>"&vbCRLF
	EMail=Email & "<QuestNotification>"&vbCRLF
	EMail=Email & "<IndexTicket>" & RS("ForeignID") & "</IndexTicket>"&vbCRLF
	EMail=Email & "<QuestTicket>" & RptNo& "</QuestTicket>"&vbCRLF
	EMail=Email & "<QuestID>" & RS("QID") & "</QuestID>"&vbCRLF
	EMail=Email & "<Status>1</Status>"&vbCRLF
	EMail=Email & "<Description>" & server.URLEncode(RS("FullDesc")) & "</Description>"&vbCRLF
	EMail=Email & "<QuestURL>" & server.URLEncode(URL) & "</QuestURL>"&vbCRLF
	EMail=Email & "</QuestNotification>"&vbCRLF
	EMail=Email & "</XML>"&vbCRLF
	QueueMail EmailAddress,EmailSubj,EMail,"","RIR"
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun NotifyIndexUser QID="&SafeNum(QPID)
		End If
End Sub

Sub NotifyUser(EmailAddress,PreMsg,PostMsg)
	Dim EmailBody, EmailSubj, Heading,rEventKey
	Dim CRMRigName, BusSeg, ServerLink, sCategories
	Dim rs2, sTemp, CRMClientName,rTemp
	On Error Resume Next
	
	rTemp=""
	set rs2=server.CreateObject("ADOdb.recordset")
	
	Heading = ""
	
	if bHSESelected then
	    rTemp = "HSE - " 
	elseif bSQSelected then
	    rTemp = "SQ - "     
	elseif bHSESelected and bSQSelected then
	    rTemp = "HSE/SQ - " 
	end if    
	Heading = Heading & rTemp
		
	If rs("Severity") <> 0 Then
		rs2.Open "SELECT SeverityDesc FROM tlkpRIRSeverity WHERE SeverityID=" & SafeNum(rs("Severity")),cn
		if not rs2.EOF Then
			Heading = Heading & trim(rs2("SeverityDesc")) & " "
		end if
		rs2.Close
	Else
	    If iClass=2 Then Heading = Heading & "Near Accident/Incident" else Heading = Heading & "Hazardous Situation"
	End if
	EmailSubj = Heading

	if Not rs("SLBInv")  Then
		EmailSubj = EmailSubj & " [Non-SLB] "
	End If
	
	EmailSubj = EmailSubj & ": " + rs("ShortDesc")	
				
	CRMClientName = "(None)"

	if not IsNull(rs("CRMClient")) and rs("CRMClient") <> "" THen 
		CRMClientName = getCRMClientName(rs("CRMClient"))	
	End if

	CRMRigName = "(None)"

	if not IsNull(rs("CRMRigID")) and rs("CRMRigID") <> "" THen 
		CRMRigName = getCRMRigName(rs("CRMRigID"))
	End if

	EmailBody =             "RIR number        : " & GetReportNumber(RS("RptDate")) & vbCRLF 
	EmailBody = EmailBody & "Classification    : " & Heading & vbCRLF
	EmailBody = EmailBody & "Event Date        : " & RS("RptDate") & vbCRLF
	EmailBody = EmailBody & "Reporter          : " & RS("Reporter") & vbCRLF
	EmailBody = EmailBody & "Location          : " & getLongName(RS("OrgNo")) & vbCRLF
	EmailBody = EmailBody & "CRM Rig Name      : " & CRMRigName & vbCRLF
	EmailBody = EmailBody & "CRM Client        : " & CRMClientName & vbCRLF
	EmailBody = EmailBody & "SLB Involved      : " & iif(rs("SLBInv"),"Yes","No") & vbCRLF
	If RS("SLBInv") Then
		EmailBody = EmailBody & "Ind. Recognized   : " & iif(rs("IndRec"),"Yes","No") & vbCRLF
		EmailBody = EmailBody & "Concerned         : " & "--" & vbCRLF
	Else
		EmailBody = EmailBody & "Ind. Recognized   : " & "--" & vbCRLF
		EmailBody = EmailBody & "Concerned         : " & iif(rs("SLBConcerned"),"Yes","No") & vbCRLF
	End IF
	
	EmailBody = EmailBody & "Service Quality   : " & iif(rs("ServiceQuality"),"Yes","No") & vbCRLF
	EmailBody = EmailBody & "HSE               : " & iif(rs("HSE"),"Yes","No") & vbCRLF
	
	BusSeg = "(none)"
		rs2.Open "select * from tlkpBusinessSegments where BusinessSegmentID='" & SafeNum(trim(rs("BusinessSegment")))&"'",cn
		if not rs2.EOF Then BusSeg = rs2("BusinessSegmentDesc")
		rs2.Close
	EmailBody = EmailBody & "Sub-Segment       : " & BusSeg & vbCRLF
	
	sCategories = ""
		rs2.Open "SELECT LC.Description + ' (' + LSC.Description + ')' AS Cat FROM tlkpLossCategories LC INNER JOIN tlkpLossSubCategories LSC ON LC.ID = LSC.LossCatID WHERE ColumnName IN (" & CategoryList & ") ORDER BY LC.Description, LSC.Description",cn
		do while not rs2.EOF
			sCategories = sCategories & rs2("Cat") & "  "
			rs2.MoveNext
		loop
		rs2.Close
	EmailBody = EmailBody & "Selected Category : " & sCategories & vbCRLF
	EmailBody = EmailBody & "Short Description : " & rs("ShortDesc") & vbCRLF
	EmailBody = EmailBody & "Full Description  : " & rs("FullDesc") & vbCRLF & vbCRLF
	EmailBody = EmailBody & "To view the complete report click on :" & vbCRLF & GetQUESTServer() & "SR.asp?Q=" & TRIM(rs("QID")) & vbCRLF & vbCRLF
	EmailBody = EmailBody & "The URL to QUEST is " & GetQUESTServer() & vbCRLF
	
	If MsgID = 121 then rEventKey = "RIRClosed"
	If MsgID = 122 then rEventKey = "RIRAcknowledged"
	QueueMail EmailAddress,EmailSubj,PreMsg & EmailBody & PostMsg,"",rEventKey
	
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun NotifyUser QID="&SafeNum(QPID)
		End If
End Sub

Sub InitSeverityAndClass()
Dim SQL,RS1
	On Error Resume Next
	SQL=" Select 'S' as Type,SeverityID,SeverityDesc from tlkpRIRSeverity"
	SQL=SQL & " Union "
	SQL=SQL & " Select 'C' as Type,ClassID,ClassDesc from tlkpRIRClass"
	Set RS1=cn.execute(SQL)
	arrSevClass=RS1.getRows()
	rs1.close
	Set rs1=nothing
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun InitSeverityAndClass QID="&SafeNum(QPID)
		End If
End Sub

Function GetSevClass(Ty,Id)
Dim i
	On Error Resume Next
	GetSevClass="-"
    if Id="" then ID=0
    for i=0 to ubound(arrSevClass,2)
		If(arrSevClass(0,i)=Ty and cint(arrSevClass(1,i))=cint(Id)) then 
			GetSevClass=arrSevClass(2,i)
			Exit For
		End IF
	Next
	If trim(GetSevClass)="" then GetSevClass="-"
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun GetSevClass QID="&SafeNum(QPID)
		End If
End Function

Sub DeleteDGTreeInfo(QPID,cn)
Dim SQL
On Error Resume Next
	If CategoryList<>"''" Then
		SQL = " Delete  from tblririnvDGTree  Where QPID="&SafeNum(QPID) & " and invType='E' and InvPID not in "
		SQL = SQL & " (Select Distinct A.LossCatID from tlkpDG_LossSubCat A inner join tlkpLossSubCategories B on A.QuestLossID=B.ID Where B.ColumnName in ("&CategoryList&"))"
		cn.execute(SQL)	
		SQL = " Delete  Inv from tblririnvDGTree  Inv Where Inv.QPID="&SafeNum(QPID) & " and Inv.invType='AT' and Not Exists "
		SQL = SQL & " (Select InvID from tblririnvDGTree T Where T.QPID=Inv.QPID and T.InvType='E' and T.InvID=Inv.InvPID)"
		cn.execute(SQL)	
	Else
		SQL = " Delete  from tblririnvDGTree  Where QPID="&SafeNum(QPID) 
		cn.execute(SQL)	
	End IF
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun DeleteDGTreeInfo QID="&SafeNum(QPID)
		End If
End Sub

Function fncRemoveMultiSeg(mQPID)
	Dim mSQL

	On Error Resume Next
	mSQL="Delete from tblRIR_SQInvSegment where qpid =" & mQPID
	cn.execute(mSQL)

	mSQL="update tblRIRp1 set TCCInvolved =0 where qid=" & mQPID
	cn.execute(mSQL)
    ' Removed GRC Tab details
	cn.Execute "DELETE FROM tblRIR_SQTCCDetails WHERE QPID = " & mQPID
	cn.Execute "DELETE FROM tblRIR_SQTCCCatDetails WHERE QPID = " & mQPID
	cn.Execute "DELETE FROM tblRIR_SQTCCLocation WHERE QPID = " & mQPID
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSUbFun fncRemoveMultiSeg QID="&SafeNum(QPID)
		End If
End Function

Function fncRemoveIOGPUpdatedTabDetails(mQPID)
	Dim mSQL,sSQL
	On Error Resume Next
	sSQL="Select * from tblIPM_PSEdata where qpid =" & mQPID
	Set iogpRs = cn.execute(sSQL)
		if not iogpRs.EOF then
			iogpTier=iogpRs("Tier")
		End if
	Set iogpRs=Nothing
	if iogpTier < 3 then
	mSQL="Delete from tblIPM_PSEdata where qpid =" & mQPID
	cn.execute(mSQL)
	End if
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun fncRemoveIOGPUpdatedTabDetails QID="&SafeNum(QPID)
		End If
End Function

Function fncRemoveIOGPTabDetails(mQPID)
	Dim mSQL	
		On Error Resume Next
	mSQL="Delete from tblIPM_PSEdata where qpid =" & mQPID
	cn.execute(mSQL)
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun fncRemoveIOGPTabDetails QID="&SafeNum(QPID)
		End If
End Function

Function fncDeleteOPFDetails()
	Dim idelID
	Dim cmdDel
	On Error Resume Next
	Set cmdDel = Server.CreateObject("ADODB.Command")	
	With cmdDel
		.ActiveConnection = GetNewCn()
		.CommandType = adCmdStoredProc
		.CommandText = "spRIR_SQOPFDelete"
		.Parameters.Append .CreateParameter ("@QPId", adInteger, adParamInput, , iQPID)			
		.Execute()		
	End With
	Set cmdDel = Nothing
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun fncDeleteOPFDetails QID="&SafeNum(iQPID)
		End If
	LogAuditTrail  iQPID,lOrgNo,dtRptdate,"R","RIRdsp2.asp",1321,""	
	End Function
Set cn = Nothing

%>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:RIRdsp2.asp;66 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1255593] 17-AUG-2009 16:18:31 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02" %>
<% '       3*[1261359] 20-AUG-2009 15:24:29 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation" %>
<% '       4*[1271052] 25-SEP-2009 20:36:20 (GMT) VGrandhi %>
<% '         "SWIFT #2403986 - Well Services SQ Tabs" %>
<% '       5*[1275699] 29-SEP-2009 15:48:26 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '       6*[1277927] 02-OCT-2009 17:50:40 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '       7*[1287289] 28-OCT-2009 16:23:18 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab" %>
<% '       8*[1292460] 17-NOV-2009 17:42:23 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - Changes for Section 3 for Rig Related." %>
<% '       9*[1295941] 19-NOV-2009 16:46:01 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - 11-16-2009 Jon Changes" %>
<% '      10*[1303867] 23-DEC-2009 17:31:09 (GMT) VGrandhi %>
<% '         "SWIFT #2448303 - Develop EMS SQ Tab" %>
<% '      11*[1333698] 16-MAR-2010 05:59:52 (GMT) DMohanty %>
<% '         "Swift #2463864 - Data Gathering informations disappearing from RIR investigation tab" %>
<% '      12*[1348998] 29-APR-2010 12:44:21 (GMT) NNaik %>
<% '         "SWIFT # 2417393 - Lost configuration parameters" %>
<% '      13*[1345547] 25-MAY-2010 16:06:55 (GMT) DMohanty %>
<% '         "Swift # 2474157 - D&M SQ tab Development" %>
<% '      14*[1359004] 04-JUN-2010 21:43:45 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Develop interface to extract RIG NAME and associated info from CRM" %>
<% '      15*[1362560] 10-JUN-2010 06:00:52 (GMT) SKadam3 %>
<% '         "SWIFT #2474157 - D&M SQ tab (Web Services Call)" %>
<% '      16*[1371202] 22-JUN-2010 22:47:32 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Making IPM Project Location to Optional Rig Name Selection" %>
<% '      17*[1422556] 30-NOV-2010 10:45:00 (GMT) DMohanty %>
<% '         "Swift 2502891 - Error in RIR Delete / UnDelete" %>
<% '      18*[1458214] 16-MAR-2011 08:24:48 (GMT) PMakhija %>
<% '         "Swift#2541978-Create a new tab 'HOC' for M-I Swaco segment" %>
<% '      19*[1470763] 15-APR-2011 08:18:06 (GMT) MPatil2 %>
<% '         "SWIFT #2542386 - Add Q-STOP tab to SQ RIR" %>
<% '      20*[1485797] 23-MAY-2011 19:24:55 (GMT) APrakash6 %>
<% '         "SWIFT #2553287 - Incorrect Sub Sub Segment display in email instead of Sub Segment" %>
<% '      21*[1529537] 14-OCT-2011 12:05:24 (GMT) PMakhija %>
<% '         "SWIFT #2574362 - Path Finder TAB" %>
<% '      22*[1533099] 31-OCT-2011 14:49:39 (GMT) MPatil2 %>
<% '         "SWIFT #2588673 - Enable multisegment reporting via Involved Segments/Functions tab" %>
<% '      23*[1546388] 01-DEC-2011 14:15:44 (GMT) MPatil2 %>
<% '         "SWIFT #2594250 - GSS ML SQ Detail tab" %>
<% '      24*[1550566] 12-DEC-2011 09:37:47 (GMT) AGazi %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab" %>
<% '      25*[1557658] 12-DEC-2011 14:57:12 (GMT) KIrani %>
<% '         "SWIFT #2594250 - GSS ML SQ Detail tab" %>
<% '      26*[1565535] 29-DEC-2011 11:58:05 (GMT) MPatil2 %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab" %>
<% '      27*[1565888] 30-DEC-2011 10:42:08 (GMT) MPatil2 %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab" %>
<% '      28*[1570284] 12-JAN-2012 00:01:01 (GMT) APrakash6 %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab^HIDE MS FOR PROD" %>
<% '      29*[1588034] 27-FEB-2012 10:50:23 (GMT) MPatil2 %>
<% '         "SWIFT #2608547 - Multi Segment - Phase 2" %>
<% '      30*[1633354] 07-AUG-2012 16:27:13 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      31*[1693437] 16-NOV-2012 08:36:13 (GMT) MSaxena2 %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      32*[1690601] 12-DEC-2012 19:40:48 (GMT) APrakash6 %>
<% '         "SWIFT #2680271 - Feature: RIRs can record Well Barrier Events in SQ and HSE" %>
<% '      33*[1695958] 19-DEC-2012 14:51:00 (GMT) ATuscano %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      34*[1724644] 31-JAN-2013 12:18:03 (GMT) ATuscano %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      35*[1743176] 26-MAR-2013 10:16:18 (GMT) MSaxena2 %>
<% '         "SWIFT #2697446 - Feature: WIS SQ Detail Tab data flow into RITE.NET" %>
<% '      36*[1772304] 02-AUG-2013 12:02:35 (GMT) ATuscano %>
<% '         "SWIFT #2706924 - Well Barrier tab update - first and secondary envelope integrity" %>
<% '      37*[1783405] 04-OCT-2013 10:16:38 (GMT) ATuscano %>
<% '         "SWIFT #2713511 - Need to extract ASL Data (Contractors) from the WebService and UI Changes." %>
<% '      38*[1803272] 08-NOV-2013 06:58:27 (GMT) ATuscano %>
<% '         "ENH009582:Feature: Adopted ASL Master Data in QUEST" %>
<% '      39*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      40*[1835929] 14-MAY-2014 13:30:31 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      41*[1836849] 15-MAY-2014 10:19:46 (GMT) BGohil2 %>
<% '         "NFT014129 NPT/CMSL/TNCR data historical capture" %>
<% '      42*[1838252] 23-MAY-2014 11:14:16 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      43*[1838660] 27-MAY-2014 15:05:37 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '      44*[1843663] 23-JUN-2014 09:02:29 (GMT) Rbhalave %>
<% '         "ENH026013 - Defect# DEF033749" %>
<% '      45*[1853090] 14-AUG-2014 09:18:45 (GMT) VGrandhi %>
<% '         "Issue with RIR NPT History from the Main Page" %>
<% '      46*[1863565] 09-OCT-2014 10:45:28 (GMT) Rbhalave %>
<% '         "NFT039565 New SPS indicator on the use of SWI" %>
<% '      47*[1867842] 13-NOV-2014 11:40:26 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '      48*[1883751] 18-FEB-2015 15:10:07 (GMT) VSharma16 %>
<% '         "ENH053415 - Addition of PTEC project acknowledgement on HSE and SQ RIR and O/I reports" %>
<% '      49*[1890939] 19-MAR-2015 14:28:46 (GMT) Rbhalave %>
<% '         "NFT056368  FEATURE QUEST Upgrades for Facilities" %>
<% '      50*[1896385] 17-APR-2015 11:42:05 (GMT) Rbhalave %>
<% '         "ENH059068 P&AM Process Safety (Updates in tab and the report)" %>
<% '      51*[1903829] 08-JUN-2015 10:50:32 (GMT) Rbhalave %>
<% '         "ENH055492  SPS request - addition of SPS Process categories to SQ RIR report - all Segments" %>
<% '      52*[1911595] 17-JUL-2015 10:27:05 (GMT) SChaudhari %>
<% '         "ENH077170 - SAXON - Operations at Time of Event (Category and Sub-Category)" %>
<% '      53*[1915630] 19-AUG-2015 06:50:17 (GMT) SChaudhari %>
<% '         "ENH077312 - SAXON - Subscription notification change" %>
<% '      54*[1927564] 20-NOV-2015 12:43:22 (GMT) SChaudhari %>
<% '         "ENH086389 - Saxon - SQ Categories" %>
<% '      55*[1932385] 24-NOV-2015 10:37:52 (GMT) Rbhalave %>
<% '         "NFT087279 - TS segment becoming a forced segment - Rig related flag at sub sub segment level" %>
<% '      56*[1937795] 31-DEC-2015 07:55:55 (GMT) MPatel13 %>
<% '         "ENH092498-TLM Sub-Segments to invoke Tool Parent SQ Tabs" %>
<% '      57*[1939904] 12-JAN-2016 10:09:21 (GMT) MPatel13 %>
<% '         "REMOVING-TLM-ENH092498-TLM Sub-Segments to invoke Tool Parent SQ Tabs" %>
<% '      58*[1940182] 14-JAN-2016 07:42:22 (GMT) MPatel13 %>
<% '         "TLM-Redo-ENH092498-TLM Sub-Segments" %>
<% '      59*[1941773] 16-FEB-2016 13:03:45 (GMT) VSharma16 %>
<% '         "ENH100497: SPS addition of Categories (SUPPORT ITT PROJECT)" %>
<% '      60*[1948724] 12-APR-2016 07:05:05 (GMT) MPatel13 %>
<% '         "ENH095322-SLIM - SQ RIR report changes to form fields" %>
<% '      61*[1944390] 12-APR-2016 11:19:28 (GMT) VSharma16 %>
<% '         "ENH096140 - Integrated Projects" %>
<% '      62*[1951386] 12-APR-2016 14:13:20 (GMT) Rbhalave %>
<% '         "ENH101167 - SLIM ROOT CAUSE CLASSIFICATION" %>
<% '      63*[1953251] 13-MAY-2016 09:25:28 (GMT) MPatel13 %>
<% '         "Updated Files" %>
<% '      64*[1953356] 13-MAY-2016 10:07:23 (GMT) Rbhalave %>
<% '         "ENH101167 SLIM Root cause classification" %>
<% '      65*[1878140] 24-JUN-2016 06:56:22 (GMT) VSharma16 %>
<% '         "NFT101068- RIR locking to prevent data integrity issues" %>
<% '      66*[1897136] 24-JUN-2016 07:39:46 (GMT) SChaudhari %>
<% '         "ENH115371 - <<MM>> Change TLM SQ Tab Selection to Sub Sub-Segment" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:RIRdsp2.asp;66 %>
