<%@  language="VBScript" %>
<%option explicit
	'****************************************************************************************
	'1. File Name		              :  RIRdsp.asp
	'2. Description           	      :  Main RIR Page
	'3. Calling Forms   	          : 
	'4. Stored Procedures Used        : 
	'5. Views Used	   	              : 
	'6. Module	   	                  : RIR (HSE/SQ)				
	'7. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'    5-Aug-2009			     Nilesh Naik        	 	Modified - changed for NPT <<2401608>>
	'    2-Sep-2009              Nilesh Naik                Modified - changed for SWIFT # 2411197 
	'    15-Sep-2009            Visali Grandhi              Modified for well services - #2403986 
	'    29-Sep-2009            Nilesh Naik                 Modified - changed for NPT # 2401608 
	'    15-Oct-2009			Micheal Anthony				Modified - SWIFT # 2409503 
	'    27-Oct-2009			Visali Grandhi				Modified - SWIFT #2430251 - IPM SQ Event Tab
	'    30-Oct-2009			Shailesh					Modified - To Exclude some segment for severity check Swift# 2438856
	'    16-Nov-2009            Visali Grandhi              Added IPM General Condition for RIR Close.
	'    14-Dec-2009            Visali Grandhi              SWIFT #2448303 - Develop EMS SQ Tab
	'    18-Dec-2009            Visali Grandhi              IPM Tab Conditions and Close Condition.
	'     6-Apr-2010            Nilesh Naik                 Modified - Changed for 2468497 - to Add Pop-up tp Quality Loss selections 
	'     7-May-2014            Varun Sharma                Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
	'    19-Aug-2014            Rohan Bhalave               Modified for Time loss tab validation while closing the report.
	'   28-OCT-2014             Varun Sharma                ENH043340  Changes in PF same as D & M Phase 1 & 2
	'   31-OCT-2014             Varun Sharma                ENH044752  HSE locking of lagging indicators - safety net - key to unlock
	'17-Feb-2015				sagar Chaudhari			    Modified - ENH053415-Addition of PTEC project	 acknowledgement on HSE and SQ RIR and OI reports
	'13-March-2015				sagar Chaudhari				ENH059674: 	Javascript forcing of PTEC selection - RIR and Obs/Int
	'   17-March-2015           Varun Sharma                ENH054251 D&M Tab updates: Phase11
	' 29-June-2015              Rohan Bhalave               ENH060157 Removal of M-I Swaco HSE RIR feature page HOC
	' 30-June-2015              Rohan Bhalave 				ENH075535  - SQ SPS Process Phase 2
	' 10-July-2015				Sagar Chaudhari				ENH077170 - SAXON - Operations at Time of Event (Category and Sub-Category)	
	' 17-July-2015				Sagar Chaudhari				NFT075965 -D&M Phase 12 - Part 1
	' 	31-Aug-2015			    Varun Sharma				ENH081653:  D&M Phase 12 - Part 2
	' 	20-OCT-2015			    Sagar Chaudhari 	
	' 	18-NOV-2015			    Varun Sharma				ANO091859:D&M - Post SCAT - Root Cause logic changes
	' 	19-NOV-2015			    Rohan Bhalave				NFT087279 - TS segment becoming a forced segment - Rig related flag at sub sub segment level
	' 	25-FEB-2016			Varun Sharma				  ENH104074:SPWL move to WL - requires SQ Wireline SQ tab to be implemente
	' 	28-MAR-2016			   Varun Sharma				ENH095327-SLIM - Job ID changes
	' 	17-FEB-2016			Varun Sharma				 ENH096140 - Integrated Projects
	' 	12-APR-2016			Varun Sharma				 ANO103787:  SQ RIR SPS Categories userability
	'   21-Jun-2016              Varun Sharma                 NFT101068- RIR locking to prevent data integrity issues
	' 10-Apr-2017                                           cameron merge
	'11-July-2017
	'******************************************************************************************************************************
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
		'Response.Expires = 0
		'Response.ExpiresAbsolute = Now() - 1
		'Response.AddHeader "pragma","no-cache"
		'Response.AddHeader "cache-control","private" 
		'Response.CacheControl = "no-cache"
		'debugPrint("Process Started" & Now())
%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<!-- #INCLUDE FILE="RIRHelpText.asp"-->
<!-- #INCLUDE FILE="../Inc/Inc_CRMClient.asp"-->
<!-- #INCLUDE FILE="../Inc/Inc_PLBSFunctions.asp"-->
<!-- #INCLUDE FILE="../Inc/Inc_Security.asp"-->
<!-- #INCLUDE FILE="../Inc_NavFrame.asp"-->
<%


	'print IsIE
	If IsIE then Response.Expires = -1
	checktimeout()

	Dim dtRptDate, RSEmployees, RSBSegments, strHeader, iQPID,SPS_L2,SPS_L3,SPS_L4,spsl45,SPS_L4_VALUE,rs1,VarHideLegacyInvestigation,LegacyChkFlg
	Dim iTemp, iTemp2, sTemp, sTemp1, sTemp2, sTemp3, bTemp
	Dim lBSegmentID, lPL, iBSID, lDepartmentID,TCCInv,tNPT,rNPT,SegInv,Grctcc '***** (MS HIDDEN) - Commented "SegInv"  ***** 
	Dim tlPL ,RSHazard,opfepcc,opfonm,valopf
	Dim iCntr, lOrgno, bNR, dtRptDateMast, cn, sHref, RSOrg, sArea, bReturnWorkDateClosureExclusion
	Dim sGeo, RSCust, sSel, lCustID, dtTemp, sSQL, RS, RS2, iOpenActions, iClosedActions
	Dim sKey , RSSQ,SQL, Source, isTimeLossTab, isTimeLossEntered
	Dim PostVars, HSESQMEssage,DMRecs,DMRS,DMRS1,FireMode
	DIM ACLDefined,sUID,sUName
	Dim PFRecs,PFRS,PFRS1
	dim chkopfhazornear,delopfdetail,ActLegacyChkFlg
	Dim WlRo,SQRo

	'Post variables
	Dim RptDate, ReportNumber, RevDate, bReviewed, bClosed, UpdatedBy, CreatedBy, CreateUID, UpdateUID
	Dim BusinessSegment, Client,Contractor,CRMClient,CRMRigID
	Dim Reporter, ReporterUID, EventDate, EventTime
	Dim Loctn
	Dim SiteName, PRiskClass, RRiskClass, ShortDescription, LongDescription
	Dim iClass ,iSev ,bSQ,bHSE, HazardCat,iClassn,LOCCountry
	Dim SQFCatID , SQDCatID ,SQCCatID, SQDelayHrs , SQNRedone,SQStandard
	Dim SQ_Process, SQ_ProcessOwn,SQ_MetroStop, SQ_Activity
	Dim SQSPCatID, SQSPSubCatID, IsInvSQ, SQPFailure 
	Dim SQFSubcatID, SQDSubcatID, SQCSubcatID, JobID,SQInvment,External,ClientAffect
	Dim SLBInv, IndRec, RegRec, SLBRel, bPostVars, sServerVariables, SLBCon,SLBInvment
	Dim iHSESev, iSQSev ,IPMInv,PTECInv,ContractorInv, IsSPRequired, IsCauseRequired,isDamageRequired,HideSQ,ShowIPMSQ,FinanceInv, ROPInv '***** (MS HIDDEN) - Commented "FinanceInv"  ***** 
	Dim sSQSeverityMatrix, sHSESeverityMatrix, WidthFactor,AccUnit,WellSite,HideWSSQ,strNPT_LossCat_G1,strNPT_LossCat_G2,strTLMatrix
	Dim bWIBEvent,bAccDischarge,bFireExpl, RsHaz, HazID,lossSafetynetVal
	Dim SwiQn1, SwiQn2A, SwiQn2B, SwiQn3,swistyle,RSswi,SwiSQL,isSWI 
	dim  RSPTEC, PTECSQL , isPTEC ,EnableOperation, isROP,RSPI,SPISQL
	Dim SQProcCats,SQCats,spsl2,spsl3,spsl4,isSPS,descp, IniStyle,OperationCat,OperationSubCat,spsb2
    Dim sSQLrole,rsDatarole,ToCheck,sSQLroleAttribID,rsDataroleAttribID,BCID
	Dim sSQLroleMaint,rsDataroleMaint,ToCheckMaint,sSQLroleAttribIDMaint,rsDataroleAttribIDMaint,BCIDMaint
	Dim sSQLFailureCatID,rsFailureCatID,FailureCatID,sSQLroleCatName,rsDataroleCatName,DMCatNameval, Qlocation, LocPLID,LocPID
	Dim countvaluetotal,rsDatasqlcountvalue,sqlcountvalue,countvalue, EnforceFlag
	dim Saxoncats,SAXCats,RSpl, SQMappingID,SPS_ID1,SQCategoryMappingID
	dim projectNOVal,projectIDSVal,projectIPSVal,projectIFSVal,projectISMVal,projectSPMVal,IPMNo,IPMIDS,IPMIPS,IPMIFS,IPMISM,IPMSPM,optInvnoVal 
	Dim JobLbl,projectunknown,projectvisible
	DIM iSubBSID 'SlimCntrl, SlimTech, SlimProc, SlimCompet, SlimBehav, RSslim
    Dim EvtClassSaf,EvtClassChoice,EvtSubClassSaf,bPLSSInv
	Dim DtEvtClass,DiffEvtClass, BusWrkFLowHubLink
	Dim VarJquery,VarSLBHub,VarSharePointURL,VarHideFailureListing,VarHideLegacyCSUR,VarHideWellBarrierInv,VarReplaceNAMwithGeomarket
	Dim VarActionItemValidation,ComparedateSSONew,HideFAILEvents,varactlegacy
	On Error Resume Next			
	VarHideFailureListing = fncD_Configuration("HideFailureListing")
	comparedateOST = fncD_Configuration("OSTMove")
	VarJquery = fncD_CommonURL("Jquery")
	VarSLBHub = fncD_CommonURL("SLBHub")
	VarSharePointURL = fncD_CommonURL("SharePointURL")
	VarHideLegacyInvestigation = fncD_Configuration("HideLegacyInvestigation")
	VarHideLegacyCSUR  = fncD_Configuration("CSURNewUI")
	VarHideWellBarrierInv  = fncD_Configuration("HideWellBarrierElementInvolved")
	VarReplaceNAMwithGeomarket = fncD_Configuration("RITE_AU_Geomarket")
	VarActionItemValidation = fncD_Configuration("ActionItemValidation")
	comparedateOPF=fncD_Configuration("OPFMove")
	ComparedateSSONew = fncD_Configuration("SSONewUI")
	HideFAILEvents = fncD_Configuration("HideFAILEvent")
	varactlegacy = fncD_Configuration("LegacyAccountingUnit")
	

	
	JobLbl="Job / Service Order Id"
	OperationSubCat= 0	
	WidthFactor = 1
	projectNOVal=""
	projectIDSVal=""
	projectIPSVal=""
	projectISMVal=""
	projectSPMVal=""
	optInvnoVal=""
	projectunknown=""
	projectvisible=""


	if GetBrowserType() = "MSIE" Then
		widthfactor = 1.8
	End iF
	sUID=lcase(Session("UID"))
	sUName=Session("UserName")
	FireMode = Request.Form("FireMode")
    BusWrkFLowHubLink = fncD_Configuration("BusWorkFlow_HUBLink")
	'SetLocation QTID
	'TS changes
	If Request.QueryString("postvars")=1 then 
		Set PostVars = GetVariables()
		If PostVars.Exists("QuestLoc") Then 
			SetLocation PostVars("QuestLoc")

			Qlocation = request.form("Qloc")
		End if
	End if
	
	
	'NBL Aug 1 2002 Added Stored Server Variables into a string
	sServerVariables = StoreSessionInfo

	Set cn = getNewCN()

	Set RsHaz = Server.CreateObject("ADODB.Recordset")	
	RsHaz.open "SELECT  HazardID FROM tlkpRIRHazardCategory WHERE HazardDesc = 'Fire/Flammables'", cn
	HazID = RsHaz("HazardID")
	RsHaz.Close

	If Request.QueryString("postvars")=1 then 
		bPostVars = True
		'Set PostVars = GetVariables()
		If PostVars.Exists("cat_sq") Then OperationCat = PostVars("cat_sq")
		If PostVars.Exists("subcat_sq") Then OperationSubCat = PostVars("subcat_sq")	
	End If
		
	If Request.QueryString("NR")= "1" then
		bNR=True 
	Else 
		bNR=False
	End if
	 If Err.Number <> 0 Then
		' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description & "-At Lineno181 QID= " &iQPID
	End If	

	iOpenActions = 0
	iClosedActions = 0
	lDepartmentID = 0
	isTimeLossEntered = 0
	LocPLID = 0
   On Error Resume Next
	If bNR Then
	On Error Resume Next
		if not IsActiveNode(Session("HomeNo")) then Response.Redirect "../Disabled.htm"
		lOrgNo = Trim(Session("HomeNo"))
		dtRptDate = now()	
		BusinessSegment=GetSubBusinessSegID(lOrgNo)
		IniStyle = "display:none;"
		LocPLID = GetProductLine(lOrgNo)    'TS Changes

	else
		lOrgNo = Request.QueryString("OrgNo")
		'LocPLID = GetProductLine(lOrgNo)	'TS Changes
		dtRptDate = Request.QueryString("rptDate")
		iQPID = Request.QueryString("QPID")
		
		sSQL = "SELECT QT.BSID,tblRIRp1.*, CASE WHEN EXISTS (SELECT * FROM tblRIRPers WHERE QPID = tblRIRp1.QID AND ReturnWorkDate IS NULL AND (DaysLost > 0)) THEN 1 ELSE 0 END AS ReturnWorkDateClosureExclusion FROM tblRIRp1 "
		sSQL=sSQL & " INNER JOIN tblQT_QuestTree QT ON tblRIRp1.OrgNo = QT.ID "
				
		If iQPID<>"" Then 
			sSQL=sSQL & " WHERE tblRIRP1.QID=" & SafeNum(iQPID)		
		Else
			sSQL=sSQL & " WHERE tblRIRP1.OrgNo='" & SafeNum(lOrgNo) & "' AND tblRIRP1.RptDate='" & dtRptDate & "'"		
		End IF
		
		Set RS = Server.CreateObject("ADODB.Recordset")		
        RS.Open sSQL, cn
		
		If rs.EOF or rs.BOF then 
			Response.Redirect("../Utils/RecNotFound.htm")
		Else
			iQPID=RS("QID")
			lOrgNo = RS("OrgNo")
			dtRptDate = RS("rptDate")
			SPS_L2=RS("SPS_L2")
			SPS_L3=RS("SPS_L3")
			SPS_L4=RS("SPS_L4")
			iClassn = RS("Class")
			'Saxon Changes		
			OperationCat  = iif (RS("OperationCat")="Null", 0 , RS("OperationCat"))
			OperationSubCat = iif(RS("OperationSubCat")="Null",0 , RS("OperationSubCat"))
			
				'swi changes
		SwiQn1  =  iif (RS("Qn1")="Null",0,RS("Qn1"))
		SwiQn2A =  iif (RS("Qn2A")="Null",0,RS("Qn2A"))
		SwiQn2B =  iif (RS("Qn2B")="Null",0,RS("Qn2B"))
		SwiQn3  =  iif (RS("Qn3")="Null",0,RS("Qn3"))
		
		
		spsl2 = iif (RS("SPS_L2")="Null",0,RS("SPS_L2"))
		spsl3 = iif (RS("SPS_L3")="Null",0,RS("SPS_L3"))
		spsl4 = iif (RS("SPS_L4")="Null",0,RS("SPS_L4"))
		spsb2 = iif (RS("SPS_B2")="Null",0,RS("SPS_B2"))
		
		Set RsPI = Server.CreateObject("ADODB.Recordset")
		SPISQL = "sp_GetSeverityValidation " & SafeNum(iQPID) & ""
		RsPI.open SPISQL, cn
		If not RsPI.EOF then
		lossSafetynetVal = RsPI("Val")
		end if 
		RsPI.Close
		
		if  spsl2 <>"" then 
				dim spsl4_new
				spsl4_new =GETVALSUFFIX(spsl4)
			END IF
		
		IniStyle = " "
		if spsl4 = 0 then IniStyle = "display:none;"
		
			GetRACount iQPID,"R",iOpenActions,iClosedActions	
			isTimeLossTab = isTime(RS)
			bReturnWorkDateClosureExclusion = RS("ReturnWorkDateClosureExclusion")		
			BusinessSegment = RS("BSID")
			Set RS2 = Server.CreateObject("ADODB.Recordset")	
			RS2.open "select Cost,CostCatID from tblRIRCosts where QPID = " & iQPID & " and CostCatID in (30,33)", cn
			if not rs2.eof then
				while not rs2.eof			
					if rs2("CostCatID") = 30 then
						strNPT_LossCat_G1 = rs2("Cost")
					elseif rs2("CostCatID") = 33 then
						strNPT_LossCat_G2 = rs2("Cost")
					end if
					rs2.movenext
				Wend
			end if
			rs2.close
			RS2.open "select count(*) as TimeLossRec from tblRIRTime where QPID = " & iQPID, cn
			if not rs2.eof then
				if cint(rs2("TimeLossRec")) > 0 then isTimeLossEntered = 1
			end if
			rs2.close
			set rs2 = nothing
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description & "-At Lineno290 Qid=" & iQPID 
		End If	
		End if
		
		' SLIM Changes
		'Commented SLIM Changes
		'sSQL = "SELECT [ControlDomain],[TechDomain] , [ProcDomain], [CompetDomain], [BehaviorDomain] "
		'sSQL = sSQL & " FROM tblRIR_SLIMdata WHERE QPID = " & iQPID & " AND SlimType = 'SLIM-INIT' "
		
		'set RSslim = Server.CreateObject("ADODB.Recordset")
		'RSslim.Open sSQL, cn
		
		'If NOT RSslim.EOF then 
			
			'SlimCntrl = RSslim("ControlDomain")
			'SlimTech = RSslim("TechDomain") 
			'SlimProc = RSslim("ProcDomain")
			'SlimCompet = RSslim("CompetDomain") 
			'SlimBehav = RSslim("BehaviorDomain")
		'END IF
		'RSslim.close
		'set RSslim = nothing
		
	end If
	
	If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(varactlegacy))  then
	ActLegacyChkFlg="True"
	End If
	
	If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideLegacyInvestigation)) and bHSE<>true  then
	LegacyChkFlg="True"
	End If

	If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideLegacyInvestigation)) and bHSE=true and bSQ=true then
	LegacyChkFlg="True"
	End If
	On Error Resume Next


		 SPS_L4_VALUE=""
	 'response.write SQL
	 Set rs1 = Server.CreateObject("ADODB.RecordSet")
		    SQL ="select Description from tblRIR_SPSData where  id in (select SPS_L4 from tblRIRp1 where QID="&SafeNum(iQPID)&")"
		    'SQL ="select SPS_L4 from tblRIRInvBusWorFlows where QPID="&SafeNum(QPID)&" and bseqid="&SelID
		rs1.open SQL, cn
		'response.write SQL
		if not rs1.eof then
			SPS_L4_VALUE = RS1("Description")
		end if
		if SPS_L4_VALUE<>"" then 
		SPS_L4_VALUE="yes"
		end if 
			rs1.close
	set rs1=nothing
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description & "-At Lineno341  Qid=" & iQPID 
		End If
	
LocPLID = GetProductLine(lOrgNo)	'TS Changes
		DtEvtClass = fncD_Configuration("ShowEventClassification")
		DiffEvtClass = DateDiff("n",DtEvtClass,dtRptDate)
	' TS Changes
	On Error Resume Next
	sSQL = "SELECT EnforceSelection FROM tblProductLines WHERE PLID = " & LocPLID
	set RSpl = Server.CreateObject("ADODB.Recordset")
	RSpl.Open sSQL, cn
	
	If NOT RSpl.EOF then 
		EnforceFlag=RSpl("EnforceSelection")
	END IF
	RSpl.close
	set RSpl = nothing


		On Error Resume Next
		' SWI change
		Set RSswi = Server.CreateObject("ADODB.Recordset")
		SwiSQL = "SELECT isNULL(PL.EnableSwiQn,0) as EnableSwiQn ,isNULL(PL.EnablePTEC,0) as EnablePTEC, isNULL(PL.EnableSPS,0) as EnableSPS , isNULL(PL.EnableOperation,0) as EnableOperation, isNULL(PL.EnableROP,0) as EnableROP  FROM tblProductLines  PL "
		SwiSQL = SwiSQL & " LEFT JOIN tlkpBusinessSegments BS on BS.PPLID = PL.PLID "
		SwiSQL = SwiSQL & " WHERE BS.BusinessSegmentID =" & BusinessSegment 
	
		RSswi.Open SwiSQL, cn

		isSWI = 0
		If not RSswi.EOF then 
			isSWI = RSswi("EnableSwiQn")
		End If

		If not RSswi.EOF then 
			isPTEC = RSswi("EnablePTEC")
		End If	
		
		If not RSswi.EOF then 
			isROP = RSswi("EnableROP")
		End If

		isSPS = 0
		If not RSswi.EOF then 
			isSPS = RSswi("EnableSPS")
		End If
		
		EnableOperation=0
		If not RSswi.EOF then 
			EnableOperation = RSswi("EnableOperation")
		End If	
		RSswi.close
		Set RSswi=nothing

		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  & "- At Lineno393  Qid=" & iQPID  
		End If
		'PTEC Changes 
		' Set RSPTEC = Server.CreateObject("ADODB.Recordset")
		' PTECSQL = "SELECT isNULL(PL.EnablePTEC,0) as EnablePTEC FROM tblProductLines PL "
		' PTECSQL = PTECSQL & " LEFT JOIN tlkpBusinessSegments BS on BS.PPLID = PL.PLID "
		' PTECSQL = PTECSQL & " WHERE BS.BusinessSegmentID =" & BusinessSegment 

		' RSPTEC.Open PTECSQL, cn

		' isPTEC = 0
		' If not RSPTEC.EOF then 
			' isPTEC = RSPTEC("EnablePTEC")
		' End If
		' RSPTEC.close
		' Set RSPTEC=nothing

		
		
	' change for comp loss
		On Error Resume Next
		dim isExist, RSComp, CompStatus, CompOpen, RSQL
		Set RSComp = Server.CreateObject("ADODB.Recordset")
		
		RSQL = "SELECT isNULL(CL.SeqID,0) AS isExist, AT.Seq AS ItemNo, AT.Description "
		RSQL = RSQL & " FROM tblRIRAssets AT "
		RSQL = RSQL & " Left join tblRIRCompLoss CL on AT.QPID = CL.QPID AND AT.Seq = CL.SeqID "
		RSQL = RSQL & " WHERE AT.QPID=" & SafeNum(iQPID) &" AND AT.Computer =1 ORDER BY AT.seq"
		
		RSComp.Open RSQL, cn
		If not RSComp.EOF then 
		While not RSComp.EOF
			isExist = RSComp("isExist")
			CompStatus = iif(isExist=0,"Add","Edit")
			if CompStatus = "Add" Then
			CompOpen = 1
			'Msg=Msg & "\nComplete the computer loss details."	

			End if 

		RSComp.MoveNext
		WEND
		End If
		RSComp.close
		Set RSComp=nothing
		
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  & "- At Lineno441  Qid=" & iQPID  
		End If
        On Error Resume Next
	Set RS2 = Server.CreateObject("ADODB.Recordset")	
	RS2.open "select MatrixID,CustSLBLoss1,NPTHours1 from tblRIRTimeLossMatrix order by MatrixID", cn
	if not rs2.eof then
		strTLMatrix = "##"
		while not rs2.eof
			strTLMatrix = strTLMatrix & rs2("MatrixID") & "#" & rs2("CustSLBLoss1") & "#" & rs2("NPTHours1") & "##"
			rs2.movenext
		wend
	end if
	rs2.close
	set rs2 = nothing


	'Deepak Added for D&M Changes
	If ShowDM() Then
		Set DMRS = Server.CreateObject("ADODB.Recordset")
		DMRS.Open "Select Count(*) AS RecCount FROM tblRIRP1 With (NOLOCK) WHERE QID='" & SafeNum(iQPID) & "' AND (SQSPCatID <> '' OR SQSPSubCatID <> '' OR SQFCatID <> '' OR SQFSubCatID <> '' OR SQDCatID <> '' OR SQDSubCatID <> '')",cn
		DMRecs = DMRS("RecCount")
		DMRS.Close
	End If
	'Deepak- End

	If ShowPF() Then
		Set PFRS = Server.CreateObject("ADODB.Recordset")
		PFRS.Open "Select Count(*) AS RecCount FROM tblRIRP1 With (NOLOCK) WHERE QID='" & SafeNum(iQPID) & "' AND (SQSPCatID <> '' OR SQSPSubCatID <> '' OR SQFCatID <> '' OR SQFSubCatID <> '' OR SQDCatID <> '' OR SQDSubCatID <> '')",cn
		PFRecs = PFRS("RecCount")
		PFRS.Close
	End If

	'Shailesh 30-Oct-2009 Swift# 2438856
	Dim blnNPT_Exempt
	blnNPT_Exempt = 0
	blnNPT_Exempt = ChkNPT_Exemption()
	'Shailesh 30-Oct-2009

	sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
	'Check for user access.
	ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	lPL=GetProductLine(lOrgNo)
	iBSID=GetSubBusinessSegID(lOrgNo)
	iSubBSID = GetSubSubBusinessSegID(lOrgNo)
	SQMappingID = GetSQMappingID(iSubBSID)
	SQCategoryMappingID = GetSQCategoryMappingID(iSubBSID)
	
	If (SQCategoryMappingID > 0)  then 
		tlPL = SQCategoryMappingID
	Else 
		tlPL = lPL	
	End IF
	
	InitVars()
	
	IF dtRptDatetmp<>"" then
		If (SQMappingID = 7 and (CDate(dtRptDatetmp) < CDate(comparedateOST)) and iBSID=9263) Then SQMappingID=3             'To handle legacy data of One Stim this conversion is done forcefully
	End IF 
	'If (SQMappingID = 7 and (CDate(dtRptDatetmp) < CDate(comparedateOST)) and (iBSID=9262 Or iBSID=9265 or iBSID=9266)) Then SQMappingID=5 'To handle legacy data of One Stim this conversion is done forcefully
	SetSeverityMatrix(lPL)	
	 'if isMNSIT(iBSID) then sSQSeverityMatrix = VarSLBHub &"Docs/qhse/IM/OFS_SQ_Severity_Matrix.htm"	
	  if isMNSIT(iBSID) then sSQSeverityMatrix = VarSharePointURL &"/sites/OSPerformance/Shared%20Documents/Severity%20Matrix/OFS_SQ_Severity_Matrix.aspx"

	IsInvSQ = IsCompleteInvSQ()
	SetCookieList "R",iQPID
	HideWSSQ = 0
	if not (isREWSQMapping(SQMappingID) or isSPWL(iBSID) or (isOFS(iBSID) and not (isMNSIT(iBSID))) or isWTSSQMapping(SQMappingID) or isEMS(SQMappingID) or isWSSQMapping(SQMappingID) or isIPMSeg(SQMappingID) or isSWACO(SQMappingID) or (isOne(SQMappingID) and not (isOneCPL(iBSID)))) then HideSQ = 1 else HideSQ=0    'isWTS(lPL)  --isRew(lPL)  isWS(lPL)
	if isWSSQMapping(SQMappingID) then HideWSSQ = 1   'isWS(lPL)
	if isIPMSeg(SQMappingID) and SQSPCatID>0 Then ShowIPMSQ=1 else ShowIPMSQ=0
	GetMatrixData()

	rNPT = getRigNPT()

	' SPS Changes 
	SQProcCats=ExtractSQProcCats()

	Saxoncats = ExtractSAXONCats()
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  & "- At Lineno512 QID="&iQPID
	End If
	
	Dim rsIPMInv,strSQLIPMInv
	SET cn = GetNewCN()
	On Error Resume Next
	strSQLIPMInv = "select IPMInv from tblrirp1  where QID=" & SafeNum(iQPID)
	SET rsIPMInv = cn.execute (strSQLIPMInv)
	If NOT rsIPMInv.eof Then IPMInv = rsIPMInv("IPMInv")
	rsIPMInv.close
	Set rsIPMInv = Nothing
	If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  & "- At Lineno529"
		End If


	'WLRo involved
	Dim strWL,rsWL,strWLSQ,rsWLSQ
	SET cn = GetNewCN()
	strWL = "select distinct(isroinvolved) from tblRIR_SQWLIncidents where qpid="& SafeNum(iQPID)&" and isroinvolved=1"	
	SET rsWL = cn.execute (strWL)
	If NOT rsWL.eof Then			
	WlRo = 1
	Else
	WlRo = 0
	End If	
	rsWL.close
	Set rsWL = Nothing

	strWLSQ="Select ROP from tblrirp1 where qid="& SafeNum(iQPID)	
	SET rsWLSQ = cn.execute (strWLSQ)
	If NOT rsWLSQ.eof Then
	if rsWLSQ("ROP") ="False" Then
	SQRo=0
	Else
	SQRo=1
	End If
	End if
	rsWLSQ.close
	Set rsWLSQ= Nothing
	
	'if IPMInv="False" and IPMNo=0 and IPMIDS=0 and IPMIFS=0 and IPMISM=0 and IPMSPM=0 and iQPID>0 then
	if IPMInv="False" and IPMNo=0 and IPMIDS=0 and IPMIFS=0 and IPMSPM=0 and iQPID>0 then
		dim strSQL
		SET cn = GetNewCN()
		If bnr then 
			strSQL = "update tblrirp1 set ProjectNO=1 where QID=" & SafeNum(iQPID)
			IPMNo=1
			cn.Execute (strSQL)
		End if 	
		
	end if	

	'if  len(request.form("projectNO")) =0 and  not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMPRSS(iBSID) and not isIPMIFS(iBSID) and IPMIDS=0 and IPMIFS=0 and IPMISM=0 and IPMSPM=0 and  (iQPID="" or iQPID=0) then 
	if  len(request.form("projectNO")) =0 and  not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMIFS(iBSID) and IPMIDS=0 and IPMIFS=0 and IPMSPM=0 and  (iQPID="" or iQPID=0) then
	projectNOVal=" checked=""checked"""
	else
	projectNOVal=""
	end if
	
	
	if len(request.form("projectIDS")) > 0 then
	projectIDSVal=" checked=""checked"""
	IPMIDS=2
	end if
	
	'if len(request.form("projectIPS")) > 0 then
	'projectIPSVal=" checked=""checked"""
	'IPMIPS=3
	'end if
	
	if len(request.form("projectIFS")) > 0 then
	projectIFSVal=" checked=""checked"""
	IPMIFS=3
	end if
	
	'if len(request.form("projectISM")) > 0 then
	'projectISMVal=" checked=""checked"""
	'IPMISM=4
	'end if
	
	if len(request.form("projectSPM")) > 0 then
	projectSPMVal=" checked=""checked"""
	IPMSPM=5
	end if
	
	if len(request.form("projectNO")) > 0 then
	projectNOVal=" checked=""checked"""
	IPMNo=1
	end if
	

	if  ((isIPMWCSS(iBSID) and IPMNo=0) or (isIPMWCSS(iBSID) and IPMIDS =2)) then
	projectIDSVal=" checked=""checked"" disabled=""disabled"""
		'If IPMIPS =3 then projectIPSVal=" checked=""checked"""
		If IPMIFS =3 then projectIFSVal=" checked=""checked"""
		'If IPMISM =4 then projectISMVal=" checked=""checked"""
		If IPMSPM =5 then projectSPMVal=" checked=""checked"""
	end if
	
	if   ((isIPMIFS(iBSID) and IPMNo=0) or (isIPMIFS(iBSID) and IPMIFS =3)) then 
	'projectIPSVal=" checked=""checked"" disabled=""disabled"""
	projectIFSVal=" checked=""checked"" disabled=""disabled"""
	
		If IPMIDS =2 then projectIDSVal=" checked=""checked"""
		'If IPMISM =4 then projectISMVal=" checked=""checked"""
		If IPMSPM =5 then projectSPMVal=" checked=""checked"""
		
	end if
	
	'if ((isIPMPRSS(iBSID) and IPMNo=0) or (isIPMPRSS(iBSID) and IPMISM =4)) then
	'projectISMVal=" checked=""checked"" disabled=""disabled"""
	
		'If IPMIDS =2 then projectIDSVal=" checked=""checked"""
		''If IPMIPS =3 then projectIPSVal=" checked=""checked"""
		'If IPMIFS =3 then projectIFSVal=" checked=""checked"""
		'If IPMSPM =5 then projectSPMVal=" checked=""checked"""
		
	'end if
	
	if  ((isIPMAPS(iBSID) and IPMNo=0) or (isIPMAPS(iBSID) and IPMSPM =5))  then
	projectSPMVal=" checked=""checked"" disabled=""disabled"""
	
		If IPMIDS =2 then projectIDSVal=" checked=""checked"""
		'If IPMIPS =3 then projectIPSVal=" checked=""checked"""
		If IPMIFS =3 then projectIFSVal=" checked=""checked"""
		'If IPMISM =4 then projectISMVal=" checked=""checked"""
		
	end if
	
	'if not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMPRSS(iBSID) and not isIPMIFS(iBSID) then 
	if not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMIFS(iBSID) then 
		If IPMIDS =2 then projectIDSVal=" checked=""checked"""
		'If IPMIPS =3 then projectIPSVal=" checked=""checked"""
		If IPMIFS =3 then projectIFSVal=" checked=""checked"""
		'If IPMISM =4 then projectISMVal=" checked=""checked"""
		If IPMSPM =5 then projectSPMVal=" checked=""checked"""
	end if
	
	If IPMNo =1  then projectNOVal=" checked=""checked"""
	
	'if IPMInv="True" and IPMNo=0 and IPMIDS=0 and IPMIFS=0 and IPMISM=0 and IPMSPM=0 and not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMPRSS(iBSID) and not isIPMIFS(iBSID)  and iQPID<>"" and iQPID > 0  then 
	if IPMInv="True" and IPMNo=0 and IPMIDS=0 and IPMIFS=0 and IPMSPM=0 and not isIPMWCSS(iBSID) and not isIPMAPS(iBSID) and not isIPMIFS(iBSID)  and iQPID<>"" and iQPID > 0  then 
		projectunknown=" checked=""checked"""

	end if
   If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  & "- At Lineno529 QID="& iQPID
		End If

	Function ExtractSQProcCats()	
	Dim SQL, RS,Str,i,rsn,SQLn,idn
	On Error Resume Next
		    Dim DtBusWrkFlw,DiffBusWrkFlw
			DtBusWrkFlw = fncD_Configuration("BusWorkFlowValues")
			DiffBusWrkFlw = DateDiff("n",DtBusWrkFlw,dtRptDate)	

			SQLn="select id from tblRIR_SPSBusWorkflow  where Name='All'"
	    	Set RSn=cn.Execute(SQLn)
					If Not RSn.EOF Then
					idn=RSn(0)
					else
					idn=27
					end if
			
		'if lPL = 2 or lPL = 1 or lPL = 6 or lPL = 7 or lPL = 107 or lPL = 116 or lPL = 3 or lPL = 119 then  
            'SQL="  Select SPS.[ID],SPS.[RemoteID],SPS.[ActivityName],SPS.[PID],SPS.[Type],SPS.[Status],SPS.[Suffix],SPS.[Order], SPS.[Objective],SPS.[Description] from tblRIR_SPSData SPS inner join tblRIR_SPSDataMapping SPSMap on SPS.[ID]=SPSMap.[RefID] where SPSMap.[PLID] in ("&lPL&",0) order by [Type], [order] "
        '    SQL="spRP_SPSMapping " & lPL & ", " & DiffBusWrkFlw
       ' else
            SQL= " Select ID,'9999' as [RemoteID] ,Name as [ActivityName],0 as [PID], "
            SQL= SQL & " 'B2' as [Type],'1' as [Status],'1' as [Suffix],DisplayOrder as [Order],NULL as [Objective], NULL as [Description]  "
            SQL= SQL & " from tblRIR_SPSBusWorkflow where active=1 "
            IF (DiffBusWrkFlw < 1)  then 
                SQL= SQL & " and ID in (4,9) "
            ELSE
                SQL= SQL & " and ID NOT in (9) "
            End if 
			
			
			
			
            SQL= SQL & " UNION "
            SQL= SQL & " Select [ID],[RemoteID],[ActivityName], "
            SQL= SQL & " Case when [PID] = 0 Then BusWorkFlow Else PID end "
            SQL= SQL & " ,[Type],[Status],[Suffix],[Order], [Objective],[Description]  "
            SQL= SQL & " from tblRIR_SPSData "
			SQL= SQL & " union  Select [ID],[RemoteID],[ActivityName],"&idn&"  as [PID],[Type],[Status],[Suffix],[Order],  [Objective],'N8' from tblRIR_SPSData where [Type]='L2' "

            SQL= SQL & " order by  [TYPE] , [order] "
			
			
        'end if

		Set RS=cn.Execute(SQL)
		'response.Write(SQL)
		'response.write SafeNum(iQPID)
		'response.end
		If Not RS.EOF Then SQCats=RS.GetRows() else SQCats=NULL
		RS.Close
		Set RS=Nothing	
		Str=""
		For i=0 to UBound(SQCats,2)
			If SQCats(3,i)>0 and SQCats(5,i)=1 Then	
			    If trim(SQCats(4,i))="L2" Then Str=Str & "affservice.load('" & SQCats(3,i) & "','" & SQCats(0,i) & "','" & replace(SQCats(2,i),"'","\'") & "'," & SQCats(5,i) & ",true," & SQCats(7,i) & ");" & vbCRLF
				If trim(SQCats(4,i))="L3" Then Str=Str & "subaffservice.load('" & SQCats(3,i) & "','" & SQCats(0,i) & "','" & replace(SQCats(2,i),"'","\'") & "'," & SQCats(5,i) & ",true," & SQCats(7,i) & ");" & vbCRLF
				If trim(SQCats(4,i))="L4" Then Str=Str & "incident.load('" & SQCats(3,i) & "','" & SQCats(6,i) & "','" & replace(SQCats(2,i),"'","\'") & "'," & SQCats(5,i) & ",true," & SQCats(7,i) & ");" & vbCRLF
			End IF
		Next
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description & "- In Function ExtractSQProcCats QID="& iQPID
	End If	
		ExtractSQProcCats=Str	
	End Function

	Function ExtractSAXONCats()
	Dim SQL, RS,Str,i
	On Error Resume Next
	'SQL="Select SCS_PID,SCS_ID,SCS_Type,SCS_Name,convert(int,SCS_InActive) as Status from tlkpSQ_SCSCategories with (NOLOCK) Where SCS_Type in ('A','SA','I') Order by SCS_Type,SCS_PID,SCS_SortOrder,SCS_Name"
	SQL="select ID , 0 as PID , 'C' As Type,Description,status from tlkpSQCategories   where PLID = 127 and TYPE= 'F' "
    SQL= SQL & " Union All "
    SQL= SQL & " Select ID , PID ,'C' ,SubDescription , substatus  from tlkpSQSubCategories where PLID = 127 and substatus = 0"
	SQL= SQL & " Union All "
	SQL= SQL &"select ID , failureid as PID , 'C' ,Description,status from tlkpSQCategories   where PLID = 127 and TYPE= 'D' "
	
	Set RS=cn.Execute(SQL)
		If Not RS.EOF Then SAxCats=RS.GetRows() else SAxCats=NULL
		RS.Close
		Set RS=Nothing	
		Str=""
		For i=0 to UBound(SAxCats,2)
				If trim(SAxCats(2,i))="C" Then Str=Str & "saxcat.load('" & SAxCats(1,i) & "','" & SAxCats(0,i) & "','" & replace(SAxCats(3,i),"'","\'") & "'," & SAxCats(4,i) & ",true);" & vbCRLF
		Next
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description & "-In Fun ExtractSAXONCats QID="& iQPID
	End If	
		ExtractSAXONCats=Str	
	End Function
	
	
Function GETVALSUFFIX(Val)
	DIM spsl4ID1 
	On Error Resume Next
	set spsl4ID1 = server.createobject("ADODB.recordset")
    GETVALSUFFIX=0
    if not isnull(Val) Then 
        spsl4ID1.open "SELECT Suffix FROM tblRIR_SPSData WHERE ID =" & Val , cn 

			 if not spsl4ID1.EOF then
				SPS_ID1=spsl4ID1("Suffix")
			 else
				SPS_ID1=0
			end if
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  & "-In Fun GETVALSUFFIX QID="& iQPID
	End If
			GETVALSUFFIX=SPS_ID1
	End if		
End Function
	
	'getSelBox("L2",0,0,0)
	Function getSelBox(Ty,SeqNo,Val,PID)
		dim Str,JSFun,SNm,Nm,sINm,SANm,selStyle
		On Error Resume Next
		JSFun=""
		selStyle = "style='width: 150px'"
		if Ty="L2" Then 
			sANm="SQL3_"&SeqNo
			JSFun=" onchange='subaffservice.setSL(this,"&sANM&");' "
			selStyle = "style='width: 172px';"		
		End IF
        
        if Ty="B2" Then 
			sANm="SQL2_"&SeqNo

			'response.write ExtractSQBusinessWorkFlow()
            JSFun=" onchange='affservice.setSL(this,"&sANM&");' "
            selStyle = "style='width: 172px';"		
		End IF
				
		if Ty="L3" Then 
			sINm="SQL4_"&SeqNo
			JSFun=" onchange='incident.setSL(this,"&sINM&");' "	
			selStyle = "style='width: 140px'"
		End IF
		
		if Ty="L4" Then 
			JSFun=" onchange='getDescription(this);' "	
			selStyle = "style='width: 140px'"
		End IF
		 
		Nm="SQ"&Ty&"_"&SeqNo
		Str="<select Id = '"&Nm&"' Name='"&Nm&"' " & selStyle & " "& JSFun&">"    ' 
		Str=Str & getSelOptions(PID,Ty,Val)
		Str=Str&"</select>"

		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun getSelBox QID="& iQPID
	End If
		getSelBox=Str	
	End Function


	Function getSelOptions(PID,Ty,Val)
		Dim Str,i,cTy
		On Error Resume Next
		'response.write "PID = " & PID & "Ty=" & Ty & "val" &val
        if Ty="B2" or val="0.8" Then 		
		else
		getSelOptions="<option value=0>(Selection Required)"
		end if  
			Str=getSelOptions
		
		For i=0 to UBound(SQCats,2)-1
			If trim(Ty)=trim(SQCats(4,i))  Then 
			
			
				'If SQCats(3,i)=PID  and (SQCats(5,i)=1 or (SQCats(0,i)=Val)) Then
				
				
						If (SQCats(5,i)=1 or (SQCats(0,i)=Val))  Then
							if Ty="L2" and SQCats(9,i)="N8" then
							
							else if  Val="0.8" then 
							
							else
								Str=Str & "<option value="&SQCats(0,i)&" "
								If SQCats(0,i)=Val Then Str=Str & " Selected"
									Str=Str & ">"&SQCats(2,i)&vbCRLF
								end if
							end if
						End IF
			end IF	
		Next	
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description & "-In Fun getSelOptions QID="& iQPID 
	End If
		getSelOptions=Str	

	End Function

	Function getSelBox_New(Ty,SeqNo,Val,PID)
		dim Str,JSFun,SNm,Nm,sINm,SANm,selStyle,myname,mysymbol1
		On Error Resume Next
		JSFun=""
		selStyle = "style='width: 150px'"
		if Ty="C" Then 
			myname ="txtFailure"
			sINm= "txtSubFailure"
			JSFun=" onchange='if(this.value!= ""0""){saxcat.setSL(this,"&sINM&");selval(document.frmRIR.txtDamage,0);selval(document.frmRIR.txtSubDamage,0);}else{selval(document.frmRIR.txtSubFailure,1);selval(document.frmRIR.txtDamage,0);selval(document.frmRIR.txtSubDamage,0);}document.getElementById(""Saxon_Damage_cat"").style.display=""none"";document.frmRIR.ShowDamage.value = ""N"";' "	
			selStyle = "style=' '"
			
		Elseif Ty="C2" Then 
			myname= "txtSubFailure"
			sINm="txtDamage"
			JSFun=" onchange='if(this.value!= ""0""){saxcat.setSL(this,"&sINM&");selval(document.frmRIR.txtSubDamage,0);}else{selval(document.frmRIR.txtDamage,1);selval(document.frmRIR.txtSubDamage,0);}' "	
			selStyle = "style=' '"	
			
		Elseif Ty="C3" Then 
			myname= "txtDamage"
			sINm="txtSubDamage"
			JSFun=" onchange='if(this.value!= ""0""){saxcat.setSL(this,"&sINM&");}else{selval(document.frmRIR.txtSubDamage,1);}' "	
			selStyle = "style=' '"
		Elseif Ty="C4" Then 
			myname= "txtSubDamage"
			' JSFun=" onchange='saxcat.setSL(this,"&sINM&");' "	
			JSFun=""
			selStyle = "style=' '"		
			
		End IF
		
		Nm="SQ"&Ty&"_"&SeqNo
		Str="<select Id = '"&myname&"' Name='"&myname&"' " & selStyle & " "& JSFun&">"    ' 
		Str=Str & getSelOptions_New(PID,Ty,Val)
		Str=Str&"</select>" & mSymbol

		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  & "-In Fun getSelBox_New QID="& iQPID
	End If
		getSelBox_New=Str	
	End Function
	
	
	Function getSelOptions_New(PID,Ty,Val)
		Dim Str,i,cTy
		On Error Resume Next
		getSelOptions_New="<option value=0>(Selection Required)"
		'getSelOptions_New=""
		Str=getSelOptions_New
		For i=0 to UBound(SAxCats,2)-1
			'Print Ty & "-" & SQCats(0,i) & "-" & SQCats(1,i) & "-" & SQCats(2,i) 
			If trim(Ty)=trim(SAxCats(2,i)) Then 
				If SAxCats(1,i)=PID  and (SAxCats(4,i)=0 or (SAxCats(0,i)=Val)) Then
					Str=Str & "<option value="&SAxCats(0,i)&" "
					If SAxCats(0,i)=Val Then Str=Str & " Selected"
					Str=Str & ">"&SAxCats(3,i)&vbCRLF
				End IF
			end IF	
		Next	
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description & "-In Fun getSelOptions_New QID="& iQPID 
	End If
		getSelOptions_New=Str
	
	End Function
	
	
	Function getDescription(id)
	'response.write "hdsfjhfdshflkdsjfds"&id
		Dim Str,i
		On Error Resume Next
	 Str = " "
		For i=0 to UBound(SQCats,2)
			If SQCats(4,i) = "L4" and SQCats(5,i)=1 Then	
				Str = Str & "<div id='"& SQCats(6,i) &"' "
				Str =  Str & iif(SQCats(6,i)=id," style='display:block; height: 60px;'"," style='height: 60px; display:none;'") 
				
				
				'Str =  Str & iif(SQCats(0,i)=id," style='visibility:visible;'"," style='visibility:hidden;'")    
				
				'Str =  Str & " style='display:block;'" 				
				Str = Str &  ">"& SQCats(9,i) &"</div> "
				'Str = Str + 
				'response.write Str
				'response.end
			End IF

		Next

		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun getDescription QID="& iQPID
	End If
		getDescription = Str
	
	End Function 
	On Error Resume Next
	sqlcountvalue="select count(*) as countval from tblririnvtree with(NOLOCK) where InvType='JF' and QPID='" & SafeNum(iQPID) & "' and InvID in (select BCID FROM tlkp_InvbasicCause where SubBCID=0)"
            Set rsDatasqlcountvalue = cn.execute(sqlcountvalue)
            If not rsDatasqlcountvalue.eof then 
			countvalue=rsDatasqlcountvalue("countval")
			End if
			rsDatasqlcountvalue.Close
			set rsDatasqlcountvalue = nothing

		
	sSQLroleAttribID = "SELECT  BCID FROM tlkp_InvbasicCause WHERE Description ='Inadequate Engineering/Manufacturing' "
	Set rsDataroleAttribID = cn.execute(sSQLroleAttribID)
			 if not rsDataroleAttribID.EOF then
			BCID=rsDataroleAttribID("BCID")
			 else
			BCID=0
			end if

	sSQLrole = "select * from tblririnvtree with(NOLOCK) where InvType='JF' and InvID = "&BCID&" and QPID='" & SafeNum(iQPID) & "'"
			 Set rsDatarole = cn.execute(sSQLrole)
			 if rsDatarole.EOF then
			 ToCheck=true
			 else
			 ToCheck=false
			 end if
		 
		
	sSQLroleAttribIDMaint = "SELECT  BCID FROM tlkp_InvbasicCause WHERE Description in ('Inadequate Maintenance/Repair','Inadequate maintenance') "
	Set rsDataroleAttribIDMaint = cn.execute(sSQLroleAttribIDMaint)
	if not rsDataroleAttribIDMaint.eof then
	Do While Not rsDataroleAttribIDMaint.EOF
			 if not rsDataroleAttribIDMaint.EOF then
			BCIDMaint=rsDataroleAttribIDMaint("BCID")
			 else
			BCIDMaint=0
			 end if

	sSQLroleMaint = "select * from tblririnvtree with(NOLOCK) where InvType='JF' and InvID = "&BCIDMaint&" and QPID='" & SafeNum(iQPID) & "'"
			 Set rsDataroleMaint = cn.execute(sSQLroleMaint)
			 if rsDataroleMaint.EOF then
			 ToCheckMaint=true
			 else
			 ToCheckMaint=false
			 end if
			 If ToCheckMaint=false Then
             exit do
             End If
	
		rsDataroleAttribIDMaint.movenext
				Loop		 
End if 
	rsDataroleAttribIDMaint.close
	set rsDataroleAttribIDMaint=nothing
	
	
	sSQLFailureCatID = "select FailureCat from tblRIR_SQDM_ProcessStd where QPID= '" & SafeNum(iQPID) & "'"
	Set rsFailureCatID = cn.execute(sSQLFailureCatID)
			 if not rsFailureCatID.EOF then
			FailureCatID=rsFailureCatID("FailureCat")
			 else
			FailureCatID=0
			end if
	
	sSQLroleCatName = "select DMCatName from tlkpSQ_DMCats where DMCatID= '" & FailureCatID & "'"
			 Set rsDataroleCatName = cn.execute(sSQLroleCatName)
			 if not rsDataroleCatName.EOF then
			DMCatNameval=rsDataroleCatName("DMCatName")
			 else
			DMCatNameval=""
			end if
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-Line 1023 QID="& iQPID
	End If
	Function CheckRIRClose()
	

	    Dim Msg,intJobMangeGin,intBHAAftFail,intActPracID,intSpecIncCondID,intWorkProcID,intLacContID, sJobManagerAlias,intCatCount,SPFishing,intFailureDept,intFailureCat,intIncClass
		Dim pfJobMangeGin,pfBHAAftFail,pfActPracID,pfSpecIncCondID,pfWorkProcID,pfLacContID, pfJobManagerAlias,pfCatCount,pfSPFishing,pfFailureDept,pfFailureCat,pfIncClass
		Dim showAccRevMsg,DMRS_ACC,SQLstrRite,SubAssetRite,PartRite,SQLRite,SQLstrMaximo,SQLMaximo,SubAssetMaximo,PartMaximo,SubAssetRiteWPS,PartRiteWPS,SubAssetMaximoWPS,PartMaximoWPS
		Dim flgOthermockupEquipment,SQLOthermockupEquipment,SQLstrOthermockupEquipment,SQLWellCSURPart,SQLstrWellCSURPart,flgWellheadequipment,flgFracEquipment,SQLFracEquipment,SQLstrFracEquipment
		Dim SQLFracOperation,SQLstrFracOperation,flgFracOperation,flgProppant_Detail,SQLstrProppant_Detail,SQLProppant_Detail,flgCSURPart
		Dim WCSQPIDRite,WCSQPIDMaximo,WPSQPIDRite,WPSQPIDMaximo,CTSQPIDRite,SubAssetRiteCTS,PartRiteCTS,CTSQPIDMaximo,SubAssetMaximoCTS,PartMaximoCTS,SQLWellheadequipment,SQLstrWellheadequipment,SQLCSURVAL_Detail,SQLstrCSURVAL_Detail,flgCSURVAL_Detail
		On Error Resume Next
		set DMRS_ACC = server.createobject("adodb.recordset")
		
		'For WCS Part and SubAsset
		Set SQLRite = server.createobject("adodb.recordset")
		SQLstrRite = "select QPID,EQCatSubSubID,PartID from tblRIR_SQWCSImportedEquipments where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLRite = cn.execute(SQLstrRite)
		if not SQLRite.EOF then
			WCSQPIDRite=SQLRite("QPID")
			SubAssetRite=SQLRite("EQCatSubSubID")
			PartRite=SQLRite("PartID")
		else
			WCSQPIDRite=0
			SubAssetRite=0
			PartRite=0
		end if
		Set SQLRite = nothing
		
		Set SQLMaximo = server.createobject("adodb.recordset")
		SQLstrMaximo = "select QPID,EQCatSubSubID,PartID from tblRIR_SQWCSImportedEquipmentsMaximo where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLMaximo = cn.execute(SQLstrMaximo)
		if not SQLMaximo.EOF then
			WCSQPIDMaximo=SQLMaximo("QPID")
			SubAssetMaximo=SQLMaximo("EQCatSubSubID")
			PartMaximo=SQLMaximo("PartID")
		else
			WCSQPIDMaximo=0
			SubAssetMaximo=0
			PartMaximo=0
		end if
		Set SQLMaximo = nothing		
		SQLstrRite = ""
		SQLstrMaximo = ""
		
		'For WPS Part and SubAsset
		Set SQLRite = server.createobject("adodb.recordset")
		SQLstrRite = "select QPID,EQCatSubSubID,PartID from tblRIR_SQWPSImportedEquipments where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLRite = cn.execute(SQLstrRite)
		if not SQLRite.EOF then
			WPSQPIDRite=SQLRite("QPID")
			SubAssetRiteWPS=SQLRite("EQCatSubSubID")
			PartRiteWPS=SQLRite("PartID")
		else
			WPSQPIDRite=0
			SubAssetRiteWPS=0
			PartRiteWPS=0
		end if
		Set SQLRite = nothing
		
		Set SQLMaximo = server.createobject("adodb.recordset")
		SQLstrMaximo = "select QPID,EQCatSubSubID,PartID from tblRIR_SQWPSImportedEquipmentsMaximo where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLMaximo = cn.execute(SQLstrMaximo)
		if not SQLMaximo.EOF then
			WPSQPIDMaximo=SQLMaximo("QPID")
			SubAssetMaximoWPS=SQLMaximo("EQCatSubSubID")
			PartMaximoWPS=SQLMaximo("PartID")
		else
			WPSQPIDMaximo=0
			SubAssetMaximoWPS=0
			PartMaximoWPS=0
		end if
		Set SQLMaximo = nothing		
		SQLstrRite = ""
		SQLstrMaximo = ""
		
		'For CTS Part and SubAsset
		Set SQLRite = server.createobject("adodb.recordset")
		SQLstrRite = "select QPID,EQCatSubSubID,PartID from tblRIR_SQCTSImportedEquipments where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLRite = cn.execute(SQLstrRite)
		if not SQLRite.EOF then
			CTSQPIDRite=SQLRite("QPID")
			SubAssetRiteCTS=SQLRite("EQCatSubSubID")
			PartRiteCTS=SQLRite("PartID")
		else
			CTSQPIDRite=0
			SubAssetRiteCTS=0
			PartRiteCTS=0
		end if
		Set SQLRite = nothing
		
		Set SQLMaximo = server.createobject("adodb.recordset")
		SQLstrMaximo = "select QPID,EQCatSubSubID,PartID from tblRIR_SQCTSImportedEquipmentsMaximo  where QPID = '" & SafeNum(iQPID) & "'"
		Set SQLMaximo = cn.execute(SQLstrMaximo)
		if not SQLMaximo.EOF then
			CTSQPIDMaximo=SQLMaximo("QPID")
			SubAssetMaximoCTS=SQLMaximo("EQCatSubSubID")
			PartMaximoCTS=SQLMaximo("PartID")
		else
			CTSQPIDMaximo=0
			SubAssetMaximoCTS=0
			PartMaximoCTS=0
		end if
		Set SQLMaximo = nothing		
		SQLstrRite = ""
		SQLstrMaximo = ""
		
		
		
		
	if isCSUR(SQMappingID)=true and  bSQ=true  then 
		Set SQLWellheadequipment = server.createobject("adodb.recordset")				
			SQLstrWellheadequipment = "select count(*) from tbl_WellheadEquipment where qpid=" & SafeNum(iQPID) & "  and (SubAssemblyPN='' or  SubAssemblySN='')"
		Set SQLWellheadequipment = cn.execute(SQLstrWellheadequipment)
		if not SQLWellheadequipment.EOF then
			flgWellheadequipment=SQLWellheadequipment(0)			
		else
			flgWellheadequipment=0
		end if
		Set SQLWellheadequipment = nothing		
		SQLstrWellheadequipment = ""
		
					
			dim sql2,Service,IsIncidentPSD,ProductFamily,ProductSubFamily,EquipmentType,EquipmentSubType,flgclosechk
			sql2 = "SELECT * FROM tbl_ProductData WHERE QPID=" & SafeNum(iQPID)
			SET RS1=cn.execute((sql2))
			If Not RS1.EOF Then 
				Service = RS1("Service")
				IsIncidentPSD = RS1("IsIncidentPSD")
				ProductFamily = RS1("ProductFamily")
				ProductSubFamily = RS1("ProductSubFamily")
				EquipmentType = RS1("EquipmentType")
				EquipmentSubType = RS1("EquipmentSubType")
			Else
				flgclosechk="noclose"
				Service = 0
				IsIncidentPSD = 0
				ProductFamily = 0
				ProductSubFamily = 0
				EquipmentType = 0
				EquipmentSubType = 0
			End If
			RS1.Close
			Set RS1=Nothing  
		
			
		
		Set SQLWellCSURPart = server.createobject("adodb.recordset")				
		SQLstrWellCSURPart = "select count(*) from tbl_CSURPart  where qpid=" & SafeNum(iQPID) & " and  FailedPartPN<>'' and FailedPartSN<>'' and PartType>0 and DamageCategory>0 and DamageSubCategory>0 and ResponsbileFunction>0 and RootCause>0 "
		Set SQLWellCSURPart = cn.execute(SQLstrWellCSURPart)
		if not SQLWellCSURPart.EOF then
			flgCSURPart=SQLWellCSURPart(0)			
		else
			flgCSURPart=0
		end if
		Set SQLWellCSURPart = nothing		
		SQLstrWellCSURPart = ""
		
		
				
		Set SQLOthermockupEquipment = server.createobject("adodb.recordset")				
		SQLstrOthermockupEquipment = "select  count(*)  from tbl_OthermockupEquipment  where  qpid=" & SafeNum(iQPID) & " and   ((RTRIM(LTRIM(EquipmentPN)))='' or  (RTRIM(LTRIM(EquipmentSN)))='') "
		Set SQLOthermockupEquipment = cn.execute(SQLstrOthermockupEquipment)
		if not SQLOthermockupEquipment.EOF then
			flgOthermockupEquipment=SQLOthermockupEquipment(0)			
		else
			flgOthermockupEquipment=0
		end if
		Set SQLOthermockupEquipment = nothing		
		SQLstrOthermockupEquipment = ""
		
		
		
		Set SQLFracEquipment = server.createobject("adodb.recordset")				
		SQLstrFracEquipment = "select count(*) from tbl_FracEquipment where  qpid=" & SafeNum(iQPID) & " and   ((RTRIM(LTRIM(EquipmentPN)))='' or  (RTRIM(LTRIM(EquipmentSN)))='') "
		Set SQLFracEquipment = cn.execute(SQLstrFracEquipment)
		if not SQLFracEquipment.EOF then
			flgFracEquipment=SQLFracEquipment(0)			
		else
			flgFracEquipment=0
		end if
		Set SQLFracEquipment = nothing		
		SQLstrFracEquipment = ""
		
		
				
		Set SQLFracOperation = server.createobject("adodb.recordset")		
		SQLstrFracOperation = "select  *  from dbo.tbl_FracOperation where qpid=" & SafeNum(iQPID) & " and  PressurePumping>0  and FluidType>0 and PumpRate>=0 and PumpRate_Unit>0 and TotalProppant>=0 and TotalProppant_Unit>0 and TotalFluids>=0 and TotalFluids_Unit>=0; "
		Set SQLFracOperation = cn.execute(SQLstrFracOperation)
			if not SQLFracOperation.EOF then
					if SQLFracOperation("PressurePumping")=443 and trim(SQLFracOperation("OtherPPCompany"))="" then
					flgFracOperation=0
					else
					flgFracOperation=SQLFracOperation(0)
					end if						
			else
				flgFracOperation=0
			end if
		Set SQLFracOperation = nothing		
		SQLstrFracOperation = ""
		
		
		
		Set SQLProppant_Detail = server.createobject("adodb.recordset")				
		SQLstrProppant_Detail = "select  count(*)  from  tblProppant_Detail where  qpid=" & SafeNum(iQPID) & " and ProppantType>0 and ProppantSize>0 and  ProppantPercent>=0"
		Set SQLProppant_Detail = cn.execute(SQLstrProppant_Detail)
		if not SQLProppant_Detail.EOF then
			flgProppant_Detail=SQLProppant_Detail(0)			
		else
			flgProppant_Detail=0
		end if
		Set SQLProppant_Detail = nothing		
		SQLstrProppant_Detail = ""

		
		Set SQLCSURVAL_Detail = server.createobject("adodb.recordset")				
		SQLstrCSURVAL_Detail = "Select dbo.getcsurvalidation("&SafeNum(iQPID)&","&ProductFamily&","&ProductSubFamily&")"
		Set SQLCSURVAL_Detail = cn.execute(SQLstrCSURVAL_Detail)
		if not SQLCSURVAL_Detail.EOF then
			flgCSURVAL_Detail=SQLCSURVAL_Detail(0)			
		else
			flgCSURVAL_Detail=0
		end if
		Set SQLCSURVAL_Detail = nothing		
		SQLstrCSURVAL_Detail = ""

		Msg=""
		     If   Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideLegacyCSUR)) then
					if flgCSURVAL_Detail="TRUE" then
					Msg="\nPlease add Part for each Equipment and enter all part details which has blue asterisk."
					end if
			 end if
				

		 If   Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideLegacyCSUR)) then
			if ProductFamily=10  then				
				if flgWellheadequipment>0 then
					Msg=Msg&"\nPlease enter Equipment details which has blue asterisk"
					
				end if
			end if		
				
			if flgclosechk="noclose"  or (IsIncidentPSD=1 and (ProductFamily=0 or  ProductSubFamily=0 or  EquipmentType=0 or EquipmentSubType=0)) then			      
				Msg=Msg&"\nPlease enter CSUR SQ\\PQ details."				
			End if			
			
			if ProductFamily=11 or ProductFamily=12  then 			
				if flgOthermockupEquipment>0 then
					Msg=Msg&"\nPlease enter Equipment details which has blue asterisk."
				end if		
			end if
			
			if ProductFamily=9 and (ProductSubFamily=14 or ProductSubFamily=15) then			
				if flgOthermockupEquipment>0 then
					Msg=Msg&"\nPlease enter Equipment details which has blue asterisk."
				end if		
			end if
		
		
			if ProductFamily=9 and ProductSubFamily=13 then					
				if flgFracEquipment>0 then
					Msg=Msg&"\nPlease enter Equipment details which has blue asterisk."
				end if		
				if flgFracOperation=0 then
					Msg=Msg&"\nPlease enter Frac - Operation Data details which has blue asterisk."
				end if		
				if flgProppant_Detail=0 then
					Msg=Msg&"\nPlease enter Proppant Type,Proppant Size and Proppant Percent."
				end if	
			end if
		End If
	End If	
		If Not bNR Then
		
		
		'if ToCheck=false and ToCheckMaint=false and ShowDM() and bSQ and iClass=1 and countvalue <=1 then
		
		'if DMCatNameval <> "Execution - Inadequate generic maintenance procedure (EMS)" and DMCatNameval <> "Technology - Design/Reliability - DM equipment" and DMCatNameval <> "Execution - Inadequate local maintenance process (Maintenance)" and DMCatNameval <> "Execution - Maintenance Procedural adherence" and DMCatNameval <> "Execution - Maintenance Competency" and  DMCatNameval <> "" then Msg=Msg & "\nRoot Cause from D&M Post SCAT does not match with Root Cause from SCAT."

		'elseif ToCheck=false and ShowDM() and bSQ and iClass=1 and countvalue <=1 then
		
		'if DMCatNameval <> "Execution - Inadequate generic maintenance procedure (EMS)" and DMCatNameval <> "Technology - Design/Reliability - DM equipment" and  DMCatNameval <> "" then Msg=Msg & "\nRoot Cause from D&M Post SCAT does not match with Root Cause from SCAT. "
		
		'elseif ToCheckMaint=false and ShowDM() and bSQ and iClass=1 and countvalue <=1 then
		
		'if DMCatNameval <> "Execution - Inadequate local maintenance process (Maintenance)" and DMCatNameval <> "Execution - Maintenance Procedural adherence" and DMCatNameval <> "Execution - Maintenance Competency" and  DMCatNameval <> "" then Msg=Msg & "\nRoot Cause from D&M Post SCAT does not match with Root Cause from SCAT. "
		'End if
		
							If IsPers(RS) and LockCount <> 0 then 
								If bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true)) then    	                    
									   if iClass=1 And Not IsCompletePers() then Msg=Msg & "\nComplete entry of the Personal Loss on the Pers Loss Tab. "
								End if 
							End If
							
							
							'If ((isDMSQMapping(SQMappingID) or  isPF(iBSID)) and (bSQ and  ShowDM()) and isDMgss(iBSID)) Then  'isDM(lPL)
							If (isDMSQMapping(SQMappingID) and (bSQ and  ShowDM())) Then  'isDM(lPL) 'removed isPF(iBSID)) no pf tab is present and isDMgss(iBSID) as not required
								 if  Not isCompleteDM(iQPID) then Msg=Msg & "\nComplete entry for all mandatory fields of the D&M Details Tab. "		
							End If
							 
							if Not IsCompleteActionItems() then 
								 If bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true)) or (bSQ and Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarActionItemValidation))) then 'Validation added for SQ reports as well. 
								  Msg=Msg & "\nComplete entry of the Action Items on the Action Items Tab. "  
								 End if  
							End if 
							
							If iOpenActions<>0 then 
								   If bSQ=true then 
										 Msg=Msg & "\nClose open and pending action items." 
								   ElseIf bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true))then 
										 Msg=Msg & "\nClose open and pending action items."   
								   End if
							End IF  
						   
							If bSQ  and SegInv1 Then '***** (MS HIDDEN) - Commented complete If loop section  ***** 
								If not IsCompleteSeg() Then 
								  Msg=Msg & "\nComplete entry of Involved Segments/Functions on the Involved Segments/Functions Tab."
								End If
							End If
						
					 If iClass=1 Then
					 
						   If (IsEnv(RS)) then  
								 If bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true)) then 
								  'Env
									If Not IsCompleteEnv() then Msg=Msg & "\nComplete entry of the Environment Loss on the Env Loss Tab. "   
								 End if 
						   End IF
												
							If (IsAuto(RS)) then 
								 if bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true))    then 
								  'Auto 
									   If Not IsCompleteAuto() Then Msg=Msg & "\nComplete entry of the Auto Loss on the Auto Loss Tab."							
								 End IF
							End if 
							
							If (IsAssets(RS)) then 
								   if bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true))   then 
								  ' ASSET
									  If Not IsCompleteAssets() Then Msg=Msg & "\nComplete entry of the Asset Loss on the Asset Loss Tab."
								   End IF
						   End if 
						   
						   If (IsInfo(RS)) then 
									if bHSE and ((SLBInv=true) or (IndRec =true) or (SLBCon =true))  then 
								 'Information
									  If Not IsCompleteInfo() Then Msg=Msg & "\nComplete entry of the Information  Loss on the Info Loss Tab."		
									End if   
							End IF
							
					 End if
					 
					 If bWIBEvent and not isIPMWCI(GetSubBusinessSegID(lOrgNo)) and not isIPMRigM(GetSubBusinessSegID(lOrgNo)) and not isSPMSeg(SQMappingID) and not IsCompleteWB() then Msg=Msg & "\nComplete entry of the Well Barrier(s) on the Well Barrier Tab."	
					 If IsCompleteRIRRisk("P") and bWIBEvent Then Msg=Msg & "\nThe Potential Risk should not be less than -12 when Well Barrier Tab triggers a serious event or a catastrophic event."	
					 'If bWIBEvent and Not isCompleted("WBRisk") Then Msg=Msg & "\nThe Well Barrier Tab is showing red because the Potential Risk is less that -10."	
					 

			If CompOpen = 1 then Msg=Msg & "\nComplete entry of the Computer Loss Investigation Review on Investigation/Review tab."
			 
			If bHSE or (bSQ and Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarActionItemValidation))) Then		'Validation added for SQ reports as well. 
				'debugprint "Potential Risk:" & PRisk_C
				'debugprint "Residual Risk:" & RRisk_C
				if PRisk_C=0 then Msg=Msg & "\nComplete entry of the  Potential Risk on the Potential Risk page." 	
				if ChkPLoss_Fatality(rs("QID"))=0 then
					if ChkFAILEvent(rs("QID"),"P") and Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(HideFAILEvents)) then Msg=Msg & "\nThis event is defined as High Potential (HiPo), please categorize the event as being either a FAIL Safe or FAIL Lucky under Potential Risk page." 	
				end if
				if RRisk_C=0 then Msg=Msg & "\nComplete entry of the Residual Risk on the Residual Risk page." 
			End IF

			If bSQ  Then
				If isWPSSQMapping(SQMappingID) or isDataExistWPSSQ() Then				 'isWPS(iBSID)
					If Not isCompleteWPSSQ() And RS("SQSeverity") > 0  Then 
						Msg=Msg & "\nComplete WPS SQ Details page first."
					Elseif Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideFailureListing)) then
						if WPSQPIDRite <> 0 then
							If SubAssetRiteWPS = 0 Then Msg=Msg & "\nPlease select Sub Asset of Rite."	
							If PartRiteWPS = 0 Then Msg=Msg & "\nPlease select Part of Rite."
						End IF
						If WPSQPIDMaximo <> 0 then
								If SubAssetMaximoWPS = 0 Then Msg=Msg & "\nPlease select Sub Asset of Maximo."	
								If PartMaximoWPS = 0 Then Msg=Msg & "\nPlease select Part of Maximo."
						End IF
					End IF
				End If
			End If		

			If IsRequiredIPM() Then '@Visali Close Condition. 18-Dec-2009
				if not bIPM_C and IPMNo="0" then Msg=Msg & "\nComplete entry of " & IPMText_SPM(1) & " SQ and/or Project Details on " & IPMText_SPM(1) & " SQ/Proj Details tab." 
			End if
			
			
			If bHSE and Not IOGPrequired then
			    Msg=Msg & "\nComplete entry of IOGP tab."
			End if
			
			if not bContractor_C then Msg=Msg & "\nComplete entry of Contractor data on Contractor page."
			if Not (bRIRInv_C or (bHSE and bHSEInv_C)) then Msg=Msg & "\nComplete Investigation/Review page first."		
			If bSQ and iSQSev > 0 and (LockCountSQ<>0 or chkSQLockingMgmt())  then 
				If IsCompleteCost() = 0 then Msg=Msg & "\nAn SQ Non-Conformance Report by definition means that Loss has occurred. \nPlease go to the Loss Tabs and ensure that Red Money is entered where appropriate."   
			end If
			
		
			
			If bSQ and isREWSQMapping(SQMappingID) and (not(isSPWL(iBSID)) or (CDate(dtRptDatetmp) > CDate(comparedate)))  Then ''isREW(lPL)
				If Not isCompleteREWSQ(iQPID) and iClass<3  Then 
					Msg=Msg & "\nComplete WL SQ Details page first."	
				ElseIf Not isCompleteWLSW_Severity() Then 
					Msg=Msg & "\nComplete WL SQ Details page first."	
				ElseIf Not isCompleteWLSW_NPT() Then 
					Msg=Msg & "\nComplete WL SQ Details page first."	
				End If
				
			END If
			
			'OST Close Condition checking
			if Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(comparedateOST)) then
				If bSQ and isOSTPDPMapping(SQMappingID)  Then ''isREW(lPL)
					If Not isCompleteOSTPDP(iQPID) and iClass<3  Then 
						Msg=Msg & "\nComplete PDP SQ Details page first."	
					ElseIf Not isCompleteOSTPDP_Severity() Then 
						Msg=Msg & "\nComplete PDP SQ Details page first."	
					ElseIf Not isCompleteOSTPDP_NPT() Then 
						Msg=Msg & "\nComplete PDP SQ Details page first."	
					End If				
				END If
			End If	
			
			IF (CDate(dtRptDatetmp) > CDate(ComparedateSSONew)) Then 'For legacy RIR Closure logic handling	
				If bSQ and (isSSOiPLID(SQMappingID) or isSCSLocation(lOrgNo) or IsSpecialSCSOrg(iQPID)) Then
					If Not IsCompleteSCSNew() Then Msg=Msg & "\nComplete SSO SQ Details page first."	
				END IF
			END IF
			
			If bSQ and isSPWL(iBSID) and iClass<3 and (CDate(dtRptDatetmp) < CDate(comparedate)) Then
				If Not isCompleteSPWLSQ(iQPID) Then Msg=Msg & "\nComplete CWS SQ Details page first."	
			END IF
			If bSQ and rs("TCCInvolved") Then 
				If Not IsCompleteTCC() Then Msg=Msg & "\nComplete GRC Details page first."	
			End IF
			If bSQ And IsGSStab(SQMappingID) And ShowGSS Then 
				If Not IsCompleteGSS() Then Msg=Msg & "\nComplete GSS ML SQ Details page first."	
			End If
			If bSQ And  isCTSSQMapping(SQMappingID) And RS("SQSeverity") > 0  Then 		
				If Not isCompleteCTSSQ() Then 
					Msg=Msg & "\nComplete WIS SQ Details page first."
				Elseif Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideFailureListing)) then
					if CTSQPIDRite <> 0 then
						If SubAssetRiteCTS = 0 Then Msg=Msg & "\nPlease select Sub Asset of Rite."	
						If PartRiteCTS = 0 Then Msg=Msg & "\nPlease select Part of Rite."
					End IF
					If  CTSQPIDMaximo <> 0 then
							If SubAssetMaximoCTS = 0 Then Msg=Msg & "\nPlease select Sub Asset of Maximo."	
							If PartMaximoCTS = 0 Then Msg=Msg & "\nPlease select Part of Maximo."							
					End If
				End If
			End If


			If bSQ And IsGSStab(SQMappingID) And ShowGSS Then 
				If Not IsCompleteWellDataSQ() Then 
					If rs("SLBRelated") AND rs("RIRExternal")	 Then
						Msg=Msg & "\nComplete Welldata SQ Details page first." 
					End If
				End If
			End If

			If bSQ and isWCSSQMapping(SQMappingID) and iSQSev>1 Then   'isWCS(BusinessSegment)
				If Not isCompleteWCSSQ() Then 
					Msg=Msg & "\nComplete WIT SQ Details page first."
				Elseif Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideFailureListing)) then
					if WCSQPIDRite <> 0 then
						If SubAssetRite = 0 Then Msg=Msg & "\nPlease select Sub Asset of Rite."	
						If PartRite = 0 Then Msg=Msg & "\nPlease select Part of Rite."
					End If
					If  WCSQPIDMaximo <> 0 then
							If SubAssetMaximo = 0 Then Msg=Msg & "\nPlease select Sub Asset of Maximo."	
							If PartMaximo = 0 Then Msg=Msg & "\nPlease select Part of Maximo."
					End If
				End If
			End if
			
			'' ********************************************************************************************
			'' Added one condition for NPT <<2401608>>,if Time loss tab is incomplete the RIR can not be closed
			'' ********************************************************************************************
			If bSQ and iClass =1  and isTimeLossTab Then 
				If Not IsCompleteTime() Then Msg=Msg & "\nComplete Time Loss page first."	
			End IF
			'' ********************************************************************************************
			
			if isOPF(SQMappingID) and bSQ Then							
			IF (CDate(dtRptDatetmp) > CDate(comparedateOPF)) Then 'For legacy red tab logic handling
			Dim RSOPF, sSQLOPF ,sSQL2,RSOPF2,clientradioval,RSOPF3,sSQL3,opfprodata
			
			opfprodata=0
			opfepcc=0
			opfonm=0
			valopf=0
			sSQLOPF = "SELECT * from  tbl_OPFProjectData "
			sSQLOPF = sSQLOPF & " WHERE losstype>0 and  QPID=" & SafeNum(iQPID) & ""
			Set RSOPF = Server.CreateObject("ADODB.Recordset")
			RSOPF.Open sSQLOPF, cn	
			If Not RSOPF.EOF then 
			opfprodata=1			
			End IF
			
			sSQL2 = "select * from tbl_EPCCProjectLoss where QPID="&SafeNum(iQPID)
			Set RSOPF2 = Server.CreateObject("ADODB.Recordset")
			RSOPF2.Open sSQL2, cn
			If Not RSOPF2.EOF then 
			opfepcc=1
			End IF
			
			sSQL3 = "select * from tbl_OandMProjectLoss where QPID="&SafeNum(iQPID)
			Set RSOPF3 = Server.CreateObject("ADODB.Recordset")
			RSOPF3.Open sSQL3, cn
			If Not RSOPF3.EOF then  			
			opfonm=1
			End IF
			
			if iClass=2 or iClass=3 then
				if opfprodata=1 then
					valopf=1
				End if 
			else
				if opfprodata=1 and (opfepcc=1 or opfonm=1) then
					valopf=1
				End if 			
			END if
			
				If valopf=0 Then Msg=Msg & "\nPlease enter MPS SQ details."	
			
			
		End If
		END IF

	dim optcls
		optcls=request.form("optClass")
	
		if cint(RS("Class"))<>cint(optcls) and isOPF(SQMappingID) and bSQ and (opfepcc=1 or opfonm=1) then
			delopfdetail=1
		end if 
		
		if isOPF(SQMappingID) and bSQ and (opfepcc=1 or opfonm=1) then
			chkopfhazornear=1
		end if 
		
		
		
			''***********************************************************************************************************
			'Condition if RigNPT is greater than Time Loss NPT
			if isTimeLossTab and iClass =1 then
				if cdbl(rNPT) > cdbl(tNPT) then Msg=Msg & "\nRig NPT is greater than Time Loss NPT. Please correct the NPT in the appropriate tabs."	
			end if
			'Condition for EMS SWIFT #2448303 - Develop EMS SQ Tab
			'If bSQ and isEMS(lPL) Then
			If bSQ and isEMS(SQMappingID) Then
				If Not isCompleteEMS() Then Msg=Msg & "\nComplete EMS PQ Details page first."	
			END IF
			If bSQ and isSWACO(SQMappingID) Then
				If Not isCompleteSWACO() Then Msg=Msg & "\nComplete M-I Quality Details page first."	
			END IF
			
			'If bHSE and isSWACO(lPL) AND iClass=3 Then
				'If Not IsCompleteSwacoHOC() Then Msg=Msg & "\nComplete HOC Details page first."	
			'END IF
			
			'Added by Sagar For Is accountability field for D&M tab
			    If isDMSQMapping(SQMappingID) and ShowDM() then   'isDM(lPL)
					DMRS_ACC.Open "Select DMSQPS.* , L.* FROM tblRIR_SQDM_ProcessStd DMSQPS With (NOLOCK) LEFT JOIN tblLinks L ON  L.pQID = DMSQPS.QPID WHERE QPID='" & SafeNum(iQPID) & "' and  FailureCat in(Select DMCatID from tlkpSQ_DMCats with (NOLOCK) where DMCatName in ('Execution - Maintenance Procedural adherence','Planning - Drilling Engineering: DE Procedural Adherence','Execution - Field Procedural Adherence','Planning - Equipment configuration - Line Mgt Procedural adherence')) order by rptnumber desc", cn
					Do while  not DMRS_ACC.EOF 
						if  DMRS_ACC("AccRevwCompleted")  and isnull(DMRS_ACC("rptnumber")) Then
							showAccRevMsg = true   
						elseif  DMRS_ACC("AccRevwCompleted")  and  not isnull(DMRS_ACC("rptnumber")) Then 
							showAccRevMsg = false
							exit do 
						elseif 	not DMRS_ACC("AccRevwCompleted")  Then 
							showAccRevMsg = false
							exit do 
						End IF 	
						DMRS_ACC.movenext 
					Loop
					if showAccRevMsg = true  then  
						Msg=Msg & "\nComplete Accountability review completed details of D&M Details for D&M Incident Categorization page first."				
					End if 
				End if 
			'Till here 
			
			'Added By Deepak for D&M Tab Development
			If bSQ and isDMSQMapping(SQMappingID) and ShowDM()  Then  'isDM(lPL)
				Dim iReportType, NonToolFailure
				Set DMRS1 = Server.CreateObject("ADODB.Recordset")
				DMRS1.Open "Select * FROM tblRIR_SQDM_Main With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not DMRS1.EOF) Then
					intJobMangeGin = DMRS1("JobManager_GIN")
					sJobManagerAlias = DMRS1("JobManager_Alias")
					intBHAAftFail = DMRS1("BHAAfterFailure")
					iReportType = DMRS1("ReportType")
					NonToolFailure = DMRS1("NonToolFailure")
					If isnull(NonToolFailure) Then NonToolFailure=0
				End If
				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then
					If SafeNum(intJobMangeGin) = 0 and len(trim(sJobManagerAlias)) = 0 Then Msg=Msg & "\nPut a Valid Job Manager in D&M Details Tab."
					If SafeNum(intBHAAftFail) = 0 Then Msg=Msg & "\nSelect BHA Activity After Failure."
				End If
				DMRS1.Close

				DMRS1.Open "Select * FROM tblRIR_SQDM_Incidents With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' and ClassId = "  &  SafeNum(iClass) & " And (Inc_ActPrac_ID != 0)", cn
				If (not DMRS1.EOF) Then
					intActPracID = DMRS1("Inc_ActPrac_ID")
					intSpecIncCondID = DMRS1("SpecInc_Cond_ID")
				End If
				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then			
					If (SafeNum(intActPracID) = 0) Then Msg=Msg & "\nIncident Categeory, D & M Specific Category is required if Class = Accident &  SQ Severity = CMS and Source = eTrace." 
				End If
				DMRS1.Close
				
				DMRS1.Open "Select * FROM tblRIR_SQDM_ProcessStd With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not DMRS1.EOF) Then
					intFailureDept  = DMRS1("FailureDept")
					intFailureCat   = DMRS1("FailureCat")
					intIncClass		= DMRS1("PostRevwIncClassification")
					
				End If
				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then
					If (SafeNum(intFailureDept) = 0 or SafeNum(intFailureCat) = 0 or SafeNum(intIncClass) = 0) Then Msg=Msg & "\nPOST REVIEW INCIDENT CLASSIFICATION or RESPONSIBLE PARTY FOR SYSTEM FAILURE or WHERE DID THE SYSTEM FAIL is required if Class = Accident &  SQ Severity = CMS and Source = eTrace."
				End If
				DMRS1.Close
				
				DMRS1.Open "Select count(*) AS Count FROM tblRIR_SQDM_Incidents With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND Inc_ActPrac_ID = 11 AND SpecInc_Cond_ID = 58 ", cn
				 If (not DMRS1.EOF) Then
					intCatCount = DMRS1("Count")
				 END if				
				DMRS1.Close

				If intCatCount=0 Then 
				DMRS1.Open "SELECT DMCatName as CatName from tlkpSQ_DMCats where DMCatID In (select Inc_ActPrac_ID  FROM tblRIR_SQDM_Incidents with (NOLOCK) WHERE QPID=" & SafeNum(iQPID) & ")", cn
				If Not DMRS1.EOF Then
					Do While Not DMRS1.EOF
					
						If DMRS1("CatName") = "Stuck pipe" Then
						
							intCatCount = 1		
						End If
						
						
						DMRS1.movenext
					Loop 
					
				End If 
				DMRS1.Close
				End IF


				Dim R1Count
				DMRS1.Open "SELECT count(*) as Count FROM tblRIR_SQDM_EquipSum WHERE SECTIONNO in (1,2,3)AND QPID ='" &  SafeNum(iQPID) & "'", cn
				 If (not DMRS1.EOF) Then
					R1Count = DMRS1("Count")
				 END if				
				DMRS1.Close

				If NonToolFailure <> True Then '// NonToolFailure = No
					If ((Source = 14 Or Source = 1 or Source = 25) and ((iReportType = 408 Or iReportType = 410) and R1Count > 0) and intCatCount = 0 and iClass = 1) Then
						DMRS1.Open "Select Count(*) AS ChkCount FROM tblRIR_SQDM_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND FailureIndicated = 1", cn
						If DMRS1("ChkCount") = 0 Then Msg=Msg & "\nAt Least one Failure Indicator is required for Equipment Summary Details."
						DMRS1.Close
					End If
					
					If ((Source = 14 Or Source = 1 or Source = 25) and ((iReportType = 408 Or iReportType = 410) and R1Count > 0) and intCatCount = 0 and iClass = 1) Then
						DMRS1.Open "Select Count(*) AS ChkCount FROM tblRIR_SQDM_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND workorderstatus = 'OPEN'", cn
						If DMRS1("ChkCount") > 0 Then Msg=Msg & "\nAll of the Field Failure work orders should be closed or cancelled."
						DMRS1.Close
					End If	
					
					
					If ((Source = 14 Or Source = 1 or Source = 25) and ((iReportType = 408 Or iReportType = 410) and R1Count > 0) and intCatCount = 0 and iClass = 1) Then
						
						DMRS1.Open "Select Count(*) AS ChkTFCount FROM tblRIR_SQDM_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND ToolFailure = 1", cn
						If DMRS1("ChkTFCount") = 0 Then Msg=Msg & "\nAt Least one Tool Failure is required for Equipment Summary Details."
						DMRS1.Close
					
						DMRS1.Open "Select Count(*) AS ChkCFCount FROM tblRIR_SQDM_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND ComponentFailure = 1", cn
						If DMRS1("ChkCFCount") = 0 Then 
							Msg=Msg & "\nAt Least one Component Failure  is required for Equipment Summary Details."
							DMRS1.Close
						Else
							DMRS1.Close
							 
							if Source = 25 then
							DMRS1.Open "spSQ_GetDMComponentFailure_FDP " & SafeNum(iQPID)
							else
							DMRS1.Open "spSQ_GetDMComponentFailure " & SafeNum(iQPID)							
							end if 
							If DMRS1("ChkCount") = 0 Then Msg=Msg & "\nAt least one Component Failure is required for a selected Tool Failure for Equipment Summary Details."
							'DMRS1.Open "select count(id) as ChkIds from tblRIR_SQDM_EquipSum where qpid='" & SafeNum(iQPID) & "' and toolfailure=1 and parent_id=0 and id not in (select Case When parent_id = 0 Then ID Else PARENT_ID End from tblRIR_SQDM_EquipSum where qpid='" & SafeNum(iQPID) & "' and componentfailure=1)"
							'If DMRS1("ChkIds") > 0 Then Msg=Msg & "\nAt least one Component Failure is required for a selected Tool Failure for Equipment Summary Details."
							DMRS1.Close
						End If

					End If
				End If	
				
				DMRS1.Open "Select SPFishing AS Fishing FROM tblRIR_SQDM_StuckPipedetails With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not DMRS1.EOF) Then
					SPFishing = DMRS1("Fishing")
				End if				
				DMRS1.Close
				
				If (SPFishing = 2517 or SPFishing = 2518 or SPFishing = 2525) Then			
					Msg=Msg & "\nCannot close incident if Planned, In Progress and Pending is chosen as Fishing."
				End If
				
			End If
			
			If bSQ and isPF(lBSegmentID) Then
				Dim pfReportType, PFNonToolFailure
				Set PFRS1 = Server.CreateObject("ADODB.Recordset")
				PFRS1.Open "Select * FROM tblRIR_SQPF_Main With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not PFRS1.EOF) Then
					pfJobMangeGin = PFRS1("JobManager_GIN")
					pfJobManagerAlias = PFRS1("JobManager_Alias")
					pfBHAAftFail = PFRS1("BHAAfterFailure")
					pfReportType = PFRS1("ReportType")
					PFNonToolFailure = PFRS1("NonToolFailure")
					If isnull(PFNonToolFailure) Then PFNonToolFailure=0
				End If
				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then
					If SafeNum(pfJobMangeGin) = 0 and len(trim(pfJobManagerAlias)) = 0 Then Msg=Msg & "\nPut a Valid Job Manager in PathFinder Details Tab."
					If SafeNum(pfBHAAftFail) = 0 Then Msg=Msg & "\nSelect BHA Activity After Failure."
				End If
				PFRS1.Close

				PFRS1.Open "Select * FROM tblRIR_SQPF_Incidents With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' and ClassId = "  &  SafeNum(iClass) & " And (Inc_ActPrac_ID != 0)", cn
				If (not PFRS1.EOF) Then
					pfActPracID = PFRS1("Inc_ActPrac_ID")
					pfSpecIncCondID = PFRS1("SpecInc_Cond_ID")
				End If
				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then
					If (SafeNum(pfActPracID) = 0) Then Msg=Msg & "\nIncident Categeory, PathFinder Specific Category is required if Class = Accident &  SQ Severity = CMS and Source = eTrace."
				End If
				PFRS1.Close
				
				PFRS1.Open "Select * FROM tblRIR_SQPF_ProcessStd With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not PFRS1.EOF) Then
					pfIncClass    = PFRS1("PostRevwIncClassification")		
					pfFailureDept = PFRS1("FailureDept")
					pfFailureCat  = PFRS1("FailureCat")
				End If

				If (iClass = 1) and (iSQSev >1) and (Source = 14) Then				
					If (SafeNum(pfFailureDept) = 0 or SafeNum(pfFailureCat) = 0 or SafeNum(pfIncClass) = 0) Then Msg=Msg & "\nPOST REVIEW INCIDENT CLASSIFICATION or RESPONSIBLE PARTY FOR SYSTEM FAILURE or WHERE DID THE SYSTEM FAIL is required if Class = Accident &  SQ Severity = CMS and Source = eTrace."
				End If
				PFRS1.Close
				
				PFRS1.Open "Select count(*) AS Count FROM tblRIR_SQPF_Incidents With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND Inc_ActPrac_ID = 11 AND SpecInc_Cond_ID = 58 ", cn
				 If (not PFRS1.EOF) Then
					pfCatCount = PFRS1("Count")
				 END if				
				PFRS1.Close

				If pfCatCount=0 Then 
				PFRS1.Open "SELECT PFCatName as CatName from tlkpSQ_PFCats where PFCatID In (select Inc_ActPrac_ID  FROM tblRIR_SQPF_Incidents with (NOLOCK) WHERE QPID=" & SafeNum(iQPID) & ")", cn
				If Not PFRS1.EOF Then
					Do While Not PFRS1.EOF
						If PFRS1("CatName") = "Stuck Pipe" Then 
							intCatCount = 1		
						End If
						PFRS1.movenext
					Loop 
				End If 
				PFRS1.Close
				End If
				

				Dim P1Count
				PFRS1.Open "SELECT count(*) as Count FROM tblRIR_SQPF_EquipSum WHERE SECTIONNO in (1,2,3)AND QPID ='" &  SafeNum(iQPID) & "'", cn
				 If (not PFRS1.EOF) Then
					P1Count = PFRS1("Count")
				 END if				
				PFRS1.Close
				
				If ((Source = 14 Or Source = 1 or Source = 25) and ((pfReportType = 408 or pfReportType = 410) and P1Count > 0) and pfCatCount = 0) Then	
				PFRS1.Open "Select Count(*) AS ChkCount FROM tblRIR_SQPF_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND FailureIndicated = 1", cn
					If PFRS1("ChkCount") = 0 Then Msg=Msg & "\nAt Least one Failure Indicator is required for Equipment Summary Details."
					PFRS1.Close
				End If

				If ((Source = 14 Or Source = 1 or Source = 25) and ((pfReportType = 408 or pfReportType = 410) and P1Count > 0) and pfCatCount = 0) Then		
					PFRS1.Open "Select Count(*) AS ChkTFCount FROM tblRIR_SQPF_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND ToolFailure = 1", cn
					If PFRS1("ChkTFCount") = 0 Then Msg=Msg & "\nAt Least one Tool Failure is required for Equipment Summary Details."
					PFRS1.Close
						
				End If


				If PFNonToolFailure <> True Then 
					If ((Source = 14 Or Source = 1 or Source = 25) and ((pfReportType = 408 or pfReportType = 410) and P1Count > 0) and pfCatCount = 0) Then	

						PFRS1.Open "Select Count(*) AS ChkCFCount FROM tblRIR_SQPF_EquipSum With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "' AND ComponentFailure = 1", cn
						If PFRS1("ChkCFCount") = 0 Then 
							Msg=Msg & "\nAt Least one Component Failure is required for Equipment Summary Details."
							PFRS1.Close
						Else
							PFRS1.Close
							PFRS1.Open "select count(a.id) as ChkIds from tblRIR_SQPF_EquipSum a inner join tblRIR_SQPF_Main b on a.qpid = b.qpid and b.NonToolFailure=0 where a.qpid='" & SafeNum(iQPID) & "' and a.toolfailure=1 and a.parent_id=0 and a.id not in (select Case When parent_id = 0 Then ID Else PARENT_ID End from tblRIR_SQPF_EquipSum where qpid='" & SafeNum(iQPID) & "' and componentfailure=1)"
							If PFRS1("ChkIds") > 0 Then Msg=Msg & "\nAt least one Component Failure is required for a selected Tool Failure for Equipment Summary Details."
							PFRS1.Close
						End If
					End If
				End If
				
				PFRS1.Open "Select SPFishing AS Fishing FROM tblRIR_SQPF_StuckPipedetails With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'", cn
				If (not PFRS1.EOF) Then
					pfSPFishing = PFRS1("Fishing")
				End if				
				PFRS1.Close
				
				If (pfSPFishing = 2517 or pfSPFishing = 2518 or pfSPFishing = 2525) Then			
					Msg=Msg & "\nCannot close incident if Planned, In Progress and Pending is chosen as Fishing."
				End If
				
			End If
			
		End IF

	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun CheckRIRClose QID="&iQPID
	End If
		CheckRIRClose=Msg
	End Function

	'****************************************************************************************
	'1. Function/Procedure Name          : getRigNPT
	'2. Description           	         : To get RigNPT value from IPM tab
	'3. Calling Forms:   	             : RIRdsp.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   10-Nov-2009			    Visali Grandhi	         	 Added for IPM tab Change.
	'****************************************************************************************
	Function getRigNPT()
		Dim rs,sql
		On Error Resume Next
		getRigNPT = 0
		sql = "Select RigNPT from tblRIRipm with (NOLOCK) WHERE QPID=" &  SafeNum(iQPID) &  ""
		set rs = cn.execute(sql)
		if not rs.eof then
			getRigNPT = SafeNum(trim(RS("RigNpt")))
		end if
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  & "-In Fun getRigNPT QID="&iQPID
	End If

	End Function

	Function imploc()

	'****************************************************************************************
	'1. Function/Procedure Name          : imploc
	'2. Description           	         : To check ipm location
	'3. Calling Forms:   	             : RIRdisp.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   5-Aug-2009			    Nilesh Naik	         	 added for NPT <<2401608>> 

	'****************************************************************************************
	Dim blnimploc
	On Error Resume Next
	blnimploc = false 

	' if ((not (isIPMSeg(lPL))) and (iClass = 1)  and (bSQ) ) then 
	 if ((not (isIPMSeg(SQMappingID))) and (iClass = 1)  and (bSQ) ) then 
		 blnimploc = false
	  else
		blnimploc = true
	 end if  
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  & "-In Fun imploc QID="&iQPID
	End If	 
	   imploc = blnimploc	
	End Function

	Function chktimedata()

	'****************************************************************************************
	'1. Function/Procedure Name          : chktimedata
	'2. Description           	         : To check Time loss data present or not for this RIR
	'3. Calling Forms:   	             : RIRdisp.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   5-Aug-2009			    Nilesh Naik	         	 added for NPT <<2401608>> 

	'****************************************************************************************
	Dim blnflag ,strsql,RS
	On Error Resume Next
	blnflag = false
	If Not bNR Then 
	  

	   
	 strSQL = "SELECT  1  from tblRIRtime a with (NOLOCK) INNER JOIN TLKPLOSSSUBCATEGORIES b with (NOLOCK) ON  b.id  = a.type " 
	 strSQL = strSQL + " WHERE b.losscatid = 7 AND qpid = " &  SafeNum(iQPID) &  ""
	   
	   Set RS = Server.CreateObject("ADODB.Recordset")		
	   RS.Open strSQL, cn
		
		If (RS.EOF or RS.BOF) then 
			blnflag = false 
		else
			blnflag = true 
		End if
	end if 

	
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun chktimedata QID="&iQPID
	End If
	chktimedata = blnflag 
	  
	End Function

	Sub GetMatrixData()
		Dim cmdMatrix, rsMatrix, Cn
		On Error Resume Next
		Set Cn = GetNewCn()
		Cn.CursorLocation = 3
		Set rsMatrix = Server.CreateObject("ADODB.RecordSet")
		Set cmdMatrix = Server.CreateObject("ADODB.Command")
		
		With cmdMatrix
			.ActiveConnection = Cn
			.CommandType = adCmdStoredProc
			.CommandText = "SPRIR_SeverityCheck"
			
			.Parameters.Append .CreateParameter("@QPID", adInteger, adParamInput, ,iQPID)

			Set rsMatrix = .Execute()
		End With
		Response.Write "<Script Language=JavaScript>" & vbCrLf
		IF NOT rsMatrix.EOF then 
			 Response.Write "var intRseverity = "&rsMatrix("severity") &";" & vbCrLf
		end if     
		Response.Write "</Script>" & vbCrLf
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In SubFun GetMatrixData QID="&iQPID
	End If
	End Sub	

	Function Severitycheck (intqpid,intSeverity)
	'****************************************************************************************
	'1. Function/Procedure Name          : Severitycheck
	'2. Description           	         : To check severity is matching with losses in Time-loss tab
	'3. Calling Forms:   	             : RIRDSP.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   28-Aug-2009			    Nilesh Naik	        	   Added for NPT SWIFT #2401608 

	'****************************************************************************************
	On Error Resume Next
	Dim Cn, rs_Severity, cmdErrorMsg
				
			conn.CursorLocation = 3
			Set rs_Severity = Server.CreateObject("ADODB.RecordSet")
			Set cmdErrorMsg = Server.CreateObject("ADODB.Command")
			
			With cmdErrorMsg
				.ActiveConnection = conn
				.CommandType = adCmdStoredProc
				.CommandText = "SPRIR_SeverityCheck"
				
				.Parameters.Append .CreateParameter("@QPID", adInteger, adParamInput, ,intqpid)
				.Parameters.Append .CreateParameter("@RIRSeverity", adInteger, adParamInput, ,intSeverity)
				
				Set rs_Severity = .Execute()
			End With
			If Err.Number <> 0 Then
			LogEntry 2,"RIRDsp.asp",err.Description &"-In Fun Severitycheck QID="&iQPID
			End If
	
			if (rs_Severity.state = 1) then 
				If not rs_Severity.eof then
					 Severitycheck = rs_Severity("cond")
				end if
			end if
			
			Set rs_Severity = Nothing
			'Cn.Close
			'Set Cn = Nothing
					 
	
			
	End Function

    SLBInvment = 0
		 If SLBInv Then
		 If IndRec Then SLBInvment = 1 Else SLBInvment = 2
			Else
					If SLBCon Then SLBInvment = 3 Else SLBInvment = 4
			End If	
					 
	If SLBRel Then
		If External Then SQInvment = 2 Else SQInvment = 1
		Else
			SQInvment = 3
	End If
	''end changes for NPT  
	Function getSupplier(CID)
		Dim Str,intEntry,strFcn, valtxtpop
		On Error Resume Next
		valtxtpop = ""    
		CID = split(CID, ",")            
		Squery="Select ContractorName,SeqID from tblRIRContractors C With (NOLOCK)  inner join tblcontractorslist CL With (NOLOCK) on C.ContractorId = CL.ContractorId"		            
		Squery=Squery& " WHERE QPID=" & iQPID & " "
		Squery=Squery& " Order by SeqID"
			
		set innerRS = cn.execute(Squery)
		If NOT innerRS.EOF or NOT innerRS.BOF Then
			valtxtpop = innerRS("ContractorName")
		Else
			valtxtpop = "No Supplier Selected"
		End If
	  
	  
		Str = "<input onclick='showPos();' onmouseover='setimage(1);' onmouseout='setimage(2);' style='border: 1px solid #ddd; border-color:#ADA9A9; border-right: 0;background: #fff url(../Images/dropdown.png) no-repeat center right;' name='txtPopup' readonly id='txtPopup' value='"&safedisplay(valtxtpop)&"'>"  

		Str = Str & "<BR><span style='vertical-align:top;'><DIV id='txtTPSupplier' style='display: none; position: absolute; border: solid black 1px; "

		Str = Str & "padding: 10px; background-color:White; width:auto;'>" 
		Str = Str & "<span style='cursor:default;display: inline-block; width:150px;' onmouseover='changeSpanColor(this,1);' onmouseout='changeSpanColor(this,2);' "
		Str = Str & "onclick='SettxtPopup("""");'>No Supplier Selected</span><br />"
		Str = Str & "<span style='cursor:default;display: inline-block; width:150px;' onmouseover='changeSpanColor(this,1);' onmouseout='changeSpanColor(this,2);' "
		Str = Str & "onclick='Supplier_onchange();'>Select Supplier</span><br />"
		If ubound(CID)>0 Then 
			For intEntry = lbound(CID) To ubound(CID)-1
				strFcn = "onclick='SettxtPopup(this);'"            
				Str=Str & "<Input Name='chkTPSupplier' "&strFcn&" Type='CheckBox' value='"&CID(intEntry)& ":" & Replace(GetSupplierName(CID(intEntry)),"'","") 
				Str=Str & "' Checked >" & GetSupplierName(CID(intEntry)) & "</input></BR>" &vbCRLF            
			Next		    
		End IF	        
	   
		Str=Str & " </div></span>"
		Response.Write Str
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  & "-In Fun getSupplier QID="&iQPID
	End If

	End Function



%>
<html>
<head>
    <link rel="stylesheet" href="../style/QUEST.css" type="text/css">

    <script src="<%=VarJquery%>"></script>

    <script language="JavaScript1.2" src="../Calendar1-82.js"></script>

    <script language="JavaScript" src="../inc/timepicker.js"></script>

    <script type="text/javascript">



        function showPos() {
            var el;
            el = document.getElementById('txtTPSupplier');
            if (el.style.display == 'block') {
                el.style.display = "none";
            } else {
                el.style.display = "block";
            }
        }
        function changeSpanColor(obj, inOut) {
            if (inOut == '1') {
                obj.style.backgroundColor = '#39f';
                obj.style.color = 'white';
            } else {
                obj.style.backgroundColor = 'White';
                obj.style.color = 'Black';
            }
        }

        function setimage(mode) {

            if (mode == 1) {
                document.getElementById("txtPopup").style.background = "#fff url(../Images/dropdown_hover.png) no-repeat right center";
            }
            else {
                document.getElementById("txtPopup").style.background = "#fff url(../Images/dropdown.png) no-repeat right center";
            }
        }



        function hideDropdown() {

            if (document.activeElement.name == 'imgPopup' || document.activeElement.id == 'txtPopup' || document.activeElement.name == 'chkTPSupplier' || document.activeElement.id == 'chkTPSupplier') {

            }
            else {
                var el = document.getElementById('txtTPSupplier');
                el.style.display = "none";
            }

        }

        function hidediv(id) {
            if (document.getElementById) { // DOM3 = IE5, NS6
                if (document.getElementById(id) != null) {
					document.getElementById(id).style.display = 'none';
				}
            }
            else {
                if (document.layers) { // Netscape 4
                    document.id.display = 'none';
                }
                else { // IE 4
                    document.all.id.style.display = 'none';
                }
            }
        }
        //swi changes
        function showdiv(id) {

            if (document.getElementById) { // DOM3 = IE5, NS6
                document.getElementById(id).style.display = 'block';
            }
            else {
                if (document.layers) { // Netscape 4
                    document.id.display = 'block';
                }
                else { // IE 4
                    document.all.id.style.display = 'block';
                }
            }
        }

        function SettxtPopup(status) {
            if (status == '') {
                document.getElementById('txtPopup').value = 'No Supplier Selected'
                var chk_arr = frmRIR.chkTPSupplier;
                if (chk_arr) {
                    var chklength = chk_arr.length;
                    if (chklength == undefined) {
                        frmRIR.chkTPSupplier.checked = false;
                    }

                    for (k = 0; k < chklength; k++) {
                        chk_arr[k].checked = false;
                    }
                }
            }
            else {
                document.getElementById('txtPopup').value = ''
                var chk_arr = frmRIR.chkTPSupplier;
                var chklength = chk_arr.length;
                if (chklength == undefined) {
                    if (frmRIR.chkTPSupplier.checked == false) {
                        document.getElementById('txtPopup').value = 'No Supplier Selected'
                    }
                    else {
                        var arrVal = chk_arr.value.split(':');
                        document.getElementById('txtPopup').value = arrVal[4];
                        return;
                    }
                }
                for (k = 0; k < chklength; k++) {
                    if (chk_arr[k].checked) {
                        var arrVal = chk_arr[k].value.split(':');
                        var setcheck = 1;
                        document.getElementById('txtPopup').value = arrVal[4];
                        return;
                    }
                    if (setcheck != 1) {
                        document.getElementById('txtPopup').value = 'No Supplier Selected'
                    }
                }
            }
        }
		
    </script>

    <!-- #INCLUDE FILE="../Inc_Java_Functions.asp"-->

    <script id="clientEventHandlersJS" language="JavaScript">
		<!--
    var isSubmit=false;
    var DisG1=0,DisG2=0;
    var globalProc = '';
    var globalPer='';
            
    function showSQMatrix() {
        window.open('<%=sSQSeverityMatrix%>','Severity','height=600,width=600,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
    }
	function showSPXexplorer() {
        window.open('https://spx.slb.com/MetroMap/Details?Type=KeyProcessAndFunction' ,'Severity','height=600,width=600,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
    }
    function showBusWorkFlow() {
        window.open('<%=BusWrkFLowHubLink%>','Business Workflow','height=470,width=700,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
    }
    function showHSEMatrix() {
        window.open('<%=sHSESeverityMatrix%>','Severity','height=600,width=600,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
    }

    function showHazardCategory() {
        window.open('HazardCategory.asp','Severity','height=600,width=600,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1')
    }

			

    function LossCat_A2_onclick() {
        if(document.frmRIR.LossCat_A2.checked==true)
        {
            window.alert('               MEDICAL CONFIDENTIALITY REQUIREMENTS\n\nPlease make sure you do NOT enter the name of the sick person in the \nbrief, detailed descriptions or anywhere else within this RIR. \n\nPersonnel names under Personnel Loss will default to "Medically Confidential".');				
        }
        return true 
    }

    function inputTime() {
        if (document.frmRIR.txtEvTime.value == "24Hr")
        { 
            setLocalTime(document.frmRIR.txtEvTime);
        }
    }

    function inputDate() {
        if (document.frmRIR.txtEvDate.value == "mmm dd, yyyy") 
        { 
            setLocalDate(document.frmRIR.txtEvDate);
        }
    }

    function OnChange_SiteType(obj,onchng){
        var ty=obj.value;
        var f=window.document.frmRIR;
        if (ty.indexOf("RIG")>0){				
            f.txtCRMRigID.disabled=false;
            showhideElement("req_RigName","show")
            //showhideElement("req_SiteName","hide")
        }else if (ty.indexOf("OPT")>0){
            f.txtCRMRigID.disabled=false;
            showhideElement("req_RigName","hide")
            //showhideElement("req_SiteName","hide")
            document.getElementById("wellsitename").innerHTML = 'N/A'; 
        }else{
            f.txtCRMRigID.disabled=true;
            showhideElement("req_RigName","hide")
            //showhideElement("req_SiteName","show")
            document.getElementById("wellsitename").innerHTML = 'N/A';
        }
		<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
        if (onchng ==1){
             <%IF ((DiffEvtClass >= 1) and  (bHSE))  then%>
            SetEventCategorisation()
            <%end if%>
            }
		<%End IF%>
    }
			

    function initForm() {
        var f=window.document.frmRIR;
        f.txtCreateDate.value=getLocalDate();
        <% if bHSE then %>
           // var ActType=0;
            //var radios = document.getElementsByName('optSLBInvment');
            //for (var i = 0, length = radios.length; i < length; i++) {
              //  if (radios[i].checked) {
                    // do whatever you want with the checked radio
                    //alert("I am In Initform" + radios[i].value);
                //    ActType = radios[i].value
                    // only one radio can be logically checked, don't check the rest
                  //  break;
                //}
            //}
            //alert(ActType);
            //if ((ActType == '1') || (ActType == '2')) {
              //  SetProc();
        //}
			<%If ((DiffEvtClass >= 1) and (Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)))) then%>
				if (!(document.frmRIR.rdEventChoiceMan.checked)) { SetEventCategorisation(); }
			<%Elseif ((DiffEvtClass >= 1) and bHSE and bSQ and not bNR) then%>
				if (!(document.frmRIR.rdEventChoiceMan.checked)) { SetEventCategorisation(); }
			<%Elseif bHSE and not bSQ then%>	
				var ActivityType
				var radios = document.getElementsByName('rdEventSubCat');
				for (var i = 0, length = radios.length; i < length; i++) {
					if (radios[i].checked) {
						ActivityType = radios[i].value;  // alert(radios[i].value);
						break; // only one radio can be logically checked, don't check the rest
					}
				}
				if(document.getElementById("rdEventSafetyProc").checked && ActivityType > 0)
				{
					document.getElementById("rdEventSubCat"+ActivityType).checked=false;
				}
			<% End if%>    
        <% End if%>
        
		<%If bSQ and Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
			if(f.optHSE.checked)
			{
				//document.frmRIR.rdPLSSInv[0].checked=true;
				//document.frmRIR.rdPLSSInv[1].disabled=true;
			}
			else
			{
				document.frmRIR.rdPLSSInv[1].checked=true;
				document.frmRIR.rdPLSSInv[1].disabled=false;
			}
		<%End IF%>
        
            <%If issps >0 Then %>
           // if (document.getElementById("SpanProcess") != null) { document.getElementById("SpanProcess").style.display= "none"; }
        <%end If %>


        <%If (((bSQ and ShowDM() and (not isDMSQMapping(SQMappingID) or SafeNum(DMRecs) >0)) Or (bSQ And not ShowDM())))Then %>  //isDM(lPL)
            <%'if SQFcatID <>0  and  issaxon(lPL)  then
				if SQFcatID <>0  and  issaxon(iSubBSID)  then%>
                if(f.txtSubSPCategory){spcategory.setSL(f.txtSPCategory,f.txtSubSPCategory);}
        SQCat_SelAssign(<%=SQFcatID%>, <%=SQFSubcatID%>, <%=SQDCatID%>, <%=SQDSubcatID%>);
        <%Else%>	
            if(f.txtSubSPCategory){spcategory.setSL(f.txtSPCategory,f.txtSubSPCategory);}
        if(f.txtSubFailure){failure.setSL(f.txtFailure,f.txtSubFailure)}
        if(f.txtSubDamage){damage.setSL(f.txtDamage,f.txtSubDamage)}
        <%End if%>  
    <%End If%>
    OnChange_SiteType(f.txtLocation)
    <%If bSQ and iClass=1 and (LockCountSQ<>0 or chkSQLockingMgmt()) Then %>
        DisG1=0
        DisG2=0
        if ((document.frmRIR.optSQInvment[0].checked) || (document.frmRIR.optSQInvment[1].checked)){
					
            f.LossCat_G2.checked=true
					
            DisG2=1
        }		    
        if ((document.frmRIR.optCAffect[0].checked) && (document.frmRIR.optSQInvment[2].checked==false)){
					
            f.LossCat_G1.checked=true
					
            DisG1=1		         
        }

        if (document.frmRIR.optClass[document.frmRIR.optClass.selectedIndex].value==1 && document.frmRIR.optSQInvment[0].checked && document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value >= 1)
        {
            document.getElementById("tblNPT").style.display = "";
            if (document.frmRIR.isTimeLossEntered.value == "0")
            {
                if (document.frmRIR.rdNPT[0].checked)
                {
                    document.getElementById("txtNPT_M").style.visibility="hidden";
                    document.getElementById("txtNPT_LossCat_G1_M").style.visibility="hidden";
                    document.getElementById("txtNPT_LossCat_G2_M").style.visibility="hidden";
                }
                else
                {
                    document.getElementById("txtNPT_M").style.visibility="visible";
                    document.getElementById("txtNPT_LossCat_G1_M").style.visibility="visible";
                    document.getElementById("txtNPT_LossCat_G2_M").style.visibility="visible";
                }
            }
        }
        else 
            document.getElementById("tblNPT").style.display = "none";


				
        <%end If%>
				
        <%if not EnforceFlag then %>
        if(f.txtBSegment.options.length==2){f.txtBSegment.options[1].selected=true;}
        <%End if%>
				
        <%if bNR Then %>
             checkifDirectDMSQ();
        <%End IF %>
        
        
        }


    function cmdDelete_onclick() {
        var bConfirm = window.confirm('Are you sure you wish to DELETE this record');
        return (bConfirm) 
    }


    function txtReporter_onchange() {
        if (document.frmRIR.txtReporter.options[document.frmRIR.txtReporter.selectedIndex].value == ''){
            open("../Utils/searchLDAP.asp","searchemployees","height=400,width=400,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1")
        }
    }


    function ldapSearch(inp) {
        var query;
        if (typeof inp != "undefined") {
            query = inp.options[inp.selectedIndex].value;
            open("../Utils/searchLDAP.asp?lookup=1&txtName=" + escape(query),"ldap","height=400,width=400,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1")
        }
    }
		
    <%=Client_JSFun("frmRIR",lOrgNo,Client,CRMClient,"../Utils/")%>
    <%=CRMRig_JSFun("frmRIR",lOrgNo,"txtEvDate",CRMRigID,"../Utils/")%>

    function addOptionContractor(sText,sValue){
        sValue=sValue.replace(/'/i, "''");
        document.frmRIR.txtContractor.options[document.frmRIR.txtContractor.options.length] = new Option(sText, sValue);
        document.frmRIR.txtContractor.selectedIndex = document.frmRIR.txtContractor.options.length-1;
    }
    function c_ContractorName_onchange() {
        var Contractor='<%=Contractor%>'
        var orgno='<%=lorgno%>'
        var UrlStr="../Utils/searchContractor.asp?flag=1&Contractor="+Contractor+"&orgno="+orgno
        if (document.frmRIR.txtContractor.options[document.frmRIR.txtContractor.selectedIndex].text == "(SEARCH CONTRACTOR)"){
            open(UrlStr,"SearchContractor","height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=1,alwaysRaised=1")
        }
    }
			
    function AccountingUnitMaximo_onchange(){
        var UrlStr="../Utils/SearchMaximoAccUnits.asp?skip_inactive=Y"//?txtSearch
        if (document.frmRIR.txtAccountUnit.options[document.frmRIR.txtAccountUnit.selectedIndex].text == "(SEARCH MORE ACCOUNTING UNITS)"){
            open(UrlStr,"SearchAccountingUnit","height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=1,alwaysRaised=1")
        }
    }
	
	 function AccountingUnit_onchange(){
        var UrlStr="../Utils/SearchAccUnits.asp?skip_inactive=Y"//?txtSearch
        if (document.frmRIR.txtAccountUnit.options[document.frmRIR.txtAccountUnit.selectedIndex].text == "(SEARCH MORE ACCOUNTING UNITS)"){
            open(UrlStr,"SearchAccountingUnit","height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=1,alwaysRaised=1")
        }
    }
	
			
    function addOptionAccUnit(sText,sValue){
        sValue=sValue.replace(/'/i, "''");
        document.frmRIR.txtAccountUnit.options[document.frmRIR.txtAccountUnit.options.length] = new Option(sText, sValue);
        document.frmRIR.txtAccountUnit.selectedIndex = document.frmRIR.txtAccountUnit.options.length-1;
    }

    function addOption(sText,sValue){
        sValue=sValue.replace(/'/i, "''");
        document.frmRIR.txtReporter.options[document.frmRIR.txtReporter.options.length] = new Option(sText, sValue);
        document.frmRIR.txtReporter.selectedIndex = document.frmRIR.txtReporter.options.length-1;
    }

			
			
    function Env_onclick(chk){
        var Envmsg
        Envmsg="WARNING:\n\tThe Environmental loss category should only be used to report "
        Envmsg=Envmsg + "\n\tevents that are related to the natural environment.\nEXAMPLE:\n\tPhysical damage to "
        Envmsg=Envmsg + "vehicles, assets or people should be reported \n\tusing the appropriate health "
        Envmsg=Envmsg + "or safety loss category."
        if(chk.checked){
            if(!confirm(Envmsg)) chk.checked=false
        }
    }
			
    function Haz_onchange(sel){
        var msg,i,frmdoc
        frmdoc = document.frmRIR;
        msg="WARNING:\n\tThe Air Transport hazard category should only be used for hazards "
        msg=msg + "\n\t associated with Aircraft/Helicopter operation, procedures and systems. "
        i=sel.selectedIndex;
         <%IF ((DiffEvtClass >= 1) and  (bHSE))  then%>
			 <%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then %>
			 if (!(frmdoc.optSLBInvment[2].checked || frmdoc.optSLBInvment[3].checked)) 
			 {
				SetEventCategorisation()
				SetEventSubSafety()
			 }
			 <%Else%>
				SetEventCategorisation()
			 <%End IF%>
        <%end if%>
if(sel.options[i].text=='Air Transport'){
            if(!confirm(msg)) sel.selectedIndex=0;
        }
			  
    }
		
    function SetEventCategorisation()
    {
        var site, hazval,frmdoc,ActivityType
        frmdoc = document.frmRIR;
        site=frmdoc.txtLocation.options[frmdoc.txtLocation.selectedIndex].value;
        if (typeof document.getElementById("txtHazard") != 'undefined' && document.getElementById("txtHazard") != null)  
        {
            hazval  = document.getElementById("txtHazard").value;
			<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
            if (typeof frmdoc.rdWIBEventHSE[0] != 'undefined') 
            {
			<%End IF%>
                var radios = document.getElementsByName('optSLBInvment');
                for (var i = 0, length = radios.length; i < length; i++) {
                    if (radios[i].checked) {
                        // do whatever you want with the checked radio
                        ActivityType = radios[i].value;  // alert(radios[i].value);
                        break; // only one radio can be logically checked, don't check the rest
                    }
                }
				
				if ((ActivityType== '3') || (ActivityType == '4')) {
                        SetPers();
						setSysMan();						
                }
                else if ((ActivityType == '1') || (ActivityType == '2')) 
				{
                    if((hazval=='7') || (hazval=='8') || (hazval=='9') <%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> || (hazval=='12') || (hazval=='4') <%End IF%>){   //7=Explosives, 8=Radiation, 9=Pressure
                        SetProc();
                    }
                    else {
                        if (<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>((site.indexOf("RIG") > 0) || (site == '3:NON') || (site == '15:NON')) && <%End IF%>(hazval=='5')) {  //5=Fire/Flammables, AnyRIG, 3=FieldLocation,  15=Vessel
                            SetProc();
                        }
						else{ 	SetPers();}
                    }
				/*	else if (   ((hazval=='12') &&  (frmdoc.LossCat_C1.checked == true)) && 
								(document.frmRIR.rdWIBEventHSE[0].checked) && 
								((ActivityType == '1') || (ActivityType == '2')) )
					{   //12=Toxic, AccidentDischanged=CHECKED, HSEWellBarrier=CHECKED
						SetProc();
					}
					else if ((hazval=='19') || (hazval=='21')){   //, 19=Information and 21=Lifting, Mechanical
						SetProc();
					}
					else{ 	SetPers(); }*/					
				}
				else {
					if ((ActivityType != '1') || (ActivityType != '2')) {
						//SetPers();
						setSysMan();
					}
					SetPers();
                }

                /** Override all the selection based on this ANO166702 **/
				if ((hazval=='19') || (hazval=='21'))
				{
				    SetPers();
				}
				<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
                if ((document.frmRIR.rdWIBEventHSE[0].checked) || ((hazval=='12') && (frmdoc.LossCat_C1.checked == true) ) )
                {   //19=Information ---- 21=Lifting, Mechanical ---- 12=Toxics ---- rdWIBEventHSE[0].checked Well Integrity selected
				    SetProc();
				}
				
            } // if rdWIBEventHSE Closed
            <%End IF%>
        }  // if txtHazard closed
    }  //fnc end
				
	function SetProc()
    {
	    if (typeof document.frmRIR.rdEventSafetyProc != 'undefined')
	    {
	        document.getElementById("rdEventSafetyProc").checked = true;
	        globalProc = true;
	        globalPer = false;
	        document.getElementById("rdEventChoiceSys").checked = true;
	        document.getElementById("rdEventChoiceSys").disabled = false;
	        document.getElementById("rdEventChoiceMan").disabled = true;
	    }
	} 

	function SetEventSubSafety()
    {
	    if (document.getElementById("rdEventSafetyPers").checked)
	    {	
			document.getElementById("rdEventSubCat1").checked = false;
			document.getElementById("rdEventSubCat2").checked = false;
			document.getElementById("rdEventSubCat3").checked = false;
			document.getElementById("rdEventSubCat4").checked = false;
			document.getElementById("rdEventSubCat5").checked = false;
			document.getElementById("rdEventSubCat6").checked = false;
			document.getElementById("rdEventSubCat7").checked = false;
			document.getElementById("rdEventSubCat8").checked = false;
			document.getElementById("rdEventSubCat9").checked = false;	
			
	        document.getElementById("rdEventSubCat1").disabled = true;
			document.getElementById("rdEventSubCat2").disabled = true;
			document.getElementById("rdEventSubCat3").disabled = true;
			document.getElementById("rdEventSubCat4").disabled = true;
			document.getElementById("rdEventSubCat5").disabled = true;
			document.getElementById("rdEventSubCat6").disabled = true;
			document.getElementById("rdEventSubCat7").disabled = true;
			document.getElementById("rdEventSubCat8").disabled = true;
			document.getElementById("rdEventSubCat9").disabled = true;	        
	    }
		else
		{						
	        document.getElementById("rdEventSubCat1").disabled = false;
			document.getElementById("rdEventSubCat2").disabled = false;
			document.getElementById("rdEventSubCat3").disabled = false;
			document.getElementById("rdEventSubCat4").disabled = false;
			document.getElementById("rdEventSubCat5").disabled = false;
			document.getElementById("rdEventSubCat6").disabled = false;
			document.getElementById("rdEventSubCat7").disabled = false;
			document.getElementById("rdEventSubCat8").disabled = false;
			document.getElementById("rdEventSubCat9").disabled = false;
		}
	} 
        
	function SetPers(){
	    if (typeof document.frmRIR.rdEventSafetyPers != 'undefined')
	    {
			
	        if (document.getElementById("rdEventSafetyProc").checked && <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>((document.getElementById("txtHazard").value)!='11')<%Else%>((document.getElementById("txtHazard").value)=='4' || (document.getElementById("txtHazard").value)=='5' || (document.getElementById("txtHazard").value)=='7' || (document.getElementById("txtHazard").value)=='8' || (document.getElementById("txtHazard").value)=='9' || (document.getElementById("txtHazard").value)=='12')<%End IF%> )   // Code to avoid default selection of Personal safety radio buttion without checking Personal safety button value.
			{
	        document.getElementById("rdEventSafetyPers").checked = false;
			}
			else
			{
			document.getElementById("rdEventSafetyPers").checked = true;
			<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
				SetEventSubSafety();
			<%End If%>
			}
	        globalProc = false;
	        globalPer = true;
	        document.getElementById("rdEventChoiceSys").checked = true;
	        document.getElementById("rdEventChoiceSys").disabled = false;
	        document.getElementById("rdEventChoiceMan").disabled = true;
	    }
	}
    
    function setSysMan()
	{
        //alert("Hello");
        if (globalProc==true && (document.getElementById("rdEventSafetyProc").checked == true) )
        {
            document.getElementById("rdEventChoiceSys").checked = true;
            document.getElementById("rdEventChoiceSys").disabled = false;
            document.getElementById("rdEventChoiceMan").disabled = true;
            document.getElementById("rdEventChoiceMan").checked = false;
            //rdEventSafety.value=1
            //rdEventChoice.value=1
        }
        else if (globalProc==false && (document.getElementById("rdEventSafetyProc").checked == true) )
        {
            document.getElementById("rdEventChoiceSys").checked = false;
            document.getElementById("rdEventChoiceSys").disabled = true;
            document.getElementById("rdEventChoiceMan").disabled = false;
            document.getElementById("rdEventChoiceMan").checked = true;
        }
        else if (globalPer==true && (document.getElementById("rdEventSafetyPers").checked == true) )
        {
            document.getElementById("rdEventChoiceSys").checked = true;
            document.getElementById("rdEventChoiceSys").disabled = false;
            document.getElementById("rdEventChoiceMan").disabled = true;
            document.getElementById("rdEventChoiceMan").checked = false;
        }
        else if (globalPer==false && (document.getElementById("rdEventSafetyPers").checked == true) )
        {
            document.getElementById("rdEventChoiceSys").checked = false;
            document.getElementById("rdEventChoiceSys").disabled = true;
            document.getElementById("rdEventChoiceMan").disabled = false;
            document.getElementById("rdEventChoiceMan").checked = true;
        }
        //00, 01, 10, 11
    }
    function CheckText()
    {   
        var tx;
        tx=document.frmRIR.txtJobID.value;
		var ty = /^[a-zA-Z0-9.\-]+$/;   
        return tx.match(ty);   
    } 
	function CheckTextWithoutDot()
    {   
        var tx;
        tx=document.frmRIR.txtJobID.value;
        var ty = /^[a-zA-Z0-9\-]+$/;   
        return tx.match(ty);   
    }   
	
	function CheckNumberText(tx)
    {   
		 var ty = /^[0-9]*$/; 
         return tx.match(ty);   
    }
			
    function checkIlluminaID()
    {
        if(document.frmRIR.HdnIlluminaJobAID.value == document.frmRIR.txtJobID.value)
        {	
            return false;
        }
        else 
        {
            var x=confirm("illumina data saved in QUEST will be deleted. You will need to reimport from the ALS Details page. Do you wish to continue?");
            if (x==true)
            {
                document.frmRIR.ALSDelete.value = 1;			        
                document.frmRIR.submit();
            }
            else
            {	
                document.frmRIR.txtJobID.value = document.frmRIR.HdnIlluminaJobAID.value
                return false;		
            }
        }
			
    }
			
    function convertstringtodate (dateString) {
        //var dateParts = dateString.split(/-/);
        var t_month = dateString.split(",")[0].split(" ")[0];
        var t_day = dateString.split(",")[0].split(" ")[1];
        var t_year = dateString.split(",")[1];
        //return new Date((dateParts[2] * 1), ($.inArray(dateParts[1].toUpperCase(), ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]) * 1), dateParts[0] * 1);
        return new Date((t_year * 1), ($.inArray(t_month.toUpperCase(), ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]) * 1), t_day * 1);
    }

		
    function datediff (firstDay, lastDay) {
        var diffDays = parseInt((lastDay - firstDay) / (1000 * 60 * 60 * 24));
        return diffDays;
    }


    function verifydata() {
			
        var d;
        var evndate =convertstringtodate(frmRIR.txtEvDate.value);
        <%If bNR  Then%>
        d = new Date();
        var datefifferencepast=datediff(evndate,d);
        <%else%>
			
        <%if (CDate(dtRptDatetmp) >= CDate(comparedateEventdate)) Then%>
			
        d = '<%=dtRptDatetmp%>';
        var d1= Date.parse(d);
        var datefifferencepast=datediff(evndate,d1);
        <%else%>
        datefifferencepast=0;
        <%end if%>
        <%end if%>

            var errorsheader = '';
		var errorsmsg = '';
        errorsmsg = errorsmsg + validateLossSafetynet();
        var frmdoc = document.frmRIR;
        var msg,tmp,TotLoss;	
        TotLoss = 0;
        errorsheader += '_______________________________________________\n\n';
        errorsheader += 'The form was not saved because of the following error(s).\n';
        errorsheader += '_______________________________________________\n\n';
		

        if (frmdoc.optSQ.checked)
        {
			<%If isALSSeg(iBSID) then%>
				if(!(CheckNumberText(document.frmRIR.txtJobID.value)))
			{
				errorsmsg += 'Please enter valid Numeric value in Job / Service Order Id. \n';
			}	
			
            else if(document.frmRIR.HdnIlluminaJobAID.value!='')
            {			
                checkIlluminaID();
            }
			<%End If%>
        }

		//WL Ro
		if (document.frmRIR.optSQ.checked) {		
        <%if (isREWSQMapping(SQMappingID)) then %>
		<%If isROP > 0 and bSQ Then%>		
		if (frmdoc.chkClosed) {	
			if (document.frmRIR.chkClosed.checked) {		
				if (document.frmRIR.optROPInv.value != <%= wlro %>)
				{			
				errorsmsg += 'Remote Operations Involved value should match in SQ main and WL SQ Details tab \n';			
				}
			}
		}
		<% End If %>
		<% End If %>
		}		
			
        /*if (frmdoc.lstDepartment.value !=0)
        {
            <%if not EnforceFlag then %>
            if (frmdoc.txtBSegment.options[frmdoc.txtBSegment.selectedIndex].value != frmdoc.hidBSID.value)
                {
                    var a=confirm("Warning! Department is not related to the Sub-Segment.\n Save anyhow?");
                    if (a==true)
                    {
                        frmdoc.submit();
                    }
                    else
                    {
                        return a;		
                    }
                }
                <%end if%>
        }*/	
				

      //  if(document.frmRIR.PLID.value in { 1:1,120:1}) 
      //  {
	  //PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.
	  //Checks only for D&M Sub Segments
		<%If isDMSeg(iBSID) then%> 
            if (frmdoc.optSQ.checked)
            {
                if (frmdoc.txtJobID.value !='')
                {
                    var strJob;
                    strJob=frmdoc.txtJobID.value;	
                    if (!(CheckText()))
                    {
						if (((strJob.charAt(0)).toUpperCase() == "O" || (strJob.charAt(0)).toUpperCase() == "A") && strJob.indexOf('.')>0){
						errorsmsg += 'No special characters allowed except for "-" and "." in Job / Service Order Id. \n';
						}
						else{
						errorsmsg += 'No special characters allowed except for "-" in Job / Service Order Id. \n';
						}
					
					}
						
                    if (!(frmdoc.txtJobID.readOnly) && (CheckText()))
                    {
						if (((strJob.charAt(0)).toUpperCase() == "O" || (strJob.charAt(0)).toUpperCase() == "A") && strJob.indexOf('.')>0)
						{	
							if(strJob.split('.').length != 3 || strJob.charAt(1) != ".")
							{
								errorsmsg += 'Invalid Job  / Service Order Id. \n';
							}
							else
							{
								var str1 = strJob.indexOf('.')+1;
								var str2 = strJob.lastIndexOf('.');
								var str3 = strJob.lastIndexOf('.')+1;
								if(strJob.indexOf('-')>0){
									var str4 = strJob.lastIndexOf('-');		
								}else{
									var str4 =strJob.length;
								}
								
								if(isNaN(strJob.substring(str1,str2)) || isNaN(strJob.substring(str3, str4)))
								{
									errorsmsg += 'Only numeric value allowed for Job  / Service Order Id. \n';
								}	
								if((strJob.substring(str1,str2).length < 6) || (strJob.substring(str3, str4).length < 2))
								{
									errorsmsg += 'Invalid Job  / Service Order Id. \n';
								}
							
							}
							if ((strJob.indexOf('-')>0))
							{
								if(strJob.substring(strJob.indexOf('-')+1).length == "0")//there should only be 1 digit after -
								{
									errorsmsg += 'Invalid Run Number. \n';
								}
								if ((strJob.indexOf('-') == strJob.lastIndexOf('-')))
								{
									if (isNaN(strJob.substring(strJob.indexOf('-')+1)))
									{
										errorsmsg += 'Run Number should be numeric only. \n';
									}
								}
							}
							
						}
						else
						{	
						if (!(CheckTextWithoutDot()))
						{
							errorsmsg += 'No special characters allowed except for "-" in Job / Service Order Id. \n';
						}
						
                        
                        if ((strJob.indexOf('-')<0) || (strJob.indexOf('-') != strJob.lastIndexOf('-')))
                        {
                            errorsmsg += 'Invalid Job  / Service Order Id. \n';
                        }
                        else
                        {
                            if(strJob.substring(strJob.indexOf('-')+1).length == "0")
                            {
                                errorsmsg += 'Invalid Run Number. \n';
                            }
										
                            if ((strJob.indexOf('-') == strJob.lastIndexOf('-')))
                            {
                                if (isNaN(strJob.substring(strJob.indexOf('-')+1)))
                                {
                                    errorsmsg += 'Run Number should be numeric only. \n';
                                }
                            }
                        }
					
						}
                    }  // end readonly here
							
						
                }
            }
		<%End If%>
       // }	
					
        if (!(frmdoc.optSQ.checked || frmdoc.optHSE.checked))
        {
            errorsmsg += 'You must select HSE or Service Quality. \n';
        }
        <%if bSQ and Session("UType")<>"O" then %>
        if (datefifferencepast >21)
        {	
            errorsmsg += 'Event Date cannot be more than 21 days in the past . \n';
        }
        <%end if%>
				
				
        <%if bHSE and Session("UType")<>"O"  then %>
        if (datefifferencepast >21)
        {
            errorsmsg += 'Event Date cannot be more than 21 days in the past . \n';
        }
        <%end if%>
	
        <%if not EnforceFlag then %>
        if (frmdoc.txtBSegment.options[frmdoc.txtBSegment.selectedIndex].value == '')
        {
            errorsmsg += 'Sub-Segment is a required field. \n';
        }
        <%end if%>
        if (frmdoc.txtReporter.options[frmdoc.txtReporter.selectedIndex].value == '')
        {
            errorsmsg += 'Reporter  is a required field. \n';
        }
				
        //if (!(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proISM").checked==true || document.getElementById("proSPM").checked==true || document.getElementById("proNO").checked==true))
		if (!(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proSPM").checked==true || document.getElementById("proNO").checked==true))
        {
            //errorsmsg += ' A selection for '+ '<%=IPMText(1)%>'+' Related? option has not yet been made? \nPlease '
            //errorsmsg += ' click the Hyperlink  for '+ '<%=IPMText(1)%>' +' Related? before making your choice.\n'
			errorsmsg += ' A selection for Integrated Performance Management (IPM) Related? option has not yet been made? \nPlease '
            errorsmsg += ' click the Hyperlink  for Integrated Performance Management (IPM) Related? before making your choice.\n'
        }
        <% ' This code is commented as of now and needs to be uncomment in future whenever we will get such intimation from business
        'If isPTEC > 0 then 
        %>
        /*	if (!(frmdoc.optPTECInv[0].checked || frmdoc.optPTECInv[1].checked))
            {
                errorsmsg += ' A selection for "'+ '<%=PTEC(1)%>'+'?" option has not yet been made?  \n'
                //errorsmsg += '\nPlease click the Hyperlink  for '+ '<%=PTEC(1)%>' +'? before making your choice.\n'
            } */
        <%
        'End If
        %>
        /* (#2588673) */
        //if(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proISM").checked==true || document.getElementById("proSPM").checked==true){
		if(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proSPM").checked==true){
            if (frmdoc.optContractorInv) //***** (MS HIDDEN) - Commented If loop  ***** 
            {
                if (!(frmdoc.optContractorInv[0].checked || frmdoc.optContractorInv[1].checked))
                {
                    errorsmsg += ' A selection for "Contractor Involved?" option has not yet been made?\n'
                    errorsmsg += ' Selection is required to be able to SAVE the report.\n'
                }
            }
        }
        /*(#2588673) */

        if ((frmdoc.txtEvDate.value == 'mmm dd, yyyy') || (frmdoc.txtEvDate.value == ''))
        {
            errorsmsg += 'Invalid event date. \n';
        }
        else if (chkdate(frmdoc.txtEvDate) == false)
        {
            errorsmsg += 'Invalid event date. \n';
        }

        if ((frmdoc.txtEvTime.value == '24Hr') || (frmdoc.txtEvTime.value.length > 5))
        {
            errorsmsg += 'Invalid event  time. \n';
        }
        else if (IsValidTime(frmdoc.txtEvTime.value) == false)
        {
            errorsmsg += 'Invalid event time - Enter hh:mm format. \n';
        }
        tmp=frmdoc.txtLocation.options[frmdoc.txtLocation.selectedIndex].value;		
        if (tmp == '0')
        {
            errorsmsg += 'Site is a required field. \n';
        }else{
					 
            if(tmp.indexOf("RIG") > 0){
                if(frmdoc.txtCRMRigID.options[frmdoc.txtCRMRigID.selectedIndex].value==''){
                    errorsmsg += 'CRM RIG Name is a required field. \n';
                }
                if((frmdoc.txtCRMRigID.options[frmdoc.txtCRMRigID.selectedIndex].value=='NO-CRM-RIG')&&(frmdoc.txtLoc.value == '')){
                    errorsmsg += 'Site Name is a required When Rig Name is Not Available in CRM. \n';
                }		
            }/*else{
                if((frmdoc.txtCRMRigID.options[frmdoc.txtCRMRigID.selectedIndex].value=='')&&(frmdoc.txtLoc.value == '')){
                    errorsmsg += 'Select Either CRM Rig Name or Enter Site Name. \n';
                }else{

                    if (frmdoc.txtLoc.value == ''){errorsmsg += 'Site Name is a required field. \n';}
                }
            }*/

        }

				
        <%if EnableOperation>0 Then %>
        if (frmdoc.cat_sq.value == ' ')
        {
            errorsmsg += 'Operation Category is a required field. \n';
        }

        if ( frmdoc.subcat_sq.value == ' ')
        {
            errorsmsg += 'Operation Subcategory is a required field.\n';
        }
        <%End if%>
				
        if (frmdoc.txtShortDesc.value == '')
        {
            errorsmsg += 'Brief description is a required field. \n';
        }

        if (frmdoc.txtFullDesc.value == '')
        {
            errorsmsg += 'Detailed description is a required field. \n';
        }
        else if (frmdoc.txtFullDesc.value.length < 50) 
        {
            errorsmsg += 'Detailed description should be atleast 50 characters. \n';
        }

        if (frmdoc.optClass[0].checked)
        {
            if (frmdoc.cmbSeverity.options[frmdoc.cmbSeverity.selectedIndex].value == '0')
            {
                errorsmsg += 'Accident classification  is a required field. \n';
            }				
        }
			
						
        <%If bHSE Then%>
				
            if (frmdoc.optHSE.checked)
        {
            if (frmdoc.optClass.options[frmdoc.optClass.selectedIndex].value == '1')
            {
                <%if bHSE and (LockCount<>0 or chkHSELockingMgmt()) then%>
                    if (frmdoc.cmbHSESeverity.selectedIndex == 4) {
                        errorsmsg += 'HSE Severity must be indicated. \n';
                    }
                <%end if%>
                }

            if (frmdoc.RegRec[0].checked)
            {
                if (frmdoc.optSLBInvment[3].checked == true) {
                    errorsmsg += 'You cannot select Regulatory Recordable without also selecting SLB Involved or Concerned.\n';
                }
            }
					
					

            if (frmdoc.LossCat_A2.checked) //Occupational Illness 
            {
                if (!(frmdoc.optSLBInvment[0].checked || frmdoc.optSLBInvment[1].checked)) {
                    errorsmsg += 'You cannot select Occupational Illness without also selecting SLB Involved.\n';
                }
            }

            if (frmdoc.LossCat_A3.checked) //Non-Occupational Illness 
            {
                if (!(frmdoc.optSLBInvment[2].checked || frmdoc.optSLBInvment[3].checked)) {
                    errorsmsg += 'You cannot select Non-occupational Illness when also selecting SLB Involved.\n';
                }
            }

            if (frmdoc.LossCat_A2.checked | frmdoc.LossCat_A3.checked) //Illnesses
            {
                if (frmdoc.LossCat_A1.checked) { //Injury
                    errorsmsg += 'You cannot select an Illness and also select Injury.\n';
                }
            }


            <%if bHSE and (LockCount<>0 or chkHSELockingMgmt()) then %>
            if (!(<% GenerateCatJS Application("a_LossCategories"),"hse"%>))
            {
                errorsmsg += 'At least one HSE category (Health, Safety, Environment) needs to be checked. \n';
        }

        if (frmdoc.txtHazard.options[frmdoc.txtHazard.selectedIndex].value == '0')
        {
            errorsmsg += 'Hazard Category is a required field. \n';
        }
        <%End If%>
        }
    <%End If%>
	
	<%IF bHSE and (Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv))) then%>
		var ActivityType
		var radios = document.getElementsByName('rdEventSubCat');
		for (var i = 0, length = radios.length; i < length; i++) {
			if (radios[i].checked) {
				ActivityType = radios[i].value;  // alert(radios[i].value);
				break; // only one radio can be logically checked, don't check the rest
			}
		}
		if(document.getElementById("rdEventSafetyProc").checked && isNaN(ActivityType))
		{
			errorsmsg += 'Please select one sub category from "Event Categorization". \n';
		}
		if (document.getElementById("rdEventSafetyPers").checked)
		{
			SetEventSubSafety();
		}
	<%End IF%>
	
    <%If bSQ  Then%>

        if (frmdoc.optSQ.value == 'on')
    {
        if (frmdoc.optClass.options[frmdoc.optClass.selectedIndex].value == '1')
        {   
            <%if bSQ and (LockCountSQ<>0 or chkSQLockingMgmt()) then %>
            if (frmdoc.cmbSQSeverity.selectedIndex == 4) {
                errorsmsg += 'Service Quality Severity must be indicated. \n';
            }
            <%end if%>
						
            if((frmdoc.optSQInvment[0].checked) || (frmdoc.optCAffect[0].checked)){
                if (frmdoc.txtCRMClient.options[frmdoc.txtCRMClient.selectedIndex].value == ''){
                    errorsmsg += 'CRM Client is a required field when External or Client Affected is checked. \n'
                }
            }
						
            //npt 
            <%If blnNPT_Exempt = 0 Then%> //Shailesh 30-Oct-2009 Swift# 2438856 
            {
                if(((4-intRseverity) < frmdoc.cmbSQSeverity.selectedIndex) && (intRseverity != 0) && (frmdoc.cmbSQSeverity.selectedIndex != 4))
                {
                    errorsmsg += 'Selected SQ severity is not matching with Time loss data, Please adjust the data in Time Loss tab before resetting the severity  \n'
        }
    }
    <%End If%> //Shailesh 30-Oct-2009 Swift# 2438856
    }

<%if not (iClass = 1 and LockCountSQ=0 and bSQ and not chkSQLockingMgmt())  then %>

if (!(<% GenerateCatJS Application("a_LossCategories"),"sq"%>))
    {
	
    //errorsmsg += 'At least one Service Quality category (Quality) needs to be checked. \n';
    errorsmsg += 'At least one Quality Loss option needs to be checked.\n';
 
    }
					<%End If%> 
					
					if (frmdoc.CDesc) {
					    if (frmdoc.CDesc.value == '')
					    {
					        errorsmsg += 'Suspected Immediate Causes is a required field. \n';
					    }
					}
    if (frmdoc.optSQStandard) {
        if (!(frmdoc.optSQStandard[0].checked || frmdoc.optSQStandard[1].checked) ){
            errorsmsg += 'Compliance with SQ Standards selection is required. \n'; 
        }
    }
    <% If (((IsSPRequired and ShowDM() and (not isDMSQMapping(SQMappingID)  or SafeNum(DMRecs) >0)) Or (IsSPRequired And not ShowDM()))) Then %>   //isDM(lPL)   
					  
        if(frmdoc.txtSPCategory) {
            if(frmdoc.txtSPCategory.options[frmdoc.txtSPCategory.selectedIndex].value == '0') {
                errorsmsg += 'Service/Product Category is a required field. \n';
            } else {
                if(frmdoc.txtSubSPCategory.options[frmdoc.txtSubSPCategory.selectedIndex].value == '0') {
                    errorsmsg += 'Service/Product Sub-Category is a required field. \n';
                }				
            }
        }
    <% End If %>
    if(frmdoc.txtFailure){
        if(frmdoc.txtFailure.options[frmdoc.txtFailure.selectedIndex].value == '0') {
            errorsmsg += 'Failure Category is a required field. \n';
        } else {
            if(frmdoc.txtSubFailure.options[frmdoc.txtSubFailure.selectedIndex].value == '0') {
                errorsmsg += 'Failure Sub-Category is a required field. \n';
            }				
        }
    }
    <%if isDamageRequired Then%>
    if(frmdoc.txtDamage <%if  issaxon(iSubBSID) Then response.write "&& document.getElementById('Saxon_Damage_cat').style.display==''"%>){
        if(frmdoc.txtDamage.options[frmdoc.txtDamage.selectedIndex].value == '0') {
            errorsmsg += 'Damage Category is a required field. \n';
        } else {
							if(frmdoc.txtSubDamage.options[frmdoc.txtSubDamage.selectedIndex].value == '0') {
							    errorsmsg += 'Damage Sub-Category is a required field. \n';
							}
    }
    }
					<%End IF%>
					<% if isREWSQMapping(SQMappingID) or isSPWL(iBSID) Then %>   //isRew(lPL)
						if(frmdoc.txtAccountUnit){
						    if(frmdoc.txtAccountUnit.options[frmdoc.txtAccountUnit.selectedIndex].value == '') {
						        errorsmsg += 'Accounting Unit is a required field. \n';
						    }
						}
    <%End IF%>
					
    //swi changes validate
    <%If ((iClass = 2) or (iClass=1)) and (isSWI>0) then%>
						
        if (!(frmdoc.swiqn[0].checked))
    {
        if (!(frmdoc.swiqn[1].checked))
        {
            errorsmsg += 'Response to question is required: "Was the activity of this event covered by an official \'Do It Right\' SWI/Checklist/Emergency Checklist?" \n';
        }
    }
						 
    <%if (iClass=1) then%>  
    var sqsev = document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value;
    if (((sqsev ==1) || (sqsev ==2)) && (frmdoc.swiqn[0].checked))
    {
        if (!(frmdoc.swiqntwob[0].checked))
        {
            if(!(frmdoc.swiqntwob[1].checked))
            {
                errorsmsg += 'Response to question is required: "Was the loss preventable by the use of the \'Do It Right\' SWI/Checklist/Emergency Checklist?" \n';
            }
        }
							
        if (!(frmdoc.swiqnthree[0].checked)) 
        {
            if (!(frmdoc.swiqnthree[1].checked)) 
            {
                errorsmsg += 'Response to question is required: "Was the \'Do It Right\' SWI/Checklist/Emergency actually followed during the activities?" \n';
            }
        }
    }
							
							
    <%End if%>
						
    <%if (iClass=2) then%> 

    if (frmdoc.swiqn[0].checked)
    {
        if (!(frmdoc.swiqntwo[0].checked))
        {
            if (!(frmdoc.swiqntwo[1].checked))
            {
                errorsmsg += 'Response to question is required: "Was the potential loss prevented by the use of the \'Do It Right\' SWI/Checklist/Emergency Checklist?" \n';
            }
        }
    }
    <%End if%>
					
					
					
<%End if%>
					
					
    // SPS changes
<%If (isSPS>0) then%>
tmp=frmdoc.SQB2_0.options[frmdoc.SQB2_0.selectedIndex].value;
   // if (tmp == '0')
   // {
   //     errorsmsg += 'Please select the Function. \n';
							
  //  }
				if (document.frmRIR.SQL2_0.value=="")
{				
		

}

else
{		//alert(document.frmRIR.SQL2_0.value);		
    tmp=frmdoc.SQL2_0.options[frmdoc.SQL2_0.selectedIndex].value;
		
    if (tmp == '0')
    {
        errorsmsg += 'Please select the Process. \n';
							
    }
    else
    {
    if (document.getElementById("Pro").style.display == '') {
        //if (tmp != '38')
        //{
        tmp=frmdoc.SQL3_0.options[frmdoc.SQL3_0.selectedIndex].value;
        if (tmp == '0')
        {
            errorsmsg += 'Please select the Metro Stop. \n';
        }
        else
        {
		
		if(frmdoc.SQL4_0.value != "")
		{
            tmp=frmdoc.SQL4_0.options[frmdoc.SQL4_0.selectedIndex].value;
            if (tmp == '0')
            {
                errorsmsg += 'Please select the Activity. \n';
            }
			}
        }
        //}
    }				
    }
					
		}				
						
    <%End if%>
					
        // slim validation
    <%'if (iClass=1) or (iClass=2)then
    %> 
        //if (!(frmdoc.sCtrl.checked || frmdoc.sTech.checked || frmdoc.sProc.checked || frmdoc.sComp.checked || frmdoc.sBehav.checked)) 
        //{
        //  errorsmsg += 'Please select atleast one SPS Improvement Domain. \n';
        //}
    <% 'End if
    %>
					
    if (document.getElementById("tblNPT"))
    {
        var flagNPTRMVal;
        flagNPTRMVal=true;
        if(document.getElementById("tblNPT").style.display=="" && document.frmRIR.isTimeLossEntered.value == "0")
        {
            <%if LockCountSQ<>0 or chkSQLockingMgmt() then%>
            var txtValid = frmdoc.txtNPT.value;
			var txtValid2 = frmdoc.rdNPT[0].checked;
			var txtValid3 = frmdoc.txtNPT_LossCat_G1.value;
			var txtValid4 = frmdoc.txtNPT_LossCat_G2.value;
            var valSQSev = frmdoc.cmbSQSeverity[frmdoc.cmbSQSeverity.selectedIndex].value;
			var txtSQSev = frmdoc.cmbSQSeverity[frmdoc.cmbSQSeverity.selectedIndex].text;
			if(valSQSev==1 && txtValid==0 && txtValid3==0 && txtValid4==0 && txtValid2==false)
			{
				errorsmsg += 'You are attempting to create a Loss report with an NPT value of zero.\nIf this is correct - the Loss is because of another causal factor.\nPlease select Severity Escalation Applied to Yes.\n';
                flagNPTRMVal = false;
			}
			if((txtValid==0 && txtValid3!=0) || (txtValid==0 && txtValid4!=0))
			{
				errorsmsg += ' Red Money values cannot be greater than zero - because NPT is zero.\n';
                flagNPTRMVal = false;
			}
            if (txtValid != '')
            {
                if((!IsNumericval(txtValid,1)))
                {
                    errorsmsg += 'Overall NPT cannot be non-numeric.\n';
                    flagNPTRMVal = false;
                }
            }
													
											
            txtValid = frmdoc.txtNPT_LossCat_G1.value;
            if (isNaN(txtValid)==false && txtValid != "") TotLoss = parseFloat(txtValid);
            if (txtValid != '')
            {
                if((!IsNumericval(txtValid,1)))
                {
                    errorsmsg += 'Estimated Client Red Money cannot be non-numeric.\n';
                    flagNPTRMVal = false;
                }
            }
							
						
            txtValid = frmdoc.txtNPT_LossCat_G2.value;
            if (isNaN(txtValid)==false && txtValid != "") TotLoss = parseFloat(TotLoss) + parseFloat(txtValid);
            if (txtValid != '')
            {
                if((!IsNumericval(txtValid,1)))
                {
                    errorsmsg += 'Estimated SLB Red Money cannot be non-numeric.\n';
                    flagNPTRMVal = false;
                }
            }
            <%end if%>
							
            <%if LockCountSQ<>0 or chkSQLockingMgmt() then%>
            if (frmdoc.rdNPT[1].checked && flagNPTRMVal)
            {
                var valMatrixNPT,valMatrixLoss,valMatrixSev,strTLMatrix,valNPT;
                strTLMatrix = frmdoc.hdTLMatrix.value;
                valMatrixSev = valSQSev + '';
                valMatrixLoss = parseFloat(strTLMatrix.substring(strTLMatrix.indexOf("#",strTLMatrix.indexOf("##" + valMatrixSev)+2)+1,strTLMatrix.indexOf("#",strTLMatrix.indexOf("#",strTLMatrix.indexOf("##" + valMatrixSev)+2)+1)));
                valMatrixNPT = parseFloat(strTLMatrix.substring(strTLMatrix.indexOf("#",strTLMatrix.indexOf("#",strTLMatrix.indexOf("##" + valMatrixSev)+2)+1)+1,strTLMatrix.indexOf("##",strTLMatrix.indexOf("##" + valMatrixSev)+2)));
                valNPT = parseFloat(frmdoc.txtNPT.value);
				  if (isNaN(TotLoss)==true || TotLoss=='') TotLoss = 0;
                if (isNaN(valNPT)==true || valNPT=='') valNPT = 0;
				if((isNaN(TotLoss)==true || TotLoss=='') && (isNaN(valNPT)==true || valNPT==''))
				{
				errorsmsg += "\nYou have selected the SQ severity '"+ txtSQSev + "'.\nThis implies the Client + Slb $Loss >= "+ valMatrixLoss + "K$ or/and the Overall NPT >= "+ valMatrixNPT +" Hrs.";
				}
				else
				{
                if(TotLoss < valMatrixLoss && valNPT < valMatrixNPT)
                {
                    errorsmsg += "\nYou have selected the SQ severity '"+ txtSQSev + "'.\nThis implies the Client + Slb $Loss >= "+ valMatrixLoss + "K$ or/and the Overall NPT >= "+ valMatrixNPT +" Hrs.";
                }
				}
            }
            <%end if%>
            }
    }				
    }

    // *****************************************************************
    // Code added for NPT <<2401608>>
    // *****************************************************************

    if  ((frmdoc.optClass.options[frmdoc.optClass.selectedIndex].value == '1'))
    {
				
        var checkSLBRelated='False';
        var checkRIRExternal='False';
        var CheckClientAffected='False';
        if(frmdoc.optSQInvment[0].checked)
        { 
            checkSLBRelated='True';
            checkRIRExternal='True';
        }     
        if(frmdoc.optSQInvment[2].checked){ 
            checkSLBRelated='True';
            checkRIRExternal='False';
        }    
        if(frmdoc.optSQInvment[1].checked)
        { 
            checkSLBRelated='False';
            checkRIRExternal='False';
        }  
        if (frmdoc.optCAffect[1].checked)
        {
            CheckClientAffected='True';
        }
				
        if (frmdoc.optCAffect[0].checked)
        {
            CheckClientAffected='False';
        }
				
    }
	
    //****************************end change of NPT*************************************
    <%End If%>
					
					
        if (errorsmsg != '')
    {
        alert(errorsheader += errorsmsg);
        return false;
    }
    else
    {
					if(isSubmit){
					    alert('This information already submited for processing. Please wait....');
					    return false;
					}
    <%If bHSE Then%>
    if (frmdoc.optClass.options[frmdoc.optClass.selectedIndex].value == '1'){
        if ((frmdoc.cmbHSESeverity.options[frmdoc.cmbHSESeverity.selectedIndex].value==4) && (frmdoc.LossCat_A1.checked || frmdoc.LossCat_A2.checked || frmdoc.LossCat_A3.checked)){		
            errorsmsg="Warning\n\nPlease check the Brief Description and Detailed Description of this report to confirm that there are no names of injured or " ;
            errorsmsg=errorsmsg+'deceased parties included in the text. If there are, please remove the names and use the terminology "Injured Person" or "Deceased" instead.\n\n' ;
            errorsmsg=errorsmsg+"Names can be entered in the Personnel Loss page but once saved will then become Confidential and will not be displayed.\n\n";
            errorsmsg=errorsmsg+"Press OK to Save or Cancel to modify.";
            if(confirm(errorsmsg)) errorsmsg='';
        }
					
        <%If ACLDefined Then%> 
        if ((frmdoc.cmbHSESeverity.options[frmdoc.cmbHSESeverity.selectedIndex].value==4)&&(frmdoc.optSLBInvment[3].checked == true))
        {
            msg='This SLB Non Involved/Informative HSE Catastrophic Event is currently protected.\n';
            msg=msg+'Do you wish to Un-Protect the document.\n'
            msg=msg+'Press OK for Un-Protect or Cancel to Continue.'
            if(confirm(msg)) frmdoc.ProtectDoc.value='Protect'
        }
        <%End If%>
        }	
    <%End If%>
    // *****************************************************************
   //  Code if user deselects all NPT losses then show message <<2401608>>
   // *****************************************************************

    <%If (bSQ  AND iClass = 1 AND chktimedata() ) Then%>
    if (errorsmsg == '' )
    {
        <%If LockCountSQ<>0 or chkSQLockingMgmt() Then%>
        if ((!(frmdoc.LossCat_G3.checked)) && (!(frmdoc.LossCat_G2.checked))&& (!(frmdoc.LossCat_G1.checked)))
        {
					   
            var answer = confirm("You have deselected all Non-Productive Time Loss Types and this will cause the Loss Time tab data to be removed from this RIR?")
            if (answer) 
            { 
            }
            else
            {  
                errorsmsg = 'You have deselected all Loss Types and this will cause the Loss Time tab data to be removed'         
            }      
													   
        }
        <%End If%>
        }
    <%End If%>
    //****************************end changes****************************************************
    if (errorsmsg == ''){
					
        checkifDirectDMSQ();
						
        isSubmit=true;
        frmdoc.submit();					
    }
    }
    }

    function fncChangeHazardCat() 
    {
				
        <%If bHSE and bSQ Then%>
					
            if(<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>document.frmRIR.rdFireExplosion[0].checked && document.frmRIR.rdWIBEventSQ[0].checked && <%End IF%> document.frmRIR.FireMode.value == 1)
        {
            document.frmRIR.txtHazard.value = '<%=HazID%>';
        }    
        <%End If%>    
			   
        }
			
			
			
    function optHSESQ_onchange(intCaller) {
        var bConfirm;
        var isHSESQ;
        if (document.frmRIR.optSQ.checked && document.frmRIR.optHSE.checked)
            isHSESQ = true;
        else 
            isHSESQ = false;
		
		<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>	
		if (intCaller >= 3 && intCaller <=5 && isHSESQ)
        {
            bConfirm = true;
        }
        else if (intCaller >= 3 && intCaller <=5 && !(isHSESQ))
        {
            if (document.frmRIR.optSQ.checked) {
                bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ. \nSelect Cancel to continue as SQ only.');
            }
            else{
                if (document.frmRIR.optHSE.checked) {
                    bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ. \nSelect Cancel to continue as HSE only.');
                }
            }
        }
		<%Else%>
        if ((((intCaller >= 3 && intCaller <=5) || (intCaller == 7)) && isHSESQ))
        {
			if (intCaller == 4 && document.getElementById("txtHazard").value == 0)
			{	
				alert("Please select Hazard Category before selecting any event from Event Categorization");
				bConfirm = false;				
			}
			else
			{
				bConfirm = true;
			}
        }
        else if (((intCaller >= 3 && intCaller <=5) || (intCaller == 7)) && !(isHSESQ))
        {
            if (document.frmRIR.optSQ.checked) {
                if (intCaller == 7)
				{
					bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ.');
				}
				else
				{
					bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ. \nSelect Cancel to continue as SQ only.');
				}
            }
            else{
                if (document.frmRIR.optHSE.checked) {
					if (intCaller == 4)
					{
						if (document.getElementById("txtHazard").value == 0)
						{
							alert("Please select Hazard Category before selecting any event from Event Categorisation");
							bConfirm = false;
						}
						else
						{
							bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ.');
						}
					}
					else
					{
						bConfirm = window.confirm('Select OK to convert this RIR to HSE/SQ. \nSelect Cancel to continue as HSE only.');
					}
                }
            }
        }
		<%End IF%>
        else
        {
		
	       var prosel6 = document.frmRIR.iClassn;
				var prosel7 = document.frmRIR.optClass;
				var prosel8 = document.frmRIR.isOPFval;
				var prosel9 = document.frmRIR.optSQ;
				var prosel10 = document.frmRIR.opfepcc;
				var prosel11 = document.frmRIR.opfonm;
			if ((typeof prosel6 != "undefined") && (typeof prosel7 != "undefined") && (typeof prosel8 != "undefined") &&  (typeof prosel9 != "undefined") && (typeof prosel10 != "undefined") &&  (typeof prosel11 != "undefined"))
			{
			if (((document.frmRIR.iClassn.value)!=(document.frmRIR.optClass.value)) && (document.frmRIR.isOPFval.value=="True")  && (document.frmRIR.optSQ.value=="on")  && (document.frmRIR.opfepcc.value==1 || document.frmRIR.opfonm.value==1)  && document.frmRIR.optClass.value!=""  && (!((document.frmRIR.iClassn.value)== 2  && (document.frmRIR.optClass.value)==3))  && !(((document.frmRIR.iClassn.value)== 3  && (document.frmRIR.optClass.value)==2)) )
			{	
				bConfirm = window.confirm('Changing this option will require QUEST to refresh this page and MPS details will be deleted .You need to resubmit MPS detail.');
			}
			else
			{
				bConfirm = window.confirm('Changing this option will require QUEST to refresh this page');					
			}		
        }
            else
			{
				bConfirm = window.confirm('Changing this option will require QUEST to refresh this page');					
			}				
        }
        //return (bConfirm) 
        if (bConfirm){
				
            var g1 = window.document.frmRIR.LossCat_G1;
				   
            if (g1 != null )
            {
						
                if(DisG1==1)
                {
                    DisG1=0
                }
            }
								
            var g2 = window.document.frmRIR.LossCat_G2; 
            if (g2 != null)
            {
                if(DisG2==1)
                {
                    DisG2=0
                }
            }
				   
            if (intCaller==3) //Fire/Explosion
            {
                if (!(isHSESQ) && document.frmRIR.rdFireExplosion[0].checked && document.frmRIR.rdWIBEventSQ[0].checked) 
                {
                    frmRIR.optHSE.checked = true;  
                    document.frmRIR.FireMode.value = 1;                      
                }  
                else if ((isHSESQ) && document.frmRIR.rdFireExplosion[0].checked && document.frmRIR.rdWIBEventSQ[0].checked)
                {                          
                    document.frmRIR.txtHazard.value = '<%=HazID%>';
                }  
            }
            if (intCaller==4) //Well Barrier Element Involved - HSE
            {
                if (!(isHSESQ) <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> && document.frmRIR.rdWIBEventHSE[0].checked <%End IF%>) frmRIR.optSQ.checked = true; 
                <%IF ((DiffEvtClass >= 1) and  (bHSE)) and (Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv))) then%>
					SetEventCategorisation()
				<%end if%>
			}			   
            if (intCaller==5) //Accidental Discharge - SQ
            {
                if (isHSESQ <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> && document.frmRIR.rdAccDischarge[0].checked <%End IF%>) frmRIR.LossCat_C1.checked = true;
                if (isHSESQ <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> && document.frmRIR.rdAccDischarge[1].checked <%End IF%>) frmRIR.LossCat_C1.checked = false;
                if (!(isHSESQ) <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> && document.frmRIR.rdAccDischarge[0].checked <%End IF%>) frmRIR.optHSE.checked = true;
            }
			<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
			if (intCaller == 7)
			{
				if (!(isHSESQ)) frmRIR.optHSE.checked = true; document.frmRIR.rdPLSSInv[1].disabled=true;
			}
			<%End IF%>
            if ((intCaller <= 2) || (intCaller >= 3 && intCaller <=6 && !(isHSESQ)) <%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%> || (intCaller == 7 && !(isHSESQ))<%End IF%>)
            {
                document.frmRIR.action='RIRdsp.asp<%=IIF (bNR,"?NR=1", sKey)%>&postvars=1';
                document.frmRIR.submit();
            }
            return true;
        }
        else{
            if (intCaller==2)
                document.frmRIR.optClass.value = document.frmRIR.hdClass.value;		
            else if (intCaller==3)			        
                document.frmRIR.rdFireExplosion[1].checked=true;    			  
            else if (intCaller==4)			
			<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
			{
				var ActivityType
				var radios = document.getElementsByName('rdEventSubCat');
				for (var i = 0, length = radios.length; i < length; i++) {
					if (radios[i].checked) {
						ActivityType = radios[i].value;  // alert(radios[i].value);
						break; // only one radio can be logically checked, don't check the rest
					}
				}
				if(document.getElementById("rdEventSafetyProc").checked && ActivityType > 0)
				{
					document.getElementById("rdEventSubCat"+ActivityType).checked=false;
				}
			}
			<%Else%>
                document.frmRIR.rdWIBEventHSE[1].checked=true; 			
            else if (intCaller==5)
                document.frmRIR.rdAccDischarge[1].checked=true;
			<%End IF%>		
            // TS Changes
            <%if bNR and EnforceFlag Then %>
            else if (intCaller==6)
            document.frmRIR.QuestLoc.value= document.frmRIR.QLoc.value; 
            <%End if%>
			else if (intCaller==7)	
				document.frmRIR.rdPLSSInv[1].checked=true;
            if (intCaller <= 2)
                return false;
        }
    }
    // *****************************************************************
    //  javascript function added for NPT <<2401608>>
    // *****************************************************************
    /*function chkClientAffectedNo(obj)
    {
        if((document.frmRIR.optSQInvment[2].checked)){
         
             document.frmRIR.LossCat_G1.checked=true ; 
             document.frmRIR.LossCat_G1.disabled=true ;
                             
           }
    }*/
    //****************************end changes****************************************************
			
    function chkClientAffected(obj){
        if(obj.checked){
            alert("If this report is Client Affected, please ensure that a CRM Client is selected before clicking Save Data and then enter Client Red Money $ on the appropriate Loss Tab")
        }

        if  (!(obj.optClass.options[obj.optClass.selectedIndex].value == '1'))
        {
            return ;
        }
        if((document.frmRIR.optSQInvment[0].checked)||(document.frmRIR.optSQInvment[1].checked)){
					 
            document.frmRIR.LossCat_G1.checked=true ;
            document.frmRIR.LossCat_G2.checked=true ;
            DisG2=1
            DisG1=1     
        }
    }
			
    function chkRegRec(obj){
        var RMsg
        RMsg="WARNING:\n\tSelect this option only if you are legally required to inform an external authority "
        RMsg=RMsg + "\n\t(Government Agency, Regulatory Body, Social Security, Local Authority) of the event. "
        RMsg=RMsg + "\n\n\tThe RIR description should be updated to include the authority involved "
        RMsg=RMsg + "\n\tand the ID#/Reference Number of the report filed with the authority."
        if(obj.checked){
            alert(RMsg)
        }
    }
			

    // ************************************************************************
    //  Javascripts to check/uncheck quality loss , created for NPT <<2401608>>
    // ************************************************************************
			
    function chkOptSLBExternal(obj){
        if(!(obj.optClass.options[obj.optClass.selectedIndex].value == '1'))
        {
            return ;
        }
        obj.optCAffect[0].checked=true;
        obj.LossCat_G2.checked=true;
        obj.LossCat_G1.checked=true;
        DisG2=1;
        DisG1=1;
        toggleNPTTable(document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value);
    }
			
    function chkOptSLBInternal(obj){
        if  (!(obj.optClass.options[obj.optClass.selectedIndex].value == '1'))
        {
            return ;
        }
        obj.LossCat_G2.checked=true;
        obj.optCAffect[1].checked=true;
        obj.LossCat_G1.checked=false;
        DisG2=1;
        DisG1=0;
        toggleNPTTable(document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value);

    }
			
			
    function chkOptThirdparty(obj){
			 
        if(!(obj.optClass.options[obj.optClass.selectedIndex].value == '1'))
        {
            return ;
        }
        obj.optCAffect[0].checked=true;
        DisG2=0;
        DisG1=0;
        obj.LossCat_G1.checked=false;
        obj.LossCat_G2.checked=false;
        toggleNPTTable(document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value);
    }
			
    function chkSlbExtClientAffectedNo(obj){
        if((obj.optSQInvment[0].checked) || (obj.optSQInvment[1].checked)){
            obj.LossCat_G1.checked=false ;
            obj.LossCat_G2.checked=true ;
            DisG2=1;
            DisG1=0;
            toggleNPTTable(document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value);
        }
    }
    function toggleNPTTable(SQSev)
    {

        if (SQSev >= 1 && document.frmRIR.optSQInvment[0].checked)
        {
            document.getElementById("tblNPT").style.display="";
        }
        else
        {
            document.getElementById("tblNPT").style.display="none";
        }
			
        if ((SQSev >2))   //swi changes
        {	
            hidediv("Q2B");
            hidediv("Q2Bans");
            hidediv("Q3A");
            hidediv("Q3Ans");
        }
        else   
        {
		if(document.frmRIR.swiqn != null)
		{
            if (document.frmRIR.swiqn[0].checked)
            {
                showdiv("Q2B");
                showdiv("Q2Bans");
                showdiv("Q3A");
                showdiv("Q3Ans");
            }
		}
        }
			

    }  	

    function toggleNPTRM()
    {
        if (document.frmRIR.rdNPT[0].checked)
        {
            document.getElementById("txtNPT_M").style.visibility="hidden";
            document.getElementById("txtNPT_LossCat_G1_M").style.visibility="hidden";
            document.getElementById("txtNPT_LossCat_G2_M").style.visibility="hidden";
        }
        else
        {
            document.getElementById("txtNPT_M").style.visibility="visible";
            document.getElementById("txtNPT_LossCat_G1_M").style.visibility="visible";
            document.getElementById("txtNPT_LossCat_G2_M").style.visibility="visible";
        }
    }  	

    function onCheck(obj){  
        <%if (LockCountSQ <> 0  or  chkSQLockingMgmt()) then%>		
                        if ((DisG1==1) && (obj.name=='LossCat_G1')){
                            //added by Nilesh for 2468497 to show alert message
                            fndisablenpt(obj);
                            return false;
                        }		
        if ((DisG2==1) && (obj.name=='LossCat_G2')){
            //added by Nilesh for 2468497 to show alert message
            fndisablenpt(obj);
            return false;
        }	
        <%end if%>				
        }
			

    //added by Nilesh for 2468497 to show alert message
    function fndisablenpt(obj)
    {
        if (!obj.checked)
        {

            if (((obj.name.indexOf("G2")> 1)) || ((obj.name.indexOf("G1")> 1)) )
            {

                if(( document.frmRIR.optSQInvment[1].checked)|| ( document.frmRIR.optSQInvment[0].checked) )
                {
                    alert('This loss is mandatory based on the Activity/Process/Service selected.');
                    obj.checked = true;
                }
            } 

        }
    }
    // ************************************************************************
    //-->
    function Supplier_onchange() { 
        var orgno='<%=lorgno%>'	            
        var UrlStr="../Utils/SearchTPSupplier.asp?Orgno="+orgno+"&optname=txtTPSupplier&source=0"
        open(UrlStr,'searchTPSupplier','height=450,width=900,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1') 
    }

    function addOptionSupplier(sText,sValue,optName){
        var opt,seq,temp			        
        sValue=sValue.replace(/'/i, "''");		
        seq=optName.substring(13);
        opt=document.getElementById('txtTPSupplier');
        opt.innerHTML = opt.innerHTML + "<Input Name='chkTPSupplier' onclick='SettxtPopup(this);' Type='CheckBox' value='"+sValue+":"+sText.replace("'","")+"' Checked >" + sText+ "</BR> "
			
    } 
    //swi changes
    function chkOptSWIqns(obj)
    {
        <%if (iclass = 1) then %>
        var sqsev = document.frmRIR.cmbSQSeverity[document.frmRIR.cmbSQSeverity.selectedIndex].value;
        if (sqsev ==1 || sqsev ==2)
        {
            if (obj.value == 1)
            {
                showdiv("Q2B");
                showdiv("Q2Bans");
                showdiv("Q3A");
                showdiv("Q3Ans");
            }
            else
            {
                hidediv("Q2B");
                hidediv("Q2Bans");
                hidediv("Q3A");
                hidediv("Q3Ans");
            }
        }
        <%Else if (iclass = 2) then%>

            if (obj.value == 1)
        {			
            showdiv("Q2A");
            showdiv("Q2Ans");
        }
    else
    {
            hidediv("Q2A");
        hidediv("Q2Ans");
    }

    <%End if%>
    <%End if%>
    }
		


    </script>

    <script language='JavaScript'>
		
		
		
		<% If bSQ and (HideSQ=1 or ShowIPMSQ=1 or SQCategoryMappingID > 0) Then %>
			function DynaSetter() {
			    this.dump = DynaSetter_Dump;
			    this.setSL= DynaSetter_SetSL;
			    this.load = DynaSetter_Load;
				
			    this.data = new Object();
			}

//Provide some debug output if needed
function DynaSetter_Dump() {
    var layerOne;
    var layerTwo;
			
    document.write ("<PRE>");
    for (layerOne in this.data) {
        document.write(layerOne+': ');
        for (layerTwo in this.data[layerOne]) {
            document.write('\t'+layerTwo+": "+this.data[layerOne][layerTwo]+"\n");
        }
    }
    document.write ("</PRE>");
}		
			
// Regens the child Select Box
function DynaSetter_SetSL(parentSelectBox,childSelectBox) 
{
   // alert(parentSelectBox.name);
    //alert(childSelectBox.name);
    if (!parentSelectBox) return;
    var psl = parentSelectBox;
    var children = this.data[psl.options[psl.selectedIndex].value];
    var childID;
    var csl = childSelectBox;
    //alert(children.name);
    csl.options.length = 0;
    //alert(typeof children);
    if (typeof children != 'undefined'){
        <%If  issaxon(iSubBSID) then%>
            csl.options[csl.options.length]=new Option("(Selection Required)",0);
        //	alert(csl.name);
        if (psl.name == 'txtSubFailure' || csl.name == 'txtDamage')
        {
            var sel_text =document.getElementById("txtFailure").options[document.getElementById("txtFailure").selectedIndex].text;
            var sel_text_subcat =document.getElementById("txtSubFailure").options[document.getElementById("txtSubFailure").selectedIndex].text;
            if ((sel_text == 'Surface Equipment Failure or Malfunction'|| sel_text == 'Downhole Equipment Failure or Malfunction' )  )  
            {
                document.getElementById("Saxon_Damage_cat").style.display='';
                document.frmRIR.ShowDamage.value = 'Y';
            }
            else
            {
                document.getElementById("Saxon_Damage_cat").style.display='none';
                document.frmRIR.ShowDamage.value = 'N';
            }
        }
        <%else%> 
					  		
             if (children.forced) {
                 csl.options[csl.options.length]=new Option("(Selection Required)",0);
             } else {
                    csl.options[csl.options.length]=new Option("(Not Specified)",-1);
    }
    <%End if%>	
					 
  csl.options[csl.options.length-1].selected = true;

    for (childID in children) {
        if (childID != 'forced') {
            csl.options[csl.options.length]=new Option(children[childID].text,childID);
            csl.options[csl.options.length-1].selected = children[childID].selected;
        }
    }
}
				
<%If  issaxon(iSubBSID) then%>
/*else	
    {
        if (psl.name =="txtSubFailure" && document.frmRIR.txtSubFailure.options.length != "0")
        {
            var sel_text_subcat =document.getElementById("txtSubFailure").options[document.getElementById("txtSubFailure").selectedIndex].text;
            if (sel_text_subcat=='Customer or Third Party Downhole')
            {document.getElementById("Saxon_Damage_cat").style.display='none';
            document.frmRIR.ShowDamage.value = 'N';
            }
        }
    }*/	
<%End if%>
}		

    //loads the values into the object
function DynaSetter_Load(pid,id,text,selected,forced) {
if (typeof this.data[pid] == 'undefined') this.data[pid] = new Object();

this.data[pid][id] = new Object;
this.data[pid].forced = forced;
this.data[pid][id].text = text;
				
this.data[pid][id].selected = selected;
				
}
			
    var spcategory = new DynaSetter;
var failure = new DynaSetter;
var damage = new DynaSetter;
//		var cause = new DynaSetter;
var saxcat  = new DynaSetter;		
<%=Saxoncats%>

<%
		
		
SET RSSQ = Server.CreateObject("ADODB.Recordset")

'if (iBSID = 9186 or iBSID = 9187) then tlPL = 120 else tlPL = lPL
'if (isMNSIT(iBSID)) then tlPL = 113
'if (isOneCPL(iBSID)) then tlPL = 107

SQL=" SELECT PID, SC.ID,C.Type ,SubDescription AS Name, CASE  WHEN SC.ID in (0,"& SQSPSubcatID &","& SQCSubcatID &","& SQFSubcatID &","& SQDSubcatID&") then 'true' else 'false' end as Selected"
SQL=SQL & " FROM tlkpSQSubCategories SC INNER JOIN tlkpSQCategories C ON SC.PID=C.ID "
SQL=SQL & " WHERE SC.PLID IN (0," & tlPL & ")  and ((Status+substatus=0) or(SC.ID in (0,"& SQSPSubcatID &","& SQCSubcatID &","& SQFSubcatID &","& SQDSubcatID&"))) ORDER BY C.Type, PID, Name"
Response.write "//SQL:SELECT PID, SC.ID,C.Type ,SubDescription AS Name, CASE SC.ID WHEN " & SQSPSubcatID & " THEN 'true' WHEN " & SQFSubcatID & " THEN 'true'  WHEN " & SQDSubcatID & " THEN 'true'  WHEN " & SQCSubcatID & " THEN 'true' ELSE 'false' END AS Selected FROM tlkpSQSubCategories SC INNER JOIN tlkpSQCategories C ON SC.PID=C.ID WHERE SC.SubStatus=0 and SC.PLID IN (0," & lPL & ") ORDER BY C.Type, PID, Name"
Response.write "//SQL:" & SQL &vbCRLF
RSSQ.Open SQL,cn
Response.Write vbCRLF
while not RSSQ.eof
If RSSQ("Type")="S" Then Response.Write vbTab & vbTab & "spcategory.load('" & RSSQ("PID") & "','" & RSSQ("ID") & "','" & replace(RSSQ("Name"),"'","\'") & "'," & RSSQ("Selected") & ",true);" & vbCRLF
If RSSQ("Type")="F" Then Response.Write vbTab & vbTab & "failure.load('" & RSSQ("PID") & "','" & RSSQ("ID") & "','" & replace(RSSQ("Name"),"'","\'") & "'," & RSSQ("Selected") & ",true);" & vbCRLF
If RSSQ("Type")="D" Then Response.Write vbTab & vbTab & "damage.load('" & RSSQ("PID") & "','" & RSSQ("ID") & "','" & replace(RSSQ("Name"),"'","\'") & "'," & RSSQ("Selected") & ",true);" & vbCRLF
RSSQ.movenext
wend

End If
%>

        // sps changes started
        function DynaSetter1() {
            this.dump = DynaSetter_Dump1;
            this.setSL= DynaSetter_SetSL1;
            this.load = DynaSetter_Load1;
				
            this.data = new Object();
        }

//Provide some debug output if needed
function DynaSetter_Dump1() {
    var layerOne;
    var layerTwo;
			
    document.write ("<PRE>");
    for (layerOne in this.data) {
        document.write(layerOne+': ');
        for (layerTwo in this.data[layerOne]) {
            document.write('\t'+layerTwo+": "+this.data[layerOne][layerTwo]+"\n");
        }
    }
    document.write ("</PRE>");
}		
			
// Regens the child Select Box
function DynaSetter_SetSL1(parentSelectBox,childSelectBox) 
{

	//alert(parentSelectBox.name);
	//alert(parentSelectBox.value);

if(parentSelectBox.name == 'SQL3_0' &&  parentSelectBox.value=="" )
    {	
        document.getElementById("Pro").style.display='none';
       document.getElementById("Act").style.display='none';
	     document.getElementById("DescpSpan").style.display='none';
       document.getElementById("DescpSpan1").style.display='none';
	   
    }
    var f =document.frmRIR
    if (!parentSelectBox) return;
    var psl = parentSelectBox;
    if(typeof psl.options[psl.selectedIndex] != "undefined" && typeof psl.options[psl.selectedIndex] != "unknown")  //Error was throwing when parent element was not found i.e. Metro Stop and Activity were not selected.
	{	
		var children = this.data[psl.options[psl.selectedIndex].value];
		var childID;
		var csl = childSelectBox;
		if (csl.options != undefined){
			csl.options.length = 0;
		}
		if (parentSelectBox.name == 'cat_sq' && parentSelectBox.value ==' ')
		{
			csl.options[csl.options.length]=new Option("(Select a Subcategory)",' ');
		}
			
		if (typeof children != 'undefined'){
						
			if (children.forced) {
				csl.options[csl.options.length]=new Option("(Selection Required)",0);
			} else {
							
				if (childSelectBox.name == 'subcat_sq')
				{
					csl.options[csl.options.length]=new Option("(Select a Subcategory)",' ');
				}
				else
				{
					csl.options[csl.options.length]=new Option("(Not Specified)",-1);}
			}
			
			for (childID in children) {
				if (childID != 'forced') {
					csl.options[csl.options.length]=new Option(children[childID].text,children[childID].id);
					//csl.options[csl.options.length-1].selected = children[childID].selected;
				}
			}

		   
			var e = document.getElementById("SQL3_0");
			var SQL3Indx
			if (parentSelectBox.name != 'SQB2_0')
			{
				if (csl.name == "SQL3_0"){
					strUser = e.options[e.selectedIndex].value;
					if (strUser == "0"){
						//document.frmRIR.SQL4_0.options[0].selected = true;
						//document.getElementById("DescpSpan").style.display='none';
						//document.getElementById("DescpSpan1").style.display='none';
					}
					if (SQL3Indx == "0"){
						document.frmRIR.SQL3_0.options[0].selectedIndex;
					}
				}
			}else{
			
				//csl.options[csl.options.length-1].selected = true;
					  
				document.getElementById("Pro").style.display='';
			   // document.getElementById("Act").style.display='';
						
				subaffservice.setSL(f.SQL2_0,f.SQL3_0);	
				//incident.setSL(f.SQL3_0,f.SQL4_0);			
					   
				if ((csl.options != undefined) && (parentSelectBox.value == 0)) {
				
					document.getElementById("Pro").style.display='';
					document.getElementById("Act").style.display='';
				}        		   
					   
			}	
		}
	}
			
		//alert("name- " + parentSelectBox.name +" Value- " + parentSelectBox.value );	
    //alert("parent- " + parentSelectBox.name +" child- " + childSelectBox.name );
    //if(parentSelectBox.name == 'SQL2_0'  &&  parentSelectBox.value != '38' &&  parentSelectBox.value != '0' )
    if(parentSelectBox.name == 'SQL2_0' &&  parentSelectBox.value >0 )
    {	
	
        document.getElementById("Pro").style.display='';
    
    }

	
	    if(parentSelectBox.name == 'SQL3_0' &&  parentSelectBox.value >0 )
    {	
        document.getElementById("Pro").style.display='';
       document.getElementById("Act").style.display='';
	     document.getElementById("DescpSpan").style.display='';
       document.getElementById("DescpSpan1").style.display='';
	   
    }
	else
	{
	
	   document.getElementById("Act").style.display='none';
       document.getElementById("DescpSpan").style.display='none';
       document.getElementById("DescpSpan1").style.display='none';
	 
	}
	
  if(parentSelectBox.name == 'SQL2_0' &&  parentSelectBox.value == 0)  
    {
        document.getElementById("SQL3_0").value = "0";
        document.getElementById("SQL4_0").value = "0";
        document.getElementById("Pro").style.display='none';
        document.getElementById("Act").style.display='none';
        document.getElementById("DescpSpan").style.display='none';
        document.getElementById("DescpSpan1").style.display='none';
    }
	
	if ((document.getElementById("SQL2_0").value == "") || (document.getElementById("SQL3_0").value == ""))
    {
        document.getElementById("SQL3_0").value = "0";
        document.getElementById("SQL4_0").value = "0";
    }
				
    //if ((parentSelectBox.value == '0'  || parentSelectBox.value == '38') || parentSelectBox.value == '9' )   //38 for Other
    if (parentSelectBox.value == '0')   //38 for Other
    {
        $("#"+childSelectBox.name).append('<option value=0>(Selection Required)</option>');
        if (childSelectBox.name== 'SQL3_0')
        {
            incident.setSL(f.SQL3_0,f.SQL4_0);
            //$("#SQL4_0").append('<option value=0>(Selection Required)</option>');
        }
					 
        document.getElementById("DescpSpan").style.display='none';
        document.getElementById("DescpSpan1").style.display='none';
    }
	
	
				
    var countL2 = document.frmRIR.cntL2.value;
	
		//alert(parentSelectBox.name)
				
    if (parentSelectBox.name == 'SQL3_0' && countL2 > 0)
    {
	
	//alert(parentSelectBox.name)
			//	alert(countL2)
				    
        document.getElementById("DescpSpan").style.display='none';
        document.getElementById("DescpSpan1").style.display='none';
					
    }
	
	if (document.frmRIR.SQL3_0.value=="")
	{
	        document.getElementById("Pro").style.display='none';
			document.getElementById("Act").style.display='none';
			
	}
	
    document.frmRIR.cntL2.value = parseInt(document.frmRIR.cntL2.value) + 1;
	
	
	
	var prosel = document.frmRIR.SQL2_0;
	 if (typeof prosel != "undefined"){
		if (prosel.options.length == 0)
		{
		
			document.getElementById("Pro").style.display='none';
						


		}
	}
	var prosel = document.frmRIR.SQL3_0;
	 if (typeof prosel != "undefined"){
		if (prosel.options.length == 0)
		{
			document.getElementById("DescpSpan").style.display='none';
		}
	}
	var prosel14 = document.frmRIR.SQL4_0;
	 if (typeof prosel14 != "undefined"){
		if (prosel14.options.length == 0)
		{
			document.getElementById("DescpSpan").style.display='none';
			document.getElementById("Act").style.display='none';
			 
		}	
		
	}
	
}


	

function descdisp(sp4)
{


//alert("dsfdsfsd");

if (sp4!="" && sp4!=0 )
{
//alert("sadsadsa")
document.getElementById("DescpSpan").style.display='';
document.getElementById("DescpSpan1").style.display='';
}
}



	

//loads the values into the object
function DynaSetter_Load1(pid,id,text,selected,forced,order) {
    if (typeof this.data[pid] == 'undefined') this.data[pid] = new Object();

    this.data[pid][order] = new Object;
    this.data[pid].forced = forced;
    this.data[pid][order].text = text;
	this.data[pid][order].id = id;
				
    this.data[pid][order].selected = selected;
				
}
			
				
//Saxon Changes	
var DT_SQOperation = new DynaSetter1;
<%
    LoadDT_SQOperation "DT_SQOperation",OperationSubCat  
%>
function frmload(){
    DT_SQOperation.setSL(document.frmRIR.cat_sq,document.frmRIR.subcat_sq,2);
}	
//Till Here 
    // SPS changes
    <%If bSQ and isSPS>0 then%>
    var affservice = new DynaSetter1;				
var subaffservice = new DynaSetter1;
var incident = new DynaSetter1;
    
<%=SQProcCats%>
			
<%End if%>
			
		

function selval(childSelectBox,val)
{
    var csl = childSelectBox;
    csl.options.length = 0;
    if (val ==1)
{csl.options[csl.options.length]=new Option("(Selection Required)",0);}
}		 
    </script>

    <script language="JavaScript1.2">
		function IPMHelp(){
			window.open('IPMHelp.htm','IPM','height=240,width=300,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0,alwaysRaised=1')
		}
		
		function TCCHelp(){
			window.open('TCCHelp.htm','TCC','height=240,width=300,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0,alwaysRaised=1')
		}
		
		function CRMRigHelp(){
						window.open('<%=VarSLBHub%>Docs/quality/videos/Rig%20selection/Rig%20selection.html','CRMRIG','height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0,alwaysRaised=1')
		}
			
		function SFOInvHelp(){
			window.open('../Utils/VideoPg.asp?vid=SFOInv','SegFunOrgInvolved','height=600,width=750,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0,alwaysRaised=1')
		}

		function WIBHelp(){
			window.open('../Utils/VideoPg.asp?vid=WB','WellBarrier','height=600,width=750,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0,alwaysRaised=1')
		}

		function On_Failure(){
		var fi,sfi
		fi=document.frmRIR.txtFailure.selectedIndex;
		sfi=document.frmRIR.txtFailure.options[fi].value;
		if(sfi==449)
			open("../utils/FailureControl.htm","FailureControl","height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=1,alwaysRaised=1")
		
		if(sfi==1112)
			open("../utils/StuckPipe.htm","StuckPipe","height=400,width=500,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=1,alwaysRaised=1")
		}
				
			function checkifDirectDMSQ()
			{
				var SiteV,Site;
				//PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.
				//Checks only for D&M Sub Segments
				<%if bNR and isDMSeg(iBSID) Then %>
				//if(document.frmRIR.PLID.value in { 1:1,120:1})
				//{
					SiteV = document.frmRIR.txtLocation.options[document.frmRIR.txtLocation.selectedIndex].value;
					Site = SiteV.split(":",1);
					if(document.frmRIR.optSQ.checked)
					{
						if(Site in {3:1,11:1,12:1,13:1,14:1})
						{
						alert("All Job Related SQ RIR's MUST come from eTrace Incident List");
						}
					}
				//}
				<%End IF %>
			}
			
			
			function selectOpt(sel,val)
			{
					var cnt;
					if (typeof sel != "undefined")
					{
						cnt=sel.options.length;
						for(i=0;i<cnt;i++){
							if(sel.options[i].value==val){
								sel.options.selectedIndex=i;
								break;
							}
						}
					}					
			}
				// SPS changes
			function SPS_SelAssign(SPS_l2,SPS_l3,SPS_l4,SPS_b2)
			{
			//alert(SPS_l2+"SPS_l3"+SPS_l3+"SPS_l4::"+SPS_l4+"::SPS_b2::"+SPS_b2)
				var f =document.frmRIR;
				selectOpt(f.SQB2_0,SPS_b2);
				affservice.setSL(f.SQB2_0,f.SQL2_0);
				
				selectOpt(f.SQL2_0,SPS_l2);
				subaffservice.setSL(f.SQL2_0,f.SQL3_0);
				
				if((f.SQL2_0.value != '38') && (f.SQL2_0.value != '0') )
				{
				
					document.getElementById("Pro").style.display='';
					document.getElementById("Act").style.display='';
				}
				
				selectOpt(f.SQL3_0,SPS_l3);
				incident.setSL(f.SQL3_0,f.SQL4_0);
					
				selectOpt(f.SQL4_0,SPS_l4);
					
			}
			
			function getDescription (id)
			{
				
				document.getElementById("DescpSpan").style.display='';
				document.getElementById("DescpSpan1").style.display='';
			
				if (id.value==0 )
				{
					document.getElementById("DescpSpan").style.display='none';
					document.getElementById("DescpSpan1").style.display='none';
				}
				
				var Prev_L4 = document.frmRIR.crntdescp.value;   //previous value
				document.getElementById(id.value).style.display = '';
				
				//alert("previous value- " + Prev_L4);
				if ((Prev_L4 != id.value) && (Prev_L4 != 0))
				{
					//alert(Prev_L4);
					document.getElementById(Prev_L4).style.display = 'none';
							// storing for future call
				}
				document.frmRIR.crntdescp.value= id.value;
				
				
					divObj = document.getElementById(id.value);
				if ( divObj )
					{
						if ( divObj.textContent )
						{ // FF
                    //alert ( divObj.textContent );
						var txtdesc=divObj.textContent;
					
						}
						else
						{  // IE           
					
							document.getElementById("DescpSpan").style.display='none';
							document.getElementById("DescpSpan1").style.display='none'; 
                   /// alert ( "678" );  //alert ( divObj.innerHTML );
					
						} 
					} 
				//alert(document.frmRIR.crntdescp.value);
			}
			
			//SQ Category Changes
			  function SQCat_SelAssign(SQCAT_l,SQCAT_2,SQCAT_3, SQCAT_4)
			{
				var f =document.frmRIR;
				selectOpt(f.txtFailure,SQCAT_l);
				saxcat.setSL(f.txtFailure,f.txtSubFailure);
				
				selectOpt(f.txtSubFailure,SQCAT_2);
				
				
				if(SQCAT_3 !='0')
				{
					saxcat.setSL(f.txtSubFailure,f.txtDamage);
					selectOpt(f.txtDamage,SQCAT_3);
					saxcat.setSL(f.txtDamage,f.txtSubDamage);
					selectOpt(f.txtSubDamage,SQCAT_4);	
				}	
			}
    </script>

</head>
<!---<BODY onload="initForm()">--->
<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" rightmargin="0"
    onload="fncChangeHazardCat();descdisp('<%=SPS_L4_VALUE%>');" onclick="hideDropdown();">

    <script language="JavaScript" src="../inc/wz_tooltip.js">
    </script>

    <%If Not bNR Then displaymenubar(RS)%>
    <%
	  if ACLDefined Then 
		DisplayConfidential()
		
	  End IF
    %>
    <%If Not bNR Then DisplayWarning(RS("Source"))%>
    <%
	If Not bNR Then 
		lBSegmentID = Trim(RS("BusinessSegment"))
		lDepartmentID = Trim(RS("Department"))
	End If

    %>

    <script language="JavaScript1.2">
	
	function Checkdataforfilter()
			{
			 <%if bSQ then %>
			ipmnocheck();
			<%End if%> 
			
			    if(document.getElementById("proNO").checked==true)
			   {
			   <%if projectunknown<>"" then %>
				document.getElementById("prounknow").checked=false;
				<%End if%>
			 
			   document.getElementById("proIDS").checked=false;
			   //document.getElementById("proIPS").checked=false;
			   document.getElementById("proIFS").checked=false;
			   //document.getElementById("proISM").checked=false;
			   document.getElementById("proSPM").checked=false;
			   
			   document.getElementById("proIDS").disabled=false;
			   //document.getElementById("proIPS").disabled=false;
			   document.getElementById("proIFS").disabled=false;
			   //document.getElementById("proISM").disabled=false;
			   document.getElementById("proSPM").disabled=false;
			    <%if bSQ then %>
			   document.getElementById("optInv").checked=false;
			   document.getElementById("optInvno").checked=true;
				<%End if%> 
			   }

            }
	
	function CheckdataforUnkownfilter()
			{
			
			    if(document.getElementById("prounknow").checked==true)
			   {
			  
			   document.getElementById("proNO").checked=false;
			   document.getElementById("proIDS").checked=false;
			   //document.getElementById("proIPS").checked=false;
			   document.getElementById("proIFS").checked=false;
			   //document.getElementById("proISM").checked=false;
			   document.getElementById("proSPM").checked=false;
			   
			   document.getElementById("proIDS").disabled=false;
			   //document.getElementById("proIPS").disabled=false;
			   document.getElementById("proIFS").disabled=false;
			   //document.getElementById("proISM").disabled=false;
			   document.getElementById("proSPM").disabled=false;
			    <%if bSQ and SegInv="False" then %>
			   document.getElementById("optInv").checked=false;
			   document.getElementById("optInvno").checked=true;
				<%End if%> 
			   }

            }
			
	function checkdataforGRC()
	{
	if (document.getElementsByName("optgot").value=1)
	{
	document.getElementById("optInv").checked=true;
	document.getElementById("optInvno").checked=false;
	}
	}
	
	
	function CheckdataforfilterByInvoledseg()
			{
			 <%if bSQ then %>
			    if(document.getElementById("optInvno").checked==true)
			   {
			 
			   document.getElementById("proIDS").checked=false;
			   //document.getElementById("proIPS").checked=false;
			   document.getElementById("proIFS").checked=false;
			   //document.getElementById("proISM").checked=false;
			   document.getElementById("proSPM").checked=false;
			   
			   document.getElementById("proIDS").disabled=false;
			   //document.getElementById("proIPS").disabled=false;
			   document.getElementById("proIFS").disabled=false;
			   //document.getElementById("proISM").disabled=false;
			   document.getElementById("proSPM").disabled=false;
			   document.getElementById("proNO").checked=true

			   }
			   <%End if%> 
            }
	

	function Checkdataforfilterbyseg()
			{
	
			 <%if bSQ then %>
			ipmcheck();
			<%End if%> 
			
			   //if(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proISM").checked==true || document.getElementById("proSPM").checked==true)
			   if(document.getElementById("proIDS").checked==true || document.getElementById("proIFS").checked==true || document.getElementById("proSPM").checked==true)
			   {
			    document.getElementById("proNO").checked=false;
				<%if projectunknown<>"" then %>
				document.getElementById("prounknow").checked=false;
				<%End if%> 
				
				<%if bSQ then %>
				document.getElementById("optInv").checked=true;
				<%End if%> 
			   }
			   else
			   {
			    document.getElementById("proNO").checked=true;
				<%if bSQ then %>
				document.getElementById("optInv").checked=false;
				document.getElementById("optInvno").checked=true;
				<%End if%> 
			   }
			   
			   <%if isIPMWCSS(iBSID)  then %>
			   document.getElementById("proIDS").checked=true;
			   document.getElementById("proIDS").disabled=true;
			   //if(document.getElementById("proIPS").checked==false && document.getElementById("proISM").checked==false && document.getElementById("proSPM").checked==false)
			   //if(document.getElementById("proIFS").checked==false && document.getElementById("proISM").checked==false && document.getElementById("proSPM").checked==false)
			   if(document.getElementById("proIFS").checked==false && document.getElementById("proSPM").checked==false)
			   {
			   <%if bSQ then %>
			   document.getElementById("optInvno").checked=true;
			   document.getElementById("optInv").checked=false;
			   <%End if%> 
			   }
			   <%End if%> 
			   
			    <%if isIPMAPS(iBSID)  then %>
			   document.getElementById("proSPM").checked=true;
			   document.getElementById("proSPM").disabled=true;
			    //if(document.getElementById("proIDS").checked==false && document.getElementById("proIPS").checked==false && document.getElementById("proISM").checked==false)
				//if(document.getElementById("proIDS").checked==false  && document.getElementById("proIFS").checked==false &&  document.getElementById("proISM").checked==false)
				if(document.getElementById("proIDS").checked==false  && document.getElementById("proIFS").checked==false)
			   {
			   <%if bSQ then %>
			   document.getElementById("optInvno").checked=true;
			   document.getElementById("optInv").checked=false;
			   <%End if%> 
			   }
			   <%End if%> 
			   
			    <%if isIPMPRSS(iBSID)  then %>
			   //document.getElementById("proISM").checked=true;
			   //document.getElementById("proISM").disabled=true;
			   //if(document.getElementById("proIDS").checked==false && document.getElementById("proIPS").checked==false && document.getElementById("proSPM").checked==false)
			   //if(document.getElementById("proIDS").checked==false  && document.getElementById("proIFS").checked==false && document.getElementById("proSPM").checked==false)
			   //{
			   <%if bSQ then %>
			   //document.getElementById("optInvno").checked=true;
			   //document.getElementById("optInv").checked=false;
			   <%End if%> 
			   //}
			   <%End if%> 
			   
			    <%if isIPMIFS(iBSID) then %>
			   //document.getElementById("proIPS").checked=true;
			   //document.getElementById("proIPS").disabled=true;
			   document.getElementById("proIFS").checked=true;
			   document.getElementById("proIFS").disabled=true;			   
			    //if(document.getElementById("proIDS").checked==false && document.getElementById("proISM").checked==false && document.getElementById("proSPM").checked==false)
				if(document.getElementById("proIDS").checked==false && document.getElementById("proSPM").checked==false)
			   {
			   <%if bSQ then %>
			   document.getElementById("optInvno").checked=true;
			   document.getElementById("optInv").checked=false;
			     <%End if%> 
			   }
			   <%End if%> 
			   
			}
		function validateLossSafetynet(){
			var frmdoc = document.frmRIR;
			var saftynetMsg = '\nAs Total Lost Days >= 180 report needs to change to Permanent Impairment and reclassified to Major.'
						<%If lossSafetynetVal = 1 or lossSafetynetVal = 2 then  %>
							<%If bHSE then %>
							<%If lossSafetynetVal = 2 then  %>
									if(frmdoc.cmbHSESeverity.options[frmdoc.cmbHSESeverity.selectedIndex].value!=3 && frmdoc.cmbHSESeverity.options[frmdoc.cmbHSESeverity.selectedIndex].value!=4) {
										return saftynetMsg
									}
								<%End if %>
								<%If lossSafetynetVal = 1 then  %>
										return saftynetMsg
								<%End if %>
							<%End if %>	  
					<%End if %>
				return '';
		}
	
		function txtClose_onclick() {
			var bstatus = true;
			var k,msg
			msg='<%=CheckRIRClose()%>'
				//if(document.frmRIR.chkClosed.value)
				if(document.frmRIR.chkClosed.checked)
				{
					msg=msg+ validateLossSafetynet();
					if (document.frmRIR.optSQ.checked && document.frmRIR.optHSE.checked)
					{
						if (document.frmRIR.rdPLSSInv.value == document.frmRIR.rdEventSafety.value)
						{
						var result=confirm('There is a mismatch of categorization of the Service & Equipment Specific Safety categorization on this report.\n\nOK - will set the report as a Service & Equipment Specific Safety Event\n\nCancel - will REMOVE the Service & Equipment Specific Safety Event categorization from the RIR');
						if (result==true)
						{
						document.getElementById("rdEventSafetyProc").checked = true;	
						document.getElementById("rdEventSafetyPers").checked = false;
						SetEventSubSafety();
						document.getElementById("rdPLSSInvID1").checked = true;
						document.getElementById("rdPLSSInvID2").checked = false;
						bstatus = false;
						}
						else{
						document.getElementById("rdEventSafetyPers").checked = true;
						document.getElementById("rdEventSafetyProc").checked = false;
						SetEventSubSafety();
						document.getElementById("rdPLSSInvID2").checked = true;
						document.getElementById("rdPLSSInvID1").checked = false;
						bstatus = false;
						
						}
					    }
					}
		if (document.frmRIR.optSQ.checked)
		{		
		<%if (isREWSQMapping(SQMappingID) and isROP > 0 and bSQ) then %>
		
		if (document.frmRIR.optROPInv.value != <%= wlro %>)		
		{
		var result=confirm('This report is not tagged as Remote Operations Involved \n\nPlease Accept to tag the report as RO Involved\n\nCancel to remove RO NPT values from this page.');
		if (result == true)
		{
		document.getElementById("hdnchkinput").value = "True";				
		document.getElementById("optROPInvY").checked = true;		
		bstatus = false;
		}
		else
		{		
		document.getElementById("hdnchkinput").value = "False";	
		document.getElementById("optROPInvN").checked = true;		
		bstatus = false;
		}
		}
		else
		{
		<%if(WlRo <> SqRo) then %>		
		var result=confirm('This report is not tagged as Remote Operations Involved \n\nPlease Accept to tag the report as RO Involved\n\nCancel to remove RO NPT values from this page.');
		if (result == true)
		{
		document.getElementById("hdnchkinput").value = "True";				
		document.getElementById("optROPInvY").checked = true;		
		bstatus = false;
		}
		else
		{		
		document.getElementById("hdnchkinput").value = "False";	
		document.getElementById("optROPInvN").checked = true;		
		bstatus = false;
		}
		<%End if%>
		}
		<%End if%>

		}
					if(msg!=''){
						msg='This report is incomplete. The following must be completed before the report can be closed.\n\n' + msg 
						window.alert(msg);
						bstatus = false;
					}
					return (bstatus) 
				}
			}

			// *****************************************************************
			//  Javascript functions added for NPT <<2401608>> , check IPM location and bussiness logic and 
			//  give popup accordingly. 
			// *****************************************************************

			function ipmcheck()
			{
				var ipmloc ;
				var msg,msg1;
						
				msg='<%=imploc()%>';
				//msg1='<%=IPMText(1)%>'
				msg1='Integrated Performance Management (IPM) Related?'
				
				if (msg.toLowerCase()  =='false')
				{
					
					if (document.frmRIR.optSQInvment[1].checked)
					{
						alert ('You cannot set '+ msg1 +' to Yes so long as "SLB managed and delivered to internal customer" is set to Yes');
									  document.getElementById("proIDS").checked=false;
									   //document.getElementById("proIPS").checked=false;
									   document.getElementById("proIFS").checked=false;
									   //document.getElementById("proISM").checked=false;
									   document.getElementById("proSPM").checked=false;
									   document.getElementById("proIDS").disabled=false;
									   //document.getElementById("proIPS").disabled=false;
									   document.getElementById("proIFS").disabled=false;
									   //document.getElementById("proISM").disabled=false;
									   document.getElementById("proSPM").disabled=false;
									   document.getElementById("proNO").checked=true ;
						return
					}
					else
					  {
						document.frmRIR.optSQInvment[1].disabled=true;
					  }
				}
			} 
			
			function ipmnocheck()
			{
				   {
					document.frmRIR.optSQInvment[1].disabled=false;
				   }
			} 
			
			
			// ************** End changes ***************************************************

			function fncToggleWIB(SrcVal,Caller)
			{
				var isHSESQ,frm;
				frm = eval(document.frmRIR);
				if (document.frmRIR.optSQ.checked && document.frmRIR.optHSE.checked)
					isHSESQ = true;

				if (Caller == 2) //Well Barrier Element Involved - SQ
				{
					if (isHSESQ && SrcVal==1) 
					{	
						<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
						frm.rdWIBEventHSE[0].checked = true;						
						frm.rdAccDischarge[0].checked = true;
						<%End IF%>
					}
					if (isHSESQ && SrcVal==0) 
					{
						<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
						frm.rdWIBEventHSE[1].checked = true;						
						frm.rdAccDischarge[1].checked = true;
						<%End IF%>
					}
						 <%IF ((DiffEvtClass >= 1) and  (bHSE))  then%>
							SetEventCategorisation()
							<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
								SetEventSubSafety()
							<%End IF%>
						<%end if%>
					<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
					if ((!(isHSESQ)) && SrcVal==1)
					{
						frm.rdAccDischarge[0].disabled = false;
						frm.rdAccDischarge[1].disabled = false;
						frm.rdFireExplosion[0].disabled = false;
						frm.rdFireExplosion[1].disabled = false;
						
						frm.rdAccDischarge[0].checked = true;
						optHSESQ_onchange(5);
					}
					if ((!(isHSESQ)) && SrcVal==0)
					{
						frm.rdAccDischarge[1].checked = true;
						frm.rdAccDischarge[0].disabled = true;
						frm.rdAccDischarge[1].disabled = true;
						frm.rdFireExplosion[1].checked = true;
						frm.rdFireExplosion[0].disabled = true;
						frm.rdFireExplosion[1].disabled = true;
						
						frm.rdAccDischarge[1].checked = true;
					}
					<%else%>
					if ((!(isHSESQ)) && SrcVal==1)
					{
						optHSESQ_onchange(5);
					}
					<%End IF%>
				}
				<%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>
				else if (Caller == 3) //Accidental Discharge - HSE
				{
					if (isHSESQ && SrcVal==true) frm.rdAccDischarge[0].checked = true;
					if (isHSESQ && SrcVal==false) frm.rdAccDischarge[1].checked = true;								
					if ((!(isHSESQ)) && SrcVal==true)
					{                   
						frm.rdWIBEventHSE[0].disabled = false;
						frm.rdWIBEventHSE[1].disabled = false;
					}
					if ((!(isHSESQ)) && SrcVal==false)
					{                   
						frm.rdWIBEventHSE[1].checked = true;
						frm.rdWIBEventHSE[0].disabled = true;
						frm.rdWIBEventHSE[1].disabled = true;
					}					
					SetEventCategorisation();					
				}
				<%End IF%>
			}
    </script>

    <form name="frmRIR" method="post" action="RIRdsp2.asp<%=sKey%>&slbin=<%=SLBInvment%>&sqlnv=<%=SQInvment%>&sqClientAffect=<%=ClientAffect%>">
    <style>
        .LockWarning
        {
            display: inline-block;
            color: #cc0000;
            background-color: #ffff99;
            font-weight: 900;
            font-size: 14px;
            text-align: center;
        }
    </style>
    <%If iClass=1 and iHSESev <> 1 and bHSE and not chkHseLockingMgmt() Then%>
    <table border="0" align="center" width="100%">
        <tr>
            <td align="center">
                <span id='Warning21' class='LockWarning'>
                    <%Response.Write sHSEWarningText%></span>
            </td>
        </tr>
    </table>
    <%END IF%>
    <%If iClass=1 and bSQ and  not chkSQLockingMgmt() Then%>
    <table border="0" align="center" width="100%">
        <tr>
            <td align="center">
                <span id='Warning22' class='LockWarning'>
                    <%Response.Write sHSEWarningTextSQ%></span>
            </td>
        </tr>
    </table>
    <%END IF%>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            <td class="boxednote" id="styleTiny" colspan="6">
                <%=mSymbol%>
                - Mandatory fields.
            </td>
        </tr>
        <tr class="reportheading">
            <td align="left" colspan="3">
                <input type="hidden" name="txtlock" value="<%=LockCount%>">
                Report Date:&nbsp;
                <%=RptDate%>
            </td>
            <td align="right" colspan="3" class="field">
                Report Number:&nbsp; <span class="urgent">
                    <%=ReportNumber%>
                    <input type="hidden" name="CreatorID" value="<%=DisplayQuotes(sUID)%>">
                    <input type="hidden" name="CreatorNm" value="<%=DisplayQuotes(sUName)%>">
            </td>
        </tr>
        <%If Not bNR Then%>
        <%If Not isGuest() Then%>
        <tr>
            <td align="left" colspan="4">
                <span class="field">Created By:&nbsp;</span> <span class="normal">
                    <%=fncGetDirLink(CreatedBy,CreateUID)%>
                </span>
            </td>
            <td align="right" colspan="1" class="field">
                Acknowledged:&nbsp;
                <input <%If bReviewed Then Response.Write "checked "%> name="chkReviewed" type="checkbox">&nbsp;&nbsp;&nbsp;
            </td>
            <td align="right" colspan="1" class="field">
                Closed:&nbsp;
                <input <%If bClosed Then Response.Write "checked "%> name="chkClosed" type="checkbox"
                    onclick='return txtClose_onclick()'>&nbsp;&nbsp;&nbsp;
            </td>
        </tr>
        <%End If%>
        <tr id="HSEBLOCK1">
            <td align="left" colspan="4">
                <span class="field">Updated On:&nbsp;</span> <span class="normal">
                    <%=RevDate%>
                    by
                    <%=fncGetDirLink(UpdatedBy,UpdateUID)%>
                </span>
            </td>
            <td align="right" class="field">
                Service Quality:&nbsp;
                <input <%If bSQ then Response.Write "checked "%> name="optSQ" type="checkbox" onclick='return optHSESQ_onchange(1)'>&nbsp;&nbsp;&nbsp;
            </td>
            <td align="right" class="field">
                HSE:&nbsp;
                <input <%If bHSE then Response.Write "checked "%> name="optHSE" type="checkbox" onclick='return optHSESQ_onchange(1)'>&nbsp;&nbsp;&nbsp;
            </td>
        </tr>
        <tr id="HSEBLOCK2">
            <!-- @ visali 06/02/2004-->
            <td align="left" colspan="4">
                <span class="field">Source:</span> <span>
                    <%If Not bNR Then Response.write getSourceName(rs("Source"))%></span> &nbsp;
                &nbsp;
                <!--@@Visali 06-29-2004 -->
                <%If Rs("ForeignID") <> "" then %>
                <span class="field">Foreign ID:</span>&nbsp;<%=rs("ForeignID")%>
            </td>
            <%End if%>
            </TD>
            <td align="right" colspan="2" class="field">
                Classification:&nbsp;
                <%=GetClassification()%>
                <input type="hidden" name="hdClass" value="<%=iClass%>" />
                <input type="hidden" name="iClassn" value="<%=iClassn%>" />
				<input type="hidden" name="opfepcc" id="opfepcc" value="<%=opfepcc%>"/>
                <input type="hidden" name="opfonm" id="opfonm" value="<%=opfonm%>"/>
                <input type="hidden" name="isOPFval" id="isOPFval" value="<%=isOPF(SQMappingID)%>"/>
                &nbsp;&nbsp;&nbsp;
            </td>
        </tr>
        <%else%>
        <tr>
            <td align="left" colspan="4" class="field">
                Report Type:&nbsp;
            </td>
            <td align="right" colspan="1" class="field">
                Service Quality&nbsp;
                <input <%If bSQ then Response.Write "checked "%> name="optSQ" type="checkbox" onclick='return optHSESQ_onchange(1)'>&nbsp;&nbsp;&nbsp;
            </td>
            <td align="right" colspan="1" class="field">
                HSE:&nbsp;
                <input <%If bHSE then Response.Write "checked "%> name="optHSE" type="checkbox" onclick='return optHSESQ_onchange(1)'>&nbsp;&nbsp;&nbsp;
            </td>
        </tr>
        <tr>
            <td align="left" colspan="4">
                <span class="field">&nbsp;
            </td>
            <td align="right" colspan="2" class="field">
                Classification:&nbsp;
                <%=GetClassification()%>
                <input type="hidden" name="hdClass" value="<%=iClass%>" />
                &nbsp;&nbsp;&nbsp;
            </td>
        </tr>
        <%End If%>
    </table>
    <table border="1" cellpadding="1" cellspacing="1" width="100%">
        <tr>
            <td>
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="top">
                            <input type="hidden" name="txtCreateDate">
                            <%
					BuildLocInfo lOrgNo,lBSegmentID,lDepartmentID,cn	
                            %>
                        </td>
                        <td valign="top">
                            <table border="0" cellpadding="1" cellspacing="1">
                                <tr>
                                    <!--if lPl =7 or lPl =125 or lPl =129 or lPl=130 then %>
                                    <td class="field">
                                        <b>Integrated Performance Management (IPM) Related?<getHelpLink("Integrated Performance Management (IPM) Related")%></b>
                                    </td>
									
                                    <td id="styleSmall">
                                        <input name="projectNO" type="checkbox" id='proNO' onclick="Checkdataforfilter()"
                                            <projectNOVal%>>&nbsp;No
                                    </td>
                                    <else%-->
                                    <td class="field">
                                        <b>
                                         <!--   <%=IPMText(1)%>
                                            Related?<%=getHelpLink(IPMText(1)& " Related")%>-->
                                    
									Integrated Performance Management (IPM) Related?<%=getHelpLink("Integrated Performance Management (IPM) Related")%>
									</b>
									</td>
                                    <td id="styleSmall">
                                        <input name="projectNO" type="checkbox" id='proNO' onclick="Checkdataforfilter()"
                                            <%=projectNOVal%>>&nbsp;No
                                        <% if projectunknown<>"" then %>
                                        <input name="projectunknow" type="checkbox" id='prounknow' onclick="CheckdataforUnkownfilter()"
                                            <%=projectunknown%>>&nbsp;<b>YES - (Segment unknown)</b><%=getHelpLink("Segment unknown")%>
                                        <%End If%>
                                    </td>
									
                                    <!--End If%-->
                                </tr>
                               <tr>
                                    <td id="styleSmall">
                                        <input name="projectIDS" type="checkbox" id='proIDS' onclick="Checkdataforfilterbyseg()"
                                            <%=projectIDSVal%>>&nbsp;IGWC Integrated Well Construction
                                    </td>
                                    <!-- td id="styleSmall">
                                        <input name="projectIPS" type="checkbox" id='proIPS' onclick="Checkdataforfilterbyseg()"
                                            <%=projectIPSVal%>>&nbsp;Integrated Production Services
                                    </td>-->
									 <td id="styleSmall">                                       
                                        <input name="projectIFS" type="checkbox" id='proIFS' onclick="Checkdataforfilterbyseg()"
                                            <%=projectIFSVal%>>&nbsp;IFS Integrated Fracturing Services
                                    </td>
                                </tr>
                                <tr>
                                    <!--<td id="styleSmall">
                                        <input name="projectISM" type="checkbox" id='proISM' onclick="Checkdataforfilterbyseg()"
                                            <%=projectISMVal%>>&nbsp;IRP Integrated Reservoir Performance [PRSS]
                                    </td>-->
                                     <td id="styleSmall">
											 <input name="projectSPM" type="checkbox" id='proSPM' onclick="Checkdataforfilterbyseg()"
                                            <%=projectSPMVal%>>&nbsp;APS Asset Performance Solutions
                                    </td>
                                </tr>
                                <%if isPTEC > 0 then %>
                                <tr>
                        </td>
                        <td class="field">
                            <b>
                                <%=PTEC(1)%>?<%=getHelpLink(PTEC(1))%></b>
                        </td>
                        <td id="styleSmall">
                            <input type="radio" name="optPTECInv" value="1" <%=IIF(PTECInv = 1,"checked","")%>>Yes
                            <input type="radio" name="optPTECInv" value="0" <%=IIF(PTECInv = 0,"checked","")%>>No
                            <%=mSymbol%>
                        </td>
                    </tr>
                    <%End if%>
                    <% If bSQ then	'***** (MS HIDDEN) - Commented complete If loop section  ***** 
                    %>
                    <tr>
                        <td class="field">
                            <b>Other Segments/Functions/Orgs Involved?&nbsp;</b>&nbsp;
                        </td>
                        <td id="TD2">
                            <input type="radio" name="optSegInv" id="optInv" value="1" <%=IIF(SegInv = "True","checked","")%>>Yes
                            <input type="radio" name="optSegInv" id="optInvno" <%=optInvnoVal%> onclick="CheckdataforfilterByInvoledseg()"
                                value="0" <%=IIF(SegInv = "False","checked","")%>>No &nbsp;<a href="JavaScript:SFOInvHelp()"
                                    class='plain'><img src='../images/movie1.gif' border="0" vspace="0" hspace="0" height="14"
                                        width="14" align='absmiddle'></a>
                        </td>
                    </tr>
                    <% end if %>
                    <!-- (#2588673)
					<% If bSQ then	'***** (MS HIDDEN) - Uncommented complete If loop section  ***** 
					%>
					  <TR>
							<TD class=field ><b>Global Regulatory Compliance Involved?&nbsp;<%=getHelpLink("Global Regulatory Compliance Involved")%></b>&nbsp;</TD>
							<TD  id=TD2>
								<INPUT type=radio name=optTCCInv value=1 <%=IIF(TCCInv = True,"checked","")%>>Yes
								<INPUT type=radio name=optTCCInv value=0 <%=IIF(TCCInv = False,"checked","")%>>No
							</TD>
						</TR>
					<% end if %>
					(#2588673) -->
                    <tr>
                        <td class="field">
                            CRM Client:&nbsp;
                        </td>
                        <td nowrap>
                            <%=getCRMCLient(CRMClient)%>
                        </td>
                        <tr>
                            <tr>
                                <td class="field">
                                    Supplier:&nbsp;
                                </td>
                                <%
						dim CID,Squery,innerRS 
						CID = ""
						Squery="Select * from tblRIRContractors  With (NOLOCK) "		            
						Squery=Squery& " WHERE QPID=" & iQPID & " "
						Squery=Squery& " Order by SeqID"
						'innerRS.Open sTemp, cn
						set innerRS = cn.execute(Squery)
						If NOT innerRS.EOF or NOT innerRS.BOF Then
							While Not innerRS.EOF
								CID = CID & innerRS("ContractorID")&":"&innerRS("ServiceID")&":"&innerRS("RiskRate")&":"&innerRS("Mode") & ","
								innerRS.MoveNext
							WEND                    						
                                %>
                                <td nowrap>
                                    <%=getSupplier(CID)%>
                                </td>
                                <%Else%>
                                <td nowrap>
                                    <%=getSupplier(0)%>
                                </td>
                                <%
						End If		                
                                %>
                                <tr>
                                    <%If Client>0 Then%>
                                    <tr>
                                        <td class="field">
                                            QUEST Client:&nbsp;
                                        </td>
                                        <td nowrap>
                                            <%=getQUESTCLient(Client)%>
                                        </td>
                                    </tr>
                                    <%End IF%>
                                    <%					
									If bSQ and (isREWSQMapping(SQMappingID) or isWTSSQMapping(SQMappingID) or isOSTPDPMapping(SQMappingID) or isSPWL(iBSID)) Then%>
                                    <!--isREW(lPL)---isWTS(lPL)-->
									
                                    <tr>
                                        <td class="field">
                                            Accounting Unit:&nbsp;
                                        </td>
                                        <td>
										
										<% if (SQMappingID=2 or SQMappingID=3) and ActLegacyChkFlg<>"True" then %>
											<%if instr(LOCCountry,"Belarus")>0 or instr(LOCCountry,"Russian")>0 then%>

												<select name="txtAccountUnit" onchange="return AccountingUnit_onchange()" id="Select2">
													<option value=''>(No Account Unit)
														<option value=''>(SEARCH MORE ACCOUNTING UNITS)
													<%=getAccUnits(lOrgNo,AccUnit,cn,lPL)%>											
												</select>
											<%else%>
												<select name="txtAccountUnit" onchange="return AccountingUnitMaximo_onchange()" id="Select2">
													<option value=''>(No Account Unit)
														<option value=''>(SEARCH MORE ACCOUNTING UNITS)
													<%=getMaxAccUnits(lOrgNo,AccUnit,cn,lPL)%>											
												</select>
											<%end if%>
										<% else %>
										    <select name="txtAccountUnit" onchange="return AccountingUnit_onchange()" id="Select2">
                                                <option value=''>(No Account Unit)
                                                    <option value=''>(SEARCH MORE ACCOUNTING UNITS)
												<%=getAccUnits(lOrgNo,AccUnit,cn,lPL)%>											
                                            </select>					
										
										<%end if %>
											
                                        </td>
                                    </tr>
									
									
									
                                    <%End IF%>
                </table>
            </td>
            <td valign="top">
				<table>
					<tr>
						<td id="styleSmall"> <%if DateValue(dtRptDate) >= DateValue(fncD_Configuration("GRCTCC")) then%>
								<% if bSQ or (bSQ and bHSE) then %>
								  <b>GRC or TCC Involved?</b></br>									
									<input type="radio" name="optgot" value="1" onclick="checkdataforGRC()"<%=IIF(Grctcc = "1","checked","")%>>Yes
									<%if (TCCInv=True) then%>
									<input type="radio" name="optgot" value="0" <%=IIF(Grctcc = "0","checked","")%> disabled = "TRUE">No &nbsp;
									<%Else%>
									<input type="radio" name="optgot" value="0" <%=IIF(Grctcc = "0","checked","")%>>No &nbsp;
									<%End If%>
									<%=getHelpLink("GRC/TCC Involved")%>
                        
								
								<%End If%>
							<%End If%>
						</td>
					</tr>
					<tr>
						<td>
							<div align="center">
							</br></br>
								<span class="okNote"><b>Remedial Actions<br></span><span id="Span1" class="urgent"><b>
								<%=iOpenActions%>- Pending
								<br></span><span id="Span2" class="okNote"><b>
								<%=iClosedActions%>- Closed</span></b>
							</div>
						</td>
					</tr>
			</table>
                
            </td>
        </tr>
    </table>
    </TD></TR></TABLE>
    <div id='divHSEBLOCK'>
        <table border="1" cellpadding="1" cellspacing="1" width="100%">
            <tr>
                <td class="field" align="left">
                    <table>
                        <tr>
                            <td class="field">
                                Event Date:
                            </td>
                            <td>
                                <!--<INPUT type=text name=txtEvDate size="12" onfocus="inputDate()" language=javascript value="<%=EventDate%>">-->
                                <input type="text" name="txtEvDate" size="12" value="<%=EventDate%>">
                                <%=popupCalendar("frmRIR.txtEvDate")%><%=mSymbol%>
                            </td>
                            <td class="field">
                                Event Time:
                            </td>
                            <td>
                                <!--<INPUT type=text name=txtEvTime size=5 onfocus="inputTime()" language=javascript value=<%=EventTime%>><%=mSymbol%>-->
                                <input id="txtEvTime" type="text" name="txtEvTime" size="5" value="<%=EventTime%>">&nbsp;
                                <%=popupTime("frmRIR.txtEvTime",1)%><%=mSymbol%>
                            </td>
                            <%If (bSQ and (not isEMS(SQMappingID) or JobID<>"")) then  %>
                            <!--SWIFT #2448303 - Develop EMS SQ Tab -->
                            <td class="field">
                                <%=JobLbl%>:&nbsp;
                            </td>
                            <td nowrap colspan="3">
                                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                    <tr>
                                        <!--Added By Deepak For D&M Tab Development-->
                                        <%If not isDMSeg(iBSID) Then%>
                                        <%If ( lPL = 120 AND ShowPF()) Then%><!--PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.-->
                                        <%if source = 14 or source = 25 then%>
                                        <td>
													<input name="txtJobID" <% if not IsOwner() then%>readonly<%end if%> style="font-size: 9pt" type="text" size="20" maxlength="20"
														value="<%=JobID%>">
                                        </td>
                                        <%else%>
                                        <td>
                                            <input name="txtJobID" style="font-size: 9pt" type="text" size="20" maxlength="20"
                                                value="<%=JobID%>">
                                        </td>
                                        <%end if%>
                                        <%Else%>
                                        <td>
                                            <input name="txtJobID" style="font-size: 9pt" type="text" size="20" maxlength="20"
                                                value="<%=JobID%>">
                                            &nbsp;
                                            <input type="Hidden" name="HdnIlluminaJobAID" value="<%=JobID%>">
                                            <input type="Hidden" name="ALSDelete" value="">
                                            <%if isALSSeg(iBSID) Then%>
                                            <%=getHelpLink("Illumina Job AID")%>
                                            <%END IF%>
                                        </td>
                                        <%End If%>
                                        <%Else%>
                                        <%If (not ShowDM()) Then%>
                                        <td>
                                            <input name="txtJobID" style="font-size: 9pt" type="text" size="20" maxlength="20"
                                                value="<%=JobID%>">
                                        </td>
                                        <%Else%>
                                        <%if source = 14 or source = 25 then%>
                                        <td>
                                            <input name="txtJobID" readonly style="font-size: 9pt" type="text" size="20" maxlength="20"
                                                value="<%=JobID%>">
                                        </td>
                                        <%else%>
                                        <td>
                                            <input name="txtJobID" style="font-size: 9pt" type="text" size="20" maxlength="20"
                                                value="<%=JobID%>">
                                        </td>
                                        <%end if%>
                                        <%End If%>
                                        <%End If%>
                                    </tr>
                                </table>
                            </td>
                            <% End if%>
                        </tr>
                        <tr>
                            <td colspan="1" class="field">
                                Site:
                            </td>
                            <td colspan="3">
                                <%=getSiteTypes(Loctn)%><%=mSymbol%>
                            </td>
                            <td class="field">
                                CRM Rig Name:&nbsp;
                            </td>
                            <td nowrap>
                                <%=getCRMRigs(CRMRigID)%>
                                <span id="req_RigName" style='display: none;'>
                                    <%=mSymbol%></span> <a href="JavaScript:CRMRigHelp()" class='plain'>
                                        <img src='../images/movie1.gif' border="0" vspace="0" hspace="0" height="14" width="14"
                                            align='absmiddle'>
                                    </a>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="1" class="field">
                                Site Name:
                            </td>
                            <td colspan="3">
                                <input name="txtLoc" value="<%=DisplayQuotes(SiteName)%>">
                                <!--<span id="req_SiteName" style='display: none;'>-->
								<%
								'=mSymbol
								%>
								<!--</span>-->
                            </td>
							<%If isROP > 0 and bSQ Then%>
							<td class="field">
								Remote Operations Involved:
							</td>
							<td id="styleSmall">
								<input type="radio" name="optROPInv" id="optROPInvY" value="1" <%=IIF(ROPInv = 1,"checked","")%>>Yes
								<input type="radio" name="optROPInv" id="optROPInvN" value="0" <%=IIF((ROPInv = "") or (ROPInv = 0),"checked","")%>>No
								<%=getHelpLink("Remote Operations Involved")%>
							</td>
							<%End IF%>			
                            <%If (bSQ and (HideSQ=0 and HideWSSQ=0)) and (not isEMS(SQMappingID) or WellSite>0) Then %>
                            <!--SWIFT #2448303 - Develop EMS SQ Tab -->
                            <td class="field">
                                Well Site:&nbsp;
                            </td>
                            <td>
                                <!--div id = 'wellsitename'></div-->
                                <% if WellSite=1 then %>
                                <div id='wellsitename'>
                                    Offshore
                                </div>
                                <% elseif WellSite=2 then%>
                                <div id='wellsitename'>
                                    OnShore
                                </div>
                                <% Else %>
                                <div id='wellsitename'>
                                    N/A
                                </div>
                                <% End if %>
                            </td>
                            <% Else %>
                            <div id='wellsitename' style="display: none;">
                            </div>
                            <%End if%>
                        </tr>
                    </table>
                </td>
                <td align="left">
                    <table>
                        <tr class="reportheading">
                            <td colspan="2" align="center">
                                Risk Classification
                            </td>
                        </tr>
                        <tr>
                            <td class="field">
                                Potential:&nbsp;&nbsp;
                            </td>
                            <td class="data">
                                <%=getRskName(PRisk_C)%>
                            </td>
							<td class="field">
								<%=FailStatViewMain(rs("QID"),"P")%>
							</td>
                        </tr>
                        <tr>
                            <td class="field">
                                Residual:&nbsp;&nbsp;
                            </td>
                            <td class="data">
                                <%=getRskName(RRisk_C)%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    <!--Saxon Changes---->
    <% if EnableOperation>0 Then %>
    <div id='divSaxonBlock'>
        <table border="1" width="100%">
            <tr>
                <td align="left" colspan="3">
                    <table cellpadding="0" cellspacing="0" width="100%" border="0">
                        <tr>
                            <td width="30%" valign="bottom" rowspan="2" class="field">
                                <b>Operation at Time of Event:</b>
                            </td>
                            <td width="10%" align="right" class="field">
                                Category:&nbsp;&nbsp;&nbsp;
                            </td>
                            <td width="25%" valign="left">
                                <select name="cat_sq" id="Select3" onchange="DT_SQOperation.setSL(this,frmRIR.subcat_sq,2)">
                                    <option value=" ">(Select Category)
                                        <%=GetSQOperationsCat(OperationCat)%>
                                </select><%=mSymbol%>
                            </td>
                            <td width="10%" align="right" class="field">
                                Subcategory:&nbsp;&nbsp;&nbsp;
                            </td>
                            <td width="25%" align="left">
                                <select name="subcat_sq" id="Select4">
                                    <option value=" ">(Select a Subcategory)
                                        <%IF (OperationSubCat <> " " and OperationSubCat <> "0") Then%>
                                        <%=GetSQOperationsSubCatCat(OperationCat,OperationSubCat)%>
                                        <%END IF %>
                                </select><%=mSymbol%>
                            </td>
                        </tr>
                        <!--tr>
                    <td><a href="http://calsqldev02/Reports/Pages/Report.aspx?ItemPath=%2fReport+Project2%2fReport1" alt="hello">hello</a>
                    </td>
                    </tr-->
                    </table>
                </td>
            </tr>
        </table>
    </div>
    <%End If%>
    <!--Till Here-->
    <table border="1" cellpadding="1" cellspacing="1" width="100%">
        <tr class="reportheading">
            <td colspan="2" align="center">
                Description and Details of Actual or Potential Loss
            </td>
        </tr>
        <tr>
            <td class="field" align="right">
                Brief Description
            </td>
            <td>
                <input name="txtShortDesc" maxlength="50" size="68" value="<%=DisplayQuotes(ShortDescription)%>"><%=mSymbol%>
            </td>
        </tr>
        <tr>
            <td class="field" align="right" valign="top">
                Detailed Description
            </td>
            <td valign="top">
                <%'@@Sreedhar Feb 19,2009 - Removed Safedisplay due to issues with Other Languages Like Russiachinese etc 
                %>
                <table cellpadding="0" cellspacing="0">
                    <tr>
                        <td>
                            <textarea name="txtFullDesc" maxlength="4000s" wrap="virtual" cols="70" rows="7"
                                onkeyup='return validlength(document.frmRIR.txtFullDesc, msgtxtFullDesc,4000)'><%=LongDescription%></textarea>
                        </td>
                        <td valign="bottom" align="left">
                            <table>
                                <tr>
                                    <td valign="bottom">
                                        <input type="text" style="border: none; background: transparent" id="msgtxtFullDesc"
                                            disabled size="7"><br>
                                        <%=mSymbol%><font size="1" color="red"><i>(Minimum atleast 50 chars)</i></font>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <%

	If bHSE Then DisplayHSE() 
	If bSQ Then DisplaySQ () 
	if HSESQMEssage<> "" Then alert HSESQMEssage

    %>
    <%If Not bNR Then 
		RS.Close
		Set RS = Nothing
		Set cn = Nothing
	End If
	debugPrint("Process End" & Now())
	Response.write DebugMsg
    %>
    <table border="0" cellpadding="1" cellspacing="1" width="100%">
        <tr>
            <td align="right" valign="top">
                <input name="cmdSubmit" onclick="verifydata()" type="button" value="Save Data">
                <input type="hidden" name="sServerVariables" value="<%=sServerVariables%>">
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>
    </form>

    <script language="javascript">
	    initForm();
	    <%if  spsl2 <>"" and spsl2 <>0 and bSQ = true then %>
						SPS_SelAssign(<%=spsl2%>, <%=spsl3%>, '<%=spsl4_new%>', <%=spsb2 %>);
		<%elseif bSQ = true then %>	
						SPS_SelAssign(<%=SQ_Process%>, <%=SQ_MetroStop%>, '<%=SQ_Activity%>', <%=SQ_ProcessOwn%>);		
	    <%End if%>		
            
        <%If bSQ then %>
            if ((document.getElementById("SQB2_0").value == "") || (document.getElementById("SQB2_0").value == "0")) {
                document.getElementById("DescpSpan1").style.display = "none";
                document.getElementById("DescpSpan").style.display = "none";
            }
	    <% End if%>

    </script>

</body>
</html>
<%
	Sub DisplayHSE
	On Error Resume Next
%>
<input type="hidden" name="HSEDisplayed" value="1">
<div id='HSEDiv'>
    <table border="1" cellpadding="0" cellspacing="1" width="100%">
        <tr class="reportheading">
            <td valign="Middle" align="center" colspan="3">
                <b>Health, Safety and Environment Data</b>
            </td>
        </tr>
        <tr>
            <td style="font-size: 9pt" valign="Middle" align="center" colspan="3">
                <%If iClass = 1 Then%>
                <a href="<%=sHSESeverityMatrix%>" title="Display HSE Severity Matrix" onclick="showHSEMatrix();;return false">
                    <b>HSE Severity</b></a>&nbsp;
                <%=GetSeverities(cn, "cmbHSESeverity", iHSESev)%><%=mSymbol%>
                <%End If%>
                <a href="javascript:void(0)" title="Display Hazard Category" onclick="showHazardCategory();;return false">
                    <b>Hazard Category</b></a>&nbsp;
                <select name="txtHazard" id="txtHazard" style="font-size: 9pt" onchange='Haz_onchange(this)'>
                    <%
					if DateValue(dtRptDate) >= DateValue(fncD_Configuration("HAZCAT")) then
					If HazardCat=0 Then Response.Write "<option 'selected' value='0'>(Selection Required)"
					Response.Write GetHazCatOptionsnew(HazardCat)
					ELSE
					If HazardCat=0 Then Response.Write "<option 'selected' value='0'>(Selection Required)"
					Response.Write GetHazCatOptions(HazardCat)
					end if
                    %>
                </select><%=mSymbol%>
                <%
					If SLBInv Then
						If IndRec Then SLBInvment = 1 Else SLBInvment = 2
					Else
						If SLBCon Then SLBInvment = 3 Else SLBInvment = 4
					End If		
					
                %>
            </td>
        </tr>
        <tr>
            <td style="font-size: 9pt" valign="Middle" align="center" colspan="3">
                <table cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td style="font-size: 9pt">
                            &nbsp;<b>Activity Type</b>
                        </td>
                        <td style="font-size: 9pt" align="left">
                            &nbsp;&nbsp;<b>Event</b>&nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt">
                            <input type="radio" name="optSLBInvment" style="font-size: 9pt" onchange="SetEventCategorisation();<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>SetEventSubSafety();<%End IF%>"
                                value="1" <%=IIF(SLBInvment = 1,"checked","")%> />Work Related Activities
                        </td>
                        <td style="font-size: 9pt" align="left">
                            &nbsp;&nbsp;SLB Involved/Industry Recognized&nbsp;<%=getHelpLink("Industry Recognized")%>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt">
                            <input type="radio" name="optSLBInvment" style="font-size: 9pt" onchange="SetEventCategorisation();<%If Cdate(FmtDateTime(dtRptDate)) > Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>SetEventSubSafety();<%End IF%>"
                                value="2" <%=IIF(SLBInvment = 2,"checked","")%> />SLB Related Activities
                        </td>
                        <td style="font-size: 9pt" align="left">
                            &nbsp;&nbsp;SLB Involved/Non Industry Recognized&nbsp;<%=getHelpLink("SLB Involved")%>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt">
                            <input type="radio" name="optSLBInvment" style="font-size: 9pt" <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>onchange="SetEventCategorisation();<%End IF%>"
                                value="3" <%=IIF(SLBInvment = 3,"checked","")%> />Non-SLB related activities
                            concerning employees, dependents, clients and contractors
                        </td>
                        <td style="font-size: 9pt" align="left">
                            &nbsp;&nbsp;SLB Non Involved/Advisory&nbsp;<%=getHelpLink("SLB Non Involved/Advisory")%>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt">
                            <input type="radio" name="optSLBInvment" style="font-size: 9pt" <%If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then%>onchange="SetEventCategorisation();<%End IF%>"
                                value="4" <%=IIF(SLBInvment = 4,"checked","")%> />Non-SLB related activities
                            concerning all other third parties
                        </td>
                        <td style="font-size: 9pt" align="left">
                            &nbsp;&nbsp;SLB Non Involved/Informative&nbsp;<%=getHelpLink("SLB Non Involved/Informative")%>
                        </td>
                        <input type="hidden" name="ProtectDoc" value=''>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt" colspan="2">
                            <hr>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-size: 9pt">
                            <b>Regulatory Recordable<b>
                        </td>
                        <td style="font-size: 9pt">
                            <input type="radio" name="RegRec" value="1" <%=IIF(RegRec = True,"checked","")%>
                                onclick="javascript:chkRegRec(this)">Yes &nbsp;&nbsp;<input type="radio" name="RegRec"
                                    value="0" <%=IIF(RegRec = False,"checked","")%>>No
                        </td>
                    </tr>
                </table>
                <!--
				<INPUT type=checkbox name=RegRec value=1 <%=IIF(RegRec = True,"checked","")%>>				
				Regulatory Recordable
				-->
            </td>
        </tr>
        <tr>
            <%ShowCat Application("a_LossCategories"),"hse"%>
        </tr>
    </table>
</div>
<%

			if  LockCount = 0 and iClass=1 and iHSESev <> 1 and bHSE and  not chkHseLockingMgmt()  Then

						Response.Write "<script type='text/javascript'>document.getElementById('divHSEBLOCK').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('HSEDiv').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK1').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK2').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('Warning21').innerHTML = '*** Key HSE Fields in this Report are now Locked ***';</script>"
						else
						Response.Write "<script type='text/javascript'>document.getElementById('divHSEBLOCK').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('HSEDiv').style.display = '';</script>"
						Response.Write "<script type='text/javascript'> if (document.getElementById('HSEBLOCK1') != null) { document.getElementById('HSEBLOCK1').style.display = ''; }</script>"
						Response.Write "<script type='text/javascript'> if (document.getElementById('HSEBLOCK2') != null) { document.getElementById('HSEBLOCK2').style.display = ''; }</script>"
						If bNR = false  Then
						Response.Write "<script type='text/javascript'> if (document.getElementById('Warning21') != null) { document.getElementById('Warning21').innerHTML = '*** Warning this report will be locked from editing Key HSE Fields in  " & LockCount &" day(s) ***';}</script>"
						End if 

			end if
      If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSubFun Display HSE QID="&SafeNum(iQPID)
		End If

	End Sub


	Sub DisplaySQ()
	On Error Resume Next
		If SLBRel Then
			If External Then SQInvment = 2 Else SQInvment = 1
		Else
			SQInvment = 3
		End If	

        If Trim(BusinessSegment)="" Or Trim(BusinessSegment)="0" Then BusinessSegment= GetSubBusinessSegID(lOrgNo)		
%>
<input type="hidden" name="SQDisplayed" value="1">
<input type="hidden" name="isTimeLossEntered" value="<%=isTimeLossEntered%>">
<input type="hidden" name="hdTLMatrix" value="<%=strTLMatrix%>">
<table border="1" cellpadding="0" cellspacing="0" width="100%">
    <tr id="divSQBLOCK1" bgcolor="lightgrey">
        <td valign="Middle" align="center" colspan="3">
            <b>Service Quality Data</b>
        </td>
    </tr>
    <tr id="divSQBLOCK2">
        <td style="font-size: 9pt" valign="Middle" align="center" colspan="3">
            <%If iClass = 1 Then%>
            <a href="<%=sSQSeverityMatrix%>" title="Display SQ Severity Matrix" onclick="showSQMatrix();;return false">
                <b>SQ Severity</b></a>&nbsp;
            <%=GetSeverities(cn, "cmbSQSeverity", iSQSev)%><%=mSymbol%>
            <%End If%>
        </td>
    </tr>
    <%If bSQ and isSPS>0 then%>
    <tr nowrap style="width: 345px; height: 70px;">
        <td style="font-size: 7pt; margin-top: 3px; border: 0;" bgcolor="#FFFFCC" width="500px"
            valign="middle" align="left" colspan="5">
            <table border="0">
                <tr>
                    <td nowrap style="font-size: 12 pt;" colspan="3">
                        Select the <b>Process Owner</b> or <b> Process </b>to filter the <b>Metro Stop,</b> to identify the <b>Activity</b> being performed when the event occurred.  <%=getHelpLink("Process")%>
                    </td>
                </tr>
                <tr style="font-size: 7 pt;" nowrap>
                    <td nowrap style="font-size: 7 pt;">
                       <!-- <a href="<%=BusWrkFLowHubLink%>" title="Function" onclick="showBusWorkFlow();;return false">-->
                            <b>Process Owner:</b>
                        <%=getSelBox("B2",0,0,0)%>
                    </td>
                    <td nowrap>
                        <span id='SpanProcess' style="font-size: 7 pt;"><b>Process: </b>
                           <% 
						  
						   'response.end
						   
						   if iQPID<>"0" and (SPS_L2=0 or SPS_L2="")  then 
						   
						   %>
						    <%=getSelBox("L2",0,0.8,0)%>
						   <%else%>
						   <%=getSelBox("L2",0,0,0)%>
						   <%end if%><%=mSymbol%></span>
                    </td>
                    <td nowrap>
                        <span id='Pro' style="font-size: 7 pt; display: none;"><b>Metro Stop: </b>
						 <% if iQPID<>"0" and (SPS_L3=0 or SPS_L3="")  then %>
						  <%=getSelBox("L3",0,0.8,0)%>
						 <%else%>
                            <%=getSelBox("L3",0,0,0)%>
                           
							 <%end if%> <%=mSymbol%></span>
                    </td>
                    <td nowrap>
                        <span id='Act' style="font-size: 7 pt; display: none;"><b>Activity: </b>
						 <% if iQPID<>"0" and (SPS_L4=0 or SPS_L4="")  then %>
						  <%=getSelBox("L4",0,0.8,0)%>
						 <%else%>
                            <%=getSelBox("L4",0,0,0)%>
						<%end if%>
						<%=mSymbol%></span>
                    </td>
					
					<tr><td><a href="https://spx.slb.com/MetroMap/Details?Type=KeyProcessAndFunction" title="SPX Explorer" onclick="showSPXexplorer();;return false">
						<b>SPX Explorer</b></a>&nbsp; <%=getHelpLink("Click SPX Explorer")%> </td></tr>
					
					 <td>
                        <div id='DescpSpan1' style="font-size: 8pt; margin-top: 3px; width: 40px; height: 10px;
                            margin-right: 40px; <%=IniStyle%>">
                            <b>Description:</b>
                        </div>
                    </td>
                    <td>
                        <div id='DescpSpan' style="font-size: 7pt; margin-top: 3px; width: 345px; height: 60px;
                            overflow-y: scroll; <%=IniStyle%>">
                            <%=getDescription(spsl4_new)%>
                        </div>
                    </td>
                </tr>
            </table>
        </td>

    </tr>
    <input type="hidden" name="crntdescp" value="<%= iif(spsl4_new="",0,spsl4_new)%>">
    <input type="hidden" name="cntL2" value="1">
    <%End If%>
    <tr>
        <td valign="top">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <%'swi changes
					If ((iClass = 1) or (iClass = 2)) and (isSWI>0) Then
                %>
                <tr width="100%">
                    <td colspan="2">
                        <hr />
                    </td>
                </tr>
                <tr>
                    <td style="font-size: 9pt">
                        <b>Was the activity of this event covered by an
                            <br />
                            official 'Do It Right' SWI/Checklist/Emergency Checklist?</b>
                    </td>
                    <td nowrap colspan="2">
                        <input type="radio" name="swiqn" value="1" <%=iif(SwiQn1=1," checked","")%> onclick="return chkOptSWIqns(this)">Yes
                        <input type="radio" name="swiqn" value="2" <%=iif(SwiQn1=2," checked","")%> onclick="return chkOptSWIqns(this)">No
                        <%=mSymbol%>
                    </td>
                </tr>
                <%If (iClass = 2) Then
							if SwiQn1 = 1 then
							swistyle = "block"
							else 
							swistyle = "none"
							End if
					
                %>
                <tr>
                    <td style="font-size: 9pt">
                        <div id='Q2A' style="display: <%=swistyle%>">
                            <br />
                            <b>Was the potential loss prevented by the use of the
                                <br />
                                'Do It Right' SWI/Checklist/Emergency Checklist? </b>
                        </div>
                    </td>
                    <td>
                        <div id='Q2Ans' style="display: <%=swistyle%>">
                            <input type="radio" name="swiqntwo" value="1" <%=iif(SwiQn2A=1," checked","")%>>Yes
                            <input type="radio" name="swiqntwo" value="2" <%=iif(SwiQn2A=2," checked","")%>>No
                            <%=mSymbol%></div>
                    </td>
                </tr>
                <%End if%>
                <%If (iClass = 1) Then 
							if SwiQn1 = 1 and ((iSQSev=1) or (iSQSev=2)) then
							swistyle = "block"
							else 
							swistyle = "none"
							End if
                %>
                <tr>
                    <td style="font-size: 9pt">
                        <div id='Q2B' style="display: <%=swistyle%>">
                            <br />
                            <b>Was the loss preventable by the use of the
                                <br />
                                'Do It Right' SWI/Checklist/Emergency Checklist? </b>
                        </div>
                    </td>
                    <td>
                        <div id='Q2Bans' style="display: <%=swistyle%>">
                            <input type="radio" name="swiqntwob" value="1" <%=iif(SwiQn2B=1," checked","")%>>Yes
                            <input type="radio" name="swiqntwob" value="2" <%=iif(SwiQn2B=2," checked","")%>>No
                            <%=mSymbol%></div>
                    </td>
                </tr>
                <tr>
                    <td style="font-size: 9pt">
                        <div id='Q3A' style="display: <%=swistyle%>">
                            <br />
                            <b>Was the 'Do It Right' SWI/Checklist/Emergency actually
                                <br />
                                followed during the activities?</b></div>
                    </td>
                    <td>
                        <div id='Q3Ans' style="display: <%=swistyle%>">
                            <input type="radio" name="swiqnthree" value="1" <%=iif(SwiQn3=1," checked","")%>>Yes
                            <input type="radio" name="swiqnthree" value="2" <%=iif(SwiQn3=2," checked","")%>>No
                            <%=mSymbol%></div>
                    </td>
                </tr>
                <%End if%>
                <%End if%>
                <% ' End of SWI changes
                %>
                <% 
                Dim DtCompliance,Diffcomp
				DtCompliance = fncD_Configuration("ComplianceSQStandards")
				Diffcomp = DateDiff("n",DtCompliance,dtRptDate)		
				IF (Diffcomp < 1)  then %>
                <% If HideWSSQ = 1 AND not isALSSeg(iBSID) then %>
                <tr width="100%">
                    <td colspan="2">
                        <hr />
                    </td>
                </tr>
                <tr>
                    <td style="font-size: 9pt">
                        <b>Compliance with SQ Standards </b>:&nbsp;
                    </td>
                    <td>
                        <input name="optSQStandard" type="radio" style="font-size: 9pt" maxlength="4" value="1"
                            <%=IIF(SQStandard = "1"," checked ","")%>>
                        Yes&nbsp;
                        <input name="optSQStandard" type="radio" style="font-size: 9pt" maxlength="4" value="0"
                            <%=IIF(SQStandard = "0"," checked ","")%>>
                        No
                    </td>
                </tr>
                <% end if %>
                <% if HideSQ = 1 then %>
				<!--PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.-->
                <%If (((not isDMSeg(iBSID) or IsGSSshared(lorgno)) or  not ShowDM() or (SafeNum(SQStandard) = 1)) AND (lPL <> 120 or not ShowPF() or (SafeNum(SQStandard) = 1))) AND not isALSSeg(iBSID) Then%>
                <%If Not(bSQ And IsGSStab(SQMappingID) And ShowGSS) OR SQStandard="1" Then%>
                <tr width="100%">
                    <td colspan="2">
                        <hr />
                    </td>
                </tr>
                <tr>
                    <td style="font-size: 9pt">
                        <b>Compliance with SQ Standards </b>:&nbsp;
                    </td>
                    <td>
                        <input name="optSQStandard" type="radio" style="font-size: 9pt" maxlength="4" value="1"
                            <%=IIF(SQStandard = "1"," checked ","")%>>
                        Yes&nbsp;
                        <input name="optSQStandard" type="radio" style="font-size: 9pt" maxlength="4" value="0"
                            <%=IIF(SQStandard = "0"," checked ","")%>>
                        No
                    </td>
                </tr>
                <%End IF%>
                <%End IF%>
				<!--PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.-->
                <%If (((not isDMSeg(iBSID) or IsGSSshared(lorgno))or not ShowDM() or (SafeNum(SQNRedone) >0)) AND (lPL <> 120 or not ShowPF() or (SafeNum(SQNRedone) >0))) AND not isALSSeg(iBSID) Then%>
                <%If NOT (bSQ And IsGSStab(SQMappingID) And ShowGSS) OR SQNRedone <> "0" Then%>
                <tr style="height: 25px;">
                    <td style="font-size: 9pt">
                        Number of times Operation was ReDone:
                    </td>
                    <td>
                        <input name="txtSQNRedone" style="font-size: 9pt" type="text" size="1" maxlength="4"
                            value="<%=SQNRedone%>">
                    </td>
                </tr>
                <%End If%>
                <%End IF%>
				<!--PLID=120 does not exists in tblproductlines so not need to change it with BSIDs.-->
                <%If (((not isDMSeg(iBSID) or IsGSSshared(lorgno)) or not ShowDM() or (SafeNum(SQPFailure) >0)) AND (lPL <> 120 or not ShowPF() or (SafeNum(SQPFailure) >0))) AND not isALSSeg(iBSID)Then%>
                <%If NOT (bSQ And IsGSStab(SQMappingID) And ShowGSS) OR SQPFailure <> "0" Then%>
                <tr style="height: 25px;">
                    <td style="font-size: 9pt">
                        Percentage target run life at time of failure:
                    </td>
                    <td>
                        <input name="txtPFailure" style="font-size: 9pt" type="text" size="1" value="<%=SQPFailure%>">&nbsp;&nbsp;%
                    </td>
                </tr>
                <%End If%>
                <%End If%>
                <%End IF%>
                <%End IF%>
            </table>
        </td>
        <td style="font-size: 9pt" valign="Top" align="center">
            <table cellpadding="0" cellspacing="0" border="0">
                <tr id="SQDiv1">
                    <td style="font-size: 9pt">
                        &nbsp;<b>Activity/Process/Service</b>
                    </td>
                    <td style="font-size: 9pt">
                        &nbsp;<b>Related</b>&nbsp;
                    </td>
                    <td style="font-size: 9pt">
                        &nbsp;<b>External</b>&nbsp;
                    </td>
                </tr>
                <tr id="SQDiv2">
                    <td style="font-size: 9pt">
                        <input type="radio" name="optSQInvment" style="font-size: 9pt" value="2" <%=IIF(SQInvment = 2,"checked","")%>
                            onclick="return chkOptSLBExternal(this.form)">
                        SLB managed service/product delivered to an external client&nbsp;<%=getHelpLink("Schlumberger External SQ/PQ Events")%>
                    </td>
                    <td style="font-size: 9pt" align="center">
                        Yes
                    </td>
                    <td style="font-size: 9pt" align="center">
                        Yes
                    </td>
                </tr>
                <tr id="SQDiv3">
                    <td style="font-size: 9pt">
                        <% if disableprocess() then %>
                        <input type="radio" name="optSQInvment" style="font-size: 9pt" disabled="false" value="1"
                            <%=IIF(SQInvment = 1,"checked","")%> onclick="return chkOptSLBInternal(this.form)">
                        SLB managed and delivered to an internal customer&nbsp;<%=getHelpLink("Schlumberger Internal SQ/PQ Events")%>
                        <% else %>
                        <input type="radio" name="optSQInvment" style="font-size: 9pt" value="1" <%=IIF(SQInvment = 1,"checked","")%>
                            onclick="return chkOptSLBInternal(this.form)">
                        SLB managed and delivered to an internal customer&nbsp;<%=getHelpLink("Schlumberger Internal SQ/PQ Events")%>
                        <% end if %>
                    </td>
                    <td style="font-size: 9pt" align="center">
                        Yes
                    </td>
                    <td style="font-size: 9pt" align="center">
                        No
                    </td>
                </tr>
                <tr id="SQDiv4">
                    <td style="font-size: 9pt">
                        <input type="radio" name="optSQInvment" style="font-size: 9pt" value="3" <%=IIF(SQInvment = 3,"checked","")%>
                            onclick="return chkOptThirdparty(this.form)">
                        Client or Third Party managed service/product&nbsp;<%=getHelpLink("Schlumberger Non-Related")%>
                    </td>
                    <td style="font-size: 9pt" align="center">
                        No
                    </td>
                    <td style="font-size: 9pt" align="center">
                        -
                    </td>
                </tr>
                <%If iClass = 1 Then%>
                <tr id="SQDiv5">
                    <td style="font-size: 9pt" colspan="3">
                        <hr>
                    </td>
                </tr>
                <tr id="SQDiv6">
                    <td style="font-size: 9pt">
                        <b>Client Affected&nbsp;&nbsp;</b>(Results in loss to Client)
                    </td>
                    <td style="font-size: 9pt" align="center">
                        <input type="radio" name="optCAffect" style="font-size: 9pt" value="1" <%=IIF(ClientAffect,"checked","")%>
                            onclick="return chkClientAffected(this.form)">Yes
                    </td>
                    <td style="font-size: 9pt" align="center">
                        <input type="radio" name="optCAffect" style="font-size: 9pt" value="0" <%=IIF(ClientAffect,"","checked")%>
                            onclick="return chkSlbExtClientAffectedNo(this.form)">No
                    </td>
                </tr>
                <% End IF %>
                <% 
				' code change for SLIM SPS 
				'If (iClass = 1)  or (iClass=2) Then 
                %>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <!--tr>
						<TD style="font-size:9pt" align=left><b>SPS Improvement Domains </b> - SLIM root cause analysis<%=mSymbol%></TD>
					</tr>
					<tr>
						<TD style="font-size:9pt">please identify the root cause(s) of the event &nbsp; <%'getHelpLink("SLIM root cause analysis")%></TD>
					</tr>
					
					<tr>
						<td><%'getHelpLink("Process")%>
						 <INPUT id='styleSmall' name='sProc'  type='checkbox' value='1' <%'iif(SlimProc="True"," checked","")%>> Process</INPUT> 
						</td>
					</tr>
					<tr>
						<td><%'getHelpLink("Control")%>
						 <INPUT id='styleSmall' name='sCtrl'  type='checkbox' value='1' <%'iif(SlimCntrl="True"," checked","")%> > Control</INPUT> 
						</td>
					</tr>
					</tr>
						<td><%'getHelpLink("Competency")%>
						<INPUT id='styleSmall' name='sComp'  type='checkbox' value='1' <%'iif(SlimCompet="True"," checked","")%>> Competency</INPUT> 
						</td>
					</tr>
					</tr>
						<td><%'getHelpLink("Behavior")%>
						<INPUT id='styleSmall' name='sBehav'  type='checkbox' value='1' <%'iif(SlimBehav="True"," checked","")%>> Behavior</INPUT> 
						</td>
					</tr>
					</tr>
						<td><%'getHelpLink("Technology")%>
						 <INPUT id='styleSmall' name='sTech'  type='checkbox' value='1' <%'iif(SlimTech="True"," checked","")%>> Technology</INPUT> 
						</td>
					</tr-->
                <% 'End IF 
                %>
            </table>
        </td>
        <td valign="top">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <%ShowCat Application("a_LossCategories"),"sq"%>
                </tr>
            </table>
        </td>
    </tr>
    <%
		if iClass = 1 and LockCountSQ=0 and bSQ and not chkSQLockingMgmt()  then 
		if iSQSev > 1 then
		response.write ShowNPTSQBLOCKTable
		end if
		else
		if iClass = 1 then
		response.write ShowNPTTable()
		end if
		end if
    %>
    <!--'@Visali 02/15/2006 -->
    <!--'Changed By Deepak For D&M Tab Development 04/26/2006 -->
    <% '	If (((lPL <> 1 or IsGSSshared(lorgno))or SafeNum(DMRecs) > 0 or not ShowDM()) AND (lPL <> 120 or SafeNum(PFRecs) > 0 or not ShowPF())) AND lPL <> 106  AND lPL <> 139 AND lPL <> 132 AND lPL <> 133 AND lPL <> 134 AND lPL <> 135 AND Not isDMSQMapping(SQMappingID or IsGSSshared(lorgno)) Then %>
    <%' 	If (HideSQ = 1 or ShowIPMSQ=1)  Then %>
    <%'If NOT ((bSQ And IsGSStab(lorgno) And ShowGSS) or  isCTSSQMapping(SQMappingID)) OR (SQSPCatID <> "0" OR SQFCatID <> "0" OR SQDCatID <> "0") Then%>
    <%If SQCategoryMappingID > 0 then%>
	<tr>
        <td align="left" colspan="3">
            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                <tr>
                    <td valign="bottom" rowspan="2">
                        <b>Service/Product:</b>
                    </td>
                    <td>
                        <b>Category</b>
                    </td>
                    <td>
                        <b>Sub-Category</b>
                    </td>
                </tr>
                <tr>
                    <td>
                        <select name="txtSPCategory" id="txtSPCategory" onchange="spcategory.setSL(this,txtSubSPCategory)"
                            style="font-size: 9pt">
                            <%
											Set RSSQ = Server.CreateObject("ADODB.Recordset")
											sSQL = "SELECT DISTINCT C.* FROM tlkpSQCategories C INNER JOIN tlkpSQSubCategories SC ON C.ID = SC.PID WHERE  C.PLID IN (0," & tlPL & ") "
											sTemp = sSQL & " AND (status+substatus=0 or C.ID="&SQSPCatID&") AND type='S' ORDER BY C.Description"
											Response.Write "//SPCategory:SQL"& sTemp &vbCRLF
				
											RSSQ.Open sTemp, cn
											If isSPRequired Then sTemp="(Selection Required)" else sTemp="(Selection Not Required)" 
											If SQSPCatID=0 Then Response.Write "<option 'selected' value='0'>"&sTemp
											
				
											Do until RSSQ.EOF %>
                            <option <%If SQSPCatID =(Trim(RSSQ("ID"))) then Response.Write "selected" %> value="<%=RSSQ("ID")%>">
                                <%Response.Write RSSQ("Description")
												RSSQ.MoveNext
											Loop 
											RSSQ.Close
											Set RSSQ = Nothing
                                %>
                        </select><%If isSPRequired Then Response.Write mSymbol %>
                    </td>
                    <td>
                        <select name="txtSubSPCategory" style="font-size: 9pt">
                            <option value='0'>
                                <%=sTemp%>
                                <option>
                                    <option>
                        </select><%If isSPRequired Then Response.Write mSymbol %>
                    </td>
                </tr>
                <tr>
                    <td align="left">
                        <b>SQ Non-conformance:</b>
                    </td>
                    <td>
                        <%'If issaxon(lPL) then
						If issaxon(iSubBSID) then%>
                        <%=getSelBox_new("C",0,0,0)%>
                        <%Else%>
                        <select name="txtFailure" onchange="failure.setSL(this,txtSubFailure);On_Failure()"
                            style="font-size: 9pt">
                            <%			
											Set RSSQ = Server.CreateObject("ADODB.Recordset")
											sTemp = sSQL & " AND (status+substatus=0 or C.ID="&SQFCatID&") AND type='F' ORDER BY Description"
											RSSQ.Open sTemp, cn
											If SQFCatID=0 Then Response.Write "<option selected value='0'>(Selection Required)"
											Do until RSSQ.EOF %>
                            <option <%If SQFCatID =(Trim(RSSQ("ID"))) then Response.Write "selected" %> value="<%=RSSQ("ID")%>">
                                <%Response.Write RSSQ("Description")
												RSSQ.MoveNext
											Loop 
											RSSQ.Close
											Set RSSQ = Nothing
                                %>
                        </select><%=mSymbol%>
                        <%End IF%>
                    </td>
                    <td>
                        <%'If issaxon(lPL) then
						If issaxon(iSubBSID) then%>
                        <%=getSelBox_new("C2",0,0,0)%>
                        <%Else%>
                        <select name="txtSubFailure" style="font-size: 9pt">
                            <option>(Selection Required)
                                <option>
                                    <option>
                        </select><%=mSymbol%>
                        <%End If%>
                    </td>
                </tr>
                <tr <%If issaxon(iSubBSID) then response.write "id=Saxon_Damage_cat style='display:none;'" %>>
                    <td align="left">
                        <b>Damage:</b>
                    </td>
                    <td>
                        <input type="hidden" name="ShowDamage" value='N'>
                        <%'If issaxon(lPL) then
						If issaxon(iSubBSID) then%>
                        <%=getSelBox_new("C3",0,0,0)%>
                        <%Else%>
                        <select name="txtDamage" onchange="damage.setSL(this,txtSubDamage);On_Failure()"
                            style="font-size: 9pt">
                            <%			
											Set RSSQ = Server.CreateObject("ADODB.Recordset")
											sTemp = sSQL & " AND (status+substatus=0 or C.ID="&SQDCatID&") AND type='D' ORDER BY Description"
											Response.Write "//Damage:SQL"& sTemp &vbCRLF
														
											RSSQ.Open sTemp, cn
											
											If isDamageRequired Then sTemp="(Selection Required)" else sTemp="(Selection Not Required)" 
											If SQDCatID=0 Then Response.Write "<option selected value='0'>"&sTemp
				
											Do until RSSQ.EOF %>
                            <option <%If SQDCatID =(Trim(RSSQ("ID"))) then Response.Write "selected" %> value="<%=RSSQ("ID")%>">
                                <%Response.Write RSSQ("Description")
												RSSQ.MoveNext
											Loop 
											RSSQ.Close
											Set RSSQ = Nothing
                                %>
                        </select><%=mSymbol%>
                        <%End If%>
                    </td>
                    <td>
                        <%'If issaxon(lPL) then
						If issaxon(iSubBSID) then%>
                        <%=getSelBox_new("C4",0,0,0)%>
                        <%Else%>
                        <select name="txtSubDamage" style="font-size: 9pt">
                            <option value='0'>
                                <%=sTemp%>
                                <option>
                                    <option>
                        </select><%=mSymbol%>
                        <%End If%>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <input type="hidden" name='txtIsGSSML' value='YES'>
    <%Else%>
    <tr>
        <td>
            <input type="hidden" name='txtIsGSSML' value='NO'>
        </td>
    </tr>
    <%End if%>
    <%' End if 
		'End If%>
    <input type='hidden' name='IsDamageRequired' value='<%=iif(isDamageRequired,1,0)%>'>
    <input type='hidden' name='IsSPRequired' value='<%=iif(IsSPRequired,1,0)%>'>
    <input type='hidden' name='HideSQ' value='<%=HideSQ%>'>
    <input type='hidden' name='HideWSSQ' value='<%=HideWSSQ%>'>
    <input type='hidden' name='IsCauseRequired' value='0'>
	<input type='hidden' id='LegacyChkFlg' name='LegacyChkFlg' value='<%=LegacyChkFlg%>' >
	<input type='hidden' id='chkopfhazornear' name='chkopfhazornear' value='<%=chkopfhazornear%>' >
	<input type='hidden' id='delopfdetail' name='delopfdetail' value='<%=delopfdetail%>' >
	<input type='hidden' id='iclass' name='iclass' value='<%=iclass%>' >	
	<input type='hidden' id='hdnchkinput' name='hdnchkinput' value='' >


</table>
<%
				'
				if  LockCountSQ = 0 and  iClass=1 and bSQ and  not chkSQLockingMgmt()    Then
			
						Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK1').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK2').style.display = 'none';</script>"
						'Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK3').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv1').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv2').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv3').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv4').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv5').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv6').style.display = 'none';</script>"
						'Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK1').style.display = 'none';</script>"
						'Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK2').style.display = 'none';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('Warning22').innerHTML = '*** Key SQ Fields in this Report are now Locked ***';</script>"
			       else

						Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK1').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK2').style.display = '';</script>"
						'Response.Write "<script type='text/javascript'>document.getElementById('divSQBLOCK3').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv1').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv2').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv3').style.display = '';</script>"
						Response.Write "<script type='text/javascript'>document.getElementById('SQDiv4').style.display = '';</script>"
                        if iclass="1" then
						    Response.Write "<script type='text/javascript'>document.getElementById('SQDiv5').style.display = '';</script>"
						    Response.Write "<script type='text/javascript'>document.getElementById('SQDiv6').style.display = '';</script>"
		                End if				
        'Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK1').style.display = '';</script>"
						'Response.Write "<script type='text/javascript'>document.getElementById('HSEBLOCK2').style.display = '';</script>"
						If bNR = false and (CDate(dtRptDatetmp) >= CDate(comparedateEventdate))   Then
						Response.Write "<script type='text/javascript'> if (document.getElementById('Warning22') != null) { document.getElementById('Warning22').innerHTML = ' Warning this report will be locked from editing Key SQ Fields in  " & LockCountSQ &" day(s) ';}</script>"
						End if 
			end if
	
	
	  If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InSub Fun DisplaySQ QID="&SafeNum(iQPID)
		End If
	
	End Sub 

	Function GetClass(sTemp)
	On Error Resume Next
		'changed to return class name (done in conjunction with CSS upgrade) -- GCF Oct 6 2000
		Select Case sTemp
			Case "H"
				GetClass = "class=urgent"
			Case "M"
				GetClass = "class=note"
			Case "L"
				GetClass = "class=em"
			Case Else
				GetClass = ""
		End Select
				  If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun GetClass QID="&SafeNum(iQPID)
		End If
	End Function

	function getRisk(sTemp)
	On Error Resume Next
		If len(trim(sTemp))=0 or isnull(sTemp) then 
			getRisk="Undefined"
		elseif sTemp="H" then 
			getRisk="High"
		elseif sTemp="M" then 
			getRisk="Medium"
		elseif sTemp="L" then 
			getRisk="Low"
		end if
		  If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"Rirdsp2.asp",err.Description  &"-InFun getRisk QID="&SafeNum(iQPID)
		End If
	end function

	Function GetClassification()
		 
		dim rs,conn,str,sql1,sel,str1,iClass1
		On Error Resume Next
		if iClass="" then iClass=0 end if  
		str1 = ""
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("ConnStr")
		str ="<SELECT Name='optClass' onchange='return optHSESQ_onchange(2)'>"
		sql1="select *  from tlkpRIRClass"
		set rs=conn.execute(sql1)	
					  
		While Not RS.EOF 
			iClass1 = rs("Classid")
			
			IF trim(iClass) = trim(iClass1) Then sel = " selected " else sel = ""
			str1 = str1 & "<OPTION value= " & rs("Classid") & " " & sel & " >" & rs("Classdesc")& vbCRLF				                  
			rs.movenext
		Wend   
		
		str = str & str1 & "</select>"
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun  GetClassification QID="&SafeNum(iQPID)
		End If
		GetClassification=str
		set conn=nothing
	
	End Function

	Function GetHiddenBSID(dep)
		Dim sSQL, rss,tmp,sel,tmp1,conn
		On Error Resume Next
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("ConnStr")
		tmp ="<SELECT Name='hidBSID'><Option value='0'></Option>"
		tmp1=""
		sSQL = "SELECT DISTINCT BS.BusinessSegmentID, D.ID, D.Description FROM tblQT_QUESTTree QT,tblProductLines PL,tlkpBusinessSegments BS,tlkpDepartments D WHERE QT.PLID = PL.PLID AND ((PL.EnforceSelection = 1 AND QT.BSID = BS.BusinessSegmentID AND D.Status=0) OR ((D.ID In (Select Department from tblRIRp1 where QID="& iQPID &")) ) OR (PL.EnforceSelection = 0 AND BS.PPLID = PL.PLID)) AND (D.BSID = BS.BusinessSegmentID OR D.BSID = 0) AND QT.ID =" & lOrgNo 
		set rss=conn.execute(sSQL)
		While Not rss.EOF 	
		IF trim(rss("ID")) = trim(dep) Then sel = " selected " else sel = ""
			tmp1 = tmp1 & "<OPTION value= " & rss("BusinessSegmentID") & " " & sel & " >" & rss("ID")& vbCRLF
			rss.movenext
		Wend   
		tmp = tmp & tmp1 & "</select>"
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun GetHiddenBSID QID="&SafeNum(iQPID)
	End If
		GetHiddenBSID=tmp
		set conn=nothing   	
	End Function	
Function GetSeverities(cn,Name, Sev)
		Dim rs, RS1,RS2,RS3, setflag,setInvflag,SQL1, SqSev, HseSev,InvCloseDate,IsInvCloseFlag,RSinvteam,Chkinvteam,ChkinvReviewed
		On Error Resume Next
		setflag = 0
		SqSev = 0
		HseSev = 0
		setInvflag = 1
		Chkinvteam = 0
		ChkinvReviewed = 0
		Dim tmp,msgcol',HSEGuest
		msgcol=""
		If instr(1,Name,"HSE")>0 then msgcol="HSE" else msgcol="SQ"
		'If instr(1,Name,"HSE")>0 and isGuest() then HSEGuest=True else HSEGuest=False
		if NOT IsNumeric(Sev) Then Sev = 0
		Set RS  = cn.Execute("SELECT *, CASE SeverityID WHEN " & Sev & " THEN ' Selected' ELSE '' END AS Selected FROM tlkpRIRSeverity WHERE SeverityID < 5 ORDER BY SeverityID DESC")
		
		Set RS1 = cn.Execute("SELECT Ter_Failed FROM tblRIRWellBarriers WHERE Ter_Failed in (1,0) AND QPID = " & SafeNum(iQPID))	
		If NOT (RS1.EOF OR RS1.BOF) Then setflag = 1
				
		Set RS2 = cn.Execute("SELECT SQSeverity , HSESeverity FROM tblRIRp1 WHERE QID = " & SafeNum(iQPID))
		If NOT (RS2.EOF OR RS2.BOF) Then 
			SqSev   = RS2("SQSeverity") 
			HseSev  = RS2("HSESeverity")	    
		End If
		
		if iQPID<>"0" then
			Set RSinvteam = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT TOP 1 * FROM tblRIRInvTeam a With (NOLOCK) Left JOIN tblririnvdetails b ON a.QPID = b.QPID WHERE a.QPID='" & SafeNum(iQPID) & "'"  'Checking the value of Check box that Investigation is closed or not.			
			RSinvteam.Open SQL1, cn
			If not RSinvteam.EOF then 
				Chkinvteam=1
				ChkinvReviewed=RSinvteam("Reviewed")
				else
				Chkinvteam=0
			end if
			RSinvteam.close
			Set RSinvteam=nothing
		End if
		
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT Top 1 UpDateClosedate,IsInvClosed FROM tblRIR_SLIMRootCause With (NOLOCK) WHERE QPID='" & SafeNum(iQPID) & "'"  'Checking the value of Check box that Investigation is closed or not.
		RS3.Open SQL1, cn
		If not RS3.EOF then 
			InvCloseDate = RS3("UpDateClosedate")
			IsInvCloseFlag = RS3("IsInvClosed") 
			If IsInvCloseFlag = 1 then 
				setInvflag = 1
			Else
				setInvflag = 0
			end if
		end if
		RS3.close
		Set RS3=nothing
		  
		tmp = ""
		tmp = tmp & "<SCRIPT>" & vbCRLF
		tmp = tmp & "	function severityAlert(selectList,msg,setflag,setInvflag,sCur,hCur,Chkinvteam,ChkinvReviewed){" & vbCRLF
		
		tmp = tmp & "		var txt=''; " & vbCRLF
		tmp = tmp & "		var v=selectList.options[selectList.selectedIndex].value;" & vbCRLF
		tmp = tmp & "		var t=selectList.options[selectList.selectedIndex].text;" & vbCRLF
		'Anson Changes***Start
		tmp = tmp & "		if(msg=='HSE'){" & vbCRLF
		tmp = tmp & "		    var H=selectList.options[selectList.selectedIndex].value; " & vbCRLF
		tmp = tmp & "		}" & vbCRLF
		tmp = tmp & "		if(msg=='SQ'){" & vbCRLF
		tmp = tmp & "		    var S=selectList.options[selectList.selectedIndex].value;" & vbCRLF
		tmp = tmp & "		}" & vbCRLF
		
		'Check for severity change in case of Investigation tab.
		tmp = tmp & "		if(setInvflag == 0 && S > 2 && msg=='SQ' && document.frmRIR.LegacyChkFlg.value=='True'){alert('The main severity can not be changed to a severity Major or Catastrophic unless SLIM Investigation is not Closed.');" & vbCRLF 
		tmp = tmp & "		selectList.value=sCur;}" & vbCRLF
		tmp = tmp & "		if(Chkinvteam == 1 && ChkinvReviewed==0 && S < 3 && msg=='SQ' && document.frmRIR.LegacyChkFlg.value=='True' && sCur!=1 && sCur!=2){alert('The main severity can be changed to a severity Light or Serious once SCAT Investigation is completed(Reviewed).');" & vbCRLF 
		tmp = tmp & "		selectList.value=sCur;}" & vbCRLF		
		tmp = tmp & "		if(sCur > hCur){" & vbCRLF
		tmp = tmp & "		if(setflag == 1 && S < sCur && msg=='SQ'){alert('The main severity can not be changed to a lower severity when Well Barrier tab triggers a serious event or a catastrophic event.');" & vbCRLF
		tmp = tmp & "		selectList.value=sCur;}" & vbCRLF		
		tmp = tmp & "}" & vbCRLF
		tmp = tmp & "		else{if(setflag == 1 && H < hCur && msg=='HSE'){alert('The main severity can not be changed to a lower severity when Well Barrier tab triggers a serious event or a catastrophic event.');" & vbCRLF	
		tmp = tmp & "		selectList.value=hCur;}}" & vbCRLF
		'Anson Changes***End	
		Do While NOT rs.EOF
		 if (LegacyChkFlg = "True") Then ' after adding validation for investigation review tab this check was implemented
			If Not IsNULL(rs("HSEConfMSg")) Then
				If trim(rs("HSEConfMSg")) <> "" Then
					tmp = tmp & "		if(H < hCur){if ((msg=='HSE') && (v==" & rs("SeverityID") & ") && setflag==0)"& vbCRLF &" txt = '" & Replace(Replace(rs("HSEConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
					tmp = tmp & "		else{if ((msg=='HSE') && (v==" & rs("SeverityID") & ") && setInvflag == 1)"& vbCRLF &" txt = '" & Replace(Replace(rs("HSEConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
				End if
			end if
			If Not IsNULL(rs("SQConfMSg")) Then
				If trim(rs("SQConfMSg")) <> "" Then
					tmp = tmp & "		if(S < sCur){if ((msg=='SQ')  && (v==" & rs("SeverityID") & ") && setflag==0) txt = '" & Replace(Replace(rs("SQConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
					tmp = tmp & "		else{if ((msg=='SQ') && (v==" & rs("SeverityID") & ") && setInvflag == 1)"& vbCRLF &" txt = '" & Replace(Replace(rs("SQConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
				End if
			End if
		else
		    If Not IsNULL(rs("HSEConfMSg")) Then
				If trim(rs("HSEConfMSg")) <> "" Then
					tmp = tmp & "		if(H < hCur){if ((msg=='HSE') && (v==" & rs("SeverityID") & ") && setflag==0)"& vbCRLF &" txt = '" & Replace(Replace(rs("HSEConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
					tmp = tmp & "		else{if ((msg=='HSE') && (v==" & rs("SeverityID") & "))"& vbCRLF &" txt = '" & Replace(Replace(rs("HSEConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
				End if
			end if
			If Not IsNULL(rs("SQConfMSg")) Then
				If trim(rs("SQConfMSg")) <> "" Then
					tmp = tmp & "		if(S < sCur){if ((msg=='SQ')  && (v==" & rs("SeverityID") & ") && setflag==0) txt = '" & Replace(Replace(rs("SQConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
					tmp = tmp & "		else{if ((msg=='SQ') && (v==" & rs("SeverityID") & "))"& vbCRLF &" txt = '" & Replace(Replace(rs("SQConfMSg"),vbCRLF,"\n"),"'","\'") & "';}" & vbCRLF
				End if
			End if
		End If
			rs.MoveNext
		Loop
		if isGuest() then 
			tmp = tmp & "		if ((v==4) && (msg=='HSE') && setflag==0) {" & vbCRLF
			tmp = tmp & "			alert('Basic Users are restricted from creating Catastrophic Events.\n To create a catastrophic event please contact your Line Manager or QHSE Support.')" & vbCRLF
			tmp = tmp & "			selectList.selectedIndex=selectList.options.length-1;"& vbCRLF
			tmp = tmp & "			txt='';}"& vbCRLF
			
		end if
		tmp = tmp & "		if (txt!='') {" & vbCRLF
		tmp = tmp & "			txt+='\nPressing ""OK"" will display the severity matrix.\n'" & vbCRLF
		tmp = tmp & "			txt+='Pressing ""Cancel"" will undo your severity selection.\n'" & vbCRLF
		'tmp = tmp & "			if(confirm(txt)) showHSEMatrix();" & vbCRLF
		'tmp = tmp & "			else selectList.selectedIndex = selectList.options.length-1;" & vbCRLF
		'tmp = tmp & "		}" & vbCRLF
		tmp = tmp & "			if(confirm(txt)) {" & vbCRLF
		tmp = tmp & "			        if (msg=='HSE') showHSEMatrix(); else showSQMatrix();" & vbCRLF
		tmp = tmp & "			}else selectList.selectedIndex = selectList.options.length-1;" & vbCRLF
		tmp = tmp & "		}" & vbCRLF
		tmp = tmp & "	}" & vbCRLF
		tmp = tmp & "</SCRIPT>" & vbCRLF
		if bSQ and msgcol="SQ" then
			tmp = tmp & "<SELECT style=""font-size:9pt"" name=""" & name & """ onchange=""severityAlert(this,'"&msgcol&"',"&setflag&","&setInvflag&","&SqSev&","&HseSev&","&Chkinvteam&","&ChkinvReviewed&");toggleNPTTable(this[this.selectedIndex].value);"">" & vbCRLF
		else
			tmp = tmp & "<SELECT style=""font-size:9pt"" name=""" & name & """ onchange=severityAlert(this,'"&msgcol&"',"&setflag&","&setInvflag&","&SqSev&","&HseSev&","&Chkinvteam&","&ChkinvReviewed&")>" & vbCRLF
		end if
		rs.MoveFirst
		Do While NOT rs.EOF
			tmp = tmp & "<OPTION value='" & rs("SeverityID") & "'" & rs("Selected") & ">" & rs("SeverityDesc") & vbCRLF
			rs.MoveNext
		Loop
		tmp = tmp & "</SELECT>"  & vbCRLF
		rs.Close
		Set rs = nothing	
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun GetSeverities QID="&SafeNum(iQPID)
	End If
		GetSeverities = tmp	
	End Function

	function ShowNPTSQBLOCKTable()
		dim strNPTTbl
		On Error Resume Next
		strNPTTbl = "<tr><td align='left' colspan='3'><table id='tblNPT' cellpadding='0' cellspacing='0' width='100%' border='1'>"
		'strNPTTbl = strNPTTbl & "<tr bgcolor='#eeeedd'><td colspan=4 align=center><b>Non-Productive Time Details</b> " & getHelpLink("NPT Details Criteria") & "</td></tr>"
		if isTimeLossEntered = 0 then
			strNPTTbl = strNPTTbl & "<tr height='26'><td width='28%'><table cellpadding='0' cellspacing='0' width='100%' border='0'><tr><td width='10%'><b>Severity Escalation Applied: " & getHelpLink("Severity Escalation Applied") & "</b></td><td width='42%'><input type=radio name='rdNPT' value='1' "
			if not bNR then strNPTTbl = strNPTTbl & iif(rs("NPTFlag"),"checked","")
			strNPTTbl = strNPTTbl & " onclick='toggleNPTRM();'> Yes&nbsp;&nbsp;<input type=radio name='rdNPT' value='0' "
			if not bNR then strNPTTbl = strNPTTbl & iif(not rs("NPTFlag") or IsNull(rs("NPTFlag")),"checked","") else strNPTTbl = strNPTTbl & "checked"
			strNPTTbl = strNPTTbl & " onclick='toggleNPTRM();'> No</td></tr></table></td></tr>"

		else
			strNPTTbl = strNPTTbl & "<tr height='26'><td width='28%'><table cellpadding='0' cellspacing='0' width='100%' border='0'><tr><td width='10%'><b>Severity Escalation Applied: " & getHelpLink("Severity Escalation Applied") & "</b></td><td width='42%'><input type=radio name='rdNPT' value='1' "
			strNPTTbl = strNPTTbl & iif(rs("NPTFlag"),"checked","")
			strNPTTbl = strNPTTbl & "> Yes&nbsp;&nbsp;<input type=radio name='rdNPT' value='0' "
			strNPTTbl = strNPTTbl & iif(not rs("NPTFlag") or IsNull(rs("NPTFlag")),"checked","")
			strNPTTbl = strNPTTbl & "> No</td></tr></table></td></tr>"
	
		end if
		strNPTTbl = strNPTTbl & "</table></td></tr>"
			If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun ShowNPTSQBLOCKTable QID="&SafeNum(iQPID)
	End If
		ShowNPTSQBLOCKTable = strNPTTbl	
	end function
	
	
	function ShowNPTTable()
		dim strNPTTbl
		On Error Resume Next
		strNPTTbl = "<tr><td align='left' colspan='3'><table id='tblNPT' cellpadding='0' cellspacing='0' width='100%' border='1'>"
		strNPTTbl = strNPTTbl & "<tr bgcolor='#eeeedd'><td colspan=4 align=center><b>Non-Productive Time Details</b> " & getHelpLink("NPT Details Criteria") & "</td></tr>"
		if isTimeLossEntered = 0 then
			strNPTTbl = strNPTTbl & "<tr height='26'><td width='28%'><table cellpadding='0' cellspacing='0' width='100%' border='0'><tr><td width='58%'><b>Severity Escalation Applied: " & getHelpLink("Severity Escalation Applied") & "</b></td><td width='42%'><input type=radio name='rdNPT' value='1' "
			if not bNR then strNPTTbl = strNPTTbl & iif(rs("NPTFlag"),"checked","")
			strNPTTbl = strNPTTbl & " onclick='toggleNPTRM();'> Yes&nbsp;&nbsp;<input type=radio name='rdNPT' value='0' "
			if not bNR then strNPTTbl = strNPTTbl & iif(not rs("NPTFlag") or IsNull(rs("NPTFlag")),"checked","") else strNPTTbl = strNPTTbl & "checked"
			strNPTTbl = strNPTTbl & " onclick='toggleNPTRM();'> No</td></tr></table></td>"
			
			strNPTTbl = strNPTTbl & "<td width='18%'><b>Overall NPT:</b> <input type=text name='txtNPT' size=2 value='"
			if not bNR then strNPTTbl = strNPTTbl & trim(rs("NPT"))
			strNPTTbl = strNPTTbl & "' maxlength=8 style='background-color:E0A198'><input type=text style='color:red;background:none;border:0;width:6px;' value='*' name='txtNPT_M' id='txtNPT_M' onfocus='document.getElementById(""txtNPT_LossCat_G1"").focus();' readonly> Hours</td>"
			strNPTTbl = strNPTTbl & "<td width='27%'><b>Estimated Client Red Money:</b> <input type=text name='txtNPT_LossCat_G1' id='txtNPT_LossCat_G1' size=2 value='"
			if not bNR then strNPTTbl = strNPTTbl & strNPT_LossCat_G1
			strNPTTbl = strNPTTbl & "' maxlength=8 style='background-color:E0A198'><input type=text style='color:red;background:none;border:0;width:6px;' value='*' name='txtNPT_LossCat_G1_M' onfocus='document.getElementById(""txtNPT_LossCat_G2"").focus();' id='txtNPT_LossCat_G1_M' readonly> <b>K$</b></td>"
			strNPTTbl = strNPTTbl & "<td width='27%'><b>Estimated SLB Red Money:</b> <input type=text name='txtNPT_LossCat_G2' id='txtNPT_LossCat_G2' size=2 value='"
			if not bNR then strNPTTbl = strNPTTbl & strNPT_LossCat_G2
			strNPTTbl = strNPTTbl & "' maxlength=8 style='background-color:E0A198'><input type=text style='color:red;background:none;border:0;width:6px;' value='*' name='txtNPT_LossCat_G2_M' onfocus='if (document.getElementById(""txtSPCategory"")) document.getElementById(""txtSPCategory"").focus(); else document.frmRIR.cmdSubmit.focus();' id='txtNPT_LossCat_G2_M' readonly> <b>K$</b></td></tr>"
		else
			strNPTTbl = strNPTTbl & "<tr height='26'><td width='28%'><table cellpadding='0' cellspacing='0' width='100%' border='0'><tr><td width='58%'><b>Severity Escalation Applied: " & getHelpLink("Severity Escalation Applied") & "</b></td><td width='42%'><input type=radio name='rdNPT' value='1' "
			strNPTTbl = strNPTTbl & iif(rs("NPTFlag"),"checked","")
			strNPTTbl = strNPTTbl & "> Yes&nbsp;&nbsp;<input type=radio name='rdNPT' value='0' "
			strNPTTbl = strNPTTbl & iif(not rs("NPTFlag") or IsNull(rs("NPTFlag")),"checked","")
			strNPTTbl = strNPTTbl & "> No</td></tr></table></td>"
			strNPTTbl = strNPTTbl & "<td width='18%'><b>Overall NPT:</b> " & iif(bNR,"",rs("NPT")) & " Hours</td>"
			strNPTTbl = strNPTTbl & "<td width='27%'><b>Estimated Client Red Money:</b> " & iif(bNR,"",strNPT_LossCat_G1) & " <b>K$</b></td>"
			strNPTTbl = strNPTTbl & "<td width='27%'><b>Estimated SLB Red Money:</b> " & iif(bNR,"",strNPT_LossCat_G2) & " <b>K$</b></td></tr>"
		end if
		strNPTTbl = strNPTTbl & "</table></td></tr>"
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun ShowNPTTable QID="&SafeNum(iQPID)
	End If
		ShowNPTTable = strNPTTbl	
	end function


	function ShowQnsTable()
		dim strQnsTbl
		On Error Resume Next
		strQnsTbl = "<tr><td align='left' colspan='3'><table id='tblNPT' cellpadding='0' cellspacing='0' width='100%' border='1'>"
		strQnsTbl = strQnsTbl & "<tr bgcolor='#eeeedd'><td colspan=4 align=center><b>Non-Productive Time Details</b> getHelpLink NPT Details Criteri</td></tr>"
		strQnsTbl = strQnsTbl & "</table></td></tr>"
		
		ShowQnsTable = strQnsTbl
			If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun ShowQnsTable QID="&SafeNum(iQPID)
	End If
	end function

	'Display the categories
	Sub ShowCat(a, hsesq)
		Dim row, column, currentHeading, currentSubHeading, displaySubHeading
		On Error Resume Next
		For row = 0 To UBound(a, 2)
			if (instr(hsesq,"hse")>0 AND lcase(a(3, row)) <> "quality") OR (instr(hsesq,"sq")>0 AND lcase(a(3, row)) = "quality") Then
				if currentHeading <> a(3, row) Then	
					if currentHeading <> "" Then
						Response.Write "</table></td>"
					End If
					currentHeading = a(3, row)
					Response.Write "<td valign=top>" & vbCRLF & "<table cellPadding='0' cellSpacing='0' width=""100%"" border=0>" & vbCRLF
					if CurrentHeading = "Quality" and LockCountSQ = 0  and not chkSQLockingMgmt() and iClass=1   then
					else
					Response.Write "<tr bgcolor='#eeeedd'><td colspan=2 align=center><b>" & ucase(CurrentHeading) & " LOSS</b></td></tr>" & vbCRLF
					End If
					currentSubHeading = ""
					displaySubHeading = ""
				End If

				if currentSubHeading <> cstr(a(1, row)) Then
					currentSubHeading = cstr(a(1, row))
					displaySubHeading = currentSubHeading
				Else
					displaySubHeading = ""
				End If

				'start add for SWIFT # 2410608 to display horizontal bar 
				if displaySubHeading ="Other" then
					If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then
						Response.write GetWellBarrier(1)
					Else
						Response.write "<tr><td colspan = 2><hr></hr></td></tr><tr><td colspan = 2>" & vbCrLf
						Response.write "<table border='0' cellpadding='2' cellspacing='0' width=100% >" & vbCrLf
						Response.write "<tr nowrap><TD valign='top' style=""font-size:9pt; width:52%; ""><b>Service & Equipment Specific Safety involved </b></td>" & vbCrLf
						Response.write "<td ><input type='radio' id=rdPLSSInvID1 name='rdPLSSInv' value='1' " & iif(bPLSSInv," checked ","") & " onclick='optHSESQ_onchange(7);'>Yes&nbsp;" & vbCRLF						
						Response.write "<input type='radio' id=rdPLSSInvID2 name='rdPLSSInv' value='0' " & iif(not bPLSSInv," checked ","") &  iif(not bPLSSInv," disabled ","") &" >No&nbsp;&nbsp;&nbsp;" & getHelpLink("Service & Equipment Specific Safety") & "</TD></tr>" & vbCRLF
						Response.write "</table></td></tr>" & vbCrLf
					End IF
						Response.Write "<tr><td colspan = 2><hr></hr></td></tr>"
				end if 
				' end add 
				Response.Write "<tr>" & vbCRLF
				'Response.Write "	<TD valign='top' style=""font-size:9pt""><b><u>" & displaySubHeading & "</u></b></TD>" & vbCRLF
				'Response.Write "	<TD valign='top' style=""font-size:9pt"">" & GetItem(a(2,row),a(0,row)) & "</TD>" & vbCRLF
				if displaySubHeading <> "" and Trim(currentHeading)="Safety" AND Row>2 then
					if a(1, row) = "Non-Productive Time"  and LockCountSQ = 0  and  not chkSQLockingMgmt() and iClass=1 then
				
					elseif trim(ucase(currentHeading)) = "HEALTH" or trim(ucase(currentHeading)) = "SAFETY" then									
						Response.Write "<TD valign='top' style=""font-size:9pt;border-top:thin solid;border-color:#989898;border-width:1.2px;""><b><u>" & displaySubHeading & "</u></b></TD></tr>" & vbCRLF
						Response.Write "	<tr><TD valign='top' style=""font-size:9pt"">" & GetItem(a(2,row),a(0,row)) & "</TD>" & vbCRLF
					else  
							
						Response.Write "<TD valign='top' style=""font-size:9pt;border-top:thin solid;border-color:#989898;border-width:1.2px;""><b><u>" & displaySubHeading & "</u></b></TD>" & vbCRLF
						Response.Write "	<TD valign='top' style=""font-size:9pt;border-top:thin solid;border-color:#989898;border-width:1.2px;"">" & GetItem(a(2,row),a(0,row)) & "</TD>" & vbCRLF
					End If	
				else
					if  a(1, row) = "Non-Productive Time"  and LockCountSQ = 0  and  not chkSQLockingMgmt() and iClass=1 then
						
					
					elseif trim(ucase(currentHeading)) = "HEALTH" or trim(ucase(currentHeading)) = "SAFETY" then
						Response.Write "	<TD valign='top' style=""font-size:9pt;""><b><u>" & displaySubHeading & "</u></b></TD></tr>" & vbCRLF
						Response.Write "	<tr><TD valign='top' style=""font-size:9pt"">" & GetItem(a(2,row),a(0,row)) & "</TD>" & vbCRLF
					else
					Response.Write "	<TD valign='top' style=""font-size:9pt;""><b><u>" & displaySubHeading & "</u></b></TD>" & vbCRLF
					Response.Write "	<TD valign='top' style=""font-size:9pt"">" & GetItem(a(2,row),a(0,row)) & "</TD>" & vbCRLF
					end if
				end if
				Response.Write "</tr>" & vbCRLF
			End If


		Next
		if trim(ucase(currentHeading)) = "ENVIRONMENT" and (Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv))) then
			Response.write GetWellBarrier(2)
		end if
		
		'DtEvtClass = fncD_Configuration("ShowEventClassification")
		'DiffEvtClass = DateDiff("n",DtEvtClass,dtRptDate)		
		IF (DiffEvtClass >= 1)  then 
		    if (instr(hsesq,"hse")>0) Then 
				If Cdate(FmtDateTime(dtRptDate)) < Cdate(FmtDateTime(VarHideWellBarrierInv)) then
					Response.Write GetEventCategorisation
				ELSE
					Response.Write GetNewEventCategorisation
				End IF
		    End if
		End if
		Response.Write "</table></td>"
	If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-in Sub FunShowCat QID="&SafeNum(iQPID)
	End If
	End Sub

	Function GetNewEventCategorisation()
		Dim strEvent,ChkHazardCat
		On Error Resume Next
		strEvent = "<tr><td colspan = 2></td></tr><tr><td colspan = 2><table border='0' cellpadding='2' width='100%' cellspacing='0' >" & vbCrLf
		strEvent = strEvent & "<tr><TD id=styleSmall style=""font-size:10pt"" colspan=4 align='center' bgcolor='eeeedd'><b>EVENT CATEGORIZATION&nbsp;&nbsp;&nbsp;&nbsp;</b></td></tr>"
		strEvent = strEvent & "<tr><TD id=styleSmall style=""font-size:9pt"" colspan=4><b>&nbsp;&nbsp;&nbsp;&nbsp;</b>"
		If HazardCat = 0 then ChkHazardCat = true else ChkHazardCat = false
		strEvent = strEvent & "Service & Equipment Specific Safety involved&nbsp;&nbsp;<input type='radio' id=rdEventSafetyProc name='rdEventSafety' value='0' " & iif(not EvtClassSaf or (not EvtClassSaf and not bNR) or not ChkHazardCat," checked ","") & " onclick='setSysMan();SetEventSubSafety();'>&nbsp;Yes" 
		strEvent = strEvent & "<input type='radio' id=rdEventSafetyPers name='rdEventSafety' value='1' " & iif((EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat)," checked ","") & " onclick='setSysMan();SetEventSubSafety();'>&nbsp;No&nbsp;&nbsp;&nbsp;" & getHelpLink("Service & Equipment Specific Safety") & "</td></tr>"
		strEvent = strEvent & "<td id=styleSmall style='display: none;'>&nbsp;&nbsp;&nbsp;<input type='radio' id=rdEventChoiceSys  name='rdEventChoice' value='1' "& iif(EvtClassChoice or bNR," checked ","") & " disabled>System&nbsp;" & vbCRLF
		strEvent = strEvent & "<input type='radio' id=rdEventChoiceMan name='rdEventChoice' value='0' "& iif(not EvtClassChoice and not bNR," checked ","") & " disabled>Manual</TD>" & vbCRLF
		strEvent = strEvent & "<input type='hidden' name='FireMode' value="&FireMode&">" & vbCrLf
		
		strEvent = strEvent & "<tr><td id=styleSmall><b><u>Service & Equipment Specific Safety involved Loss Category</b></td>"		
		strEvent = strEvent & "<tr><td id=styleSmall><input type='radio' id=rdEventSubCat1 name='rdEventSubCat' value='1' "& iif(EvtSubClassSaf=1," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Well Integrity - Well Barrier Involved&nbsp;" & getHelpLink("Well Integrity - Well Barrier Involved") & "</td>"
		strEvent = strEvent & "<td id=styleSmall><input type='radio' id=rdEventSubCat2 name='rdEventSubCat' value='2' "& iif(EvtSubClassSaf=2," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Process Safety - Loss/Potential Loss of Primary Containment&nbsp;" & getHelpLink("Process Safety - Loss/Potential Loss of Primary Containment") & "</td></tr>" & vbCRLF
		strEvent = strEvent & "<tr><td id=styleSmall><input type='radio' id=rdEventSubCat3 name='rdEventSubCat' value='3' "& iif(EvtSubClassSaf=3," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Well Integrity - Blow out&nbsp;" & getHelpLink("Well Integrity - Blow out") & "</td>"
		strEvent = strEvent & "<td id=styleSmall><input type='radio' id=rdEventSubCat4 name='rdEventSubCat' value='4' "& iif(EvtSubClassSaf=4," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Chemical Products sold/used in Service Delivery or processed by SLB&nbsp;" & getHelpLink("Chemical Products sold/used in Service Delivery or processed by SLB") & "</td></tr>" & vbCRLF
		strEvent = strEvent & "<tr><td id=styleSmall><input type='radio' id=rdEventSubCat5 name='rdEventSubCat' value='5' "& iif(EvtSubClassSaf=5," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Pressure Containing Equipment used in Service Delivery&nbsp;" & getHelpLink("Pressure Containing Equipment used in Service Delivery") & "</td>"
		strEvent = strEvent & "<td id=styleSmall><input type='radio' id=rdEventSubCat6 name='rdEventSubCat' value='6' "& iif(EvtSubClassSaf=6," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Lithium Batteries used in Wellsite Tools&nbsp;" & getHelpLink("Lithium Batteries used in Wellsite Tools") & "</td></tr>" & vbCRLF
		strEvent = strEvent & "<tr><td id=styleSmall><input type='radio' id=rdEventSubCat7 name='rdEventSubCat' value='7' "& iif(EvtSubClassSaf=7," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Fire - At Wellsite or Production Facility/Asset&nbsp;" & getHelpLink("Fire - At Wellsite or Production Facility/Asset") & "</td>" & vbCRLF
		strEvent = strEvent & "<td id=styleSmall><input type='radio' id=rdEventSubCat8 name='rdEventSubCat' value='8' "& iif(EvtSubClassSaf=8," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Explosives used in Service Delivery/Mining&nbsp;" & getHelpLink("Explosives used in Service Delivery/Mining") & "</td></tr>"
		strEvent = strEvent & "<tr><td></td><td id=styleSmall><input type='radio' id=rdEventSubCat9 name='rdEventSubCat' value='9' "& iif(EvtSubClassSaf=9," checked ","") & iif( EvtClassSaf or (EvtClassSaf and bNR) or ChkHazardCat," disabled","")& " onclick='optHSESQ_onchange(4);'>Ionizing Radiation&nbsp;" & getHelpLink("Ionizing Radiation") & "</td></tr>" & vbCRLF

		strEvent = strEvent & "</tr></table></td></tr>" & vbCrLf
		If Err.Number <> 0 Then
	' Log the ERROR
	LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun GetNewEventCategorisation QID="&SafeNum(iQPID)
	End If
		GetNewEventCategorisation = strEvent		
		
    End Function
	
	Function GetEventCategorisation()
		Dim strEvent
		On Error Resume Next
		strEvent = "<tr><td colspan = 2><hr></hr></td></tr><tr><td colspan = 2><table border='0' cellpadding='2' cellspacing='0' >" & vbCrLf
		strEvent = strEvent & "<tr><TD valign='top' style=""font-size:9pt"" colspan = 2><b><u>Event Categorisation</u></b> " & getHelpLink("Event Categorisation") & "</td></tr>"
		
		strEvent = strEvent & "<td id=styleSmall><input type='radio' id=rdEventSafetyPers name='rdEventSafety' value='1' " & iif((EvtClassSaf or bNR)," checked ","") & " onclick='setSysMan()'>Personal Safety&nbsp;"  & vbCRLF
		strEvent = strEvent & "<input type='radio' id=rdEventSafetyProc name='rdEventSafety' value='0' " & iif(not EvtClassSaf and not bNR," checked ","") & " onclick='setSysMan()'>Service & Equipment Specific Safety&nbsp;&nbsp;" & vbCRLF & vbCRLF
		
		strEvent = strEvent & "&nbsp;&nbsp;&nbsp;<input type='radio' id=rdEventChoiceSys  name='rdEventChoice' value='1' "& iif(EvtClassChoice or bNR," checked ","") & " disabled>System&nbsp;" & vbCRLF
		strEvent = strEvent & "<input type='radio' id=rdEventChoiceMan name='rdEventChoice' value='0' "& iif(not EvtClassChoice and not bNR," checked ","") & " disabled>Manual</TD>" & vbCRLF
		strEvent = strEvent & "<input type='hidden' name='FireMode' value="&FireMode&">" & vbCrLf
		
		strEvent = strEvent & "</tr></table></td></tr>" & vbCrLf
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun GetEventCategorisation QID="&SafeNum(iQPID)
		End If
		GetEventCategorisation = strEvent		

    End Function
		
	Function GetWellBarrier(Caller)
		Dim strWellB
		On Error Resume Next
		strWellB = "<tr><td colspan = 2><hr></hr></td></tr><tr><td colspan = 2><table border='0' cellpadding='2' cellspacing='0' >" & vbCrLf
		strWellB = strWellB & "<tr><TD valign='top' style=""font-size:9pt"" colspan = 2><b><u>Well Integrity - Barriers</u></b>"
		strWellB = strWellB & "&nbsp;<a href='JavaScript:WIBHelp()' class='plain'><img src='../images/movie1.gif' border=0 VSPACE=0 HSPACE=0 height=14 width=14 align='absmiddle'></a></td></tr>" & vbCrLf
				  
		strWellB = strWellB & "<tr nowrap><TD valign='top' style=""font-size:9pt; width:52%; "">Well Barrier Element Involved </td>" & vbCrLf
		if Caller = 1 then 'SQ
			strWellB = strWellB & "<td ><input type='radio' id=styleSmall name='rdWIBEventSQ' value='1' " & iif(bWIBEvent," checked ","") & " onclick='return fncToggleWIB(1,2);'>Yes&nbsp;" & vbCRLF
			strWellB = strWellB & "<input type='radio' id=styleSmall name='rdWIBEventSQ' value='0' " & iif(not bWIBEvent," checked ","") & " onclick='return fncToggleWIB(0,2);'>No</TD></tr>" & vbCRLF
		else
			strWellB = strWellB & "<td ><input type='radio' id=styleSmall name='rdWIBEventHSE' value='1' " & iif(bWIBEvent," checked ","") & iif(not bAccDischarge or (bSQ and bHSE)," disabled ","") & " onclick='optHSESQ_onchange(4);'>Yes&nbsp;" & vbCRLF
			strWellB = strWellB & "<input type='radio' id=styleSmall name='rdWIBEventHSE' value='0' " & iif(not bWIBEvent," checked ","") & iif(not bAccDischarge or (bSQ and bHSE)," disabled ","") & ">No</TD></tr>" & vbCRLF
		end if
		if Caller = 1 then 'SQ
			strWellB = strWellB & "<tr nowrap><TD valign='top' style=""font-size:9pt; width:52%; "">Accidental Discharge or Spill </td>" & vbCrLf
			strWellB = strWellB & "<td><input type='radio' id=styleSmall name='rdAccDischarge' value='1' " & iif(bAccDischarge," checked ","") & iif(not (bWIBEvent) and not (bSQ and bHSE)," disabled ","") & " onclick='optHSESQ_onchange(5);'>Yes&nbsp;" & vbCRLF
			strWellB = strWellB & "<input type='radio' id=styleSmall name='rdAccDischarge' value='0' " & iif(not bAccDischarge," checked ","") & iif(not (bWIBEvent) and not (bSQ and bHSE)," disabled ","") & ">No</TD></tr>" & vbCRLF
		end if        
		if Caller = 1 then 'SQ        
			strWellB = strWellB & "<tr nowrap><TD valign='top' style=""font-size:9pt; width:52%; "">Fire/Explosion  </td><input type='hidden' name='FireMode' value="&FireMode&">" & vbCrLf
			strWellB = strWellB & "<td><input type='radio' id=styleSmall name='rdFireExplosion' value='1' " & iif(bFireExpl," checked ","") & iif(not (bWIBEvent) and not (bSQ and bHSE)," disabled ","") & " onclick='optHSESQ_onchange(3);'>Yes&nbsp;" & vbCRLF
			strWellB = strWellB & "<input type='radio' id=styleSmall name='rdFireExplosion' value='0' " & iif(not bFireExpl," checked ","") & iif(not (bWIBEvent) and not (bSQ and bHSE)," disabled ","") & ">No</TD></tr>" & vbCRLF   
		end if        
		strWellB = strWellB & "</table></td></tr>" & vbCrLf
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun GetWellBarrier QID="&SafeNum(iQPID)
		End If
		GetWellBarrier = strWellB		
	End Function
	
	Function GetItem(strName,strValue)
		Dim strChecked, strJSFun, strtip, Title
		On Error Resume Next
		strChecked=""
		
		Title = strName
		IF Title = "Accidental Discharge or Spill" then
			strtip = getHelpLink(Title)
		End if
		
		IF Title = "Impact on Natural Environment" then
			strtip = getHelpLink(Title)
		End if
		
		IF Title = "Regulatory Sanctions or External Scrutiny" then
			strtip = getHelpLink(Title)
		End if
		
		IF Title = "Inappropriate Disposal" then
			strtip = getHelpLink(Title)
		End if
		
		
		If Not bNR and not bPostVars then 
			If RS(strValue) then strChecked="checked"
			If strValue = "LossCat_C1" and rs("AccDischarge") then strChecked="checked"
		End If

		If bPostVars then 
			If PostVars(strValue)="1" then strChecked="checked"
			If strValue = "LossCat_C1" and bAccDischarge then strChecked="checked"
		End If
		
		strJSFun = ""	
		If strName= "Health" Then strJSFun = " onclick='return LossCat_A2_onclick()'"
		If Instr(1,strValue,"_C") then strJSFun = " "
		If strValue = "LossCat_C1" then strJSFun = " onclick='javascript:return fncToggleWIB(this.checked,3)'"
		If Instr(1,strValue,"_G") AND iClass =1 then strJSFun = " onclick='javascript:return onCheck(this)'"

		' *****************************************************************
		' Code added for NPT <<2401608>> , to disable the checkboxes  
		' *****************************************************************
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description &"-In Fun GetItem QID="&SafeNum(iQPID) 
		End If
		GetItem = "<INPUT id=styleSmall " & strChecked & " name='" & strValue & "'" & strJSFun & " type='checkbox' value='1'>" & strName  & " " & strtip		
	End Function


	Function disablenpt (strValue)
	'****************************************************************************************
	'1. Function/Procedure Name          : disablenpt
	'2. Description           	         : Check various condition to disable the checkboxes.
	'3. Calling Forms:   	             : RIRdisp.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   5-Aug-2009			    Nilesh Naik  	         	Added- for NPT <<2401608>>
	 '****************************************************************************************
	dim blnflag
		On Error Resume Next	
	blnflag  =false
	 if(SQInvment= 2 and ClientAffect and (Instr(1,strValue,"_G1")or (Instr(1,strValue,"_G2")))) then 
		   blnflag    = true   
	 elseif(SQInvment= 2 and (Instr(1,strValue,"_G2") )) then 
		   blnflag   = true               
	 elseif(SQInvment= 1 and (Instr(1,strValue,"_G2") )) then 
		   blnflag   = true
	 elseif(SQInvment= 1 and ClientAffect and (Instr(1,strValue,"_G1") or (Instr(1,strValue,"_G2")) )) then 
		   blnflag   = true
	 else
		   blnflag = false        
	 end if 
	 If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun disablenpt QID="&SafeNum(iQPID)
		End If
	 disablenpt= blnflag
	
	
	End function 

	Function disableprocess ()
	'****************************************************************************************
	'1. Function/Procedure Name          : disableprocess
	'2. Description           	         : Check various condition to disable the checkboxes.
	'3. Calling Forms:   	             : RIRdisp.asp
	'4. History
	'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
	'   5-Aug-2009			    Nitesh Naik  	         	Added- for NPT <<2401608>>
	 '****************************************************************************************
	Dim blnflag 
	On Error Resume Next
	blnflag  =false
		'if((IPMIDS =2 or IPMIFS=3 or IPMISM =4 or IPMSPM=5 ) AND (not imploc) AND(iClass =1)  AND (bsq)) then  
		if((IPMIDS =2 or IPMIFS=3 or IPMSPM=5 ) AND (not imploc) AND(iClass =1)  AND (bsq)) then
			blnflag    = true
		end if 
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun disableprocess QID="&SafeNum(iQPID)
		End If
		disableprocess= blnflag
		
		
	 End function 
	 
	Sub ShowMockCat(a, hsesq)
		dim row
		On Error Resume Next
		For row = 0 To UBound(a, 2)
			if (instr(hsesq,"hse")>0 AND lcase(a(3, row)) <> "quality") OR (instr(hsesq,"sq")>0 AND lcase(a(3, row)) = "quality") Then
				Response.Write "<INPUT type=hidden name='" & a(0,row) & "' value='preserve'>" & vbCRLF
			End If
		Next
			If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun ShowMockCat QID="&SafeNum(iQPID)
		End If
	End Sub

	Sub WriteHeading(strHeading)
	On Error Resume Next
		Response.Write "<TR><TD class=field align=center><u>" & strHeading & "</u></TD></TR>"
			If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun WriteHeading QID="&SafeNum(iQPID)
		End If
	End Sub 


	Sub GenerateCatJS(a,typeMode)
		'frmdoc.LossCat_A1.checked || frmdoc.LossCat_A2.checked  etc...
		Dim row,column , i
		On Error Resume Next
		Dim Output()
		Redim Output(ubound(a,2))
		i=0
		For row = 0 To UBound(a, 2)
			if (typeMode = "hse" AND lcase(a(3, row)) <> "quality") OR (typeMode = "sq" AND lcase(a(3, row)) = "quality")Then
				Output(i) = "frmdoc." & a(0,row)&".checked"
				i = i + 1
			End If
		Next
		if i > 0 Then
			redim Preserve Output(i-1)
			response.Write Join(output," || ")
		Else
			Response.write "true"
		End If
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In SubFun GenerateCatJS QID="&SafeNum(iQPID)
		End If
	End Sub

	Sub InitVars()
	On Error Resume Next
		RptDate = "<span class=urgent >Pending</span>"
		ReportNumber = "Pending Creation"
		RevDate = "N/A"
		bReviewed = false
		bClosed = false
		UpdatedBy = ""
		CreatedBy = ""
		if Request.QueryString("Type") = "HSE" then bHSE=True
		If Request.QueryString("Type") = "SQ" then bSQ= True
		If Request.QueryString("Type") = "HSE-SQ" then 
			bHSE=True
			bSQ= True
		end if
		BusinessSegment = "0"
		Client = 0
		Contractor =""
		If sUID <> "" Then
			ReporterUID = sUID
			Reporter = sUName
		Else 
			ReporterUID = ""
			Reporter = "(NOT AVAILABLE)"
		end if
		EventDate = "mmm dd, yyyy"
		EventTime = "24Hr"
		Loctn = 0
		SiteName =  ""
		PRiskClass = ""
		RRiskClass = ""
		ShortDescription = ""
		LongDescription = ""
		iClass = 0
		If bNR Then iClass = Request.QueryString("Class")
		iSev = 0 
		iHSESev = 0 
		iSQSev = 0

		SQ_Process = 0
		SQ_ProcessOwn=0
		SQ_MetroStop=0
		SQ_Activity=0
		
		'TS changes
		Qlocation = lOrgNo
		
		SQSPCatID = 0
		SQFCatID = 0 
		SQDCatID = 0 
		SQCCatID = 0 
		SQDelayHrs = ""
		if isWSSQMapping(SQMappingID) or isDMSQMapping(SQMappingID) OR (IsGSS(GetSubBusinessSegID(lOrgNo)) And ShowGSS) Then SQStandard = "" else SQStandard = "1"  'isDM(lPL)   isWS(lPL)
		SQPFailure = ""
		SQNRedone = ""
		JobID	  =""
		SQSPSubcatID = 0
		SQFSubcatID = 0
		SQDSubcatID = 0
		SQCSubcatID = 0
		'OPC = 0
		SLBInv = True
		IndRec = True
		RegRec = False
		SLBRel = True
		SLBCon = True
		External = true
		' *****************************************************************
		' Changed for NPT <<2401608>>
		' *****************************************************************

		''ClientAffect = false
		 ClientAffect = true
		'*******************************************************************
		HazardCat= 0 
		Source = 0 ' Visali 06/02/2004
		IPMInv = False
		IPMNo=0
		IPMIDS = 0
		'IPMIPS = 0
		IPMIFS = 0
		'IPMISM = 0
		IPMSPM = 0
		PTECInv = 2
		ROPInv = 0
		ContractorInv=False
		TCCInv = False
		SegInv=False '***** (MS HIDDEN) - Commented line ***** 
		Grctcc=0
		bAccDischarge=False
		bFireExpl=False 
		bWIBEvent=False	
		EvtClassSaf=False
		EvtClassChoice=False
		bPLSSInv=False
		EvtSubClassSaf=0
		AccUnit=""
		WellSite=0
		tNPT = 0
		rNPT = 0
		If isDMSQMapping(SQMappingID) OR isRES(SQCategoryMappingID) OR isSIS(SQCategoryMappingID) OR isWCP(SQCategoryMappingID) or isSS(SQCategoryMappingID) or isAL(SQCategoryMappingID) or isCS(SQCategoryMappingID)  or isDCS(SQCategoryMappingID)  or isMNSIT(iBSID) or isOneCPL(iBSID) Then IsSPRequired = true else IsSPRequired = false   'isDM(lPL)
		If isIPMSeg(SQMappingID) or isREWSQMapping(SQMappingID)or isSPWL(iBSID) or (isOFS(iBSID) and not (isMNSIT(iBSID))) or isWTSSQMapping(SQMappingID) or isEMS(SQMappingID) or isWSSQMapping(SQMappingID) or isSwaco(SQMappingID) or (isOne(SQMappingID) and not (isOneCPL(iBSID))) Then isDamageRequired=false else IsDamageRequired = True  'isWTS(lPL)--isREW(lPL)   isWS(lPL)
		
		If bPostVars Then
			'Alert sKey

			If PostVars.Exists("RptDate") Then RptDate = FmtDate(PostVars("RptDate")) & " (UTC)"
			If PostVars.Exists("OrgNo") Then ReportNumber = getreportnumber(PostVars("OrgNo"))	
			If PostVars("chkReviewed")="on" then bReviewed = True else bReviewed = False
			If PostVars("chkClosed")="on" then bClosed = True else bClosed = False
			If PostVars("optSQ")="on" then bSQ = True else bSQ = False
			If PostVars("optHSE")="on" then bHSE = True else bHSE = False
			'If PostVars("optIPMInv") then IPMInv = True else IPMInv = False
			' If PostVars("optPTECInv") then PTECInv = 1 else PTECInv = 2  
			' If PostVars("optPTECInv")= false  then PTECInv = 0 else  PTECInv =2  
			PTECInv = iif(PostVars("optPTECInv")=1 , 1 , iif(PostVars("optPTECInv")= 0, 0 , 2 ) )
			ROPInv = iif(PostVars("optROPInv")=1 , 1 , iif(PostVars("optROPInv")= 0, 0 , 0 ) )			
			If PostVars("optTCCInv") then TCCInv = True else TCCInv = False
			If PostVars("optSegInv") then SegInv = True else SegInv = False '***** (MS HIDDEN) - Commented line ***** 
			if PostVars("optgot")=1 then Grctcc = 1 else Grctcc = 0
			If PostVars("optContractorInv") then ContractorInv = True else ContractorInv = False
			If PostVars("rdAccDischarge") or PostVars("LossCat_C1") then bAccDischarge = True else bAccDischarge = False
			If PostVars("rdWIBEventSQ") or PostVars("rdWIBEventHSE") or (PostVars("rdEventSubCat") = 1) then bWIBEvent = True else bWIBEvent = False		
			If PostVars("rdEventSafety") then EvtClassSaf = True else EvtClassSaf = False	
			If PostVars("rdPLSSInv") then bPLSSInv = True else bPLSSInv = False
			
			Dim SubCat
			If PostVars("rdEventSubCat") = "" then SubCat = 0 Else SubCat = PostVars("rdEventSubCat")
			Select Case SubCat
			  Case 1
				EvtSubClassSaf = 1
			  Case 2
				EvtSubClassSaf = 2
			  Case 3
				EvtSubClassSaf = 3
			  Case 4
				EvtSubClassSaf = 4
			  Case 5
				EvtSubClassSaf = 5
			  Case 6
				EvtSubClassSaf = 6
			  Case 7
				EvtSubClassSaf = 7
			  Case 8
				EvtSubClassSaf = 8
			  Case 9
				EvtSubClassSaf = 9
			End Select
							 			
			If PostVars("rdEventChoice") then EvtClassChoice = True else EvtClassChoice = False		
			
			If PostVars("rdFireExplosion") then bFireExpl = True else bFireExpl = False 
			BusinessSegment = PostVars("txtBSegment")
			Client = PostVars("txtClient")
			CRMClient = PostVars("txtCRMClient")
			CRMRigID = PostVars("txtCRMRigID")
			Contractor = PostVars("txtContractor")
			RetrieveFQNValues PostVars("txtReporter"), Reporter, ReporterUID, "", ""
			EventDate = FmtDate(PostVars("txtEvDate"))
			EventTime = Right(FmtDateTime(PostVars("txtEvTime")), 5)
			Loctn = PostVars("txtLocation")
			SiteName =  PostVars("txtLoc")
			ShortDescription = PostVars("txtShortDesc")
			LongDescription = PostVars("txtFullDesc")
			iClass = PostVars("optClass")
			
			iHSESev = cint(PostVars("cmbHSESeverity"))
			iSQSev = cint(PostVars("cmbSQSeverity"))		
			If PostVars.Exists("txtAccountUnit") Then AccUnit = PostVars("txtAccountUnit")
			If PostVars.Exists("txtSPCategory") Then SQSPCatID = PostVars("txtSPCategory")
			If PostVars.Exists("txtFailure") Then SQFCatID = PostVars("txtFailure")
			If PostVars.Exists("txtDamage") Then SQDCatID = PostVars("txtDamage")
			If PostVars.Exists("txtCause") Then SQCCatID = PostVars("txtCause")
			If PostVars.Exists("txtSQDelayHrs") Then SQDelayHrs = PostVars("txtSQDelayHrs")
			If PostVars.Exists("txtPFailure") Then SQPFailure = PostVars("txtPFailure")
			If PostVars.Exists("optSQStandard") Then SQStandard = PostVars("optSQStandard")
			If PostVars.Exists("txtSQNRedone") Then SQNRedone = PostVars("txtSQNRedone")
			If PostVars.Exists("txtJobID") Then JobID = PostVars("txtJobID")
			If PostVars.Exists("txtSubSPCategory") Then SQSPSubcatID = PostVars("txtSubSPCategory")
			If PostVars.Exists("txtSubFailure") Then SQFSubcatID = PostVars("txtSubFailure")
			If PostVars.Exists("txtSubDamage") Then SQDSubcatID = PostVars("txtSubDamage")
			If PostVars.Exists("txtSubCause") Then SQCSubcatID = PostVars("txtSubCause")
			If PostVars.Exists("optDay") Then RegRec = True else RegRec = False
			If PostVars.Exists("optSLBRel") Then SLBRel = True else SLBRel = False
			If PostVars.Exists("optCAffect") Then ClientAffect = True else ClientAffect = False
			If PostVars.Exists("optWellSite") Then wellSite = PostVars("optWellSite")
			
			If PostVars.Exists("SQB2_0") Then SQ_ProcessOwn = PostVars("SQB2_0")
			If PostVars.Exists("SQL2_0") Then SQ_Process = PostVars("SQL2_0")
			If PostVars.Exists("SQL3_0") Then SQ_MetroStop = PostVars("SQL3_0")
			If PostVars.Exists("SQL4_0") Then SQ_Activity = PostVars("SQL4_0")
			'TS changes
			
			 If PostVars.Exists("QuestLoc") Then  Qlocation = PostVars("QuestLoc")
				' response.write "post var - " & Qlocation
				' response.end
			' else
				' response.write "not present"
				' reponse.end
			' end if
			
			
			If PostVars.Exists("optSLBInvment") Then
			Select Case PostVars("optSLBInvment")
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
				case 4
					SLBInv = False
					IndRec = False
					SLBCon = False
			End Select
			End If

			If PostVars.Exists("optSQInvment") Then
				Select Case PostVars("optSQInvment")
					case 1
						SLBRel = True
						External = False
					case 2
						SLBRel = True
						External = true
						ClientAffect = True
					case 3
						SLBRel = False
						External = False
				End Select
			End If

			If PostVars.Exists("txtHazard") Then HazardCat = PostVars("txtHazard")
			If not bNR Then
				if not isnull(RS("UpdatedBy")) then 
					UpdatedBy = RS("UpdatedBy")
					UpdateUID = RS("UpdateUID")
				end if
				if not isnull(RS("CreateName")) then 
					CreatedBy = RS("CreateName") 
					CreateUID = RS("CreateUID")
				Else 
					CreatedBy="&lt;Not Available&gt;" 
					CreateUID =""
				end if
				if not isnull(RS("RevDate")) then RevDate =  RS("RevDate") & " (UTC)" 
			End If
		ElseIf Not bNR Then
			RptDate =  FmtDate(RS("RptDate")) & " (UTC)"
			ReportNumber = getReportNumber(RS("RptDate"))
			RevDate =  RS("RevDate") & " (UTC)"
			bReviewed = RS("Reviewed")
			bClosed = RS("Closed")
			If not isnull(RS("UpdatedBy")) then 
				UpdatedBy = RS("UpdatedBy")
				UpdateUID = RS("UpdateUID")
			end if
			If not isnull(RS("CreateName")) then 
				CreatedBy = RS("CreateName") 
				CreateUID = RS("CreateUID")
			Else 
				CreatedBy = "&lt;Not Available&gt;" 
				CreateUID = ""
			end if
			bSQ= (RS("ServiceQuality"))	
			bHSE = (RS("HSE"))
			dim js_tmpSQLossCats, js_tmpHSELossCats, js_tmpCategory
			js_tmpSQLossCats = split("LossCat_D1 LossCat_D2 LossCat_D3 LossCat_D4 LossCat_D5 LossCat_G1 LossCat_G2 LossCat_G3 LossCat_G4 LossCat_G5")
			js_tmpHSELossCats = split("LossCat_A1 LossCat_A2 LossCat_A3 LossCat_A4 LossCat_A5 LossCat_B1 LossCat_B2 LossCat_B3 LossCat_B4 LossCat_B5 LossCat_C1 LossCat_C2 LossCat_C3 LossCat_C4 LossCat_C5 LossCat_E1 LossCat_E2 LossCat_E3 LossCat_E4 LossCat_E5 LossCat_F1 LossCat_F2 LossCat_F3 LossCat_F4 LossCat_F5")
			
			if bSQ = false Then
				for each js_tmpCategory in js_tmpSQLossCats
					if rs(js_tmpCategory) = true then
						bSQ = true
					End If
				Next
				if bSQ Then
					response.write "<html> " 
					HSESQMessage = "It was not marked as a service quality event, but it has Service Quality loss categories selected.\n\nQUEST has automatically set the Service Quality flag for you, but it has not been saved yet."
				End If			
			End If
			
			if bHSE = false Then
				for each js_tmpCategory in js_tmpHSELossCats
				   if rs(js_tmpCategory) = true then
						bHSE = true
					End If
				Next
				
				if bHSE Then
					HSESQMessage = "It was not marked as an HSE event, but it has HSE loss categories selected.\n\nQUEST has automatically set the HSE flag for you, but it has not been saved yet."
				End If			
			End If
			
			if HSESQMessage <> "" Then
				HSESQMessage = "This Report violates HSE/SQ rules:\n\n" & HSESQMessage & "\n\nPlease make any required corrections, and save it."
			End If		

			BusinessSegment = Trim(RS("BusinessSegment"))
			Client = Trim(RS("CustID"))
			CRMClient = UCase(Trim(RS("CRMClient")))
			CRMRigID = UCase(Trim(RS("CRMRigID")))
			Contractor = RS("ContractorID")
			Reporter = RS("Reporter") 
			ReporterUID = RS("ReporterUID") 
			EventDate = FmtDate(RS("EventDateTime"))
			EventTime = Right(FmtDateTime(RS("EventDateTime")), 5)
			Loctn = Trim(RS("LocType"))
			SiteName =  RS("Location")
			PRiskClass = RS("PRiskClass")
			RRiskClass = RS("RRiskClass")
			ShortDescription = RS("ShortDesc")
			LongDescription = RS("FullDesc")
			'call Str2Asc(LongDescription)
			iClass = RS("Class")
			iSev = RS("Severity")
			iHSESev = RS("HSESeverity")
			iSQSev = RS("SQSeverity")
			SQSPCatID = trim(RS("SQSPCatID"))
			SQFCatID = trim(RS("SQFCatID"))
			SQDCatID = trim(RS("SQDCatID"))
			SQCCatID = trim(RS("SQCCatID"))
			SQStandard = RS("SQDelayHrs")
			SQPFailure = RS("PFailure")
			'OPC = RS("OPC")
			SQNRedone = RS("SQNRedone")
			JobID	  = RS("JobID")
			SQSPSubcatID = RS("SQSPSubcatID")
			SQFSubcatID = RS("SQFSubcatID")
			SQDSubcatID = RS("SQDSubcatID")
			SQCSubcatID = RS("SQCSubcatID") 
			SLBInv = RS("SLBInv")
			IndRec = RS("IndRec")
			RegRec = RS("Daylight")
			SLBRel = RS("SLBRelated")
			ClientAffect = RS("ClientAffected")
			External = RS("RIRExternal")
			SLBCon = RS("SLBConcerned")
			HazardCat= Trim(RS("HazardCat"))
			IPMInv = RS("IPMInv")
			IPMNo = iif(isnull(rs("ProjectNO")), 0 , rs("ProjectNO"))  
			IPMIDS = iif(isnull(rs("ProjectIDS")), 0 , rs("ProjectIDS"))  
			'IPMIPS =iif(isnull(rs("ProjectIPS")), 0 , rs("ProjectIPS")) 
			IPMIFS =iif(isnull(rs("ProjectIFS")), 0 , rs("ProjectIFS"))
			'IPMISM = iif(isnull(rs("ProjectISM")), 0 , rs("ProjectISM"))  
			IPMSPM = iif(isnull(rs("ProjectSPM")), 0 , rs("ProjectSPM"))
			PTECInv = iif( rs("PTEC") , 1 , iif(rs("PTEC")= false , 0,2 ))
			ROPInv = iif( rs("ROP") , 1 , iif(rs("ROP")= false , 0,0 ))
			TCCInv = (rs("TCCInvolved"))
			SegInv=trim(rs("IsSegmentInv")) '***** (MS HIDDEN) - Commented line ***** 
			Grctcc= rs("GRCInv")
			
			ContractorInv=rs("ContractorID")>0
			Source = (rs("Source"))
			AccUnit = rs("AccountUnit")
			WellSite=RS("WellSite")
			tNPT = SafeNum(trim(RS("NPT")))
			FinanceInv=RS("IsFinance") '***** (MS HIDDEN) - Commented line  ***** 
			if isnull(tNPT) or tNPT = "" then tNPT = 0
			if isnull(SegInv) then SegInv=False '***** (MS HIDDEN) - Commented line  ***** 
			bAccDischarge = iif(rs("AccDischarge") or rs("LossCat_C1"),true,false)
			bWIBEvent = iif(rs("WIBEvent"),true,false)		
			
			EvtClassSaf = iif(rs("EventClassSafety"),true,false)		
			EvtClassChoice = iif(rs("EventClassChoice"),true,false)		
			bFireExpl = iif(rs("FireExplosion"),true,false)
			EvtSubClassSaf=RS("EventSubClassSafety")
			bPLSSInv = iif(rs("PLSSInv"),true,false)
		End If
		If Client="" then Client =0
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In SubFun InitVars QID="&SafeNum(iQPID)
		End If

	End Sub 

	Sub DisplayWarning (Source)
		Dim sType,Str
		On Error Resume Next
		'If source is RMS and Ritewin then display the warning
		sType=""
		If Source =6 then 	sType = "RMS"
		If Source =3 then 	sType = "RITEWIN"
		If Source =8 then 	sType = "INDEX"
		If sType <>"" and Not RS("Closed") then
			Str="<div class=Urgent align=center><br>This report originated from " & sType & " and is still open."
			If Source=3 and RS("Severity")>2 Then 
				Str=Str & " Automatic updates from RITEWIN to QUEST have been disabled for all C & M reports.<BR>" 
				Str=Str & " Follow up and closure for this report must be performed from within QUEST.</div>"
			Else
				Str=Str & " Any modifications made to this page may be overwritten by automatic updates from " & sType & " unless the report is closed.</div>"
			End IF
			Response.Write Str
		End If
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In SubFun DisplayWarning QID="&SafeNum(iQPID)
		End If

	End Sub

	Sub SetSeverityMatrix (lPl)
		dim rs
		On Error Resume Next
		Set rs=cn.execute("SELECT HSESeverityURL, SQSeverityURL FROM tblProductLines WHERE PLID = " & lPl)
		if not rs.eof Then
				sSQSeverityMatrix = RS("SQSeverityURL")
				sHSESeverityMatrix = RS("HSESeverityURL")
		End If
		Rs.Close
		Set Rs = Nothing
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Subfun  SetSeverityMatrix QID="&SafeNum(iQPID)
		End If

	End Sub

	'@Sreedhar - Mooved this function from In_Functions 27sep2006
	Function getAccUnits(OrgNo,SelVal,cn,iPL)
	Dim SQL,RS,Opts,sel,temp,ln,loc
	On Error Resume Next

		Opts = ""
		
		
		SQL="Select AccUnit_ID as AUID,AccUnit_Code+' : '+AccUnit_Description as Description "
		SQL=SQL & " from tlkpSQ_AccountingUnits "
		SQL=SQL & " Where ((AccUnit_Id='"&SelVal&"') or "
		if ( isREWSQMapping(SQMappingID) or isSPWL(iBSID) ) then SQL=SQL & " (((AccUnit_Segment='WL' and (AccUnit_SubSubSegement!='WCH1' or AccUnit_Area='NAM' or Accunit_Geomarket in (" & VarReplaceNAMwithGeomarket &"))) or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TWL')) and AccUnit_deleted=0 "   'isREW(iPL) 
		if (isWTSSQMapping(SQMappingID) and not isCSUB(SQMappingID)) then SQL=SQL & " (((AccUnit_Segment='TPS')or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TTS')) and AccUnit_deleted=0 "   'isWTS(iPL)
		if (isOSTPDPMapping(SQMappingID) ) then SQL=SQL & " ((((AccUnit_Segment='WL' and (AccUnit_SubSubSegement!= '' or AccUnit_Area='NAM' or Accunit_Geomarket in (" & VarReplaceNAMwithGeomarket &"))) or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TWL') or ( (AccUnit_Segment='WSV' and AccUnit_SubSubSegement= 'NPWL')))) and AccUnit_deleted=0 "   'isWTS(iPL) 

		
		if isCSUB(SQMappingID) then SQL=SQL & " ((AccUnit_Segment='CSUB' and AccUnit_deleted=0)  " 
		SQL=SQL & " and AccUnit_Country in (Select Abbr1 from tlkpCountries C inner join tblQT_QuestTree QT with(NOLOCK) on QT.CountryID=C.ID Where QT.ID='"&OrgNo&"')))"
		SQL=SQL & " Order by Description "
		'response.write SQL
		Set RS=cn.execute(SQL)
		Do while not rs.eof
			If trim(selVal) = trim(rs("AUID")) Then sel = "selected" Else sel = ""
			Opts = Opts & "<OPTION " & sel & " value='" & trim(rs("AUID")) & "'>" & rs("Description") & vbCRLF
			rs.movenext
		Loop
		rs.close
		Set rs = nothing
		

		
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun getAccUnits QID="&SafeNum(iQPID)
		End If
		getAccUnits=Opts
	End Function
	
	
	


Function getMaxAccUnits(OrgNo,SelVal,cn,iPL)
	Dim SQL,RS,Opts,sel,temp,ln,loc
	On Error Resume Next

		Opts = ""
		
		
		SQL="Select AccUnit_ID as AUID,AccUnit_Code+' : '+AccUnit_Description as Description "
		SQL=SQL & " from tlkpSQ_AccountingUnits_Maximo "
		SQL=SQL & " Where ((AccUnit_Id='"&SelVal&"') or "
		if ( isREWSQMapping(SQMappingID) or isSPWL(iBSID) ) then SQL=SQL & " (((AccUnit_Segment='WL' and (AccUnit_SubSubSegement!='WCH1' or AccUnit_Area='NAM' or Accunit_Geomarket in (" & VarReplaceNAMwithGeomarket &"))) or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TWL')) and AccUnit_deleted=0 "   'isREW(iPL) 
		if (isWTSSQMapping(SQMappingID) and not isCSUB(SQMappingID)) then SQL=SQL & " (((AccUnit_Segment='TPS')or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TTS')) and AccUnit_deleted=0 "   'isWTS(iPL)
		if (isOSTPDPMapping(SQMappingID) ) then SQL=SQL & " ((((AccUnit_Segment='WL' and (AccUnit_SubSubSegement!= '' or AccUnit_Area='NAM' or Accunit_Geomarket in (" & VarReplaceNAMwithGeomarket &"))) or (AccUnit_Segment='TLMS' and AccUnit_SubSegment= 'TWL') or ( (AccUnit_Segment='WSV' and AccUnit_SubSubSegement= 'NPWL')))) and AccUnit_deleted=0 "   'isWTS(iPL) 

		
		if isCSUB(SQMappingID) then SQL=SQL & " ((AccUnit_Segment='CSUB' and AccUnit_deleted=0)  " 
		SQL=SQL & " and AccUnit_Country in (Select Abbr1 from tlkpCountries C inner join tblQT_QuestTree QT with(NOLOCK) on QT.CountryID=C.ID Where QT.ID='"&OrgNo&"')))"
		SQL=SQL & " Order by Description "
		Set RS=cn.execute(SQL)
		Do while not rs.eof
			If trim(selVal) = trim(rs("AUID")) Then sel = "selected" Else sel = ""
			Opts = Opts & "<OPTION " & sel & " value='" & trim(rs("AUID")) & "'>" & rs("Description") & vbCRLF
			rs.movenext
		Loop
		rs.close
		Set rs = nothing
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-In Fun getAccUnits QID="&SafeNum(iQPID)
		End If
		getMaxAccUnits=Opts		
	End Function


	Function DisplaySQFunctions(QPID,Mode)
	Dim RS,SQL,FunID,FunName,isSel
	Dim Result,Img,optCheck,ChkCount,Ct,sClass ,oRole,oSec,cols
	On Error Resume Next
		
		SQL="SELECT A.FunID, FunName,0 as isSel "
		SQL=SQL & " From tlkpSQFunctions A "
		SQL=SQL & " Order by FunOrder,FunName "
			
		set RS=cn.execute(SQL)
		Img="<img src='../images/checkmrk.gif'>"	
		Response.Write "<table cellpadding=0 cellspacing=0 border=0 width=100% >"
		
		ct=0
		cols=4
		While Not RS.EOF 
			FunID=rs("FunID")
			FunName=rs("FunName")
			isSel=RS("isSel")			
			optCheck="<Input type=checkbox name=FunIDs value='"&FunID&"' "&iif(isSel>0," checked","")& ">"				
			If Mode=1 then optCheck=Img
					
			If Ct Mod cols =0 Then Result="<TR>"
			Result=Result & "<td align ='center' valign='center' id=styleSmall>"&optCheck&"</td>"
			Result=Result & "<td valign='center' nowrap id=styleSmall>&nbsp;"&FunName&"</td>"
			ct = ct + 1	
			If Ct Mod cols =0 then Response.Write Result & "</tr>"		
			RS.MoveNext		
		Wend
			If Ct Mod cols <>0 then	Response.Write Result & "<td colspan="&(cols-(ct mod cols))*2&">&nbsp;</td></tr>"
			
		Response.Write "</table>"
		RS.Close
		Set RS=nothing
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InFun DisplaySQFunctions QID="&SafeNum(iQPID)
		End If
	End Function


	Sub BuildLocInfo(lOrgNo,SubSeg,Dep,cn)
		Dim RS,SQL,LocNM,PLNM,Country,AreaNm,GMNm,LocLink,DepStr,SegStr,Str,PLID,EnforceSel, LocBSID
		On Error Resume Next
		SQL="Select PLID,Name,ProductLine,Country,AreaCode,GMDescription,PID,EnforceSelection,BSID from vw_QuestLocations Where ID="&lOrgNo
		Set RS=cn.execute(SQL)
		If Not RS.EOF Then
			LocNm=RS("Name")
			PLNm=RS("ProductLine")
			Country=RS("Country")
			LOCCountry=RS("Country")
			
			%>
			<input type="hidden" name="loccnt" value="<%=LOCCountry%>">
			
			<%
			PLID = RS("PLID")
			GMNm=RS("AreaCode")& " - " & RS("GMDescription")
			EnforceSel = RS("EnforceSelection")
			'TS chagnes
			LocPID = RS("PID")
			LocBSID = RS("BSID")
		End IF
		RS.Close
		Set RS=Nothing
		
		If isGuest() or bNR Then
			'TS changes
			SQL = " Select ID, Name from tblQT_QuestTree where PID = " & LocPID & "and RecType = 1 and inactivedate is null order by Name  "
			'response.write "value" & EnforceSel
			'response.end
        	if EnforceSel then 
				LocLink="<SELECT name=QuestLoc id=QuestLoc  id=styleSmall  onchange='return optHSESQ_onchange(6)'>"
				LocLink=LocLink & "<OPTION value='0'>(Selection Not Required)"
				
				LocLink=LocLink & GetOptions(SQL,Qlocation)
				LocLink=LocLink & "</SELECT>"
				
				'LocLink = GetOptions(SQL,lOrgNo)
			else
				LocLink=SafeDisplay(LocNm)
			end if
		Else
			'SWIFT # 2409503 - Start - addeded &CheckTopNodes=1
			LocLink="<a href='../Utils/ReAssignLocation.asp"&sKey&"&CallTree=1&Type=RIR&CheckTopNodes=1' "
			'SWIFT # 2409503 - End
			LocLink=LocLink & " onMouseOver=""window.status= 're-assign location';return true;"" onMouseOut=""window.status='';return true ;"">"
			LocLink=LocLink & SafeDisplay(LocNm) & "</a>"
		End IF 
		

		
		DepStr="<SELECT name=lstDepartment  id=styleSmall onChange='hidBSID.selectedIndex=this.selectedIndex;'>"
		DepStr=DepStr & "<OPTION value='0'>(Selection Not Required)"
		If bNR then
		DepStr=DepStr & GetDepartmentOptions(lOrgNo,Dep,1,iQPID,cn)
		else
		DepStr=DepStr & GetDepartmentOptions(lOrgNo,Dep,"",iQPID,cn)
		End if
		DepStr=DepStr & "</SELECT>"	
		
		Str = "<input type='hidden' name='LocBSID' value= " & LocBSID & ">"
		Str = Str & "<input type='hidden' name='EnforceSelection' value= " & EnforceSel & ">"
		
		Str=Str & "<TABLE border=0 cellPadding=1 cellSpacing=1 width='100%'>" &vbCRLF
		Str=Str & getEventSegmentInfo_RIR(lOrgNo,"txtBSegment",SubSeg,cn)
		Str=Str & "<input type=hidden name=PLID value='" & PLID & "'>"
		Str=Str & "<TR><TD class=field>QUEST Location:&nbsp;</TD><TD>&nbsp;"&LocLink&"</TD></TR>"&vbCRLF
		'Str=Str & "<TR><TD class=field>Department:&nbsp;</TD><TD>&nbsp;"&DepStr&"</TD></TR>"&vbCRLF
		Str=Str & "<TR><TD class=field>Country:&nbsp;</TD><TD>&nbsp;"&SafeDisplay(Country)&"</TD></TR>"&vbCRLF
		Str=Str & "<TR><TD style='Display:None;' class=field>TEST:&nbsp;</TD><TD style='Display:None;'>&nbsp;"&GetHiddenBSID(lDepartmentID)&"</TD></TR>"&vbCRLF
		Str=Str & "<TR><TD class=field>GeoMarket:&nbsp;</TD><TD>&nbsp;"&SafeDisplay(GMNm)&"</TD></TR>"&vbCRLF
		
		Str=Str & "<TR>"&vbCRLF
		Str=Str & "<TD class=field>Reporter:&nbsp;</TD>"&vbCRLF	
		Str=Str & "<TD nowrap>"&vbCRLF
		Str=Str & "<select name='txtReporter' LANGUAGE=javascript onChange='return txtReporter_onchange()'  id=Select1>"&vbCRLF
		If trim(sUID) <> trim(lcase(ReporterUID)) Or isNull(ReporterUID) then 
		Str=Str & "<option value="& GetFQN(sUID,NULL,sUName) & ">"
		Str=Str & sUName &vbCRLF
		Str=Str & "end if "&vbCRLF
		end if						
		Str=Str & "<option selected value="& GetFQN(ReporterUID,NULL,Reporter) & ">"
	    Str=Str &  Reporter &vbCRLF
		Str=Str & "<option value='Anonymous'>Anonymous"&vbCRLF
		Str=Str & "<option value='Consultant'>Consultant"&vbCRLF
		Str=Str & "<option value='Third Party'>Third Party"&vbCRLF
		Str=Str & "<option value='Sub-contractor'>Sub-contractor"&vbCRLF
		Str=Str & "<option value=''>(SEARCH MORE EMPLOYEES)"&vbCRLF
		Str=Str & Session("SearchedEmployees")&vbCRLF
		Str=Str & "</select>"&vbCRLF
		If Reporter <> "" then
		Str=Str & "<a href='javascript:ldapSearch(document.frmRIR.txtReporter)'><IMG SRC='../images/phonebook.gif' border=0></a>"&vbCRLF
		end if
		Str=Str & mSymbol &vbCRLF
		Str=Str & "</td>"&vbCRLF										
		Str=Str & "</tr>"&vbCRLF
		
		
		
		Str=Str & "</TABLE>"
		'TS changes
		Str=Str & " <input type='hidden' name='QLoc' value= " & Qlocation & " >"
		Response.Write Str	
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun BuildLocInfo QID="&SafeNum(iQPID)
		End If
	End Sub



	sub Str2Asc(Obj)
	dim str,i,ch,a
		On Error Resume Next
		if lcase(Session("UID"))= "svadla" Then 
			str=""		
			for i=1 to len(Obj)
				ch=Mid(Obj, i, 1)
				a=Asc(ch)
				if (a<23) then
					str=str & "("&a&")"
				else
					str=str & ch 
				end if			
			next
			Response.Write str
		End IF
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun Str2Asc QID="&SafeNum(iQPID)
		End If
	End Sub

	
	
	
	Function getSiteTypes(selID)
		Dim sRS,sSQL,sStr,selStr,optval,opttxt,optgrp
		Dim DtSiteLocation,DiffSiteLocation
		On Error Resume Next
			DtSiteLocation = fncD_Configuration("SiteLocation")
			DiffSiteLocation = DateDiff("n",DtSiteLocation,dtRptDate)		
	    
		If instr(selID,":") Then selID=split(selID,":")(0)
		selID=SafeNum(selID)
		sSQL = "SELECT LocationID,LocationDesc,isnull(LocGroup,'NON') as LocGroup FROM tlkpRIRLocationType WHERE isNULL(status,0) <> 1 "
		 
			IF (DiffSiteLocation < 1)  then 
                'sSQL= sSQL & " AND LocationID in (4,9) "
            ELSE
				sSQL= sSQL & " AND LocationID NOT in (17,18,19,20,7) "
            End if                
			sSQL = sSQL & " or locationid ='"&Trim(selID)&"'  ORDER BY LocationDesc "
			'response.write (sSQL)
		If  selID=0 Then selStr="selected"
		sStr="<select name='txtLocation' id=styleSmall OnChange='OnChange_SiteType(this,1);checkifDirectDMSQ();'>"
		sStr=sStr & "<option "&selStr&" value='0'>(Selection Required) "
		Set sRS=cn.execute(sSQL)
		While not sRS.EOF
			 optval=trim(sRS("LocationID"))
			 opttxt=sRS("LocationDesc")
			 optgrp=sRS("LocGroup")
			 selstr=""
			 If  selID=optval Then selStr="selected"
			 sStr=sStr & "<option "&selStr&" value='"&optval&":"&optgrp&"'>"&opttxt																						
			 sRS.MoveNext
		Wend
		sStr=sStr & "</select>"
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InFun getSiteTypes  QID="&SafeNum(iQPID)
		End If
		getSiteTypes=sStr		
	End Function
	
	
	Function GetSQOperationsCat(selValue)
		dim SQL
		On Error Resume Next
		SQL="SELECT ID,OperationName as Name FROM tlkpSQOperations With (NOLOCK)  where PID = 0 ORDER BY OperationName"
		GetSQOperationsCat=GetOptions(SQL,selValue)
		
			If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InFun GetSQOperationsCat  QID="&SafeNum(iQPID)
		End If
	End Function

	Function GetSQOperationsSubCatCat(ID,selValue)
		dim SQL
		On Error Resume Next
		SQL="SELECT ID,OperationName as Name FROM tlkpSQOperations With (NOLOCK)  where PID = " & ID & " ORDER BY ID"
		GetSQOperationsSubCatCat=GetOptions(SQL,selValue)
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InFun GetSQOperationsSubCatCat  QID="&SafeNum(iQPID)
		End If
	End Function
	
	Function GetOptions(SQL,selID)
		Dim RS,Cn,Str,SelStr
		On Error Resume Next
		SET cn = GetNewCN()
		SET RS=cn.execute((SQL))
		Str=""
		'Str= "<!--"&SQL&"-->"
		While Not RS.EOF 
			'Str=Str& "<!--"&rs("Name")&"-->"
			If Trim(SelID)=Trim(RS("ID")) Then SelStr=" selected" else SelStr=" "
			Str = Str & "<OPTION " & selStr & " value=" & rs("ID") & ">" & SafeDisplay(rs("Name")) & vbCRLF
			rs.movenext
		Wend
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InFun GetOptions  QID="&SafeNum(iQPID)
		End If
		GetOptions=Str
		RS.Close
		cn.Close		
	End Function
	
	Sub LoadDT_SQOperation(Obj,SelID)
		Dim PSQL,SQL
		On Error Resume Next
		'IF SelID = "0" then
			SQL="SELECT ID,PID,OperationName As Name,CASE When ID =" & SafeNum(SelID) & " Then 'true' ELSE 'false' END AS Selected FROM tlkpSQOperations With (NOLOCK)  where PID <> 0  ORDER BY ID"
		'else
			'SQL="SELECT ID,PID,OperationName As Name FROM tlkpSQOperations With (NOLOCK)  where PID =  " & SafeNum(SelID) & "  ORDER BY ID"
		'end if
		LoadJSData_SQOperation SQL,Obj,1	
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun LoadDT_SQOperation  QID="&SafeNum(iQPID)
		End If
	End Sub
	
	
	Sub LoadJSData_SQOperation(SQL,Obj,isALL)
		dim RS
		On Error Resume Next
		set RS=cn.execute((SQL))	
		While Not RS.EOF  
			Response.Write vbTab & vbTab & Obj&".load('" & RS("PID") & "','" & RS("ID") & "','" & RS("Name") & "'," & RS("Selected") & ",0," & RS("ID") & ",'1');" & vbCRLF	
			rs.movenext
		Wend
		RS.Close
		If Err.Number <> 0 Then
		' Log the ERROR
		LogEntry 2,"RIRDsp.asp",err.Description  &"-InSubFun LoadJSData_SQOperation  QID="&SafeNum(iQPID)
		End If
	End Sub


%>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:RIRdsp.asp;134 %>
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
<% '       7*[1266818] 15-SEP-2009 16:26:34 (GMT) NNaik %>
<% '         "2411197 :Expand the 'Site' dropdown with additional items" %>
<% '       8*[1271052] 25-SEP-2009 20:36:20 (GMT) VGrandhi %>
<% '         "SWIFT #2403986 - Well Services SQ Tabs" %>
<% '       9*[1275699] 29-SEP-2009 15:48:26 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      10*[1277927] 02-OCT-2009 17:50:40 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      11*[1279417] 05-OCT-2009 22:57:37 (GMT) VGrandhi %>
<% '         "SWIFT #2403986 - Fixes/Additions from Stephan in Well Services SQ Tabs" %>
<% '      12*[1277363] 07-OCT-2009 14:50:18 (GMT) MAnthony2 %>
<% '         "SWIFT# 2401608 NPT Change Request" %>
<% '      13*[1282923] 16-OCT-2009 21:28:41 (GMT) MAnthony2 %>
<% '         "SWIFT # 2409503 - Possible now to RELOCATE RIR outside topnode" %>
<% '      14*[1287289] 28-OCT-2009 16:23:18 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab" %>
<% '      15*[1289354] 03-NOV-2009 05:23:47 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      16*[1290343] 04-NOV-2009 13:41:57 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      17*[1292460] 17-NOV-2009 17:42:23 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - Changes for Section 3 for Rig Related." %>
<% '      18*[1295941] 19-NOV-2009 16:46:01 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - 11-16-2009 Jon Changes" %>
<% '      19*[1298532] 24-NOV-2009 17:49:46 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - Changed closing message" %>
<% '      20*[1303867] 23-DEC-2009 17:31:09 (GMT) VGrandhi %>
<% '         "SWIFT #2448303 - Develop EMS SQ Tab" %>
<% '      21*[1303811] 23-DEC-2009 17:53:32 (GMT) VGrandhi %>
<% '         "2447849 - Hide SQ Questions for IPM HSE events" %>
<% '      22*[1333698] 16-MAR-2010 05:59:52 (GMT) DMohanty %>
<% '         "Swift #2463864 - Data Gathering informations disappearing from RIR investigation tab" %>
<% '      23*[1340591] 29-MAR-2010 19:31:28 (GMT) DMohanty %>
<% '         "Swift # 2463864 - Data Gathering informations disappearing from RIR investigation tab" %>
<% '      24*[1340715] 29-MAR-2010 21:45:17 (GMT) DMohanty %>
<% '         "Swift # 2471699 -IPM Project / SQ Details tab" %>
<% '      25*[1342944] 07-APR-2010 11:33:24 (GMT) NNaik %>
<% '         "SWIFT # 2468497 - Add Pop-up tp Quality Loss selections" %>
<% '      26*[1348951] 21-APR-2010 10:25:19 (GMT) MAnthony2 %>
<% '         "2476332 - Selection Issue for NPT Loss- RIR Page" %>
<% '      27*[1349607] 05-MAY-2010 13:23:19 (GMT) APrakash6 %>
<% '         "SWIFT #2467716 - Add the GIN# next to the creator name" %>
<% '      28*[1345547] 25-MAY-2010 16:06:55 (GMT) DMohanty %>
<% '         "Swift # 2474157 - D&M SQ tab Development" %>
<% '      29*[1359004] 04-JUN-2010 21:43:45 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Develop interface to extract RIG NAME and associated info from CRM" %>
<% '      30*[1366377] 10-JUN-2010 06:51:40 (GMT) DMohanty %>
<% '         "Swift # 2474157 - D&M SQ tab Development (Reported error rectification)" %>
<% '      31*[1368943] 21-JUN-2010 14:44:07 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Not selected Site Type when we switch to HSE/SQ/Classification" %>
<% '      32*[1371202] 22-JUN-2010 22:47:32 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Making IPM Project Location to Optional Rig Name Selection" %>
<% '      33*[1372627] 30-JUN-2010 12:47:24 (GMT) APrakash6 %>
<% '         "SWIFT # 2488995 - Add 2 items to the hazard category dropdown in QUEST" %>
<% '      34*[1367024] 01-JUL-2010 05:54:24 (GMT) SKadam3 %>
<% '         "SWIFT #2474157 - D&M SQ tab (Re-Work)" %>
<% '      35*[1377427] 13-JUL-2010 05:09:26 (GMT) APrakash6 %>
<% '         "SWIFT #2467716 - Fix for the issue reported by Roberto" %>
<% '    33A1 [1381262] 21-JUL-2010 17:30:41 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Added Help Video for Rig Names Selection in RIR Main Page" %>
<% '      36*[1381417] 23-JUL-2010 13:38:01 (GMT) APrakash6 %>
<% '         "SWIFT # 2467716 - CR to implement links for Created By and Updated By names" %>
<% '      37*[1383408] 28-JUL-2010 14:31:25 (GMT) APrakash6 %>
<% '         "SWIFT # 2467716 - Merging the changes made for revision 33A1" %>
<% '      38*[1389978] 03-SEP-2010 07:37:22 (GMT) APrakash6 %>
<% '         "SWIFT # 2440793 - Error while saving RIR Main Page" %>
<% '      39*[1398286] 16-NOV-2010 12:29:24 (GMT) PMakhija %>
<% '         "Swift#2502389-Cross Scripting issue in RIR module except Reports" %>
<% '      40*[1441026] 11-JAN-2011 17:27:01 (GMT) VGrandhi %>
<% '         "SWIFT #2532665 - Cannot close IPM RIR when Rig NPT > NPT for Hazardous Situation" %>
<% '      41*[1446293] 25-JAN-2011 13:44:29 (GMT) PMakhija %>
<% '         "Release notes for Q2010.07.16" %>
<% '      42*[1450601] 21-FEB-2011 12:04:46 (GMT) PMakhija %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:RIRdsp.asp;171 %>
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
<% '       7*[1266818] 15-SEP-2009 16:26:34 (GMT) NNaik %>
<% '         "2411197 :Expand the 'Site' dropdown with additional items" %>
<% '       8*[1271052] 25-SEP-2009 20:36:20 (GMT) VGrandhi %>
<% '         "SWIFT #2403986 - Well Services SQ Tabs" %>
<% '       9*[1275699] 29-SEP-2009 15:48:26 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      10*[1277927] 02-OCT-2009 17:50:40 (GMT) NNaik %>
<% '         "SWIFT # 2401608 : NPT - Changes to conform to new SQ Std02 -Message creation, to fix the bugs" %>
<% '      11*[1279417] 05-OCT-2009 22:57:37 (GMT) VGrandhi %>
<% '         "SWIFT #2403986 - Fixes/Additions from Stephan in Well Services SQ Tabs" %>
<% '      12*[1277363] 07-OCT-2009 14:50:18 (GMT) MAnthony2 %>
<% '         "SWIFT# 2401608 NPT Change Request" %>
<% '      13*[1282923] 16-OCT-2009 21:28:41 (GMT) MAnthony2 %>
<% '         "SWIFT # 2409503 - Possible now to RELOCATE RIR outside topnode" %>
<% '      14*[1287289] 28-OCT-2009 16:23:18 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab" %>
<% '      15*[1289354] 03-NOV-2009 05:23:47 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      16*[1290343] 04-NOV-2009 13:41:57 (GMT) SKadam3 %>
<% '         "SWIFT #2438856 - Remove Severity Check for given Segments (as matrix is different)" %>
<% '      17*[1292460] 17-NOV-2009 17:42:23 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - Changes for Section 3 for Rig Related." %>
<% '      18*[1295941] 19-NOV-2009 16:46:01 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - 11-16-2009 Jon Changes" %>
<% '      19*[1298532] 24-NOV-2009 17:49:46 (GMT) VGrandhi %>
<% '         "SWIFT #2430251 - IPM SQ Event Tab - Changed closing message" %>
<% '      20*[1303867] 23-DEC-2009 17:31:09 (GMT) VGrandhi %>
<% '         "SWIFT #2448303 - Develop EMS SQ Tab" %>
<% '      21*[1303811] 23-DEC-2009 17:53:32 (GMT) VGrandhi %>
<% '         "2447849 - Hide SQ Questions for IPM HSE events" %>
<% '      22*[1333698] 16-MAR-2010 05:59:52 (GMT) DMohanty %>
<% '         "Swift #2463864 - Data Gathering informations disappearing from RIR investigation tab" %>
<% '      23*[1340591] 29-MAR-2010 19:31:28 (GMT) DMohanty %>
<% '         "Swift # 2463864 - Data Gathering informations disappearing from RIR investigation tab" %>
<% '      24*[1340715] 29-MAR-2010 21:45:17 (GMT) DMohanty %>
<% '         "Swift # 2471699 -IPM Project / SQ Details tab" %>
<% '      25*[1342944] 07-APR-2010 11:33:24 (GMT) NNaik %>
<% '         "SWIFT # 2468497 - Add Pop-up tp Quality Loss selections" %>
<% '      26*[1348951] 21-APR-2010 10:25:19 (GMT) MAnthony2 %>
<% '         "2476332 - Selection Issue for NPT Loss- RIR Page" %>
<% '      27*[1349607] 05-MAY-2010 13:23:19 (GMT) APrakash6 %>
<% '         "SWIFT #2467716 - Add the GIN# next to the creator name" %>
<% '      28*[1345547] 25-MAY-2010 16:06:55 (GMT) DMohanty %>
<% '         "Swift # 2474157 - D&M SQ tab Development" %>
<% '      29*[1359004] 04-JUN-2010 21:43:45 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Develop interface to extract RIG NAME and associated info from CRM" %>
<% '      30*[1366377] 10-JUN-2010 06:51:40 (GMT) DMohanty %>
<% '         "Swift # 2474157 - D&M SQ tab Development (Reported error rectification)" %>
<% '      31*[1368943] 21-JUN-2010 14:44:07 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Not selected Site Type when we switch to HSE/SQ/Classification" %>
<% '      32*[1371202] 22-JUN-2010 22:47:32 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Making IPM Project Location to Optional Rig Name Selection" %>
<% '      33*[1372627] 30-JUN-2010 12:47:24 (GMT) APrakash6 %>
<% '         "SWIFT # 2488995 - Add 2 items to the hazard category dropdown in QUEST" %>
<% '      34*[1367024] 01-JUL-2010 05:54:24 (GMT) SKadam3 %>
<% '         "SWIFT #2474157 - D&M SQ tab (Re-Work)" %>
<% '      35*[1377427] 13-JUL-2010 05:09:26 (GMT) APrakash6 %>
<% '         "SWIFT #2467716 - Fix for the issue reported by Roberto" %>
<% '    33A1 [1381262] 21-JUL-2010 17:30:41 (GMT) SVadla %>
<% '         "SWIFT #2463187 - Added Help Video for Rig Names Selection in RIR Main Page" %>
<% '      36*[1381417] 23-JUL-2010 13:38:01 (GMT) APrakash6 %>
<% '         "SWIFT # 2467716 - CR to implement links for Created By and Updated By names" %>
<% '      37*[1383408] 28-JUL-2010 14:31:25 (GMT) APrakash6 %>
<% '         "SWIFT # 2467716 - Merging the changes made for revision 33A1" %>
<% '      38*[1389978] 03-SEP-2010 07:37:22 (GMT) APrakash6 %>
<% '         "SWIFT # 2440793 - Error while saving RIR Main Page" %>
<% '      39*[1398286] 16-NOV-2010 12:29:24 (GMT) PMakhija %>
<% '         "Swift#2502389-Cross Scripting issue in RIR module except Reports" %>
<% '      40*[1441026] 11-JAN-2011 17:27:01 (GMT) VGrandhi %>
<% '         "SWIFT #2532665 - Cannot close IPM RIR when Rig NPT > NPT for Hazardous Situation" %>
<% '      41*[1446293] 25-JAN-2011 13:44:29 (GMT) PMakhija %>
<% '         "Release notes for Q2010.07.16" %>
<% '      42*[1450601] 21-FEB-2011 12:04:46 (GMT) PMakhija %>
<% '         "Swift#2540095-MI Swaco Tab" %>
<% '      43*[1458214] 16-MAR-2011 08:24:48 (GMT) PMakhija %>
<% '         "Swift#2541978-Create a new tab 'HOC' for M-I Swaco segment" %>
<% '      44*[1476152] 26-APR-2011 12:01:38 (GMT) MPatil2 %>
<% '         "SWIFT #2554297 - Opening Main page of module flashes a box on upper Left containing 'False'." %>
<% '      45*[1474070] 12-MAY-2011 06:04:13 (GMT) PMakhija %>
<% '         "Swift#2553616-Replicate WL SQ Detail tab and FTL/Rite plugin functionality for the new segment 'Production Wireline NAM' ID#122" %>
<% '      46*[1478310] 18-MAY-2011 09:59:12 (GMT) MPatil2 %>
<% '         "SWIFT #2555574 - Renaming the TCC tab to GRC and add functionalities for Radiation/Explosive/Chemical" %>
<% '      47*[1499439] 30-JUN-2011 12:20:14 (GMT) AGazi %>
<% '         "SWIFT #2525214 - Discrepancy in RIR classification dropdown values" %>
<% '      48*[1499491] 08-AUG-2011 13:37:21 (GMT) AGazi %>
<% '         "SWIFT #2531103 - Accidents can be closed with incomplete/red tabs" %>
<% '      49*[1514329] 16-AUG-2011 13:44:23 (GMT) AGazi %>
<% '         "SWIFT #2525214 - Discrepancy in RIR classification dropdown values" %>
<% '      50*[1516872] 24-AUG-2011 06:19:33 (GMT) AGazi %>
<% '         "SWIFT #2531103 - Accidents can be closed with incomplete/red tabs" %>
<% '      51*[1519315] 02-SEP-2011 12:30:56 (GMT) AGazi %>
<% '         "SWIFT #2531103 - Accidents can be closed with incomplete/red tabs" %>
<% '      52*[1491173] 07-SEP-2011 13:26:10 (GMT) PMakhija %>
<% '         "SWIFT #2574348 - D&amp;M Tab related changes" %>
<% '      53*[1521781] 12-SEP-2011 09:55:42 (GMT) AGazi %>
<% '         "SWIFT #2531103 - Accidents can be closed with incomplete/red tabs" %>
<% '      54*[1523632] 13-SEP-2011 09:09:09 (GMT) AGazi %>
<% '         "SWIFT #2531103 - Accidents can be closed with incomplete/red tabs" %>
<% '      55*[1522276] 14-SEP-2011 12:26:14 (GMT) KIrani %>
<% '         "SWIFT #2568237 - Create owner interface to control the DEPARTMENT dropdown" %>
<% '      56*[1524355] 14-SEP-2011 12:50:01 (GMT) PMakhija %>
<% '         "SWIFT #2574348 - D&M Tab related changes" %>
<% '      57*[1527073] 27-SEP-2011 08:03:26 (GMT) PMakhija %>
<% '         "SWIFT #2574348 - D&M Tab related changes" %>
<% '      58*[1531489] 05-OCT-2011 12:49:14 (GMT) KIrani %>
<% '         "SWIFT #2568237 - Create owner interface to control the DEPARTMENT dropdown" %>
<% '      59*[1533591] 14-OCT-2011 10:56:12 (GMT) KIrani %>
<% '         "SWIFT #2568237 - Create owner interface to control the DEPARTMENT dropdown" %>
<% '      60*[1534778] 14-OCT-2011 11:16:56 (GMT) KIrani %>
<% '         "SWIFT #2568237 - Create owner interface to control the DEPARTMENT dropdown" %>
<% '      61*[1529537] 14-OCT-2011 12:05:24 (GMT) PMakhija %>
<% '         "SWIFT #2574362 - Path Finder TAB" %>
<% '      62*[1533099] 31-OCT-2011 14:49:39 (GMT) MPatil2 %>
<% '         "SWIFT #2588673 - Enable multisegment reporting via Involved Segments/Functions tab" %>
<% '      63*[1543478] 08-NOV-2011 12:42:41 (GMT) MPatil2 %>
<% '         "SWIFT #2588673 - Enable multisegment reporting via Involved Segments/Functions tab" %>
<% '      64*[1551113] 28-NOV-2011 12:15:32 (GMT) PMakhija %>
<% '         "SWIFT #2574362 - Path Finder TAB" %>
<% '      65*[1546388] 01-DEC-2011 14:15:44 (GMT) MPatil2 %>
<% '         "SWIFT #2594250 - GSS ML SQ Detail tab" %>
<% '      66*[1556861] 12-DEC-2011 05:58:59 (GMT) MPatil2 %>
<% '         "SWIFT #2594250 - GSS ML SQ Detail tab" %>
<% '      67*[1550566] 12-DEC-2011 09:37:47 (GMT) AGazi %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab" %>
<% '      68*[1557658] 12-DEC-2011 14:57:12 (GMT) KIrani %>
<% '         "SWIFT #2594250 - GSS ML SQ Detail tab" %>
<% '      69*[1549216] 20-DEC-2011 09:51:19 (GMT) AGazi %>
<% '         "SWIFT # 2591239  - IPM/SPM split - adapting QUEST IPM flag;" %>
<% '      70*[1562777] 23-DEC-2011 07:12:39 (GMT) AGazi %>
<% '         "SWIFT #2591239 - IPM/SPM split - adapting QUEST IPM flag" %>
<% '      71*[1550980] 06-JAN-2012 07:47:15 (GMT) PMakhija %>
<% '         "SWIFT #2594603 - Feature: Segment specific alert for D&amp;M SQ RIR" %>
<% '      72*[1570284] 12-JAN-2012 00:01:01 (GMT) APrakash6 %>
<% '         "SWIFT #2599766 - Modifications fro Multi-segment tab^HIDE MS FOR PROD" %>
<% '      73*[1572090] 19-JAN-2012 23:50:56 (GMT) APrakash6 %>
<% '         "SWIFT #2613086 - Scripts to help with WIS creation..." %>
<% '      74*[1578585] 07-FEB-2012 07:03:40 (GMT) PMakhija %>
<% '         "SWIFT #2594603 - Feature: Segment specific alert for D&amp;M SQ RIR" %>
<% '      75*[1588034] 27-FEB-2012 10:50:23 (GMT) MPatil2 %>
<% '         "SWIFT #2608547 - Multi Segment - Phase 2" %>
<% '      76*[1589026] 29-FEB-2012 10:45:11 (GMT) MPatil2 %>
<% '         "SWIFT #2608547 - Multi Segment - Phase 2" %>
<% '      77*[1591845] 08-MAR-2012 17:21:13 (GMT) APrakash6 %>
<% '         "SWIFT #2627077 -  Feature:Prompt for Other Segments/Functions/Organizations Involved(Text Change)" %>
<% '      78*[1594071] 29-MAR-2012 06:50:14 (GMT) MPatil2 %>
<% '         "SWIFT #2622550 - Changes in M-I SWACO Quality detail tab" %>
<% '      79*[1598959] 30-MAR-2012 09:41:14 (GMT) PMakhija %>
<% '         "SWIFT #2602391 - Changes to D&M SQ Direct Entry RIR's" %>
<% '      80*[1610532] 25-APR-2012 12:57:53 (GMT) NNaik %>
<% '         "SWIFT #2602391 - Changes to D&M SQ Direct Entry RIR's" %>
<% '      81*[1592763] 25-APR-2012 14:17:14 (GMT) NNaik %>
<% '         "SWIFT #2627077 - Feature: Prompt for 'Other Segments/Functions/Organizations Involved" %>
<% '      82*[1612102] 30-APR-2012 11:06:20 (GMT) PMakhija %>
<% '         "SWIFT #2602391 - Changes to D&M SQ Direct Entry RIR's" %>
<% '      83*[1612104] 30-APR-2012 11:11:07 (GMT) MPatil2 %>
<% '         "SWIFT #2627077 - Feature: Prompt for 'Other Segments/Functions/Organizations Involved" %>
<% '      84*[1613203] 03-MAY-2012 13:33:44 (GMT) MPatil2 %>
<% '         "SWIFT #2639628 - Fix required for Job+RunNo" %>
<% '      85*[1619604] 22-MAY-2012 09:07:40 (GMT) VGrandhi %>
<% '         "SWIFT #2639057 - Feature: D&amp;M SQ Tab : New Drop-downs added for Failure Department &amp; Category" %>
<% '      86*[1619376] 22-JUN-2012 11:59:59 (GMT) ATuscano %>
<% '         "SWIFT #2638011 - Feature: Added Character Limitation Indicator and Character Counter for all Multiline Textboxes" %>
<% '      87*[1648151] 24-JUL-2012 22:16:25 (GMT) VGrandhi %>
<% '         "SWIFT #2657598 - Feature: Add new column SSR-ID in SQ Categories System Report." %>
<% '      88*[1633354] 07-AUG-2012 16:27:13 (GMT) APrakash6 %>
<% '         "SWIFT #2649311 - Feature: Quality SQ RIR enforce NPT &amp; Red Money at creation for CMS events." %>
<% '      89*[1660545] 21-AUG-2012 21:01:49 (GMT) VGrandhi %>
<% '         "SWIFT #2663845 - Feature: Red money cannot be a negative value" %>
<% '      90*[1651487] 24-AUG-2012 07:53:07 (GMT) ATuscano %>
<% '         "SWIFT #2658497 - Feature:   D&amp;M Equipment Summary has new Tool Failure column" %>
<% '      91*[1674663] 09-NOV-2012 11:47:54 (GMT) MPatil2 %>
<% '         "SWIFT #2658497 - Feature:   D&M Equipment Summary has new Tool Failure column" %>
<% '      92*[1693437] 16-NOV-2012 08:36:13 (GMT) MSaxena2 %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      93*[1690601] 12-DEC-2012 19:40:48 (GMT) APrakash6 %>
<% '         "SWIFT #2680271 - Feature: RIRs can record Well Barrier Events in SQ and HSE" %>
<% '      94*[1695958] 19-DEC-2012 14:51:00 (GMT) ATuscano %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      95*[1713926] 03-JAN-2013 15:57:14 (GMT) VGrandhi %>
<% '         "SWIFT #2686021 - FIX: D&amp;M Equipment Summary in D&amp;M Details Tab" %>
<% '      96*[1724644] 31-JAN-2013 12:18:03 (GMT) ATuscano %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released." %>
<% '      97*[1727675] 09-FEB-2013 00:17:47 (GMT) APrakash6 %>
<% '         "SWIFT #2670092 - Feature: ALS SQ Details Tab released.^fix" %>
<% '      98*[1727030] 12-FEB-2013 23:50:40 (GMT) ATuscano %>
<% '         "SWIFT #2691249 - Feature: DNM Failed Process-Equipment tab upgraded" %>
<% '      99*[1727754] 13-FEB-2013 00:04:18 (GMT) APrakash6 %>
<% '         "SWIFT #2692030 - Feature: Path Finder rolled into DNM Segment as special sub segment." %>
<% '     100*[1730861] 19-FEB-2013 23:38:26 (GMT) VGrandhi %>
<% '         "SWIFT #2691249 - Feature: DNM Failed Process-Equipment tab upgraded" %>
<% '     101*[1732469] 22-FEB-2013 22:27:18 (GMT) VGrandhi %>
<% '         "SWIFT #2692030 - Feature: Path Finder rolled into DNM Segment as special sub segment." %>
<% '     102*[1707484] 04-MAR-2013 04:12:24 (GMT) APrakash6 %>
<% '         "SWIFT #2683025 - FEATURE: Severity Escalation option for SQ Incidents" %>
<% '     103*[1730036] 18-MAR-2013 14:35:40 (GMT) MPatil2 %>
<% '         "SWIFT #2686023 - FEATURE: WIS SQ Detail Tab" %>
<% '     104*[1740918] 19-MAR-2013 11:58:41 (GMT) MPatil2 %>
<% '         "SWIFT #2686023 - FEATURE: WIS SQ Detail Tab" %>
<% '     105*[1740925] 19-MAR-2013 12:52:06 (GMT) MPatil2 %>
<% '         "SWIFT #2693931 - Move Production Wireline (E&amp;P SPWL) to WIS/CWS id 9190" %>
<% '     106*[1746162] 02-APR-2013 12:25:56 (GMT) ATuscano %>
<% '         "SWIFT #2691249 - Feature: DNM Failed Process-Equipment tab upgraded" %>
<% '     107*[1751042] 17-APR-2013 15:19:58 (GMT) APrakash6 %>
<% '         "SWIFT #2699201 - Training video clips to be attached" %>
<% '     108*[1753165] 07-MAY-2013 14:03:35 (GMT) MPatil2 %>
<% '         "SWIFT #2696315 - Modify Safety Loss categories and its format" %>
<% '     109*[1758014] 04-JUN-2013 08:24:55 (GMT) MPatil2 %>
<% '         "SWIFT #2698673 - Feature: WL SQ Tab validates Incident NPT Vs Total NPT Incident Severity Vs RIR Severity" %>
<% '     110*[1763883] 10-JUN-2013 12:29:29 (GMT) MPatil2 %>
<% '         "SWIFT #2698673 - Feature: WL SQ Tab validates Incident NPT Vs Total NPT Incident Severity Vs RIR Severity" %>
<% '     111*[1772304] 02-AUG-2013 12:02:35 (GMT) ATuscano %>
<% '         "SWIFT #2706924 - Well Barrier tab update - first and secondary envelope integrity" %>
<% '     112*[1790507] 17-SEP-2013 15:20:07 (GMT) MPatil2 %>
<% '         "SWIFT #2705555 - Feature: DNM and PF UI upgrades -Part 2" %>
<% '     113*[1791196] 19-SEP-2013 12:38:55 (GMT) MPatil2 %>
<% '         "SWIFT #2705555 - Feature: DNM and PF UI upgrades -Part 2" %>
<% '     114*[1783405] 04-OCT-2013 10:16:38 (GMT) ATuscano %>
<% '         "SWIFT #2713511 - Need to extract ASL Data (Contractors) from the WebService and UI Changes." %>
<% '     115*[1796322] 14-OCT-2013 09:00:55 (GMT) ATuscano %>
<% '         "SWIFT #2713511 - Need to extract ASL Data (Contractors) from the WebService and UI Changes." %>
<% '     116*[1795427] 14-OCT-2013 10:08:19 (GMT) MPatil2 %>
<% '         "ENH009592  CHANGE: Guest User to be renamed as Basic User" %>
<% '     117*[1798706] 28-OCT-2013 08:54:12 (GMT) ATuscano %>
<% '         "SWIFT #2713511 - Need to extract ASL Data (Contractors) from the WebService and UI Changes." %>
<% '     118*[1795807] 06-NOV-2013 15:35:22 (GMT) MPatil2 %>
<% '         "ENH009599:  WPS SQ CMSL to be only closed after completion of WPS SQ Details tab (Requested)" %>
<% '     119*[1810307] 12-DEC-2013 05:32:38 (GMT) ATuscano %>
<% '         "ENH009582 - Feature: Adopted ASL Master Data in QUEST" %>
<% '     120*[1811846] 29-JAN-2014 10:03:02 (GMT) MPatil2 %>
<% '         "Fix D&M Tab Issue for LIH Incident category" %>
<% '     121*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '     122*[1834769] 20-MAY-2014 12:05:20 (GMT) VGrandhi %>
<% '         "ENH022155:  ENH: Name of Environment Loss categories to be changed (HSE RIR)" %>
<% '     123*[1845452] 04-JUL-2014 09:21:32 (GMT) Rbhalave %>
<% '         "NFT029203 - D&M phase 10 - part 1 labeling changes for D&M SQ tab" %>
<% '     124*[1854239] 20-AUG-2014 07:30:48 (GMT) Rbhalave %>
<% '         "Time Loss Anomaly" %>
<% '     125*[1863565] 09-OCT-2014 10:45:28 (GMT) Rbhalave %>
<% '         "NFT039565 New SPS indicator on the use of SWI" %>
<% '     126*[1862023] 09-OCT-2014 14:18:39 (GMT) Rbhalave %>
<% '         "NFT039565 New SPS indicator on the use of SWI" %>
<% '     127*[1864350] 14-OCT-2014 05:24:09 (GMT) Rbhalave %>
<% '         "D&M revert back from production." %>
<% '     128*[1865353] 17-OCT-2014 11:47:21 (GMT) Rbhalave %>
<% '         "NFT039565 New SPS indicator on the use of SWI - Patch" %>
<% '     129*[1867842] 13-NOV-2014 11:40:26 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '     130*[1870634] 17-NOV-2014 13:40:52 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '     131*[1870655] 17-NOV-2014 15:03:17 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '     132*[1872365] 15-DEC-2014 14:06:37 (GMT) VSharma16 %>
<% '         "ENH046924 - D&M and Pathfinder - Complete merge of details tab and related tables" %>
<% '     133*[1876672] 08-JAN-2015 12:16:34 (GMT) VSharma16 %>
<% '         "DNM merge pathfinder" %>
<% '     134*[1879128] 08-JAN-2015 13:30:17 (GMT) Rbhalave %>
<% '         "ENH052384: Change of text SPS - to 'Do It Right" %>
<% '     135*[1883751] 18-FEB-2015 15:10:07 (GMT) VSharma16 %>
<% '         "ENH053415 - Addition of PTEC project acknowledgement on HSE and SQ RIR and O/I reports" %>
<% '     136*[1889513] 16-MAR-2015 12:49:12 (GMT) SChaudhari %>
<% '         "ENH059674:  Javascript forcing of PTEC selection - RIR and Obs/Int" %>
<% '     137*[1889337] 19-MAR-2015 08:37:56 (GMT) VSharma16 %>
<% '         "ENH054251-D&M Tab updates: Phase11" %>
<% '     138*[1903829] 08-JUN-2015 10:50:32 (GMT) Rbhalave %>
<% '         "ENH055492  SPS request - addition of SPS Process categories to SQ RIR report - all Segments" %>
<% '     139*[1908933] 30-JUN-2015 12:11:14 (GMT) Rbhalave %>
<% '         "ENH060157 Removal of M-I Swaco HSE RIR feature page HOC" %>
<% '     140*[1911892] 16-JUL-2015 13:21:58 (GMT) Rbhalave %>
<% '         "ENH075535  - SQ SPS Process Phase 2" %>
<% '     141*[1911595] 17-JUL-2015 10:27:05 (GMT) SChaudhari %>
<% '         "ENH077170 - SAXON - Operations at Time of Event (Category and Sub-Category)" %>
<% '     142*[1912660] 17-JUL-2015 11:15:26 (GMT) SChaudhari %>
<% '         "NFT075965 -D&M Phase 12 - Part 1" %>
<% '     143*[1912868] 21-JUL-2015 12:07:55 (GMT) SChaudhari %>
<% '         "Changes missed while merging file" %>
<% '     144*[1914148] 04-AUG-2015 11:57:03 (GMT) Rbhalave %>
<% '         "post deployment ENH075535  - SQ SPS Process Phase 2" %>
<% '     145*[1922253] 05-OCT-2015 07:56:28 (GMT) SChaudhari %>
<% '         "ENH086534-Operation at the time of event categories selection" %>
<% '     146*[1815665] 05-OCT-2015 10:56:31 (GMT) VSharma16 %>
<% '         "ENH081653:  D&M Phase 12 - Part 2" %>
<% '     147*[1866913] 18-NOV-2015 11:49:08 (GMT) VSharma16 %>
<% '         "ANO091859:D&M - Post SCAT - Root Cause logic changes" %>
<% '     148*[1927564] 20-NOV-2015 12:43:22 (GMT) SChaudhari %>
<% '         "ENH086389 - Saxon - SQ Categories" %>
<% '     149*[1932385] 24-NOV-2015 10:37:52 (GMT) Rbhalave %>
<% '         "NFT087279 - TS segment becoming a forced segment - Rig related flag at sub sub segment level" %>
<% '     150*[1933523] 18-DEC-2015 12:36:08 (GMT) Rbhalave %>
<% '         "NFT087279 - TS segment becoming a forced segment - Rig related flag at sub sub segment level" %>
<% '     151*[1937795] 31-DEC-2015 07:55:55 (GMT) MPatel13 %>
<% '         "ENH092498-TLM Sub-Segments to invoke Tool Parent SQ Tabs" %>
<% '     152*[1939904] 12-JAN-2016 10:09:21 (GMT) MPatel13 %>
<% '         "REMOVING-TLM-ENH092498-TLM Sub-Segments to invoke Tool Parent SQ Tabs" %>
<% '     153*[1940182] 14-JAN-2016 07:42:22 (GMT) MPatel13 %>
<% '         "TLM-Redo-ENH092498-TLM Sub-Segments" %>
<% '     154*[1941773] 16-FEB-2016 13:03:45 (GMT) VSharma16 %>
<% '         "ENH100497: SPS addition of Categories (SUPPORT ITT PROJECT)" %>
<% '     155*[1940746] 18-FEB-2016 14:08:28 (GMT) SChaudhari %>
<% '         "ANO099305 - Rite Account Unit Changes" %>
<% '     156*[1942528] 24-FEB-2016 07:49:23 (GMT) ALanger %>
<% '         "ENH097408 - GSS merge with D&M" %>
<% '     157*[1945291] 29-FEB-2016 13:10:53 (GMT) SChaudhari %>
<% '         "ENH104106 - Logic change related to Execution Error in D&M Tab" %>
<% '     158*[1945290] 04-MAR-2016 13:48:22 (GMT) VSharma16 %>
<% '         "ENH104074:SPWL move to WL - requires SQ Wireline SQ tab to be implemente" %>
<% '     159*[1938043] 31-MAR-2016 12:55:24 (GMT) VSharma16 %>
<% '         "ENH095327-SLIM - Job ID changes" %>
<% '     160*[1949038] 31-MAR-2016 14:18:54 (GMT) ALanger %>
<% '         "ENH095323  SLIM - SQ + HSE RIR report changes to form field (Department)" %>
<% '     161*[1952076] 11-APR-2016 09:48:28 (GMT) SChaudhari %>
<% '         "ENH095335-SLIM - SQ + HSE RIR changes to Site and site name fields" %>
<% '     162*[1948724] 12-APR-2016 07:05:05 (GMT) MPatel13 %>
<% '         "ENH095322-SLIM - SQ RIR report changes to form fields" %>
<% '     163*[1944390] 12-APR-2016 11:19:28 (GMT) VSharma16 %>
<% '         "ENH096140 - Integrated Projects" %>
<% '     164*[1948528] 12-APR-2016 11:41:17 (GMT) VSharma16 %>
<% '         "ANO103787:  SQ RIR SPS Categories userability" %>
<% '     165*[1953103] 12-APR-2016 13:52:20 (GMT) SChaudhari %>
<% '         "ENH095335-SLIM - SQ + HSE RIR changes to Site and site name fields(Revert changes)" %>
<% '     166*[1951386] 12-APR-2016 14:13:20 (GMT) Rbhalave %>
<% '         "ENH101167 - SLIM ROOT CAUSE CLASSIFICATION" %>
<% '     167*[1955774] 13-MAY-2016 09:33:59 (GMT) VSharma16 %>
<% '         "PNM changed files" %>
<% '     168*[1953356] 13-MAY-2016 10:07:23 (GMT) Rbhalave %>
<% '         "ENH101167 SLIM Root cause classification" %>
<% '     169*[1878140] 24-JUN-2016 06:56:22 (GMT) VSharma16 %>
<% '         "NFT101068- RIR locking to prevent data integrity issues" %>
<% '     170*[1897136] 24-JUN-2016 07:39:46 (GMT) SChaudhari %>
<% '         "ENH115371 - <<MM>> Change TLM SQ Tab Selection to Sub Sub-Segment" %>
<% '     171*[1963919] 30-JUN-2016 13:29:19 (GMT) VSharma16 %>
<% '         "NFT101068 SQ Report Locking - Discussion(change)" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:RIRdsp.asp;171 %>
