<%@ Language=VBScript %>
<%option explicit
'   Date(dd-MMM-yyyy)		Prepared/Modified By         Comments
'   7-May-2014                Varun Sharma                Modified - Changed for NFT014129 NPT/CMSL/TNCR data historical capture
'   26-AUG-2014                Varun Sharma                ENH035526 – Malaria Reporting 
'   31-OCT-2014                Varun Sharma                ENH044752  HSE locking of lagging indicators - safety net - key to unlock
%>
<!-- #INCLUDE FILE="../Inc_Functions.asp"-->
<!-- #INCLUDE FILE="RIR_Include.asp"-->
<!-- #INCLUDE FILE="../Inc/Inc_Security.asp"-->
<%checktimeout()
  If IsIE then Response.Expires = -1

Dim dtRptDate, sRptNo, RS1, conn, sKey, bDisplayInjury, bHealth, bInjury, sWhere , bEventDate
Dim iCt, iCtr, lOrgNo, bNew, cn, RS, iCntr, bMoreRows,EditMode,HSEGuest
Dim iKCost1, iKCost2, iKCost3, iKCost4, dtTemp, sHref,LockOutcome
Dim sTemp, sTemp2, sTemp3, iTemp, iQPID,HSESeverityval,VarJquery
Dim ACLDefined
lOrgNo = Request.QueryString("OrgNo")
dtRptDate = Request.QueryString("rptDate")
iQPID = Request.QueryString("QPID")

sHSEWarningText="*** Warning this report will be locked from editing Key HSE Fields in " & LockCount & " day(s) ***"

Set cn = GetNewCN()
Set RS = Server.CreateObject("ADODB.Recordset")
VarJquery = fncD_CommonURL("Jquery")
sKey = GetReportKey(iQPID, lOrgNo, dtRptDate, "R", cn)
ACLDefined=CheckAccess("R", iQPID, lOrgNo, dtRptDate, cn)
	
sTemp = "SELECT * FROM tblRIRP1 With (NOLOCK) WHERE QID=" & SafeNum(iQPID) 
Set RS1 = cn.Execute(sTemp)
HSESeverityval=RS1("HSESeverity")
bDisplayInjury=IsPers(RS1)And RS1("Class")=1 And RS1("HSE")
HSEGuest=isGuest() And RS1("HSE")
bInjury = RS1("LossCat_A1")
bHealth = RS1("LossCat_A2") OR RS1("LossCat_A3")
bEventDate = RS1("EventDateTime")
editmode=true

if bInjury then 
	if chkEditPersLoss("Inj") or (lcase(RS1("CreateUID")) = lcase(Session("UID"))) then EditMode=true else EditMode=false
End if
'If it's an illness, then we need to make sure the person logged in has access to personnel data for illnesses
If bHealth Then
	If chkEditPersLoss("Med") OR (lcase(RS1("CreateUID")) = lcase(Session("UID"))) Then
		EditMode=True
	Else
		AccessDenied "Access only available to authorized users and report creator."
	End If
End If
%>
 
<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="../style/QUEST.css" type="text/css">
	<script LANGUAGE="JavaScript1.2" SRC="../Calendar1-82.js"></script>
	<script src="<%=VarJquery%>"></script>
	<script language="JavaScript">
		<!--
    var LRDays = 0;
    var OutComeSum = 0;

    function validateSeverity(ctr) {
        var D1 = eval('window.document.frmPers.txtDays' + ctr + '.value');
        var D2 = eval('window.document.frmPers.txtRDays' + ctr + '.value');
        var o1 = eval('window.document.frmPers.outcome' + ctr);


        if ((isNaN(parseInt(D1)) & D1 != "") | (isNaN(parseInt(D2)) & D2 != "")) {
            alert('Please enter an integer value for the days lost and days of reduced work.');
        } else {
            var days = parseInt(D1) + parseInt(D2);  //get numeric value of sum of D1 and D2
            OutComeSum = days;

            if (days > 0) {

                if (days >= 180) {
                    //must be at least a Major
                    if (<%= rs1("HSESeverity") %> < 3) {
                        //but it isn't
                        alert('An injury accident with ' + days + ' days of lost/restricted time will automatically be converted to a Major if you save.');


                    }
                    o1.value = "P";
                    showMessage();
                } else if (days >= 1) {
                    //must be at least a Serious
                    if (<%= rs1("HSESeverity") %> < 2) {
                        //but it isn't
                        alert('An injury accident with ' + days + ' days of lost/restricted time will automatically be converted to a Serious if you save.');

                    }
                    showMessage();
                }
            }
        }
    }

    function showMessage() {
        var outcome = false;
        var totSum = false;
        $('.outcome').each(function (i, obj) {
            //test
            if (obj.value == "P") {
                outcome = true;
            }
        });
        var sum = 0;
        var totalEle = $(".cltxtDays").length;

        for (i = 0; i < totalEle / 2; i++) {

            var t1 = parseInt($(".cltxtDaysSum" + i).val());
            var t2 = parseInt($(".cltxtRDaysSum" + i).val());
            sum = t1 + t2;

            if (sum >= 180) { totSum = true; }
        }
        var a = document.getElementById('PImsg');
        if (totSum == true || outcome == true) { a.style.display = 'block'; }
        else {
            a.style.display = 'none';
        }
    }


    $(document).ready(function () {
        showMessage();
        var prev_val;
        $('.outcome').focus(function () {
            prev_val = $(this).val();
        }).change(function () {
            $(this).unbind('focus');
            var name = $(this).attr("name");
            var ret = name.replace('outcome', '');
            var t1 = parseInt($(".cltxtDaysSum" + ret).val());
            var t2 = parseInt($(".cltxtRDaysSum" + ret).val());
            var sum = t1 + t2;
            if (sum >= 180) {
                var v = $(this).val();
                var t = $(this).find("option:selected").text();

                if (v == "P" || v == "X") {
                    return true;
                }
                else {
                    alert("If Total Lost Days>=180 then '" + t.trim() + "' is not allowed as outcome.");
                    $(this).val(prev_val);
                    $(this).bind('focus');
                    return false;
                }
            }
        });
    });









    function extTrim(txt) {
        //Like trim, but trims CR, LF, TAB, and SPACE

        var trimChars, startPos, stopPos, foundText, idx

        trimChars = " \t\n\r";
        startPos = -1;
        stopPos = -1;

        for (idx = 0; idx < txt.length; idx++) {
            if (trimChars.indexOf(txt.charAt(idx)) == -1) {
                if (!foundText) startPos = idx
                foundText = true
                stopPos = idx
            }
        }

        if (startPos != -1)
            return txt.substr(startPos, stopPos - startPos + 1);
        return ""
    }
    function OnSLBClick(ctr) {
        var OC, TY, Obj
        OC = eval('document.frmPers.outcome' + ctr) ? eval('document.frmPers.outcome' + ctr + '.options[document.frmPers.outcome' + ctr + '.selectedIndex].value') : 0;
        TY = eval('document.frmPers.txtType' + ctr) ? eval('document.frmPers.txtType' + ctr + '.options[document.frmPers.txtType' + ctr + '.selectedIndex].value') : 0;
        if (!(OC == 'X' && TY == '3')) {
            Obj = eval('document.frmPers.SlbCon' + ctr);
            Obj.checked = false;
            alert('To Make this fatality to SLB Inv/Concerned only possible \n When the "Employee, Contractor Third Party" type is Third Party \n and the "Outcome" is Fatality');
        }
    }


    function MalariaAlert(selectList) {

        var v = selectList.options[selectList.selectedIndex].value;
        if (v == 'R') {

            open("../Utils/Malaria.asp?flag=1", "searchCRMClient", "height=520,width=550,channelmode=0,dependent=0,directories=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,toolbar=0,alwaysRaised=1")
        }
    }


    function prepSubmit() {
        var ctr = 0;
        var LD, RD, OC, SC, TY, Days, IL, PN, IJ, pct, returnWorkDate;
        var errorsheader = '';
        var errorsmsg = '';
        var frm = document.frmPers;
        pct = 0
        errorsheader += '_______________________________________________\n\n';
        errorsheader += 'The form was not saved because of the following error(s).\n';
        errorsheader += '_______________________________________________\n\n';



        for (ctr = 1; typeof eval('frm.txtName' + ctr) != 'undefined'; ctr++) {
            Days = 0
            PN = eval('frm.txtName' + ctr) ? eval('frm.txtName' + ctr + '.value') : '';
            TY = eval('frm.txtType' + ctr) ? eval('frm.txtType' + ctr + '.options[frm.txtType' + ctr + '.selectedIndex].value') : 0;
            OC = eval('frm.outcome' + ctr) ? eval('frm.outcome' + ctr + '.options[frm.outcome' + ctr + '.selectedIndex].value') : '';
            IJ = eval('frm.txtInj' + ctr) ? eval('frm.txtInj' + ctr + '.options[frm.txtInj' + ctr + '.selectedIndex].value') : '';
            IL = eval('frm.txtInjLoc' + ctr) ? eval('frm.txtInjLoc' + ctr + '.options[frm.txtInjLoc' + ctr + '.selectedIndex].value') : '';
            LD = eval('frm.txtDays' + ctr) ? eval('frm.txtDays' + ctr + '.value') : 0;
            if (extTrim(PN) == 'Medically Confidential') PN = ''
            //alert(PN +','+TY+','+IJ+','+IL+','+OC)
            if ((extTrim(PN) != '') || (TY > 0) || (IJ != '') || (IL != '') || (OC != '')) {
                if ((extTrim(PN) == '') && (extTrim(eval('frm.txtName' + ctr + '.value')) != 'Medically Confidential')) errorsmsg += 'No name entered for person ' + ctr + '\n';
                if (TY == 0) errorsmsg += 'Person ' + ctr + ': Employee, Company or 3rd Pty is not selected.\n';
                <% If bDisplayInjury Then %>
                if (IJ == 0) errorsmsg += 'Person ' + ctr + ': Injury Information is not selected.\n';
                if (IL == 0) errorsmsg += 'Person ' + ctr + ': Injury Body parts affected is not selected.\n'; 
                <% If LockCount <> 0 Then %>
                if (OC == 0) errorsmsg += 'Person ' + ctr + ': Outcome entered is not selected.\n'; 
                <% end if%>
                if (LD == 0 && OC == 'P') errorsmsg += 'Person ' + ctr + ': Days lost is not entered.\n';
                <% end if%>
                    pct = pct + 1
            }

            if (extTrim(eval('frm.txtName' + ctr + '.value')) != '') {
                if ((extTrim(PN) != '') || (TY > 0) || (IJ != '') || (IL != '') || (OC != '')) {
                    RD = eval('frm.txtRDays' + ctr) && eval('frm.txtRDays' + ctr + '.value') ? eval('frm.txtRDays' + ctr + '.value') : 0;
                    returnWorkDate = eval('frm.txtReturnWorkDate' + ctr) && eval('frm.txtReturnWorkDate' + ctr + '.value') ? new Date(eval('frm.txtReturnWorkDate' + ctr + '.value')) : '';

                    LD = eval('frm.txtDays' + ctr) && eval('frm.txtDays' + ctr + '.value') ? eval('frm.txtDays' + ctr + '.value') : 0;

                    if (RD < 0 || LD < 0) {
                        errorsmsg += 'Person ' + ctr + ': Lost Work Days or Restricted Work Days can not be negative ' + 'Please provide valid input.' + '\n';
                    }

                    if (!returnWorkDate && (OC != 'F' && OC != 'M' && OC != 'X')) {
                        var errMsg = (OC != 'R' && OC != '') ? 'Person ' + ctr + ': Return to Work Date can not be empty for Lost Days or Permanent Impairment, if the person is returning to work' + '\n' : '';
                        if (errMsg)
                            alert(errMsg);
                    }
                    else {
                        if (OC != 'F' && OC != 'M' && OC != 'X') {
                            Days = Days + parseInt(LD) + parseInt(RD)
                            var EventDate = new Date(document.frmPers.txtEventDate.value);
                            var time_difference = returnWorkDate.getTime() - EventDate.getTime();
                            var days_difference = parseInt(time_difference) / (1000 * 60 * 60 * 24) - 1;
                            if (Days != days_difference && (OC != 'F' && OC != 'M' && OC != 'X')) {
                                alert('Person ' + ctr + ': Lost Work Days Calculation incorrect, the duration between Event Date and Return to Work Date is: ' + parseInt(days_difference) + ' days Please update accordingly.' + '\n');
                            }
                        }
                    }
                }
            }
            switch (OC) {
                case 'X':
                    if (LD > 0 | RD > 0) errorsmsg += 'Person ' + ctr + ': Lost and Restricted days should be zero for "Fatality".\n';
                    break;
                case 'L':
                    if (LD == 0) errorsmsg += 'Person ' + ctr + ': Lost Days must be greater than zero if Outcome of "Lost Days" is selected.\n';
                    break;
                case 'R':
                    if (LD > 0) errorsmsg += 'Person ' + ctr + ': Cannot have Lost Days when Outcome of "Restricted Days" is selected.\n';
                    if (RD == 0) errorsmsg += 'Person ' + ctr + ': Restricted Days must be greater than zero if Outcome of "Restricted Days" is selected.\n';
                    break;
                case 'M':
                    if (LD > 0 | RD > 0) errorsmsg += 'Person ' + ctr + ': Lost and Restricted days should be zero for "Medical Treatment Only".\n';
                    break;
                case 'F':
                    if (LD > 0 | RD > 0) errorsmsg += 'Person ' + ctr + ': Lost and Restricted days should be zero for "First Aid Only".\n';
                    break;
                default:
                //they must be doing a near accident/hazardous situation
            }
            if (eval('frm.SlbCon' + ctr)) {
                SC = eval('frm.SlbCon' + ctr);
                if (SC.checked) {
                    if (!(OC == 'X' && TY == 3)) SC.checked = false;
                }
            }
        }

        
        <% If bHealth or bInjury Then %>
        if (pct == 0) {
            errorsmsg += 'Warning - You have not entered any details for the parties that suffered injury or illness.\n\n'
            errorsmsg += 'As a minimum to allow Save to be recorded, enter the names of the people who were injured or became \n'
            errorsmsg += 'ill and the required data in the "Personnel Details" table. In order to Close this report you will also be \n'
            errorsmsg += 'required to enter all required information in the "Injury/Illness Details" table for each person recorded.\n\n'
            errorsmsg += 'If there were no Injuries or Illnesses associated with this Accident Event Report please go to the \n'
            errorsmsg += '"HSE Main" page and remove the check against "Safety Loss>Personnel>Injury"'
        }
        <% End IF %>
        if (errorsmsg != '') {
            alert(errorsheader += errorsmsg);
            return false;
        }
        else {
            if (Days > 0) {
                if (parseFloat(frm.CC_27.value) != frm.CC_27.value) {
                    alert('Please enter valid number in Lost/Restricted Work Days Cost');
                    frm.CC_27.focus()
                    return false;
                }
                CalDefaultCost()
                if (frm.CC_27.value <= 0) {
                    alert('The cost for Lost/Restricted Work Days Cost should be greater than zero\nif there are Lost Days/Restricted work days involved.');
                    frm.CC_27.focus()
                    return false;
                }
            }
            return true;
        }
    }





    function fatalityAlert(selectList) {

        var v = selectList.options[selectList.selectedIndex].value;
        var t = selectList.options[selectList.selectedIndex].text;
        <%if HSEGuest then %>
                    if (v == 'X') {
            alert('Basic Users are restricted from creating Catastrophic Events.\n To create a catastrophic event please contact your Line Manager or QHSE Support.')
            selectList.selectedIndex = 0
        }	
        <% Else %>
            if (v == 'X') alert('You have selected an Outcome of \'Fatality\'.\n\nThis implies Life Loss and will force this incident to be classified as Catastrophic.\n\nA notification will be sent to all Senior Managers in SLB with a QUEST  subscription.\n\nAre you sure about the severity of this event?');		
        <% end if%>
            showMessage();


    }




    function CalDefaultCost() {
        var f = document.frmPers, tmp, msg;
        var LD = 0, RD = 0, cost = 0;

        for (var ctr = 1; typeof eval('f.txtName' + ctr) != 'undefined'; ctr++) {
            tmp = eval('f.txtDays' + ctr) ? eval('f.txtDays' + ctr + '.value') : 0;
            if (tmp == parseFloat(tmp)) LD = LD + parseFloat(tmp);
            tmp = eval('f.txtRDays' + ctr) ? eval('f.txtRDays' + ctr + '.value') : 0;
            if (tmp == parseFloat(tmp)) RD = RD + parseFloat(tmp);
        }

        if (LD + RD > 0) {
            LRDays = LD + RD;
            cost = (LD * 3) + (RD * +2);

            msg = 'Warning - The value entered in Lost/Restricted Work Days Cost is not equal to calculated default value.\n\n'
            msg = msg + 'It is calculated using the following formula ... \n\n'
            msg = msg + 'Lost/Restricted Work Days Cost=(Lost Work Days x 3k$) + (Restricted Work Days x 2k$)\n'
            msg = msg + 'Do you want to replace ' + f.CC_27.value + ' k$ with calculated default value ' + cost + ' k$.'
            if (f.CC_27.value != cost) {
                if (confirm(msg)) f.CC_27.value = cost;
            }
        }
        return false
    }

    //-->
    </script>

</head>
<body MARGINWIDTH=0 MARGINHEIGHT=0 LEFTMARGIN=0 TOPMARGIN=0 RIGHTMARGIN=0>
<%
displaymenubar(RS1)
RS1.Close

RS1.Open "SELECT Count(*) AS RecCount FROM tblRIRPers With (NOLOCK) WHERE QPID=" & SafeNum(iQPID) , cn
iCt = RS1("RecCount")+1
if iCt<2 then iCt=2 
RS1.Close
Set RS1=Nothing
if ACLDefined Then DisplayConfidential() 
%>

<form name="frmPers" method="post"  action="LOSSpers2.asp<%=sKey%>" onsubmit="return prepSubmit()">
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
<%If HSESeverityval <> 1 and bDisplayInjury and  not chkHseLockingMgmt()  Then%>
<table border=0 align=center cellPadding=0 cellSpacing=0 width=100%>
<TR>
		<TD align = center>
		<span id='Warning21' class='LockWarning'><%Response.Write sHSEWarningText%></span>			
		</TD>
	</TR>
</table>
<%END IF%>
<table border=0 cellPadding=0 cellSpacing=0 width=100%>

	

	<tr class=reportheading>
		<TD align=left colspan=3 >
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


<input type="hidden" name="txtCt" value="<%=iCt%>">
<input type="hidden" name="txtRptDate" value="<%=FmtDate(dtRptDate)%>">
<input type="hidden" name="txtEventDate" value="<%=FmtDate(bEventDate)%>">

<input type="hidden" name="bDisplayInjury" value="<%=bDisplayInjury%>">


		<table width="100%" border="1" cellPadding="0" cellSpacing="0">
		
			<tr>
				<td align = center colspan=10 class=boxednote id=styleSmall>
					To add more personnel fill all the lines displayed and click save.<br>
					To delete an entry click on the "Item Number" to the left of the name.
				</td>
			</tr>		
				
			<tr>
				<td colspan="2" valign="top" align="center">
					<b><span id=styleSmall >Name(s)</span></b></td>
				<td valign="top" align="center">
					<b><span id=styleSmall >Seniority<br>Date</span></b></td>		
                <td valign="top" align="center">
					<b><span id=styleSmall >Position<br>Experience</span></b></td>			
				<td valign="top" align="center">
					<b><span id=styleSmall >Employee, Contractor,<br>Third Party</b></span></td>				
				<td valign="top" align="center">
					<b><span id=styleSmall >Job<br>Function</span></b></td>				
				<td valign="top" align="center">
					<b><span id=styleSmall >Hrs since<br>last sleep</span></b></td>				
				<td valign="top" align="center">
					<b><span id=styleSmall >Hrs slept<br>last time</span></b></td>				
				<td valign="top" align="center">
					<b><span id=styleSmall >Hrs on<br>Duty</span></b></td></tr>
			
		<%
		sTemp = "SELECT * FROM tblRIRPers With (NOLOCK) WHERE QPID=" & iQPID & " ORDER BY Seq"
		RS.Open sTemp, cn

		For iCntr = 1 to iCt
				bMoreRows = True
				If RS.EOF or RS.BOF Then bMoreRows = False			
				sTemp = ""%>
			
			<tr>
				<td align="center" width="10">
					<%If bMoreRows and EditMode Then 
						sTemp = "<A href=LOSSpers2.asp" & sKey & "&Delete=1&ID=" & RS("SEQ") & " onclick='return cmdDelete_onclick()'>" & iCntr & "</A>" 
					else
						sTemp= iCntr
					End If
						
				%>		
					<b><span id=styleSmall >&nbsp;<%=sTemp%>&nbsp;</span></b></td>
				<td align="center">
					<%sTemp=""
					If bMoreRows = True Then sTemp = Trim(RS("Name"))%>
						<%If bHealth Then 
							Response.Write  "Medically Confidential"%>
							<input type="hidden" name=txtName<%=iCntr%> value='Medically Confidential'></td>
						<%Else			
							if EditMode=false and sTemp<>"" then sTemp="Confidential"
						%>
							<input type="text" name=txtName<%=iCntr%> value="<%=DisplayQuotes(sTemp)%>" size="15" id=stylesmall></td>
						<%End If%>
				<td align="center">
					<%If bMoreRows = True Then sTemp = Trim(RS("Seniority")) else sTemp=""%>
					<input type="text" name=txtSrDate<%=iCntr%> value='<%=sTemp%>' size="4" id=stylesmall>
					<%=popupCalendar("frmPers.txtSrDate" & iCntr)%>

					</td>
                <td align="center">
					<%If bMoreRows = True Then sTemp2 = Trim(RS("PosExp"))else sTemp2=""%>
					<%=GetExperience("txtPosExp"&iCntr,sTemp2,cn)%>
				</td>

				<td align="center">
					<%If bMoreRows = True Then sTemp2 = Trim(RS("InjuredPartyType"))else sTemp2=""%>
					<%=GetPartyTypes("txtType"&iCntr,sTemp2,cn)%>
				</td>	

                
                <td align="center">
					<%If bMoreRows = True Then sTemp2 = Trim(RS("JobFunction"))else sTemp2=""%>
                    <%if len (sTemp2)>0 and not isnull(sTemp2)then %>

                   	<input type="text" name=txtJobfcn<%=iCntr%> value="<%=DisplayQuotes(sTemp2)%>" readonly  size="15" id=stylesmall>
                 
                    <%else %>
                    <%If bMoreRows = True Then sTemp = Trim(RS("JobFID"))  else sTemp=""%>
                    <%=GetJObFunction("txtJob"&iCntr,sTemp,cn)%>
				</td>
					<%end if %>
         



				<td align='center'>
					<%If bMoreRows = True Then sTemp = Trim(RS("hrsAwake"))  else sTemp=""%>
					<input type="text" name=txtSince<%=iCntr%> value='<%=sTemp%>' size="3" id=stylesmall></td>
				<td align='center'>
					<%If bMoreRows = True Then sTemp = Trim(RS("hrsSlept"))  else sTemp=""%>
					<input type="text" name="txtSleep<%=iCntr%>" value='<%=sTemp%>' size="3" id=stylesmall></td>
				<td align='center'>
					<%If bMoreRows = True Then sTemp = Trim(RS("hrsOnDuty"))  else sTemp=""%>
					<input type="text" name=txtOnDuty<%=iCntr%> value='<%=sTemp%>' size="3" id=stylesmall></td></tr>
				<%
				If bMoreRows then RS.MoveNext 
			Next
			%>
		</table>

<%	if bDisplayInjury Then 
		DisplayInjury(RS) 
		DisplayCost cn, 6, iQPID
	End If
	RS.Close
	
	cn.Close
	Set cn = Nothing
%>

<table width=100%>
	<tr>
			<%if EditMode then %>
			<td align="right" valign="top"><input type="submit" name="cmdSubmit" value="Save Data"></td>
			<% Else%>
			<td class=urgent align=center>THIS PAGE CAN ONLY BE UPDATED BY THE REPORT CREATOR OR AUTHORIZED USER.</td>
			<%End IF%>
	</tr>			
</table>
	
</form>
</body>
</html>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

    function cmdDelete_onclick() {
        var bConfirm = window.confirm('Are you sure you wish to DELETE this record');
        return (bConfirm)
    }

    //-->
</script>

<%
Sub DisplayInjury(RS)

	Dim RSInjury, RSBodyPart, RSOutcome
	
	Set RSInjury = Server.CreateObject("ADODB.Recordset")
	Set RSBodyPart = Server.CreateObject("ADODB.Recordset")
	Set RSOutcome = Server.CreateObject("ADODB.Recordset")
	
	If bInjury AND bHealth Then
		sWhere = " WHERE Type IN ('H','I') "
	ElseIf bInjury Then
		sWhere = " WHERE Type IN ('I') "
	ElseIf bHealth Then
		sWhere = " WHERE Type IN ('H') "
	Else
		sWhere = " WHERE 1=0 "
	End If
	
	RSInjury.Open "SELECT * FROM tlkpInjuryCategories With (NOLOCK) " & sWhere & "order by Code_newOrder",cn
	RSBodyPart.Open  "SELECT * FROM tlkpBodyPartCategories With (NOLOCK) ",cn	


	sWhere = " WHERE 1=0 "
	If bInjury Then
		sWhere = sWhere & " OR Injury = 1"
	End If

	If bHealth Then
		sWhere = sWhere & " OR Health = 1"
	End If
	RSOutcome.open "SELECT * FROM tlkpInjuryOutcome With (NOLOCK) " & sWhere & " ORDER BY CASE InjuryOutcomeID WHEN 'F' THEN 1 WHEN 'M' THEN 2 WHEN 'R' THEN 3 WHEN 'L' THEN 4 WHEN 'P' THEN 5 WHEN 'X' THEN 6 ELSE 7 END", cn
	
	%>
	<br>
	<table border='1' cellPadding='0' cellSpacing='0' width='100%'>
		<tr>
			<td colspan=9 class=boxednote id=styleSmall align=center>Enter details of injuries and body parts affected for each of the persons listed above.</td>
		</tr>
		<tr>
			<td colspan=9 class=urgent id=styleSmall align=center color=red>If a SLB Involved fatality is reported, QUEST will automatically change the incident classification to Catastrophic.
				<span id='PImsg' style='display:none'>
				If an outcome is Permanent Impairment or Total Lost Days>=180, QUEST will require incident severity to be Major</span></td>
		</tr>
		<tr>
			<td colspan=2 class=field id=styleSmall align=center valign=top>Injury / Illness</td>
			<td class=field id=styleSmall align=center valign=top>Body Parts Affected</td>
			<td class=field id=styleSmall align=center valign=top>Lost<BR>Work Days</td>
			<td class=field id=styleSmall align=center valign=top>Return to<BR>Work Date</td>
			<td class=field id=styleSmall align=center valign=top>Restricted<BR>Work Days</td>
			<td class=field id=styleSmall align=center valign=top>Outcome</td>
			<%If isOwner() Then Response.Write "<td class=field id=styleSmall align=center valign=top>SLB Inv/<br>Concerned?</td>"%>
			<td class=field id=styleSmall align=center valign=top>MedEvac</td>
			<!---<td class=field id=styleSmall align=center>Fatality (Y/N)</td>--->
		</tr>
	<%
	


	'RIRHRLockingDate=CDate(outyear &"/"& outmonth  &"/"& outdays )
	'Currentdate=Date

		if  LockCount = 0 and HSESeverityval <> 1 and bDisplayInjury and  not chkHseLockingMgmt()  Then

		LockOutcome =" style=""display:none"""
		Response.Write "<script type='text/javascript'>document.getElementById('Warning21').innerHTML = '*** Key HSE Fields in this Report are now Locked ***';</script>"
				
					else
		LockOutcome=""
					
		end if	
	
	If not (RS.Bof and RS.EOF) then RS.MoveFirst

	bMoreRows = True
	For iCntr = 1 to iCt
		If RS.EOF Then bMoreRows = False
		%>
		<tr>
			<td align="center" size="10" class="field" id="styleSmall">&nbsp;<%=iCntr%>&nbsp;</td>

			<td align="center" valign="top">
			<input type="hidden" name="txtlock" value="<%=LockOutcome%>">
				<select name=txtInj<%=iCntr%>  LANGUAGE="javascript" onchange='MalariaAlert(this)' id=stylesmall>
					<%
					RSInjury.MoveFirst								
					If  bMoreRows then
						sTemp=trim(RS("InjuryType"))
					else
						sTemp = NULL
					End If

					Response.Write "<option "
					if IsNull(sTemp) Then Response.Write "selected "
					Response.Write "value=''>"
					
					Do until RSInjury.EOF %>
						<option <%If sTemp =(Trim(RSInjury("Code"))) then Response.Write "selected" %> value="<%=RSInjury("Code")%>"><%Response.Write SafeDisplay(RSInjury("ShortDescription"))
						RSInjury.MoveNext
					Loop 
					%>
				</select>
			</td>
			<td align="center" valign="top">
				<select name=txtInjLoc<%=iCntr%> id=stylesmall>
				<%
					RSBodyPart.MoveFirst								
					If  bMoreRows then
						sTemp=trim(RS("InjuryPart"))
					else
						sTemp = NULL
					End If

					Response.Write "<option "
					if IsNull(sTemp) Then Response.Write "selected "
					Response.Write "value=''>"
					
					Do until RSBodyPart.EOF %>
						<option <%If sTemp =(Trim(RSBodyPart("Code"))) then Response.Write "selected" %> value="<%=RSBodyPart("Code")%>"><%Response.Write SafeDisplay(RSBodyPart("ShortDescription"))
						RSBodyPart.MoveNext
					Loop 
				%>
				</select>
			</td>

			<td align="center" valign="top">
	<%
			If  bMoreRows then sTemp=trim(RS("DaysLost")) else sTemp = ""
					
			response.write "<input type=Hidden name='hiddentxtDays"&iCntr&"' value='"&sTemp&"' >"
			%>
				<input  type="text" name=txtDays<%=iCntr%> class="cltxtDays cltxtDaysSum<%=iCntr%>" onchange='validateSeverity(<%=iCntr%>)' value='<%If bMoreRows Then Response.write Trim(RS("DaysLost"))%>' size="1" id=stylesmall>
			</td>
			<td align="center" valign=top>
				<%If bMoreRows = True Then sTemp = Trim(RS("ReturnWorkDate"))%>
				<input type="text" name=txtReturnWorkDate<%=iCntr%> value="<%=sTemp%>" size="4" id=stylesmall>
				<%=popupCalendar("frmPers.txtReturnWorkDate" & iCntr)%>
			</td>
			<td align="center" valign="top">
	<%
			If  bMoreRows then sTemp=trim(RS("ReducedWorkDays")) else sTemp = ""
					
			response.write "<input type=Hidden name='hiddentxtRDays"&iCntr&"'  value='"&sTemp&"' >"
			%>
				<input  type="text" name=txtRDays<%=iCntr%>  class="cltxtDays cltxtRDaysSum<%=iCntr%>" onchange='validateSeverity(<%=iCntr%>)' value='<%If bMoreRows Then Response.write Trim(RS("ReducedWorkDays"))%>' size="1" id=stylesmall>
			</td>
			<td align="center" valign=top>
				
					<%
					If  bMoreRows then sTemp=trim(rs("Outcome")) else sTemp = ""
					RSOutcome.MoveFirst
response.write "<input type=Hidden name='hiddoutcomestylesmall"&iCntr&"' value='"&sTemp&"' >"
			%>
				<select name="outcome<%=iCntr%>" <%=LockOutcome%> class="outcome" onchange='fatalityAlert(this)' id=stylesmall>
				
					<%
					Response.write MakeOption("",  sTemp,"")
					Do until RSOutcome.EOF
						Response.write MakeOption(rsOutcome("InjuryOutcomeID"), sTemp,rsOutcome("InjuryOutcomeDesc"))
						RSOutcome.MoveNext
					Loop 
					%>
				</select>
			</td>
			<%
			If  bMoreRows then sTemp=trim(rs("SLBConFlag")) else sTemp = 0
			If IsOwner() Then				
				Response.Write "<td align='center' valign=top>"
				Response.Write "<input type=checkbox value=1 name='SlbCon"&iCntr&"' "&LockOutcome&" "&iif(sTemp=1,"Checked","")&" onClick='OnSLBClick("&iCntr&")'>"
				Response.Write "</TD>"
			Else
				Response.Write "<input type=hidden value="&sTemp&" name='SlbCon"&iCntr&"'>"
			End If%>
			<td align="center" valign=top>
				<select name="MedVac<%=iCntr%>" id=stylesmall>
					<%If  bMoreRows then sTemp=trim(rs("MedVac")) else sTemp = 0
						Response.Write "<option value=0>None"
						Response.Write "<option value=1 "
						If sTemp=1 Then Response.Write "selected"  
						Response.Write ">Local"
						Response.Write "<option value=2 " 
						If sTemp=2 Then Response.Write "selected"  
						Response.Write ">International"
					
					%>
					
				</select>
			</td>							
			<!---
			<td align="center" valign="top">
				<input type="text" name=txtFatal<%=iCntr%> value='<%If bMoreRows Then If  RS("Fatality") Then Response.write "Y" ELSE Response.write "N" %>' size="1">
			</td>
			--->
		</tr>
		<%If bMoreRows then RS.MoveNext
	Next
	Response.Write "</table>"
	Response.Write "<span class=boxednote id=styleSmall>"
	Response.Write "Lost work Days and Restricted Work Days commence on the Event data and include weekends and vacation days."
	Response.Write "</span>"
	Response.Write	"<br>"
	Response.Write	"<br>"
		
	Response.Write "<br><table border='1' cellPadding='0' cellSpacing='0' width='100%'> " 
	Response.Write "<tr><td class=field id=styleSmall align=center>Injury / Health</td>"
	Response.Write "<td class=field id=styleSmall align=center>Body Parts</td></tr>"
	
	Response.Write "<tr><td class=boxednote id=styleSmall align=left valign=top>"
		RSInjury.MoveFirst								
		Do until RSInjury.EOF 
			Response.Write RSInjury("ShortDescription")& "<br>"
			RSInjury.MoveNext
		Loop 
	Response.Write "</td>"	
	
	Response.Write "<td class=boxednote id=styleSmall align=left valign=top>"
		RSBodyPart.MoveFirst								
		Do until RSBodyPart.EOF 
			Response.Write RSBodyPart("LongDescription")& "<br>"
			RSBodyPart.MoveNext
		Loop 
	Response.Write "</td></tr></table>"	

	
	
	RSBodyPart.Close
	RSInjury.Close
	RSOutcome.close
	
	Set RSOutcome = Nothing
	Set RSBodyPart = Nothing
	Set RSInjury = Nothing
End Sub


Function MakeOption(txtValue,SelectedValue,Text)
	MakeOption = "<option value=""" & txtValue & """"
	if txtValue = SelectedValue Then MakeOption = MakeOption & " selected"
	MakeOption = MakeOption & ">" & Text
End Function


Function GetPartyTypes(Name,Sel,cn)
Dim Str,SQL,RS,val,sTemp
	SQL = "SELECT TypeID,TypeDescription FROM tlkpInjuredPartyType With (NOLOCK) ORDER BY TypeDescription "
	Set RS=cn.execute(SQL)
	Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
	Str=Str&"	<option value=''>(None)"&vbCRLF
	Sel=Ucase(Sel)
	While Not RS.EOF 									
		Val = UCase(Trim(RS("TypeID")))
		if Val=Sel then sTemp="SELECTED" else sTemp=""
		Str=Str&"	<option  value='"&Val&"' "&sTemp&">"& RS("TypeDescription") &vbCRLF
		RS.MoveNext
	Wend
	Str=Str&"</select>"	&vbCRLF
	RS.Close
	Set RS=Nothing
	GetPartyTypes=Str
End Function

Function GetExperience(Name,Sel,cn)
Dim Str,SQL,RS,val,sTemp
	SQL = "SELECT ExpID,ExpName FROM tlkpPersLoss_Experience  "
	Set RS=cn.execute(SQL)
	Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
	Str=Str&"	<option value=''>(None)"&vbCRLF
	Sel=Ucase(Sel)
	While Not RS.EOF 									
		Val = UCase(Trim(RS("ExpID")))
		if Val=Sel then sTemp="SELECTED" else sTemp=""
		Str=Str&"	<option  value='"&Val&"' "&sTemp&" >"& RS("ExpName") &vbCRLF
		RS.MoveNext
	Wend
	Str=Str&"</select>"	&vbCRLF
	RS.Close
	Set RS=Nothing
	GetExperience=Str
End Function

 Function GetJObFunction(Name,Sel,cn)
Dim Str,SQL,RS,val,sTemp
	SQL = "SELECT STELL,STLTX FROM tblsap_Jobtitle where STLTX <> '' ORDER BY STLTX "
	Set RS=cn.execute(SQL)

    if Sel = "" or Sel = "NULL" THEN Sel = 0
            

	Str="<select name='"&Name&"' id=styleSmall>"&vbCRLF
	Str=Str&"	<option value='0' "& iif(Sel = 0,"SELECTED","") &" >(None)"&vbCRLF
             Str=Str&"	<option value='1'"& iif(Sel=1,"SELECTED","") &" >Member of the Public"&vbCRLF
            Str=Str&"	<option value='2' "& iif(Sel=2,"SELECTED","") &" >Third Party"&vbCRLF
            Str=Str&"	<option value='3' "& iif(Sel=3,"SELECTED","") &" >Not Known"&vbCRLF
            Str=Str&"	<option value='4' "& iif(Sel=4,"SELECTED","") &" >Not Applicable"&vbCRLF
	Sel=Ucase(Sel)
	While Not RS.EOF 									
		Val = UCase(Trim(RS("STELL")))

		if Val=Sel then sTemp="SELECTED" else sTemp=""
           
        	Str=Str&"	<option  value='"&Val&"' "&sTemp&" >"& RS("STLTX")&" ("&RS("STELL")&")"&vbCRLF
            RS.MoveNext

	Wend
	Str=Str&"</select>"	&vbCRLF
	RS.Close
	Set RS=Nothing
	GetJObFunction=Str
End Function


%>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSpers.asp;9 %>
<% '       1*[1235378] 12-JUN-2009 22:11:14 (GMT) SVadla %>
<% '         "QUEST Initial Upload" %>
<% '       2*[1398286] 16-NOV-2010 12:29:24 (GMT) PMakhija %>
<% '         "Swift#2502389-Cross Scripting issue in RIR module except Reports" %>
<% '       3*[1795427] 14-OCT-2013 10:08:19 (GMT) MPatil2 %>
<% '         "ENH009592  CHANGE: Guest User to be renamed as Basic User" %>
<% '       4*[1835726] 09-MAY-2014 14:43:29 (GMT) BGohil2 %>
<% '         "NFT014129 - NPT/CMSL/TNCR data historical capture" %>
<% '       5*[1855267] 28-AUG-2014 08:45:54 (GMT) VSharma16 %>
<% '         "ENH035526 . Malaria Reporting" %>
<% '       6*[1867842] 13-NOV-2014 11:40:26 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '       7*[1870634] 17-NOV-2014 13:40:52 (GMT) VSharma16 %>
<% '         "ENH044752  HSE locking of lagging indicators - safety net - key to unlock" %>
<% '       8*[1877279] 18-DEC-2014 09:55:51 (GMT) VSharma16 %>
<% '         "update files for LOSSpers.asp" %>
<% '       9*[1915045] 18-AUG-2015 15:01:07 (GMT) VSharma16 %>
<% '         "ENH077171-SAXON - HSE Personnel Loss Tab addition" %>
<% '+- OmniWorks Replacement History - qhse`quest`Rir:LOSSpers.asp;9 %>
