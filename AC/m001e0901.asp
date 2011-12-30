<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsLetter = Server.CreateObject("ADODB.Recordset");
rsLetter.ActiveConnection = MM_cnnASP02_STRING;
rsLetter.Source = "{call dbo.cp_get_ac_crsp_hstry3("+ Request.QueryString("intLetter_id") + ",0)}";
rsLetter.CursorType = 0;
rsLetter.CursorLocation = 2;
rsLetter.LockType = 3;
rsLetter.Open();

var rsCC = Server.CreateObject("ADODB.Recordset");
rsCC.ActiveConnection = MM_cnnASP02_STRING;
rsCC.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,1,'Q',0)}";
rsCC.CursorType = 0;
rsCC.CursorLocation = 2;
rsCC.LockType = 3;
rsCC.Open();

var CCClient = false;
while ((!rsCC.EOF) && (CCClient==false)) {
	if (rsCC.Fields.Item("intContact_Id").Value==99999) CCClient = true;
	rsCC.MoveNext();
}
rsCC.Requery();

var rsCondition = Server.CreateObject("ADODB.Recordset");
rsCondition.ActiveConnection = MM_cnnASP02_STRING;
rsCondition.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,2,'Q',0)}";
rsCondition.CursorType = 0;
rsCondition.CursorLocation = 2;
rsCondition.LockType = 3;
rsCondition.Open();

var rsPurpose = Server.CreateObject("ADODB.Recordset");
rsPurpose.ActiveConnection = MM_cnnASP02_STRING;
rsPurpose.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,3,'Q',0)}";
rsPurpose.CursorType = 0;
rsPurpose.CursorLocation = 2;
rsPurpose.LockType = 3;
rsPurpose.Open();

var rsReason = Server.CreateObject("ADODB.Recordset");
rsReason.ActiveConnection = MM_cnnASP02_STRING;
rsReason.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,4,'Q',0)}";
rsReason.CursorType = 0;
rsReason.CursorLocation = 2;
rsReason.LockType = 3;
rsReason.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact("+ Request.QueryString("intAdult_id") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client(" + Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsTemplate = Server.CreateObject("ADODB.Recordset");
rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
rsTemplate.Source = "{call dbo.cp_Letter_template(0,1,'',0,'',0,0,0,0,0,0,2,'Q',0)}";
rsTemplate.CursorType = 0;
rsTemplate.CursorLocation = 2;
rsTemplate.LockType = 3;
rsTemplate.Open();
var count = 0;
while (!rsTemplate.EOF) {
	count++;
	rsTemplate.MoveNext();
}
rsTemplate.MoveFirst();

var rsPurposeOfLoan = Server.CreateObject("ADODB.Recordset");
rsPurposeOfLoan.ActiveConnection = MM_cnnASP02_STRING;
rsPurposeOfLoan.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,12,'',2,'Q',0)}";
rsPurposeOfLoan.CursorType = 0;
rsPurposeOfLoan.CursorLocation = 2;
rsPurposeOfLoan.LockType = 3;
rsPurposeOfLoan.Open();

var rsReasonForIneligibility = Server.CreateObject("ADODB.Recordset");
rsReasonForIneligibility.ActiveConnection = MM_cnnASP02_STRING;
rsReasonForIneligibility.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,21,'',2,'Q',0)}";
rsReasonForIneligibility.CursorType = 0;
rsReasonForIneligibility.CursorLocation = 2;
rsReasonForIneligibility.LockType = 3;
rsReasonForIneligibility.Open();

var rsCSGMissingDoc = Server.CreateObject("ADODB.Recordset");
rsCSGMissingDoc.ActiveConnection = MM_cnnASP02_STRING;
rsCSGMissingDoc.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,9,'',2,'Q',0)}";
rsCSGMissingDoc.CursorType = 0;
rsCSGMissingDoc.CursorLocation = 2;
rsCSGMissingDoc.LockType = 3;
rsCSGMissingDoc.Open();

var rsLoanMissingDoc = Server.CreateObject("ADODB.Recordset");
rsLoanMissingDoc.ActiveConnection = MM_cnnASP02_STRING;
rsLoanMissingDoc.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,8,'',2,'Q',0)}";
rsLoanMissingDoc.CursorType = 0;
rsLoanMissingDoc.CursorLocation = 2;
rsLoanMissingDoc.LockType = 3;
rsLoanMissingDoc.Open();
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>View Correspondence</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
		   	case 76 :
				//alert("L");
				window.location.href='m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function popDueDates(num) {
		if ((isNaN(num)) || (num <= 0) || (num == "")){
			alert("Invalid number of installments.");
			return ;
		}
		var temp = "m001pop6.asp?num=" + num;
		document.frm0901.InstallmentDueDates.value = window.showModalDialog(temp,"","dialogHeight: 200px; dialogWidth: 375px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");		
		return; 
	}
	
	var DocumentArray = new Array(<%=count%>);
<% 
var i = 0;
while (!rsTemplate.EOF) {
	if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m001") {
%>
		DocumentArray[<%=i%>] = new Array(3);
		DocumentArray[<%=i%>][0] = <%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>;
		DocumentArray[<%=i%>][1] = "<%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%>";
		DocumentArray[<%=i%>][2] = "<%=(rsTemplate.Fields.Item("chvFileName").Value)%>";
<%
	i++;
	}
	rsTemplate.MoveNext();
}
rsTemplate.MoveFirst();
%>
		
	function Save(){
		if (!CheckDate(document.frm0901.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0901.DateGenerated.focus();
			return ;
		}
		if (Trim(document.frm0901.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0901.DocumentName.focus();
			return ;
		}
		
		if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();	
		
		document.frm0901.CC.value = "";
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) {
			if (document.frm0901.CCList.options[i].selected) document.frm0901.CC.value = document.frm0901.CC.value + ":" + document.frm0901.CCList.options[i].value;
		}
		
		if (document.frm0901.CC.value.length > 0) document.frm0901.CC.value = document.frm0901.CC.value.substring(1, document.frm0901.CC.value.length);
				
		var temp = document.frm0901.action;
						
		if (document.frm0901.MailMethod.value=="0") {
			document.frm0901.action = "../TPL/"+DocumentArray[document.frm0901.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
		} else {
			document.frm0901.action = "../TPL/E-"+DocumentArray[document.frm0901.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
		}
		document.frm0901.target = "_blank";
		document.frm0901.submit();
	}
	
	function GenerateEnvelope(){
		//Print recipient
		var temp = document.frm0901.Recipient.value.split(":");
		document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=" + temp[0] + "&To=" + temp[1];		
		document.frm0901.target = "_blank";
		document.frm0901.submit();

		document.frm0901.CC.value = "";
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) {
			if (document.frm0901.CCList.options[i].selected) document.frm0901.CC.value = document.frm0901.CC.value + ":" + document.frm0901.CCList.options[i].value;
		}
		
		if (document.frm0901.CC.value.length > 0) document.frm0901.CC.value = document.frm0901.CC.value.substring(1, document.frm0901.CC.value.length);
		
		//Print CCs
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) {
			if (document.frm0901.CCList.options[i].selected) {
				document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0901.CCList.options[i].value;
				document.frm0901.target = "_blank";
				document.frm0901.submit();
			}
		}
		
		//if CC Client
		if (document.frm0901.CCClient.checked) {
			document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=Client&To=<%=Request.QueryString("intAdult_id")%>";
			document.frm0901.target = "_blank";
			document.frm0901.submit();
		}		
	}
	
	function Init(){
		ChangeTemplate();
	}
	
	function ChangeTemplate(){
		LoanMIR.style.visibility = "hidden";
		CSGMIR.style.visibility = "hidden";
		LoanDefault.style.visibility = "hidden";
		LoanAnnualEducationFollowUp.style.visibility = "hidden";
		LoanRescindDefault.style.visibility = "hidden";
		LoanPendingBuyout.style.visibility = "hidden";
		switch (String(document.frm0901.Template.value)) {
			case "865":
				CSGMIR.style.visibility = "visible";			
				buttons.style.top = "810px";
			break;
			case "869":
				LoanAnnualEducationFollowUp.style.visibility = "visible";
				buttons.style.top = "490px";				
			break;
			case "862":
				LoanDefault.style.visibility = "visible";
				buttons.style.top = "600px";
			break;
			case "866":
				LoanMIR.style.visibility = "visible";
				buttons.style.top = "870px";				
			break;
			case "864":
				LoanPendingBuyout.style.visibility = "visible";			
				buttons.style.top = "580px";
			break;
			case "863":
				LoanRescindDefault.style.visibility = "visible";			
				buttons.style.top = "450px";				
			break;
		}
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0901">
<h5>View Correspondence</h5>
<i>This page is readonly.</i>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == rsLetter.Fields.Item("intSubject_id").Value)?"SELECTED":"")%>><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		rsClient.Requery();
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="2">
			<option value="4" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Form Letter")?"SELECTED":"")%>>Form Letter
			<option value="0" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Custom Letter")?"SELECTED":"")%>>Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td>Recipient:</td>
		<td><select name="Recipient" tabindex="3">
			<% 
			while (!rsClient.EOF) {
			%>
				<option value="Client:<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsLetter.Fields.Item("chvRx_Class").Value=="Client")?"SELECTED":"")%>><%=(rsClient.Fields.Item("chvName").Value)%></option>
			<%
				rsClient.MoveNext();
			}
			while (!rsContact.EOF) {
			%>
				<option value="Contact:<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=(((rsLetter.Fields.Item("chvRx_Class").Value=="Contact")&&(rsContact.Fields.Item("intContact_id").Value==rsLetter.Fields.Item("intRecipient_id").Value))?"SELECTED":"")%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
			<%
				rsContact.MoveNext();
			}
			rsContact.Requery();		
			%>		
		</select></td>
	</tr>
	<tr>
		<td valign="top">CC:</td>
		<td valign="top"><select name="CCList" multiple size="5" tabindex="4">
		<% 
		var selected;
		while (!rsContact.EOF) {
			rsCC.Requery();
			selected = "";
			while (!rsCC.EOF) {
				if (rsContact.Fields.Item("intContact_id").Value==rsCC.Fields.Item("intContact_Id").Value) selected = "SELECTED";
				rsCC.MoveNext();
			}		
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=selected%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)		
		<%			
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="CCClient" tabindex="5" <%=((CCClient)?"CHECKED":"")%> class="chkstyle">CC the client</td>
	</tr>	
    <tr> 
		<td>Template:</td>
		<td><select name="Template" tabindex="6">
		<% 
		while (!rsTemplate.EOF) {
			if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m001") {		
		%>
				<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>" <%=((rsTemplate.Fields.Item("insTemplate_id").Value==Request.QueryString("insTemplate_id"))?"SELECTED":"")%>><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
		<%
			}
			rsTemplate.MoveNext();
		}
		%>
        </select></td>
	</tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" value="<%=rsLetter.Fields.Item("chvLetter_Name").Value%>" maxlength="50" size="30" tabindex="7"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=FilterDate(rsLetter.Fields.Item("dtsSend_Date").Value)%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="9" accesskey="L">
			<option value="0" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="Canada_Post")?"SELECTED":"")%>>Canada Post
			<option value="1" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="e-Mail")?"SELECTED":"")%>>E-Mail
		</select></td>
	</tr>
</table>
<hr>
<div id="LoanMIR" style="position: absolute; top: 342px">
<h5>Loan MIR</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Missing Documentation:</b></td>
		<td><select name="LoanMissingDocumentationOne">
				<option value="0">(none)
	<%			
	rsReason.Requery();	
	if (!rsReason.EOF) {
		while (!rsLoanMissingDoc.EOF) {
			if (rsLoanMissingDoc.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {
	%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsLoanMissingDoc.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationTwo">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsLoanMissingDoc.Requery();	
		while (!rsLoanMissingDoc.EOF) {
			if (rsLoanMissingDoc.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}		
			rsLoanMissingDoc.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationThree">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsLoanMissingDoc.Requery();	
		while (!rsLoanMissingDoc.EOF) {
			if (rsLoanMissingDoc.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}		
			rsLoanMissingDoc.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationFour">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsLoanMissingDoc.Requery();	
		while (!rsLoanMissingDoc.EOF) {
			if (rsLoanMissingDoc.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}		
			rsLoanMissingDoc.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationFive">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsLoanMissingDoc.Requery();	
		while (!rsLoanMissingDoc.EOF) {
			if (rsLoanMissingDoc.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}		
			rsLoanMissingDoc.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>	
		<td align="right">Other:</td>
		<td><input type="text" name="OtherLoanMissingDocumentation" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" maxlength="80" size="60">&nbsp;(7)</td>		
	</tr>
</table>
<br>
<b>Issues:</b>
<table cellpadding="1" cellspacing="1" align="center">
	<tr>
		<td nowrap><input type="checkbox" name="InsufficientTimeForProcessing" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_01").Value=="1")?"CHECKED":"")%>>Insufficient Time for Processing Application Prior to end of:</td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellpadding="1" cellspacing="1" align="center">	
				<tr>
					<td><input type="checkbox" name="Courses" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_02").Value=="1")?"CHECKED":"")%>>courses&nbsp;(8a)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="JobPlacement" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_03").Value=="1")?"CHECKED":"")%>>job placement&nbsp;(8b)</td>		
				</tr>
				<tr>
					<td><input type="checkbox" name="JobShadow" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_04").Value=="1")?"CHECKED":"")%>>job shadow&nbsp;(8c)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="PSTP" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_05").Value=="1")?"CHECKED":"")%>>PSTP&nbsp;(8d)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="VolunteerPosition" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_06").Value=="1")?"CHECKED":"")%>>volunteer position&nbsp;(8e)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="Practicum" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_07").Value=="1")?"CHECKED":"")%>>practicum&nbsp;(8f)</td>
				</tr>
				<tr>
					<td>Other:&nbsp;<input type="text" name="OtherComment" value="<%=Trim(rsLetter.Fields.Item("chvText_02").Value)%>" maxlength="80" size="40">&nbsp;(8g)</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="ContactATBCForCIP" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_08").Value=="1")?"CHECKED":"")%>>Contact AT-BC for CIP/clarify equip.&nbsp;(9)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="ATBCDefaultForLoan" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_09").Value=="1")?"CHECKED":"")%>>AT-BC default for loan&nbsp;(10)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="EnrollmentInOneCourse" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_10").Value=="1")?"CHECKED":"")%>>Enrollment in one course&nbsp;(11)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="InsufficientAcademicProgress" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_11").Value=="1")?"CHECKED":"")%>>Insufficient academic progress&nbsp;(12)</td>
	</tr>
	<tr>
		<td colspan="2">Other:&nbsp;<input type="text" name="OtherLoanMIRIssue" value="<%=Trim(rsLetter.Fields.Item("chvText_03").Value)%>" maxlength="80" size="60">&nbsp;(13)</td>
	</tr>	
</table>
<hr>
</div>
<div id="CSGMIR" style="position: absolute; top: 342px">
<h5>CSG MIR</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Issues</b></td>
		<td><input type="checkbox" name="NoBCSAPOrHNPT" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_01").Value=="1")?"CHECKED":"")%>>No current BCSAP or HNPT&nbsp;(1)</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="BCSAPOrHNPTErrors" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_02").Value=="1")?"CHECKED":"")%>>BCSAP or HNPT errors&nbsp;(2)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="NoFinancialNeed" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_03").Value=="1")?"CHECKED":"")%>>No Financial Need&nbsp;(3)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="CanadaStudentLoanDefault" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_04").Value=="1")?"CHECKED":"")%>>Canada Student Loan Default&nbsp;(4)</td>
	</tr>	
	<tr>
		<td></td>	
		<td><input type="checkbox" name="InsufficientTimeForProcessingTSSP" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_05").Value=="1")?"CHECKED":"")%>>Insufficient Time for Processing TSSP&nbsp;(5)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="OutstandingReceipts" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_06").Value=="1")?"CHECKED":"")%>>Outstanding Receipts&nbsp;(6)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="ATBCDefaultForBuyout" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_07").Value=="1")?"CHECKED":"")%>>AT-BC Default for Buyout&nbsp;(7)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="ContactATBCForClarifyEquipment" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_08").Value=="1")?"CHECKED":"")%>>Contact AT-BC for CIP/Clarify Equip.&nbsp;(8)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="RequestForSecondSystem" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_09").Value=="1")?"CHECKED":"")%>>Request for Second System&nbsp;(9)</td>
	</tr>
	<tr>
		<td></td>	
		<td nowrap><input type="checkbox" name="IneligibleEquipment" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_10").Value=="1")?"CHECKED":"")%>>Ineligible Equipment&nbsp;(10)<input type="text" name="Comment" value="<%=Trim(rsLetter.Fields.Item("chvText_02").Value)%>" maxlength="80" size="60"></td>
	</tr>
	<tr>
		<td align="right">Other:</td>	
		<td><input type="text" name="OtherCSGMIRIssue" value="<%=Trim(rsLetter.Fields.Item("chvText_03").Value)%>" maxlength="80" size="60">&nbsp;(11)</td>
	</tr>
</table>
<br><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><b>Missing Documentation:</b></td>
		<td><select name="CSGMissingDocumentationOne">
				<option value="0">(none)
	<%	
	rsCondition.Requery();	
	if (!rsCondition.EOF) {
		while (!rsCSGMissingDoc.EOF) {
			if (rsCSGMissingDoc.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {
	%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsCSGMissingDoc.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>				
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationTwo">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsCSGMissingDoc.Requery();	
		while (!rsCSGMissingDoc.EOF) {
			if (rsCSGMissingDoc.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%	
			}		
			rsCSGMissingDoc.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationThree">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsCSGMissingDoc.Requery();	
		while (!rsCSGMissingDoc.EOF) {
			if (rsCSGMissingDoc.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsCSGMissingDoc.MoveNext();			
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationFour">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsCSGMissingDoc.Requery();	
		while (!rsCSGMissingDoc.EOF) {
			if (rsCSGMissingDoc.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsCSGMissingDoc.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationFive">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsCSGMissingDoc.Requery();	
		while (!rsCSGMissingDoc.EOF) {
			if (rsCSGMissingDoc.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsCSGMissingDoc.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherCSGMissingDocumentation" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" maxlength="80" size="65"></td>		
	</tr>
</table>
<hr>
</div>
<div id="LoanRescindDefault" style="position: absolute; top: 342px">
<h5>Loan Rescind Default</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><b>Reason for Canceling Default:</b></td>
		<td nowrap><select name="ReasonForCancelingDefault">
			<%
			rsReason.Requery();
			%>
			<option value="0" <%if (!rsReason.EOF) { if (rsReason.Fields.Item("intDoc_Id").Value==0) {Response.Write("SELECTED"); rsReason.MoveNext();}}%>>(none)
			<option value="1" <%if (!rsReason.EOF) { if (rsReason.Fields.Item("intDoc_Id").Value==1) {Response.Write("SELECTED"); rsReason.MoveNext();}}%>>equipment return (1)
			<option value="2" <%if (!rsReason.EOF) { if (rsReason.Fields.Item("intDoc_Id").Value==2) {Response.Write("SELECTED"); rsReason.MoveNext();}}%>>purchase of equipment (2)
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="input" name="OtherReasonForCancelingDefault" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" maxlength="80" size="60">&nbsp;(3)</td>
	</tr>
</table>
<hr>
</div>
<div id="LoanPendingBuyout" style="position: absolute; top: 342px">
<h5>Loan Pending Buyout</h5>
<b>Buyout Type:</b>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td width="160"><input type="radio" name="BuyoutType" value="0" <%=((rsLetter.Fields.Item("bitCkBx_01").value=="0")?"CHECKED":"")%> class="chkstyle">Employment</td>
		<td>Employment Loan Duration</td>
		<td><select name="EmploymentLoanDuration">
				<option value="1" <%=((rsLetter.Fields.Item("insINT_01").value==1)?"SELECTED":"")%>>1 year
				<option value="2" <%=((rsLetter.Fields.Item("insINT_01").value==2)?"SELECTED":"")%>>6 months
				<option value="3" <%=((rsLetter.Fields.Item("insINT_01").value==3)?"SELECTED":"")%>>8 months
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td>Loan Expiry Date:</td>
		<td>
			<input type="text" name="LoanExpiryDate" value="<%=FilterDate(rsLetter.Fields.Item("dtsDate_01").Value)%>" size="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td><input type="radio" name="BuyoutType" value="1" <%=((rsLetter.Fields.Item("bitCkBx_01").value=="1")?"CHECKED":"")%> class="chkstyle">Other</td>
		<td></td>
		<td></td>
	</tr>
</table>
<b>Buyout Plan</b>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Discount Amount:</td>
		<td>$<input type="text" name="DiscountAmount" value="<%=Trim(rsLetter.Fields.Item("fltFloat_01").Value)%>" size="10"></td>
	</tr>
	<tr>
		<td>Buyout Cost:</td>
		<td>$<input type="text" name="BuyoutCost" value="<%=Trim(rsLetter.Fields.Item("fltFloat_02").Value)%>" size="10"></td>
	</tr>
	<tr>
		<td>Number of Installments:</td>
		<td>
			<input type="text" name="NumberOfInstallments" size="3" value="<%=Number(rsLetter.Fields.Item("insINT_02").Value)%>" onKeypress="AllowNumericOnly();">
			<input type="button" value="Enter Due Dates" onClick="popDueDates(document.frm0901.NumberOfInstallments.value);" class="btnstyle">
			<input type="hidden" name="InstallmentDueDates" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>">
		</td>
	</tr>
	<tr>
		<td>Payment in Full Date:</td>
		<td>
			<input type="text" name="PaymentInFullDate" size="11" maxlength="10" value="<%=FilterDate(rsLetter.Fields.Item("dtsDate_02").Value)%>" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
</table>
<hr>
</div>
<div id="LoanDefault" style="position: absolute; top: 342px">
<h5>Loan Default</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Purpose of Loan:</b></td>
		<td><select name="PurposeOfLoan">
				<option value="0">(none)
			<%
			while (!rsPurposeOfLoan.EOF) {
			%>
				<option value="<%=rsPurposeOfLoan.Fields.Item("intDoc_id").Value%>" <%if (!rsPurpose.EOF) { if (rsPurposeOfLoan.Fields.Item("intDoc_id").Value==rsPurpose.Fields.Item("intDoc_Id").Value) {Response.Write("SELECTED"); rsPurpose.MoveNext();}}%>><%=rsPurposeOfLoan.Fields.Item("chvDocDesc").Value%>
			<%
				rsPurposeOfLoan.MoveNext();
			}
			%>				
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="input" name="OtherPurposeOfLoan" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" maxlength="80" size="65">&nbsp;(3)</td>
	</tr>
	<tr>
		<td nowrap><b>Reason for Ineligibility:</b></td>
		<td><select name="ReasonForIneligibilityOne">
				<option value="0">(none)
	<%
	rsReason.Requery();	
	if (!rsReason.EOF) {
		while (!rsReasonForIneligibility.EOF) {
			if (rsReasonForIneligibility.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {
	%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsReasonForIneligibility.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="ReasonForIneligibilityTwo">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsReasonForIneligibility.Requery();	
		while (!rsReasonForIneligibility.EOF) {
			if (rsReasonForIneligibility.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsReasonForIneligibility.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>	
		<td><select name="ReasonForIneligibilityThree">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsReasonForIneligibility.Requery();	
		while (!rsReasonForIneligibility.EOF) {
			if (rsReasonForIneligibility.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsReasonForIneligibility.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>
		</select></td>
	</tr>	
	<tr>
		<td></td>	
		<td><select name="ReasonForIneligibilityFour">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsReasonForIneligibility.Requery();	
		while (!rsReasonForIneligibility.EOF) {
			if (rsReasonForIneligibility.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsReasonForIneligibility.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherReasonForIneligibility" value="<%=Trim(rsLetter.Fields.Item("chvText_02").Value)%>" maxlength="80" size="65">&nbsp;(9)</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Purchase Cost of Equipment:</b></td>
		<td>$<input type="text" name="PurchaseCostOfEquipment" value="<%=Trim(rsLetter.Fields.Item("fltFloat_01").Value)%>" size="10" onKeypress="AllowNumericOnly();"></td>
	</tr>
</table>
<hr>
</div>
<div id="LoanAnnualEducationFollowUp" style="position: absolute; top: 342px">
<h5>Loan Annual Education Follow-Up</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Loan Conditions:</b></td>
		<td><input type="checkbox" name="EnrolledInRequiedCourses" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_01").Value=="1")?"CHECKED":"")%>>enrolled in the minimum of&nbsp;<input type="text" name="NumberOfRequiredCourses" value="<%=Number(rsLetter.Fields.Item("insINT_01").Value)%>" size="3" onKeypress="AllowNumericOnly();">&nbsp;required courses.</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="SuccessfulCompletion" class="chkstyle" <%=((rsLetter.Fields.Item("bitCkBx_02").Value=="1")?"CHECKED":"")%>>successful completion of courses</td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherConditionToMaintainLoan" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" size="60" maxlength="80"></td>
	</tr>	
	<tr>
		<td><b>Reply by Date:</b></td>
		<td><input type="text" name="ReplyByDate" value="<%=FilterDate(rsLetter.Fields.Item("dtsDate_01").Value)%>" size="12" maxlength="10" onChange="FormatDate(this)"><span style="font-size: 7pt">(mm/dd/yyyy)</span></td>	
	</tr>
</table>
<hr>
</div>
<div id="buttons" style="position: absolute; top: 510px;">
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="View Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>';" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="CC" value="">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
rsLetter.Close();
rsCC.Close();
rsPurpose.Close();
rsCondition.Close();
rsReason.Close();
%>