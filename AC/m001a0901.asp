<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var temp = String(Request.Form("CC")).split(":");
	var CC = new Array(10);
	for (var i = 0; i < 10; i++) CC[i] = 0;
	for (var i = 0; i < temp.length; i++) CC[i] = temp[i];
	if (temp == "") CC[0] = 0;
	var Is_Recipient_Client = 1;
	var temp2 = String(Request.Form("Recipient")).split(":");
	Is_Recipient_Client = ((temp2[0]=="Client")?1:0);

	CC[9] = ((Request.Form("CCClient")=="on")?"99999":"0");
		
	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");	
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	switch(String(Request.Form("Template"))) {
		//CSG MIR
		case "865":
			var OtherCSGMissingDocumentation = String(Request.Form("OtherCSGMissingDocumentation")).replace(/'/g, "''");	
			var NoBCSAPOrHNPT = ((Request.Form("NoBCSAPOrHNPT")=="on")?"1":"0");
			var BCSAPOrHNPTErrors = ((Request.Form("BCSAPOrHNPTErrors")=="on")?"1":"0");
			var NoFinancialNeed = ((Request.Form("NoFinancialNeed")=="on")?"1":"0");
			var CanadaStudentLoanDefault = ((Request.Form("CanadaStudentLoanDefault")=="on")?"1":"0");
			var InsufficientTimeForProcessingTSSP = ((Request.Form("InsufficientTimeForProcessingTSSP")=="on")?"1":"0");
			var OutstandingReceipts = ((Request.Form("OutstandingReceipts")=="on")?"1":"0");
			var ATBCDefaultForBuyout = ((Request.Form("ATBCDefaultForBuyout")=="on")?"1":"0");
			var ContactATBCForClarifyEquipment = ((Request.Form("ContactATBCForClarifyEquipment")=="on")?"1":"0");
			var RequestForSecondSystem = ((Request.Form("RequestForSecondSystem")=="on")?"1":"0");
			var IneligibleEquipment = ((Request.Form("IneligibleEquipment")=="on")?"1":"0");
			var Comment = String(Request.Form("Comment")).replace(/'/g, "''");	
			var OtherCSGMIRIssue = String(Request.Form("OtherCSGMIRIssue")).replace(/'/g, "''");				
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_csg_mir(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("CSGMissingDocumentationOne")+","+Request.Form("CSGMissingDocumentationTwo")+","+Request.Form("CSGMissingDocumentationThree")+","+Request.Form("CSGMissingDocumentationFour")+","+Request.Form("CSGMissingDocumentationFive")+",'"+OtherCSGMissingDocumentation+"',"+NoBCSAPOrHNPT+","+BCSAPOrHNPTErrors+","+NoFinancialNeed+","+CanadaStudentLoanDefault+","+InsufficientTimeForProcessingTSSP+","+OutstandingReceipts+","+ATBCDefaultForBuyout+","+ContactATBCForClarifyEquipment+","+RequestForSecondSystem+","+IneligibleEquipment+",'"+Comment+"','"+OtherCSGMIRIssue+"',0)}";
		break;
		//Loan MIR
		case "866":
			var OtherLoanMissingDocumentation = String(Request.Form("OtherLoanMissingDocumentation")).replace(/'/g, "''");	
			var InsufficientTimeForProcessing = ((Request.Form("InsufficientTimeForProcessing")=="on")?"1":"0");
			var Courses = ((Request.Form("Courses")=="on")?"1":"0");
			var JobPlacement = ((Request.Form("JobPlacement")=="on")?"1":"0");
			var JobShadow = ((Request.Form("JobShadow")=="on")?"1":"0");
			var PSTP = ((Request.Form("PSTP")=="on")?"1":"0");
			var VolunteerPosition = ((Request.Form("VolunteerPosition")=="on")?"1":"0");
			var Practicum = ((Request.Form("Practicum")=="on")?"1":"0");
			var OtherComment = String(Request.Form("OtherComment")).replace(/'/g, "''");	
			var ContactATBCForCIP = ((Request.Form("ContactATBCForCIP")=="on")?"1":"0");
			var ATBCDefaultForLoan = ((Request.Form("ATBCDefaultForLoan")=="on")?"1":"0");
			var EnrollmentInOneCourse = ((Request.Form("EnrollmentInOneCourse")=="on")?"1":"0");
			var InsufficientAcademicProgress =  ((Request.Form("InsufficientAcademicProgress")=="on")?"1":"0");
			var OtherLoanMIRIssue = String(Request.Form("OtherLoanMIRIssue")).replace(/'/g, "''");	
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_loan_mir(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("LoanMissingDocumentationOne")+","+Request.Form("LoanMissingDocumentationTwo")+","+Request.Form("LoanMissingDocumentationThree")+","+Request.Form("LoanMissingDocumentationFour")+","+Request.Form("LoanMissingDocumentationFive")+",'"+OtherLoanMissingDocumentation+"',"+InsufficientTimeForProcessing+","+Courses+","+JobPlacement+","+JobShadow+","+PSTP+","+VolunteerPosition+","+Practicum+",'"+OtherComment+"',"+ContactATBCForCIP+","+ATBCDefaultForLoan+","+EnrollmentInOneCourse+","+InsufficientAcademicProgress+",'"+OtherLoanMIRIssue+"',0)}";
		break;
		//Loan Rescind Default
		case "863":
			var OtherReasonForCancelingDefault = String(Request.Form("OtherReasonForCancelingDefault")).replace(/'/g, "''");	
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_loan_rescind(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("ReasonForCancelingDefault")+",'"+OtherReasonForCancelingDefault+"',0)}";		
		break;
		//Loan Default
		case "862":
			var OtherPurposeOfLoan = String(Request.Form("OtherPurposeOfLoan")).replace(/'/g, "''");			
			var OtherReasonForIneligibility = String(Request.Form("OtherReasonForIneligibility")).replace(/'/g, "''");	
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_loan_default(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("PurposeOfLoan")+",'"+OtherPurposeOfLoan+"',"+Request.Form("ReasonForIneligibilityOne")+","+Request.Form("ReasonForIneligibilityTwo")+","+Request.Form("ReasonForIneligibilityThree")+","+Request.Form("ReasonForIneligibilityFour")+",'"+OtherReasonForIneligibility+"',"+Request.Form("PurchaseCostOfEquipment")+",0)}";
		break;
		//Loan Annual Education Follow-Up
		case "869":
			var OtherConditionToMaintainLoan = String(Request.Form("OtherConditionToMaintainLoan")).replace(/'/g, "''");			
			var EnrolledInRequiedCourses = ((Request.Form("EnrolledInRequiedCourses")=="on")?"1":"0");
			var SuccessfulCompletion = ((Request.Form("SuccessfulCompletion")=="on")?"1":"0");			
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_loan_annual_edu_flwup(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+EnrolledInRequiedCourses+","+Request.Form("NumberOfRequiredCourses")+","+SuccessfulCompletion+",'"+OtherConditionToMaintainLoan+"','"+Request.Form("ReplyByDate")+"',0)}";
		break;
		//Loan Pending Buyout
		case "864":
			var DiscountAmount = ((Request.Form("DiscountAmount")=="")?0:Request.Form("DiscountAmount"));						
			var BuyoutCost = ((Request.Form("BuyoutCost")=="")?0:Request.Form("BuyoutCost"));						
			var NumberOfInstallments = ((Request.Form("NumberOfInstallments")=="")?0:Request.Form("NumberOfInstallments"));			
			rsTemplate.Source = "{call dbo.cp_Insert_CrspLtr_Loan_Pending_B0(0,"+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("BuyoutType")+","+Request.Form("EmploymentLoanDuration")+",'"+Request.Form("LoanExpiryDate")+"',"+DiscountAmount+","+BuyoutCost+","+NumberOfInstallments+",'"+Request.Form("InstallmentDueDates")+"','"+Request.Form("PaymentInFullDate")+"',0)}";		
		break;
	}
	rsTemplate.CursorType = 0;
	rsTemplate.CursorLocation = 2;
	rsTemplate.LockType = 3;
	rsTemplate.Open();	
	Response.Redirect("InsertSuccessful.html");
}

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
<html>
<head>
	<title>New Correspondence</title>
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
				window.close();
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
		document.frm0901.InstallmentDueDates.value = window.showModalDialog("m001pop6.asp?num="+num,"","dialogHeight: 200px; dialogWidth: 375px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");		
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
			
		document.frm0901.CC.value = "";
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) if (document.frm0901.CCList.options[i].selected) document.frm0901.CC.value = document.frm0901.CC.value + ":" + document.frm0901.CCList.options[i].value;
		
		if (document.frm0901.CC.value.length > 0) document.frm0901.CC.value = document.frm0901.CC.value.substring(1, document.frm0901.CC.value.length);
				
		var temp = document.frm0901.action;
						
		if (document.frm0901.MailMethod.value=="0") {
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();			
			document.frm0901.action = "../TPL/"+DocumentArray[document.frm0901.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
		} else {
			document.frm0901.action = "../TPL/E-"+DocumentArray[document.frm0901.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
		}
		document.frm0901.target = "_blank";
		document.frm0901.submit();
		
		document.frm0901.action = temp;
		document.frm0901.target = "_self";
		document.frm0901.submit();		
	}
	
	function ChangeType(){
		if (document.frm0901.Type.value == "0") {
			window.location.href = "m001a0904.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
		}
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
		document.frm0901.Subject.focus();
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
<h5>New Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == "Request.QueryString(\"intAdult_id\")")?"SELECTED":"")%> ><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		rsClient.Requery();
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="2" onChange="ChangeType();">
			<option value="4" SELECTED>Form Letter
			<option value="0">Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="3">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="Client:<%=(rsClient.Fields.Item("intAdult_Id").Value)%>"><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		while (!rsContact.EOF) {
		%>
			<option value="Contact:<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
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
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
		<%
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="CCClient" tabindex="5" class="chkstyle">CC the client</td>
	</tr>
    <tr> 
		<td>Template:</td>
		<td><select name="Template" tabindex="6" onChange="ChangeTemplate();">
	<% 
	while (!rsTemplate.EOF) {
		if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m001") {		
	%>
			<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>"><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
	<%
		}
		rsTemplate.MoveNext();
	}
	%>
        </select></td>
	</tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" maxlength="50" size="30" tabindex="7"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="9" accesskey="L">
			<option value="0">Canada Post
			<option value="1">E-Mail
		</select></td>
	</tr>
</table>
<hr>
<div id="LoanMIR" style="position: absolute; top: 322px">
<h5>Loan MIR</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Missing Documentation:</b></td>
		<td><select name="LoanMissingDocumentationOne">
				<option value="0">(none)
			<%
			while (!rsLoanMissingDoc.EOF) {
			%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanMissingDoc.MoveNext();
			}
			%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationTwo">
				<option value="0">(none)
			<%
			rsLoanMissingDoc.Requery();
			while (!rsLoanMissingDoc.EOF) {
			%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanMissingDoc.MoveNext();
			}
			%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationThree">
				<option value="0">(none)
			<%
			rsLoanMissingDoc.Requery();
			while (!rsLoanMissingDoc.EOF) {
			%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationFour">
				<option value="0">(none)
			<%
			rsLoanMissingDoc.Requery();
			while (!rsLoanMissingDoc.EOF) {
			%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanMissingDocumentationFive">
				<option value="0">(none)
			<%
			rsLoanMissingDoc.Requery();
			while (!rsLoanMissingDoc.EOF) {
			%>
				<option value="<%=rsLoanMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsLoanMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>	
		<td align="right">Other:</td>
		<td><input type="text" name="OtherLoanMissingDocumentation" maxlength="80" size="60">&nbsp;(7)</td>		
	</tr>
</table>
<br>
<b>Issues:</b>
<table cellpadding="1" cellspacing="1" align="center">
	<tr>
		<td nowrap colspan="2"><input type="checkbox" name="InsufficientTimeForProcessing" class="chkstyle">Insufficient Time for Processing Application Prior to end of:</td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellpadding="1" cellspacing="1" align="center">
				<tr>
					<td><input type="checkbox" name="Courses" class="chkstyle">courses&nbsp;(8a)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="JobPlacement" class="chkstyle">job placement&nbsp;(8b)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="JobShadow" class="chkstyle">job shadow&nbsp;(8c)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="PSTP" class="chkstyle">PSTP&nbsp;(8d)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="VolunteerPosition" class="chkstyle">volunteer position&nbsp;(8e)</td>
				</tr>
				<tr>
					<td><input type="checkbox" name="Practicum" class="chkstyle">practicum&nbsp;(8f)</td>
				</tr>
				<tr>
					<td>Other:&nbsp;<input type="text" name="OtherComment" maxlength="80" size="40">&nbsp;(8g)</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="ContactATBCForCIP" class="chkstyle">Contact AT-BC for CIP/clarify equip.&nbsp;(9)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="ATBCDefaultForLoan" class="chkstyle">AT-BC default for loan&nbsp;(10)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="EnrollmentInOneCourse" class="chkstyle">Enrollment in one course&nbsp;(11)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="InsufficientAcademicProgress" class="chkstyle">Insufficient academic progress&nbsp;(12)</td>
	</tr>
	<tr>
		<td colspan="2">Other:&nbsp;<input type="text" name="OtherLoanMIRIssue" maxlength="80" size="60">&nbsp;(13)</td>
	</tr>	
</table>
<hr>
</div>
<div id="CSGMIR" style="position: absolute; top: 322px">
<h5>CSG MIR</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Issues</b></td>
		<td><input type="checkbox" name="NoBCSAPOrHNPT" class="chkstyle">No current BCSAP or HNPT (1)</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="BCSAPOrHNPTErrors" class="chkstyle">BCSAP or HNPT errors (2)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="NoFinancialNeed" class="chkstyle">No Financial Need (3)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="CanadaStudentLoanDefault" class="chkstyle">Canada Student Loan Default (4)</td>
	</tr>	
	<tr>
		<td></td>	
		<td><input type="checkbox" name="InsufficientTimeForProcessingTSSP" class="chkstyle">Insufficient Time for Processing TSSP (5)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="OutstandingReceipts" class="chkstyle">Outstanding Receipts (6)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="ATBCDefaultForBuyout" class="chkstyle">AT-BC Default (7)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="ContactATBCForClarifyEquipment" class="chkstyle">Contact AT-BC for CIP/Clarify Equip. (8)</td>
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="RequestForSecondSystem" class="chkstyle">Request for Second System (9)</td>
	</tr>
	<tr>
		<td></td>	
		<td nowrap><input type="checkbox" name="IneligibleEquipment" class="chkstyle">Ineligible Equipment <input type="text" name="Comment" maxlength="80" size="60">&nbsp;(10)</td>
	</tr>
	<tr>
		<td align="right">Other:</td>	
		<td><input type="text" name="OtherCSGMIRIssue" maxlength="80" size="60">&nbsp;(11)</td>
	</tr>	
</table>
<br><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><b>Missing Documentation:</b></td>
		<td><select name="CSGMissingDocumentationOne">
				<option value="0">(none)
			<%
			while (!rsCSGMissingDoc.EOF) {
			%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsCSGMissingDoc.MoveNext();
			}
			%>				
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationTwo">
				<option value="0">(none)
			<%
			rsCSGMissingDoc.Requery();
			while (!rsCSGMissingDoc.EOF) {
			%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsCSGMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationThree">
				<option value="0">(none)
			<%
			rsCSGMissingDoc.Requery();
			while (!rsCSGMissingDoc.EOF) {
			%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsCSGMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationFour">
				<option value="0">(none)
			<%
			rsCSGMissingDoc.Requery();
			while (!rsCSGMissingDoc.EOF) {
			%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsCSGMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="CSGMissingDocumentationFive">
				<option value="0">(none)
			<%
			rsCSGMissingDoc.Requery();
			while (!rsCSGMissingDoc.EOF) {
			%>
				<option value="<%=rsCSGMissingDoc.Fields.Item("intDoc_id").Value%>"><%=rsCSGMissingDoc.Fields.Item("chvDocDesc").Value%>
			<%
				rsCSGMissingDoc.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherCSGMissingDocumentation" maxlength="80" size="65">&nbsp;(19)</td>		
	</tr>
</table>
<hr>
</div>
<div id="LoanRescindDefault" style="position: absolute; top: 322px">
<h5>Loan Rescind Default</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><b>Reason for Canceling Default:</b></td>
		<td><select name="ReasonForCancelingDefault">
				<option value="0">(none)
				<option value="1">equipment return (1)
				<option value="2">purchase of equipment (2)
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="input" name="OtherReasonForCancelingDefault" maxlength="80" size="60">&nbsp;(3)</td>
	</tr>
</table>
<hr>
</div>
<div id="LoanPendingBuyout" style="position: absolute; top: 322px">
<h5>Loan Pending Buyout</h5>
<b>Buyout Type:</b>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td width="160"><input type="radio" name="BuyoutType" value="0" CHECKED class="chkstyle">Employment</td>
		<td>Employment Loan Duration</td>
		<td><select name="EmploymentLoanDuration">
				<option value="1">1 year
				<option value="2">8 months
				<option value="3">6 months
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td>Loan Expiry Date:</td>
		<td>
			<input type="text" name="LoanExpiryDate" size="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td><input type="radio" name="BuyoutType" value="1" class="chkstyle">Other</td>
		<td></td>
		<td></td>
	</tr>
</table>
<b>Buyout Plan</b>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Discount Amount:</td>
		<td>$<input type="text" name="DiscountAmount" size="10" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td>Buyout Cost:</td>
		<td>$<input type="text" name="BuyoutCost" size="10" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td>Number of Installments:</td>
		<td>
			<input type="text" name="NumberOfInstallments" size="3" onKeypress="AllowNumericOnly();">
			<input type="button" value="Enter Due Dates" onClick="popDueDates(document.frm0901.NumberOfInstallments.value);" class="btnstyle">
			<input type="hidden" name="InstallmentDueDates">
		</td>
	</tr>
	<tr>
		<td>Payment in Full Date:</td>
		<td>
			<input type="text" name="PaymentInFullDate" size="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
</table>
<hr>
</div>
<div id="LoanDefault" style="position: absolute; top: 322px">
<h5>Loan Default</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Purpose of Loan:</b></td>
		<td><select name="PurposeOfLoan">
				<option value="0">(none)
			<%
			while (!rsPurposeOfLoan.EOF) {
			%>
				<option value="<%=rsPurposeOfLoan.Fields.Item("intDoc_id").Value%>"><%=rsPurposeOfLoan.Fields.Item("chvDocDesc").Value%>
			<%
				rsPurposeOfLoan.MoveNext();
			}
			%>				
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="input" name="OtherPurposeOfLoan" maxlength="80" size="65">&nbsp;(3)</td>
	</tr>
	<tr>
		<td nowrap><b>Reason for Ineligibility:</b></td>
		<td><select name="ReasonForIneligibilityOne">
				<option value="0">(none)
			<%
			while (!rsReasonForIneligibility.EOF) {
			%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>"><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
			<%
				rsReasonForIneligibility.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="ReasonForIneligibilityTwo">
				<option value="0">(none)
			<%
			rsReasonForIneligibility.Requery();
			while (!rsReasonForIneligibility.EOF) {
			%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>"><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
			<%
				rsReasonForIneligibility.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>	
		<td><select name="ReasonForIneligibilityThree">
				<option value="0">(none)
			<%
			rsReasonForIneligibility.Requery();
			while (!rsReasonForIneligibility.EOF) {
			%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>"><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
			<%
				rsReasonForIneligibility.MoveNext();
			}
			%>
		</select></td>
	</tr>	
	<tr>
		<td></td>	
		<td><select name="ReasonForIneligibilityFour">
				<option value="0">(none)
			<%
			rsReasonForIneligibility.Requery();
			while (!rsReasonForIneligibility.EOF) {
			%>
				<option value="<%=rsReasonForIneligibility.Fields.Item("intDoc_id").Value%>"><%=rsReasonForIneligibility.Fields.Item("chvDocDesc").Value%>
			<%
				rsReasonForIneligibility.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherReasonForIneligibility" maxlength="80" size="65">&nbsp;(9)</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Purchase Cost of Equipment:</b></td>
		<td>$<input type="text" name="PurchaseCostOfEquipment" size="10" onKeypress="AllowNumericOnly();"></td>
	</tr>
</table>
<hr>
</div>
<div id="LoanAnnualEducationFollowUp" style="position: absolute; top: 322px">
<h5>Loan Annual Education Follow-Up</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Loan Conditions:</b></td>
		<td><input type="checkbox" name="EnrolledInRequiedCourses" class="chkstyle">enrolled in the minimum of&nbsp;<input type="text" name="NumberOfRequiredCourses" size="3" onKeypress="AllowNumericOnly();">&nbsp;required courses.</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="SuccessfulCompletion" class="chkstyle">successful completion of courses</td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherConditionToMaintainLoan" size="60" maxlength="80"></td>
	</tr>	
	<tr>
		<td><b>Reply by Date:</b></td>
		<td><input type="text" name="ReplyByDate" size="12" maxlength="10" onChange="FormatDate(this)"><span style="font-size: 7pt">(mm/dd/yyyy)</span></td>	
	</tr>
</table>
<hr>
</div>
<div id="buttons" style="position: absolute; top: 510px;">
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="CC" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
%>