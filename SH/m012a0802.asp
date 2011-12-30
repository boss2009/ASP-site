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
	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");		
	var OtherEquipmentList = String(Request.Form("OtherEquipmentList")).replace(/'/g, "''");			
	var OtherEquipmentConditions = String(Request.Form("OtherEquipmentConditions")).replace(/'/g, "''");
	var OtherDocumentCondition = String(Request.Form("OtherDocumentCondition")).replace(/'/g, "''");
	var TrainingRequested = ((Request.Form("TrainingRequested")=="on")?"1":"0");
	
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;	
	rsTemplate.CursorType = 0;
	rsTemplate.CursorLocation = 2;
	rsTemplate.LockType = 3;
	if (String(Request.Form("TransactionType"))=="Loan") {
		rsTemplate.Source = "{call dbo.cp_insert_crspltr_pilat_accept("+Request.Form("intLoan_req_id")+",0,"+Request.QueryString("insSchool_id")+","+Session("insStaff_id")+",0,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',0,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("PILATAcceptReferralType")+","+Request.Form("EquipmentList")+",'"+OtherEquipmentList+"',"+Request.Form("EquipmentConditions")+",'"+Request.Form("LoanReviewDate")+"','"+Request.Form("ReturnDate")+"','"+OtherEquipmentConditions+"',"+Request.Form("DocumentConditionOne")+","+Request.Form("DocumentConditionTwo")+","+Request.Form("DocumentConditionThree")+","+Request.Form("DocumentConditionFour")+",'"+OtherDocumentCondition+"',"+TrainingRequested+",0)}";	
		rsTemplate.Open();				
		Response.Redirect("../LN/m008FS01.asp?intLoan_req_id="+Request.Form("intLoan_req_id"));
	} else {
		rsTemplate.Source = "{call dbo.cp_insert_crspltr_pilat_accept(0,"+Request.Form("intBuyout_req_id")+","+Request.QueryString("insSchool_id")+","+Session("insStaff_id")+",0,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',0,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("PILATAcceptReferralType")+","+Request.Form("EquipmentList")+",'"+OtherEquipmentList+"',"+Request.Form("EquipmentConditions")+",'','','"+OtherEquipmentConditions+"',"+Request.Form("DocumentConditionOne")+","+Request.Form("DocumentConditionTwo")+","+Request.Form("DocumentConditionThree")+","+Request.Form("DocumentConditionFour")+",'"+OtherDocumentCondition+"',"+TrainingRequested+",0)}";	
		rsTemplate.Open();				
		Response.Redirect("../BO/m010FS01.asp?intBuyout_req_id="+Request.Form("intBuyout_req_id"));
	}
}

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_school_contacts("+ Request.QueryString("insSchool_id") + ",0,0,0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2("+Request.QueryString("insSchool_id")+",'',0,0,0,0,0,0,0,'',1,'Q',0)}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();

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

var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentCondition.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,19,'',2,'Q',0)}";
rsDocumentCondition.CursorType = 0;
rsDocumentCondition.CursorLocation = 2;
rsDocumentCondition.LockType = 3;
rsDocumentCondition.Open();
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
	var DocumentArray = new Array(<%=count%>);
<% 
var i = 0;
while (!rsTemplate.EOF) {
	if (rsTemplate.Fields.Item("chvFileName").Value == "m012tpl002.asp") {
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

	function GenerateEnvelope(){
		//Print recipient
		document.frm0802.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0802.Recipient.value;
		document.frm0802.target = "_blank";
		document.frm0802.submit();

		document.frm0802.CC.value = "";
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) {
				document.frm0802.CC.value = document.frm0802.CC.value + ":" + document.frm0802.CCList.options[i].value;
			}
		}
		
		if (document.frm0802.CC.value.length > 0) {
			document.frm0802.CC.value = document.frm0802.CC.value.substring(1, document.frm0802.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) {
				document.frm0802.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0802.CCList.options[i].value;
				document.frm0802.target = "_blank";
				document.frm0802.submit();
			}
		}		
	}
		
	function Save(){
		if (!CheckDate(document.frm0802.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0802.DateGenerated.focus();
			return ;
		}
		if (Trim(document.frm0802.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0802.DocumentName.focus();
			return ;
		}
				
		document.frm0802.CC.value = "";
		
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) document.frm0802.CC.value = document.frm0802.CC.value + ":" + document.frm0802.CCList.options[i].value;
		}

		if (document.frm0802.CC.value.length > 0) document.frm0802.CC.value = document.frm0802.CC.value.substring(1, document.frm0802.CC.value.length);

		var temp = document.frm0802.action;
		
		if (document.frm0802.TransactionType.value=="Buyout") {
			if (document.frm0802.MailMethod.value=="0") {
				if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();					
				document.frm0802.action = "../TPL/"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
			} else {
				document.frm0802.action = "../TPL/E-"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";			
			}
		} else {
			if (document.frm0802.MailMethod.value=="0") {
				if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();			
				document.frm0802.action = "../TPL/"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";		
			} else {
				document.frm0802.action = "../TPL/E-"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";			
			}
		}
		
		document.frm0802.target = "_blank";
		document.frm0802.submit();
		
		document.frm0802.action = temp;
		document.frm0802.target = "_self";
		document.frm0802.submit();
	}
	
	function ChangeType(){
		if (document.frm0802.Type.value == "0") {
			window.location.href = "m012a0803.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&Type=<%=Request.QueryString("Type")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
		}
	}
	</script>
</head>
<body onLoad="document.frm0802.Type.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0802">
<h5>New Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" accesskey="F" onChange="ChangeType();">
			<option value="4" SELECTED>Form Letter
			<option value="0">Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="2">
		<%
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvFst_Name").Value%>&nbsp;<%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<%
			rsContact.MoveNext();
		}
		rsContact.Requery();		
		%>		
		</select></td>
	</tr>
	<tr>
		<td valign="top">CC:</td>
		<td valign="top"><select name="CCList" multiple size="5" tabindex="3">
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvFst_Name").Value%>&nbsp;<%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<%
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>
    <tr> 
		<td>Template:</td>
		<td><select name="Template" tabindex="4">
	<% 
	while (!rsTemplate.EOF) {
		if (rsTemplate.Fields.Item("chvFileName").Value == "m012tpl002.asp") {		
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
		<td nowrap><input type="text" name="DocumentName" maxlength="50" size="30" tabindex="5"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="7" accesskey="L">
			<option value="0">Canada Post
			<option value="1">E-Mail
		</select></td>
	</tr>	
</table>
<hr>
<h5>PILAT Accept</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PilatAcceptReferralType" tabindex="8">
			<%
			if (Request.QueryString("Type")=="Loan") {
			%>						
				<option value="1">Low Utilization
				<option value="2">Interim
				<option value="3">Donation
			<%
			} else {
			%>				
				<option value="4">Purchase
			<%
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td><b>Equipment List:</b></td>
		<td><select name="EquipmentList" tabindex="9">
				<option value="0">None
			<%
			if (Request.QueryString("Type")=="Loan") {
			%>								
				<option value="1">Loan Equipment			
				<option value="2">Donation Equipment
			<%
			} else {
			%>								
				<option value="3">Buyout Equipment
			<%
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherEquipmentList" tabindex="10" maxlength="80" size="60"></td>
	</tr>
	<tr>
		<td><b>Equipment Conditions:</b></td>
		<td></td>
	</tr>
<%
if (Request.QueryString("Type") == "Loan") {
%>	
	<tr>
		<td></td>
		<td>
			<input type="radio" name="EquipmentConditions" value="1" class="chkstyle" CHECKED tabindex="11">Low Utilization - Loan Review Date&nbsp;<input type="text" name="LoanReviewDate" maxlength="10" size="12" onChange="FormatDate(this)"><span style="font-size: 7pt">&nbsp;(mm/dd/yyyy)</span>
		</td>	
	</tr>	
	<tr>
		<td></td>
		<td>
			<input type="radio" name="EquipmentConditions" value="2" class="chkstyle" tabindex="12">Interim Loan - Return Date&nbsp;<input type="text" name="ReturnDate" maxlength="10" size="12" onChange="FormatDate(this)"><span style="font-size: 7pt">&nbsp;(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="radio" name="EquipmentConditions" class="chkstyle" value="3" tabindex="13">Donation</td>
	</tr>	
<%
} else {
%>
	<tr>
		<td></td>
		<td><input type="radio" name="EquipmentConditions" class="chkstyle" value="4" CHECKED tabindex="14">Purchase</td>
	</tr>	
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherEquipmentConditions" maxlength="80" size="60" tabindex="16"></td>
	</tr>
<%
}
%>
	<tr>
		<td><b>Document Conditions:</b></td>
		<td><select name="DocumentConditionOne" tabindex="17">
				<option value="0">(none)
			<%
			while (!rsDocumentCondition.EOF) {
			%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>"><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentCondition.MoveNext();
			}
			rsDocumentCondition.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionTwo" tabindex="18">
				<option value="0">(none)
			<%
			while (!rsDocumentCondition.EOF) {
			%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>"><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentCondition.MoveNext();
			}
			rsDocumentCondition.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionThree" tabindex="19">
				<option value="0">(none)
			<%
			while (!rsDocumentCondition.EOF) {
			%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>"><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentCondition.MoveNext();
			}
			rsDocumentCondition.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionFour" tabindex="20">
				<option value="0">(none)
			<%
			while (!rsDocumentCondition.EOF) {
			%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>"><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentCondition.MoveNext();
			}
			rsDocumentCondition.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherDocumentCondition" maxlength="80" size="60" tabindex="21"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="TrainingRequested" class="chkstyle" tabindex="22"><b>Training Requested</b></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" tabindex="23" class="btnstyle"></td>
<%
if (Request.QueryString("Type") == "Loan") {
%>		
		<td><input type="button" value="Close" onClick="window.location.href='../LN/m008FS01.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>';" tabindex="24" class="btnstyle"></td>
<%
} else {
%>
		<td><input type="button" value="Close" onClick="window.location.href='../BO/m010FS01.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>';" tabindex="24" class="btnstyle"></td>
<%
}
%>
    </tr>
</table>
<input type="hidden" name="CC" value="">
<input type="hidden" name="MM_insert" value="true">
<input type="hidden" name="TransactionType" value="<%=Request.QueryString("Type")%>">
<input type="hidden" name="intLoan_req_id" value="<%=Request.QueryString("intLoan_req_id")%>">
<input type="hidden" name="intBuyout_req_id" value="<%=Request.QueryString("intBuyout_req_id")%>">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsContact.Close();
rsTemplate.Close();
%>