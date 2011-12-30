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
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	switch(String(Request.Form("Template"))) {
		//PILAT Decline
		case "867":
			var OtherDeclineReason = String(Request.Form("OtherDeclineReason")).replace(/'/g, "''");	
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_pilat_decline(0,"+Request.QueryString("insSchool_id")+","+Session("insStaff_id")+",0,0,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',0,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("PilatDeclineReferralType")+","+Request.Form("DeclineReasonOne")+","+Request.Form("DeclineReasonTwo")+","+Request.Form("DeclineReasonThree")+","+Request.Form("DeclineReasonFour")+",'"+OtherDeclineReason+"',0)}";
		break;
		//PILAT Accept Consult/Training
		case "870":
			var OtherConditions = String(Request.Form("OtherConditions")).replace(/'/g, "''");	
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_pilat_accept_ct(0,"+Request.QueryString("insSchool_id")+","+Session("insStaff_id")+",0,0,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',0,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("PILATAcceptReferralType")+","+Request.Form("Conditions")+",'"+OtherConditions+"',0)}";
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

var rsDeclineReason = Server.CreateObject("ADODB.Recordset");
rsDeclineReason.ActiveConnection = MM_cnnASP02_STRING;
rsDeclineReason.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,7,'',2,'Q',0)}";
rsDeclineReason.CursorType = 0;
rsDeclineReason.CursorLocation = 2;
rsDeclineReason.LockType = 3;
rsDeclineReason.Open();
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
	if ((rsTemplate.Fields.Item("chvFileName").Value=="m012tpl001.asp") || (rsTemplate.Fields.Item("chvFileName").Value=="m012tpl003.asp")) {	
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
		document.frm0801.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0801.Recipient.value;
		document.frm0801.target = "_blank";
		document.frm0801.submit();

		document.frm0801.CC.value = "";
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.CC.value = document.frm0801.CC.value + ":" + document.frm0801.CCList.options[i].value;
			}
		}
		
		if (document.frm0801.CC.value.length > 0) {
			document.frm0801.CC.value = document.frm0801.CC.value.substring(1, document.frm0801.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0801.CCList.options[i].value;
				document.frm0801.target = "_blank";
				document.frm0801.submit();
			}
		}		
	}
		
	function Save(){
		if (!CheckDate(document.frm0801.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0801.DateGenerated.focus();
			return ;
		}
		
		if (Trim(document.frm0801.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0801.DocumentName.focus();
			return ;
		}
		
		document.frm0801.CC.value = "";
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.CC.value = document.frm0801.CC.value + ":" + document.frm0801.CCList.options[i].value;
			}
		}
		if (document.frm0801.CC.value.length > 0) {
			document.frm0801.CC.value = document.frm0801.CC.value.substring(1, document.frm0801.CC.value.length);
		}
		
		var temp = document.frm0801.action;
		
		if (document.frm0801.MailMethod.value=="0") {
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();						
			document.frm0801.action = "../TPL/"+DocumentArray[document.frm0801.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		} else {
			document.frm0801.action = "../TPL/E-"+DocumentArray[document.frm0801.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		}
		
		document.frm0801.target = "_blank";
		document.frm0801.submit();
		
		document.frm0801.action = temp;
		document.frm0801.target = "_self";
		document.frm0801.submit();
	}
	
	function ChangeType(){
		if (document.frm0801.Type.value == "0") {
			window.location.href = "m012a0803.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		}
	}

	function ChangeTemplate(){
		PILATAccept.style.visibility = "hidden";
		PILATDecline.style.visibility = "hidden";
		switch (String(document.frm0801.Template.value)) {
			case "870":
				PILATAccept.style.visibility = "visible";			
				buttons.style.top = "390px";
			break;
			case "867":
				PILATDecline.style.visibility = "visible";
				buttons.style.top = "450px";
			break;
		}
	}
	
	function Init(){
		document.frm0801.Type.focus();
		ChangeTemplate();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0801">
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
		<td><select name="Template" tabindex="4" onChange="ChangeTemplate();">
	<% 
	while (!rsTemplate.EOF) {
		if ((rsTemplate.Fields.Item("chvFileName").Value=="m012tpl001.asp") || (rsTemplate.Fields.Item("chvFileName").Value=="m012tpl003.asp")) {
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
<div id="PILATAccept" style="position: absolute; top: 285px">
<h5>PILAT Accept Consult/Training</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PILATAcceptReferralType">
				<option value="5">Training Only
				<option value="6">Consultation Only
		</select></td>
	</tr>
	<tr>
		<td><b>Conditions:</b></td>
		<td><select name="Conditions">
				<option value="1">Training Only Conditions
				<option value="2">Consultation Only Conditions
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherConditions" size="60" maxlength="80"></td>
	</tr>
</table>
</div>
<div id="PILATDecline" style="position: absolute; top: 285px">
<h5>PILAT Decline</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PilatDeclineReferralType">
				<option value="1">Low Utilization
				<option value="2">Interim
				<option value="3">Donation
				<option value="4">Purchase
		</select></td>
	</tr>
	<tr>
		<td><b>Reasons:</b></td>
		<td><select name="DeclineReasonOne">
				<option value="0">(none)
			<%
			while (!rsDeclineReason.EOF) {
			%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>"><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsDeclineReason.MoveNext();
			}
			rsDeclineReason.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonTwo">
				<option value="0">(none)
			<%
			while (!rsDeclineReason.EOF) {
			%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>"><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsDeclineReason.MoveNext();
			}
			rsDeclineReason.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonThree">
				<option value="0">(none)
			<%
			while (!rsDeclineReason.EOF) {
			%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>"><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsDeclineReason.MoveNext();
			}
			rsDeclineReason.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonFour">
				<option value="0">(none)
			<%
			while (!rsDeclineReason.EOF) {
			%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>"><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsDeclineReason.MoveNext();
			}
			rsDeclineReason.MoveFirst();
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherDeclineReason" maxlength="80" size="60"></td>		
	</tr>
</table>
</div>
<div id="buttons" style="position: absolute; top: 450px">
<hr>
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
rsInstitution.Close();
rsContact.Close();
rsTemplate.Close();
%>