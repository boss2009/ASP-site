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
		
	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");	
	var CustomLetterContent = String(Request.Form("CustomLetterContent")).replace(/'/g, "''");	
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	rsTemplate.Source = "{call dbo.cp_insert_crspltr_custom(0,0,0,"+Request.QueryString("insSchool_id")+","+Session("insStaff_id")+",0,4,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+",'"+DocumentName+"',0,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+",'"+CustomLetterContent+"',0)}";
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
	function GenerateEnvelope(){
		//Print recipient
		document.frm0803.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0803.Recipient.value;
		document.frm0803.target = "_blank";
		document.frm0803.submit();

		document.frm0803.CC.value = "";
		for (var i = 0; i < document.frm0803.CCList.options.length; i++) {
			if (document.frm0803.CCList.options[i].selected) {
				document.frm0803.CC.value = document.frm0803.CC.value + ":" + document.frm0803.CCList.options[i].value;
			}
		}
		
		if (document.frm0803.CC.value.length > 0) {
			document.frm0803.CC.value = document.frm0803.CC.value.substring(1, document.frm0803.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0803.CCList.options.length; i++) {
			if (document.frm0803.CCList.options[i].selected) {
				document.frm0803.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0803.CCList.options[i].value;
				document.frm0803.target = "_blank";
				document.frm0803.submit();
			}
		}		
	}
			
	function Save(){
		if (!CheckDate(document.frm0803.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0803.DateGenerated.focus();
			return ;
		}
		
		if (Trim(document.frm0803.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0803.DocumentName.focus();
			return ;
		}

		if (document.frm0803.CustomLetterContent.value.length > 4000) {
			alert("Custom letter content cannot exceed 4000 characters.");
			document.frm0803.CustomLetterContent.focus();
			return ;
		}
		
		document.frm0803.CC.value = "";
		for (var i = 0; i < document.frm0803.CCList.options.length; i++) {
			if (document.frm0803.CCList.options[i].selected) {
				document.frm0803.CC.value = document.frm0803.CC.value + ":" + document.frm0803.CCList.options[i].value;
			}
		}
		
		if (document.frm0803.CC.value.length > 0) {
			document.frm0803.CC.value = document.frm0803.CC.value.substring(1, document.frm0803.CC.value.length);
		}

		var temp = document.frm0803.action;		
		
		if (document.frm0803.MailMethod.value=="0") {				
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();		
			document.frm0803.action = "../TPL/CustomLetterTemplate.asp";
		} else {
			document.frm0803.action = "../TPL/E-CustomLetterTemplate.asp";
		}				
		document.frm0803.target = "_blank";		
		document.frm0803.submit();
		
		document.frm0803.action = temp;
		document.frm0803.target = "_self";
		document.frm0803.submit();		
	}
	
	function ChangeType(){
		if (document.frm0803.Type.value == "4") {	
	<%
	if ((String(Request.QueryString("intBuyout_req_id")) != "undefined") || (String(Request.QueryString("intLoan_req_id")) != "undefined")) {
	%>
			window.location.href = "m012a0802.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
	<%
	} else {	
	%>
			window.location.href = "m012a0801.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
	<%
	}
	%>
		}
	}
	</script>
</head>
<body onLoad="document.frm0803.Recipient.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0803">
<h5>New Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" accesskey="F" onChange="ChangeType();">
			<option value="4">Form Letter
			<option value="0" SELECTED>Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="2">
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvFst_Name").Value%> <%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
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
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvFst_Name").Value%> <%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<%
			rsContact.MoveNext();
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
<textarea name="CustomLetterContent" cols="90" rows="50"></textarea>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="CC" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsContact.Close();
%>