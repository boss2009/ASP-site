<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var IsLoanDocument = ((Request.Form("IsLoanDocument")=="on") ? "1":"0");
	var IsOutstandingDocument = ((Request.Form("IsOutstandingDocument")=="on") ? "1":"0");
	var IsDeclineDocument = ((Request.Form("IsDeclineDocument")=="on") ? "1":"0");
	var IsPendingDocument = ((Request.Form("IsPendingDocument")=="on") ? "1":"0");
	var IncludeEquipment = ((Request.Form("IncludeEquipment")=="on") ? "1":"0");
	var TemplateName = String(Request.Form("TemplateName")).replace(/'/g, "''");
	var FileName = String(Request.Form("FileName")).replace(/'/g, "''");
	var rsLetterTemplate = Server.CreateObject("ADODB.Recordset");
	rsLetterTemplate.ActiveConnection = MM_cnnASP02_STRING;
	rsLetterTemplate.Source = "{call dbo.cp_Letter_template(0,"+Request.Form("TemplateType")+",'" + TemplateName + "',"+Request.Form("DocumentType")+",'" + FileName + "'," + IsLoanDocument + "," + IsOutstandingDocument + "," + IsDeclineDocument + "," + IsPendingDocument + "," + IncludeEquipment + "," + Session("insStaff_id")+",0,'A',0)}";
	rsLetterTemplate.CursorType = 0;
	rsLetterTemplate.CursorLocation = 2;
	rsLetterTemplate.LockType = 3;
	rsLetterTemplate.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Letter Template</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0341.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0341.TemplateName.value)==""){
			alert("Enter Letter Template Name.");
			document.frm0341.TemplateName.focus();
			return ;
		}
		document.frm0341.submit();
	}
	</script>
</head>
<body onLoad="document.frm0341.TemplateName.focus();">
<form name="frm0341" method="POST" action="<%=MM_editAction%>">
<h5>New Letter Template</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>Letter Template Name:</td>
		<td nowrap><input type="text" name="TemplateName" size="40" maxlength="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr>
		<td nowrap>Template Type:</td>
		<td nowrap><select name="TemplateType" tabindex="2">
			<option value="0">Letter
			<option value="1">Form
		</select></td>
    </tr>
    <tr>
		<td nowrap>Document Type:</td>
		<td nowrap><select name="DocumentType" tabindex="3">
			<option value="0">Others
			<option value="1">Accept
			<option value="2">PILAT
			<option value="3">Decline
			<option value="4">Pending
		</select></td>
    </tr>
    <tr>
		<td nowrap>File Name:</td>
		<td nowrap><input type="textbox" name="FileName" tabindex="4"></td>
    </tr>
    <tr>
		<td nowrap>Is Loan Document:</td>
		<td nowrap><input type="checkbox" name="IsLoanDocument" tabindex="5" class="chkstyle"></td>
	</tr>
    <tr>
		<td nowrap>Is Outstanding Document:</td>
		<td nowrap><input type="checkbox" name="IsOutstandingDocument" tabindex="6" class="chkstyle"></td>
	</tr>
    <tr>
		<td nowrap>Is Decline Document:</td>
		<td nowrap><input type="checkbox" name="IsDeclineDocument" tabindex="7" class="chkstyle"></td>
	</tr>
    <tr>
		<td nowrap>Is Pending Document:</td>
		<td nowrap><input type="checkbox" name="IsPendingDocument" tabindex="8" class="chkstyle"></td>
	</tr>
    <tr>
		<td nowrap>Include Equipment:</td>
		<td nowrap><input type="checkbox" name="IncludeEquipment" tabindex="9" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="10" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="11" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>