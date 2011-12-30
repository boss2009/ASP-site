<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_Insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsClient = ((Request.Form("IsClient")=="on") ? "1":"0");
	var IsInstitution = ((Request.Form("IsInstitution")=="on") ? "1":"0");
	var IsLoan = ((Request.Form("IsLoan")=="on") ? "1":"0");
	var IsBuyout = ((Request.Form("IsBuyout")=="on") ? "1":"0");
	var rsDocumentType = Server.CreateObject("ADODB.Recordset");
	rsDocumentType.ActiveConnection = MM_cnnASP02_STRING;
	rsDocumentType.Source = "{call dbo.cp_doc_type(0,'"+ Description + "'," + IsClient + "," + IsInstitution + "," + IsLoan + "," + IsBuyout + ",0,'A',0)}";
	rsDocumentType.CursorType = 0;
	rsDocumentType.CursorLocation = 2;
	rsDocumentType.LockType = 3;
	rsDocumentType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Document Type Lookup</title>
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
				document.frm03153.reset();
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
		if (Trim(document.frm03153.Description.value)==""){
			alert("Enter Description.");
			document.frm03153.Description.focus();
			return ;
		}
		document.frm03153.submit();
	}
	</script>
</head>
<body onLoad="document.frm03153.Description.focus();">
<form name="frm03153" method="POST" action="<%=MM_editAction%>">
<h5>New Document Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr>
		<td>Is Client:</td>
		<td><input type="checkbox" name="IsClient" tabindex="2" class="chkstyle"></td>
	</tr>
    <tr>
		<td>Is Institution:</td>
		<td><input type="checkbox" name="IsInstitution" tabindex="3" class="chkstyle"></td>
	</tr>
    <tr>
		<td>Is Loan:</td>
		<td><input type="checkbox" name="IsLoan" tabindex="4" class="chkstyle"></td>
	</tr>
    <tr>
		<td>Is Buyout:</td>
		<td><input type="checkbox" name="IsBuyout" tabindex="5" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="6" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="7" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>