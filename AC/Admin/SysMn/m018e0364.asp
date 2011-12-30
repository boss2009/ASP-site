<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsLoanStatus = Server.CreateObject("ADODB.Recordset");
	rsLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsLoanStatus.Source = "{call dbo.cp_loan_status2("+Request.Form("MM_recordId")+",'"+Description+"',0,'E',0)}";
	rsLoanStatus.CursorType = 0;
	rsLoanStatus.CursorLocation = 2;
	rsLoanStatus.LockType = 3;
	rsLoanStatus.Open();
	Response.Redirect("m018q0364.asp");
}

var rsLoanStatus = Server.CreateObject("ADODB.Recordset");
rsLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
rsLoanStatus.Source = "{call dbo.cp_loan_status2("+ Request.QueryString("intloan_status_id") + ",'',1,'Q',0)}";
rsLoanStatus.CursorType = 0;
rsLoanStatus.CursorLocation = 2;
rsLoanStatus.LockType = 3;
rsLoanStatus.Open();
%>
<html>
<head>
	<title>Update Loan Status Lookup</title>
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
			document.frm0364.reset();
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
		if (Trim(document.frm0364.Description.value)==""){
			alert("Enter Description.");
			document.frm0364.Description.focus();
			return ;		
		}
		document.frm0364.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0364.Description.focus();">
<form name="frm0364" method="POST" action="<%=MM_editAction%>">
<h5>Update Loan Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsLoanStatus.Fields.Item("chvname").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F" ></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsLoanStatus.Fields.Item("intloan_status_id").Value %>">
</form>
</body>
</html>
<%
rsLoanStatus.Close();
%>