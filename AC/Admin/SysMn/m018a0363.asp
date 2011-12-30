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
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var IsInstitutionLoan = ((Request.Form("IsInstitutionLoan")=="1")?"1":"0");
	var rsLoanType = Server.CreateObject("ADODB.Recordset");
	rsLoanType.ActiveConnection = MM_cnnASP02_STRING;
	rsLoanType.Source = "{call dbo.cp_loan_type2(0,'"+Description+"',"+IsInstitutionLoan+",0,'A',0)}";
	rsLoanType.CursorType = 0;
	rsLoanType.CursorLocation = 2;
	rsLoanType.LockType = 3;
	rsLoanType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Loan Type</title>
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
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0363.Description.value)=="") {
			alert("Enter Description.");
			document.frm0363.Description.focus();
			return ;
		}
		document.frm0363.submit();
	}
	</script>
</head>
<body onLoad="document.frm0363.Description.focus();">
<form name="frm0363" method="POST" action="<%=MM_editAction%>">
<h5>New Loan Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Institution Loan:</td>
		<td nowrap><input type="checkbox" name="IsInstitutionLoan" value="1" tabindex="2" accesskey="L" class="chkstyle"></td>        
    </tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="4" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>