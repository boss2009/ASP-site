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
	var IsInstitutionLoan = ((Request.Form("IsInstitutionLoan")=="1")?"1":"0");
	var rsLoanType = Server.CreateObject("ADODB.Recordset");
	rsLoanType.ActiveConnection = MM_cnnASP02_STRING;
	rsLoanType.Source = "{call dbo.cp_loan_type2("+Request.Form("MM_recordId")+",'"+Description+"',"+IsInstitutionLoan+",0,'E',0)}";
	rsLoanType.CursorType = 0;
	rsLoanType.CursorLocation = 2;
	rsLoanType.LockType = 3;
	rsLoanType.Open();
	Response.Redirect("m018q0363.asp");
}

var rsLoanType = Server.CreateObject("ADODB.Recordset");
rsLoanType.ActiveConnection = MM_cnnASP02_STRING;
rsLoanType.Source = "{call dbo.cp_loan_type2("+ Request.QueryString("intloan_type_id") + ",'',0,1,'Q',1)}";
rsLoanType.CursorType = 0;
rsLoanType.CursorLocation = 2;
rsLoanType.LockType = 3;
rsLoanType.Open();
%>
<html>
<head>
	<title>Update Loan Type Lookup</title>
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
			document.frm0363.reset();
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
		if (Trim(document.frm0363.Description.value)==""){
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
<h5>Update Loan Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsLoanType.Fields.Item("chvname").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>Is Institution Loan:</td>
		<td><input type="checkbox" name="IsInstitutionLoan" <%=((rsLoanType.Fields.Item("bitIs_Institution_App").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>        
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsLoanType.Fields.Item("intloan_type_id").Value %>">
</form>
</body>
</html>
<%
rsLoanType.Close();
%>