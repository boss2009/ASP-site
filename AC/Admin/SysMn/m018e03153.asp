<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var IsClient = ((Request.Form("IsClient")=="1") ? "1":"0");
	var IsInstitution = ((Request.Form("IsInstitution")=="1") ? "1":"0");
	var IsLoan = ((Request.Form("IsLoan")=="1") ? "1":"0");
	var IsBuyout = ((Request.Form("IsBuyout")=="1") ? "1":"0");
	var rsDocumentType = Server.CreateObject("ADODB.Recordset");
	rsDocumentType.ActiveConnection = MM_cnnASP02_STRING;
	rsDocumentType.Source = "{call dbo.cp_doc_type("+ Request.Form("MM_recordId") + ",'" + Request.Form("Description") + "'," + IsClient + "," + IsInstitution + "," + IsLoan + "," + IsBuyout + ",0,'E',0)}";
	rsDocumentType.CursorType = 0;
	rsDocumentType.CursorLocation = 2;
	rsDocumentType.LockType = 3;
	rsDocumentType.Open();
	Response.Redirect("m018q03153.asp");
}

var rsDocumentType = Server.CreateObject("ADODB.Recordset");
rsDocumentType.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentType.Source = "{call dbo.cp_doc_type("+ Request.QueryString("intDoc_Type_Id") + ",'',0,0,0,0,1,'Q',0)}";
rsDocumentType.CursorType = 0;
rsDocumentType.CursorLocation = 2;
rsDocumentType.LockType = 3;
rsDocumentType.Open();
%>
<html>
<head>
	<title>Update Inventory Status Lookup</title>
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
<h5>Update Inventory Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsDocumentType.Fields.Item("chvDoc_Type_Desc").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Is Client:</td>
		<td><input type="checkbox" name="IsClient" <%=((rsDocumentType.Fields.Item("bitIs_Client").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" class="chkstyle"></td>
	</tr>
    <tr> 
		<td>Is Institution:</td>
		<td><input type="checkbox" name="IsInstitution" <%=((rsDocumentType.Fields.Item("bitIs_School").Value == 1)?"CHECKED":"")%> value="1" tabindex="3" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Is Loan:</td>
		<td><input type="checkbox" name="IsLoan" <%=((rsDocumentType.Fields.Item("bitIs_Loan").Value == 1)?"CHECKED":"")%> value="1" tabindex="4" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Is Buyout:</td>
		<td><input type="checkbox" name="IsBuyout" <%=((rsDocumentType.Fields.Item("bitIs_Buyout").Value == 1)?"CHECKED":"")%> value="1" tabindex="5" class="chkstyle" accesskey="L"></td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="6" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="7" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="8" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsDocumentType.Fields.Item("intDoc_Type_Id").Value%>">
</form>
</body>
</html>
<%
rsDocumentType.Close();
%>