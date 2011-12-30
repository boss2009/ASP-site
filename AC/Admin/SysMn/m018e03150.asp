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
	var Code = String(Request.Form("Code")).replace(/'/g, "''");	
	var rsEmploymentType = Server.CreateObject("ADODB.Recordset");
	rsEmploymentType.ActiveConnection = MM_cnnASP02_STRING;
	rsEmploymentType.Source = "{call dbo.cp_employ_type("+ Request.Form("MM_recordId") + ",'" + Code + "','" + Description + "',0,'E',0)}";
	rsEmploymentType.CursorType = 0;
	rsEmploymentType.CursorLocation = 2;
	rsEmploymentType.LockType = 3;
	rsEmploymentType.Open();
	Response.Redirect("m018q03150.asp");
}

var rsEmploymentType = Server.CreateObject("ADODB.Recordset");
rsEmploymentType.ActiveConnection = MM_cnnASP02_STRING;
rsEmploymentType.Source = "{call dbo.cp_employ_type("+ Request.QueryString("insEmploy_Type_Id") + ",'','',1,'Q',0)}";
rsEmploymentType.CursorType = 0;
rsEmploymentType.CursorLocation = 2;
rsEmploymentType.LockType = 3;
rsEmploymentType.Open();
%>
<html>
<head>
	<title>Update Employment Type Lookup</title>
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
			document.frm03150.reset();
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
		if (Trim(document.frm03150.Description.value)==""){
			alert("Enter Description.");
			document.frm03150.Description.focus();
			return ;		
		}
		if (Trim(document.frm03150.Code.value)==""){
			alert("Enter Code.");
			document.frm03150.Code.focus();
			return ;		
		}
		document.frm03150.Code.value = document.frm03150.Code.value.toUpperCase();
		if (!isNaN(document.frm03150.Code.value)) {
			alert("Code must be an alphabet.");
			document.frm03150.Code.focus();
			return ;
		}
		document.frm03150.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03150.Description.focus();">
<form name="frm03150" method="POST" action="<%=MM_editAction%>">
<h5>Update Employment Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsEmploymentType.Fields.Item("chvEmploy_Desc").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Code:</td>
		<td>
			<input type="text" name="Code" value="<%=(rsEmploymentType.Fields.Item("chrEmploy_Type").Value)%>" tabindex="2" size="1" maxlength="2" accesskey="L">
			<span style="font-size:7pt">(Must not be assigned to another employment type.)</span>			
		</td>
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
<input type="hidden" name="MM_recordId" value="<%=rsEmploymentType.Fields.Item("insEmploy_Type_Id").Value %>">
</form>
</body>
</html>
<%
rsEmploymentType.Close();
%>