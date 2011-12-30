<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true") {
	var rsWaiver = Server.CreateObject("ADODB.Recordset");
	rsWaiver.ActiveConnection = MM_cnnASP02_STRING;
	rsWaiver.Source = "{call dbo.cp_Insert_waiver("+ Request.QueryString("intAdult_id") + ",'"+ Request.Form("DateReceived") + "',0)}";
	rsWaiver.CursorType = 0;
	rsWaiver.CursorLocation = 2;
	rsWaiver.LockType = 3;
	rsWaiver.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New Waiver</title>
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
	function Save(){
		if (!CheckDate(document.frm0604.DateReceived.value)) {
			alert("Invalid Date Received.");
			document.frm0604.DateReceived.focus();
			return ;
		}
		document.frm0604.submit();
		document.frm0604.btnSave.disabled = true;
	}
	</script>
</head>
<body onLoad="javascript:document.frm0604.DateReceived.focus()" >
<form name="frm0604" method="POST" action="<%=MM_editAction%>">
<h5>New Waiver</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=CurrentDate()%>" size="11" maxlength="10" accesskey="F" tabindex="1" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><input type="button" name="btnSave" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
