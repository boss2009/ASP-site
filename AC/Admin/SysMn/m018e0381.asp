<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_edit"))=="true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");	
	var rsBuyoutProcess = Server.CreateObject("ADODB.Recordset");
	rsBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutProcess.Source = "{call dbo.cp_buyout_process("+Request.Form("MM_record_id")+",'" + Description + "'," + IsActive + ",0,'E',0)}";
	rsBuyoutProcess.CursorType = 0;
	rsBuyoutProcess.CursorLocation = 2;
	rsBuyoutProcess.LockType = 3;
	rsBuyoutProcess.Open();
	Response.Redirect("m018q0381.asp");
}

var rsBuyoutProcess = Server.CreateObject("ADODB.Recordset");
rsBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutProcess.Source = "{call dbo.cp_buyout_process("+Request.QueryString("insBuyout_process_id")+",'',0,1,'Q',0)}";
rsBuyoutProcess.CursorType = 0;
rsBuyoutProcess.CursorLocation = 2;
rsBuyoutProcess.LockType = 3;
rsBuyoutProcess.Open();
%>
<html>
<head>
	<title>Update Buyout Process Lookup</title>
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
			document.frm0381.reset();
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
		if (Trim(document.frm0381.Description.value)==""){
			alert("Enter Description.");
			document.frm0381.Description.focus();
			return ;		
		}
		document.frm0381.submit();
	}
	</script>
</head>
<body onLoad="document.frm0381.Description.focus();">
<form name="frm0381" method="POST" action="<%=MM_editAction%>">
<h5>Update Buyout Process Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F" value="<%=(rsBuyoutProcess.Fields.Item("chvBuyout_process").Value)%>"></td>
	</tr>
	<tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" tabindex="2" value="1" accesskey="L" <%=((rsBuyoutProcess.Fields.Item("bitIs_Active").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
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
<input type="hidden" name="MM_edit" value="true">
<input type="hidden" name="MM_record_id" value="<%=Request.QueryString("insBuyout_process_id")%>">
</form>
</body>
</html>
<%
rsBuyoutProcess.Close();
%>