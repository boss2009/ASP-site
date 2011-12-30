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
	var AllowLaborCost = ((Request.Form("AllowLaborCost")=="on")?"1":"0");	
	var rsRepairReason = Server.CreateObject("ADODB.Recordset");
	rsRepairReason.ActiveConnection = MM_cnnASP02_STRING;
	rsRepairReason.Source = "{call dbo.cp_eqsrv_repair_reason("+ Request.Form("MM_recordId") + ",'" + Description + "',"+AllowLaborCost+",0,0,'E',0)}";
	rsRepairReason.CursorType = 0;
	rsRepairReason.CursorLocation = 2;
	rsRepairReason.LockType = 3;
	rsRepairReason.Open();
	Response.Redirect("m018q03151.asp");
}

var rsRepairReason = Server.CreateObject("ADODB.Recordset");
rsRepairReason.ActiveConnection = MM_cnnASP02_STRING;
rsRepairReason.Source = "{call dbo.cp_eqsrv_repair_reason("+ Request.QueryString("insRepair_Reason_Id") + ",'',0,0,1,'Q',0)}";
rsRepairReason.CursorType = 0;
rsRepairReason.CursorLocation = 2;
rsRepairReason.LockType = 3;
rsRepairReason.Open();
%>
<html>
<head>
	<title>Update Reason for Repair Lookup</title>
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
			document.frm03151.reset();
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
		if (Trim(document.frm03151.Description.value)==""){
			alert("Enter Description.");
			document.frm03151.Description.focus();
			return ;		
		}
		document.frm03151.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03151.Description.focus();">
<form name="frm03151" method="POST" action="<%=MM_editAction%>">
<h5>Update Reason for Repair Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowraps><input type="text" name="Description" value="<%=(rsRepairReason.Fields.Item("chvRepair_Reason_Desc").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
	<tr>
		<td nowrap>Allow Labor Cost</td>
		<td nowrap>
			<input type="checkbox" name="AllowLaborCost" class="chkstyle" <%=((rsRepairReason.Fields.Item("bitIs_Labor_cost").Value=="1")?"CHECKED":"")%> tabindex="2" accesskey="L">
			<span style="font-size:7pt">(If checked, repair involves a labor cost.)</span>
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
<input type="hidden" name="MM_recordId" value="<%=rsRepairReason.Fields.Item("insRepair_Reason_Id").Value %>">
</form>
</body>
</html>
<%
rsRepairReason.Close();
%>