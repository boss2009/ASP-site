<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var IsActive = ((Request.Form("IsActive")=="1")?"1":"0");
	var rsPurchaseStatus = Server.CreateObject("ADODB.Recordset");
	rsPurchaseStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchaseStatus.Source = "{call dbo.cp_purchase_status("+Request.QueryString("insPurchase_sts_id")+",'" + Description + "'," + IsActive + ",'E',0,0)}";
	rsPurchaseStatus.CursorType = 0;
	rsPurchaseStatus.CursorLocation = 2;
	rsPurchaseStatus.LockType = 3;
	rsPurchaseStatus.Open();
	Response.Redirect("m018q0354.asp");
}

var rsPurchaseStatus = Server.CreateObject("ADODB.Recordset");
rsPurchaseStatus.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseStatus.Source = "{call dbo.cp_purchase_status("+Request.QueryString("insPurchase_sts_id")+",'',0,'Q',1,0)}";
rsPurchaseStatus.CursorType = 0;
rsPurchaseStatus.CursorLocation = 2;
rsPurchaseStatus.LockType = 3;
rsPurchaseStatus.Open();
%>
<html>
<head>
	<title>Update Purchase Status Lookup</title>
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
			document.frm0354.reset();
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
		if (Trim(document.frm0354.Description.value)==""){
			alert("Enter Description.");
			document.frm0354.Description.focus();
			return ;		
		}
		document.frm0354.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0354.Description.focus();">
<form name="frm0354" method="POST" action="<%=MM_editAction%>">
<h5>Update Purchase Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsPurchaseStatus.Fields.Item("chvPurchase_name").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" tabindex="2" value="1" accesskey="L" <%=((rsPurchaseStatus.Fields.Item("bitis_Active").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value %>">
</form>
</body>
</html>
<%
rsPurchaseStatus.Close();
%>