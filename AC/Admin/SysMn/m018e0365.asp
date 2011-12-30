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
	var IsActive = ((Request.Form("IsActive")=="1")?"1":"0");
	var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
	rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
	rsShippingMethod.Source = "{call dbo.cp_shipping_method2("+Request.Form("MM_recordId")+",'"+Description+"',"+IsActive+",0,'E',0)}";
	rsShippingMethod.CursorType = 0;
	rsShippingMethod.CursorLocation = 2;
	rsShippingMethod.LockType = 3;
	rsShippingMethod.Open();
	Response.Redirect("m018q0365.asp");
}

var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
rsShippingMethod.Source = "{call dbo.cp_shipping_method2("+ Request.QueryString("intship_method_id") + ",'',0,1,'Q',0)}";
rsShippingMethod.CursorType = 0;
rsShippingMethod.CursorLocation = 2;
rsShippingMethod.LockType = 3;
rsShippingMethod.Open();
%>
<html>
<head>
	<title>Update Shipping Method Lookup</title>
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
			document.frm0365.reset();
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
		if (Trim(document.frm0365.Description.value)==""){
			alert("Enter Description.");
			document.frm0365.Description.focus();
			return ;		
		}
		document.frm0365.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0365.Description.focus();">
<form name="frm0365" method="POST" action="<%=MM_editAction%>">
<h5>Update Shipping Method Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsShippingMethod.Fields.Item("chvname").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" <%=((rsShippingMethod.Fields.Item("bitis_active").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsShippingMethod.Fields.Item("intship_method_id").Value %>">
</form>
</body>
</html>
<%
rsShippingMethod.Close();
%>