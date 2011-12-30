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
	var rsEquipUserType = Server.CreateObject("ADODB.Recordset");
	var IsActive = ((Request.Form("IsActive")=="1")?"1":"0");
	var IsBuyout = ((Request.Form("IsBuyout")=="1")?"1":"0");	
	rsEquipUserType.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipUserType.Source = "{call dbo.cp_eq_user_type2("+Request.QueryString("insEq_user_type")+",'" + Description + "'," + IsActive + "," + IsBuyout + ",0,'E',0)}";
	rsEquipUserType.CursorType = 0;
	rsEquipUserType.CursorLocation = 2;
	rsEquipUserType.LockType = 3;
	rsEquipUserType.Open();
	Response.Redirect("m018q03122.asp");
}
var rsEquipUserType = Server.CreateObject("ADODB.Recordset");
rsEquipUserType.ActiveConnection = MM_cnnASP02_STRING;
rsEquipUserType.Source = "{call dbo.cp_eq_user_type2("+Request.QueryString("insEq_user_type")+",'',0,0,1,'Q',0)}";
rsEquipUserType.CursorType = 0;
rsEquipUserType.CursorLocation = 2;
rsEquipUserType.LockType = 3;
rsEquipUserType.Open();
%>
<html>
<head>
	<title>Update Equipment User Type Lookup</title>
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
				document.frm03122.reset();
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
		if (Trim(document.frm03122.Description.value)==""){
			alert("Enter Description.");
			document.frm03122.Description.focus();
			return ;		
		}
		document.frm03122.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03122.Description.focus();">
<form name="frm03122" method="POST" action="<%=MM_editAction%>">
<h5>Update Equipment User Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsEquipUserType.Fields.Item("chvEq_user_type").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" value="1" tabindex="2" <%=((rsEquipUserType.Fields.Item("bitIs_active").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
    </tr>	
    <tr> 
		<td nowrap>Is Buyout:</td>
		<td nowrap><input type="checkbox" name="IsBuyout" value="1" tabindex="3" accesskey="L" <%=((rsEquipUserType.Fields.Item("bitIs_Buyout_Applic").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
    </tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsEquipUserType.Fields.Item("insEq_user_type").Value %>">
</form>
</body>
</html>
<%
rsEquipUserType.Close();
%>