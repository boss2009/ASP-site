<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var rsDesktop = Server.CreateObject("ADODB.Recordset");
	rsDesktop.ActiveConnection = MM_cnnASP02_STRING;
	rsDesktop.Source = "{call dbo.cp_Insert_DeskTop(" + Session("insStaff_id") + "," + Request.Form("intEquip_Set_id") + ",18,'" + CurrentDate() + "',0)}";
	rsDesktop.CursorType = 0;
	rsDesktop.CursorLocation = 2;
	rsDesktop.LockType = 3;
	rsDesktop.Open();
	Response.Redirect("InsertSuccessful.html");	
}

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + Request.QueryString("intEquip_Set_id") + ",0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();

var rsModule = Server.CreateObject("ADODB.Recordset");
rsModule.ActiveConnection = MM_cnnASP02_STRING;
rsModule.Source = "{call dbo.cp_ASP_Lkup(700)}";
rsModule.CursorType = 0;
rsModule.CursorLocation = 2;
rsModule.LockType = 3;
rsModule.Open();
%>
<html>
<head>
	<title>Save to Desktop</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				document.frm03j01.submit();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
</head>
<body onLoad="window.focus();">
<form name="frm03j01" method="POST" action="<%=MM_editAction%>">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Inventory Name:</td>
		<td><input type="text" name="InventoryName" value="<%=(rsInventory.Fields.Item("chvInventory_Name").Value)%>" readonly size="35" tabindex="1" accesskey="F" ></td>
    </tr>	
    <tr> 
		<td>Save To:</td>
		<td><select name="SaveTo" tabindex="2">
			<% 
			while (!rsModule.EOF) {
			%>
				<option value="<%=(rsModule.Fields.Item("insId").Value)%>" <%=((rsModule.Fields.Item("insId").Value == 22)?"SELECTED":"")%> ><%=(rsModule.Fields.Item("ncvMODname").Value)%>
			<%
				rsModule.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td>Date Copied:</td>
		<td>
			<input type="text" name="DateCopied" value="<%=CurrentDate()%>" size="10" readonly tabindex="3" accesskey="L">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>						
		</td>
    </tr>
    <tr> 
		<td colspan="2" align="center">
			<input type="submit" value="Save to Desktop" class="btnstyle">
			<input type="button" value="Cancel" onClick="self.close();" class="btnstyle">
		</td>
    </tr>
</table>
<input type="hidden" name="intEquip_Set_id" value="<%=rsInventory.Fields.Item("intEquip_Set_id").Value%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsInventory.Close();
rsModule.Close();
%>