<!--------------------------------------------------------------------------
* File Name: m014a01j.asp
* Title: Save to Desktop
* Main SP: cp_Insert_Desktop
* Description: This page adds a purchase requisition to current user's desktop.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var rsDesktop = Server.CreateObject("ADODB.Recordset");
	rsDesktop.ActiveConnection = MM_cnnASP02_STRING;
	rsDesktop.Source = "{call dbo.cp_Insert_DeskTop(" + Session("insStaff_id") + "," + Request.Form("insPurchase_Req_id") + ",29,'" + CurrentDate() + "',0)}";
	rsDesktop.CursorType = 0;
	rsDesktop.CursorLocation = 2;
	rsDesktop.LockType = 3;
	rsDesktop.Open();
	Response.Redirect("InsertSuccessful.html");	
}

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
			document.frm01j.submit();
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
<form name="frm01j" method="POST" action="<%=MM_editAction%>">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Save To:</td>
		<td><select name="SaveTo" tabindex="1" accesskey="F">
			<% 
			while (!rsModule.EOF) {
			%>
				<option value="<%=(rsModule.Fields.Item("insId").Value)%>" <%=((rsModule.Fields.Item("insId").Value == 29)?"SELECTED":"")%> ><%=(rsModule.Fields.Item("ncvMODname").Value)%>
			<%
				rsModule.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td>Date Copied:</td>
		<td>
			<input type="text" name="DateCopied" value="<%=CurrentDate()%>" size="10" readonly tabindex="2" accesskey="L">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>						
		</td>
    </tr>
    <tr> 
		<td colspan="2" align="center">
			<input type="submit" value="Save to Desktop" class="btnstyle" accesskey="3">
			<input type="button" value="Cancel" onClick="self.close();" accesskey="4" class="btnstyle">
		</td>
    </tr>
</table>
<input type="hidden" name="insPurchase_Req_id" value="<%=Request.QueryString("insPurchase_Req_id")%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsModule.Close();
%>