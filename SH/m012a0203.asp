<!--------------------------------------------------------------------------
* File Name: m012a0401.asp
* Title: New Inventory Request
* Main SP: cp_school_ref_eqp_requested
* Comments: This page validates the equipment class first then inserts a 
* new inventory request.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_Insert")) == "true") {	
	var Comments = String(Request.Form("Comments")).replace(/'/g, "''");			
	var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
	rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryRequest.Source = "{call dbo.cp_school_ref_eqp_requested(0,"+Request.QueryString("intReferral_id")+","+Request.Form("ClassID")+",1,"+Request.Form("Quantity")+",'"+Comments+"',0,'A',0)}";
	rsInventoryRequest.CursorType = 0;
	rsInventoryRequest.CursorLocation = 2;
	rsInventoryRequest.LockType = 3;
	rsInventoryRequest.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Inventory Request</title>
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
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   
	
	function Save(){
		if (document.frm0203.ClassID.value==0) {
			alert("Select a class.");
			document.frm0203.ListClass.focus();
			return ;
		}
		document.frm0203.MM_Insert.value="true";
		document.frm0203.submit();
	}
	
	function Init(){
		document.frm0203.ListClass.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0203" method="POST" action="<%=MM_editAction%>">
<h5>New Inventory Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Inventory Class:</td>
		<td nowrap>
			<input type="text" name="ClassName" size="65" value="<%=Request.QueryString("ClassName")%>" tabindex="1" accesskey="F" readonly>
			<input type="button" name="ListClass" value="List Class" tabindex="2" onClick="openWindow('m012p01FS.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>','');" class="btnstyle">
		</td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="10" cols="65" tabindex="3"></textarea></td>
	</tr>
	<tr>
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="4" accesskey="F" value="0" onKeypress="AllowNumericOnly();"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="6" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<input type="hidden" name="MM_Insert" value="false">
</form>
</body>
</html>