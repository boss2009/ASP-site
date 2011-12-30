<!--------------------------------------------------------------------------
* File Name: m012e0203.asp
* Title: Edit Inventory Request
* Main SP: cp_school_ref_eqp_requested
* Description: This page updates inventory requested.
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

if (String(Request("MM_update")) == "true") {	
	var Comments = String(Request.Form("Comments")).replace(/'/g, "''");
	var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
	rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryRequest.Source = "{call dbo.cp_school_ref_eqp_requested("+Request.QueryString("intEqpRequest_id")+","+Request.QueryString("intReferral_id")+","+Request.Form("ClassID")+",1,"+Request.Form("Quantity")+",'"+Comments+"',0,'E',0)}";
	rsInventoryRequest.CursorType = 0;
	rsInventoryRequest.CursorLocation = 2;
	rsInventoryRequest.LockType = 3;
	rsInventoryRequest.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012q0203.asp&intReferral_id="+Request.QueryString("intReferral_id"));
}

var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequest.Source = "{call dbo.cp_school_ref_eqp_requested("+Request.QueryString("intEqpRequest_id")+",0,0,0,0,'',1,'Q',0)}";
rsInventoryRequest.CursorType = 0;
rsInventoryRequest.CursorLocation = 2;
rsInventoryRequest.LockType = 3;
rsInventoryRequest.Open();
%>
<html>
<head>
	<title>Update Inventory Request</title>
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
				top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';
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
		if (document.frm0201.Quantity.value=="") document.frm0201.Quantity.value=0;
		document.frm0203.MM_update.value="true";
		document.frm0203.submit();
	}
	
	function Init(){
		document.frm0203.ListClass.focus();
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm0203" method="POST" action="<%=MM_editAction%>">
<h5>Inventory Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Inventory Class:</td>
		<td nowrap> 
			<input type="text" name="ClassName" value="<%=((String(Request.QueryString("ClassName"))=="undefined")?rsInventoryRequest.Fields.Item("chvEqp_Class_Name").Value:Request.QueryString("ClassName"))%>" size="40" tabindex="1" accesskey="F" readonly>
			<input type="button" name="ListClass" value="List Class" tabindex="2" onClick="openWindow('m012p01FS.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>','');" class="btnstyle">
		</td>
    </tr>
	<tr> 
		<td valign="top">Comments:</td>
		<td valign="top"><textarea name="Comments" rows="5" cols="65" tabindex="3"><%=(rsInventoryRequest.Fields.Item("chvComments").Value)%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="4" accesskey="L" onKeypress="AllowNumericOnly();" value="<%=(rsInventoryRequest.Fields.Item("insQuantity").Value)%>"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>'" tabindex="6" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ClassID" value="<%=((String(Request.QueryString("ClassID"))=="undefined")?rsInventoryRequest.Fields.Item("insClass_bundle_id").Value:Request.QueryString("ClassID"))%>">
<input type="hidden" name="MM_update" value="false">
</form>
</body>
</html>
<%
rsInventoryRequest.Close();
%>