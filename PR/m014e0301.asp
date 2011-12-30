<!--------------------------------------------------------------------------
* File Name: m014e0301.asp
* Title: Edit Inventory Received
* Main SP: cp_purchase_requisition_received
* Description: This page updates inventory received. 
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

if (String(Request.Form("MM_update"))=="true") {
 	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var rsInventoryReceived = Server.CreateObject("ADODB.Recordset");
	rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received("+Request.QueryString("insPurchase_Req_id")+","+Request.QueryString("insRqst_received_id")+","+Request.Form("QuantityReceived")+",'"+Request.Form("DateReceived")+"',"+Session("insStaff_id")+",'"+Description+"',0,'E',0)}";
	rsInventoryReceived.CursorType = 0;
	rsInventoryReceived.CursorLocation = 2;
	rsInventoryReceived.LockType = 3;
	rsInventoryReceived.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m014q0301.asp&insPurchase_Req_id="+Request.QueryString("insPurchase_Req_id"));
}

var rsInventoryReceived = Server.CreateObject("ADODB.Recordset");
rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received("+Request.QueryString("insPurchase_Req_id")+","+Request.QueryString("insRqst_received_id")+",0,'',0,'',1,'Q',0)}";
rsInventoryReceived.CursorType = 0;
rsInventoryReceived.CursorLocation = 2;
rsInventoryReceived.LockType = 3;
rsInventoryReceived.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsRequisition = Server.CreateObject("ADODB.Recordset");
rsRequisition.ActiveConnection = MM_cnnASP02_STRING;
rsRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition(0,0,'',1,"+ Request.QueryString("insPurchase_Req_id")+ ",0)}";
rsRequisition.CursorType = 0;
rsRequisition.CursorLocation = 2;
rsRequisition.LockType = 3;
rsRequisition.Open();
%>
<html>
<head>
	<title>Update Inventory Received</title>
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
		if ((!CheckDate(document.frm0301.DateReceived.value)) || (document.frm0301.DateReceived.value=="")) {
			alert("Invalid Date Received.");
			document.frm0301.DateReceived.focus();
			return ;
		}
		if (document.frm0301.QuantityReceived.value < 1) {
			alert("Enter Quantity Received.");
			document.frm0301.QuantityReceived.focus();
			return ;
		}
		var received = new Number(document.frm0301.QuantityReceived.value);
		var ordered = new Number(document.frm0301.QuantityOrdered.value);
		if (received > ordered){
			alert("Quantity received exceeds quantity ordered.");
			document.frm0301.QuantityReceived.focus();
			return ;
		}		
		document.frm0301.submit();
	}
	</script>
</head>
<body onLoad="document.frm0301.ClassName.focus();">
<form name="frm0301" method="POST" action="<%=MM_editAction%>">
<h5>Inventory Received</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Inventory Class:</td>
		<td nowrap><input type="text" name="ClassName" size="50" tabindex="1" accesskey="F" readonly value="<%=(rsInventoryReceived.Fields.Item("chvClass_name").Value)%>"></td>
    </tr>
    <tr> 
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" size="11" maxlength="10" tabindex="2" value="<%=(((FilterDate(rsInventoryReceived.Fields.Item("dtsReceived").Value)=="") || (rsInventoryReceived.Fields.Item("dtsReceived").Value==null))?CurrentDate():FilterDate(rsInventoryReceived.Fields.Item("dtsReceived").Value))%>" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Quantity Ordered:</td>
		<td nowrap><input type="text" name="QuantityOrdered" size="6" tabindex="3" onKeypress="AllowNumericOnly();" readonly value="<%=(rsInventoryReceived.Fields.Item("intQuantity_Ordered").Value)%>"></td>
    </tr>
    <tr>
		<td nowrap>Quantity Received:</td>
		<td nowrap><input type="text" name="QuantityReceived" size="6" tabindex="4" onKeypress="AllowNumericOnly();" value="<%=(rsInventoryReceived.Fields.Item("intQuantity_Received").Value)%>"></td>
    </tr>
<!--
	<tr>
		<td nowrap>Received By:</td>
		<td nowrap><select name="ReceivedBy" tabindex="5">
		<%
//		staff_id = ((rsInventoryReceived.Fields.Item("").Value != null)?rsInventoryReceived.Fields.Item("").Value:Session("insStaff_id");
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((Session("insStaff_id")==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>		
		</select></td>
	</tr>
-->	
	<tr>
		<td nowrap valign="top">Description:</td>
		<td nowrap valign="top"><textarea name="Description" tabindex="5" accesskey="L" rows="3" cols="65"><%=(rsInventoryReceived.Fields.Item("chvNotes").Value)%></textarea></td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="history.back();" tabindex="7" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="ClassID" value="<%=(rsInventoryReceived.Fields.Item("insEquip_class_id").Value)%>">
</form>
</body>
</html>
<%
rsStaff.Close();
rsInventoryReceived.Close();
rsRequisition.Close();
%>