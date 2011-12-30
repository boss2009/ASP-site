<!--------------------------------------------------------------------------
* File Name: m014a0501.asp
* Title: New Backorder Received
* Main SP: cp_PR_BackOrder_Rx
* Description: This page is used to insert a record when a backorder has been
* received.
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
	var rsPurchaseStatus = Server.CreateObject("ADODB.Recordset");
	rsPurchaseStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchaseStatus.Source = "{call dbo.cp_PR_BackOrder_Rx(0,"+Request.QueryString("insPurchase_Req_id")+",'"+ Request.Form("BackOrderReceivedDate") + "',"+Request.Form("InventoryClass")+","+Request.Form("QuantityReceived")+","+Request.Form("ReceivedBy")+",'A',0,0)}";
	rsPurchaseStatus.CursorType = 0;
	rsPurchaseStatus.CursorLocation = 2;
	rsPurchaseStatus.LockType = 3;
	rsPurchaseStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsOutstandingInventory = Server.CreateObject("ADODB.Recordset");
rsOutstandingInventory.ActiveConnection = MM_cnnASP02_STRING;
rsOutstandingInventory.Source = "{call dbo.cp_PR_BORx_EqpClass("+Request.QueryString("insPurchase_Req_id")+",0)}";
rsOutstandingInventory.CursorType = 0;
rsOutstandingInventory.CursorLocation = 2;
rsOutstandingInventory.LockType = 3;
rsOutstandingInventory.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>New Backorder Received</title>
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
		if ((!IsID(document.frm0501.QuantityReceived.value)) || (Number(document.frm0501.QuantityReceived.value<1))) {
			alert("Invalid Quantity.");
			document.frm0501.QuantityReceived.focus();
			return ;
		}
		if (!CheckDate(document.frm0501.BackOrderReceivedDate.value)) {
			alert("Invalid Backorder Received Date.");
			document.frm0501.BackOrderReceivedDate.focus();
			return ;
		}
		if (document.frm0501.InventoryClass.length > 1){
			if (Number(document.frm0501.QuantityReceived.value) > Number(document.frm0501.MaximumQuantity[document.frm0501.InventoryClass.selectedIndex].value)) {
				alert("Quantity cannot exceed "+ document.frm0501.MaximumQuantity[document.frm0501.InventoryClass.selectedIndex].value+".");
				document.frm0501.QuantityReceived.focus();
				return;
			}		
		} else {
			if (Number(document.frm0501.QuantityReceived.value) > Number(document.frm0501.MaximumQuantity.value)) {
				alert("Quantity cannot exceed "+ document.frm0501.MaximumQuantity.value)
				document.frm0501.QuantityReceived.focus();
				return;
			}
		}
		document.frm0501.submit();
	}
	</script>	
</head>
<body onLoad="<%=((!rsOutstandingInventory.EOF)?"document.frm0501.BackOrderReceivedDate.focus();":"")%>">
<form name="frm0501" method="POST" action="<%=MM_editAction%>">
<h5>New Backorder Received</h5>
<hr>
<% 
if (rsOutstandingInventory.EOF) {
%>
<i>There are currently no outstanding backorders.</i><br>
<br>
<input type="button" value="Close" onClick="window.close();" tabindex="1" class="btnstyle">
<%
} else {
%>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Backorder Received Date:</td>
		<td nowrap>
			<input type="text" name="BackOrderReceivedDate" maxlength="10" size="11" tabindex="1" accesskey="F"  value=<%=CurrentDate()%> onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Inventory Class:</td>
		<td nowrap><select name="InventoryClass" tabindex="2">
		<%
		while (!rsOutstandingInventory.EOF) {
		%>
			<option value="<%=rsOutstandingInventory.Fields.Item("insEquip_class_id").Value%>"><%=rsOutstandingInventory.Fields.Item("chvEqp_Class").Value%> 
		<%
			rsOutstandingInventory.MoveNext();
		}
		rsOutstandingInventory.MoveFirst();
		%>		
		</select></td>		
	</tr>
	<tr>
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="QuantityReceived" value="1" size="3" tabindex="3" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap>Received By:</td>
		<td nowrap><select name="ReceivedBy" tabindex="4" accesskey="L">
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((Session("insStaff_id")==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
			<%
				rsStaff.MoveNext();
			}
			%>
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="6" class="btnstyle"></td>
    </tr>
</table>
<%
}
%>
<input type="hidden" name="MM_insert" value="true">
<%
while (!rsOutstandingInventory.EOF) {
%>
<input type="hidden" name="MaximumQuantity" value="<%=rsOutstandingInventory.Fields.Item("intTotal").Value%>">
<%
	rsOutstandingInventory.MoveNext();
}
%>
</form>
</body>
</html>
<%
rsStaff.Close();
rsOutstandingInventory.Close();
%>