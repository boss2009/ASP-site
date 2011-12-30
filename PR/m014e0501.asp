<!--------------------------------------------------------------------------
* File Name: m014e0501.asp
* Title: Backorder Received Date Edit
* Main SP: cp_PR_BORx_EqpClass, cp_PR_BackOrder_Rx
* Description: Edit page for backorder received.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JavaScript"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true") {
	var rsBackOrder = Server.CreateObject("ADODB.Recordset");
	rsBackOrder.ActiveConnection = MM_cnnASP02_STRING;
	rsBackOrder.Source = "{call dbo.cp_PR_BackOrder_Rx("+Request.QueryString("intBack_Ord_id")+","+Request.QueryString("insPurchase_Req_id")+",'"+Request.Form("BackOrderReceivedDate")+"',"+Request.Form("ClassID")+","+Request.Form("QuantityReceived")+","+Request.Form("ReceivedBy")+",'E',0,0)}";	
	rsBackOrder.CursorType = 0;
	rsBackOrder.CursorLocation = 2;
	rsBackOrder.LockType = 3;
	rsBackOrder.Open();
	Response.Redirect("m014q0501.asp?insPurchase_Req_id="+Request.QueryString("insPurchase_Req_id"));
}

var rsBackOrder = Server.CreateObject("ADODB.Recordset");
rsBackOrder.ActiveConnection = MM_cnnASP02_STRING;
rsBackOrder.Source = "{call dbo.cp_PR_BackOrder_Rx("+Request.QueryString("intBack_Ord_id")+","+Request.QueryString("insPurchase_Req_id")+",'',0,0,0,'Q',1,0)}";
rsBackOrder.CursorType = 0;
rsBackOrder.CursorLocation = 2;
rsBackOrder.LockType = 3;
rsBackOrder.Open();

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
	<title>Backorder Received</title>
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
			case 85:
				//alert("U");
				document.frm0501.reset();
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
		if (Trim(document.frm0501.BackOrderReceivedDate.value)==""){
			alert("Enter Backorder Received Date.");
			document.frm0501.BackOrderReceivedDate.focus();
			return ;		
		}
		document.frm0501.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0501.BackOrderReceivedDate.focus();">
<form name="frm0501" method="POST" action="<%=MM_editAction%>">
<h5>Backorder Received</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Backorder Received Date:</td>
		<td nowrap>
			<input type="text" name="BackOrderReceivedDate" value="<%=FilterDate(rsBackOrder.Fields.Item("dtsBack_Ord_rx").Value)%>" maxlength="10" size="11" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Inventory Received:</td>
		<td nowrap><input type="text" name="InventoryReceived" value="<%=(rsBackOrder.Fields.Item("chvEqp_Class").Value)%>" size="40" style="border: none" tabindex="2" readonly></td>	
	</tr>
	<tr>
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="QuantityReceived" value="<%=(rsBackOrder.Fields.Item("intQuantity").Value)%>" size="3" style="border: none" tabindex="3" readonly></td>
	</tr>
	<tr>
		<td nowrap>Received By:</td>
		<td nowrap><select name="ReceivedBy" tabindex="4" accesskey="L">
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsBackOrder.Fields.Item("insBORx_by_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
			<%
				rsStaff.MoveNext();
			}
			rsStaff.MoveFirst();
			%>
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="7" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="ClassID" value="<%=(rsBackOrder.Fields.Item("insEquip_class_id").Value)%>">
</form>
</body>
</html>
<%
rsBackOrder.Close();
rsStaff.Close();
%>