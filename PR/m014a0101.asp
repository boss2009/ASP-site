<!--------------------------------------------------------------------------
* File Name: m014a0101.asp
* Title: New Purchase Requisition
* Main SP: cp_Insert_Purchase_Requisition
* Description: This page inserts a new purchase requisition by saving the
* general information then redirects the user to edit page.
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

if (String(Request.Form("MM_Insert"))=="true") {
//	var OnBackOrder = ((Request.Form("OnBackOrder")=="on")?"1":"0");
//	var Received = ((Request.Form("Received")=="on")?"1":"0");	
//	var ReceivedBy = ((Request.Form("Received")=="on")?String(Request.Form("ReceivedBy")):"0");
//	var DateReceived = ((Request.Form("Received")=="on")?String(Request.Form("DateReceived")):"");
	var cmdInsertPR = Server.CreateObject("ADODB.Command");
	cmdInsertPR.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertPR.CommandText = "dbo.cp_Insert_Purchase_Requisition";
	cmdInsertPR.CommandType = 4;
	cmdInsertPR.CommandTimeout = 0;
	cmdInsertPR.Prepared = true;
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insPurchase_sts_id", 2, 1,1,Request.Form("PurchaseStatus")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@bitInv_on_bk_order", 2, 1,1,0));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insRequest_type_id", 2, 1,1,Request.Form("RequestedType")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insWork_order_id", 2, 1,10,Request.Form("WorkOrderNumber")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@chvContract_PO_no", 200, 1,20,""));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@dtsDate_Requested", 200, 1,30,Request.Form("DateRequested")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insReq_by_id", 2, 1,1,Request.Form("RequestedBy")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@dtsDate_Ordered", 200, 1,30,Request.Form("DateOrdered")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insOrdered_by_id", 2, 1,1,Request.Form("OrderedBy")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@dtsDate_Received", 200, 1,30,""));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insReceived_by_id", 2, 1,1,0));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insVendor_id", 2, 1,1,0));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@bitIs_received", 2, 1,1,0));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@intBack_Ord_id", 3, 1,10,0));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insCreator_user_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@chvNote", 200, 1,4000,Request.Form("Notes")));
	cmdInsertPR.Parameters.Append(cmdInsertPR.CreateParameter("@insPurchase_Req_id", 2, 2));
	cmdInsertPR.Execute();
	Response.Redirect("m014FS3.asp?insPurchase_Req_id="+cmdInsertPR.Parameters.Item("@insPurchase_Req_id").Value);
}

var rsPurchaseType = Server.CreateObject("ADODB.Recordset");
rsPurchaseType.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseType.Source = "{call dbo.cp_ASP_lkup(55)}";
rsPurchaseType.CursorType = 0;
rsPurchaseType.CursorLocation = 2;
rsPurchaseType.LockType = 3;
rsPurchaseType.Open();

var rsWorkOrder = Server.CreateObject("ADODB.Recordset");
rsWorkOrder.ActiveConnection = MM_cnnASP02_STRING;
rsWorkOrder.Source = "{call dbo.cp_ASP_lkup(59)}";
rsWorkOrder.CursorType = 0;
rsWorkOrder.CursorLocation = 2;
rsWorkOrder.LockType = 3;
rsWorkOrder.Open();

var rsPurchaseStatus = Server.CreateObject("ADODB.Recordset");
rsPurchaseStatus.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseStatus.Source = "{call dbo.cp_ASP_lkup(54)}";
rsPurchaseStatus.CursorType = 0;
rsPurchaseStatus.CursorLocation = 2;
rsPurchaseStatus.LockType = 3;
rsPurchaseStatus.Open();

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
	<title>New Purchase Requisition</title>
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
	function Init(){
		document.frm0101.PurchaseStatus.focus();
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0101.Notes, 4000)) {
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (!CheckDate(document.frm0101.DateRequested.value)) {
			alert("Invalid Date Requested.  Use (mm/dd/yyyy).");
			document.frm0101.DateRequested.focus();
			return ;
		}
		
		if (!CheckDate(document.frm0101.DateOrdered.value)) {
			alert("Invalid Date Ordered.  Use (mm/dd/yyyy).");
			document.frm0101.DateOrdered.focus();
			return ;
		}
		
//		if (!CheckDate(document.frm0101.DateReceived.value)) {
//			alert("Invalid Date Received.  Use (mm/dd/yyyy).");
//			document.frm0101.DateReceived.focus();
//			return ;
//		}
		document.frm0101.submit();	
	}
	
	function ChangeStatus() {
		if (document.frm0101.PurchaseStatus.value=="3") {
			document.frm0101.DateOrdered.value="<%=CurrentDate()%>";
			document.frm0101.OrderedBy.value="<%=Session("insStaff_id")%>";
		}
	}
	</script>
</head>
<body onLoad="Init();">
<form name="frm0101" method=POST action="<%=MM_editAction%>">
<h5>New Purchase Requisition</h5>
<hr>
<table cellpadding="1" cellspacing="1">
   	<tr> 
		<td nowrap>Purchase Status:</td>
		<td nowrap><select name="PurchaseStatus" onChange="ChangeStatus();" style="width: 170px" tabindex="1" accesskey="F">
	<%
	while (!rsPurchaseStatus.EOF) {
		if ((rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value=="3") || (rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value=="1")) {
	%>
			<option value="<%=rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value%>" <%=((rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value=="1")?"SELECTED":"")%>><%=rsPurchaseStatus.Fields.Item("chvPurchase_name").Value%> 
	<%
		}
		rsPurchaseStatus.MoveNext();
	}
	rsPurchaseStatus.MoveFirst();
	%>
		</select></td>
		<td colspan="2"></td>
	</tr>
	<tr> 
		<td nowrap>Work Order:</td>
		<td nowrap><select name="WorkOrderNumber" style="width: 170px" tabindex="2">
			<option value="0">None 
		<%
		while (!rsWorkOrder.EOF) {
		%>
			<option value="<%=rsWorkOrder.Fields.Item("insWork_order_id").Value%>" <%=((rsWorkOrder.Fields.Item("insWork_order_id").Value=="1")?"SELECTED":"")%>><%=rsWorkOrder.Fields.Item("chvWork_order_no").Value%> 
		<%
			rsWorkOrder.MoveNext();
		}
		rsWorkOrder.MoveFirst();
		%>
		</select></td>
		<td colspan="2"></td>
	</tr>
	<tr height="15"> 
		<td colspan="4"></td>
	</tr>
	<tr> 
		<td colspan="2"><b>Requested</b></td>
		<td colspan="2"><b>Ordered</b></td>
	</tr>
	<tr> 
		<td nowrap align="right">Date:</td>
		<td nowrap>
			<input type="text" name="DateRequested" size="11" maxlength="10" value="<%=CurrentDate()%>" tabindex="5" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap align="right">Date:</td>
		<td nowrap> 
			<input type="text" name="DateOrdered" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap align="right">By:</td>
		<td nowrap><select name="RequestedBy" style="width: 170px" tabindex="6">
			<option value="0">None 
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==Session("insStaff_id"))?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>
		</select></td>
		<td nowrap align="right">By:</td>
		<td nowrap><select name="OrderedBy" style="width: 170px" tabindex="9">
			<option value="0">None 
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>"><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap align="right">Type:</td>
		<td nowrap><select name="RequestedType" style="width: 170px" tabindex="7">
		<%
		while (!rsPurchaseType.EOF) {
		%>
			<option value="<%=rsPurchaseType.Fields.Item("insPur_type_id").Value%>" <%=((rsPurchaseType.Fields.Item("insPur_type_id").Value=="5")?"SELECTED":"")%>><%=rsPurchaseType.Fields.Item("chvname").Value%> 
		<%
			rsPurchaseType.MoveNext();
		}
		rsPurchaseType.MoveFirst();
		%>
		</select></td>
		<td width="80"></td>
		<td width="258"></td>
	</tr>
<!--	
	<tr> 
		<td nowrap colspan="2"><b>Received</b></td>
		<td nowrap colspan="2"></td>
    </tr>
	<tr> 
		<td nowrap align="right" width="97">Date:</td>
		<td nowrap width="197"> 
			<input type="text" name="DateReceived" size="11" maxlength="10" disabled tabindex="11" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td colspan="2"></td>
	</tr>
	<tr> 
		<td nowrap align="right">By:</td>
		<td nowrap><select name="ReceivedBy" disabled style="width: 170px" tabindex="12">
			<option value="0">None 
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>"><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		%>
		</select></td>
		<td colspan="2"></td>
    </tr>
-->	
</table>
<!--
<br>
Equipment on backorder: <input type="checkbox" name="OnBackOrder" tabindex="13" class="chkstyle">
<br>
-->	
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" rows="4" cols="65" tabindex="14" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="16" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>
<%
rsStaff.Close();
rsPurchaseStatus.Close();
rsWorkOrder.Close();
rsPurchaseType.Close();
%>