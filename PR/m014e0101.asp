<!--------------------------------------------------------------------------
* File Name: m014e0101.asp
* Title: General Information
* Main SP: cp_update_purchase_requisition
* Description: This page updates general information of a purchase requisition.
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

var PurchaseCardNumber = "5550-0000-0104-7763";

if (String(Request.Form("MM_Update"))=="true") {
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");
	var bitOnBackOrder = ((Request.Form("OnBackOrder")=="on") ? "1":"0");
	var ReceivedBy = ((String(Request.Form("ReceivedBy"))=="undefined")?"0":String(Request.Form("ReceivedBy")));
	var DateReceived = ((String(Request.Form("DateReceived"))=="undefined")?"":String(Request.Form("DateReceived")));
	var OrderedBy = ((String(Request.Form("OrderedBy"))=="undefined")?"0":String(Request.Form("OrderedBy")));
	var DateOrdered = ((String(Request.Form("DateOrdered"))=="undefined")?"":String(Request.Form("DateOrdered")));
	var ContractPO = ((Request.Form("RequestedType")=="5")?Request.Form("ContractPO"):PurchaseCardNumber);
	var rsRequisition = Server.CreateObject("ADODB.Recordset");
	rsRequisition.ActiveConnection = MM_cnnASP02_STRING;
	rsRequisition.Source = "{call dbo.cp_Update_Purchase_Requisition("+Request.QueryString("insPurchase_Req_id")+","+Request.Form("PurchaseStatus")+","+bitOnBackOrder+","+Request.Form("RequestedType")+","+Request.Form("WorkOrderNumber")+",'"+ContractPO+"','"+Request.Form("DateRequested")+"',"+Request.Form("RequestedBy")+",'"+DateOrdered+"',"+OrderedBy+",'"+DateReceived+"',"+ReceivedBy+",0,"+Session("insStaff_id")+",'"+Notes+"',0)}";
	rsRequisition.CursorType = 0;
	rsRequisition.CursorLocation = 2;
	rsRequisition.LockType = 3;
	rsRequisition.Open();
	Response.Redirect("UpdateSuccessful2.asp?page=m014e0101.asp&insPurchase_Req_id="+Request.QueryString("insPurchase_Req_id"));
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

var rsVendor = Server.CreateObject("ADODB.Recordset");
rsVendor.ActiveConnection = MM_cnnASP02_STRING;
rsVendor.Source = "{call dbo.cp_ASP_lkup(3)}";
rsVendor.CursorType = 0;
rsVendor.CursorLocation = 2;
rsVendor.LockType = 3;
rsVendor.Open();
var rsVendor_total = 0;
for (rsVendor_total=0; !rsVendor.EOF; rsVendor.MoveNext()) {
    rsVendor_total++;
}
if (rsVendor.CursorType > 0) {
	if (!rsVendor.BOF) rsVendor.MoveFirst();
} else {
	rsVendor.Requery();
}

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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>General Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0101.reset();
			break;
		   	case 76 :
				//alert("L");
				top.window.close();
			break;
		}
	}
	</script>
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function ShowContractPO(type){
		if (type == "5") {
			openWindow('m014e0102.asp?Type=Standing&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','');
		} else {
			openWindow('m014e0102.asp?Type=Purchase&PurchaseCardNumber=<%=PurchaseCardNumber%>&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','');			
		}		
	}
	
	function Init(){
		if (document.frm0101.PurchaseStatus.value!="6") {
			if (document.frm0101.DateReceived.value=="") {
				document.frm0101.DateReceived.disabled = true;
				document.frm0101.ReceivedBy.disabled = true;					
			}
		}
		
		if (document.frm0101.PurchaseStatus.value!="3") {		
			if (document.frm0101.DateOrdered.value=="") {
				document.frm0101.DateOrdered.disabled = true;
				document.frm0101.OrderedBy.disabled = true;					
			}
		}
		//ChangeStatus();		
		document.frm0101.PurchaseStatus.focus();
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0101.Notes, 256)){
			alert("Text area cannot exceed 256 characters.");
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
		if (!CheckDate(document.frm0101.DateReceived.value)) {
			alert("Invalid Date Received.  Use (mm/dd/yyyy).");
			document.frm0101.DateReceived.focus();
			return ;
		}
		document.frm0101.submit();	
	}
	
	function ChangeStatus(){
		switch (document.frm0101.PurchaseStatus.value) {
			//complete
			case "6":
				if (document.frm0101.DateReceived.value == "") {
					document.frm0101.DateReceived.disabled = false;
					document.frm0101.ReceivedBy.disabled = false;
					document.frm0101.DateReceived.value = "<%=CurrentDate()%>";
					document.frm0101.ReceivedBy.value = "<%=Session("insStaff_id")%>";					
				}
				document.frm0101.OnBackOrder.checked = false;						
			break;
			//incomplete
			case "7":
				document.frm0101.OnBackOrder.checked = true;			
			break;
			//ordered
			case "3":
				if (document.frm0101.DateOrdered.value == "") {
					document.frm0101.DateOrdered.disabled = false;
					document.frm0101.OrderedBy.disabled = false;
					document.frm0101.DateOrdered.value = "<%=CurrentDate()%>";
					document.frm0101.OrderedBy.value = "<%=Session("insStaff_id")%>";					
				}
				document.frm0101.OnBackOrder.checked = false;				
			break;
			default:
				document.frm0101.OnBackOrder.checked = false;
			break;
		}	
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=400,height=200,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm0101" method=POST action="<%=MM_editAction%>">
<h5>General Information</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Purchase Status:</td>
		<td nowrap><select name="PurchaseStatus" style="width: 170px" onChange="ChangeStatus();" tabindex="1" accesskey="F">
		<%
		while (!rsPurchaseStatus.EOF) {
		%>
			<option value="<%=rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value%>" <%=((rsRequisition.Fields.Item("insPurchase_sts_id").Value==rsPurchaseStatus.Fields.Item("insPurchase_sts_id").Value)?"SELECTED":"")%>><%=rsPurchaseStatus.Fields.Item("chvPurchase_name").Value%> 
		<%
			rsPurchaseStatus.MoveNext();
		}
		rsPurchaseStatus.MoveFirst();
		%>
		</select></td>
		<td colspan="2" align="left" valign="top"><input type="button" value="Show Contract PO" tabindex="2" onClick="ShowContractPO(document.frm0101.RequestedType.value);" class="btnstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Work Order:</td>
		<td nowrap><select name="WorkOrderNumber" style="width: 170px" tabindex="3">
			<option value="0">None
		<%
		while (!rsWorkOrder.EOF) {
		%>
			<option value="<%=rsWorkOrder.Fields.Item("insWork_order_id").Value%>" <%=((rsRequisition.Fields.Item("insWork_order_id").Value==rsWorkOrder.Fields.Item("insWork_order_id").Value)?"SELECTED":"")%>><%=rsWorkOrder.Fields.Item("chvWork_order_no").Value%> 
		<%
			rsWorkOrder.MoveNext();
		}
		rsWorkOrder.MoveFirst();
		%>
		</select></td>	
		<td colspan="2"></td>
    </tr>
    <tr> 
		<td colspan="2"><b>Requested</b></td>
		<td colspan="2"><b>Ordered</b></td>
    </tr>
    <tr> 
		<td align="right">Date:</td>
		<td nowrap>
			<input type="text" name="DateRequested" size="11" maxlength="10" value="<%=FilterDate(rsRequisition.Fields.Item("dtsDate_Requested").Value)%>" tabindex="4" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
		<td align="right">Date:</td>
		<td nowrap>
			<input type="text" name="DateOrdered" size="11" maxlength="10" value="<%=FilterDate(rsRequisition.Fields.Item("dtsDate_Ordered").Value)%>" tabindex="8" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
    </tr>
    <tr> 
		<td align="right">By:</td>
		<td nowrap><select name="RequestedBy" style="width: 170px" tabindex="5">
			<option value="0">None
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsRequisition.Fields.Item("insReq_by_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>
        </select></td>
		<td align="right">By:</td>
		<td nowrap><select name="OrderedBy" style="width: 170px" tabindex="9">
			<option value="0" <%=((rsRequisition.Fields.Item("insOrdered_by_id").Value==0)?"SELECTED":"")%>>None
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsRequisition.Fields.Item("insOrdered_by_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>
		</select></td>
	</tr>
	<tr> 
		<td align="right">Type:</td>
		<td nowrap><select name="RequestedType" style="width: 170px" tabindex="6">
		<%
		while (!rsPurchaseType.EOF) {
		%>
			<option value="<%=rsPurchaseType.Fields.Item("insPur_type_id").Value%>" <%=((rsRequisition.Fields.Item("insRequest_type_id").Value==rsPurchaseType.Fields.Item("insPur_type_id").Value)?"SELECTED":"")%>><%=rsPurchaseType.Fields.Item("chvname").Value%> 
		<%
			rsPurchaseType.MoveNext();
		}
		rsPurchaseType.MoveFirst();
		%>
        </select></td>
		<td colspan="2"></td>
    </tr>
	<tr> 
		<td colspan="2"><b>Received</b></td>
		<td colspan="2"></td>
	</tr>
	<tr> 
		<td align="right" width="97">Date:</td>
		<td nowrap width="197"> 
			<input type="text" name="DateReceived" size="11" maxlength="10" value="<%=FilterDate(rsRequisition.Fields.Item("dtsDate_Received").Value)%>" tabindex="11" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td colspan="2"></td>
	</tr>
	<tr>		
		<td align="right" width="97">By:</td>		
		<td nowrap><select name="ReceivedBy" readonly style="width: 170px" tabindex="12">
			<option value="0" <%=((rsRequisition.Fields.Item("insReceived_by_id").Value==0)?"SELECTED":"")%>>None
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsRequisition.Fields.Item("insReceived_by_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		%>
        </select></td>
		<td colspan="2"></td>				
	</tr>	
</table>
<br>
Equipment on backorder: <input type="checkbox" name="OnBackOrder" <%=((rsRequisition.Fields.Item("bitInv_on_bk_order").Value)?"CHECKED":"")%> tabindex="13" class="chkstyle">
<br>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td valign="top">Notes:</td>
		<td valign="top"><textarea name="Notes" rows="5" cols="65" tabindex="14" accesskey="L"><%=(rsRequisition.Fields.Item("chvNote").Value)%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="top.window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Update" value="true">
<input type="hidden" name="ContractPO" value="<%=Trim(rsRequisition.Fields.Item("chvContract_PO_no").Value)%>">
</form>
</body>
</html>
<%
rsRequisition.Close();
rsVendor.Close();
rsStaff.Close();
rsPurchaseStatus.Close();
rsWorkOrder.Close();
rsPurchaseType.Close();
%>