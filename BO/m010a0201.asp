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
	var OnBackOrder = ((Request.Form("OnBackOrder")=="on") ? "1":"0");		
	var ClassID = ((String(Request.Form("ClassID"))!="undefined")?Request.Form("ClassID"):0);
	var Quantity = ((String(Request.Form("Quantity"))=="")?Request.Form("Quantity"):0);
	var ListUnitCost = ((String(Request.Form("ListUnitCost"))=="")?Request.Form("ListUnitCost"):0);	
	var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
	rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryRequest.Source = "{call dbo.cp_buyout_eqp_requested(0,"+Request.QueryString("intBuyout_req_id")+","+Request.Form("ClassID")+","+Request.Form("ClassBundle")+","+Request.Form("Quantity")+","+Request.Form("ListUnitCost")+","+OnBackOrder+",'"+Comments+"',0,'A',0)}";
	rsInventoryRequest.CursorType = 0;
	rsInventoryRequest.CursorLocation = 2;
	rsInventoryRequest.LockType = 3;
	rsInventoryRequest.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Equipment Request</title>
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
		if (document.frm0201.ClassID.value==0) {
			alert("Select a Class.");
			document.frm0201.ListClass.focus();
			return ;
		}
		if ((isNaN(document.frm0201.Quantity.value)) || (document.frm0201.Quantity.value < 1)){
			alert("Invalid Quantity.");
			document.frm0201.Quantity.focus();
			return ;
		}
		document.frm0201.MM_Insert.value="true";
		document.frm0201.submit();
	}
	
	function ViewAcc(){
		if (document.frm0201.ClassID.value > 0) temp=window.showModalDialog("m010pop.asp?ClassID="+document.frm0201.ClassID.value,"","dialogHeight: 200px; dialogWidth: 375px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");	
	}
	
	function CalculateTotal(){
		var temp = new Number("0");
		var temp1 = new Number(document.frm0201.Quantity.value);
		var temp2 = new Number(document.frm0201.ListUnitCost.value);
		temp = Math.round(temp1 * temp2 * 100)/100;
		document.frm0201.Total.value= FormatCurrency(temp.toString());
	}
	
	function Init(){
		CalculateTotal();	
		document.frm0201.ListClass.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0201" method="POST" action="<%=MM_editAction%>">
<h5>New Equipment Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Class/Bundle:</td>
		<td nowrap>
			<input type="text" name="ClassName" value="" tabindex="1" size="60" accesskey="F" readonly >
			<input type="button" name="ListClass" value="List" tabindex="2" onClick="openWindow('m010p01FS.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>','');" class="btnstyle">
			<input type="button" name="ViewAccessory" value="View Accessory" tabindex="3" onClick="ViewAcc();" disabled class="btnstyle">			
		</td>
	</tr>
	<tr>
		<td nowrap>On Backorder:</td>
		<td nowrap><input type="checkbox" name="OnBackOrder" tabindex="4" class="chkstyle"></td>
	</tr>
	<tr>
		<td nowrap>List Unit Cost:</td>
		<td nowrap><input type="text" name="ListUnitCost" size="8" tabindex="5" readonly></td>
	</tr>
	<tr>
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="6" value="1" onKeypress="AllowNumericOnly();" onChange="CalculateTotal();" ></td>
	</tr>
	<tr>
		<td nowrap>Total:</td>
		<td nowrap><input type="text" name="Total" size="10" tabindex="7" value="0" readonly></td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="10" cols="65" tabindex="8" accesskey="L"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="9" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="10" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ClassID" value="">
<input type="hidden" name="ClassBundle" value="">
<input type="hidden" name="MM_Insert" value="false">
</form>
</body>
</html>