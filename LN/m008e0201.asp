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
	var IsBackOrdered = ((Request.Form("IsBackOrdered")=="on") ? "1":"0");
	var Comments = String(Request.Form("Comments")).replace(/'/g, "''");
	var ClassID = ((String(Request.Form("ClassID"))!="undefined")?Request.Form("ClassID"):0);	
	var Quantity = ((String(Request.Form("Quantity"))!="undefined")?Request.Form("Quantity"):0);	
	var ListUnitCost = ((String(Request.Form("ListUnitCost"))!="undefined")?Request.Form("ListUnitCost"):0);			
	var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
	rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryRequest.Source = "{call dbo.cp_eqp_requested("+Request.QueryString("intEqpReq_Id")+","+Request.QueryString("intLoan_req_id")+","+ClassID+","+Request.Form("ClassBundle")+","+Quantity+",'"+Comments+"',"+ListUnitCost+","+IsBackOrdered+",0,'E',0)}";
	rsInventoryRequest.CursorType = 0;
	rsInventoryRequest.CursorLocation = 2;
	rsInventoryRequest.LockType = 3;
	rsInventoryRequest.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m008q0201.asp&intLoan_req_id="+Request.QueryString("intLoan_req_id"));
}

var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequest.Source = "{call dbo.cp_eqp_requested("+Request.QueryString("intEqpReq_Id")+",0,0,0,0,'',0.0,0,1,'Q',0)}";
rsInventoryRequest.CursorType = 0;
rsInventoryRequest.CursorLocation = 2;
rsInventoryRequest.LockType = 3;
rsInventoryRequest.Open();
%>
<html>
<head>
	<title>Update Equipment Request</title>
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
				window.location.href='m008q0201.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>';
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
			alert("Select a class.");
			document.frm0201.ListClass.focus();
			return ;
		}
		CalculateTotal();		
		document.frm0201.MM_update.value="true";
		document.frm0201.submit();
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
<h5>Equipment Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td nowrap>Class/Bundle:</td>
		<td nowrap> 
			<input type="text" name="ClassName" value="<%=((rsInventoryRequest.Fields.Item("bitIs_class").Value==1)?rsInventoryRequest.Fields.Item("chvEqp_Class_Name").Value:rsInventoryRequest.Fields.Item("chvEqp_Bundle_Name").Value)%>" tabindex="1" size="40" accesskey="F" readonly>
			<input type="button" name="ListClass" value="List" tabindex="2" onClick="openWindow('m008p01FS.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>','');" class="btnstyle">
		</td>
	</tr>
    <tr> 
		<td nowrap>List Unit Cost:</td>
		<td nowrap><input type="text" name="ListUnitCost" size="10" tabindex="3" value="<%=rsInventoryRequest.Fields.Item("fltList_unit_cost").Value%>" readonly></td>
    </tr>
    <tr> 
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="4" onKeypress="AllowNumericOnly();" onChange="CalculateTotal();" value="<%=((String(Request.QueryString("ListUnitCost"))=="undefined")?rsInventoryRequest.Fields.Item("insQuantity").Value:"0")%>"></td>
    </tr>
    <tr> 
		<td nowrap>Total:</td>
		<td nowrap><input type="text" name="Total" size="10" tabindex="5" readonly></td>
    </tr>
	<tr>
		<td nowrap>Is Backordered:</td>
		<td nowrap><input type="checkbox" name="IsBackOrdered" tabindex="6" <%=(((rsInventoryRequest.Fields.Item("bitIs_BO").Value)=="1")?"CHECKED":"")%> class="chkstyle"></td>
	</tr>
    <tr> 
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="5" cols="65" tabindex="7"><%=(rsInventoryRequest.Fields.Item("chvComments").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="9" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m008q0201.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>'" tabindex="10" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="false">
<input type="hidden" name="ClassID" value="<%=rsInventoryRequest.Fields.Item("insClass_bundle_id").Value%>">
<input type="hidden" name="ClassBundle" value="<%=((rsInventoryRequest.Fields.Item("bitIs_Class").Value==1)?"1":"0")%>">					
</form>
</body>
</html>
<%
rsInventoryRequest.Close();
%>