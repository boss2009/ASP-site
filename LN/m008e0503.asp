<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_get_loan_ship_name("+ Request.QueryString("intLoan_Req_id") + ",0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var intShip_dtl_id = 0;
if ((!rsLoan.EOF) && (rsLoan.Fields.Item("intShip_dtl_id").Value != null)) intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value;

if (String(Request("MM_action")) == "update") {
	var rsUpdateSchedule = Server.CreateObject("ADODB.Recordset");
	rsUpdateSchedule.ActiveConnection = MM_cnnASP02_STRING;
	rsUpdateSchedule.Source = "{call dbo.cp_loan_ship_schedule("+Request.Form("MM_recordId")+","+Request.Form("DeliveryOnSchedule")+","+Request.Form("DeliveryStatus")+","+Request.Form("Shipper")+",'"+Trim(Request.Form("ShipperPhoneAreaCode"))+"','"+Trim(Request.Form("ShipperPhoneNumber"))+"','"+Trim(Request.Form("ShipperPhoneExtension"))+"','E',0)}";
	rsUpdateSchedule.CursorType = 0;
	rsUpdateSchedule.CursorLocation = 2;
	rsUpdateSchedule.LockType = 3;	
	rsUpdateSchedule.Open();
}

var rsShippingSchedule = Server.CreateObject("ADODB.Recordset");
rsShippingSchedule.ActiveConnection = MM_cnnASP02_STRING;
rsShippingSchedule.Source = "{call dbo.cp_loan_ship_schedule("+intShip_dtl_id+",0,0,0,'','','','Q',0)}";
rsShippingSchedule.CursorType = 0;
rsShippingSchedule.CursorLocation = 2;
rsShippingSchedule.LockType = 3;
rsShippingSchedule.Open();

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
	<title>Shipping Schedule</title>
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
				window.location.href='m008e0501.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
	<%
	if (!rsShippingSchedule.EOF) {
	%>
		if (document.frm0503.DeliveryOnSchedule.value=="0") {
			document.frm0503.DeliveryStatus.style.visibility = "visible";
		} else {
			document.frm0503.DeliveryStatus.style.visibility = "hidden";		
		}
		document.frm0503.ScheduledArrivalDate.focus();
	<%
	}
	%>
	}

	function ChangeDeliveryStatus(){
		if (document.frm0503.DeliveryOnSchedule.value=="0") {
			document.frm0503.DeliveryStatus.style.visibility = "visible";
			document.frm0503.DeliveryStatus.value="1";
		} else {
			document.frm0503.DeliveryStatus.style.visibility = "hidden";		
		}
	}
		
	function Save(){
		document.frm0503.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0503">
<h5>Shipping Schedule</h5>
<hr>
<%
if (rsShippingSchedule.EOF) {
%>
<i>Please go to Method page and save first, before entering shipping schedule.</i>
<%
} else {
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Scheduled Arrival Date:</td>
		<td nowrap>
			<input type="text" name="ScheduledArrivalDate" value="<%=FilterDate(rsShippingSchedule.Fields.Item("dtsSch_Arv_date").Value)%>" size="11" readonly maxlength="10" tabindex="1" style="border: none" accesskey="F" onChange="FormatDate(this)">
<!--		<span style="font-size: 7pt">(mm/dd/yyyy)</span>-->
		</td>
	</tr>
	<tr>
		<td nowrap>Delivery on Schedule:</td>
		<td nowrap>
			<select name="DeliveryOnSchedule" tabindex="2" onChange="ChangeDeliveryStatus();">
				<option value="1" <%=((rsShippingSchedule.Fields.Item("bitIsDlvy_onshdl").Value == "1")?"SELECTED":"")%>>Yes
				<option value="0" <%=((rsShippingSchedule.Fields.Item("bitIsDlvy_onshdl").Value != "1")?"SELECTED":"")%>>No
			</select>
			<select name="DeliveryStatus" tabindex="3">
				<option value="1" <%=((rsShippingSchedule.Fields.Item("bitIsDlvy_delay").Value == "1")?"SELECTED":"")%>>Delay
				<option value="0" <%=((rsShippingSchedule.Fields.Item("bitIsDlvy_delay").Value != "1")?"SELECTED":"")%>>Delivery Resolved
			</select>
		</td>
	</tr>
	<tr> 
		<td nowrap>Shipper:</td>
		<td nowrap><select name="Shipper" tabindex="4" accesskey="L">
		<% 
		var staffid = Session("insStaff_id");
		if (rsShippingSchedule.Fields.Item("insMail_Staff_id").Value != null) staffid = rsShippingSchedule.Fields.Item("insMail_Staff_id").Value;
		while (!rsStaff.EOF) {
		%>
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==staffid)?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvName").Value)%></option>
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
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="window.location.href='m008e0501.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ShipperPhoneAreaCode" value="<%=rsShippingSchedule.Fields.Item("chvMSPhone_Arcd").value%>">
<input type="hidden" name="ShipperPhoneNumber" value="<%=(rsShippingSchedule.Fields.Item("chvMSPhone_Num").Value)%>">
<input type="hidden" name="ShipperPhoneExtension" value="<%=(rsShippingSchedule.Fields.Item("chvMSPhone_Ext").Value)%>">
<%
}
%>
<input type="hidden" name="MM_action" value="update">
<input type="hidden" name="MM_recordId" value="<%=intShip_dtl_id%>">
</form>
</body>
</html>
<%
rsLoan.Close();
rsShippingSchedule.Close();
rsStaff.Close();
%>