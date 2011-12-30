<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var EquipUserID = 0;
	var IsBackOrdered = ((Request.Form("EquipmentOnBackOrder")=="on")?"1":"0");

	switch (String(Request.Form("BuyerType"))) {
		//client
		case "3":
			EquipUserID = Request.Form("ClientBuyerID");
		break;
		//institution
		case "4":
			EquipUserID = Request.Form("InstitutionBuyerID");
		break;
		//none
		default:
			EquipUserID = 0;
		break;
	}

	var rsBuyout = Server.CreateObject("ADODB.Recordset");
	rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyout.Source = "{call dbo.cp_Buyout_request3("+Request.Form("MM_recordId")+","+Request.Form("BuyerType")+","+EquipUserID+",'"+Request.Form("DateRequested")+"',"+Request.Form("ApprovedBy")+",'"+Request.Form("DateApproved")+"',"+IsBackOrdered+","+Request.Form("BuyoutStatus")+","+Request.Form("BuyoutProcess")+","+Session("insStaff_id")+",0,'E',0)}";
	rsBuyout.CursorType = 0;
	rsBuyout.CursorLocation = 2;
	rsBuyout.LockType = 3;
	rsBuyout.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m010e0101.asp&intBuyout_Req_id="+Request.Form("MM_recordId"));	
}

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var rsBuyoutProcess = Server.CreateObject("ADODB.Recordset");
rsBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutProcess.Source = "{call dbo.cp_buyout_process(0,'',1,0,'Q',0)}";
rsBuyoutProcess.CursorType = 0;
rsBuyoutProcess.CursorLocation = 2;
rsBuyoutProcess.LockType = 3;
rsBuyoutProcess.Open();

var rsBuyoutStatus = Server.CreateObject("ADODB.Recordset");
rsBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutStatus.Source = "{call dbo.cp_buyout_status(0,'',0,'Q',0)}"
rsBuyoutStatus.CursorType = 0;
rsBuyoutStatus.CursorLocation = 2;
rsBuyoutStatus.LockType = 3;
rsBuyoutStatus.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsUserType = Server.CreateObject("ADODB.Recordset");
rsUserType.ActiveConnection = MM_cnnASP02_STRING;
rsUserType.Source = "{call dbo.cp_eq_user_type2(0,'',1,0,2,'Q',0)}";
rsUserType.CursorType = 0;
rsUserType.CursorLocation = 2;
rsUserType.LockType = 3;
rsUserType.Open();

var InstUserName = "";
var InstUserId = 0;
var IdvUserName = "";
var IdvUserId = 0;
switch (String(rsBuyout.Fields.Item("insEq_user_type").Value)) {
	//client
	case "3":		
		var rsIndClient = Server.CreateObject("ADODB.Recordset");
		rsIndClient.ActiveConnection = MM_cnnASP02_STRING;
		rsIndClient.Source = "{call dbo.cp_Idv_Adult_Client("+rsBuyout.Fields.Item("intEq_user_id").Value+")}";
		rsIndClient.CursorType = 0;
		rsIndClient.CursorLocation = 2;
		rsIndClient.LockType = 3;
		rsIndClient.Open();
		IdvUserName = rsIndClient.Fields.Item("chvLst_Name").Value + ", " + rsIndClient.Fields.Item("chvFst_Name").Value;		
		IdvUserId = rsBuyout.Fields.Item("intEq_user_id").Value;
		rsIndClient.Close();
	break;
	//institution
	case "4":
		var rsIndInstitution = Server.CreateObject("ADODB.Recordset");
		rsIndInstitution.ActiveConnection = MM_cnnASP02_STRING;		
		rsIndInstitution.Source = "{call dbo.cp_school3("+rsBuyout.Fields.Item("intEq_user_id").Value+",'',0,0,0,0,0,0,0,'',1,'Q',0)}";
		rsIndInstitution.CursorType = 0;
		rsIndInstitution.CursorLocation = 2;
		rsIndInstitution.LockType = 3;
		rsIndInstitution.Open();		
		InstUserName = rsIndInstitution.Fields.Item("chvSchool_Name").Value;
		InstUserId = rsBuyout.Fields.Item("intEq_user_id").Value;		
		rsIndInstitution.Close();		
	break;
	default:
		IdvUserName = "";
		InstUserName = "";				
	break;
}
%>									
<html>
<head>
	<title>General Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
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
		}
	}
	</script>	
	<script language="Javascript">	
	function ChangeBuyerType(){
		switch ((document.frm0101.BuyerType.value)){
			//none
			case "0":
				oClientBuyerLabel.style.visibility = "hidden";
				document.frm0101.ClientBuyerName.style.visibility = "hidden";				
				document.frm0101.ListClientBuyer.style.visibility = "hidden";
				
				oInstitutionBuyerLabel.style.visibility = "hidden";				
				document.frm0101.ListInstitutionBuyer.style.visibility = "hidden";
				document.frm0101.InstitutionBuyerName.style.visibility = "hidden";				
			break;
			//client
			case "3":
				oClientBuyerLabel.style.visibility = "visible";
				document.frm0101.ClientBuyerName.style.visibility = "visible";																				
				document.frm0101.ListClientBuyer.style.visibility = "visible";
				
				oInstitutionBuyerLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionBuyer.style.visibility = "hidden";
				document.frm0101.InstitutionBuyerName.style.visibility = "hidden";																
			break;
			//institution
			case "4":
				oClientBuyerLabel.style.visibility = "hidden";				
				document.frm0101.ListClientBuyer.style.visibility = "hidden";
				document.frm0101.ClientBuyerName.style.visibility = "hidden";								
				
				oInstitutionBuyerLabel.style.visibility = "visible";								
				document.frm0101.ListInstitutionBuyer.style.visibility = "visible";
				document.frm0101.InstitutionBuyerName.style.visibility = "visible";
			break;
		}
	}
	
	function Init(){
		ChangeBuyerType();
		document.frm0101.DateRequested.focus();
	}
	
	function ChangeStatus(){
		if (document.frm0101.BuyoutStatus.value=="2") {
			if (Trim(document.frm0101.DateApproved.value)=="") {
				document.frm0101.DateApproved.value = "<%=CurrentDate()%>";
			}
			if (document.frm0101.ApprovedBy.value <= 0) {
				document.frm0101.ApprovedBy.value = "<%=Session("insStaff_id")%>";
			}
		}
	}
		
	function Save(){
		if (!CheckDate(document.frm0101.DateRequested.value)){
			alert("Invalid Date Requested.");
			document.frm0101.DateRequested.focus();
			return ;
		}
		if (document.frm0101.BuyerType.value == "3") {
			if (document.frm0101.ClientBuyerID.value == "0") {
				alert("Select a client.");
				return ;
			}
		}
		if (document.frm0101.BuyerType.value == "4") {
			if (document.frm0101.InstitutionBuyerID.value == "0") {
				alert("Select an institution");
				return ;
			}
		}
		document.frm0101.submit();
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=460,height=430,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td>Date Requested:</td>
		<td>
			<input type="text" name="DateRequested" value="<%=FilterDate(rsBuyout.Fields.Item("dtsRequest_date").Value)%>" size="11" maxlength="10" readonly tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
		<td>Buyout Status:</td>
		<td><select name="BuyoutStatus" tabindex="2" onChange="ChangeStatus();" style="width: 160px">
			<%
			while (!rsBuyoutStatus.EOF) {
			%>
				<option value="<%=(rsBuyoutStatus.Fields.Item("insbuyout_status_id").Value)%>" <%=((rsBuyoutStatus.Fields.Item("insbuyout_status_id").Value==rsBuyout.Fields.Item("insBuyout_Status_id").Value)?"SELECTED":"")%>><%=(rsBuyoutStatus.Fields.Item("chvBuyout_status").Value)%>
			<%
				rsBuyoutStatus.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr> 
		<td>Date Approved:</td>
		<td>
			<input type="text" name="DateApproved" value="<%=FilterDate(rsBuyout.Fields.Item("dtsApprvd_Date").Value)%>" size="11" maxlength="10" readonly tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td>Approved By:</td>
		<td><select name="ApprovedBy" tabindex="4">
				<option value="0">(none)
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsBuyout.Fields.Item("insApprvd_Staff_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
			<%
				rsStaff.MoveNext();
			}
			%>		
		</select></td>		
	</tr>
    <tr> 
		<td>Buyout Process:</td>
		<td><select name="BuyoutProcess" tabindex="5" style="width: 160px">
			<%
			while (!rsBuyoutProcess.EOF) {
			%>
				<option value="<%=(rsBuyoutProcess.Fields.Item("insBuyout_process_id").Value)%>" <%=((rsBuyoutProcess.Fields.Item("insBuyout_process_id").Value==rsBuyout.Fields.Item("insBuyout_Prc_id").Value)?"SELECTED":"")%>><%=(rsBuyoutProcess.Fields.Item("chvBuyout_process").Value)%>
			<%
				rsBuyoutProcess.MoveNext();
			}
			%>
		</select></td>
		<td>Buyer Type:</td>		
		<td><select name="BuyerType" tabindex="6" onChange="ChangeBuyerType();">
			<%
			while (!rsUserType.EOF) {
			%>
				<option value="<%=(rsUserType.Fields.Item("insEq_user_type").Value)%>" <%=((rsUserType.Fields.Item("insEq_user_type").Value==rsBuyout.Fields.Item("insEq_user_type").Value)?"SELECTED":"")%>><%=(rsUserType.Fields.Item("chvEq_user_type").Value)%></option>
			<%
				rsUserType.MoveNext();
			}
			%>		
		</select></td>				
	</tr>	
	<tr>
		<td colspan="2"><input type="checkbox" name="EquipmentOnBackOrder" <%=((rsBuyout.Fields.Item("bitIsBack_Ordered").Value=="1")?"CHECKED":"")%> tabindex="7" readonly class="chkstyle">Equipment on Backorder</td>				
		<td nowrap><div id="oClientBuyerLabel">Client Buyer:</div></td>
		<td nowrap>
			<input type="text" name="ClientBuyerName" value="<%=IdvUserName%>" readonly tabindex="8">
			<input type="button" name="ListClientBuyer" value="List" onClick="openWindow('m010p0202.asp','wPopUser');" tabindex="9" class="btnstyle">
		</td>
    </tr>
    <tr> 
		<td colspan="2"></td>
		<td nowrap><div id="oInstitutionBuyerLabel">Institution Buyer:</div></td>
		<td nowrap>
			<input type="text" name="InstitutionBuyerName" value="<%=InstUserName%>" readonly tabindex="10">
			<input type="button" name="ListInstitutionBuyer" value="List" onClick="openWindow('m010p0201.asp','wPopUser');" tabindex="11" accesskey="L" class="btnstyle">
		</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="12" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="13" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ClientBuyerID" value="<%=IdvUserId%>">
<input type="hidden" name="InstitutionBuyerID" value="<%=InstUserId%>">
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsBuyout.Fields.Item("intBuyout_Req_id").Value %>">
</form>
</body>
</html>
<%
rsBuyoutProcess.Close();
rsBuyoutStatus.Close();
rsBuyout.Close();
rsStaff.Close();
%>