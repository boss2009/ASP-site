<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var EquipUserID = 0;
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
	var cmdInsertBuyout = Server.CreateObject("ADODB.Command");
	cmdInsertBuyout.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertBuyout.CommandText = "dbo.cp_Buyout_Request3";
	cmdInsertBuyout.CommandType = 4;
	cmdInsertBuyout.CommandTimeout = 0;
	cmdInsertBuyout.Prepared = true;
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insEq_user_type", 2, 1,1,Request.Form("BuyerType")));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@intEq_user_id", 3, 1,1,EquipUserID));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@dtsRequest_date", 135, 1,1,Request.Form("DateRequested")));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insApprvd_Staff_id", 2, 1,1,0));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@dtsApprvd_Date", 200, 1,30,"1/1/1900"));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@bitIsBack_Ordered", 2, 1,1,0));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insBuyout_Status_id", 2, 1,1,Request.Form("BuyoutStatus")));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insBuyout_Prc_id", 2, 1,1,0));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertBuyout.Parameters.Append(cmdInsertBuyout.CreateParameter("@intRtnFlag", 3, 2));
	cmdInsertBuyout.Execute();

	Response.Redirect("m010FS3.asp?intBuyout_Req_id="+cmdInsertBuyout.Parameters.Item("@intRtnFlag").Value);
}

var rsBuyoutStatus = Server.CreateObject("ADODB.Recordset");
rsBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutStatus.Source = "{call dbo.cp_buyout_status(0,'',0,'Q',0)}"
rsBuyoutStatus.CursorType = 0;
rsBuyoutStatus.CursorLocation = 2;
rsBuyoutStatus.LockType = 3;
rsBuyoutStatus.Open();

var rsUserType = Server.CreateObject("ADODB.Recordset");
rsUserType.ActiveConnection = MM_cnnASP02_STRING;
rsUserType.Source = "{call dbo.cp_eq_user_type2(0,'',1,0,2,'Q',0)}";
rsUserType.CursorType = 0;
rsUserType.CursorLocation = 2;
rsUserType.LockType = 3;
rsUserType.Open();

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
	<title>New Buyout Request</title>
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
		   	case 76 :
				//alert("L");
				window.close();
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
				document.frm0101.ClientBuyerName.value="";
				document.frm0101.ClientBuyerID.value="0";				
				
				oInstitutionBuyerLabel.style.visibility = "hidden";				
				document.frm0101.ListInstitutionBuyer.style.visibility = "hidden";
				document.frm0101.InstitutionBuyerName.style.visibility = "hidden";				
				document.frm0101.InstitutionBuyerName.value="";
				document.frm0101.InstitutionBuyerID.value="0";								
			break;
			//client
			case "3":
				oClientBuyerLabel.style.visibility = "visible";
				document.frm0101.ClientBuyerName.style.visibility = "visible";																				
				document.frm0101.ListClientBuyer.style.visibility = "visible";
				
				oInstitutionBuyerLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionBuyer.style.visibility = "hidden";
				document.frm0101.InstitutionBuyerName.style.visibility = "hidden";																
				document.frm0101.InstitutionBuyerName.value="";
				document.frm0101.InstitutionBuyerID.value="0";								
			break;
			//institution
			case "4":
				oClientBuyerLabel.style.visibility = "hidden";				
				document.frm0101.ListClientBuyer.style.visibility = "hidden";
				document.frm0101.ClientBuyerName.style.visibility = "hidden";								
				document.frm0101.ClientBuyerName.value="";
				document.frm0101.ClientBuyerID.value="0";				
				
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
	
	function Save(){
		if (!CheckDate(document.frm0101.DateRequested.value)){
			alert("Invalid Date Requested.");
			document.frm0101.DateRequested.focus();
			return ;
		}
		if (document.frm0101.BuyerType.value=="0"){
			alert("Select Buyer Type.");
			document.frm0101.BuyerType.focus();
			return ;
		}
		if ((document.frm0101.BuyerType.value=="3") && (document.frm0101.ClientBuyerID.value<="0")) {
			alert("Select a Client.");
			return ;
		}		
		if ((document.frm0101.BuyerType.value=="4") && (document.frm0101.InstitutionBuyerID.value<="0")) {
			alert("Select an Institution.");
			return ;
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
<h5>New Buyout Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Date Requested:</td>
		<td nowrap width="200">
			<input type="text" name="DateRequested" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Buyer Type:</td>
		<td nowrap><select name="BuyerType" tabindex="7" onChange="ChangeBuyerType();" style="width: 160px" accesskey="L">
			<%
			while (!rsUserType.EOF) {
			%>
				<option value="<%=(rsUserType.Fields.Item("insEq_user_type").Value)%>" <%=((rsUserType.Fields.Item("insEq_user_type").Value==3)?"SELECTED":"")%>><%=(rsUserType.Fields.Item("chvEq_user_type").Value)%></option>
			<%
				rsUserType.MoveNext();
			}
			%>		
		</select></td>		
	</tr>
	<tr> 
		<td nowrap>Buyout Status:</td>
		<td nowrap><select name="BuyoutStatus" tabindex="2" style="width: 160px">
		<%
		while (!rsBuyoutStatus.EOF) {
		%>
			<option value="<%=(rsBuyoutStatus.Fields.Item("insbuyout_status_id").Value)%>" <%=((rsBuyoutStatus.Fields.Item("insbuyout_status_id").Value=="1")?"SELECTED":"")%>><%=(rsBuyoutStatus.Fields.Item("chvBuyout_status").Value)%> 
		<%
			rsBuyoutStatus.MoveNext();
		}
		%>
		</select></td>
		<td nowrap><div id="oClientBuyerLabel">Client Buyer:</div></td>
		<td nowrap> 
			<input type="text" name="ClientBuyerName" readonly tabindex="8">
			<input type="button" name="ListClientBuyer" value="List" onClick="openWindow('m010p0202.asp','wPopUser');" tabindex="9" class="btnstyle">
		</td>
    </tr>
    <tr> 
		<td></td>
		<td></td>
		<td nowrap><div id="oInstitutionBuyerLabel">Institution Buyer:</div></td>
		<td nowrap> 
			<input type="text" name="InstitutionBuyerName" readonly tabindex="10">
			<input type="button" name="ListInstitutionBuyer" value="List" onClick="openWindow('m010p0201.asp','wPopUser')" tabindex="11" class="btnstyle">
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="12" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Cancel" tabindex="13" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ClientBuyerID" value="0">
<input type="hidden" name="InstitutionBuyerID" value="0">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsBuyoutStatus.Close();
rsStaff.Close();
%>