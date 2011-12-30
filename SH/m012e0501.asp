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
	var Year = ((Request.Form("Year")!="")?Request.Form("Year"):"0");
	var Cycle= ((Request.Form("Cycle")!="")?Request.Form("Cycle"):"0");	
	var ServiceNotes = String(Request.Form("ServiceNotes")).replace(/'/g, "''");
	var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
	rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+Request.QueryString("insSchool_id")+","+Request.QueryString("intSrv_Note_id")+",'"+Request.Form("DateRequested")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+ServiceNotes+"','"+Request.Form("ServiceRequestHexCode")+"',0,'E',0)}";
	rsServiceRequested.CursorType = 0;
	rsServiceRequested.CursorLocation = 2;
	rsServiceRequested.LockType = 3;
	rsServiceRequested.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012q0501.asp&insSchool_id="+Request.QueryString("insSchool_id"));
}	

var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+Request.QueryString("insSchool_id")+","+Request.QueryString("intSrv_Note_id")+",'',0,0,0,'','',1,'Q',0)}";
rsServiceRequested.CursorType = 0;
rsServiceRequested.CursorLocation = 2;
rsServiceRequested.LockType = 3;
rsServiceRequested.Open();
var count = 0;
while (!rsServiceRequested.EOF) {
	count ++;
	rsServiceRequested.MoveNext();
}
rsServiceRequested.MoveFirst();

var rsServiceRequestedType = Server.CreateObject("ADODB.Recordset");
rsServiceRequestedType.ActiveConnection = MM_cnnASP02_STRING;
rsServiceRequestedType.Source = "{call dbo.cp_service_type(0,0,1,2)}";
rsServiceRequestedType.CursorType = 0;
rsServiceRequestedType.CursorLocation = 2;
rsServiceRequestedType.LockType = 3;
rsServiceRequestedType.Open();

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source_attributes(0,0,0,0,0,0,0,1,2,'Q',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();

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
	<title>Update Service Requested</title>
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
			case 85 :
				//alert("U");
				frm0501.reset;
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}	
	</script>	
	<script language="Javascript">	
	var count = <%=count%>;
	
	function Save(){
		if (!CheckDate(document.frm0501.DateRequested.value)){
			alert("Invalid Date Requested.");
			document.frm0501.DateRequested.focus();
			return;
		}
		
		if (count < 1){
			alert("Select Service Type.");
			document.frm0501.ServiceType[0].focus();
			return ;
		}
		
		var temp = "";
		for (var i=0; i < count; i++){
			temp = temp + PadDecToHex(document.frm0501.ServiceType[i].value) + PadDecToHex(document.frm0501.FundingSource[i].value);
		}
		var zero = 40 - temp.length;
		for (var j = 0; j < zero; j++){
			temp = temp + String("0");
		}
		document.frm0501.ServiceRequestHexCode.value=temp;
		document.frm0501.submit();
	}
	
	function AddService(){
		document.frm0501.ServiceType[count].style.visibility = "visible";
		document.frm0501.FundingSource[count].style.visibility = "visible";
		count++;
		document.frm0501.Remove.disabled = false;			
		if (count == 10) document.frm0501.Add.disabled = true;
	}
	
	function RemoveService(){
		count--;	
		document.frm0501.ServiceType[count].style.visibility = "hidden";
		document.frm0501.FundingSource[count].style.visibility = "hidden";
		document.frm0501.Add.disabled = false;
		if (count == 1) document.frm0501.Remove.disabled = true;
	}
	
	function Init(){
	<%
	for (var i = 0; i< count; i++){
	%>
		document.frm0501.ServiceType[<%=i%>].style.visibility = "visible";
		document.frm0501.FundingSource[<%=i%>].style.visibility = "visible";
	<%
	}
	if (count == 1) {
	%>
		document.frm0501.Remove.disabled = true;
	<%
	}
	%>
		document.frm0501.DateRequested.focus();
	}
	</script>
</head>
<body onLoad="Init();" >
<form action="<%=MM_editAction%>" method="POST" name="frm0501">
<h5>Service Requested</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Date Requested:</td>
		<td nowrap><input type="textbox" name="DateRequested" value="<%=rsServiceRequested.Fields.Item("dtsRequest_Date").Value%>" accesskey="F" tabindex="1" size="11" maxlength="10" onChange="FormatDate(this)"></td>
	</tr>
    <tr> 
		<td nowrap>Service Provider:</td>
		<td nowrap><select name="ServiceProvider" tabindex="2">
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsServiceRequested.Fields.Item("insSrv_Staff_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%>
		<%
			rsStaff.MoveNext();
		}
		rsStaff.MoveFirst();
		%>				
		</select></td>
	</tr>	
	<tr>
		<td nowrap>Year/Cycle:</td>
		<td nowrap>
			<input type="textbox" name="Year" value="<%=rsServiceRequested.Fields.Item("chvYear").Value%>" size="4" maxlength="4" tabindex="3" onKeypress="AllowNumericOnly();">
			<input type="textbox" name="Cycle" value="<%=rsServiceRequested.Fields.Item("chvCycle").Value%>" size="2" maxlength="2" tabindex="4" onKeypress="AllowNumericOnly();">
		</td>
	</tr>
    <tr>
		<td nowrap valign="top">Service Notes:</td>
		<td nowrap valign="top"><textarea name="ServiceNotes" rows="6" cols="57" tabindex="5" accesskey="L"><%=rsServiceRequested.Fields.Item("chvNotes").Value%></textarea></td>
    </tr>
	<tr>
		<td valign="top">Service Received:</td>
		<td valign="top"><div class="BrowsePanel" style="width: 362px; height: 139px;"> 
			<table cellpadding="1" cellspacing="1">
				<tr> 
					<td>Service Type</td>
					<td>Funding Source</td>
				</tr>
				<%
				rsServiceRequested.MoveFirst();
				for (var i=5; i< 25; i=i+2){
				%>
				<tr> 
					<td><select name="ServiceType" tabindex="<%=i%>" style="width: 180px; visibility='hidden'">
						<%
						while (!rsServiceRequestedType.EOF) {
						%>
							<option value="<%=rsServiceRequestedType.Fields.Item("insService_type_id").Value%>" <%if (!rsServiceRequested.EOF) Response.Write((rsServiceRequested.Fields.Item("insSrv_Code_id").Value==rsServiceRequestedType.Fields.Item("insService_type_id").Value)?"SELECTED":"")%>><%=rsServiceRequestedType.Fields.Item("chvname").Value%> 
						<%
							rsServiceRequestedType.MoveNext();
						}
						rsServiceRequestedType.MoveFirst();
						%>
					</select></td>
					<td><select name="FundingSource" tabindex="<%=i+1%>" style="width: 150px; visibility='hidden'">
						<%
						while (!rsFundingSource.EOF) {
						%>
							<option value="<%=rsFundingSource.Fields.Item("insFunding_source_id").Value%>" <%if (!rsServiceRequested.EOF) Response.Write((rsFundingSource.Fields.Item("insFunding_source_id").Value==rsServiceRequested.Fields.Item("insFunded_by_id").Value)?"SELECTED":"")%>><%=rsFundingSource.Fields.Item("chvfunding_source_name").Value%> 
						<%
							rsFundingSource.MoveNext();
						}
						rsFundingSource.MoveFirst();
						%>
					</select></td>
				</tr>
				<%
					if (!rsServiceRequested.EOF) rsServiceRequested.MoveNext();
				}
				%>
			</table>
        </div></td>
	</tr>
</table>
<div style="position: absolute; top: 340px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" name="Add" value="Add Service" onClick="AddService();" tabindex="26" class="btnstyle">
<input type="button" name="Remove" value="Remove Service" onClick="RemoveService();" tabindex="27" class="btnstyle">
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="28" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="29" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="history.go('-1');" tabindex="30" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="ServiceRequestHexCode" value="">
</form>
</body>
</html>
<%
rsServiceRequested.Close();
rsServiceRequestedType.Close();
rsFundingSource.Close();
%>