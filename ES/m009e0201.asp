<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true"){
	var IsReceived = ((Request.Form("IsReceived")=="1")?"1":"0");	
	var DateReceived = ((String(Request.Form("DateReceived"))=="undefined")?"1/1/1900":Request.Form("DateReceived"));		
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
	rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceRequested.Source = "{call dbo.cp_eqpSrv_Request("+ Request.Form("MM_recordId") + ",'"+Request.Form("DateRequested")+"',"+IsReceived+",'"+DateReceived+"','"+Description+"',1,'E',0)}";
	rsServiceRequested.CursorType = 0;
	rsServiceRequested.CursorLocation = 2;
	rsServiceRequested.LockType = 3;
	rsServiceRequested.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m009e0201.asp&intEquip_srv_id="+Request.QueryString("intEquip_srv_id"));
}

var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
rsServiceRequested.Source = "{call dbo.cp_eqpSrv_Request("+ Request.QueryString("intEquip_srv_id") + ",'',0,'','',1,'Q',0)}"
rsServiceRequested.CursorType = 0;
rsServiceRequested.CursorLocation = 2;
rsServiceRequested.LockType = 3;
rsServiceRequested.Open();
%>
<html>
<head>
	<title>Service Requested</title>
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
				document.frm0201.reset();
			break;
			case 76 :
				//alert("L");
				window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>';
			break;
		}
	}
	</script>		
	<script language="Javascript">
	function ChangeReceived() {
		if (document.frm0201.IsReceived.checked) {
			document.frm0201.DateReceived.disabled = false;
			if (Trim(document.frm0201.DateReceived.value)=="") {			
				document.frm0201.DateReceived.value = "<%=CurrentDate()%>";
			}
		} else {
			document.frm0201.DateReceived.disabled = true;
			document.frm0201.DateReceived.value = "";		
		}		
	}
	
	function Init(){
		ChangeReceived();
		document.frm0201.DateRequested.focus();
	}
	
	function Save(){
		if (document.frm0201.IsReceived.checked) {
			if (!CheckDate(document.frm0201.DateReceived.value)) {
				alert("Invalid Date Received.");
				document.frm0201.DateReceived.focus();
				return ;
			}
		}		
		if (!CheckTextArea(document.frm0201.Description, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
		
		document.frm0201.submit();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0201" method="POST" action="<%=MM_editAction%>">
<h5>Service Requested</h5>
<hr>
<table cellpadding="2" cellspacing="3">
	<tr>		
		<td nowrap>Date Requested:</td>
		<td nowrap>
			<input type="text" name="DateRequested" value="<%=FilterDate(rsServiceRequested.Fields.Item("dtsRequested_date").Value)%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>		
		<td nowrap><input type="checkbox" name="IsReceived" value="1" tabindex="2" <%=((rsServiceRequested.Fields.Item("bitIs_received").Value=="1")?"CHECKED":"")%> onClick="ChangeReceived();" class="chkstyle">Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=FilterDate(rsServiceRequested.Fields.Item("dtsReceived_Date").Value)%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>	
	<tr>
		<td valign="top">Description:</td>
		<td valign="top"><textarea name="Description" cols="65" rows="10" tabindex="4" accesskey="L"><%=rsServiceRequested.Fields.Item("chvNote_Desc").Value%></textarea>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>'" class="btnstyle"></td>		
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("intEquip_Srv_id")%>">
</form>
</body>
</html>
<%
rsServiceRequested.Close();
%>