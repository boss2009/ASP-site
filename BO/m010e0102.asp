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
	var rsBuyoutCIP = Server.CreateObject("ADODB.Recordset");
	rsBuyoutCIP.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutCIP.Source = "{call dbo.cp_buyout_cip("+Request.Form("MM_recordId")+",'',"+Request.Form("MM_noteId")+",0,'E',0)}";
	rsBuyoutCIP.CursorType = 0;
	rsBuyoutCIP.CursorLocation = 2;
	rsBuyoutCIP.LockType = 3;
	rsBuyoutCIP.Open();	
	Response.Redirect("m010e0101.asp?intBuyout_Req_id="+Request.Form("MM_recordId"));
}

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var IsClient = true;
if (!rsBuyout.EOF) {
	if ((rsBuyout.Fields.Item("insEq_user_type").Value == 3) && (rsBuyout.Fields.Item("intEq_user_id").Value > 0)) {
		var rsBuyoutCIP = Server.CreateObject("ADODB.Recordset");
		rsBuyoutCIP.ActiveConnection = MM_cnnASP02_STRING;
		rsBuyoutCIP.Source = "{call dbo.cp_buyout_cip("+Request.QueryString("intBuyout_Req_id")+","+rsBuyout.Fields.Item("intEq_user_id").Value+",0,0,'Q',0)}";
		rsBuyoutCIP.CursorType = 0;
		rsBuyoutCIP.CursorLocation = 2;
		rsBuyoutCIP.LockType = 3;
		rsBuyoutCIP.Open();	
	} else {
		IsClient = false;
	}
} else {
	IsClient = false;
}
%>									
<html>
<head>
	<title>Update TAP Date</title>
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
				document.frm0102.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">	
	function Link(){
		if (document.frm0102.ServiceDate.length>1){
			for (var i=0; i < document.frm0102.ServiceDate.length; i++) {
				if (document.frm0102.ServiceDate[i].checked) document.frm0102.MM_noteId.value=document.frm0102.ServiceDate[i].value;	
			}
		} else {
			document.frm0102.MM_noteId.value=document.frm0102.ServiceDate.value;	
		}
		if (document.frm0102.MM_noteId.value > 0){	
			document.frm0102.submit();
		} else {
			alert("Select a service date.");
		}
	}
	</script>
</head>
<body>
<form action="<%=MM_editAction%>" method="POST" name="frm0102">
<h5>Update TAP Date</h5>
<%
if (!IsClient) {
%>
Information not available.  Either the buyer is an institution or the client is not found.
<%
} else {
%>
<i>Select one of the following service date as the TAP date for this buyout.</i>
<hr>
<table cellpadding="1" cellspacing=1" style="border: 1px solid">
	<tr>
		<th class="headrow" nowrap>&nbsp;</th>
		<th class="headrow" nowrap>Request Date</th>
		<th class="headrow" nowrap>Year</th>
		<th class="headrow" nowrap>Service Provider</th>
		<th class="headrow" nowrap>Service Code</th>
<%
if (String(Request.QueryString("ShowFS"))=="1") {
%>
		<th class="headrow" nowrap>Funding Source</th>
<%
}
%>
	</tr>
<%
var count = 0
while (!rsBuyoutCIP.EOF) {
%>
	<tr class=<%=(((count%2)=="0")?"rowa":"rowb")%>>
		<td style="background-color: #FFFFFF"><input type="radio" name="ServiceDate" value="<%=rsBuyoutCIP.Fields.Item("intSrv_Note_id").Value%>" <%=((rsBuyoutCIP.Fields.Item("bitIs_BO_CIP").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
		<td align="center"><%=FilterDate(rsBuyoutCIP.Fields.Item("dtsRequest_Date").Value)%></td>
		<td align="center"><%=rsBuyoutCIP.Fields.Item("chvYear").Value%></td>
		<td align="left"><%=rsBuyoutCIP.Fields.Item("chvService_provider").Value%></td>
		<td align="center"><%=rsBuyoutCIP.Fields.Item("chvSrv_Code").Value%></td>
<%
if (String(Request.QueryString("ShowFS"))=="1") {
%>
		<td align="center"><%=rsBuyoutCIP.Fields.Item("chvFunding_Source").Value%></td>
<%
}
%>

	</tr>
<%
	rsBuyoutCIP.MoveNext();
	count++;
}
%>		
</table>
<hr>
<%
if (count > 0){
%>
<input type="button" value="Link" onClick="Link();" class="btnstyle">
<input type="button" value="Cancel" onClick="window.location.href='m010e0101.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>';" class="btnstyle">
<%
if (String(Request.QueryString("ShowFS"))=="1") {
%>
<input type="button" value="Hide Funding Source" onClick="window.location.href='m010e0102.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>&ShowFS=0';" class="btnstyle">
<%
} else {
%>
<input type="button" value="Show Funding Source" onClick="window.location.href='m010e0102.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>&ShowFS=1';" class="btnstyle">
<%
}
}
}
%>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsBuyout.Fields.Item("intBuyout_Req_id").Value %>">
<input type="hidden" name="MM_noteId" value="">
</form>
</body>
</html>