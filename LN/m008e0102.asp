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
	var rsLoanCIP = Server.CreateObject("ADODB.Recordset");
	rsLoanCIP.ActiveConnection = MM_cnnASP02_STRING;
	rsLoanCIP.Source = "{call dbo.cp_loan_cip2("+Request.Form("MM_recordId")+","+Request.Form("MM_noteId")+",0,'E',0)}";
	rsLoanCIP.CursorType = 0;
	rsLoanCIP.CursorLocation = 2;
	rsLoanCIP.LockType = 3;
	rsLoanCIP.Open();
	Response.Redirect("m008e0101.asp?intLoan_Req_id="+Request.Form("MM_recordId"));
}

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var IsClient = true;
if (!rsLoan.EOF) {
	if ((rsLoan.Fields.Item("insEq_user_type").Value == 3) && (rsLoan.Fields.Item("intEq_user_id").Value > 0)) {
		var rsLoanCIP = Server.CreateObject("ADODB.Recordset");
		rsLoanCIP.ActiveConnection = MM_cnnASP02_STRING;
		rsLoanCIP.Source = "{call dbo.cp_loan_cip2("+Request.QueryString("intLoan_req_id")+",0,"+rsLoan.Fields.Item("intEq_user_id").Value+",'Q',0)}";
		rsLoanCIP.CursorType = 0;
		rsLoanCIP.CursorLocation = 2;
		rsLoanCIP.LockType = 3;
		rsLoanCIP.Open();	
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
		if (document.frm0102.ServiceDate.length > 1) {
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
<h5>TAP Date</h5>
<%
if (!IsClient) {
%>
<i>Informatin not available.  Either this is an institutional loan or the client was not found.</i>
<%
} else {
%>
<i>Select one of the following service date as the TAP date for this loan.</i>
<hr>
<table cellpadding="1" cellspacing=1" style="border: 1px solid">
	<tr>
		<th class="headrow">&nbsp;</th>
<!--	<th class="headrow" nowrap>Service Date</th>-->
		<th class="headrow" nowrap>TAP Date</th>
<!--	<th class="headrow" nowrap>Year/Cycle</th>-->
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
while (!rsLoanCIP.EOF) {
%>
	<tr class="<%=(((count%2)=="0")?"rowa":"rowb")%>">
		<td style="background-color: #FFFFFF"><input type="radio" name="ServiceDate" value="<%=rsLoanCIP.Fields.Item("intSrv_Note_id").Value%>" <%=((rsLoanCIP.Fields.Item("bitIs_LN_CIP").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
<!--	<td align="center"><%=FilterDate(rsLoanCIP.Fields.Item("dtsService_Date").Value)%></td>-->
		<td align="center"><%=rsLoanCIP.Fields.Item("chvCIP_Date").Value%></td>
<!--	<td align="center"><%=rsLoanCIP.Fields.Item("insYearCycle").Value%></td>-->
		<td align="left"><%=rsLoanCIP.Fields.Item("chvService_provider").Value%></td>
		<td align="center"><%=rsLoanCIP.Fields.Item("chvSrv_Code").Value%></td>
<%
if (String(Request.QueryString("ShowFS"))=="1") {
%>		
		<td align="center"><%=rsLoanCIP.Fields.Item("chvFunding_Source").Value%></td>
<%
}
%>
	</tr>
<%
	rsLoanCIP.MoveNext();
	count++;
}
%>		
</table>
<hr>
<%
if (count > 0) {
%>
<input type="button" value="Link" onClick="Link();" class="btnstyle">
<input type="button" value="Cancel" onClick="window.location.href='m008e0101.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>';" class="btnstyle">
<%
if (String(Request.QueryString("ShowFS"))=="1") {
%>
<input type="button" value="Hide Funding Source" onClick="window.location.href='m008e0102.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>&ShowFS=0';" class="btnstyle">
<%
} else {
%>
<input type="button" value="Show Funding Source" onClick="window.location.href='m008e0102.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>&ShowFS=1';" class="btnstyle">
<%
}
}
}
%>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("intLoan_Req_id")%>">
<input type="hidden" name="MM_noteId" value="">
</form>
</body>
</html>
<%
rsLoan.Close();
%>