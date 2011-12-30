<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsBuyoutFundingSource = Server.CreateObject("ADODB.Recordset");
rsBuyoutFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+ Request.QueryString("intBuyout_req_id") + ",0,0,0,0,'Q',0)}";
rsBuyoutFundingSource.CursorType = 0;
rsBuyoutFundingSource.CursorLocation = 2;
rsBuyoutFundingSource.LockType = 3;
rsBuyoutFundingSource.Open();
%>
<html>
<head>
	<title>Funding Source</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=500,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>	
</head>
<body>
<h5>Funding Source</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
    <tr> 
		<th class="headrow" nowrap align="left">Referral ID</th>	
		<th class="headrow" nowrap align="left">Referral Date</th>
		<th class="headrow" nowrap align="left">Referral Type</th>				
		<th class="headrow" nowrap align="center">Funding Source</th>
		<th class="headrow" nowrap align="left">Selected</th>
    </tr>
<% 
while (!rsBuyoutFundingSource.EOF) { 
%>
    <tr>         
		<td valign="top" align="center" nowrap><a href="m010e0103.asp?intReferral_id=<%=(rsBuyoutFundingSource.Fields.Item("intReferral_id").Value)%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&intAdult_id=<%=rsBuyoutFundingSource.Fields.Item("intEq_user_id").Value%>"><%=ZeroPadFormat(rsBuyoutFundingSource.Fields.Item("intReferral_id").Value,8)%></a></td>
		<td valign="top" align="left"><%=FilterDate(rsBuyoutFundingSource.Fields.Item("dtsRefral_date").Value)%></td>
		<td valign="top" align="center"><%=(rsBuyoutFundingSource.Fields.Item("chvRefAgt").Value)%>&nbsp;</td>	
		<td valign="top" align="left"><%=(rsBuyoutFundingSource.Fields.Item("chvfunding_source_name").Value)%>&nbsp;</td>
		<td valign="top" align="center"><%=(rsBuyoutFundingSource.Fields.Item("bitIs_Sel_FundingSrc").Value)%>&nbsp;</td>
    </tr>
<%
	rsBuyoutFundingSource.MoveNext();
}
%>
</table>
<!--
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('m010a0103.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>','wA0103');">Add Funding Source</a></td>
	</tr>
</table>
-->
</body>
</html>
<%
rsBuyoutFundingSource.Close();
%>