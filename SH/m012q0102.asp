<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_school_funding_src(0,"+ Request.QueryString("insSchool_id") + ",0,0,0,0,'Q',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();
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
		<th class="headrow" align="left">Referral ID</th>	
		<th class="headrow" align="left">Referral Date</th>
		<th class="headrow" align="left">Referral Type</th>				
		<th class="headrow" align="left">Funding Source</th>
		<th class="headrow" align="left">Selected</th>
    </tr>
<% 
while (!rsFundingSource.EOF) { 
%>
    <tr>         
		<td align="center" nowrap><a href="m012e0102.asp?insRefAgt_id=<%=(rsFundingSource.Fields.Item("insRefAgt_id").Value)%>&intReferral_id=<%=(rsFundingSource.Fields.Item("intReferral_id").Value)%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>"><%=ZeroPadFormat(rsFundingSource.Fields.Item("intReferral_id").Value,8)%></a></td>
		<td align="center"><%=FilterDate(rsFundingSource.Fields.Item("dtsRefral_date").Value)%></td>
		<td align="center"><%=(rsFundingSource.Fields.Item("chvRefAgt").Value)%>&nbsp;</td>	
		<td align="center"><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%>&nbsp;</td>
		<td align="center"><%=(rsFundingSource.Fields.Item("bitIs_Sel_FundingSrc").Value)%>&nbsp;</td>
    </tr>
<%
	rsFundingSource.MoveNext();
}
%>
</table>
<hr>
</body>
</html>
<%
rsFundingSource.Close();
%>